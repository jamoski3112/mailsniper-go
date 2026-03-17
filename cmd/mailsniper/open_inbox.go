package cmd

import (
	"fmt"
	"sync"

	"github.com/spf13/cobra"
	"mailsniper-go/internal/ews"
)

func newOpenInboxCmd() *cobra.Command {
	var (
		hostname    string
		ewsURL      string
		username    string
		password    string
		accessToken string
		version     string
		emailList   string
		threads     int
		skipTLS     bool
	)

	cmd := &cobra.Command{
		Use:   "open-inbox",
		Short: "Check if current user can read other users' inboxes (Invoke-OpenInboxFinder)",
		Long: `Attempts to read the Inbox of each email address in a list using the current
user's credentials, identifying any accessible mailboxes.

Equivalent to the PowerShell Invoke-OpenInboxFinder function.`,
		RunE: func(cmd *cobra.Command, args []string) error {
			emails, err := readLines(emailList)
			if err != nil {
				return fmt.Errorf("read email list: %w", err)
			}

			if ewsURL == "" && hostname != "" {
				ewsURL = ews.BuildEWSURL(hostname)
			}
			if ewsURL == "" {
				return fmt.Errorf("provide --hostname or --ews-url")
			}

			fmt.Printf("[*] Checking inbox access for %d addresses using %s...\n", len(emails), username)

			sem := make(chan struct{}, threads)
			var wg sync.WaitGroup
			var mu sync.Mutex
			var accessible []string

			for _, email := range emails {
				email := email
				wg.Add(1)
				sem <- struct{}{}
				go func() {
					defer wg.Done()
					defer func() { <-sem }()

					client, err := buildEWSClient(hostname, ewsURL, username, password, accessToken, version, skipTLS)
					if err != nil {
						return
					}
					// Attempt to list one item in the target inbox.
					client.ImpersonateEmail = ""
					_, err = client.FindItems(email, "inbox", 1)
					if err == nil {
						fmt.Printf("[+] ACCESSIBLE: %s\n", email)
						mu.Lock()
						accessible = append(accessible, email)
						mu.Unlock()
					}
				}()
			}
			wg.Wait()

			fmt.Printf("\n[+] %d accessible inbox(es) found\n", len(accessible))
			return nil
		},
	}

	f := cmd.Flags()
	f.StringVar(&hostname, "hostname", "", "Exchange server hostname")
	f.StringVar(&ewsURL, "ews-url", "", "Full EWS URL (overrides --hostname)")
	f.StringVar(&username, "username", "", "Username to authenticate as")
	f.StringVar(&password, "password", "", "Password")
	f.StringVar(&accessToken, "access-token", "", "OAuth2 Bearer token")
	f.StringVar(&version, "exchange-version", "Exchange2010", "Exchange server version")
	f.StringVar(&emailList, "email-list", "", "File with email addresses to check (one per line)")
	f.IntVar(&threads, "threads", 5, "Number of concurrent checks")
	f.BoolVar(&skipTLS, "skip-tls", false, "Skip TLS certificate verification")

	_ = cmd.MarkFlagRequired("email-list")
	return cmd
}
