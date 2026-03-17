package cmd

import (
	"fmt"
	"strings"
	"sync"

	"github.com/spf13/cobra"
	"mailsniper-go/internal/ews"
	"mailsniper-go/internal/output"
)

func newGlobalSearchCmd() *cobra.Command {
	var (
		impersonationAccount string
		hostname             string
		ewsURL               string
		autoDiscoverEmail    string
		adminUsername        string
		adminPassword        string
		accessToken          string
		version              string
		mailsPerUser         int
		termsArg             []string
		outputFile           string
		outputFmt            string
		folder               string
		regexStr             string
		checkAttachments     bool
		downloadDir          string
		skipTLS              bool
		emailListFile        string
		threads              int
	)

	cmd := &cobra.Command{
		Use:   "global-search",
		Short: "Search all mailboxes via EWS impersonation (Invoke-GlobalMailSearch)",
		Long: `Connects to Exchange via EWS using ApplicationImpersonation to search
every user's mailbox for specific terms.

Equivalent to the PowerShell Invoke-GlobalMailSearch function.`,
		RunE: func(cmd *cobra.Command, args []string) error {
			terms := termsArg
			if len(terms) == 0 {
				terms = defaultTerms
			}

			// Resolve EWS URL.
			if ewsURL == "" && hostname == "" && autoDiscoverEmail != "" {
				fmt.Printf("[*] Autodiscovering from %s...\n", autoDiscoverEmail)
				discovered, err := ews.Autodiscover(autoDiscoverEmail, adminPassword, skipTLS)
				if err == nil {
					ewsURL = discovered
					fmt.Printf("[+] Autodiscover: %s\n", ewsURL)
				} else {
					return fmt.Errorf("autodiscover failed: %w", err)
				}
			}

			if ewsURL == "" && hostname != "" {
				ewsURL = ews.BuildEWSURL(hostname)
			}
			if ewsURL == "" {
				return fmt.Errorf("provide --hostname, --ews-url, or --autodiscover-email")
			}

			// Build impersonation account credentials.
			impUser := impersonationAccount
			if impUser == "" {
				impUser = adminUsername
			}

			// Fetch mailbox list.
			var mailboxes []string
			if emailListFile != "" {
				var err error
				mailboxes, err = readLines(emailListFile)
				if err != nil {
					return fmt.Errorf("read email list: %w", err)
				}
				fmt.Printf("[*] Loaded %d mailboxes from %s\n", len(mailboxes), emailListFile)
			} else {
				// Try to get GAL from EWS.
				adminClient, err := buildEWSClient(hostname, ewsURL, adminUsername, adminPassword, accessToken, version, skipTLS)
				if err != nil {
					return err
				}
				fmt.Println("[*] Fetching mailbox list from GAL via EWS...")
				entries, err := adminClient.GetGALEWS(0)
				if err != nil {
					return fmt.Errorf("failed to get GAL: %w", err)
				}
				for _, e := range entries {
					if e.EmailAddress != "" {
						mailboxes = append(mailboxes, e.EmailAddress)
					}
				}
				fmt.Printf("[*] Found %d mailboxes in GAL\n", len(mailboxes))
			}

			if len(mailboxes) == 0 {
				return fmt.Errorf("no mailboxes found; provide --email-list")
			}

			fmt.Printf("[*] Searching %d mailboxes with impersonation account: %s\n", len(mailboxes), impUser)
			fmt.Printf("[*] Terms: %s | Folder: %s | Max per mailbox: %d\n",
				strings.Join(terms, ", "), folder, mailsPerUser)

			sem := make(chan struct{}, threads)
			var wg sync.WaitGroup
			var allResults []output.MailResult

			for _, mb := range mailboxes {
				mb := mb
				wg.Add(1)
				sem <- struct{}{}
				go func() {
					defer wg.Done()
					defer func() { <-sem }()

					client, err := buildEWSClient(hostname, ewsURL, impUser, adminPassword, accessToken, version, skipTLS)
					if err != nil {
						fmt.Printf("[-] %s: client error: %v\n", mb, err)
						return
					}
					client.ImpersonateEmail = mb

					results, err := searchMailbox(client, mb, folder, mailsPerUser, terms, regexStr, checkAttachments, downloadDir)
					if err != nil {
						fmt.Printf("[-] %s: %v\n", mb, err)
						return
					}
					if len(results) > 0 {
						fmt.Printf("[+] %s: %d match(es)\n", mb, len(results))
						appendResults(&allResults, results)
					}
				}()
			}
			wg.Wait()

			fmt.Printf("\n[+] Total matches: %d\n", len(allResults))

			if outputFile != "" {
				fmt.Printf("[*] Writing results to %s\n", outputFile)
				if err := output.WriteMailResults(outputFile, output.ParseFormat(outputFmt), allResults); err != nil {
					return fmt.Errorf("write results: %w", err)
				}
			}
			return nil
		},
	}

	f := cmd.Flags()
	f.StringVar(&impersonationAccount, "impersonation-account", "", "Account to grant/use ApplicationImpersonation role")
	f.StringVar(&hostname, "hostname", "", "Exchange server hostname")
	f.StringVar(&ewsURL, "ews-url", "", "Full EWS URL (overrides --hostname)")
	f.StringVar(&autoDiscoverEmail, "autodiscover-email", "", "Email address for autodiscovery")
	f.StringVar(&adminUsername, "username", "", "Admin username (domain\\username or UPN)")
	f.StringVar(&adminPassword, "password", "", "Admin password")
	f.StringVar(&accessToken, "access-token", "", "OAuth2 Bearer access token")
	f.StringVar(&version, "exchange-version", "Exchange2010", "Exchange server version")
	f.IntVar(&mailsPerUser, "mails-per-user", 100, "Number of emails to retrieve per mailbox")
	f.StringArrayVar(&termsArg, "terms", nil, "Search terms (default: *password*,*creds*,*credentials*)")
	f.StringVar(&outputFile, "output", "", "Output file path")
	f.StringVar(&outputFmt, "output-format", "csv", "Output format: csv, json, txt")
	f.StringVar(&folder, "folder", "inbox", "Folder to search (use 'all' for all folders)")
	f.StringVar(&regexStr, "regex", "", "Regex pattern (overrides --terms)")
	f.BoolVar(&checkAttachments, "check-attachments", false, "Search attachment content")
	f.StringVar(&downloadDir, "download-dir", "", "Directory to save matched attachments")
	f.BoolVar(&skipTLS, "skip-tls", false, "Skip TLS certificate verification")
	f.StringVar(&emailListFile, "email-list", "", "File with email addresses to search (one per line)")
	f.IntVar(&threads, "threads", 5, "Number of concurrent mailbox searches")

	return cmd
}
