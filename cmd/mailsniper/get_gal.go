package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
	"mailsniper-go/internal/ews"
	"mailsniper-go/internal/output"
	owaClient "mailsniper-go/internal/owa"
)

func newGetGALCmd() *cobra.Command {
	var (
		hostname    string
		username    string
		password    string
		accessToken string
		version     string
		outFile     string
		outFmt      string
		maxResults  int
		useOWA      bool
		skipTLS     bool
	)

	cmd := &cobra.Command{
		Use:   "get-gal",
		Short: "Retrieve the Global Address List (Get-GlobalAddressList)",
		Long: `Attempts to retrieve the Global Address List via OWA FindPeople API
(Exchange 2013+) and falls back to EWS FindPeople if OWA fails.

Equivalent to the PowerShell Get-GlobalAddressList function.`,
		RunE: func(cmd *cobra.Command, args []string) error {
			var galEntries []output.GALResult

			if useOWA {
				fmt.Printf("[*] Fetching GAL via OWA FindPeople from %s...\n", hostname)
				oc, err := owaClient.NewClient(hostname, skipTLS)
				if err != nil {
					return fmt.Errorf("OWA client: %w", err)
				}
				entries, err := oc.GetGALFindPeople(username, password, maxResults)
				if err != nil {
					fmt.Printf("[-] OWA GAL failed: %v — falling back to EWS\n", err)
				} else {
					for _, e := range entries {
						galEntries = append(galEntries, output.GALResult{
							DisplayName:  e.DisplayName,
							EmailAddress: e.EmailAddress,
						})
					}
				}
			}

			if len(galEntries) == 0 {
				fmt.Printf("[*] Fetching GAL via EWS FindPeople from %s...\n", hostname)
				ewsURL := ews.BuildEWSURL(hostname)
				var client *ews.Client
				if accessToken != "" {
					client = ews.NewBearerClient(ewsURL, accessToken, version, skipTLS)
				} else {
					client = ews.NewClient(ewsURL, username, password, version, skipTLS)
				}
				entries, err := client.GetGALEWS(maxResults)
				if err != nil {
					return fmt.Errorf("EWS GAL failed: %w", err)
				}
				for _, e := range entries {
					galEntries = append(galEntries, output.GALResult{
						DisplayName:  e.DisplayName,
						EmailAddress: e.EmailAddress,
					})
				}
			}

			fmt.Printf("[+] Found %d GAL entries\n", len(galEntries))
			for _, e := range galEntries {
				fmt.Printf("  %s <%s>\n", e.DisplayName, e.EmailAddress)
			}

			if outFile != "" {
				fmt.Printf("[*] Writing GAL to %s\n", outFile)
				if err := output.WriteGALResults(outFile, output.ParseFormat(outFmt), galEntries); err != nil {
					return fmt.Errorf("write GAL: %w", err)
				}
			}
			return nil
		},
	}

	f := cmd.Flags()
	f.StringVar(&hostname, "hostname", "", "Exchange/OWA server hostname")
	f.StringVar(&username, "username", "", "Username (domain\\username or UPN)")
	f.StringVar(&password, "password", "", "Password")
	f.StringVar(&accessToken, "access-token", "", "OAuth2 Bearer token")
	f.StringVar(&version, "exchange-version", "Exchange2013", "Exchange server version (2013+ for FindPeople)")
	f.StringVar(&outFile, "output", "", "Output file path")
	f.StringVar(&outFmt, "output-format", "txt", "Output format: csv, json, txt")
	f.IntVar(&maxResults, "max", 0, "Maximum number of entries (0 = all)")
	f.BoolVar(&useOWA, "owa", false, "Use OWA FindPeople API first")
	f.BoolVar(&skipTLS, "skip-tls", false, "Skip TLS certificate verification")

	_ = cmd.MarkFlagRequired("hostname")
	return cmd
}
