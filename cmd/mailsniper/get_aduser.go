package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
	"mailsniper-go/internal/output"
)

func newGetADUserCmd() *cobra.Command {
	var (
		emailList   string
		hostname    string
		ewsURL      string
		username    string
		password    string
		accessToken string
		version     string
		outFile     string
		outFmt      string
		skipTLS     bool
	)

	cmd := &cobra.Command{
		Use:   "get-aduser",
		Short: "Resolve AD usernames from email addresses via EWS (Get-ADUsernameFromEWS)",
		Long: `Uses the EWS ResolveNames operation to map email addresses to their
Active Directory display names / usernames.

Equivalent to the PowerShell Get-ADUsernameFromEWS function.`,
		RunE: func(cmd *cobra.Command, args []string) error {
			emails, err := readLines(emailList)
			if err != nil {
				return fmt.Errorf("read email list: %w", err)
			}

			client, err := buildEWSClient(hostname, ewsURL, username, password, accessToken, version, skipTLS)
			if err != nil {
				return err
			}

			fmt.Printf("[*] Resolving %d email(s) via EWS...\n", len(emails))

			var userResults []output.UserResult
			for _, email := range emails {
				name, err := client.ResolveNames(email)
				if err != nil {
					fmt.Printf("[-] %s: error: %v\n", email, err)
					continue
				}
				if name != "" {
					fmt.Printf("[+] %s -> %s\n", email, name)
				} else {
					fmt.Printf("[-] %s: not resolved\n", email)
				}
				userResults = append(userResults, output.UserResult{
					Username:   email,
					Valid:      name != "",
					ADUsername: name,
				})
			}

			if outFile != "" {
				fmt.Printf("[*] Writing results to %s\n", outFile)
				if err := output.WriteUserResults(outFile, output.ParseFormat(outFmt), userResults); err != nil {
					return fmt.Errorf("write results: %w", err)
				}
			}
			return nil
		},
	}

	f := cmd.Flags()
	f.StringVar(&emailList, "email-list", "", "File with email addresses (one per line)")
	f.StringVar(&hostname, "hostname", "", "Exchange server hostname")
	f.StringVar(&ewsURL, "ews-url", "", "Full EWS URL (overrides --hostname)")
	f.StringVar(&username, "username", "", "Username for authentication")
	f.StringVar(&password, "password", "", "Password")
	f.StringVar(&accessToken, "access-token", "", "OAuth2 Bearer token")
	f.StringVar(&version, "exchange-version", "Exchange2010", "Exchange server version")
	f.StringVar(&outFile, "output", "", "Output file path")
	f.StringVar(&outFmt, "output-format", "txt", "Output format: csv, json, txt")
	f.BoolVar(&skipTLS, "skip-tls", false, "Skip TLS certificate verification")

	_ = cmd.MarkFlagRequired("email-list")
	return cmd
}
