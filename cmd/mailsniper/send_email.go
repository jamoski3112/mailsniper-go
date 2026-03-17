package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
)

func newSendEmailCmd() *cobra.Command {
	var (
		hostname    string
		ewsURL      string
		username    string
		password    string
		accessToken string
		version     string
		recipient   string
		subject     string
		body        string
		skipTLS     bool
	)

	cmd := &cobra.Command{
		Use:   "send-email",
		Short: "Send an email via EWS (Send-EWSEmail)",
		Long: `Connects to EWS and sends an email to the specified recipient.

Equivalent to the PowerShell Send-EWSEmail function.`,
		RunE: func(cmd *cobra.Command, args []string) error {
			client, err := buildEWSClient(hostname, ewsURL, username, password, accessToken, version, skipTLS)
			if err != nil {
				return err
			}

			fmt.Printf("[*] Sending email to %s via %s...\n", recipient, ewsURL)
			if err := client.SendEmail(recipient, subject, body); err != nil {
				return fmt.Errorf("send email: %w", err)
			}
			fmt.Printf("[+] Email sent successfully to %s\n", recipient)
			return nil
		},
	}

	f := cmd.Flags()
	f.StringVar(&hostname, "hostname", "", "Exchange server hostname")
	f.StringVar(&ewsURL, "ews-url", "", "Full EWS URL (overrides --hostname)")
	f.StringVar(&username, "username", "", "Username for authentication")
	f.StringVar(&password, "password", "", "Password")
	f.StringVar(&accessToken, "access-token", "", "OAuth2 Bearer token")
	f.StringVar(&version, "exchange-version", "Exchange2010", "Exchange server version")
	f.StringVar(&recipient, "recipient", "", "Recipient email address")
	f.StringVar(&subject, "subject", "", "Email subject")
	f.StringVar(&body, "body", "", "Email body (HTML)")
	f.BoolVar(&skipTLS, "skip-tls", false, "Skip TLS certificate verification")

	_ = cmd.MarkFlagRequired("recipient")
	_ = cmd.MarkFlagRequired("subject")
	_ = cmd.MarkFlagRequired("body")
	return cmd
}
