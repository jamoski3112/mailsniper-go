package cmd

import (
	"fmt"
	"strings"

	"github.com/spf13/cobra"
	owaClient "mailsniper-go/internal/owa"
)

func newHarvestDomainCmd() *cobra.Command {
	var (
		hostname string
		skipTLS  bool
	)

	cmd := &cobra.Command{
		Use:   "harvest-domain",
		Short: "Discover the login domain from OWA headers (Invoke-DomainHarvestOWA)",
		Long: `Sends a request to an OWA portal and inspects the WWW-Authenticate
response header to determine the Active Directory domain name.

Equivalent to the PowerShell Invoke-DomainHarvestOWA function.`,
		RunE: func(cmd *cobra.Command, args []string) error {
			fmt.Printf("[*] Probing OWA at %s for domain information...\n", hostname)
			oc, err := owaClient.NewClient(hostname, skipTLS)
			if err != nil {
				return fmt.Errorf("OWA client: %w", err)
			}

			header, err := oc.GetWWWAuthenticate()
			if err != nil {
				return fmt.Errorf("OWA request failed: %w", err)
			}

			if header == "" {
				fmt.Println("[-] No WWW-Authenticate header returned. Server may not use Windows auth.")
				return nil
			}

			fmt.Printf("[+] WWW-Authenticate: %s\n", header)

			// Extract realm/domain from NTLM or Negotiate header.
			domain := extractDomain(header)
			if domain != "" {
				fmt.Printf("[+] Discovered domain: %s\n", domain)
			} else {
				fmt.Println("[*] Could not parse domain from header; check raw value above.")
			}
			return nil
		},
	}

	f := cmd.Flags()
	f.StringVar(&hostname, "hostname", "", "OWA server hostname")
	f.BoolVar(&skipTLS, "skip-tls", false, "Skip TLS certificate verification")

	_ = cmd.MarkFlagRequired("hostname")
	return cmd
}

func extractDomain(header string) string {
	// Look for realm="..." in NTLM or Negotiate responses.
	lower := strings.ToLower(header)
	if idx := strings.Index(lower, `realm="`); idx != -1 {
		rest := header[idx+7:]
		end := strings.Index(rest, `"`)
		if end != -1 {
			return rest[:end]
		}
	}
	// Look for domain=... (some NTLM implementations).
	if idx := strings.Index(lower, "domain="); idx != -1 {
		rest := header[idx+7:]
		rest = strings.Trim(rest, `"`)
		end := strings.IndexAny(rest, " ,\r\n")
		if end != -1 {
			rest = rest[:end]
		}
		return rest
	}
	return ""
}
