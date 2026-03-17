package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
	"mailsniper-go/internal/ews"
	"mailsniper-go/internal/output"
)

func newListFoldersCmd() *cobra.Command {
	var (
		mailbox     string
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
		Use:   "list-folders",
		Short: "List all mailbox folders via EWS (Get-MailboxFolders)",
		Long: `Connects to EWS and retrieves a list of all folders and subfolders
in the specified mailbox.

Equivalent to the PowerShell Get-MailboxFolders function.`,
		RunE: func(cmd *cobra.Command, args []string) error {
			if ewsURL == "" && hostname == "" && mailbox != "" {
				fmt.Printf("[*] Attempting autodiscover for %s...\n", mailbox)
				discovered, err := ews.Autodiscover(mailbox, password, skipTLS)
				if err == nil {
					ewsURL = discovered
				}
			}

			client, err := buildEWSClient(hostname, ewsURL, username, password, accessToken, version, skipTLS)
			if err != nil {
				return err
			}

			target := mailbox
			if target == "" {
				target = username
			}
			fmt.Printf("[*] Listing folders for %s...\n", target)

			folders, err := client.FindFolders(target)
			if err != nil {
				return fmt.Errorf("find folders: %w", err)
			}

			fmt.Printf("[+] Found %d folder(s)\n\n", len(folders))
			fmt.Printf("%-45s %8s %8s\n", "Folder Name", "Total", "Unread")
			fmt.Printf("%-45s %8s %8s\n", "--------------------------------------------", "--------", "--------")
			var folderResults []output.FolderResult
			for _, f := range folders {
				fmt.Printf("%-45s %8d %8d\n", f.DisplayName, f.TotalCount, f.UnreadCount)
				folderResults = append(folderResults, output.FolderResult{
					ID:               f.ID,
					DisplayName:      f.DisplayName,
					TotalCount:       f.TotalCount,
					UnreadCount:      f.UnreadCount,
					ChildFolderCount: f.ChildFolderCount,
				})
			}

			if outFile != "" {
				fmt.Printf("\n[*] Writing results to %s\n", outFile)
				if err := output.WriteFolderResults(outFile, output.ParseFormat(outFmt), folderResults); err != nil {
					return fmt.Errorf("write results: %w", err)
				}
			}
			return nil
		},
	}

	f := cmd.Flags()
	f.StringVar(&mailbox, "mailbox", "", "Email address of the mailbox")
	f.StringVar(&hostname, "hostname", "", "Exchange server hostname")
	f.StringVar(&ewsURL, "ews-url", "", "Full EWS URL (overrides --hostname)")
	f.StringVar(&username, "username", "", "Username for authentication")
	f.StringVar(&password, "password", "", "Password")
	f.StringVar(&accessToken, "access-token", "", "OAuth2 Bearer token")
	f.StringVar(&version, "exchange-version", "Exchange2010", "Exchange server version")
	f.StringVar(&outFile, "output", "", "Output file path")
	f.StringVar(&outFmt, "output-format", "txt", "Output format: csv, json, txt")
	f.BoolVar(&skipTLS, "skip-tls", false, "Skip TLS certificate verification")

	return cmd
}
