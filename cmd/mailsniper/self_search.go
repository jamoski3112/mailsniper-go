package cmd

import (
	"bufio"
	"fmt"
	"os"
	"strings"
	"sync"

	"github.com/spf13/cobra"
	"mailsniper-go/internal/ews"
	"mailsniper-go/internal/output"
)

// shared search helper used by both self-search and global-search
func searchMailbox(
	client *ews.Client,
	mailbox string,
	folder string,
	mailsPerUser int,
	terms []string,
	regex string,
	checkAttachments bool,
	downloadDir string,
) ([]output.MailResult, error) {
	var items []struct {
		ID         string
		ChangeKey  string
		Subject    string
		DateSent   string
		From       string
		FolderName string
	}
	var err error

	if strings.ToLower(folder) == "all" {
		items, err = client.FindItemsAllFolders(mailbox, mailsPerUser)
	} else {
		raw, e := client.FindItems(mailbox, folder, mailsPerUser)
		if e != nil {
			return nil, e
		}
		for _, r := range raw {
			items = append(items, struct {
				ID         string
				ChangeKey  string
				Subject    string
				DateSent   string
				From       string
				FolderName string
			}{r.ID, r.ChangeKey, r.Subject, r.DateSent, r.From, folder})
		}
	}
	if err != nil {
		return nil, err
	}

	var results []output.MailResult
	for _, item := range items {
		msg, err := client.GetItem(item.ID, item.ChangeKey)
		if err != nil {
			continue
		}

		matched := matchTerms(msg.Subject, msg.Body, terms, regex)
		attachMatched := ""
		if checkAttachments {
			for _, att := range msg.Attachments {
				if isSearchableAttachment(att.Name) {
					m := matchTerms(att.Name, att.Content, terms, regex)
					if m != "" {
						attachMatched = m
						if downloadDir != "" {
							saveAttachment(downloadDir, mailbox, att.Name, att.Content)
						}
					}
				}
			}
		}

		finalMatch := matched
		if finalMatch == "" {
			finalMatch = attachMatched
		}
		if finalMatch == "" {
			continue
		}

		results = append(results, output.MailResult{
			Mailbox:   mailbox,
			Folder:    item.FolderName,
			From:      msg.From,
			DateSent:  msg.DateSent,
			Subject:   msg.Subject,
			Body:      truncateBody(msg.Body, 500),
			MatchedBy: finalMatch,
		})
	}
	return results, nil
}

func matchTerms(subject, body string, terms []string, regexStr string) string {
	if regexStr != "" {
		re, err := compileRegex(regexStr)
		if err == nil {
			text := subject + " " + body
			if re.MatchString(text) {
				return regexStr
			}
		}
		return ""
	}
	haystack := strings.ToLower(subject + " " + body)
	for _, t := range terms {
		pattern := strings.ToLower(strings.Trim(t, "*"))
		if strings.Contains(haystack, pattern) {
			return t
		}
	}
	return ""
}

func isSearchableAttachment(name string) bool {
	name = strings.ToLower(name)
	for _, ext := range []string{".bat", ".htm", ".msg", ".pdf", ".txt", ".ps1", ".doc", ".xls"} {
		if strings.HasSuffix(name, ext) {
			return true
		}
	}
	return false
}

func saveAttachment(dir, mailbox, name, content string) {
	os.MkdirAll(dir, 0o755)
	safe := strings.ReplaceAll(mailbox, "@", "_at_")
	safe = strings.ReplaceAll(safe, "/", "_")
	path := fmt.Sprintf("%s/%s_%s", dir, safe, name)
	os.WriteFile(path, []byte(content), 0o644)
}

func truncateBody(s string, n int) string {
	runes := []rune(s)
	if len(runes) <= n {
		return s
	}
	return string(runes[:n]) + "..."
}

func readLines(path string) ([]string, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	var lines []string
	sc := bufio.NewScanner(f)
	for sc.Scan() {
		l := strings.TrimSpace(sc.Text())
		if l != "" {
			lines = append(lines, l)
		}
	}
	return lines, sc.Err()
}

func buildEWSClient(hostname, ewsURL, username, password, accessToken, version string, skipTLS bool) (*ews.Client, error) {
	if ewsURL == "" && hostname != "" {
		ewsURL = ews.BuildEWSURL(hostname)
	}
	if ewsURL == "" {
		return nil, fmt.Errorf("provide --hostname or --ews-url")
	}
	if accessToken != "" {
		return ews.NewBearerClient(ewsURL, accessToken, version, skipTLS), nil
	}
	return ews.NewClient(ewsURL, username, password, version, skipTLS), nil
}

// newMu is a package-level mutex for concurrent result collection.
var newMu sync.Mutex

func appendResults(dst *[]output.MailResult, src []output.MailResult) {
	newMu.Lock()
	defer newMu.Unlock()
	*dst = append(*dst, src...)
}

// defaultTerms matches the original MailSniper defaults.
var defaultTerms = []string{"*password*", "*creds*", "*credentials*"}

// newSelfSearchCmd builds the self-search subcommand.
func newSelfSearchCmd() *cobra.Command {
	var (
		mailbox          string
		hostname         string
		ewsURL           string
		username         string
		password         string
		accessToken      string
		version          string
		mailsPerUser     int
		termsArg         []string
		outputFile       string
		outputFmt        string
		folder           string
		regexStr         string
		checkAttachments bool
		downloadDir      string
		skipTLS          bool
		otherMailbox     string
	)

	cmd := &cobra.Command{
		Use:   "self-search",
		Short: "Search the current user's mailbox via EWS (Invoke-SelfSearch)",
		Long: `Connects to Exchange via EWS and searches the current user's mailbox
for specific terms in email subjects and bodies.

Equivalent to the PowerShell Invoke-SelfSearch function.`,
		RunE: func(cmd *cobra.Command, args []string) error {
			terms := termsArg
			if len(terms) == 0 {
				terms = defaultTerms
			}

			// Resolve EWS URL via autodiscover if needed.
			if ewsURL == "" && hostname == "" && mailbox != "" {
				fmt.Printf("[*] Attempting autodiscover for %s...\n", mailbox)
				discovered, err := ews.Autodiscover(mailbox, password, skipTLS)
				if err == nil {
					ewsURL = discovered
					fmt.Printf("[+] Autodiscover: %s\n", ewsURL)
				}
			}

			client, err := buildEWSClient(hostname, ewsURL, username, password, accessToken, version, skipTLS)
			if err != nil {
				return err
			}

			target := mailbox
			if otherMailbox != "" {
				target = otherMailbox
				client.ImpersonateEmail = otherMailbox
			}

			fmt.Printf("[*] Searching mailbox: %s\n", target)
			fmt.Printf("[*] Terms: %s\n", strings.Join(terms, ", "))
			fmt.Printf("[*] Folder: %s | Max: %d emails\n", folder, mailsPerUser)

			results, err := searchMailbox(client, target, folder, mailsPerUser, terms, regexStr, checkAttachments, downloadDir)
			if err != nil {
				return fmt.Errorf("search failed: %w", err)
			}

			fmt.Printf("[+] Found %d matching email(s)\n", len(results))
			for _, r := range results {
				fmt.Printf("  [%s] From: %s | Subject: %s | Match: %s\n",
					r.DateSent, r.From, r.Subject, r.MatchedBy)
			}

			if outputFile != "" {
				fmt.Printf("[*] Writing results to %s\n", outputFile)
				if err := output.WriteMailResults(outputFile, output.ParseFormat(outputFmt), results); err != nil {
					return fmt.Errorf("write results: %w", err)
				}
			}
			return nil
		},
	}

	f := cmd.Flags()
	f.StringVar(&mailbox, "mailbox", "", "Email address of the mailbox to search")
	f.StringVar(&hostname, "hostname", "", "Exchange server hostname")
	f.StringVar(&ewsURL, "ews-url", "", "Full EWS URL (overrides --hostname)")
	f.StringVar(&username, "username", "", "Username for authentication")
	f.StringVar(&password, "password", "", "Password for authentication")
	f.StringVar(&accessToken, "access-token", "", "OAuth2 Bearer access token")
	f.StringVar(&version, "exchange-version", "Exchange2010", "Exchange server version")
	f.IntVar(&mailsPerUser, "mails-per-user", 100, "Number of emails to retrieve")
	f.StringArrayVar(&termsArg, "terms", nil, "Search terms (default: *password*,*creds*,*credentials*)")
	f.StringVar(&outputFile, "output", "", "Output file path")
	f.StringVar(&outputFmt, "output-format", "csv", "Output format: csv, json, txt")
	f.StringVar(&folder, "folder", "inbox", "Folder to search (use 'all' for all folders)")
	f.StringVar(&regexStr, "regex", "", "Regex pattern (overrides --terms)")
	f.BoolVar(&checkAttachments, "check-attachments", false, "Search attachment content")
	f.StringVar(&downloadDir, "download-dir", "", "Directory to save matched attachments")
	f.BoolVar(&skipTLS, "skip-tls", false, "Skip TLS certificate verification")
	f.StringVar(&otherMailbox, "other-mailbox", "", "Search a different user's mailbox (requires impersonation)")

	return cmd
}
