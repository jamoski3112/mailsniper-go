// Package output provides CSV, JSON, and plain-text writers for MailSniper results.
package output

import (
	"encoding/csv"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// Format selects the output format.
type Format string

const (
	FormatCSV  Format = "csv"
	FormatJSON Format = "json"
	FormatText Format = "txt"
)

// ParseFormat converts a string to a Format.
func ParseFormat(s string) Format {
	switch strings.ToLower(s) {
	case "json":
		return FormatJSON
	case "txt", "text":
		return FormatText
	default:
		return FormatCSV
	}
}

// MailResult is the universal record written to output files.
type MailResult struct {
	Mailbox   string `json:"mailbox"`
	Folder    string `json:"folder"`
	From      string `json:"from"`
	DateSent  string `json:"date_sent"`
	Subject   string `json:"subject"`
	Body      string `json:"body"`
	MatchedBy string `json:"matched_by"`
}

// SprayResult records a credential spray attempt.
type SprayResult struct {
	Username string `json:"username"`
	Password string `json:"password"`
	Valid    bool   `json:"valid"`
}

// GALResult records a single GAL entry.
type GALResult struct {
	DisplayName  string `json:"display_name"`
	EmailAddress string `json:"email_address"`
}

// UserResult records a harvested user entry.
type UserResult struct {
	Username   string `json:"username"`
	Valid      bool   `json:"valid"`
	ADUsername string `json:"ad_username,omitempty"`
}

// FolderResult records a mailbox folder.
type FolderResult struct {
	ID               string `json:"id"`
	DisplayName      string `json:"display_name"`
	TotalCount       int    `json:"total_count"`
	UnreadCount      int    `json:"unread_count"`
	ChildFolderCount int    `json:"child_folder_count"`
}

// WriteMailResults writes MailResult records to a file.
func WriteMailResults(path string, format Format, records []MailResult) error {
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		return err
	}
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()

	switch format {
	case FormatJSON:
		enc := json.NewEncoder(f)
		enc.SetIndent("", "  ")
		return enc.Encode(records)
	case FormatText:
		for _, r := range records {
			fmt.Fprintf(f, "[%s] [%s] From: %s | Subject: %s | Match: %s\n",
				r.Mailbox, r.DateSent, r.From, r.Subject, r.MatchedBy)
		}
		return nil
	default: // CSV
		w := csv.NewWriter(f)
		_ = w.Write([]string{"Mailbox", "Folder", "From", "DateSent", "Subject", "Body", "MatchedBy"})
		for _, r := range records {
			_ = w.Write([]string{r.Mailbox, r.Folder, r.From, r.DateSent, r.Subject, r.Body, r.MatchedBy})
		}
		w.Flush()
		return w.Error()
	}
}

// WriteSprayResults writes credential spray results to a file.
func WriteSprayResults(path string, format Format, records []SprayResult) error {
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		return err
	}
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()

	switch format {
	case FormatJSON:
		enc := json.NewEncoder(f)
		enc.SetIndent("", "  ")
		return enc.Encode(records)
	case FormatText:
		for _, r := range records {
			if r.Valid {
				fmt.Fprintf(f, "[+] %s:%s\n", r.Username, r.Password)
			}
		}
		return nil
	default:
		w := csv.NewWriter(f)
		_ = w.Write([]string{"Username", "Password", "Valid"})
		for _, r := range records {
			valid := "false"
			if r.Valid {
				valid = "true"
			}
			_ = w.Write([]string{r.Username, r.Password, valid})
		}
		w.Flush()
		return w.Error()
	}
}

// WriteGALResults writes Global Address List results to a file.
func WriteGALResults(path string, format Format, records []GALResult) error {
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		return err
	}
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()

	switch format {
	case FormatJSON:
		enc := json.NewEncoder(f)
		enc.SetIndent("", "  ")
		return enc.Encode(records)
	case FormatText:
		for _, r := range records {
			fmt.Fprintf(f, "%s <%s>\n", r.DisplayName, r.EmailAddress)
		}
		return nil
	default:
		w := csv.NewWriter(f)
		_ = w.Write([]string{"DisplayName", "EmailAddress"})
		for _, r := range records {
			_ = w.Write([]string{r.DisplayName, r.EmailAddress})
		}
		w.Flush()
		return w.Error()
	}
}

// WriteUserResults writes user harvest results to a file.
func WriteUserResults(path string, format Format, records []UserResult) error {
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		return err
	}
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()

	switch format {
	case FormatJSON:
		enc := json.NewEncoder(f)
		enc.SetIndent("", "  ")
		return enc.Encode(records)
	case FormatText:
		for _, r := range records {
			status := "[-]"
			if r.Valid {
				status = "[+]"
			}
			fmt.Fprintf(f, "%s %s\n", status, r.Username)
		}
		return nil
	default:
		w := csv.NewWriter(f)
		_ = w.Write([]string{"Username", "Valid", "ADUsername"})
		for _, r := range records {
			valid := "false"
			if r.Valid {
				valid = "true"
			}
			_ = w.Write([]string{r.Username, valid, r.ADUsername})
		}
		w.Flush()
		return w.Error()
	}
}

// WriteFolderResults writes folder listing results to a file.
func WriteFolderResults(path string, format Format, records []FolderResult) error {
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		return err
	}
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()

	switch format {
	case FormatJSON:
		enc := json.NewEncoder(f)
		enc.SetIndent("", "  ")
		return enc.Encode(records)
	case FormatText:
		for _, r := range records {
			fmt.Fprintf(f, "%-40s  Total: %5d  Unread: %4d\n",
				r.DisplayName, r.TotalCount, r.UnreadCount)
		}
		return nil
	default:
		w := csv.NewWriter(f)
		_ = w.Write([]string{"ID", "DisplayName", "TotalCount", "UnreadCount", "ChildFolderCount"})
		for _, r := range records {
			_ = w.Write([]string{
				r.ID, r.DisplayName,
				fmt.Sprintf("%d", r.TotalCount),
				fmt.Sprintf("%d", r.UnreadCount),
				fmt.Sprintf("%d", r.ChildFolderCount),
			})
		}
		w.Flush()
		return w.Error()
	}
}
