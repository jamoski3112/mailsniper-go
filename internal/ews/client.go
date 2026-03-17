// Package ews implements a minimal Exchange Web Services (EWS) SOAP client.
// It supports Basic, NTLM, and Bearer-token authentication and covers the
// EWS operations required by the MailSniper feature set.
package ews

import (
	"bytes"
	"crypto/tls"
	"encoding/base64"
	"fmt"
	"io"
	"net/http"
	"strings"
	"time"

	"github.com/Azure/go-ntlmssp"
)

// AuthType selects the authentication mode.
type AuthType int

const (
	AuthBasic  AuthType = iota // HTTP Basic (username:password)
	AuthNTLM                   // NTLM (via Authorization header negotiation)
	AuthBearer                 // OAuth2 Bearer token
)

// Client is a thin EWS SOAP client.
type Client struct {
	EWSURL      string
	Username    string
	Password    string
	AccessToken string
	AuthType    AuthType
	HTTPClient  *http.Client
	// ExchangeVersion is the RequestServerVersion SchemaType value, e.g. "Exchange2010".
	ExchangeVersion string
	// ImpersonateEmail, when set, adds an ExchangeImpersonation header to every request.
	ImpersonateEmail string
}

// NewClient creates a Client using NTLM auth (with Basic fallback).
// Exchange 2013+ on-prem typically advertises Negotiate/NTLM.
func NewClient(ewsURL, username, password, version string, skipTLS bool) *Client {
	return &Client{
		EWSURL:          ewsURL,
		Username:        username,
		Password:        password,
		AuthType:        AuthNTLM,
		ExchangeVersion: version,
		HTTPClient:      newNTLMHTTPClient(skipTLS),
	}
}

// NewBasicClient creates a Client explicitly using HTTP Basic auth.
func NewBasicClient(ewsURL, username, password, version string, skipTLS bool) *Client {
	return &Client{
		EWSURL:          ewsURL,
		Username:        username,
		Password:        password,
		AuthType:        AuthBasic,
		ExchangeVersion: version,
		HTTPClient:      newHTTPClient(skipTLS),
	}
}

// NewBearerClient creates a Client using an OAuth2 Bearer token.
func NewBearerClient(ewsURL, token, version string, skipTLS bool) *Client {
	return &Client{
		EWSURL:          ewsURL,
		AccessToken:     token,
		AuthType:        AuthBearer,
		ExchangeVersion: version,
		HTTPClient:      newHTTPClient(skipTLS),
	}
}

// newNTLMHTTPClient returns an http.Client whose transport handles the
// NTLM/Negotiate multi-step handshake transparently.
func newNTLMHTTPClient(skipTLS bool) *http.Client {
	tlsCfg := &tls.Config{InsecureSkipVerify: skipTLS} //nolint:gosec
	inner := &http.Transport{TLSClientConfig: tlsCfg}
	return &http.Client{
		Timeout:   30 * time.Second,
		Transport: ntlmssp.Negotiator{RoundTripper: inner},
	}
}

func newHTTPClient(skipTLS bool) *http.Client {
	tr := &http.Transport{
		TLSClientConfig: &tls.Config{InsecureSkipVerify: skipTLS}, //nolint:gosec
	}
	return &http.Client{
		Timeout:   30 * time.Second,
		Transport: tr,
		CheckRedirect: func(req *http.Request, via []*http.Request) error {
			// Preserve auth header across redirects.
			if len(via) > 0 {
				for key, val := range via[0].Header {
					if _, ok := req.Header[key]; !ok {
						req.Header[key] = val
					}
				}
			}
			return nil
		},
	}
}

// Do sends a raw SOAP body wrapped in the standard envelope and returns the response body bytes.
func (c *Client) Do(soapBody string) ([]byte, error) {
	envelope := c.buildEnvelope(soapBody)
	envBytes := []byte(envelope)

	req, err := http.NewRequest("POST", c.EWSURL, bytes.NewReader(envBytes))
	if err != nil {
		return nil, err
	}
	req.Header.Set("Content-Type", "text/xml; charset=utf-8")
	// Provide GetBody so go-ntlmssp can replay the body on the auth round-trip.
	req.GetBody = func() (io.ReadCloser, error) {
		return io.NopCloser(bytes.NewReader(envBytes)), nil
	}
	c.applyAuth(req)

	resp, err := c.HTTPClient.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return nil, err
	}
	if resp.StatusCode >= 400 {
		return nil, fmt.Errorf("EWS HTTP %d: %s", resp.StatusCode, truncate(string(body), 300))
	}
	return body, nil
}

func (c *Client) applyAuth(req *http.Request) {
	switch c.AuthType {
	case AuthBasic:
		encoded := base64.StdEncoding.EncodeToString([]byte(c.Username + ":" + c.Password))
		req.Header.Set("Authorization", "Basic "+encoded)
	case AuthBearer:
		req.Header.Set("Authorization", "Bearer "+c.AccessToken)
	case AuthNTLM:
		// go-ntlmssp reads credentials from a Basic Authorization header on the
		// first request, then transparently upgrades to NTLM Negotiate/Challenge/Auth.
		encoded := base64.StdEncoding.EncodeToString([]byte(c.Username + ":" + c.Password))
		req.Header.Set("Authorization", "Basic "+encoded)
	}
}

func (c *Client) buildEnvelope(body string) string {
	var impersonation string
	if c.ImpersonateEmail != "" {
		impersonation = fmt.Sprintf(`<soap:Header>
    <t:ExchangeImpersonation>
      <t:ConnectingSID>
        <t:PrimarySmtpAddress>%s</t:PrimarySmtpAddress>
      </t:ConnectingSID>
    </t:ExchangeImpersonation>
    <t:RequestServerVersion Version="%s"/>
  </soap:Header>`, c.ImpersonateEmail, c.ExchangeVersion)
	} else {
		impersonation = fmt.Sprintf(`<soap:Header>
    <t:RequestServerVersion Version="%s"/>
  </soap:Header>`, c.ExchangeVersion)
	}

	return fmt.Sprintf(`<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  %s
  <soap:Body>
    %s
  </soap:Body>
</soap:Envelope>`, impersonation, body)
}

// Autodiscover attempts to find the EWS endpoint for the given email address.
// It tries the standard autodiscover URL patterns.
func Autodiscover(email, password string, skipTLS bool) (string, error) {
	parts := strings.SplitN(email, "@", 2)
	if len(parts) != 2 {
		return "", fmt.Errorf("invalid email address: %s", email)
	}
	domain := parts[1]

	candidates := []string{
		fmt.Sprintf("https://autodiscover.%s/autodiscover/autodiscover.xml", domain),
		fmt.Sprintf("https://%s/autodiscover/autodiscover.xml", domain),
	}

	httpClient := newHTTPClient(skipTLS)
	reqBody := fmt.Sprintf(`<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006">
  <Request>
    <EMailAddress>%s</EMailAddress>
    <AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>
  </Request>
</Autodiscover>`, email)

	for _, url := range candidates {
		req, err := http.NewRequest("POST", url, strings.NewReader(reqBody))
		if err != nil {
			continue
		}
		req.Header.Set("Content-Type", "text/xml")
		encoded := base64.StdEncoding.EncodeToString([]byte(email + ":" + password))
		req.Header.Set("Authorization", "Basic "+encoded)

		resp, err := httpClient.Do(req)
		if err != nil {
			continue
		}
		defer resp.Body.Close()
		b, _ := io.ReadAll(resp.Body)
		if resp.StatusCode == 200 {
			ewsURL := extractXMLValue(string(b), "EwsUrl")
			if ewsURL != "" {
				return ewsURL, nil
			}
		}
	}
	return "", fmt.Errorf("autodiscover failed for %s", email)
}

// BuildEWSURL constructs a default EWS URL from a hostname.
func BuildEWSURL(hostname string) string {
	hostname = strings.TrimPrefix(hostname, "https://")
	hostname = strings.TrimPrefix(hostname, "http://")
	hostname = strings.TrimSuffix(hostname, "/")
	return fmt.Sprintf("https://%s/EWS/Exchange.asmx", hostname)
}

// truncate shortens a string for error messages.
func truncate(s string, n int) string {
	if len(s) <= n {
		return s
	}
	return s[:n] + "..."
}

// extractXMLValue is a quick-and-dirty XML value extractor (avoids heavy imports).
func extractXMLValue(body, tag string) string {
	open := "<" + tag + ">"
	close := "</" + tag + ">"
	start := strings.Index(body, open)
	if start == -1 {
		return ""
	}
	start += len(open)
	end := strings.Index(body[start:], close)
	if end == -1 {
		return ""
	}
	return body[start : start+end]
}
