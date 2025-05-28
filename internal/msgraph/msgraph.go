package msgraph

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/mail"
	"strings"
	"time"

	"github.com/axllent/mailpit/internal/logger"
)

const (
	// GraphAPIEndpoint is the base URL for Microsoft Graph API
	GraphAPIEndpoint = "https://graph.microsoft.com/v1.0"

	// TokenEndpoint is the URL for obtaining OAuth tokens
	TokenEndpoint = "https://login.microsoftonline.com/%s/oauth2/v2.0/token"
)

// Config holds the configuration for Microsoft Graph API
type Config struct {
	TenantID     string
	ClientID     string
	ClientSecret string
	Enabled      bool
}

// TokenResponse represents the OAuth token response
type TokenResponse struct {
	AccessToken  string `json:"access_token"`
	TokenType    string `json:"token_type"`
	ExpiresIn    int    `json:"expires_in"`
	RefreshToken string `json:"refresh_token,omitempty"`
}

// Client represents a Microsoft Graph API client
type Client struct {
	Config      Config
	AccessToken string
	ExpiresAt   time.Time
}

// NewClient creates a new Microsoft Graph client
func NewClient(config Config) *Client {
	return &Client{
		Config: config,
	}
}

// GetToken obtains a new access token
func (c *Client) GetToken() error {
	// Check if we already have a valid token
	if c.AccessToken != "" && time.Now().Before(c.ExpiresAt) {
		return nil
	}

	// Prepare the token request
	data := strings.NewReader(fmt.Sprintf(
		"grant_type=client_credentials&client_id=%s&client_secret=%s&scope=https://graph.microsoft.com/.default",
		c.Config.ClientID,
		c.Config.ClientSecret,
	))

	// Create the request
	url := fmt.Sprintf(TokenEndpoint, c.Config.TenantID)
	req, err := http.NewRequest("POST", url, data)
	if err != nil {
		return fmt.Errorf("error creating token request: %v", err)
	}

	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")

	// Send the request
	client := &http.Client{}
	resp, err := client.Do(req)
	if err != nil {
		return fmt.Errorf("error sending token request: %v", err)
	}
	defer resp.Body.Close()

	// Read the response
	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return fmt.Errorf("error reading token response: %v", err)
	}

	// Check for error response
	if resp.StatusCode != http.StatusOK {
		return fmt.Errorf("error getting token: %s", string(body))
	}

	// Parse the token response
	var tokenResp TokenResponse
	if err := json.Unmarshal(body, &tokenResp); err != nil {
		return fmt.Errorf("error parsing token response: %v", err)
	}

	// Store the token
	c.AccessToken = tokenResp.AccessToken
	c.ExpiresAt = time.Now().Add(time.Duration(tokenResp.ExpiresIn-300) * time.Second) // Subtract 5 minutes for safety

	return nil
}

// SendMail sends an email using the Microsoft Graph API
func (c *Client) SendMail(from string, to []string, rawMessage []byte) error {
	// Get a valid token
	if err := c.GetToken(); err != nil {
		return err
	}

	// Parse the raw email message
	reader := bytes.NewReader(rawMessage)
	msg, err := mail.ReadMessage(reader)
	if err != nil {
		return fmt.Errorf("error parsing email message: %v", err)
	}

	// Extract subject
	subject := msg.Header.Get("Subject")

	// Extract message body
	body, err := io.ReadAll(msg.Body)
	if err != nil {
		return fmt.Errorf("error reading message body: %v", err)
	}

	// Determine content type
	contentType := msg.Header.Get("Content-Type")
	isHTML := strings.Contains(strings.ToLower(contentType), "text/html")

	// Prepare the message request
	type Recipient struct {
		EmailAddress struct {
			Address string `json:"address"`
		} `json:"emailAddress"`
	}

	type Message struct {
		Subject      string      `json:"subject"`
		Body         interface{} `json:"body"`
		ToRecipients []Recipient `json:"toRecipients"`
	}

	type SendMailRequest struct {
		Message            Message `json:"message"`
		SaveToSentItems   bool    `json:"saveToSentItems"`
	}

	// Create recipients
	recipients := make([]Recipient, len(to))
	for i, addr := range to {
		recipients[i] = Recipient{
			EmailAddress: struct {
				Address string `json:"address"`
			}{
				Address: addr,
			},
		}
	}

	// Create message body based on content type
	var messageBody interface{}
	if isHTML {
		messageBody = struct {
			ContentType string `json:"contentType"`
			Content     string `json:"content"`
		}{
			ContentType: "html",
			Content:     string(body),
		}
	} else {
		messageBody = struct {
			ContentType string `json:"contentType"`
			Content     string `json:"content"`
		}{
			ContentType: "text",
			Content:     string(body),
		}
	}

	// Create the request payload
	sendMailReq := SendMailRequest{
		Message: Message{
			Subject:      subject,
			Body:         messageBody,
			ToRecipients: recipients,
		},
		SaveToSentItems: true,
	}

	// Convert to JSON
	jsonData, err := json.Marshal(sendMailReq)
	if err != nil {
		return fmt.Errorf("error marshaling request: %v", err)
	}

	// Create the HTTP request
	url := fmt.Sprintf("%s/users/%s/sendMail", GraphAPIEndpoint, from)
	req, err := http.NewRequest("POST", url, bytes.NewBuffer(jsonData))
	if err != nil {
		return fmt.Errorf("error creating request: %v", err)
	}

	req.Header.Set("Authorization", "Bearer "+c.AccessToken)
	req.Header.Set("Content-Type", "application/json")

	// Send the request
	client := &http.Client{}
	resp, err := client.Do(req)
	if err != nil {
		return fmt.Errorf("error sending request: %v", err)
	}
	defer resp.Body.Close()

	// Check for error response
	if resp.StatusCode != http.StatusAccepted && resp.StatusCode != http.StatusOK {
		respBody, _ := io.ReadAll(resp.Body)
		return fmt.Errorf("error sending mail: %s", string(respBody))
	}

	logger.Log().Debugf("[msgraph] sent message to %s", strings.Join(to, ", "))
	return nil
}

// SendMailFromRaw sends an email using the Microsoft Graph API from a raw email message
func SendMailFromRaw(config Config, from string, to []string, rawMessage []byte) error {
	if !config.Enabled {
		return fmt.Errorf("Microsoft Graph API is not enabled")
	}

	client := NewClient(config)
	return client.SendMail(from, to, rawMessage)
}
