package ews

import (
	"encoding/xml"
	"fmt"
	"strings"
)

// -------- Generic response-code helpers --------

// checkResponseCode parses the first ResponseCode found in an EWS SOAP reply
// and returns an error if it is not "NoError".
// Exchange returns elements with namespace prefixes (e.g. <m:ResponseCode>)
// so we search for ">ResponseCode</" pattern to be prefix-agnostic.
func checkResponseCode(raw []byte) error {
	s := string(raw)

	// Match both <ResponseCode> and <m:ResponseCode> (or any prefix).
	codeEnd := strings.Index(s, ">ResponseCode>")
	if codeEnd == -1 {
		// Try namespace-prefixed variant like <m:ResponseCode>value</m:ResponseCode>
		// by finding the closing tag first.
		codeEnd = strings.Index(s, "ResponseCode>")
		if codeEnd == -1 {
			return nil // no ResponseCode element — assume success
		}
	}
	// Find the closing > that terminates the opening tag.
	valueStart := strings.Index(s, "ResponseCode>")
	if valueStart == -1 {
		return nil
	}
	valueStart += len("ResponseCode>")
	valueEnd := strings.Index(s[valueStart:], "<")
	if valueEnd == -1 {
		return nil
	}
	code := s[valueStart : valueStart+valueEnd]
	if code == "NoError" {
		return nil
	}
	// Also extract MessageText for a useful error message.
	var msgText string
	mtStart := strings.Index(s, "MessageText>")
	if mtStart != -1 {
		mtStart += len("MessageText>")
		mtEnd := strings.Index(s[mtStart:], "<")
		if mtEnd != -1 {
			msgText = s[mtStart : mtStart+mtEnd]
		}
	}
	if msgText != "" {
		return fmt.Errorf("EWS %s: %s", code, msgText)
	}
	return fmt.Errorf("EWS %s", code)
}

// MailItem represents a single email result.
type MailItem struct {
	Subject    string
	Body       string
	From       string
	DateSent   string
	Mailbox    string
	FolderName string
	Matched    string
}

// FindItemsResult holds results from a FindItem call.
type FindItemsResult struct {
	Items []MailItem
}

// -------- FindItem (list messages) --------

type findItemResponse struct {
	XMLName xml.Name `xml:"Envelope"`
	Body    struct {
		FindItemResponse struct {
			ResponseMessages struct {
				FindItemResponseMessage struct {
					RootFolder struct {
						Items struct {
							Messages []struct {
								ItemID struct {
									ID        string `xml:"Id,attr"`
									ChangeKey string `xml:"ChangeKey,attr"`
								} `xml:"ItemId"`
								Subject  string `xml:"Subject"`
								DateSent string `xml:"DateTimeSent"`
								From     struct {
									Mailbox struct {
										EmailAddress string `xml:"EmailAddress"`
									} `xml:"Mailbox"`
								} `xml:"From"`
							} `xml:"Message"`
						} `xml:"Items"`
					} `xml:"RootFolder"`
				} `xml:"FindItemResponseMessage"`
			} `xml:"ResponseMessages"`
		} `xml:"FindItemResponse"`
	} `xml:"Body"`
}

// FindItems returns message summaries from a folder (default Inbox).
func (c *Client) FindItems(mailbox, folder string, maxItems int) ([]struct {
	ID        string
	ChangeKey string
	Subject   string
	DateSent  string
	From      string
}, error) {
	if folder == "" {
		folder = "inbox"
	}
	body := fmt.Sprintf(`<m:FindItem Traversal="Shallow">
  <m:ItemShape>
    <t:BaseShape>AllProperties</t:BaseShape>
  </m:ItemShape>
  <m:IndexedPageItemView MaxEntriesReturned="%d" Offset="0" BasePoint="Beginning"/>
  <m:ParentFolderIds>
    <t:DistinguishedFolderId Id="%s">
      <t:Mailbox>
        <t:EmailAddress>%s</t:EmailAddress>
      </t:Mailbox>
    </t:DistinguishedFolderId>
  </m:ParentFolderIds>
</m:FindItem>`, maxItems, strings.ToLower(folder), mailbox)

	raw, err := c.Do(body)
	if err != nil {
		return nil, err
	}
	if err := checkResponseCode(raw); err != nil {
		return nil, err
	}

	var resp findItemResponse
	if err := xml.Unmarshal(raw, &resp); err != nil {
		return nil, fmt.Errorf("findItem unmarshal: %w", err)
	}

	msgs := resp.Body.FindItemResponse.ResponseMessages.FindItemResponseMessage.RootFolder.Items.Messages
	out := make([]struct {
		ID        string
		ChangeKey string
		Subject   string
		DateSent  string
		From      string
	}, 0, len(msgs))
	for _, m := range msgs {
		out = append(out, struct {
			ID        string
			ChangeKey string
			Subject   string
			DateSent  string
			From      string
		}{
			ID:        m.ItemID.ID,
			ChangeKey: m.ItemID.ChangeKey,
			Subject:   m.Subject,
			DateSent:  m.DateSent,
			From:      m.From.Mailbox.EmailAddress,
		})
	}
	return out, nil
}

// -------- GetItem (fetch body) --------

type getItemResponse struct {
	XMLName xml.Name `xml:"Envelope"`
	Body    struct {
		GetItemResponse struct {
			ResponseMessages struct {
				GetItemResponseMessage []struct {
					Items struct {
						Message struct {
							Body struct {
								Content string `xml:",chardata"`
							} `xml:"Body"`
							Subject  string `xml:"Subject"`
							DateSent string `xml:"DateTimeSent"`
							From     struct {
								Mailbox struct {
									EmailAddress string `xml:"EmailAddress"`
								} `xml:"Mailbox"`
							} `xml:"From"`
							Attachments struct {
								FileAttachment []struct {
									Name    string `xml:"Name"`
									Content string `xml:"Content"`
								} `xml:"FileAttachment"`
							} `xml:"Attachments"`
						} `xml:"Message"`
					} `xml:"Items"`
				} `xml:"GetItemResponseMessage"`
			} `xml:"ResponseMessages"`
		} `xml:"GetItemResponse"`
	} `xml:"Body"`
}

// GetItem fetches the full body of a message by ItemID.
func (c *Client) GetItem(id, changeKey string) (*struct {
	Subject     string
	Body        string
	DateSent    string
	From        string
	Attachments []struct{ Name, Content string }
}, error) {
	body := fmt.Sprintf(`<m:GetItem>
  <m:ItemShape>
    <t:BaseShape>AllProperties</t:BaseShape>
    <t:IncludeMimeContent>false</t:IncludeMimeContent>
    <t:BodyType>Text</t:BodyType>
    <t:AdditionalProperties>
      <t:FieldURI FieldURI="item:Attachments"/>
    </t:AdditionalProperties>
  </m:ItemShape>
  <m:ItemIds>
    <t:ItemId Id="%s" ChangeKey="%s"/>
  </m:ItemIds>
</m:GetItem>`, id, changeKey)

	raw, err := c.Do(body)
	if err != nil {
		return nil, err
	}

	var resp getItemResponse
	if err := xml.Unmarshal(raw, &resp); err != nil {
		return nil, fmt.Errorf("getItem unmarshal: %w", err)
	}

	msgs := resp.Body.GetItemResponse.ResponseMessages.GetItemResponseMessage
	if len(msgs) == 0 {
		return nil, fmt.Errorf("no message returned")
	}
	m := msgs[0].Items.Message
	result := &struct {
		Subject     string
		Body        string
		DateSent    string
		From        string
		Attachments []struct{ Name, Content string }
	}{
		Subject:  m.Subject,
		Body:     m.Body.Content,
		DateSent: m.DateSent,
		From:     m.From.Mailbox.EmailAddress,
	}
	for _, a := range m.Attachments.FileAttachment {
		result.Attachments = append(result.Attachments, struct{ Name, Content string }{a.Name, a.Content})
	}
	return result, nil
}

// -------- GetFolder (list all folders) --------

type getFolderResponse struct {
	XMLName xml.Name `xml:"Envelope"`
	Body    struct {
		FindFolderResponse struct {
			ResponseMessages struct {
				FindFolderResponseMessage struct {
					RootFolder struct {
						Folders struct {
							Folders []struct {
								FolderID struct {
									ID string `xml:"Id,attr"`
								} `xml:"FolderId"`
								DisplayName      string `xml:"DisplayName"`
								TotalCount       int    `xml:"TotalCount"`
								UnreadCount      int    `xml:"UnreadCount"`
								ChildFolderCount int    `xml:"ChildFolderCount"`
							} `xml:"Folder"`
						} `xml:"Folders"`
					} `xml:"RootFolder"`
				} `xml:"FindFolderResponseMessage"`
			} `xml:"ResponseMessages"`
		} `xml:"FindFolderResponse"`
	} `xml:"Body"`
}

// FolderInfo holds brief folder metadata.
type FolderInfo struct {
	ID               string
	DisplayName      string
	TotalCount       int
	UnreadCount      int
	ChildFolderCount int
}

// FindFolders returns all folders in the given mailbox.
func (c *Client) FindFolders(mailbox string) ([]FolderInfo, error) {
	body := fmt.Sprintf(`<m:FindFolder Traversal="Deep">
  <m:FolderShape>
    <t:BaseShape>AllProperties</t:BaseShape>
  </m:FolderShape>
  <m:ParentFolderIds>
    <t:DistinguishedFolderId Id="root">
      <t:Mailbox>
        <t:EmailAddress>%s</t:EmailAddress>
      </t:Mailbox>
    </t:DistinguishedFolderId>
  </m:ParentFolderIds>
</m:FindFolder>`, mailbox)

	raw, err := c.Do(body)
	if err != nil {
		return nil, err
	}

	var resp getFolderResponse
	if err := xml.Unmarshal(raw, &resp); err != nil {
		return nil, fmt.Errorf("findFolder unmarshal: %w", err)
	}

	folders := resp.Body.FindFolderResponse.ResponseMessages.FindFolderResponseMessage.RootFolder.Folders.Folders
	out := make([]FolderInfo, 0, len(folders))
	for _, f := range folders {
		out = append(out, FolderInfo{
			ID:               f.FolderID.ID,
			DisplayName:      f.DisplayName,
			TotalCount:       f.TotalCount,
			UnreadCount:      f.UnreadCount,
			ChildFolderCount: f.ChildFolderCount,
		})
	}
	return out, nil
}

// -------- ResolveNames (AD username lookup) --------

type resolveNamesResponse struct {
	XMLName xml.Name `xml:"Envelope"`
	Body    struct {
		ResolveNamesResponse struct {
			ResponseMessages struct {
				ResolveNamesResponseMessage struct {
					ResolutionSet struct {
						Resolution []struct {
							Mailbox struct {
								Name         string `xml:"Name"`
								EmailAddress string `xml:"EmailAddress"`
								RoutingType  string `xml:"RoutingType"`
							} `xml:"Mailbox"`
							Contact struct {
								AssistantName    string `xml:"AssistantName"`
								BusinessHomePage string `xml:"BusinessHomePage"`
							} `xml:"Contact"`
						} `xml:"Resolution"`
					} `xml:"ResolutionSet"`
				} `xml:"ResolveNamesResponseMessage"`
			} `xml:"ResponseMessages"`
		} `xml:"ResolveNamesResponse"`
	} `xml:"Body"`
}

// ResolveNames resolves an email address to its AD display name / username.
func (c *Client) ResolveNames(email string) (string, error) {
	body := fmt.Sprintf(`<m:ResolveNames ReturnFullContactData="true" SearchScope="ActiveDirectory">
  <m:UnresolvedEntry>%s</m:UnresolvedEntry>
</m:ResolveNames>`, email)

	raw, err := c.Do(body)
	if err != nil {
		return "", err
	}

	var resp resolveNamesResponse
	if err := xml.Unmarshal(raw, &resp); err != nil {
		return "", fmt.Errorf("resolveNames unmarshal: %w", err)
	}

	resolutions := resp.Body.ResolveNamesResponse.ResponseMessages.ResolveNamesResponseMessage.ResolutionSet.Resolution
	if len(resolutions) == 0 {
		return "", nil
	}
	return resolutions[0].Mailbox.Name, nil
}

// -------- GetGAL via EWS FindPeople --------

type findPeopleResponse struct {
	XMLName xml.Name `xml:"Envelope"`
	Body    struct {
		FindPeopleResponse struct {
			People struct {
				Persona []struct {
					EmailAddress struct {
						EmailAddress string `xml:"EmailAddress"`
					} `xml:"EmailAddress"`
					DisplayName string `xml:"DisplayName"`
				} `xml:"Persona"`
			} `xml:"People"`
			TotalNumberOfPeopleInView int `xml:"TotalNumberOfPeopleInView"`
		} `xml:"FindPeopleResponse"`
	} `xml:"Body"`
}

// GALEntry is a single GAL record.
type GALEntry struct {
	DisplayName  string
	EmailAddress string
}

// GetGALEWS retrieves the Global Address List entries via EWS FindPeople.
func (c *Client) GetGALEWS(maxResults int) ([]GALEntry, error) {
	var all []GALEntry
	offset := 0
	pageSize := 100
	if maxResults > 0 && pageSize > maxResults {
		pageSize = maxResults
	}

	for {
		body := fmt.Sprintf(`<m:FindPeople>
  <m:PersonaShape>
    <t:BaseShape>AllProperties</t:BaseShape>
  </m:PersonaShape>
  <m:IndexedPageItemView MaxEntriesReturned="%d" Offset="%d" BasePoint="Beginning"/>
  <m:ParentFolderId>
    <t:DistinguishedFolderId Id="directory"/>
  </m:ParentFolderId>
</m:FindPeople>`, pageSize, offset)

		raw, err := c.Do(body)
		if err != nil {
			return all, err
		}

		var resp findPeopleResponse
		if err := xml.Unmarshal(raw, &resp); err != nil {
			return all, fmt.Errorf("findPeople unmarshal: %w", err)
		}

		personas := resp.Body.FindPeopleResponse.People.Persona
		for _, p := range personas {
			all = append(all, GALEntry{
				DisplayName:  p.DisplayName,
				EmailAddress: p.EmailAddress.EmailAddress,
			})
		}

		offset += len(personas)
		total := resp.Body.FindPeopleResponse.TotalNumberOfPeopleInView
		if len(personas) == 0 || offset >= total {
			break
		}
		if maxResults > 0 && len(all) >= maxResults {
			break
		}
	}
	return all, nil
}

// -------- SendItem --------

// SendEmail sends an email via EWS CreateItem.
func (c *Client) SendEmail(to, subject, emailBody string) error {
	body := fmt.Sprintf(`<m:CreateItem MessageDisposition="SendAndSaveCopy">
  <m:SavedItemFolderId>
    <t:DistinguishedFolderId Id="sentitems"/>
  </m:SavedItemFolderId>
  <m:Items>
    <t:Message>
      <t:Subject>%s</t:Subject>
      <t:Body BodyType="HTML">%s</t:Body>
      <t:ToRecipients>
        <t:Mailbox>
          <t:EmailAddress>%s</t:EmailAddress>
        </t:Mailbox>
      </t:ToRecipients>
    </t:Message>
  </m:Items>
</m:CreateItem>`, xmlEscape(subject), xmlEscape(emailBody), xmlEscape(to))

	_, err := c.Do(body)
	return err
}

func xmlEscape(s string) string {
	s = strings.ReplaceAll(s, "&", "&amp;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	s = strings.ReplaceAll(s, "\"", "&quot;")
	return s
}

// -------- FindItemsAllFolders --------

// FindItemsAllFolders returns message summaries from every folder in the mailbox.
func (c *Client) FindItemsAllFolders(mailbox string, maxPerFolder int) ([]struct {
	ID         string
	ChangeKey  string
	Subject    string
	DateSent   string
	From       string
	FolderName string
}, error) {
	folders, err := c.FindFolders(mailbox)
	if err != nil {
		return nil, err
	}

	var all []struct {
		ID         string
		ChangeKey  string
		Subject    string
		DateSent   string
		From       string
		FolderName string
	}

	for _, folder := range folders {
		items, err := c.FindItemsByFolderID(mailbox, folder.ID, maxPerFolder)
		if err != nil {
			continue
		}
		for _, item := range items {
			all = append(all, struct {
				ID         string
				ChangeKey  string
				Subject    string
				DateSent   string
				From       string
				FolderName string
			}{
				ID:         item.ID,
				ChangeKey:  item.ChangeKey,
				Subject:    item.Subject,
				DateSent:   item.DateSent,
				From:       item.From,
				FolderName: folder.DisplayName,
			})
		}
	}
	return all, nil
}

type findItemByIDResponse struct {
	XMLName xml.Name `xml:"Envelope"`
	Body    struct {
		FindItemResponse struct {
			ResponseMessages struct {
				FindItemResponseMessage struct {
					RootFolder struct {
						Items struct {
							Messages []struct {
								ItemID struct {
									ID        string `xml:"Id,attr"`
									ChangeKey string `xml:"ChangeKey,attr"`
								} `xml:"ItemId"`
								Subject  string `xml:"Subject"`
								DateSent string `xml:"DateTimeSent"`
								From     struct {
									Mailbox struct {
										EmailAddress string `xml:"EmailAddress"`
									} `xml:"Mailbox"`
								} `xml:"From"`
							} `xml:"Message"`
						} `xml:"Items"`
					} `xml:"RootFolder"`
				} `xml:"FindItemResponseMessage"`
			} `xml:"ResponseMessages"`
		} `xml:"FindItemResponse"`
	} `xml:"Body"`
}

// FindItemsByFolderID returns message summaries from a folder addressed by its ID.
func (c *Client) FindItemsByFolderID(mailbox, folderID string, maxItems int) ([]struct {
	ID        string
	ChangeKey string
	Subject   string
	DateSent  string
	From      string
}, error) {
	body := fmt.Sprintf(`<m:FindItem Traversal="Shallow">
  <m:ItemShape>
    <t:BaseShape>AllProperties</t:BaseShape>
  </m:ItemShape>
  <m:IndexedPageItemView MaxEntriesReturned="%d" Offset="0" BasePoint="Beginning"/>
  <m:ParentFolderIds>
    <t:FolderId Id="%s">
      <t:Mailbox>
        <t:EmailAddress>%s</t:EmailAddress>
      </t:Mailbox>
    </t:FolderId>
  </m:ParentFolderIds>
</m:FindItem>`, maxItems, folderID, mailbox)

	raw, err := c.Do(body)
	if err != nil {
		return nil, err
	}

	var resp findItemByIDResponse
	if err := xml.Unmarshal(raw, &resp); err != nil {
		return nil, fmt.Errorf("findItemByID unmarshal: %w", err)
	}

	msgs := resp.Body.FindItemResponse.ResponseMessages.FindItemResponseMessage.RootFolder.Items.Messages
	out := make([]struct {
		ID        string
		ChangeKey string
		Subject   string
		DateSent  string
		From      string
	}, 0, len(msgs))
	for _, m := range msgs {
		out = append(out, struct {
			ID        string
			ChangeKey string
			Subject   string
			DateSent  string
			From      string
		}{
			ID:        m.ItemID.ID,
			ChangeKey: m.ItemID.ChangeKey,
			Subject:   m.Subject,
			DateSent:  m.DateSent,
			From:      m.From.Mailbox.EmailAddress,
		})
	}
	return out, nil
}
