package ews

import (
	"encoding/xml"
	"errors"
	"time"
)

type CreateItem struct {
	XMLName                struct{}          `xml:"m:CreateItem"`
	MessageDisposition     string            `xml:"MessageDisposition,attr"`
	SendMeetingInvitations string            `xml:"SendMeetingInvitations,attr"`
	SavedItemFolderId      SavedItemFolderId `xml:"m:SavedItemFolderId"`
	Items                  Items             `xml:"m:Items"`
}

type Items struct {
	Message      []Message      `xml:"t:Message"`
	CalendarItem []CalendarItem `xml:"t:CalendarItem"`
}

type SavedItemFolderId struct {
	DistinguishedFolderId DistinguishedFolderId `xml:"t:DistinguishedFolderId"`
}

type Message struct {
	ItemClass    string     `xml:"t:ItemClass"`
	Subject      string     `xml:"t:Subject"`
	Body         Body       `xml:"t:Body"`
	Sender       OneMailbox `xml:"t:Sender"`
	ToRecipients XMailbox   `xml:"t:ToRecipients"`
}

type CalendarItem struct {
	Subject                    string      `xml:"t:Subject,omitempty"`
	Body                       Body        `xml:"t:Body,omitempty"`
	ReminderIsSet              bool        `xml:"t:ReminderIsSet,omitempty"`
	ReminderMinutesBeforeStart int         `xml:"t:ReminderMinutesBeforeStart,omitempty"`
	Start                      time.Time   `xml:"t:Start,omitempty"`
	End                        time.Time   `xml:"t:End,omitempty"`
	IsAllDayEvent              bool        `xml:"t:IsAllDayEvent,omitempty"`
	LegacyFreeBusyStatus       string      `xml:"t:LegacyFreeBusyStatus,omitempty"`
	Location                   string      `xml:"t:Location,omitempty"`
	RequiredAttendees          []Attendees `xml:"t:RequiredAttendees,omitempty"`
	OptionalAttendees          []Attendees `xml:"t:OptionalAttendees,omitempty"`
	Resources                  []Attendees `xml:"t:Resources,omitempty"`
}

type Body struct {
	BodyType string `xml:"BodyType,attr"`
	Body     []byte `xml:",chardata"`
}

type OneMailbox struct {
	Mailbox Mailbox `xml:"t:Mailbox"`
}

type XMailbox struct {
	Mailbox []Mailbox `xml:"t:Mailbox"`
}

type Mailbox struct {
	EmailAddress string `xml:"t:EmailAddress"`
}

type Attendee struct {
	Mailbox Mailbox `xml:"t:Mailbox"`
}

type Attendees struct {
	Attendee []Attendee `xml:"t:Attendee"`
}

type createItemResponseBodyEnvelop struct {
	XMLName struct{}               `xml:"Envelope"`
	Body    createItemResponseBody `xml:"Body"`
}

type createItemResponseBody struct {
	CreateItemResponse ItemOperationResponse `xml:"CreateItemResponse"`
}

type ItemOperationResponse struct {
	ResponseMessages ResponseMessages `xml:"ResponseMessages"`
}

type ResponseMessages struct {
	CreateItemResponseMessage Response `xml:"CreateItemResponseMessage"`
	UpdateItemResponseMessage Response `xml:"UpdateItemResponseMessage"`
	DeleteItemResponseMessage Response `xml:"DeleteItemResponseMessage"`
}

// CreateMessageItem
// https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/createitem-operation-email-message
func CreateMessageItem(c Client, m ...Message) ([]ItemId, error) {

	item := &CreateItem{
		MessageDisposition: "SendAndSaveCopy",
		SavedItemFolderId:  SavedItemFolderId{DistinguishedFolderId{Id: "sentitems"}},
	}
	item.Items.Message = append(item.Items.Message, m...)

	xmlBytes, err := xml.MarshalIndent(item, "", "  ")
	if err != nil {
		return nil, err
	}

	bb, err := c.SendAndReceive(xmlBytes)
	if err != nil {
		return nil, err
	}

	items, err := checkCreateItemResponseForErrors(bb)
	if err != nil {
		return nil, err
	}

	return items.Message, nil
}

// CreateCalendarItem
// https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/createitem-operation-calendar-item
func CreateCalendarItem(c Client, ci ...CalendarItem) ([]ItemId, error) {

	item := &CreateItem{
		SendMeetingInvitations: "SendToAllAndSaveCopy",
		SavedItemFolderId:      SavedItemFolderId{DistinguishedFolderId{Id: "calendar"}},
	}
	item.Items.CalendarItem = append(item.Items.CalendarItem, ci...)

	xmlBytes, err := xml.MarshalIndent(item, "", "  ")
	if err != nil {
		return nil, err
	}

	bb, err := c.SendAndReceive(xmlBytes)
	if err != nil {
		return nil, err
	}

	items, err := checkCreateItemResponseForErrors(bb)
	if err != nil {
		return nil, err
	}

	return items.CalendarItem, nil
}

func checkCreateItemResponseForErrors(bb []byte) (items ResponseItems, err error) {
	var soapResp createItemResponseBodyEnvelop
	if err = xml.Unmarshal(bb, &soapResp); err != nil {
		return
	}

	resp := soapResp.Body.CreateItemResponse.ResponseMessages.CreateItemResponseMessage
	if resp.ResponseClass == ResponseClassError {
		err = errors.New(resp.MessageText)
		return
	}
	return resp.ResponseItems, nil
}
