package ews

import (
	"encoding/xml"
	"errors"
	"reflect"
	"strings"
)

type UpdateStrategy struct {
	ConflictResolution                    string `xml:"ConflictResolution,attr,omitempty"`
	MessageDisposition                    string `xml:"MessageDisposition,attr,omitempty"`
	SendMeetingInvitationsOrCancellations string `xml:"SendMeetingInvitationsOrCancellations,attr,omitempty"`
}

type UpdateItem struct {
	XMLName struct{} `xml:"m:UpdateItem"`

	ItemChanges ItemChanges `xml:"m:ItemChanges"`
	UpdateStrategy
}

type ItemChanges struct {
	XMLName xml.Name `xml:"m:ItemChanges"`

	ItemChanges []ItemChange
}

type ItemChange struct {
	XMLName xml.Name `xml:"t:ItemChange"`

	ItemId  ItemId  `xml:"t:ItemId"`
	Updates Updates `xml:"t:Updates"`
}

type Updates struct {
	XMLName xml.Name `xml:"t:Updates"`

	Updates []SetItemField `xml:"t:Updates"`
}

type SetItemField struct {
	XMLName xml.Name `xml:"t:SetItemField"`

	FieldURI     FieldURI `xml:"t:FieldURI"`
	CalendarItem []Field  `xml:"t:CalendarItem,omitempty"`
	Message      []Field  `xml:"t:Message,omitempty"`
}

type Field struct {
	Name       string
	Attributes map[string]string
	Value      interface{}
}

type updateItemResponseBodyEnvelop struct {
	XMLName struct{}               `xml:"Envelope"`
	Body    updateItemResponseBody `xml:"Body"`
}

type updateItemResponseBody struct {
	UpdateItemResponse ItemOperationResponse `xml:"UpdateItemResponse"`
}

func (f Field) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	e.EncodeToken(start)
	index, attr := 0, make([]xml.Attr, len(f.Attributes))
	for k, v := range f.Attributes {
		attr[index].Name = xml.Name{Local: k}
		attr[index].Value = v
	}
	var t = xml.StartElement{Name: xml.Name{Local: f.Name}, Attr: attr}
	// e.EncodeToken(t)
	e.EncodeElement(f.Value, t)
	// e.EncodeToken(xml.EndElement{Name: t.Name})
	e.EncodeToken(xml.EndElement{Name: start.Name})
	return e.Flush()
}

func getFields(obj interface{}) (fields []Field) {
	val := reflect.Indirect(reflect.ValueOf(obj))
	t := val.Type()
	length := t.NumField()
	for index := 0; index < length; index++ {
		value := val.Field(index)
		if !value.IsZero() {
			val, ok := t.Field(index).Tag.Lookup("xml")
			if !ok {
				val = t.Field(index).Name
			} else {
				val = strings.Split(val, ",")[0]
			}
			fields = append(fields, Field{
				Name:  val,
				Value: value.Interface(),
			})
		}
	}
	return
}

func (field Field) uri(replace string) string {
	index := strings.Index(field.Name, ":")
	if index != -1 {
		return replace + field.Name[index:]
	}
	return field.Name
}

var itemFields = map[string]bool{
	"t:Subject": true,
	"t:Body":    true,
}

func getSetItemField(prefix string, fields ...Field) []SetItemField {
	var setFields = make([]SetItemField, len(fields))
	for index, field := range fields {
		setFields[index].CalendarItem = []Field{field}
		replace := prefix
		if itemFields[field.Name] {
			replace = "item"
		}
		setFields[index].FieldURI = FieldURI{
			FieldURI: strings.Replace(field.Name, "t", replace, 1),
		}
	}
	return setFields
}

// UpdateCalendarItem
// https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-update-appointments-and-meetings-by-using-ews-in-exchange
func UpdateCalendarItem(c Client, id ItemId, ci CalendarItem, strategy ...UpdateStrategy) ([]ItemId, error) {

	strategy = append(strategy, UpdateStrategy{
		ConflictResolution:                    "AlwaysOverwrite",
		MessageDisposition:                    "SaveOnly",
		SendMeetingInvitationsOrCancellations: "SendToAllAndSaveCopy",
	})

	var setFields = getSetItemField("calendar", getFields(ci)...)

	item := UpdateItem{
		ItemChanges: ItemChanges{ItemChanges: []ItemChange{
			ItemChange{
				ItemId: id,
				Updates: Updates{
					Updates: setFields,
				},
			},
		}},
		UpdateStrategy: strategy[0],
	}

	xmlBytes, err := xml.MarshalIndent(item, "", "  ")
	if err != nil {
		return nil, err
	}

	bb, err := c.SendAndReceive(xmlBytes)
	if err != nil {
		return nil, err
	}

	items, err := checkUpdateItemResponseForErrors(bb)
	if err != nil {
		return nil, err
	}

	return items.CalendarItem, nil
}

func checkUpdateItemResponseForErrors(bb []byte) (items ResponseItems, err error) {
	var soapResp updateItemResponseBodyEnvelop
	if err = xml.Unmarshal(bb, &soapResp); err != nil {
		return
	}

	resp := soapResp.Body.UpdateItemResponse.ResponseMessages.UpdateItemResponseMessage
	if resp.ResponseClass == ResponseClassError {
		err = errors.New(resp.MessageText)
		return
	}
	return resp.ResponseItems, nil
}
