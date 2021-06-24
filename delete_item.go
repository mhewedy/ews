package ews

import (
	"encoding/xml"
	"errors"
)

type DeleteStrategy struct {
	DeleteType               string `xml:"DeleteType,attr,omitempty"`
	SendMeetingCancellations string `xml:"SendMeetingCancellations,attr,omitempty"`
}

type DeleteItem struct {
	XMLName struct{} `xml:"m:DeleteItem"`
	ItemIds ItemIds  `xml:"m:ItemIds"`

	DeleteStrategy
}

type ItemIds struct {
	XMLName struct{} `xml:"m:ItemIds"`

	ItemId []ItemId `xml:"t:ItemId"`
}

type deleteItemResponseBodyEnvelop struct {
	XMLName struct{}               `xml:"Envelope"`
	Body    deleteItemResponseBody `xml:"Body"`
}

type deleteItemResponseBody struct {
	DeleteItemResponse ItemOperationResponse `xml:"DeleteItemResponse"`
}

// DeleteItems
// https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-delete-appointments-and-cancel-meetings-by-using-ews-in-exchange
func DeleteItems(c Client, id []ItemId, strategy ...DeleteStrategy) error {

	strategy = append(strategy, DeleteStrategy{
		DeleteType:               "MoveToDeletedItems",
		SendMeetingCancellations: "SendToAllAndSaveCopy",
	})

	item := DeleteItem{
		ItemIds: ItemIds{
			ItemId: id,
		},
		DeleteStrategy: strategy[0],
	}

	xmlBytes, err := xml.MarshalIndent(item, "", "  ")
	if err != nil {
		return err
	}

	bb, err := c.SendAndReceive(xmlBytes)
	if err != nil {
		return err
	}
	return checkDeleteItemResponseForErrors(bb)
}

func checkDeleteItemResponseForErrors(bb []byte) (err error) {
	var soapResp deleteItemResponseBodyEnvelop
	if err = xml.Unmarshal(bb, &soapResp); err != nil {
		return
	}

	resp := soapResp.Body.DeleteItemResponse.ResponseMessages.DeleteItemResponseMessage
	if resp.ResponseClass == ResponseClassError {
		err = errors.New(resp.MessageText)
	}
	return
}
