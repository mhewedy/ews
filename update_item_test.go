package ews

import (
	"encoding/xml"
	"log"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

func Test_marshal_UpdateItems(t *testing.T) {

	attendee := make([]Attendee, 0)
	attendee = append(attendee,
		Attendee{Mailbox: Mailbox{EmailAddress: "User1@example.com"}},
		Attendee{Mailbox: Mailbox{EmailAddress: "User2@example.com"}},
	)
	attendees := make([]Attendees, 0)
	attendees = append(attendees, Attendees{Attendee: attendee})

	start, _ := time.Parse(time.RFC3339, "2006-11-02T14:00:00Z")
	end, _ := time.Parse(time.RFC3339, "2006-11-02T15:00:00Z")

	citem := &CalendarItem{
		Subject: "Planning Meeting",
		Body: Body{
			BodyType: "Text",
			Body:     []byte("Plan the agenda for next week's meeting."),
		},
		Start:                start,
		End:                  end,
		IsAllDayEvent:        false,
		LegacyFreeBusyStatus: "Busy",
		Location:             "Conference Room 721",
		RequiredAttendees:    attendees,
	}

	uitem := &Updates{
		Updates: getSetItemField("calendar", getFields(citem)...),
	}

	xmlBytes, err := xml.MarshalIndent(uitem, "", "  ")
	if err != nil {
		log.Fatal(err)
	}

	assert.Equal(t, `<t:Updates>
  <t:SetItemField>
    <t:FieldURI FieldURI="item:Subject"></t:FieldURI>
    <t:CalendarItem>
      <t:Subject>Planning Meeting</t:Subject>
    </t:CalendarItem>
  </t:SetItemField>
  <t:SetItemField>
    <t:FieldURI FieldURI="item:Body"></t:FieldURI>
    <t:CalendarItem>
      <t:Body BodyType="Text">Plan the agenda for next week&#39;s meeting.</t:Body>
    </t:CalendarItem>
  </t:SetItemField>
  <t:SetItemField>
    <t:FieldURI FieldURI="calendar:Start"></t:FieldURI>
    <t:CalendarItem>
      <t:Start>2006-11-02T14:00:00Z</t:Start>
    </t:CalendarItem>
  </t:SetItemField>
  <t:SetItemField>
    <t:FieldURI FieldURI="calendar:End"></t:FieldURI>
    <t:CalendarItem>
      <t:End>2006-11-02T15:00:00Z</t:End>
    </t:CalendarItem>
  </t:SetItemField>
  <t:SetItemField>
    <t:FieldURI FieldURI="calendar:LegacyFreeBusyStatus"></t:FieldURI>
    <t:CalendarItem>
      <t:LegacyFreeBusyStatus>Busy</t:LegacyFreeBusyStatus>
    </t:CalendarItem>
  </t:SetItemField>
  <t:SetItemField>
    <t:FieldURI FieldURI="calendar:Location"></t:FieldURI>
    <t:CalendarItem>
      <t:Location>Conference Room 721</t:Location>
    </t:CalendarItem>
  </t:SetItemField>
  <t:SetItemField>
    <t:FieldURI FieldURI="calendar:RequiredAttendees"></t:FieldURI>
    <t:CalendarItem>
      <t:RequiredAttendees>
        <t:Attendee>
          <t:Mailbox>
            <t:EmailAddress>User1@example.com</t:EmailAddress>
          </t:Mailbox>
        </t:Attendee>
        <t:Attendee>
          <t:Mailbox>
            <t:EmailAddress>User2@example.com</t:EmailAddress>
          </t:Mailbox>
        </t:Attendee>
      </t:RequiredAttendees>
    </t:CalendarItem>
  </t:SetItemField>
</t:Updates>`, string(xmlBytes))
}
