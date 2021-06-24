package ewsutil

import (
	"time"

	"github.com/mhewedy/ews"
)

func UpdateHTMLEvent(
	c ews.Client, id ews.ItemId, to, optional []string, subject, body, location string, from time.Time, duration time.Duration,
) ([]ews.ItemId, error) {
	return updateEvent(c, id, to, optional, subject, body, location, "HTML", from, duration)
}

// UpdateEvent helper method to update Message
func UpdateEvent(
	c ews.Client, id ews.ItemId, to, optional []string, subject, body, location string, from time.Time, duration time.Duration,
) ([]ews.ItemId, error) {
	return updateEvent(c, id, to, optional, subject, body, location, "Text", from, duration)
}

func updateEvent(
	c ews.Client, id ews.ItemId, to, optional []string, subject, body, location, bodyType string, from time.Time, duration time.Duration,
) ([]ews.ItemId, error) {

	requiredAttendees := make([]ews.Attendee, len(to))
	for i, tt := range to {
		requiredAttendees[i] = ews.Attendee{Mailbox: ews.Mailbox{EmailAddress: tt}}
	}

	optionalAttendees := make([]ews.Attendee, len(optional))
	for i, tt := range optional {
		optionalAttendees[i] = ews.Attendee{Mailbox: ews.Mailbox{EmailAddress: tt}}
	}

	room := make([]ews.Attendee, 1)
	room[0] = ews.Attendee{Mailbox: ews.Mailbox{EmailAddress: location}}

	m := ews.CalendarItem{
		Subject: subject,
		Body: ews.Body{
			BodyType: bodyType,
			Body:     []byte(body),
		},
		// ReminderIsSet: duration != 0,
		// ReminderMinutesBeforeStart: 15,
		Start:                from,
		End:                  from.Add(duration),
		IsAllDayEvent:        duration == 0,
		LegacyFreeBusyStatus: ews.BusyTypeBusy,
		Location:             location,
		RequiredAttendees:    []ews.Attendees{{Attendee: requiredAttendees}},
		OptionalAttendees:    []ews.Attendees{{Attendee: optionalAttendees}},
		Resources:            []ews.Attendees{{Attendee: room}},
	}

	return ews.UpdateCalendarItem(c, id, m)
}
