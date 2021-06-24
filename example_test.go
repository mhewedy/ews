package ews_test

import (
	"fmt"
	"io/ioutil"
	"math"
	"os"
	"testing"
	"time"

	. "github.com/mhewedy/ews"
	"github.com/mhewedy/ews/ewsutil"
)

func Test_Example(t *testing.T) {

	c := NewClient(
		"https://outlook.office365.com/EWS/Exchange.asmx",
		"daziplqa@daziplqa.onmicrosoft.com",
		"systemsystem@123",
		&Config{Dump: true, NTLM: true, SkipTLS: true},
	)

	/*c := NewClient(
		"https://mail.mycompany/EWS/Exchange.asmx",
		"company\\user",
		"password",
		&Config{Dump: true, NTLM: true},
	)*/

	//err := testSendEmail(c)

	//err := testCreateCalendarItem(c)

	//err := testGetUserAvailability(c)

	//err := testListUsersEvents(c)

	//err := testCreateEvent(c)

	//err := testGetRoomLists(c)

	//err := testFindPeople(c)

	//err := testGetUserPhoto(c)

	//err := testEWSUtilFindPeople(c)

	//err := testGetPersona(c)

	//err := testCreateHTMLEvent(c)

	//if err != nil {
	//	log.Fatal("err>: ", err.Error())
	//}

	_ = c

	fmt.Println("--- success ---")
}

func testSendEmail(c Client) error {
	return ewsutil.SendEmail(c,
		[]string{"mhewedy@gmail.com", "someone@else.com"},
		"An email subject",
		"The email body, as plain text",
	)
}

func testCreateCalendarItem(c Client) error {
	attendee := make([]Attendee, 0)
	attendee = append(attendee,
		Attendee{Mailbox: Mailbox{EmailAddress: "mhewedy@mhewedy.onmicrosoft.com"}},
	)
	attendees := make([]Attendees, 0)
	attendees = append(attendees, Attendees{Attendee: attendee})

	var _, err = CreateCalendarItem(c, CalendarItem{
		Subject: "Planning Meeting",
		Body: Body{
			BodyType: "Text",
			Body:     []byte("Plan the agenda for next week's meeting."),
		},
		ReminderIsSet:              true,
		ReminderMinutesBeforeStart: 60,
		Start:                      time.Now().Add(24 * time.Hour),
		End:                        time.Now().Add(24 * time.Hour).Add(30 * time.Minute),
		IsAllDayEvent:              false,
		LegacyFreeBusyStatus:       "Busy",
		Location:                   "Conference Room 721",
		RequiredAttendees:          attendees,
	})
	return err
}

func testGetUserAvailability(c Client) error {

	mb := make([]MailboxData, 0)
	mb = append(mb, MailboxData{
		Email: Email{
			Name:        "",
			Address:     "mhewedy@mhewedy.onmicrosoft.com",
			RoutingType: "SMTP",
		},
		AttendeeType:     AttendeeTypeRequired,
		ExcludeConflicts: false,
	}, MailboxData{
		Email: Email{
			Name:        "",
			Address:     "example2@mhewedy.onmicrosoft.com",
			RoutingType: "SMTP",
		},
		AttendeeType:     AttendeeTypeRoom,
		ExcludeConflicts: false,
	})

	start, _ := time.Parse(time.RFC3339, "2019-11-28T00:00:00Z")
	end, _ := time.Parse(time.RFC3339, "2019-12-10T00:00:00Z")

	req := &GetUserAvailabilityRequest{
		TimeZone: TimeZone{
			Bias: -180,
			StandardTime: TimeZoneTime{
				Bias:      0,
				Time:      "02:00:00",
				DayOrder:  5,
				Month:     10,
				DayOfWeek: "Sunday",
			},
			DaylightTime: TimeZoneTime{
				Bias:      0,
				Time:      "02:00:00",
				DayOrder:  1,
				Month:     4,
				DayOfWeek: "Sunday",
			},
		},
		MailboxDataArray: MailboxDataArray{MailboxData: mb},
		FreeBusyViewOptions: FreeBusyViewOptions{
			TimeWindow: TimeWindow{
				StartTime: start,
				EndTime:   end,
			},
			MergedFreeBusyIntervalInMinutes: 30,
			RequestedView:                   RequestedViewFreeBusy,
		},
	}

	response, err := GetUserAvailability(c, req)
	if err != nil {
		return err
	}

	fmt.Println(response)

	return nil
}

func testListUsersEvents(c Client) error {

	eventUsers := []ewsutil.EventUser{
		{
			Email:        "mhewedy@mhewedy.onmicrosoft.com",
			AttendeeType: AttendeeTypeRequired,
		},
		{
			Email:        "example@mhewedy.onmicrosoft.com",
			AttendeeType: AttendeeTypeRequired,
		},
		{
			Email:        "room001@mhewedy.onmicrosoft.com",
			AttendeeType: AttendeeTypeResource,
		},
	}
	start, _ := time.Parse(time.RFC3339, "2019-12-10T11:00:00+03:00")

	events, err := ewsutil.ListUsersEvents(c, eventUsers, start, 48*time.Hour)

	if err != nil {
		return err
	}

	fmt.Println(events)

	return nil
}

func testCreateEvent(c Client) error {

	var _, err = ewsutil.CreateEvent(c,
		[]string{"mhewedy@mhewedy.onmicrosoft.com", "example2@mhewedy.onmicrosoft.com"},
		[]string{},
		"An Event subject",
		"An Event body, as plain text",
		"Room 55",
		time.Now().Add(24*time.Hour),
		30*time.Minute,
	)
	return err
}

func testCreateHTMLEvent(c Client) error {

	var _, err = ewsutil.CreateHTMLEvent(c,
		[]string{"mhewedy@mhewedy.onmicrosoft.com", "example@mhewedy.onmicrosoft.com"},
		[]string{},
		"An Event subject",
		`An Event body, as HTML <br/><br/><br/><div style="color: gray; font-size: x-small;">
		Sent by <a style="color: gray; text-decoration: none;" href="https://github.com/mhewedy/mego"
		target="_blank">MEGO</a> - the meeting organizer</div>`,
		"Room 55",
		time.Now().Add(24*time.Hour),
		30*time.Minute,
	)
	return err
}

func testGetRoomLists(c Client) error {
	response, err := GetRoomLists(c)
	if err != nil {
		return err
	}
	fmt.Println(response)

	return nil
}

func testFindPeople(c Client) error {

	req := &FindPeopleRequest{IndexedPageItemView: IndexedPageItemView{
		MaxEntriesReturned: math.MaxInt32,
		Offset:             0,
		BasePoint:          BasePointBeginning,
	}, ParentFolderId: ParentFolderId{
		DistinguishedFolderId: DistinguishedFolderId{Id: "directory"}},
		PersonaShape: &PersonaShape{BaseShape: BaseShapeIdOnly,
			AdditionalProperties: AdditionalProperties{
				FieldURI: []FieldURI{
					{FieldURI: "persona:DisplayName"},
					{FieldURI: "persona:Title"},
					{FieldURI: "persona:EmailAddress"},
				},
			}},
		QueryString: "ex",
	}

	resp, err := FindPeople(c, req)

	if err != nil {
		return err
	}

	fmt.Println(resp)

	return nil
}

func testGetUserPhoto(c Client) error {

	bytes, err := ewsutil.GetUserPhoto(c, "mhewedy@mhewedy.onmicrosoft.com")

	if err != nil {
		return err
	}

	err = ioutil.WriteFile("/tmp/file.png", bytes, os.ModePerm)
	fmt.Println("written to: /tmp/file.png")

	return err
}

func testEWSUtilFindPeople(c Client) error {

	resp, err := ewsutil.FindPeople(c, "test")

	if err != nil {
		return err
	}

	fmt.Println(resp)

	return nil
}

func testGetPersona(c Client) error {

	personas, _ := ewsutil.FindPeople(c, "hewedy")

	resp, err := GetPersona(c, &GetPersonaRequest{
		PersonaId: personas[0].PersonaId,
	})

	if err != nil {
		return err
	}

	fmt.Println(resp)

	return nil
}
