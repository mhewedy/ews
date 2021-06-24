package ewsutil

import "github.com/mhewedy/ews"

// SendEmail helper method to send Message
func SendEmail(c ews.Client, to []string, subject, body string) (err error) {

	m := ews.Message{
		ItemClass: "IPM.Note",
		Subject:   subject,
		Body: ews.Body{
			BodyType: "Text",
			Body:     []byte(body),
		},
		Sender: ews.OneMailbox{
			Mailbox: ews.Mailbox{
				EmailAddress: c.GetUsername(),
			},
		},
	}
	mb := make([]ews.Mailbox, len(to))
	for i, addr := range to {
		mb[i].EmailAddress = addr
	}
	m.ToRecipients.Mailbox = append(m.ToRecipients.Mailbox, mb...)

	_, err = ews.CreateMessageItem(c, m)
	return
}
