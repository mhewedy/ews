package ews

import (
	"encoding/xml"
	"log"
	"testing"

	"github.com/stretchr/testify/assert"
)

func Test_marshal_DeleteItems(t *testing.T) {

	ditem := &DeleteItem{
		ItemIds: ItemIds{
			ItemId: []ItemId{
				{"ID", "Key"},
			},
		},
	}

	xmlBytes, err := xml.MarshalIndent(ditem, "", "  ")
	if err != nil {
		log.Fatal(err)
	}

	assert.Equal(t, `<m:DeleteItem>
  <m:ItemIds>
    <t:ItemId Id="ID" ChangeKey="Key"></t:ItemId>
  </m:ItemIds>
</m:DeleteItem>`, string(xmlBytes))
}
