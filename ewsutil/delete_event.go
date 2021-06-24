package ewsutil

import (
	"github.com/mhewedy/ews"
)

func DeleteEvent(
	c ews.Client, id ...ews.ItemId,
) error {
	return ews.DeleteItems(c, id)
}
