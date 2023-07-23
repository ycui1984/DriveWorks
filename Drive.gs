  /**
 * Callback for rendering the card for specific Drive items.
 * @param {Object} e The event object.
 * @return {CardService.Card} The card to show to the user.
 */

function onDriveItemsSelected(e) {
  console.log(e);
  return createHomeCard(e.drive.activeCursorItem);
}

