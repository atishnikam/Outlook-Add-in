Office.onReady().then(function () {
  const item = Office.context.mailbox.item;
  if (!item) return;

  // Insert signature into new items (compose mode)
  if (
    item.itemType === Office.MailboxEnums.ItemType.Appointment ||
    item.itemType === Office.MailboxEnums.ItemType.Message
  ) {
    item.body.getAsync(Office.CoercionType.Html, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const currentBody = result.value;
        const signature = "<br><br>--<br>External Meeting | Set Signature";

        if (!currentBody.includes("External Meeting | Set Signature")) {
          item.body.setAsync(currentBody + signature, { coercionType: "html" }, function (asyncResult) {
            console.log("Auto signature appended to viewed/edited item:", JSON.stringify(asyncResult));
          });
        }
      }
    });
  }
});

// Launch event for read mode (when item is opened)
function onMessageOrAppointmentRead(event) {
  const item = Office.context.mailbox.item;
  if (!item) {
    event.completed();
    return;
  }

  item.body.getAsync(Office.CoercionType.Html, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const body = result.value;
      const signature = "<br><br>--<br>External Meeting | Set Signature";

      if (!body.includes("External Meeting | Set Signature")) {
        // **Don’t modify the actual item** — just log or handle preview
        console.log("Signature missing in read item");
      }
    }
    event.completed();
  });
}
