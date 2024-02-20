Office.actions.associate("triggerComposeEvent", triggerComposeEvent);

async function triggerComposeEvent(eventObj) {
  const data = Office.context.roamingSettings.get("data");

  if (!Office.context.mailbox.item) return;

  Office.context.mailbox.item.notificationMessages.addAsync("key", {
    type: "informationalMessage",
    message: stringify(data).slice(0, 150),
    icon: "Icon.64x64",
    persistent: false,
    key: "info",
  });

  eventObj.completed();
}

function stringify(obj) {
  if (!obj) return "[no return] (type: " + typeof obj + ")";

  const stringified = JSON.stringify(obj);

  if (!stringified) return "[empty] (type: " + typeof obj + ")";

  return stringified;
}

Office.onReady();
