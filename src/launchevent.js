/*
 * Auto CC: runs when the user opens a new email, reply, reply all or forward,
 * and adds the user's saved addresses to the CC line. The user can see them and
 * remove them before sending.
 *
 * No addresses are stored in the code. Each user enters their own on the settings
 * page (settings.html, opened from the add-in's button under Apps), which saves
 * them in that user's mailbox.
 */

var CC_SETTING = "ccAddresses";

function onNewMessageComposeHandler(event) {
  var addresses = Office.context.roamingSettings.get(CC_SETTING);
  if (!addresses || !addresses.length) {
    event.completed();
    return;
  }
  Office.context.mailbox.item.cc.addAsync(addresses, function () {
    event.completed();
  });
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
