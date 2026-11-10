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

// TEMPORARY: shows a bar in the compose window saying what the add-in did.
// Set to false once Auto CC is confirmed working.
var SHOW_DIAGNOSTICS = true;

function showDiagnostic(message, done) {
  if (!SHOW_DIAGNOSTICS) {
    done();
    return;
  }
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "autoCcDiagnostic",
    { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message: ("Auto CC test: " + message).substring(0, 150) },
    function () {
      done();
    }
  );
}

function onNewMessageComposeHandler(event) {
  var finish = function () {
    event.completed();
  };
  try {
    var addresses = Office.context.roamingSettings.get(CC_SETTING);
    if (!addresses || !addresses.length) {
      showDiagnostic("ran, but no saved addresses were found.", finish);
      return;
    }
    Office.context.mailbox.item.cc.addAsync(addresses, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        showDiagnostic("added " + addresses.length + " address(es) to CC.", finish);
      } else {
        showDiagnostic("could not add CC: " + result.error.message, finish);
      }
    });
  } catch (e) {
    var detail = String(e);
    if (e && e.message) {
      detail = e.message;
    }
    showDiagnostic("error: " + detail, finish);
  }
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
