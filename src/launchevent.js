/*
 * Auto CC: runs when the user opens a new email, reply, reply all or forward,
 * and adds the user's saved addresses to the CC line unless they are already
 * in To, CC or BCC. A red notice in the compose window says what was added; the
 * user can delete the address from CC as normal.
 *
 * No addresses are stored in the code. Each user enters their own on the settings
 * page (settings.html, opened from the add-in's button under Apps), which saves
 * them in that user's mailbox.
 *
 * Classic Outlook loads this file directly; it supports only ES2016-era JavaScript,
 * so no arrow functions, async/await or the ?: operator here.
 */

var CC_SETTING = "ccAddresses";
var NOTICE_KEY = "autoCcNotice";

// Set to true to show a bar when nothing was added or something failed (for troubleshooting).
// Bump ?v= in commands.html after changing this file.
var SHOW_DIAGNOSTICS = false;

function showDiagnostic(message, done) {
  if (!SHOW_DIAGNOSTICS) {
    done();
    return;
  }
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    NOTICE_KEY,
    { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message: ("Auto CC test: " + message).substring(0, 150) },
    function () {
      done();
    }
  );
}

// Collects the lower-cased addresses already in To, CC and BCC.
function getExistingRecipients(item, callback) {
  var fields = [item.to, item.cc, item.bcc];
  var found = [];
  var index = 0;

  function next() {
    if (index >= fields.length) {
      callback(found);
      return;
    }
    var field = fields[index];
    index++;
    if (!field || !field.getAsync) {
      next();
      return;
    }
    field.getAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        for (var i = 0; i < result.value.length; i++) {
          if (result.value[i].emailAddress) {
            found.push(result.value[i].emailAddress.toLowerCase());
          }
        }
      }
      next();
    });
  }

  next();
}

// Red bar at the top of the compose window, so the user notices the extra CC.
function showAddedNotice(item, added, done) {
  item.notificationMessages.replaceAsync(
    NOTICE_KEY,
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: ("Added " + added.join(", ") + " to CC.").substring(0, 150)
    },
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
    var saved = Office.context.roamingSettings.get(CC_SETTING);
    if (!saved || !saved.length) {
      showDiagnostic("no saved addresses were found.", finish);
      return;
    }
    var item = Office.context.mailbox.item;
    getExistingRecipients(item, function (existing) {
      var missing = saved.filter(function (address) {
        return existing.indexOf(address.toLowerCase()) === -1;
      });
      if (!missing.length) {
        showDiagnostic("all saved addresses were already recipients.", finish);
        return;
      }
      item.cc.addAsync(missing, function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          showDiagnostic("could not add CC: " + result.error.message, finish);
          return;
        }
        showAddedNotice(item, missing, finish);
      });
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
