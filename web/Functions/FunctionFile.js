/**
 * FunctionFile.js
 * Hosts UI-less ExecuteFunction actions for the PGP add-in.
 * These functions are triggered directly from ribbon buttons without
 * opening a task pane.
 */

Office.initialize = function () {};

/**
 * Display a transient info-bar notification on the current mail item.
 *
 * @param {string} icon    - Icon resource ID (e.g. "icon16")
 * @param {string} message - Text to display (max ~150 chars)
 */
function showNotification(icon, message) {
  Office.context.mailbox.item.notificationMessages.replaceAsync('pgp_status', {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    icon,
    message,
    persistent: false,
  });
}

/**
 * Placeholder for any future one-click ExecuteFunction actions.
 * Currently all PGP operations are performed inside task panes.
 */
function pgpDefaultAction(event) {
  showNotification('icon16', 'Use the Encrypt or Decrypt buttons to get started.');
  event.completed();
}

/**
 * Open a reply compose form for the current message, then prompt the user
 * to click the Encrypt ribbon button before sending.
 */
function replyEncrypted(event) {
  try {
    Office.context.mailbox.item.displayReplyForm('');
    showNotification('iconEncrypt16',
      'Reply opened — click Encrypt in the ribbon to encrypt before sending.');
  } catch (e) {
    showNotification('icon16', 'Could not open reply: ' + e.message);
  }
  event.completed();
}

/**
 * Open a reply-all compose form for the current message, then prompt the user
 * to click the Encrypt ribbon button before sending.
 */
function replyAllEncrypted(event) {
  try {
    Office.context.mailbox.item.displayReplyAllForm('');
    showNotification('iconEncrypt16',
      'Reply All opened — click Encrypt in the ribbon to encrypt before sending.');
  } catch (e) {
    showNotification('icon16', 'Could not open reply: ' + e.message);
  }
  event.completed();
}
