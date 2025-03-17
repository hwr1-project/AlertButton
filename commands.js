/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Sends an emergency alert to all users.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const alertMessage = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "🚨 Emergency Alert Sent to All Employees 🚨",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message in Outlook.
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "EmergencyAlertNotification",
    alertMessage
  );

  // Send a request to a central server (You need to replace this URL).
  fetch("https://your-server.com/trigger-alert", { method: "POST" })
    .then(response => response.json())
    .then(data => console.log("Alert Sent Successfully:", data))
    .catch(error => console.error("Error sending alert:", error));

  // Mark the command as completed.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);