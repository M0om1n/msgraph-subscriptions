// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Connect to the WebSocket server
const socket = new WebSocket(`wss://hello-msgraph.onrender.com/ws`);

// Listen for notification received messages
socket.onmessage = (event) => {
  const notificationData = JSON.parse(event.data);
  console.log(`Received notification: ${JSON.stringify(notificationData)}`);

  // Create a new table row with data from the notification
  const tableRow = document.createElement('tr');

  if (notificationData.type == 'user_event') {
    // Delegated notifications log event subject and start time
    const subjectCell = document.createElement('td');
    subjectCell.innerText = notificationData.resource.subject || '(no subject)';
    tableRow.appendChild(subjectCell);

    const startCell = document.createElement('td');
    startCell.innerText = formatEventDateTime(notificationData.resource.start);
    tableRow.appendChild(startCell);
  }
  if (notificationData.type === 'app_event') {
    // App-only notifications log organizer and event details
    const organizerCell = document.createElement('td');
    organizerCell.innerText =
      notificationData.resource.organizer?.emailAddress?.name ||
      notificationData.resource.organizer?.emailAddress?.address ||
      'Unknown';
    tableRow.appendChild(organizerCell);

    const subjectCell = document.createElement('td');
    subjectCell.innerText = notificationData.resource.subject || '(no subject)';
    tableRow.appendChild(subjectCell);

    const startCell = document.createElement('td');
    startCell.innerText = formatEventDateTime(notificationData.resource.start);
    tableRow.appendChild(startCell);
  }

  document.getElementById('notifications').appendChild(tableRow);
};

function formatEventDateTime(start) {
  if (!start || !start.dateTime) {
    return 'Unknown';
  }

  const timezone = start.timeZone || 'UTC';
  return `${start.dateTime} (${timezone})`;
}
