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

  /* 
  if (notificationData.type == 'message') {
    // Email messages log subject and message ID
    const subjectCell = document.createElement('td');
    subjectCell.innerText = notificationData.resource.subject;
    tableRow.appendChild(subjectCell);

    const idCell = document.createElement('td');
    idCell.innerText = notificationData.resource.id;
    tableRow.appendChild(idCell);
  }
  */

  if (notificationData.type === 'message') {
    // Teams channel messages log sender and text
    const senderCell = document.createElement('td');
    senderCell.innerText =
      notificationData.resource.from?.emailAddress?.name || 'Unknown' + ' <' + notificationData.resource.from?.emailAddress?.address + '>';
    tableRow.appendChild(senderCell);

    const messageCell = document.createElement('td');
    messageCell.innerText = `subject: ${notificationData.resource.subject || ''} body: ${notificationData.resource.bodyPreview || ''}`;
    tableRow.appendChild(messageCell);
  }

  document.getElementById('notifications').appendChild(tableRow);
};
