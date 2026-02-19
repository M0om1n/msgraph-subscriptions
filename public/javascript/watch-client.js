// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Connect to the Socket.io server
const socket = io();

// Listen for notification received messages
socket.on('notification_received', (notificationData) => {
  console.log(`Received notification: ${JSON.stringify(notificationData)}`);

  // Create a new table row with data from the notification
  const tableRow = document.createElement('tr');

  if (notificationData.type == 'message') {
    // Email messages log subject and message ID
    const subjectCell = document.createElement('td');
    subjectCell.innerText = notificationData.resource.subject;
    tableRow.appendChild(subjectCell);

    const idCell = document.createElement('td');
    idCell.innerText = notificationData.resource.id;
    tableRow.appendChild(idCell);
  } else if (notificationData.type === 'chatMessage') {
    // Teams channel messages log sender and text
    const senderCell = document.createElement('td');
    senderCell.innerText =
      notificationData.resource.from.user?.displayName || 'Unknown';
    tableRow.appendChild(senderCell);

    const messageCell = document.createElement('td');
    messageCell.innerText = notificationData.resource.body?.content || '';
    tableRow.appendChild(messageCell);
  }

  document.getElementById('notifications').appendChild(tableRow);
});

// Create a room for the subscription ID
socket.emit('create_room', subscriptionId);
