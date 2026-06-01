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
    // Delegated notifications log event details
    const subjectCell = document.createElement('td');
    subjectCell.innerText = notificationData.resource.subject || '(no subject)';
    tableRow.appendChild(subjectCell);

    const startCell = document.createElement('td');
    startCell.innerText = formatEventDateTime(notificationData.resource.start);
    tableRow.appendChild(startCell);

    appendEventDetailCells(tableRow, notificationData.resource);
  }
  if (notificationData.type === 'app_event') {
    // App-only notifications log organizer and event details
    const organizerCell = document.createElement('td');
    const organizer = notificationData.resource.organizer;
    const organizerEmail = organizer && organizer.emailAddress;
    organizerCell.innerText =
      (organizerEmail && organizerEmail.name) ||
      (organizerEmail && organizerEmail.address) ||
      'Unknown';
    tableRow.appendChild(organizerCell);

    const subjectCell = document.createElement('td');
    subjectCell.innerText = notificationData.resource.subject || '(no subject)';
    tableRow.appendChild(subjectCell);

    const startCell = document.createElement('td');
    startCell.innerText = formatEventDateTime(notificationData.resource.start);
    tableRow.appendChild(startCell);

    appendEventDetailCells(tableRow, notificationData.resource);
  }

  document.getElementById('notifications').appendChild(tableRow);
};

function appendEventDetailCells(tableRow, eventResource) {
  const locationCell = document.createElement('td');
  const location = eventResource.location;
  locationCell.innerText =
    (location && location.displayName) ||
    (location && location.locationUri) ||
    '(none)';
  tableRow.appendChild(locationCell);

  const attendeesCell = document.createElement('td');
  appendExpandableText(attendeesCell, formatAttendees(eventResource.attendees), 90);
  tableRow.appendChild(attendeesCell);

  const bodyPreviewCell = document.createElement('td');
  appendExpandableText(bodyPreviewCell, eventResource.bodyPreview || '(none)', 120);
  tableRow.appendChild(bodyPreviewCell);

  const bodyCell = document.createElement('td');
  appendExpandableText(bodyCell, formatBody(eventResource.body), 180);
  tableRow.appendChild(bodyCell);
}

function appendExpandableText(cell, text, maxLength) {
  const safeText = String(text || '(none)');

  if (safeText.length <= maxLength) {
    cell.innerText = safeText;
    return;
  }

  const content = document.createElement('div');
  content.className = 'expandable-cell-text';
  content.innerText = `${safeText.slice(0, maxLength)}...`;

  const toggle = document.createElement('button');
  toggle.type = 'button';
  toggle.className = 'btn btn-sm btn-link p-0 ms-1 expandable-cell-toggle';
  toggle.innerText = 'Show more';

  let expanded = false;
  toggle.addEventListener('click', () => {
    expanded = !expanded;
    content.innerText = expanded ? safeText : `${safeText.slice(0, maxLength)}...`;
    toggle.innerText = expanded ? 'Show less' : 'Show more';
  });

  cell.appendChild(content);
  cell.appendChild(toggle);
}

function formatEventDateTime(start) {
  if (!start || !start.dateTime) {
    return 'Unknown';
  }

  const timezone = start.timeZone || 'UTC';
  return `${start.dateTime} (${timezone})`;
}

function formatAttendees(attendees) {
  if (!Array.isArray(attendees) || attendees.length === 0) {
    return '(none)';
  }

  return attendees
    .map((attendee) => {
      const emailAddress = attendee && attendee.emailAddress;
      return (emailAddress && (emailAddress.name || emailAddress.address)) || null;
    })
    .filter((value) => value)
    .join(', ') || '(none)';
}

function formatBody(body) {
  if (!body || !body.content) {
    return '(none)';
  }

  // Graph event body can be HTML, so render it as plain text in the table.
  const bodyHtml = String(body.content);
  const tempContainer = document.createElement('div');
  tempContainer.innerHTML = bodyHtml;
  const plainText = (tempContainer.textContent || tempContainer.innerText || '').trim();

  return plainText || '(none)';
}
