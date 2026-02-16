// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import express from "express";
import http from "http";
import * as io from "socket.io";

const socketServer = http.createServer(express);

const redirectUri = process.env.OAUTH_REDIRECT_URI || 'https://localhost:3000/delegated/callback';

// Create a Socket.io server
const ioServer = new io.Server(socketServer, {
  cors: {
    // Allow requests from the server only
    origin: [
      redirectUri.substring(0, redirectUri.indexOf('/', 'https://'.length)),
    ],
    methods: ['GET', 'POST'],
  },
});

ioServer.on('connection', (socket) => {
  // Create rooms by subscription ID
  socket.on('create_room', (subscriptionId) => {
    socket.join(subscriptionId);
  });
});

// Listen on port 3001
socketServer.listen(3001);
console.log('Socket.io listening on port 3001');

export default ioServer;
