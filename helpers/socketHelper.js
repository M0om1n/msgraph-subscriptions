// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import express from "express";
import http from "http";
import * as io from "socket.io";

const socketServer = http.createServer(express);

// Create a Socket.io server
const ioServer = new io.Server(socketServer, {
  cors: {
    // Allow requests from the server only
    origin: [
      process.env.OAUTH_REDIRECT_URI.substring(
        0,
        process.env.OAUTH_REDIRECT_URI.indexOf('/', 'https://'.length),
      ),
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
