// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import express from "express";
import http from "http";
import { Server } from "socket.io";

const socketServer = http.createServer(express);

const redirectUri = process.env.OAUTH_REDIRECT_URI;
const originUrl = redirectUri.substring(0, redirectUri.indexOf('/', 'https://'.length));

// Create a Socket.io server
const ioServer = new Server(socketServer, {
  cors: {
    // Allow requests from the server only
    origin: [originUrl],
    methods: ['GET', 'POST'],
  },
});

ioServer.on('connection', (socket) => {
  // Create rooms by subscription ID
  socket.on('create_room', (subscriptionId) => {
    socket.join(subscriptionId);
  });
});

// Listen on port
socketServer.listen(3002);
console.log(`Socket.io listening on port 3002`);

export default ioServer;
