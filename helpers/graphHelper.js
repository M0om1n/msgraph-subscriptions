// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as graph from "@microsoft/microsoft-graph-client";

export default {
  /**
   * Gets a Graph client configured to use delegated auth
   * @param  {IConfidentialClientApplication} msalClient - The MSAL client used to retrieve user tokens
   * @param  {string} userAccountId - The user's account ID
   */
  getGraphClientForUser(msalClient, userAccountId) {
    if (!msalClient || !userAccountId) {
      throw new Error(
        `Invalid MSAL state. Client: ${
          msalClient ? 'present' : 'missing'
        }, User Account ID: ${userAccountId ? 'present' : 'missing'}`,
      );
    }

    // Initialize Graph client
    return graph.Client.init({
      // Implement an auth provider that gets a token
      // from the app's MSAL instance
      authProvider: async (done) => {
        try {
          // Get the user's account
          const account = await msalClient
            .getTokenCache()
            .getAccountByHomeId(userAccountId);

          if (account) {
            // Attempt to get the token silently
            // This method uses the token cache and
            // refreshes expired tokens as needed
            const response = await msalClient.acquireTokenSilent({
              scopes: process.env.OAUTH_SCOPES.split(','),
              redirectUri: process.env.OAUTH_REDIRECT_URI,
              account: account,
            });

            // First param to callback is the error,
            // Set to null in success case
            done(null, response.accessToken);
          }
        } catch (err) {
          console.log(JSON.stringify(err, Object.getOwnPropertyNames(err)));
          done(err, null);
        }
      },
    });
  },
  /**
   * Gets a Graph client configured to use app-only auth
   * @param  {IConfidentialClientApplication} msalClient - The MSAL client used to retrieve app-only tokens
   */
  getGraphClientForApp(msalClient) {
    if (!msalClient) {
      throw new Error('Invalid MSAL state. MSAL client is missing.');
    }

    // Initialize Graph client
    return graph.Client.init({
      // Implement an auth provider that gets a token
      // from the app's MSAL instance
      authProvider: async (done) => {
        try {
          // Get a token using client credentials
          const response = await msalClient.acquireTokenByClientCredential({
            scopes: ['https://graph.microsoft.com/.default'],
          });

          // First param to callback is the error,
          // Set to null in success case
          done(null, response.accessToken);
        } catch (err) {
          console.log(JSON.stringify(err, Object.getOwnPropertyNames(err)));
          done(err, null);
        }
      },
    });
  },
};
