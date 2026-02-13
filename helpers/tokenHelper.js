// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as jwt from "jsonwebtoken";
import jwksClient from "jwks-rsa";

// Configure JSON web key set client to get keys
// from well-known Microsoft identity endpoint
const client = jwksClient({
  jwksUri: 'https://login.microsoftonline.com/common/discovery/v2.0/keys',
});

/**
 * Gets the key specified in header
 * @param  {JwtHeader} header - The header containing the key ID
 * @param  {function} callback - The callback function
 */
async function getKey(header, callback) {
  try {
    const key = await client.getSigningKey(header.kid);
    const signingKey = key.publicKey || key.rsaPublicKey;
    callback(null, signingKey);
  } catch (err) {
    callback(err, null);
  }
}

export default {
  /**
   * Validates a token has a valid signature and has the expected audience and issuer
   * @param  {string} token - The token to verify
   * @param  {string} appId - The application ID expected in the audience claim
   * @param  {string} tenantId - The tenant ID expected in the issuer claim
   */
  isTokenValid: (token, appId, tenantId) => {
    return new Promise((resolve) => {
      const options = {
        audience: [appId],
        issuer: [`https://sts.windows.net/${tenantId}/`],
      };

      jwt.verify(token, getKey, options, (err) => {
        if (err) {
          console.log(`Token validation error: ${err.message}`);
          resolve(false);
        }
        resolve(true);
      });
    });
  },
};
