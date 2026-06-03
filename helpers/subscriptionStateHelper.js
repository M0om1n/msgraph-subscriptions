// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const CLIENT_STATE_PREFIX = 'msgraph-sample-v1';
const MAX_CLIENT_STATE_LENGTH = 255;

function sanitizeText(value, maxLength) {
  return String(value || '').trim().slice(0, maxLength);
}

function safeParseMetadata(encodedPayload) {
  try {
    const decoded = Buffer.from(encodedPayload, 'base64url').toString('utf8');
    const parsed = JSON.parse(decoded);
    return {
      userName: sanitizeText(parsed.userName, 80),
      calendarName: sanitizeText(parsed.calendarName, 80),
    };
  } catch {
    return null;
  }
}

export default {
  buildClientState: (baseState, metadata = {}) => {
    const normalizedBaseState = sanitizeText(baseState, 120);
    if (!normalizedBaseState) {
      return '';
    }

    const normalizedMetadata = {
      userName: sanitizeText(metadata.userName, 80),
      calendarName: sanitizeText(metadata.calendarName, 80),
    };

    const hasMetadata = normalizedMetadata.userName || normalizedMetadata.calendarName;
    if (!hasMetadata) {
      return normalizedBaseState;
    }

    const payload = Buffer.from(
      JSON.stringify(normalizedMetadata),
      'utf8',
    ).toString('base64url');

    const extendedState = `${CLIENT_STATE_PREFIX}|${normalizedBaseState}|${payload}`;
    return extendedState.length <= MAX_CLIENT_STATE_LENGTH
      ? extendedState
      : normalizedBaseState;
  },

  parseClientState: (rawClientState) => {
    const clientState = String(rawClientState || '');
    const defaultResult = {
      raw: clientState,
      isExtended: false,
      baseState: clientState,
      metadata: null,
    };

    if (!clientState.startsWith(`${CLIENT_STATE_PREFIX}|`)) {
      return defaultResult;
    }

    const segments = clientState.split('|');
    if (segments.length < 3) {
      return defaultResult;
    }

    const baseState = segments[1] || '';
    const encodedPayload = segments.slice(2).join('|');
    const metadata = safeParseMetadata(encodedPayload);

    return {
      raw: clientState,
      isExtended: true,
      baseState,
      metadata,
    };
  },

  isExpectedClientState: (rawClientState, expectedBaseState) => {
    const normalizedExpected = String(expectedBaseState || '');
    if (!normalizedExpected) {
      return false;
    }

    if (String(rawClientState || '') === normalizedExpected) {
      return true;
    }

    const parsedResult = (function parse() {
      const result = String(rawClientState || '');
      if (!result.startsWith(`${CLIENT_STATE_PREFIX}|`)) {
        return null;
      }

      const parts = result.split('|');
      if (parts.length < 3) {
        return null;
      }

      return {
        baseState: parts[1] || '',
      };
    }());

    return parsedResult ? parsedResult.baseState === normalizedExpected : false;
  },
};
