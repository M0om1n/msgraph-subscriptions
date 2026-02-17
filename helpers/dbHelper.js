// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const storage = new Map();

export default {
  /**
   * Gets a single subscription by ID
   * @param  {string} subscriptionId - The ID of the subscription to get
   * @returns {object} The subscription
   */
  getSubscription: (subscriptionId) => {
    const userAccountId = storage.get(subscriptionId);
    if (userAccountId) {
      return {
        subscriptionId,
        userAccountId
      };
    }
    return null;
  },
  /**
   * Gets all subscriptions for a user account
   * @param  {string} userAccountId - The user account ID
   * @returns {Array} An array of subscriptions for the user
   */
  getSubscriptionsByUserAccountId: (userAccountId) => {
    return [...storage]
      .filter(([key, value]) => value === userAccountId) 
      .map(([key, value]) => key);
  },
  /**
   * Adds a subscription to the database
   * @param  {string} subscriptionId - The subscription ID
   * @param  {string} userAccountId - The user account ID (use 'APP-ONLY' for subscriptions owned by the app)
   */
  addSubscription: (subscriptionId, userAccountId) => {
    storage.set(subscriptionId, userAccountId);
  },
  /**
   * Deletes a subscription from the database
   * @param  {string} subscriptionId - The ID of the subscription to delete
   */
  deleteSubscription: (subscriptionId) => {
    storage.delete(subscriptionId);
  },
};
