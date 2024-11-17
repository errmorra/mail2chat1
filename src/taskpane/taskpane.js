/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  Office.onReady(() => {
    if (Office.context.mailbox.item) {
      const item = Office.context.mailbox.item;
      initializeChatView(item);
    }
  });
  
  // Initialize and populate the chat view
  async function initializeChatView(item) {
    const chatContainer = document.getElementById("chat-container");
    chatContainer.innerHTML = "<p>Loading conversation...</p>";
  
    try {
      // Get conversation ID
      const conversationId = item.conversationId;
  
      // Fetch email thread messages
      const messages = await fetchConversationMessages(conversationId);
  
      // Render chat messages
      renderChatMessages(messages);
    } catch (error) {
      console.error("Error loading conversation:", error);
      chatContainer.innerHTML = "<p>Failed to load conversation.</p>";
    }
  }
  
  // Render messages in the chat interface
  function renderChatMessages(messages) {
    const chatContainer = document.getElementById("chat-container");
    chatContainer.innerHTML = ""; // Clear loading text
  
    messages.forEach((message) => {
      const messageDiv = document.createElement("div");
      messageDiv.className = message.isSender ? "message sender" : "message recipient";
      messageDiv.innerHTML = `<strong>${message.sender}:</strong><br>${message.body}`;
      chatContainer.appendChild(messageDiv);
    });
  }
  

  const item = Office.context.mailbox.item;
  let insertAt = document.getElementById("item-subject");
  let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(item.subject));
  insertAt.appendChild(document.createElement("br"));
}
