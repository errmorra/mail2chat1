// Run when the Office Add-in is loaded
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
  
  // Simulate fetching email thread (Replace with Microsoft Graph API call in real app)
  async function fetchConversationMessages(conversationId) {
    // Simulate a delay for fetching data
    return new Promise((resolve) => {
      setTimeout(() => {
        resolve([
          { sender: "Alice", body: "Hi! How are you?", isSender: true },
          { sender: "Bob", body: "I'm good, thanks! How about you?", isSender: false },
          { sender: "Alice", body: "Doing great, thanks for asking!", isSender: true },
        ]);
      }, 1000);
    });
  }
  