const {ActivityHandler, MessageFactory} = require("botbuilder");
const {OpenAI} = require("openai");

class OpenAIBot extends ActivityHandler {
    constructor() {
        super();

        // Initialize OpenAI client
        this.openai = new OpenAI({
            apiKey: process.env.OPENAI_API_KEY, // OpenAI API key from .env
        });

        // Initialize an object to track conversation history for each user
        this.conversations = {};

        // Handle incoming messages
        this.onMessage(async (context, next) => {
            const userMessage = context.activity.text;
            const userId = context.activity.from.id; // Get user ID to store their conversation context

            try {
                // Retrieve the conversation history for the user, or initialize it if not exists
                let conversationHistory = this.conversations[userId] || [];

                // If conversation is starting, add a system message to provide context to OpenAI
                if (conversationHistory.length === 0) {
                    const systemMessage = {
                        role: "system",
                        content:
                            "You are an intelligent assistant bot, named BookWorm, at the company BooksTime. You can assist bookkeepers, senior accountants, IT department, Senior Mangers and client service advisors with their queries to the best of your ability. You can provide sales support and management insights. You can advise staffs at BooksTime, a bookkeeping company, and answer their questions, and help them draft emails. If someone asks you, what is your name, you tell them your name is BookWorm",
                    };
                    conversationHistory.push(systemMessage); // Add initial system context
                }

                // Append the new user message to the conversation history
                conversationHistory.push({role: "user", content: userMessage});

                // If the length of conversation history exceeds 10, remove the second message
                // instead of first one in order to keep system prompt intact
                if (conversationHistory.length > 10) {
                    // remove the oldest message in order to keep the systemMessage intact
                    conversationHistory.splice(1, 1);
                    // conversationHistory.shift(); // Remove the first (oldest) message
                }

                // Send a typing indicator before making the OpenAI request
                await context.sendActivity({
                    type: "typing",
                });

                // Get the reply from OpenAI
                let replyText = await this.getOpenAIResponse(
                    conversationHistory
                );

                // Save the bot's response in the conversation history
                conversationHistory.push({
                    role: "assistant",
                    content: replyText,
                });

                // Update the conversation history for the user
                this.conversations[userId] = conversationHistory;

                if (
                    context.activity.attachments &&
                    context.activity.attachments.length > 0
                ) {
                    // Process each attachment
                    for (const attachment of context.activity.attachments) {
                        const downloadUrl = await this.getAttachmentUrl(
                            attachment
                        );
                        if (downloadUrl != undefined) {
                            replyText =
                                replyText +
                                " The linked attachment has DOWNLOAD URL: " +
                                downloadUrl;
                            replyText =
                                replyText +
                                ". Here is the context: " +
                                context.activity.attachments;
                        }
                    }
                }

                // Send the OpenAI response back to the user
                await context.sendActivity(
                    MessageFactory.text(replyText, replyText)
                );
            } catch (error) {
                console.error(
                    "Error while getting response from OpenAI:",
                    error
                );
                await context.sendActivity(
                    "Sorry, I can not answer your question at the moment. Please try again later. If this issue still persists, please reach out to the IT Team at BooksTime."
                );
            }

            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            const welcomeText =
                "Hello BooksTimer! I am BookWorm, an Intelligent Conversational Chatbot.\nHow can I help you today?";

            for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await context.sendActivity(
                        MessageFactory.text(welcomeText, welcomeText)
                    );
                }
            }

            await next();
        });
    }

    async getAttachmentUrl(attachment) {
        // For Teams, this would typically be the contentUrl
        // and you may need to ensure you have permission to access the file
        return attachment.contentUrl; // This should be the URL to the file
    }

    // Function to get response from OpenAI
    async getOpenAIResponse(conversationHistory) {
        try {
            const response = await this.openai.chat.completions.create({
                model: process.env.OPENAI_MODEL,
                messages: conversationHistory, // Pass the entire conversation history
            });

            // Return the bot's response (assistant's message)
            return response.choices[0].message.content;
        } catch (error) {
            throw new Error(`OpenAI API error: ${error.message}`);
        }
    }
}

module.exports.OpenAIBot = OpenAIBot;
