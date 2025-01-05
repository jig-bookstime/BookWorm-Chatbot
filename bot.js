const {ActivityHandler, MessageFactory} = require("botbuilder");
const {OpenAI} = require("openai");

const axios = require("axios");

const MAX_PAST_MESSAGE_FOR_CONTEXT = process.env.MAX_PAST_MESSAGE_FOR_CONTEXT;

class OpenAIBot extends ActivityHandler {
    constructor() {
        super();

        // Initialize OpenAI client
        this.openai = new OpenAI({
            apiKey: process.env.OPENAI_API_KEY, // OpenAI API key environment variable
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
                            "You are an intelligent assistant bot, named BookWorm, at the company BooksTime. You can assist Bookkeepers, Senior Accountants, IT Department, Senior Managers and Client Service Advisors at BooksTime with their queries to the best of your ability. You can provide sales support and management insights. You can help Bookstimers (staffs at BooksTime) in analyzing financial statements, proofreading proposals for grammar errors, upselling opportunities, finding answers to questions in bank statements, help them draft emails and much much more. If someone asks you, what is your name, you tell them your name is BookWorm.",
                    };
                    conversationHistory.push(systemMessage); // Add initial system context
                }

                let fileId = null;
                const attachments = context.activity.attachments;
                if (attachments && attachments[0]) {
                    const attachment = attachments[0];
                    const downloadUrl = await this.getAttachmentUrl(attachment);
                    console.log("Download URL: " + downloadUrl);
                    if (downloadUrl != undefined) {
                        const fileResponse = await axios.get(downloadUrl, {
                            responseType: "arraybuffer",
                        });
                        console.log("Logging fileResponse");

                        console.log(fileResponse);
                        const fileBuffer = Buffer.from(fileResponse.data);

                        // Upload the file to OpenAI with the "answers" purpose
                        const uploadResponse = await this.openai.files.create({
                            purpose: "answers", // Purpose changed from "fine-tune" to "answers"
                            file: fileBuffer,
                            filename: attachment.name,
                        });
                        console.log("Logging uploadResponse");

                        // Capture the file ID from the upload response
                        fileId = uploadResponse.id;
                    }
                }

                // Append the new user message to the conversation history
                let userMessageWithFileContext = userMessage;
                if (fileId) {
                    userMessageWithFileContext = `Answer the questions based on the context of following file: [file: ${fileId}]. ${userMessage}`;
                }

                // Append the new user message to the conversation history
                // conversationHistory.push({role: "user", content: userMessage});
                conversationHistory.push({
                    role: "user",
                    content: userMessageWithFileContext,
                });

                // If the length of conversation history exceeds 10
                // remove the second oldest message
                // in order to keep the systemMessage intact
                // prettier-ignore
                if (conversationHistory.length > MAX_PAST_MESSAGE_FOR_CONTEXT + 1)
                 {
                    conversationHistory.splice(1, 1);
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

                if (fileId) {
                    replyText = replyText + " File ID: " + fileId;
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

    async getAttachmentUrl(file) {
        // and you may need to ensure you have permission to access the file
        // return file.contentUrl; // This should be the URL to the file
        return file.content.downloadUrl;
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
