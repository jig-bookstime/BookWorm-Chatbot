const {ActivityHandler, MessageFactory} = require("botbuilder");
const {OpenAI} = require("openai");

const axios = require("axios");
const pdfParse = require("pdf-parse"); // A library to parse PDFs
const mammoth = require("mammoth"); // Library to parse DOCX files

class OpenAIBot extends ActivityHandler {
    constructor() {
        super();

        // Initialize OpenAI client
        this.openai = new OpenAI({
            apiKey: process.env.OPENAI_API_KEY,
        });

        // Initialize an object to track conversation history for each user
        this.conversations = {};

        // Handle incoming messages
        this.onMessage(async (context, next) => {
            const userMessage = context.activity.text;
            const userId = context.activity.from.id; // Get user ID to store their conversation context

            try {
                // Check for attachments (file uploads)
                if (
                    context.activity.attachments &&
                    context.activity.attachments.length > 0
                ) {
                    const attachment = context.activity.attachments[0]; // Assuming you want the first attachment
                    const fileUrl = attachment.contentUrl; // The URL of the uploaded file

                    // Fetch the file using axios
                    const response = await axios.get(fileUrl, {
                        responseType: "arraybuffer",
                    });

                    // Determine the file type and parse accordingly
                    const contentType = attachment.contentType;

                    let extractedText = "";

                    if (contentType === "application/pdf") {
                        // If the file is a PDF, parse it using pdf-parse
                        const parsedData = await pdfParse(response.data);
                        extractedText = parsedData.text;
                    } else if (
                        contentType ===
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    ) {
                        // If the file is a DOCX, parse it using mammoth
                        extractedText = await parseDocx(response.data);
                    } else {
                        await context.sendActivity(
                            `File type not supported for parsing: ${contentType}`
                        );
                        return;
                    }

                    // Generate a question based on the extracted text
                    const generatedQuestion = `Based on the following content, can you summarize or provide insights?\n\n${extractedText.slice(
                        0,
                        500
                    )}...`; // You can adjust the question as per your requirements

                    // Retrieve the conversation history for the user, or initialize it if not exists
                    let conversationHistory = this.conversations[userId] || [];

                    // If conversation is starting, add a system message to provide context to OpenAI
                    if (conversationHistory.length === 0) {
                        const systemMessage = {
                            role: "system",
                            content:
                                "You are an intelligent assistant bot, named BookWorm, at the company BooksTime. You can assist with various tasks including providing insights from documents.",
                        };
                        conversationHistory.push(systemMessage); // Add initial system context
                    }

                    // Append the generated question to the conversation history
                    conversationHistory.push({
                        role: "user",
                        content: generatedQuestion,
                    });

                    // If the length of conversation history exceeds 10, remove the first element
                    if (conversationHistory.length > 10) {
                        conversationHistory.splice(1, 1); // Removes the first (oldest) message
                    }

                    // Send a typing indicator before making the OpenAI request
                    await context.sendActivity({type: "typing"});

                    // Get the reply from OpenAI based on the generated question
                    const replyText = await this.getOpenAIResponse(
                        conversationHistory
                    );

                    // Save the bot's response in the conversation history
                    conversationHistory.push({
                        role: "assistant",
                        content: replyText,
                    });

                    // Update the conversation history for the user
                    this.conversations[userId] = conversationHistory;

                    // Send the OpenAI response back to the user
                    await context.sendActivity(
                        MessageFactory.text(replyText, replyText)
                    );
                } else {
                    // Handle regular user messages
                    let conversationHistory = this.conversations[userId] || [];

                    if (conversationHistory.length === 0) {
                        const systemMessage = {
                            role: "system",
                            content:
                                "You are an intelligent assistant bot, named BookWorm...",
                        };
                        conversationHistory.push(systemMessage); // Add initial system context
                    }

                    // Append the new user message to the conversation history
                    conversationHistory.push({
                        role: "user",
                        content: userMessage,
                    });

                    // If the length of conversation history exceeds 10, remove the first element
                    if (conversationHistory.length > 10) {
                        conversationHistory.splice(1, 1); // Removes the first (oldest) message
                    }

                    // Send a typing indicator before making the OpenAI request
                    await context.sendActivity({type: "typing"});

                    // Get the reply from OpenAI
                    const replyText = await this.getOpenAIResponse(
                        conversationHistory
                    );

                    // Save the bot's response in the conversation history
                    conversationHistory.push({
                        role: "assistant",
                        content: replyText,
                    });

                    // Update the conversation history for the user
                    this.conversations[userId] = conversationHistory;

                    // Send the OpenAI response back to the user
                    await context.sendActivity(
                        MessageFactory.text(replyText, replyText)
                    );
                }
            } catch (error) {
                console.error("Error processing message:", error);
                await context.sendActivity(
                    "Sorry, I couldn't process your request at the moment."
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

    // Helper function to parse DOCX files
    async parseDocx(data) {
        const result = await mammoth.extractRawText({buffer: data});
        return result.value; // This is the extracted text from the Word document
    }

    // Function to get response from OpenAI
    async getOpenAIResponse(conversationHistory) {
        try {
            const response = await this.openai.chat.completions.create({
                model: process.env.OPENAI_MODEL,
                messages: conversationHistory, // Pass the entire conversation history
            });

            // console.log(response);

            // Return the bot's response (chatbot's message)
            return response.choices[0].message.content;
        } catch (error) {
            throw new Error(`OpenAI API error: ${error.message}`);
        }
    }
}

module.exports.OpenAIBot = OpenAIBot;
