const {ActivityHandler, MessageFactory} = require("botbuilder");
const {OpenAI} = require("openai");
const axios = require("axios");
const pdfParse = require("pdf-parse");

// Utility function to split text into chunks
function splitIntoChunks(text, maxChunkSize = 1000) {
    const chunks = [];
    const sentences = text.split(/[.!?]+/); // Split by sentence boundaries
    let currentChunk = "";

    for (const sentence of sentences) {
        if (currentChunk.length + sentence.length > maxChunkSize) {
            chunks.push(currentChunk.trim());
            currentChunk = "";
        }
        currentChunk += sentence + ". ";
    }

    if (currentChunk) {
        chunks.push(currentChunk.trim());
    }

    return chunks;
}

// Calculate cosine similarity between two vectors
function cosineSimilarity(vecA, vecB) {
    const dotProduct = vecA.reduce((sum, a, i) => sum + a * vecB[i], 0);
    const normA = Math.sqrt(vecA.reduce((sum, a) => sum + a * a, 0));
    const normB = Math.sqrt(vecB.reduce((sum, b) => sum + b * b, 0));
    return dotProduct / (normA * normB);
}

// Find most similar chunks using embeddings
function findMostSimilarChunks(
    chunkEmbeddings,
    questionEmbedding,
    numChunks = 3
) {
    const similarities = chunkEmbeddings.map((embedding) =>
        cosineSimilarity(
            embedding.embedding,
            questionEmbedding.data[0].embedding
        )
    );

    // Get indices of top N similar chunks
    const topIndices = similarities
        .map((sim, idx) => ({sim, idx}))
        .sort((a, b) => b.sim - a.sim)
        .slice(0, numChunks)
        .map((item) => item.idx);

    return topIndices;
}

class OpenAIBot extends ActivityHandler {
    constructor() {
        super();

        this.openai = new OpenAI({
            apiKey: process.env.OPENAI_API_KEY,
        });

        this.conversations = {};
        this.documentEmbeddings = {};

        this.onMessage(async (context, next) => {
            const userMessage = context.activity.text;
            const userId = context.activity.from.id;

            try {
                let conversationHistory = this.conversations[userId] || [];

                if (conversationHistory.length === 0) {
                    const systemMessage = {
                        role: "system",
                        content:
                            "You are an intelligent assistant bot, named BookWorm, at the company BooksTime. You can assist Bookkeepers, Senior Accountants, IT Department, Senior Managers and Client Service Advisors at BooksTime with their queries to the best of your ability. You can provide sales support and management insights. You can help Bookstimers (staffs at BooksTime) in analyzing financial statements, proofreading proposals for grammar errors, upselling opportunities, finding answers to questions in bank statements, help them draft emails and much much more. If someone asks you, what is your name, you tell them your name is BookWorm.",
                    };
                    conversationHistory.push(systemMessage);
                }

                const attachments = context.activity.attachments;

                if (attachments && attachments[0]) {
                    try {
                        const downloadUrl = await this.getAttachmentUrl(
                            attachments[0]
                        );
                        if (downloadUrl) {
                            const fileResponse = await axios.get(downloadUrl, {
                                responseType: "arraybuffer",
                            });

                            const fileBuffer = Buffer.from(fileResponse.data);
                            const fileSizeInMB =
                                fileBuffer.length / (1024 * 1024);

                            if (fileSizeInMB > 20) {
                                await context.sendActivity(
                                    MessageFactory.text(
                                        "The file is too large. Please upload a smaller file (under 20MB)."
                                    )
                                );
                                return;
                            }

                            // Extract text from PDF and process document
                            const documentText = await extractTextFromPDF(
                                fileBuffer
                            );
                            await this.processDocument(userId, documentText);
                        }
                    } catch (error) {
                        console.error("Error processing file:", error);
                        await context.sendActivity(
                            MessageFactory.text(
                                "I encountered an error while processing your file. Please make sure it's a valid PDF document."
                            )
                        );
                        return;
                    }
                }

                // Get relevant context if document exists
                const relevantContext = await this.getRelevantContext(
                    userId,
                    userMessage
                );

                // Prepare the message with document context if available
                let messageWithContext = userMessage;
                if (relevantContext) {
                    messageWithContext = `Using the following relevant document context: "${relevantContext}" \n\nUser question: ${userMessage}`;
                }

                conversationHistory.push({
                    role: "user",
                    content: messageWithContext,
                });

                if (
                    conversationHistory.length >
                    process.env.MAX_PAST_MESSAGE_FOR_CONTEXT + 1
                ) {
                    conversationHistory.splice(1, 1);
                }

                await context.sendActivity({type: "typing"});

                const replyText = await this.getOpenAIResponse(
                    conversationHistory
                );

                conversationHistory.push({
                    role: "assistant",
                    content: replyText,
                });

                this.conversations[userId] = conversationHistory;

                await context.sendActivity(
                    MessageFactory.text(replyText, replyText)
                );
            } catch (error) {
                console.error("Error while processing message:", error);
                await context.sendActivity(
                    "Sorry, I cannot answer your question at the moment. Please try again later or contact the IT Team at BooksTime."
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

    async processDocument(userId, documentText) {
        // Split document into chunks
        const chunks = splitIntoChunks(documentText);

        // Get embeddings for all chunks
        const embeddings = await this.openai.embeddings.create({
            model: "text-embedding-ada-002",
            input: chunks,
        });

        // Store chunks and embeddings for the user
        this.documentEmbeddings[userId] = {
            chunks,
            embeddings: embeddings.data,
        };
    }

    async getRelevantContext(userId, question) {
        const docData = this.documentEmbeddings[userId];
        if (!docData) return null;

        // Get embedding for the question
        const questionEmbedding = await this.openai.embeddings.create({
            model: "text-embedding-ada-002",
            input: question,
        });

        // Find most similar chunks
        const topIndices = findMostSimilarChunks(
            docData.embeddings,
            questionEmbedding
        );

        // Combine relevant chunks
        const relevantText = topIndices
            .map((idx) => docData.chunks[idx])
            .join("\n\n");

        return relevantText;
    }

    async getAttachmentUrl(file) {
        return file.content.downloadUrl;
    }

    async getOpenAIResponse(conversationHistory) {
        try {
            const response = await this.openai.chat.completions.create({
                model: process.env.OPENAI_MODEL,
                messages: conversationHistory,
                temperature: 0.7,
                max_tokens: 1000,
            });

            return response.choices[0].message.content;
        } catch (error) {
            throw new Error(`OpenAI API error: ${error.message}`);
        }
    }
}

async function extractTextFromPDF(fileBuffer) {
    try {
        const data = await pdfParse(fileBuffer);
        return data.text;
    } catch (error) {
        throw new Error(`Error extracting text from PDF: ${error.message}`);
    }
}

module.exports.OpenAIBot = OpenAIBot;
