const {ActivityHandler, MessageFactory} = require("botbuilder");
const {OpenAI} = require("openai");
const axios = require("axios");
const pdfParse = require("pdf-parse");
const mammoth = require("mammoth");
const XLSX = require("xlsx");

// Update supported file types to include Excel formats
const supportedFileTypes = ["pdf", "doc", "docx", "xlsx", "xls"];

// New function to handle Excel files
async function extractTextFromExcel(fileBuffer) {
    try {
        const workbook = XLSX.read(fileBuffer, {
            type: "buffer",
            cellDates: true,
            cellNF: false,
            cellText: false,
            compression: true,
        });

        let fullText = [];

        for (const sheetName of workbook.SheetNames) {
            const worksheet = workbook.Sheets[sheetName];

            // Convert to JSON with header row
            const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                raw: true,
                dateNF: "yyyy-mm-dd",
            });

            if (jsonData.length === 0) continue;

            fullText.push(`Table: ${sheetName}\n`);

            // Get headers
            const headers = jsonData[0].map((h) =>
                h ? h.toString().trim() : ""
            );

            // Create markdown table
            const tableRows = [];

            // Add header row
            tableRows.push(`| ${headers.join(" | ")} |`);

            // Add separator row
            tableRows.push(`| ${headers.map(() => "---").join(" | ")} |`);

            // Add data rows
            for (let i = 1; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (!row || row.length === 0) continue;

                // Format each cell value
                const formattedRow = headers.map((_, index) => {
                    const value = row[index];
                    if (value === undefined || value === null) return "";
                    if (value instanceof Date)
                        return value.toISOString().split("T")[0];
                    if (typeof value === "number") return value.toString();
                    return value.toString().trim();
                });

                tableRows.push(`| ${formattedRow.join(" | ")} |`);
            }

            // Add the table to fullText
            fullText.push(tableRows.join("\n"));
            fullText.push("\n---\n");
        }

        // Add a summary for GPT
        const tableSummary =
            "The above data is presented in table format. Each table represents a sheet from the Excel file. " +
            "You can perform calculations, analyze trends, and answer questions about the data using the tabular information provided.";

        fullText.push(tableSummary);

        return fullText.join("\n");
    } catch (error) {
        console.error("Excel processing error:", error);
        return null;
    }
}

// Extract text based on file type
async function extractTextFromDocument(fileBuffer, fileType) {
    switch (fileType) {
        case "pdf":
            const pdfData = await pdfParse(fileBuffer);
            return pdfData.text;
        case "doc":
        case "docx":
            const result = await mammoth.extractRawText({buffer: fileBuffer});
            return result.value;
        case "xlsx":
        case "xls":
            // return extractTextFromExcel(fileBuffer);
            const excelExtractedData = await extractTextFromExcel(fileBuffer);
            if (!excelExtractedData) {
                console.log("NULL extracted!");

                return "";
            }
            return excelExtractedData;
        default:
            throw new Error(`Unsupported file type: ${fileType}`);
    }
}

// Rest of the utility functions remain the same
function splitIntoChunks(text, maxChunkSize = 1000) {
    const chunks = [];
    const sentences = text.split(/[.!?]+/);
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

function cosineSimilarity(vecA, vecB) {
    const dotProduct = vecA.reduce((sum, a, i) => sum + a * vecB[i], 0);
    const normA = Math.sqrt(vecA.reduce((sum, a) => sum + a * a, 0));
    const normB = Math.sqrt(vecB.reduce((sum, b) => sum + b * b, 0));
    return dotProduct / (normA * normB);
}

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
            let isValidAttachment = false;
            try {
                let conversationHistory = this.conversations[userId] || [];

                if (conversationHistory.length === 0) {
                    const systemMessage = {
                        role: "system",
                        content:
                            "You are an intelligent assistant bot, named BookWorm, at the company BooksTime. You can assist Bookkeepers, Senior Accountants, IT Department, Senior Managers and Client Service Advisors at BooksTime with their queries to the best of your ability. You can provide sales support and management insights. You can help Bookstimers (staffs at BooksTime) in analyzing financial statements, proofreading proposals for grammar errors, upselling opportunities, finding answers to questions in bank statements, help them draft emails and much much more. When analyzing Excel data, provide specific insights and calculations based on the numerical data provided. For financial data, include relevant metrics and trends when appropriate. If someone asks you, what is your name, you tell them your name is BookWorm. When working with Excel data, the data will be presented in markdown table format,you can perform calculations on the numerical values, You can analyze trends and patterns in the data, you can compare values across different columns and rows, you should provide specific numerical insights when relevant, for financial data, include relevant metrics and calculations. Always show your calculations when performing numerical analysis.",
                    };
                    conversationHistory.push(systemMessage);
                }

                const attachments = context.activity.attachments;

                if (attachments && attachments[0]) {
                    try {
                        const attachment = attachments[0];
                        const downloadUrl = await this.getAttachmentUrl(
                            attachment
                        );

                        if (downloadUrl) {
                            const fileType = await this.getFileType(attachment);

                            if (!supportedFileTypes.includes(fileType)) {
                                await context.sendActivity(
                                    MessageFactory.text(
                                        "Sorry, I can only process PDF, DOC, DOCX, XLSX, and XLS files at the moment."
                                    )
                                );
                                return;
                            }
                            console.log("UPLOADED FILE TYPE: " + fileType);

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

                            isValidAttachment = true;

                            // Extract text from document
                            const documentText = await extractTextFromDocument(
                                fileBuffer,
                                fileType
                            );
                            await this.processDocument(userId, documentText);

                            // Acknowledge successful processing of Excel file
                            // if (fileType === "xlsx" || fileType === "xls") {
                            //     await context.sendActivity(
                            //         MessageFactory.text(
                            //             "I've processed your Excel file. You can now ask questions about the data!"
                            //         )
                            //     );
                            // }
                        }
                    } catch (error) {
                        console.error("Error processing file:", error);
                        await context.sendActivity(
                            MessageFactory.text(
                                "I encountered an error while processing your file. Please make sure it's a valid document file."
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
                if (isValidAttachment) {
                    if (relevantContext) {
                        messageWithContext = `Using the relevant document context: "${relevantContext}" \n\nUser question: ${userMessage}`;
                    }
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
                    "Sorry, I cannot answer your question at the moment. Please try again later and if the issue persists, contact the IT Team at BooksTime."
                );
            }

            await next();
        });

        // Rest of the class implementation remains the same...
        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            const welcomeText =
                "Hello BooksTimer! I am BookWorm, an intelligent conversational Chatbot.\nHow can I help you today?";

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
        const chunks = splitIntoChunks(documentText);
        const embeddings = await this.openai.embeddings.create({
            model: "text-embedding-ada-002",
            input: chunks,
        });
        this.documentEmbeddings[userId] = {
            chunks,
            embeddings: embeddings.data,
        };
    }

    async getRelevantContext(userId, question) {
        const docData = this.documentEmbeddings[userId];
        if (!docData) return null;

        const questionEmbedding = await this.openai.embeddings.create({
            model: "text-embedding-ada-002",
            input: question,
        });

        const topIndices = findMostSimilarChunks(
            docData.embeddings,
            questionEmbedding
        );
        const relevantText = topIndices
            .map((idx) => docData.chunks[idx])
            .join("\n\n");

        return relevantText;
    }

    async getAttachmentUrl(attachment) {
        return attachment.content.downloadUrl;
    }

    async getFileType(attachment) {
        return attachment.content.fileType;
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

module.exports.OpenAIBot = OpenAIBot;
