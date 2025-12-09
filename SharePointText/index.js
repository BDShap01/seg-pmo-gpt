const pdfParse = require('pdf-parse');
const { Buffer } = require('buffer');
const path = require('path');
const { OpenAI } = require("openai");

// Import shared functions
const { getGraphToken, sendResponse} = require('../shared/utilities');

// Import Graph client initialization
const { initGraphClient } = require('../shared/graph-client');

// --------- DOCUMENT PROCESSING ---------
// Function to fetch drive item content and convert to text
const getDriveItemContent = async (client, driveId, itemId, name) => {
    try {
        const fileType = path.extname(name).toLowerCase();
        
        // Handle .txt files first - they're simple text content
        if (fileType === '.txt') {
            const response = await client.api(`/drives/${driveId}/items/${itemId}/content`).get();
            return response;
        } 
        
        // Handle .csv files - return as UTF-8 string
        else if (fileType === '.csv') {
            const response = await client.api(`/drives/${driveId}/items/${itemId}/content`).getStream();
            let chunks = [];
            for await (let chunk of response) {
                chunks.push(chunk);
            }
            let buffer = Buffer.concat(chunks);
            return buffer.toString('utf-8');
        } 
        
        // Handle files that can be converted to PDF for text extraction
        // See https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format?view=graph-rest-1.0&tabs=http
        else if (['.pdf', '.doc', '.docx', '.odp', '.ods', '.odt', '.pot', '.potm', '.potx', 
                  '.pps', '.ppsx', '.ppsxm', '.ppt', '.pptm', '.pptx', '.rtf'].includes(fileType)) {
            // For PDFs, get content directly. For others, convert to PDF first
            const filePath = `/drives/${driveId}/items/${itemId}/content${fileType === '.pdf' ? '' : '?format=pdf'}`;
            const response = await client.api(filePath).getStream();
            
            let chunks = [];
            for await (let chunk of response) {
                chunks.push(chunk);
            }
            let buffer = Buffer.concat(chunks);
            
            // Extract text from the PDF
            const pdfContents = await pdfParse(buffer);
            return pdfContents.text;
        } 
        
        else {
            return 'Unsupported File Type';
        }
     
    } catch (error) {
        console.error('Error fetching drive content:', error);
        throw new Error(`Failed to fetch content for ${name}: ${error.message}`);
    }
};

// Function to get relevant parts of text using gpt-4o-mini
const getRelevantParts = async (text, query) => {
    try {
        // We use your OpenAI key to initialize the OpenAI client
        const openAIKey = process.env["OPENAI_API_KEY"];
        const openai = new OpenAI({
            apiKey: openAIKey,
        });
        const response = await openai.chat.completions.create({
            // Using gpt-4o-mini due to speed to prevent timeouts. You can tweak this prompt as needed
            model: "gpt-4o-mini",
            messages: [
                {"role": "system", "content": "You are a helpful assistant that finds relevant content in text based on a query. You only return the relevant sentences, and you return a maximum of 10 sentences"},
                {"role": "user", "content": `Based on this question: **"${query}"**, get the relevant parts from the following text:*****\n\n${text}*****. If you cannot answer the question based on the text, respond with 'No information provided'`}
            ],
            // using temperature of 0 since we want to just extract the relevant content
            temperature: 0,
            // using max_tokens of 1000, but you can customize this based on the number of documents you are searching. 
            max_tokens: 1000
        });
        return response.choices[0].message.content;
    } catch (error) {
        console.error('Error with OpenAI:', error);
        return 'Error processing text with OpenAI: ' + error.message;
    }
};

// --------- AZURE FUNCTION LOGIC ---------
// Below is what the Azure Function executes
module.exports = async function (context, req) {
    const debug = {};

    const query = req.query.query || (req.body && req.body.query);
    debug.query = query;
    const searchTerm = req.query.searchTerm || (req.body && req.body.searchTerm);
    debug.searchTerm = searchTerm;
    //return sendResponse(context, debug);

    let graphToken;
    try {
        graphToken = await getGraphToken(req, context, debug);
    } catch (error) {
        return sendResponse(context, debug, 500, `Error generating Graph token: ${error.message}`);
    }
    // return sendResponse(context, debug);

    // Initialize the Graph Client using the initGraphClient function defined above
    let client = initGraphClient(graphToken);
    
    // Search body for Microsoft Graph Search API
    // https://learn.microsoft.com/en-us/graph/search-concept-files
    const requestBody = {
        requests: [
            {
                entityTypes: ['driveItem'],
                query: {
                    queryString: searchTerm
                },
                from: 0,
                // the below is set to summarize the top 10 search results from the Graph API, but can configure based on your documents. 
                size: 10
            }
        ]
    };
    debug.requestBody = requestBody;
    // return sendResponse(context, debug);

    try { 
        // Function to tokenize content (word-based approximation)
        const tokenizeContent = (content) => {
            return content.split(/\s+/);
        };
    
        // Function to break tokens into chunks for gpt-4o-mini
        // Using 7500 words as max to account for ~0.75 word-to-token ratio
        // This gives us ~10k actual GPT tokens per window with safety margin
        const breakIntoTokenWindows = (tokens) => {
            const tokenWindows = [];
            const maxWindowTokens = 7500; // Conservative estimate for ~10k actual tokens
            let startIndex = 0;
    
            while (startIndex < tokens.length) {
                const window = tokens.slice(startIndex, startIndex + maxWindowTokens);
                tokenWindows.push(window);
                startIndex += maxWindowTokens;
            }
    
            return tokenWindows;
        };
        
        // Perform the search
        const list = await client.api('/search/query').post(requestBody);
    
        const processList = async () => {
            // This will go through and for each search response, grab the contents of the file and summarize with gpt-4o-mini
            const results = [];
    
            // Process each container (typically just one)
            for (const container of list.value[0].hitsContainers) {
                // Use Promise.allSettled to handle failures gracefully
                const hitPromises = container.hits.map(async (hit) => {
                    if (hit.resource["@odata.type"] === "#microsoft.graph.driveItem") {
                        const { name, id } = hit.resource;
                        const webUrl = hit.resource.webUrl.replace(/\s/g, "%20");
                        const rank = hit.rank;
                        const driveId = hit.resource.parentReference.driveId;
                        
                        try {
                            const contents = await getDriveItemContent(client, driveId, id, name);
                            
                            if (contents !== 'Unsupported File Type') {
                                // Tokenize content
                                const tokens = tokenizeContent(contents);
                                
                                // Break into chunks
                                const tokenWindows = breakIntoTokenWindows(tokens);
                                
                                // Process each chunk and combine results
                                const relevantPartsPromises = tokenWindows.map(window => 
                                    getRelevantParts(window.join(' '), query)
                                );
                                const relevantParts = await Promise.all(relevantPartsPromises);
                                const combinedResults = relevantParts.join('\n');
                                
                                return { name, webUrl, rank, contents: combinedResults, status: 'success' };
                            } else {
                                return { name, webUrl, rank, contents: 'Unsupported File Type', status: 'unsupported' };
                            }
                        } catch (error) {
                            console.error(`Error processing ${name}:`, error);
                            return { name, webUrl, rank, contents: `Error: ${error.message}`, status: 'error' };
                        }
                    }
                    return null;
                });
                
                // Wait for all hits to process (or fail)
                const settledResults = await Promise.allSettled(hitPromises);
                
                // Extract successful results
                settledResults.forEach(result => {
                    if (result.status === 'fulfilled' && result.value !== null) {
                        results.push(result.value);
                    }
                });
            }
    
            return results;
        };
               
        let results;
        if (list.value[0].hitsContainers[0].total == 0) {
            // Return no results found to the API if the Microsoft Graph API returns no results
            results = { message: 'No results found', openaiFileResponse: [] };
        } else {
            // If the Microsoft Graph API does return results, then run processList to iterate through.
            results = await processList();
            // Sort by rank (lower is better)
            results.sort((a, b) => a.rank - b.rank);
        }
        
        debug.results = results;
        return sendResponse(context, debug);
        // return sendResponse(context, results); // final success return
    } catch (error) {
        return sendResponse(context, debug, 500, `Error performing search or processing results: ${error.message}`);
    }
};