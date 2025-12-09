const { Buffer } = require('buffer');
const path = require('path');

//Import utilities
const { getGraphToken, sendResponse} = require('../shared/utilities');

//Import Graph client initialization
const { initGraphClient } = require('../shared/graph-client');

//// --------- DOCUMENT PROCESSING ---------
// Function to fetch drive item content and convert to text
const getDriveItemContent = async (client, driveId, itemId, name) => {
    try {
        const filePath = `/drives/${driveId}/items/${itemId}`;
        const downloadPath = filePath + `/content`
      
        try{
            const fileStream = await client.api(downloadPath).getStream();
            let chunks = [];
            for await (let chunk of fileStream) {
                chunks.push(chunk);
            }
            const base64String = Buffer.concat(chunks).toString('base64');
            const file = await client.api(filePath).get();
            const mime_type = file.file.mimeType;
            const name = file.name;
            return {"name":name, "mime_type":mime_type, "content":base64String}
        } catch (error) {
            console.error(`Could not download content for '${name}':`, error.message);
            return { name, error: 'Unable to download content' };
        }
    } catch (error) {
        console.error('Error fetching drive content:', error);
        throw new Error(`Failed to fetch content for ${name}: ${error.message}`);
    }
};

// Function to fetch drive item metadata
const getDriveItemMetadata = async (client, driveId, itemId, name) => {
    try {
        const filePath = `/drives/${driveId}/items/${itemId}`;
        const file = await client.api(filePath).get();

        return {
            title: file.name,
            url: file.webUrl,
            lastModified: file.lastModifiedDateTime,
            modifiedBy: file.lastModifiedBy?.user?.displayName || 'Unknown'
        };
    } catch (error) {
        console.error(`Failed to fetch metadata for '${name}' (itemId: ${itemId}, driveId: ${driveId}):`, error.message);
        return {
            title: name || null,
            url: null,
            lastModified: null,
            modifiedBy: null,
            error: 'Unable to retrieve metadata'
        };
    }
};

//// --------- AZURE FUNCTION LOGIC ---------
// Below is what the Azure Function executes
module.exports = async function (context, req) {
    const debug = {};
   
    try {
        const searchTerm = req.query.searchTerm || (req.body && req.body.searchTerm);
        debug.searchTerm = searchTerm;
        debug.step = "1-params-parsed";

        let graphToken;
        try {
            graphToken = await getGraphToken(req, context, debug);
            debug.step = "2-token-acquired";
        } catch (error) {
            debug.step = "2-token-failed";
            return sendResponse(context, debug, 500, `Error generating Graph token: ${error.message}`);
        }

        // Initialize the Graph Client
        let client = initGraphClient(graphToken);
        debug.step = "3-client-initialized";
        
        const requestBody = {
            requests: [
                {
                    entityTypes: ['driveItem'],
                    query: {
                        queryString: searchTerm
                    },
                    from: 0,
                    size: 10
                }
            ]
        };
        debug.requestBody = requestBody;
        debug.step = "4-request-body-created";

        // This is where we are doing the search
        const list = await client.api('/search/query').post(requestBody);
        debug.step = "5-search-completed";
        debug.totalHits = list.value[0].hitsContainers[0].total;

        const processList = async () => {
            const results = [];

            // FIXED: Proper Promise.all usage
            for (const container of list.value[0].hitsContainers) {
                const hitPromises = container.hits.map(async (hit) => {
                    if (hit.resource["@odata.type"] === "#microsoft.graph.driveItem") {
                        const { name, id } = hit.resource;
                        const driveId = hit.resource.parentReference.driveId;
                        return await getDriveItemMetadata(client, driveId, id, name);
                    }
                    return null;
                });
                
                const hitResults = await Promise.all(hitPromises);
                results.push(...hitResults.filter(r => r !== null));
            }

            return results;
        };
        
        let results;
        if (list.value[0].hitsContainers[0].total == 0) {
            debug.step = "6-no-results";
            results = { message: 'No results found', openaiFileResponse: [] };
        } else {
            debug.step = "6-processing-results";
            results = await processList();
            debug.step = "7-processing-complete";
            results = {'openaiFileResponse': results}
        }

        debug.results = results;
        debug.step = "8-success";
        return sendResponse(context, debug);

    } catch (error) {
        debug.step = "ERROR";
        debug.errorMessage = error.message;
        debug.errorStack = error.stack;
        console.error("FATAL ERROR in SharePointDocs:", error);
        return sendResponse(context, debug, 500, `Error: ${error.message}`);
    }
};