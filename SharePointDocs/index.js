const { Buffer } = require('buffer');
const path = require('path');

//Import utilities
const { getGraphToken, sendResponse} = require('../shared/utilities');

//Import client authorization
// const { getOboToken, getServiceAccountToken, getAccessToken } = require('../shared/sharepoint-auth');

//Import Graph client initialization
const { initGraphClient } = require('../shared/graph-client');

//// --------- DOCUMENT PROCESSING ---------
// Function to fetch drive item content and convert to text
const getDriveItemContent = async (client, driveId, itemId, name) => {
    try {
        // const fileType = path.extname(name).toLowerCase();
        // the below files types are the ones that are able to be converted to PDF to extract the text. See https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format?view=graph-rest-1.0&tabs=http
        // const allowedFileTypes = ['.pdf', '.doc', '.docx', '.odp', '.ods', '.odt', '.pot', '.potm', '.potx', '.pps', '.ppsx', '.ppsxm', '.ppt', '.pptm', '.pptx', '.rtf'];
        // filePath changes based on file type, adding ?format=pdf to convert non-pdf types to pdf for text extraction, so all files in allowedFileTypes above are converted to pdf
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
   
    const searchTerm = req.query.searchTerm || (req.body && req.body.searchTerm);
    debug.searchTerm = searchTerm;

    let graphToken;
    try {
        graphToken = await getGraphToken(req, context, debug);
    } catch (error) {
        return sendResponse(context, debug, 500, `Error generating Graph token: ${error.message}`);
    }

    // Initialize the Graph Client using the initGraphClient function defined above
    let client = initGraphClient(graphToken);
    // this is the search body to be used in the Microsft Graph Search API: https://learn.microsoft.com/en-us/graph/search-concept-files
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
    debug.requestBody = requestBody

    try { 
        // This is where we are doing the search
        const list = await client.api('/search/query').post(requestBody);

        const processList = async () => {
            // Loop through each search response to retreive desired file data
            const results = [];

            await Promise.all(list.value[0].hitsContainers.map(async (container) => {
                for (const hit of container.hits) {
                    if (hit.resource["@odata.type"] === "#microsoft.graph.driveItem") {
                        const { name, id } = hit.resource;
                        const driveId = hit.resource.parentReference.driveId;
                        // const fileData = await getDriveItemContent(client, driveId, id, name);
                        const fileData = await getDriveItemMetadata(client, driveId, id, name);
                        results.push(fileData)
                    }
                }
            }));

            return results;
        };
        let results;
        if (list.value[0].hitsContainers[0].total == 0) {
            // Return no results found to the API if the Microsoft Graph API returns no results
            results = { message: 'No results found', openaiFileResponse: [] };
        } else {
            // If the Microsoft Graph API does return results, then run processList to iterate through.
            results = await processList();
            results = {'openaiFileResponse': results}
            // results.sort((a, b) => a.rank - b.rank);
        }

        debug.results = results;
        return sendResponse(context, debug);
        // return sendResponse(context, results); // final success return

    } catch (error) {
        return sendResponse(context, debug, 500, `Error performing search or processing results: ${error.message}`);
    }
};
