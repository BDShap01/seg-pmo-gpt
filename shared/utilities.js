// shared/utilities.js

const { resolveGraphToken } = require('./sharepoint-auth');

/**
 * Truncates a token to show only the beginning and end.
 * Useful for debugging without leaking the full token.
 */
function _truncateToken(token, segmentLength = 15) {
    if (typeof token !== 'string') return '';
    if (token.length <= segmentLength * 2) return token;
    return (
        token.substring(0, segmentLength) +
        '...' +
        token.substring(token.length - segmentLength)
    );
}

/**
 * Generic response formatter for Azure Functions.
 * Automatically sets content-type and optionally includes error message.
 */
function sendResponse(context, debugData, statusCode = 200, errorMessage = null) {
    const body = errorMessage ? { debug: debugData, error: errorMessage } : debugData;

    context.res = {
        status: statusCode,
        headers: { 'Content-Type': 'application/json' },
        body
    };
    return context.res;
}

/**
 * Extracts bearer token (if present), obtains a valid Graph token (OBO or service),
 * enriches the debug object, and handles error responses.
  */
async function getGraphToken(req, context, debug) {
    let bearerToken;

    if (req.headers.authorization) {
        debug.authHeaders = true;
        bearerToken = req.headers.authorization.split(' ')[1];
        debug.bearerToken = _truncateToken(bearerToken);
    } else {
        debug.authHeaders = false;
    }

    try {
        const graphToken = await resolveGraphToken(bearerToken);
        debug.graphToken = _truncateToken(graphToken);
        return graphToken;
    } catch (error) {
        sendResponse(context, debug, 500, `Graph token error: ${error.message}`);
        throw error;
    }
}

module.exports = {
    getGraphToken,
    sendResponse    
};
