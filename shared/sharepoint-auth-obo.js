const axios = require('axios');
const qs = require('querystring');

// Function to obtain OBO token. This will take the access token in request header (scoped to this Function App) and generate a new token to use for Graph API
const getOboToken = async (userAccessToken) => {
    const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
    
    if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
        throw new Error('Missing environment variables for OBO token exchange');
    }
    
    const scope = 'https://graph.microsoft.com/.default';
    const oboTokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

    const params = {
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        assertion: userAccessToken,
        requested_token_use: 'on_behalf_of',
        scope: scope
    };

    try {
        const response = await axios.post(oboTokenUrl, qs.stringify(params), {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        });
        return response.data.access_token; // OBO token
    } catch (error) {
        const errorMsg =
            error.response?.data?.error_description ||
            error.response?.data?.error ||
            error.message ||
            'Unknown error occurred';
        throw new Error(errorMsg);
    }
};

// DISABLED: Service account authentication is not supported in OBO-only mode
// This ensures all requests are scoped to the authenticated user's permissions
const getServiceAccountToken = async () => {
    throw new Error('Service account authentication is not supported in OBO-only mode. User authentication required.');
    
    // Original service account code preserved but unreachable:
    const { TENANT_ID, CLIENT_ID, SERVICE_USERNAME, SERVICE_PASSWORD } = process.env;
    
    if (!TENANT_ID || !CLIENT_ID || !SERVICE_USERNAME || !SERVICE_PASSWORD) {
        throw new Error('Missing environment variables for Service Account token exchange');
    }
    
    const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

    const params = {
        client_id: CLIENT_ID,
        scope: 'https://graph.microsoft.com/.default',
        username: SERVICE_USERNAME,
        password: SERVICE_PASSWORD,
        grant_type: 'password'
    };

    try {
        const response = await axios.post(tokenUrl, qs.stringify(params), {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        });
        return response.data.access_token;
    } catch (error) {
        console.error('Error getting service account token:', error.response?.data || error.message);
        throw error;
    }
};

// PLACEHOLDER: Future user token validation will go here
// This function will validate the bearer token before attempting OBO exchange
// Validations to implement:
// - Check token expiration
// - Verify token issuer matches expected EntraID tenant
// - Validate required claims (aud, sub, etc.)
// - Check for required scopes/permissions
const validateUserToken = async (bearerToken) => {
    // TODO: Implement JWT validation
    // Example checks:
    // - Decode JWT header and payload
    // - Verify signature against EntraID public keys
    // - Check 'exp' claim is not expired
    // - Verify 'aud' matches this application's client ID
    // - Verify 'iss' matches expected tenant issuer
    // - Check for required claims
    
    if (!bearerToken) {
        throw new Error('Bearer token is required for authentication');
    }
    
    // For now, basic null check
    // TODO: Replace with proper JWT validation
    console.log('[validateUserToken] Token validation placeholder - implement JWT checks');
    return true;
};

// MODIFIED: OBO-only token resolution - no fallback to service account
// Returns an OBO token if bearerToken is provided, otherwise throws error
const resolveGraphToken = async (bearerToken) => {
    if (!bearerToken) {
        console.error('[resolveGraphToken] No bearer token provided - authentication required');
        throw new Error('User authentication required. No bearer token provided in Authorization header.');
    }
    
    console.log('[resolveGraphToken] Bearer token provided, attempting OBO exchange');
    
    // PLACEHOLDER: Add user token validation before OBO exchange
    try {
        await validateUserToken(bearerToken);
    } catch (validationError) {
        console.error('[resolveGraphToken] Token validation failed:', validationError.message);
        throw new Error(`Token validation failed: ${validationError.message}`);
    }
    
    // Proceed with OBO token exchange
    try {
        const oboToken = await getOboToken(bearerToken);
        console.log('[resolveGraphToken] OBO token successfully acquired');
        return oboToken;
    } catch (oboError) {
        console.error('[resolveGraphToken] OBO token exchange failed:', oboError.message);
        throw new Error(`OBO token exchange failed: ${oboError.message}`);
    }
};

module.exports = {
    getOboToken,
    getServiceAccountToken, // Exported but throws error - maintained for backwards compatibility
    resolveGraphToken,
    validateUserToken // Export for testing purposes
};