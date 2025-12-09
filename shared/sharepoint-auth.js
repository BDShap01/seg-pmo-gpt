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

// Function to obtain ServiceAccount token token. This will take the access token in request header (scoped to this Function App) and generate a new token to use for Graph API
const getServiceAccountToken = async () => {
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

// Returns an OBO token if bearerToken is provided, otherwise falls back to service account
const resolveGraphToken = async (bearerToken) => {
    return bearerToken
    ? await getOboToken(bearerToken)
    : await getServiceAccountToken();
};

module.exports = {
    getOboToken,
    getServiceAccountToken,
    resolveGraphToken
}