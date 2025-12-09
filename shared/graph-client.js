const { Client } = require('@microsoft/microsoft-graph-client');

//// --------- ENVIRONMENT CONFIGURATION AND INITIALIZATION ---------
// Function to initialize Microsoft Graph client
const initGraphClient = (accessToken) => {
    return Client.init({
        authProvider: (done) => {
            done(null, accessToken); // Pass the access token for Graph API calls
        }
    });
};

module.exports = {
    initGraphClient
}