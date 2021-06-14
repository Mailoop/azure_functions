var msal = require("@azure/msal-node");
function getToken() {
  return new Promise((resolve, reject) => {
    const msalConfig = {
      auth: {
        clientId: process.env["CLIENT_ID"],
        authority: `https://login.microsoftonline.com/${process.env["TENANT"]}`,
        clientSecret: process.env["CLIENT_SECRET"],
      },
    };
    // Create msal application object
    const cca = new msal.ConfidentialClientApplication(msalConfig);
    // With client credentials flows permissions need to be granted in the portal by a tenant administrator.
    // The scope is always in the format "<resource>/.default"
    const clientCredentialRequest = {
      scopes: ["https://graph.microsoft.com/.default"],
    };
    cca
      .acquireTokenByClientCredential(clientCredentialRequest)
      .then((response) => {
        resolve(response.accessToken);
      })
      .catch((error) => {
        reject(error);
      });
  });
}
module.exports = getToken
