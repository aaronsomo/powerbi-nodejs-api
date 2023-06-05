const getAccessToken = async function () {
  const config = require(__dirname + "/../config/config.json");

  const msal = require("@azure/msal-node");

  const msalConfig = {
    auth: {
      clientId: config.clientId,
      authority: `${config.authorityUrl}${config.tenantId}`,
    },
  };

  if (config.authenticationMode.toLowerCase() === "masteruser") {
    const clientApplication = new msal.PublicClientApplication(msalConfig);

    const usernamePasswordRequest = {
      scopes: [config.scopeBase],
      username: config.pbiUsername,
      password: config.pbiPassword,
    };

    return clientApplication.acquireTokenByUsernamePassword(
      usernamePasswordRequest
    );
  }

  if (config.authenticationMode.toLowerCase() === "serviceprincipal") {
    msalConfig.auth.clientSecret = config.clientSecret;
    const clientApplication = new msal.ConfidentialClientApplication(
      msalConfig
    );

    const clientCredentialRequest = {
      scopes: [config.scopeBase],
    };

    return clientApplication.acquireTokenByClientCredential(
      clientCredentialRequest
    );
  }
};

module.exports.getAccessToken = getAccessToken;
