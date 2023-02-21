import * as msal from "@azure/msal-node";

async function getToken() {
  const msalConfig = {
    auth: {
      clientId: process.env.CLIENT_ID!,
      authority: process.env.AAD_ENDPOINT! + process.env.TENANT_ID,
      clientSecret: process.env.CLIENT_SECRET,
    },
  };

  const tokenRequest = {
    scopes: [process.env.GRAPH_ENDPOINT + ".default"], // e.g. 'https://graph.microsoft.com/.default'
  };
  const cca = new msal.ConfidentialClientApplication(msalConfig);

  return await cca.acquireTokenByClientCredential(tokenRequest);
}

export { getToken };
