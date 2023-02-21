import * as msal from "@azure/msal-node";

/**
 *
 * @param clientId
 * @param clientSecret
 * @param aadEndpoint AZURE ACTIVE DIRECTORY // https://login.microsoftonline.com/
 * @param graphEndpoint https://graph.microsoft.com/
 * @returns
 */
async function getToken(
  clientId: string,
  clientSecret: string,
  tenantId: string,
  aadEndpoint?: string,
  graphEndpoint?: string
) {
  const msalConfig = {
    auth: {
      clientId,
      authority: aadEndpoint ? aadEndpoint : "https://login.microsoftonline.com/" + tenantId,
      clientSecret,
    },
  };

  const tokenRequest = {
    scopes: [graphEndpoint ? graphEndpoint : "https://graph.microsoft.com/" + ".default"], // e.g. 'https://graph.microsoft.com/.default'
  };
  const cca = new msal.ConfidentialClientApplication(msalConfig);

  return await cca.acquireTokenByClientCredential(tokenRequest);
}

export { getToken };
