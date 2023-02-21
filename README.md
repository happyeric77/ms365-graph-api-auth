# Microsoft 365 GraphAPI auth tool

This package allows you to get the authenticated Azure app to CRUD Sharepoint by Microsoft Graph API.

The project is built on top of [Azure auth example](https://github.com/Azure-Samples/ms-identity-javascript-nodejs-console)

# How to use

## Register a Azure client app. Check the [tutorial here](TODO)

## Install [The Package](TODO)

```
npm install ms365-graph-api-auth
// or
yarn add ms365-graph-api-auth
```

## Import relevant modules to get started

```
import { ApiEndpoint, callApi, IListItems, ILists, getAccessToken, IListItem } from "../src";

// Credentials
const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;

(async () => {
  try {
    const authResponse = await getAccessToken(clientId!, clientSecret!, tenantId!);
    if (!authResponse) return;
    const res = await callApi({
      endpoint: ApiEndpoint.sites,
      accessToken: authResponse.accessToken,
      getParams: { search: "<your-site-name>" },
    });
    const siteId = res.value[0].id;
    const lists: ILists = await callApi({
      endpoint: ApiEndpoint.listsInSite(siteId),
      accessToken: authResponse.accessToken,
    });
    const sortedLists: { id: string; name: string }[] = lists.value.map((list) => ({
      id: list.id,
      name: list.name,
    }));
    const wptListId = sortedLists.find((list) => list.name === "<your-list-name>")!.id;

    const listItems: IListItems = await callApi({
      endpoint: ApiEndpoint.itemsInListInSite(wptListId, siteId),
      accessToken: authResponse.accessToken,
    });
    // console.log(listItems.value.map((item) => item));
    const firstItem: IListItem = await callApi({
      endpoint: ApiEndpoint.itemInListInSite(wptListId, siteId, listItems.value[0].id),
      accessToken: authResponse.accessToken,
    });
    console.log(firstItem.fields);
  } catch (error) {
    console.log(error);
  }
})();
```

# TODO (Open to contribute)

- Only support part of graphAPI (Users/ Sites/ Lists)
- Only support graphAPI "get" methods (Read status only)
