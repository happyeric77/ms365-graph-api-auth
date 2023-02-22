# Microsoft 365 GraphAPI auth tool

This package allows you to get the authenticated Azure app to CRUD Sharepoint by Microsoft Graph API.

The project is built on top of [Azure auth example](https://github.com/Azure-Samples/ms-identity-javascript-nodejs-console)

# How to use

## Register a Azure client app. Check the [tutorial here](TODO)

## Install [The Package](https://www.npmjs.com/package/ms365-graph-api-auth)

```
npm install ms365-graph-api-auth
// or
yarn add ms365-graph-api-auth
```

## Import relevant modules to get started

```
require("dotenv").config();
import { getAccessToken, GraphApiQuery } from "../src";

// Credentials
const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;

(async () => {
  try {
    const authResponse = await getAccessToken(clientId!, clientSecret!, tenantId!);
    if (!authResponse) return;
    const query = new GraphApiQuery(authResponse.accessToken);
    //1. GET SHAREPOINT SITES
    console.log("1. GET SHAREPOINT SITES ***");
    // SEARCH CERTAIN SITE BY NAME.
    const mySite = await query.getSites("<your-sharepoint-site-name>");
    console.log({ sites: mySite.value.map((site) => ({ name: site.name, displayName: site.displayName })) });
    // FETCH ALL SITES : await query.getSites();

    //2. GET LISTS IN CERTAIN SHAREPOINT SITE
    console.log("\n\n\n2. GET LISTS IN CERTAIN SHAREPOINT SITE ***");
    // SEARCH CERTAIN LIST BY NAME
    const myList = await query.getListsInSite(mySite.value[0].id, "Work progress tracker");
    console.log(myList.value.map((list) => ({ name: list.name, displayName: list.displayName })));

    //3. GET ITEMS IN LIST
    console.log("\n\n\n3. GET ITEMS IN LIST ***");
    const items = await query.getItemsInList(mySite.value[0].id, myList.value[0].id);
    const itemInfos = items.value.map((item) => item.fields);
    console.log(itemInfos.map((info) => info?.["Title"]));
    console.log(items);

    //4. CRUD ITEM IN LIST // https://learn.microsoft.com/ja-jp/graph/api/resources/listitem?view=graph-rest-1.0
    //4.1 CREATE
    console.log("\n\n\n4.1. CREATE NEW ITEM IN LIST ***");
    const newItem = {
      Title: `My new item create at ${Date.now()}`,
      Description: `Created by ms365-graph-api test`,
    };
    const createdItem = await query.postCreateListItem(mySite.value[0].id, myList.value[0].id, newItem);
    console.log(createdItem);

    //4.2 UPDATE // https://learn.microsoft.com/ja-jp/graph/api/listitem-update?view=graph-rest-1.0&tabs=http
    console.log("\n\n\n4.2 UPDATE ITEM IN LIST ***");
    const data = {
      Notes: `Updated by ms365-graph-api test`,
    };
    const updatedItem = await query.patchUpdateListItem(mySite.value[0].id, myList.value[0].id, createdItem.id, data);
    console.log(updatedItem);

    //4.3 DELETE // https://learn.microsoft.com/ja-jp/graph/api/listitem-delete?view=graph-rest-1.0&tabs=http
    // It will not be working using application permission.
    // The only way to delete objects is using user delegated auth with a token from a user that has sufficient permissions to do so (generally an admin).
    console.log("\n\n\n4.3 DELETE ITEM IN LIST ***");
    // const res = await query.deleteListItem(mySite.value[0].id, myList.value[0].id, items.value[0].id);
    // console.log(res);
  } catch (error) {
    console.log(error);
  }
})();


```

# TODO (Open to contribute)

- Only support part of graphAPI (Sites/ Lists/ Items)
