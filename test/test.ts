require("dotenv").config();
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
      getParams: { search: "colorfullife" },
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
    const wptListId = sortedLists.find((list) => list.name === "Work progress tracker")!.id;

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
