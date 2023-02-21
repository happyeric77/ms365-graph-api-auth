require("dotenv").config();

import { ApiEndpoint, callApi } from "./fetch";
import { IListItems, ILists } from "./types";
const auth = require("./auth");

async function main() {
  try {
    const authResponse = await auth.getToken();
    const res = await callApi({
      endpoint: ApiEndpoint.site,
      accessToken: authResponse.accessToken,
      getParams: { search: "colorfullife" },
    });
    const siteId = res.value[0].id;
    const lists: ILists = await callApi({
      endpoint: ApiEndpoint.listInSite(siteId),
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
    // console.log(listItems);
  } catch (error) {
    console.log(error);
  }
}

main();
