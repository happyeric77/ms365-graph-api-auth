import axios, { AxiosRequestConfig, AxiosResponse } from "axios";
import { IDriveItem, IDriveItems, IDrives, IListItem, IListItemField, IListItems, ILists, ISites } from "./types";

interface ICallApi {
  endpoint: string;
  accessToken: string;
  method?: "get" | "post" | "patch" | "delete" | "put";
  params?: { [key: string]: string };
  data?: { [key: string]: any };
  headers?: { [key: string]: string };
}
class GraphApiQuery {
  constructor(private accessToken: string) {}

  async getSites(siteName?: string): Promise<ISites> {
    let sites: ISites = await this.callApi({
      endpoint: `v1.0/sites`,
      accessToken: this.accessToken,
      // params: siteName ? { search: siteName } : undefined,
    });
    if (siteName) sites = { ...sites, value: sites.value.filter((site) => site.name === siteName) };
    return sites;
  }
  async getListsInSite(siteId: string, listName?: string) {
    let lists: ILists = await this.callApi({
      endpoint: `v1.0/sites/${siteId}/lists`,
      accessToken: this.accessToken,
    });
    if (listName) lists = { ...lists, value: lists.value.filter((list) => list.name === listName) };
    return lists;
  }
  async getItemsInList(siteId: string, listId: string) {
    let items: IListItems = await this.callApi({
      endpoint: `v1.0/sites/${siteId}/lists/${listId}/items?expand=fields`, // OPTION: ?expand=fields(select=Column1,Column2) https://learn.microsoft.com/ja-jp/graph/api/listitem-list?view=graph-rest-1.0&tabs=http
      accessToken: this.accessToken,
    });
    return items;
  }
  async postCreateListItem(siteId: string, listId: string, fields: { [key: string]: string }) {
    const result: IListItemField = await this.callApi({
      endpoint: `v1.0/sites/${siteId}/lists/${listId}/items`,
      accessToken: this.accessToken,
      method: "post",
      headers: {
        "Content-Type": "application/json",
      },
      data: {
        fields: fields,
      },
    });
    return result;
  }
  async patchUpdateListItem(siteId: string, listId: string, itemId: string, fields: { [key: string]: string }) {
    const result: IListItem = await this.callApi({
      endpoint: `v1.0/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
      accessToken: this.accessToken,
      method: "patch",
      headers: {
        "Content-Type": "application/json",
      },
      data: fields,
    });
    return result;
  }

  async deleteListItem(siteId: string, listId: string, itemId: string) {
    const result: any = await this.callApi({
      endpoint: `v1.0/sites/${siteId}/lists/${listId}/items/${itemId}`,
      accessToken: this.accessToken,
      method: "delete",
    });
    return result; // RETURN HTTP/1.1 204 No Content
  }

  async getDrives(siteId: string, driveName?: string) {
    let drives: IDrives = await this.callApi({
      endpoint: `v1.0/sites/${siteId}/drives/`,
      accessToken: this.accessToken,
      method: "get",
      headers: {
        "Content-Type": "application/json",
      },
    });
    if (driveName) drives = { ...drives, value: drives.value.filter((site) => site.name === driveName) };
    return drives;
  }

  async getDriveItems(siteId: string, driveId: string) {
    let items: IDriveItems = await this.callApi({
      endpoint: `v1.0/sites/${siteId}/drives/${driveId}/root/children`,
      accessToken: this.accessToken,
      method: "get",
      headers: {
        "Content-Type": "application/json",
      },
    });
    return items;
  }

  async getDriveItemByFileName(siteId: string, driveId: string, itemName: string, loadContent?: boolean) {
    let item: IDriveItem = await this.callApi({
      endpoint: `v1.0/sites/${siteId}/drives/${driveId}/root:/${loadContent ? `${itemName}:/content` : itemName}`,
      accessToken: this.accessToken,
      method: "get",
      headers: {
        "Content-Type": "application/json",
      },
    });
    return item;
  }

  async getDriveItemById(siteId: string, driveId: string, itemId: string, loadContent?: boolean) {
    let item: IDriveItem = await this.callApi({
      endpoint: `v1.0/sites/${siteId}/drives/${driveId}/items/${loadContent ? `${itemId}/content` : itemId}`,
      accessToken: this.accessToken,
      method: "get",
      headers: {
        "Content-Type": "application/json",
      },
    });
    return item;
  }

  async putUploadDriveItem(siteId: string, driveId: string, fileName: string, data: any, folder?: string) {
    const driveItem = await this.callApi({
      endpoint: `v1.0/sites/${siteId}/drives/${driveId}/root:/${folder ? `${folder}` : ""}${fileName}:/content`,
      accessToken: this.accessToken,
      method: "put",
      data: data,
    });
    return driveItem;
  }

  async deleteDriveItem(siteId: string, driveId: string, itemId: string) {
    let item: IDriveItem = await this.callApi({
      endpoint: `v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}`,
      accessToken: this.accessToken,
      method: "delete",
    });
    return item;
  }

  private async callApi(configs: ICallApi, graphEndpoint?: string): Promise<any> {
    configs.method = configs.method ? configs.method : "get";
    const options: AxiosRequestConfig = {
      headers: {
        ...configs.headers,
        Authorization: `Bearer ${configs.accessToken}`,
      },
      params: configs.params,
    };

    console.log("request made to web API at: " + new Date().toString());

    try {
      let result: any;
      switch (configs.method) {
        case "get":
          result = (
            await axios.get(graphEndpoint ? graphEndpoint : "https://graph.microsoft.com/" + configs.endpoint, options)
          ).data;
          break;
        case "post":
          result = await (
            await axios.post(
              graphEndpoint ? graphEndpoint : "https://graph.microsoft.com/" + configs.endpoint,
              configs.data,
              options
            )
          ).data;
          break;
        case "patch":
          result = await (
            await axios.patch(
              graphEndpoint ? graphEndpoint : "https://graph.microsoft.com/" + configs.endpoint,
              configs.data,
              options
            )
          ).data;
          break;
        case "put":
          result = await (
            await axios.put(
              graphEndpoint ? graphEndpoint : "https://graph.microsoft.com/" + configs.endpoint,
              configs.data,
              options
            )
          ).data;
          break;
        case "delete":
          // This applies to both the Azure AD Graph and the Microsoft Graph. The only way to delete objects is using user delegated auth with a token from a user that has sufficient permissions to do so (generally an admin).
          result = await (
            await axios.delete(
              graphEndpoint ? graphEndpoint : "https://graph.microsoft.com/" + configs.endpoint,
              options
            )
          ).data;
          break;
        default:
          throw Error(`ERROR: callApi - does not support this method - ${configs.method}`);
      }

      return result;
    } catch (error) {
      console.log(error);
      return error;
    }
  }
}

export { GraphApiQuery };
