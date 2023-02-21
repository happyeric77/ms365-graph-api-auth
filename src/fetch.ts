import axios, { AxiosRequestConfig } from "axios";
import { IListItems, ILists } from "./types";

class ApiEndpoint {
  static users = "v1.0/users";
  static sites = `v1.0/sites`;
  static site = (id: string) => `v1.0/sites/${id}`;
  static listsInSite = (id: string) => `v1.0/sites/${id}/lists`;
  static listInSite = (listId: string, siteId: string) => `v1.0/sites/${siteId}/lists/${listId}`;
  static itemsInListInSite = (listId: string, siteId: string) => `v1.0/sites/${siteId}/lists/${listId}/items`;
  static itemInListInSite = (listId: string, siteId: string, itemId: string) =>
    `v1.0/sites/${siteId}/lists/${listId}/items/${itemId}`;
}

interface ICallApi {
  endpoint: ApiEndpoint;
  accessToken: string;
  method?: "get" | "post";
  getParams?: { [key: string]: string };
  postData?: { [key: string]: string };
  headers?: { [key: string]: string };
}

const callApi = async (params: ICallApi, graphEndpoint?: string): Promise<any> => {
  params.method = params.method ? params.method : "get";
  const options: AxiosRequestConfig = {
    headers: {
      ...params.headers,
      Authorization: `Bearer ${params.accessToken}`,
    },
    params: params.getParams,
  };

  console.log("request made to web API at: " + new Date().toString());

  try {
    let result: any;
    switch (params.method) {
      case "get":
        result = (
          await axios.get(graphEndpoint ? graphEndpoint : "https://graph.microsoft.com/" + params.endpoint, options)
        ).data;
        break;
      case "post": // TODO
        result = await (
          await axios.post(
            graphEndpoint ? graphEndpoint : "https://graph.microsoft.com/" + params.endpoint,
            params.postData,
            options
          )
        ).data;
        break;
      default:
    }

    return result;
  } catch (error) {
    console.log(error);
    return error;
  }
};

export { callApi, ApiEndpoint };
