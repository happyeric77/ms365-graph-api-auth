import axios, { AxiosRequestConfig } from "axios";

class ApiEndpoint {
  static users = "v1.0/users";
  static site = `v1.0/sites`;
  static listInSite = (id: string) => `v1.0/sites/${id}/lists`;
  static itemsInListInSite = (listId: string, siteId: string) => `v1.0/sites/${siteId}/lists/${listId}/items`;
}

interface ICallApi {
  endpoint: ApiEndpoint;
  accessToken: string;
  method?: "get" | "post";
  getParams?: { [key: string]: string };
  postData?: { [key: string]: string };
  headers?: { [key: string]: string };
}

const callApi = async (params: ICallApi): Promise<any> => {
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
        result = (await axios.get(process.env.GRAPH_ENDPOINT! + params.endpoint, options)).data;
        break;
      case "post":
        result = await (await axios.post(process.env.GRAPH_ENDPOINT! + params.endpoint, params.postData)).data;
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
