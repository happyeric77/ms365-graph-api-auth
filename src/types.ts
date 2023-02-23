export interface ISites {
  "@odata.context": string;
  value: ISite[];
}

export interface ISite {
  createdDateTime: string;
  id: string;
  lastModifiedDateTime: string;
  name: string;
  webUrl: string;
  displayName: string;
  root: Object;
  siteCollection: [Object];
}
export interface IListItems {
  "@odata.context": string;
  value: IListItem[];
}

export interface IListItemField {
  [key: string]: any;
}
export interface IListItem {
  "@odata.etag": string;
  createdDateTime: string;
  eTag: string;
  id: string;
  lastModifiedDateTime: string;
  webUrl: string;
  createdBy: [Object];
  lastModifiedBy: [Object];
  parentReference: [Object];
  contentType: [Object];
  fields?: IListItemField; // ONLY SHOWS WHEN QUERY SINGLE ITEM BY ID
}

export interface ILists {
  "@odata.context": string;
  value: IList[];
}

export interface IList {
  "@odata.etag": string;
  createdDateTime: string;
  description: string;
  eTag: string;
  id: string;
  lastModifiedDateTime: string;
  name: string;
  webUrl: string;
  displayName: string;
  createdBy: [Object];
  parentReference: [Object];
  list: [Object];
}
