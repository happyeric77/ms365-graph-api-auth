export interface IListItems {
  "@odata.context": string;
  value: IListItem[];
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
