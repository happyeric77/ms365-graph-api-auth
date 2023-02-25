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

export interface IDrives {
  "@odata.context": string;
  value: IDrive[];
}

export interface IDrive {
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#drives/$entity";
  createdDateTime: string;
  description: string;
  id: string;
  lastModifiedDateTime: string;
  name: string;
  webUrl: string;
  driveType: string;
  createdBy: [Object];
  lastModifiedBy: [Object];
  owner: [Object];
  quota: [Object];
}

export interface IDriveItems {
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#sites('51tvnm.sharepoint.com%2Cbdaafbd4-08f0-421c-877e-4b87dc459784%2C22c9a737-a9b6-4279-885e-2231d39a0627')/drives('b%211PuqvfAIHEKHfkuH3EWXhDenySK2qXlCiF4iMdOaBid8qWThmL8QQaczJWXd_EOO')/root/children";
  value: IDriveItem[];
}

export interface IDriveItem {
  "@microsoft.graph.downloadUrl": string;
  createdDateTime: string;
  eTag: string;
  id: string;
  lastModifiedDateTime: string;
  name: string;
  webUrl: string;
  cTag: string;
  size: number;
  createdBy: [Object];
  lastModifiedBy: [Object];
  parentReference: [Object];
  file: [Object];
  fileSystemInfo: [Object];
  image: {};
  shared: [Object];
}
