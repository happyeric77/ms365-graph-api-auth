require("dotenv").config();
import { getAccessToken, GraphApiQuery } from "../src";
import fs from "fs";

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
    const mySite = await query.getSites(process.env.SITE_NAME);
    console.log({ sites: mySite.value.map((site) => ({ name: site.name, displayName: site.displayName })) });
    // FETCH ALL SITES : await query.getSites();

    //2. GET DRIVES
    console.log("\n\n\n2. GET DRIVES IN CERTAIN SHAREPOINT SITE ***");
    const myDrives = await query.getDrives(mySite.value[0].id, process.env.DRIVE_NAME);
    console.log(myDrives);

    //3.1. GET ALL ITEMS IN DRIVE
    console.log("\n\n\n3.1. GET ALL ITEMS IN CERTAIN DRIVE ***");
    const myItems = await query.getDriveItems(mySite.value[0].id, myDrives.value[0].id);
    console.log(myItems);

    // 3.2 GET CERTAIN ITEM IN DRIVE BY FILE NAME
    console.log("\n\n\n3.2. GET CERTAIN ITEM IN DRIVE BY FILE NAME ***");
    const myItem1 = await query.getDriveItemByFileName(mySite.value[0].id, myDrives.value[0].id, "TainanView2.jpg");
    console.log(myItem1);

    // 3.3 GET CERTIAN ITEM IN DRIVE BY FILE ID
    console.log("\n\n\n3.3. GET CERTAIN ITEM IN DRIVE BY FILE ID ***");
    const myItem2 = await query.getDriveItemById(mySite.value[0].id, myDrives.value[0].id, myItems.value[0].id);
    console.log(myItem2);

    // 4. UPLOAD DRIVE ITEM
    console.log("\n\n\n4. UPLOAD DRIVE ITEM ***");
    const filePath = process.env.UPLOAD_FILE_PATH;
    const fileData = fs.readFileSync(filePath!);
    /**
     * @fileName w/o slash
     * @folder with trailing slash
     */
    const createdItem = await query.putUploadDriveItem(
      mySite.value[0].id,
      myDrives.value[0].id,
      `file${Date.now()}.jpg`,
      fileData,
      `folder${Date.now()}/`
    );
    console.log(createdItem);

    // 5. DELETE DRIVE ITEM
    console.log("\n\n\n5. DELETE DRIVE ITEM ***");
    const res = await query.deleteDriveItem(mySite.value[0].id, myDrives.value[0].id, myItems.value[0].id);
    console.log(res);
  } catch (error) {
    console.log(error);
  }
})();
