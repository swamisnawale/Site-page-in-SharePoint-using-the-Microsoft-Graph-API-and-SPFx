import { spfi, SPFI, SPFx } from "@pnp/sp"; // npm i @pnp/sp@3.8.0
import { WebPartContext } from "@microsoft/sp-webpart-base";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/fields";
// import "@pnp/sp/items/get-all";
import "@pnp/sp/lists/web";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups";

import "@pnp/sp/lists/web";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
let sp: SPFI | any = null;

export const getSP = (context: WebPartContext) => {
  if (sp === null && context != null) {
    sp = spfi().using(SPFx(context));
  }
  return sp;
};
