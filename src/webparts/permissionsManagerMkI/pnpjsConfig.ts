import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/security";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export const getSP = (context: WebPartContext) => {
    return spfi().using(SPFx(context));
};