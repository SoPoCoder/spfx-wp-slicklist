import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx  } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/fields";
import "@pnp/sp/items/get-all";

// create an instance of SharePoint Factory Interface for use in the project
// for more information see https://pnp.github.io/pnpjs/getting-started/
let _sp: SPFI;
export const getSP = (context: WebPartContext): SPFI => {
    if (context)
        _sp = spfi().using(SPFx(context));
    return _sp;
};