import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/security/list";
import "@pnp/sp/security/web";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export const getSP = (context: WebPartContext) => {
  return spfi().using(SPFx(context));
};

export const logAction = async (
  sp: ReturnType<typeof getSP>,
  action: "Consultation" | "Ajout" | "Suppression",
  bibliotheque: string,
  cible: string = "",
  details: string = ""
): Promise<void> => {
  try {
    await sp.web.lists.getByTitle("Logs").items.add({
      Title: `${action} - ${bibliotheque}`,
      Action: action,
      Bibliotheque: bibliotheque,
      Cible: cible,
      Details: details
    });
  } catch (e) {
    // On ne bloque jamais l'action principale si le log échoue
    console.error("Erreur lors de l'écriture du log", e);
  }
}; 