// Configuration PnPjs et journalisation.
// - getSP(context) : renvoie une instance PnPjs authentifiée avec les droits délégués de l'utilisateur connecté.
//   Les imports "@pnp/sp/..." ci-dessous activent sélectivement les fonctionnalités utilisées (listes, sécurité, utilisateurs, groupes).
// - logAction(...) : écrit une entrée dans la liste SharePoint "Logs" (colonnes Action, Bibliotheque, Cible, Details),
//   ce qui déclenche le flux Power Automate d'alerte par email. Un échec de log ne bloque JAMAIS l'action principale.

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
  action: "Consultation" | "Ajout" | "Suppression" | "Initialisation",
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

// ---------- Feedback ----------
// Adresses du service informatique qui reçoivent les retours envoyés depuis la page "Votre avis".
// >>> Pour modifier ou ajouter un destinataire, c'est ICI : une adresse par ligne, entre guillemets, séparées par des virgules.
// Elles sont copiées dans la colonne "Destinataires" de chaque retour, que le flux Power Automate utilise comme champ "À".
export const FEEDBACK_DESTINATAIRES: string[] = [
  "leopold.roux-decorzent@cofidest.com",
  "service.informatique@cofidest.com"
];

export const FEEDBACK_LIST = "Feedback";

export type FeedbackCategorie = "Retour" | "Suggestion" | "Question";

// Écrit un retour dans la liste SharePoint "Feedback" (colonnes : Categorie, Message, Destinataires), ce qui
// déclenche le flux Power Automate d'envoi de l'email. L'auteur est fourni par la colonne native "Créé par".
// Contrairement à logAction, l'écriture EST l'action principale : une erreur est remontée à l'appelant.
export const sendFeedback = async (
  sp: ReturnType<typeof getSP>,
  categorie: FeedbackCategorie,
  objet: string,
  message: string
): Promise<void> => {
  await sp.web.lists.getByTitle(FEEDBACK_LIST).items.add({
    Title: objet,
    Categorie: categorie,
    Message: message,
    Destinataires: FEEDBACK_DESTINATAIRES.join(";")
  });
};