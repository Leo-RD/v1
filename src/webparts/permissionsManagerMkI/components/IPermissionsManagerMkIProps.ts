// Propriétés (props) du composant React principal : uniquement le contexte SPFx, nécessaire pour initialiser PnPjs.

import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IPermissionsManagerMkIProps {
  context: WebPartContext;
}