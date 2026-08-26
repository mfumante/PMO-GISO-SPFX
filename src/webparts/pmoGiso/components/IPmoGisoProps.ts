import type { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IPmoGisoProps {
  // Passato per gli step successivi (accesso alle liste SharePoint).
  context: WebPartContext;
  userDisplayName: string;
}
