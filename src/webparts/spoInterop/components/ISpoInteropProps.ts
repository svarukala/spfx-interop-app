import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISpoInteropProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userLoginName: string;
  userEmail: string;
  spoContext: WebPartContext;
}
