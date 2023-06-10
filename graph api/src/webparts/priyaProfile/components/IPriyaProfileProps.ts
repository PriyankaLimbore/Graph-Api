import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IPriyaProfileProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  MyUserInformation:string;
  context:WebPartContext;
}
