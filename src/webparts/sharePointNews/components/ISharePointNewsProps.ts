import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISharePointNewsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}

export interface IState {
  AllNews: NewsItem[];
  ShowForm: boolean;
  IsLoading: boolean;
  SiteID: string;

  // Form
  Title: string;
  Description: string;
  Summary: string;
}

export interface NewsItem {
  Title: string;
  Description: string;
  Summary: string;

  ID: number;
  Created: string;
  PostLink: {
    Url: string;
  };
}
