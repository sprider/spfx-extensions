import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IEmailFileUrlComponentProps {
  siteUrl: string;
  listTitle: string;
  itemId: number;
  fileName: string;
  fileRelativePath: string;
  spHttpClient: SPHttpClient;
}