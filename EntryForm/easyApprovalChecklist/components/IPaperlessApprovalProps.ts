import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface IPaperlessApprovalProps {
  description: string;
  context: WebPartContext;
  siteUrl: string;
}
