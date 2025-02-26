import { WebPartContext } from '@microsoft/sp-webpart-base';
import { BaseComponentContext } from '@microsoft/sp-component-base';
export interface IEasyApprovalDocumentFormProps extends BaseComponentContext{
  description: string;
  context: WebPartContext;
  siteUrl: string;
}
