import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'EasyApprovalDocumentsViewWebPartStrings';
import EasyApprovalDocumentsView from './components/EasyApprovalDocumentsView';
import { IEasyApprovalDocumentsViewProps } from './components/IEasyApprovalDocumentsViewProps';

export interface IEasyApprovalDocumentsViewWebPartProps {
  description: string;
  context:WebPartContext; 
}
import "@pnp/polyfill-ie11";  
import * as pnp from "sp-pnp-js";
import { sp } from '@pnp/sp';  
import * as jQuery from 'jquery'; 
export default class EasyApprovalDocumentsViewWebPart extends BaseClientSideWebPart<IEasyApprovalDocumentsViewWebPartProps> {
  protected onInit(): Promise<void> {  
    jQuery(".ms-compositeHeader").hide();
    jQuery(".o365cs-navMenuButton").hide();
    jQuery('div[data-automationid="SiteHeader"]').remove();
    jQuery('.commandBarWrapper').hide();
    // jQuery('div [id="SuiteNavPlaceHolder"]').hide();
    
    jQuery('div .commandBarWrapper').hide();
    jQuery('div [data-automation-id="pageHeader"]').hide();
    setTimeout(() => {
      jQuery('#O365_MainLink_Settings').remove();
    }, 3000);
  pnp.setup({
    spfxContext: this.context
  });
  return new Promise<void>((resolve: () => void, reject: (error?: any) => void): void => {  
    sp.setup({  
      sp: {  
        headers: {  
          "Accept": "application/json; odata=nometadata"  
        }  
      }  
    });  
    resolve();  
  });  
} 
  public render(): void {
    const element: React.ReactElement<IEasyApprovalDocumentsViewProps > = React.createElement(
      EasyApprovalDocumentsView,
      {
        description: this.properties.description,
        context: this.context 
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
