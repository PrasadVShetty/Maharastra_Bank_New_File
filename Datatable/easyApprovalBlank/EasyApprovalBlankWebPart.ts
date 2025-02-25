import * as React from 'react';
import * as ReactDom from 'react-dom';
//import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'EasyApprovalBlankWebPartStrings';
import EasyApprovalBlank from './components/EasyApprovalBlank';
import { IEasyApprovalBlankProps } from './components/IEasyApprovalBlankProps';
import "@pnp/polyfill-ie11";  
import * as pnp from "sp-pnp-js";
import { sp } from '@pnp/sp';  
import * as jQuery from 'jquery'; 
export interface IEasyApprovalBlankWebPartProps {
  description: string;
}

export default class EasyApprovalBlankWebPart extends BaseClientSideWebPart<IEasyApprovalBlankWebPartProps> {
  protected onInit(): Promise<void> {  
  debugger;
    jQuery(".ms-compositeHeader").hide();
  jQuery(".o365cs-navMenuButton").hide();
  jQuery('div[data-automationid="SiteHeader"]').remove();
 // jQuery('#O365_SearchBoxContainer_container').remove();
  jQuery('.commandBarWrapper').hide();
  jQuery('div [id="SuiteNavPlaceHolder"]').hide();
  jQuery('div [id="O365_SearchBoxContainer_container"]').hide();
 
  jQuery('div .commandBarWrapper').hide();
  jQuery('div [data-automation-id="pageHeader"]').hide();
  jQuery(("a[aria-label^='Get the mobile app']")).remove();
  setTimeout(() => {
    const searchBox = document.getElementById('O365_SearchBoxContainer_container');
    if (searchBox) {
        searchBox.remove();
    }
}, 1000);
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
    const element: React.ReactElement<IEasyApprovalBlankProps > = React.createElement(
      EasyApprovalBlank,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
