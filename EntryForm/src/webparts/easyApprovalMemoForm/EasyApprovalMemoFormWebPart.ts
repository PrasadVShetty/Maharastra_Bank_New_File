import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'EasyApprovalMemoFormWebPartStrings';
import EasyApprovalMemoForm from './components/EasyApprovalMemoForm';
import EasyApprovalMemoEditForm from './components/EasyApprovalMemoEditForm';
import { IEasyApprovalMemoFormProps } from './components/IEasyApprovalMemoFormProps';
import * as jQuery from 'jquery';
export interface IEasyApprovalMemoFormWebPartProps {
  description: string;
}

export default class EasyApprovalMemoFormWebPart extends BaseClientSideWebPart<IEasyApprovalMemoFormWebPartProps> {
  public async onInit(): Promise<void> {
    let qstring='1';
      jQuery(".ms-compositeHeader").hide();
      jQuery(".o365cs-navMenuButton").hide();
      jQuery('div[data-automationid="SiteHeader"]').remove();
      jQuery('div [data-sp-feature-tag="Comments"]').remove();
      jQuery('div [id="SuiteNavPlaceHolder"]').hide();
      jQuery('div .commandBarWrapper').hide();
      jQuery('div [data-automation-id="pageHeader"]').hide();
      jQuery('.commandBarWrapper').hide();
    }
  public render(): void {
    let qstr=window.location.search.split('uid=');
 let qstring='';
   if(qstr.length>1){qstring=qstr[1];}
    const element: React.ReactElement<IEasyApprovalMemoFormProps > = React.createElement(
      (qstring=='' )?EasyApprovalMemoForm:EasyApprovalMemoEditForm,
      {
        context: this.context,
        description: this.properties.description,
        siteUrl: this.context.pageContext.web.absoluteUrl,
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
