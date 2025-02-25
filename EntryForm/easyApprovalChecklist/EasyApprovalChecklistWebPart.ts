import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,  
} from '@microsoft/sp-webpart-base';

import {
IPropertyPaneConfiguration,
  PropertyPaneTextField
}from '@microsoft/sp-property-pane'

import * as strings from 'EasyApprovalChecklistWebPartStrings';
import EasyApprovalChecklist from './components/EasyApprovalChecklist';
import { IEasyApprovalChecklistProps } from './components/IEasyApprovalChecklistProps';
import PNoteForms from './components/PNoteForm';
import PNoteFormsEdit from './components/PNoteFormEdit';
import PNoteDraft from './components/PNoteDraft';
import * as jQuery from 'jquery';

export interface IEasyApprovalChecklistWebPartProps {
  description: string;
}

export default class EasyApprovalChecklistWebPart extends BaseClientSideWebPart<IEasyApprovalChecklistWebPartProps> {
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
    debugger;
    let qstr=window.location.search.split('uid=');
    let qstrNid=window.location.search.split('Nid=');
    let qstrPid=window.location.search.split('Pid=');
 let qstring='';
 let qstringNid='';
 let qstringPid='';
 if(qstrNid.length>1){qstringNid=qstr[1];}
 if(qstrPid.length>1){qstringPid=qstr[1];}
   if(qstr.length>1){qstring=qstr[1];}
    const element: React.ReactElement<IEasyApprovalChecklistProps > = React.createElement(
      //(qstring=='')?PNoteForms:PNoteFormEditable,
      (qstring=='' && qstringNid=='' && qstringPid=='')?PNoteForms:(qstringNid =='' && qstringPid=='')?PNoteFormsEdit:PNoteDraft,
      {
        context: this.context,
        description: this.properties.description,
        siteUrl: this.context.pageContext.web.absoluteUrl,
      }
    );
    ReactDom.render(element, this.domElement);
    // ReactDom.unmountComponentAtNode(this.domElement);
    // ReactDom.render(element, this.domElement);

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
