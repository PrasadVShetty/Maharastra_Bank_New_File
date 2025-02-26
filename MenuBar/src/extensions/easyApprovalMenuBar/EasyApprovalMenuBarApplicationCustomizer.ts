import { override } from '@microsoft/decorators';
//import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';
//import { Dialog } from '@microsoft/sp-dialog';

//import * as strings from 'EasyApprovalMenuBarApplicationCustomizerStrings';
import { MegaMenu, IMegaMenuProps } from './components/MegaMenu';
import { MegaMenuService } from './service/MegaMenuService';
import { TopLevelMenu } from './model/TopLevelMenu';
//import * as jQuery from 'jquery';

//const LOG_SOURCE: string = 'EasyApprovalMenuBarApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IEasyApprovalMenuBarApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class EasyApprovalMenuBarApplicationCustomizer
  extends BaseApplicationCustomizer<IEasyApprovalMenuBarApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;
  // @override
  // public onInit(): Promise<void> {
  //   MegaMenuService.getMenuItems(this.context.pageContext.site.absoluteUrl)
  //     .then((topLevelMenus: TopLevelMenu[]) => {
  //       this._renderPlaceHolders(topLevelMenus);
        
        
  //     }).catch((error) => { console.log("Error in loading the Mega Menu" + error); });
  //     return Promise.resolve();
  // }

@override
public async onInit(): Promise<void> {
  console.log("EasyApprovalMenuBarApplicationCustomizer Initialized");

  // 🔹 Ensure `pageContext` is defined
  if (!this.context || !this.context.pageContext || !this.context.pageContext.site) {
    console.error("SPFx Context is undefined! Cannot retrieve site URL.");
    return;
  }

  const siteUrl = this.context.pageContext.site.absoluteUrl;
  console.log("Site URL:", siteUrl);

  try {
    const topLevelMenus: TopLevelMenu[] = await MegaMenuService.getMenuItems(siteUrl);
    this._renderPlaceHolders(topLevelMenus);
  } catch (error) {
    console.error("Error loading the Mega Menu:", error);
  }

  return Promise.resolve();
}


  private _renderPlaceHolders(menuItems: any): void {
    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose });

      if (!this._topPlaceholder) {
        console.error('The expected placeholder (Top) was not found.');
        return;
      }
      const element: React.ReactElement<IMegaMenuProps> = React.createElement(
        MegaMenu,
        {
          topLevelMenuItems: menuItems
        });

      ReactDom.render(element, this._topPlaceholder.domElement);

    }
  
  }

  private _onDispose(): void {
    console.log('[MegaMenuApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
