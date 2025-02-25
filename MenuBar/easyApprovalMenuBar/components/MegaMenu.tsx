import * as React from 'react';
//import { useState } from "react";
// import { withResponsiveMode} from '@fluentui/react/lib/utilities/decorators/withResponsiveMode';
//import { useResponsiveMode } from '@fluentui/react';
import { SPComponentLoader } from '@microsoft/sp-loader';  
import styles from './MegaMenu.module.scss';
// SPComponentLoader.loadCss('/sites/EasyApproval/SiteAssets/css/styles.css'); 
SPComponentLoader.loadCss('/sites/EasyApproval/SiteAssets/css/styles.css');  
import { TopLevelMenu as TopLevelMenuModel } from '../model/TopLevelMenu';
import { CommandBar } from '@fluentui/react/lib/CommandBar';
import { IContextualMenuItem, ContextualMenuItemType } from '@fluentui/react/lib/ContextualMenu';
import { Icon } from '@fluentui/react/lib/Icon';
const logo: any = require('../Images/BOM-Caps.jpg');
const home: any = require('../Images/BOM-Home.jpg');
const logout: any = require('../Images/BOM-Logout.jpg');

export interface IMegaMenuProps {
    topLevelMenuItems: TopLevelMenuModel[];
}

export interface IMegaMenuState {
}

export interface IMegaMenuItems {
    identity: string;
    name: string;
    id: number;
    child: IMegaMenuItems[];
}

//@useResponsiveMode
export class MegaMenu extends React.Component<IMegaMenuProps, IMegaMenuState> {
    
    public state={
        isNavExpanded:false
    }
    constructor(props : any) {
        super(props);
        this.state = {
            isNavExpanded:false
        };
    }
    

    public render(): React.ReactElement<IMegaMenuProps> {
debugger;
        const commandBarItems: IContextualMenuItem[] = this.props.topLevelMenuItems.map((i) => {
            return (this.projectMenuItem(i, ContextualMenuItemType.Header));
        });

        debugger;

        return (
           
            <div className={`ms-bgColor-neutralLighter ms-fontColor-white ${styles.app}`}>
                <div className={` ${styles.divAtt}`}></div>
                <div className={`ms-bgColor-neutralLighter ms-fontColor-white ${styles.top}`}>
                <img src={logo} style={{width:"50px",height:"42px",marginLeft:"10px",marginRight:"10px"}}></img>
              <div className={styles.mobClass}> <a href="/sites/BOMCompliance_UAT"> <img src={home} style={{width:"35px",height:"35px",marginLeft:"10px",marginRight:"10px"}}></img></a></div>
                <CommandBar
                        // className={styles.commandBar + ' '+ {this.isNavExpanded ? "navigation-menu expanded": "commandBar"}}
                        className={(this.state.isNavExpanded ? "commandBar expanded": "commandBar")}
                        isSearchBoxVisible={false}
                        elipisisAriaLabel='More options'
                        items={commandBarItems}
                    />
                     <div>
                        <button className="hamburger" onClick={() => {
                            if(this.state.isNavExpanded===false)
                            {
                                this.setState({isNavExpanded: true })
                            }
                            else if(this.state.isNavExpanded===true){
                                this.setState({isNavExpanded: false })

                            }
          //setIsNavExpanded(!isNavExpanded);
        }}>
       
          <Icon iconName='GlobalNavButton' />
         </button>
          </div>
                   <div className={styles.mobClass + ' ' + "logoutmobile"}> <a href="/sites/EasyApproval/_layouts/signout.aspx"> <img src={logout} style={{width:"35px",height:"35px",marginRight:"10px"}}></img></a></div>
                    <br/>
                       </div>
                       <div className={` ${styles.divAtt}`}></div>
            </div>
           
        );
    }
    // private projectMenuItem(menuItem: any, itemType: ContextualMenuItemType): IContextualMenuItem {
    //     return ({
    //         key: menuItem.text,
    //         name: menuItem.text,
    //         href: menuItem.columns.length == 0 ?
    //             (menuItem["url"] != undefined ?
    //                 menuItem["url"]
    //                 : null)
    //             : null,
    //         subMenuProps: menuItem.columns.length > 0 ?
    //             {
    //                 items: menuItem.columns.map((i) => {
    //                     return (this.projectMenuHeading(i, ContextualMenuItemType.Normal));
    //                 })
    //             }
    //             : null
    //     });
    // }

    private projectMenuItem(menuItem: any, itemType: ContextualMenuItemType): IContextualMenuItem {
      return {
          key: menuItem.text,
          name: menuItem.text,
          href: menuItem.columns.length === 0 ? menuItem["url"] : undefined,
          subMenuProps: menuItem.columns.length > 0
              ? {
                  items: menuItem.columns.map((i:number) => {
                      return this.projectMenuHeading(i, ContextualMenuItemType.Normal);
                  })
              }
              : undefined // Assign undefined when there are no submenus instead of null
      };
  }
  

    // private projectMenuHeading(menuItem: any, itemType: ContextualMenuItemType): IContextualMenuItem {
    //     return ({
    //         key: menuItem.heading.text,
    //         name: menuItem.heading.text,
    //         href: menuItem.links.length == 0 ?
    //             (menuItem.heading.url != undefined ?
    //                 menuItem.heading.url
    //                 : null)
    //             : null,
    //         subMenuProps: menuItem.links.length > 0 ?
    //             { items: menuItem.links.map((i:number) => { return (this.projectMenuThirdLevel(i, ContextualMenuItemType.Normal)); }) }
    //             : null
    //     });
    // }
    private projectMenuHeading(menuItem: any, itemType: ContextualMenuItemType): IContextualMenuItem {
      return {
          key: menuItem.heading.text,
          name: menuItem.heading.text,
          href: menuItem.links.length === 0 ? 
              (menuItem.heading.url !== undefined ? menuItem.heading.url : undefined) 
              : undefined, // Return undefined instead of null when there are no links
          subMenuProps: menuItem.links.length > 0
              ? {
                  items: menuItem.links.map((i: number) => {
                      return this.projectMenuThirdLevel(i, ContextualMenuItemType.Normal);
                  })
              }
              : undefined // Return undefined instead of null when there are no sub-menu items
      };
  }
  

    private projectMenuThirdLevel(menuItem: any, itemType: ContextualMenuItemType): IContextualMenuItem {
        return ({
            key: menuItem.text,
            name: menuItem.text,
            href: menuItem["url"] != undefined ? menuItem["url"] : null
        });
    }
}
