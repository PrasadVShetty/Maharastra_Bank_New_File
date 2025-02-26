import pnp from 'sp-pnp-js';
import { Web } from "sp-pnp-js/lib/sharepoint/webs";

// import { Web } from "@pnp/sp/webs";
// import { spfi } from "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/site-users/web";
// import "@pnp/sp/site-groups/web";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";

import { TopLevelMenu } from '../model/TopLevelMenu';
import { FlyoutColumn } from '../model/FlyoutColumn';
//import { Link } from '../model/Link';
//import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';

import { sampleData } from './MegaMenuSampleData';
import { SiteUserProps, SiteUser } from 'sp-pnp-js/lib/sharepoint/siteusers';
//import { SiteUser,CurrentUser } from 'sp-pnp-js/lib/sharepoint/siteusers';

export class MegaMenuService {

    private static readonly useSampleData: boolean = false;

    private static readonly level1ListName: string = "Mega Menu - Level 1";
    private static readonly level2ListName: string = "Mega Menu - Level 2";
  //  private static readonly level3ListName: string = "Mega Menu - Level 3";

    private static readonly cacheKey: string = "MegaMenuTopLevelItems";


    // Get items for the menu and cache the result in session cache.
    public static getMenuItems(siteCollectionUrl: string): Promise<TopLevelMenu[]> {
debugger;
        if (!MegaMenuService.useSampleData) {

            return new Promise<TopLevelMenu[]>((resolve, reject) => {

                // See if we've cached the result previously.
                var topLevelItems: TopLevelMenu[] = pnp.storage.session.get(MegaMenuService.cacheKey);

                if (topLevelItems) {
                    console.log("Found mega menu items in cache.");
                    resolve(topLevelItems);
                }
                else {
                    
                    this.getUserGroups().then((uid)=>{

                        let UIDlen=uid.length;
                        console.log("Didn't find mega menu items in cache, getting from list.");

                        var level1ItemsPromise = MegaMenuService.getMenuItemsFromSp(MegaMenuService.level1ListName, siteCollectionUrl,UIDlen);
                        var level2ItemsPromise = MegaMenuService.getMenuItemsFromSp(MegaMenuService.level2ListName, siteCollectionUrl,UIDlen);
                //        var level3ItemsPromise = MegaMenuService.getMenuItemsFromSp(MegaMenuService.level3ListName, siteCollectionUrl);
    
                       // Promise.all([level1ItemsPromise, level2ItemsPromise, level3ItemsPromise])
                       Promise.all([level1ItemsPromise, level2ItemsPromise])
                            .then((results: any[][]) => {
                                topLevelItems = MegaMenuService.convertItemsFromSp(results[0], results[1]);
                              //   topLevelItems = MegaMenuService.convertItemsFromSp(results[0], results[1], results[2]);
                                // Store in session cache.
                                pnp.storage.session.put(MegaMenuService.cacheKey, topLevelItems);
                                resolve(topLevelItems);
                            });

                    });

                   
                }
            });
        }
        else {
            return new Promise<TopLevelMenu[]>((resolve, reject) => {
                resolve(sampleData);
            });
        }

    }

    private static  getUserGroups(): Promise<any[]>{
    return new Promise<any[]>((resolve, reject) => {
            let UserIDs:Number[]=[] ;let uid:Number;
            //let web = new Web('https://isritechnologiessolutions.sharepoint.com/sites/EasyApproval');
            let web = new Web('https://bankofmaha.sharepoint.com/sites/BOMCompliance_UAT');
            web.currentUser.get().then((r:SiteUserProps) => {
            uid=r.Id;
            web.siteGroups.getByName("BOMCompliance_UAT Owners").users.get().then((u: any) => {
            u.forEach((user: SiteUserProps) =>{
            if(uid==user.Id){ UserIDs.push(user["Id"]);}
            resolve(UserIDs);
          });                
        });           
      });        
     });
    }

    

    // private static async getUserGroups(): Promise<number[]> {
    //     try {
    //         // Initialize SharePoint context
    //         //const sp = spfi("https://isritechnologiessolutions.sharepoint.com/sites/EasyApproval");
    //         const sp = spfi("https://isritechnologiessolutions.sharepoint.com/sites/EasyApproval");
    //         console.log(sp);
    //         // Get current user
    //         const currentUser = await sp.web.currentUser();
    //         const currentUserId = currentUser.Id;
    
    //         // Initialize an array to store user IDs
    //         let userIDs: number[] = [];
    
    //         // Get users from the "EasyApproval Owners" group
    //         const groupUsers = await sp.web.siteGroups.getById(3).users();
            
    //         // Filter the group users based on the current user's ID
    //         groupUsers.forEach(user => {
    //             if (user.Id === currentUserId) {
    //                 userIDs.push(user.Id);  // Add the current user's ID to the array
    //             }
    //         });
    
    //         return userIDs;  // Return an array of user IDs (could be empty if user is not found)
    //     } catch (error) {
    //         console.error("Error fetching user groups:", error);
    //         throw error;
    //     }
    // }    
    

    //Get raw results from SP.
    private static getMenuItemsFromSp(listName: string, siteCollectionUrl: string,UIDlen:number): Promise<any[]> {        
                return new Promise<TopLevelMenu[]>((resolve, reject) => {
                    let filterOption ="Admin eq 'No'";
                    // this.getUserGroups().then((uid)=>{
                    if(UIDlen>0){filterOption="Admin eq 'Yes' or Admin eq 'No'";}

                    let web = new Web(siteCollectionUrl);

                    // TODO : Note that passing in url and using this approach is a workaround. I would have liked to just
                    // call pnp.sp.site.rootWeb.lists, however when running this code on SPO modern pages, the REST call ended
                    // up with a corrupt URL. However it was OK on View All Site content pages, etc.
                    // Added by Amar - .filter("substringof('"+filterOption+"',Admin)")
                    web.lists
                    .getByTitle(listName)
                    .items
                    .orderBy("SortOrder")
                    .filter(filterOption)
                    .get()
                    .then((items: any[]) => {
                    console.log(items.length);
                    resolve(items);
                })
                .catch((error: any) => {
                reject(error);
            });
        });
    }

    // private static async getMenuItemsFromSp(listName: string, siteCollectionUrl: string, UIDlen: number): Promise<any[]> {        
    //     try {
    //         let filterOption = "Admin eq 'No'";
    //         if (UIDlen > 0) {
    //             filterOption = "Admin eq 'Yes' or Admin eq 'No'";
    //         }
    
    //         // ✅ Correct way to create a Web instance
    //         const web = new Web(siteCollectionUrl);
    
    //         // ✅ Correct way to retrieve list items
    //         const items = await web.lists
    //             .getByTitle(listName)
    //             .items
    //             .select("Title", "SortOrder", "Admin") // Select only necessary fields
    //             .orderBy("SortOrder", true) // Sorting (true = ascending)
    //             .filter(filterOption)
    //             .top(5000) // Limit results if needed (SharePoint threshold)
    //             ();
    
    //         console.log(items.length);
    //         return items;
    //     } catch (error) {
    //         console.error("Error fetching menu items:", error);
    //         throw error;
    //     }
    // }


    // Convert results from SP into actual entities with correct relationships.
    private static convertItemsFromSp(level1: any[], level2: any[]): TopLevelMenu[] {

        var level1Dictionary: { [id: number]: TopLevelMenu; } = {};
        var level2Dictionary: { [id: number]: FlyoutColumn; } = {};

        // Convert level 1 items and store in dictionary.
        var level1Items: TopLevelMenu[] = level1.map((item: any) => {
            var newItem = {
                key: item.ID,
                id: item.Id,
                text: item.Title,
                columns: []
            };

            level1Dictionary[newItem.id] = newItem;

            return newItem;
        });

        // Convert level 2 items and store in dictionary.
        var level2Items: FlyoutColumn[] = level2.map((item: any) => {
            var newItem = {
                id: item.Id,
                heading: {
                    key: item.ID,
                    text: item.Title,
                    url: item.Url ? item.Url.Url : "",
                    openInNewTab: item.OpenInNewTab
                },
                links: [],
                level1ParentId: item.Level1ItemId
            };

            level2Dictionary[newItem.id] = newItem;

            return newItem;
        });

     

        // Now link the entities into the desired structure.
        

        for (let l2 of level2Items) {
            const level1Parent = level1Dictionary[l2.level1ParentId];
    
            if (level1Parent) {
                if (level1Parent.columns) {
                    // If 'columns' is defined, push to it
                    level1Parent.columns.push(l2);
                } else {
                    // If 'columns' is undefined, initialize it first
                    level1Parent.columns = [l2];
                }
            }
        }

        var retVal: TopLevelMenu[] = [];

        for (let l1 of level1Items) {
            retVal.push(l1);
        }

        return retVal;

    }
}