import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
//import { Attachment } from 'csstype';

 export interface CIState {
     selectedItems: any[];
   name: string; 
   description: string; 
    pplPickerType:string;
   userManagerIDs: number[];
   status: string;
   showPanel: boolean;
    onSubmission:boolean;
     ManagerEmail:string[];
     seqno: string;
   attachments:any[];
   Note:any[];
   AttachType:string;
    MgrName:string;
   files:any;
   UserID:number;
   UserEmail:string;
   ImgUrl:string;
   CurrentItemId:number;
    NoteType:String;
   Notefilename:string;
   Sitename:string;
   Absoluteurl:string;
   RadioClient:string;
   hideDialog: boolean;
   DepartmentItems:string[];
   selectedUsers: { id: string, text: string }[];
   }