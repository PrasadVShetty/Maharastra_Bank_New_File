   import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
   //import { Attachment } from 'csstype';

   export interface CIState {
      selectedItems: any[];
      name: string; 
      description: string; 
      dpselectedItem?: { key: string | number | undefined };
      termKey?: string | number;
      dpselectedItems: IDropdownOption[];
      dropdownOptions:IDropdownOption[];
      disableToggle: boolean;
      defaultChecked: boolean;
      pplPickerType:string;
      userManagerIDs: number[];
      ccIDS:number[];
      ccName:string[];
      iframeDialog:boolean;
      hideDialog: boolean;
      status: string;
      isChecked: boolean;
      showPanel: boolean;
      required:string;
      onSubmission:boolean;
      termnCond:boolean;
      ManagerEmail:string[];
      ccEmail:string[];
      seqno: string;
      attachments:any[];
      Note:any[];
      AttachType:string;
      Appstatus:string;
      MgrName:string;
      files:any;
      UserID:number;
      UserEmail:string;
      ImgUrl:string;
      CurrentItemId:number;
      RecpEmail:string[];
      RecpID:number[];
      RecpName:string[];
      NoteType:String;
      Notefilename:string;
      Sitename:string;
      Absoluteurl:string;
      AppSeqNo:number;
      RecommSeqNo:number;
      ccSelectedItems: any[];
      InwarddocketnoSet:string;
      Outwarddocketno: any[];
      OutwarddocketnoSet:string;
      RadioClient:string;
      controllerPPId:number;
      controllerName:string;
      RestrictedEmails:string[];
      RestrictedEmailsMsg:string[];
      DepartmentItems:string[];
      FinNotes:string[];
      DOPItems:string[];
      selectedDate?: Date;
   }