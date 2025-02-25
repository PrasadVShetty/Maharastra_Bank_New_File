import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
//import { Attachment } from 'csstype';

 export interface EditState {
    selectedItems: any[];
    CommentsLog: any[];
    WFHistoryLog:any[];
    name: string; 
    description: string; 
    dpselectedItem?: { key: string | number | undefined };
    termKey?: string | number;
    dpselectedItems: IDropdownOption[];
    RecomselectedItems: IDropdownOption[];
    RecomNewselectedItems: IDropdownOption[];
    ReferselectedItems: IDropdownOption[];
    ControlselectedItems: IDropdownOption[];
    dropdownOptions:IDropdownOption[];
    disableToggle: boolean;
    defaultChecked: boolean;
    pplPickerType:string;
    userManagerIDs: number[];
    iframeDialog:boolean;
    hideDialog: boolean;
    status: string;
    statusno:number;
    isChecked: boolean;
    showPanel: boolean;
    required:string;
    onSubmission:boolean;
    termnCond:boolean;
    ManagerEmail:string[];
    seqno: string;
    attachments:any[];
    Note:any[];
    AppAttachments:any[];
    AttachType:string;
    Appstatus:string;
    MgrName:string;
    files:any;
    UserID:number;
    UserEmail:string;
    ImgUrl:string;
    ReturnedByID:number;
    ReturnedByName:string;
    ReferredByID:number;
    ReferredByName:string;
    ReqID:number;
    ReqName:string;
    pplTo:number;
    To:string;
    NoteType:String;
    Notefilename:string;
    Sitename:string;
    Absoluteurl:string;
    CurrApproverEmail:string;
    CurrAppID:number;
    AdminFlag:string;
    ccIDS:number[];
    ccName:string;
    ccEmail:string;
    ModifiedDate:Date | null;
    MarkIDs:number[];
    MarkName:string[];
    MarkEmails:string[];
    MarkItems:any[];   
    AllApprovers:any[];
    RecpID:number[];
    RecpName:string[];
    RecpEmail:string[];
    ReferredCasesCount:number;
    ReferredCasesLastCount:number;
    Charsleft:number;
    RestrictedEmails:string[];
   }