import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';

export interface CIState {
    selectedItems: any[];
    name: string; 
    description: string; 
    pplPickerType: string;
    userManagerIDs: number[];
    status: string;
    showPanel: boolean;
    onSubmission: boolean;
    ManagerEmail: string[];
    seqno: string;
    attachments: File[]; // If you're expecting file objects
    Note: any[];
    AttachType: string;
    MgrName: string;
    files: any;
    UserID: number;
    UserEmail: string;
    ImgUrl: string;
    CurrentItemId: number;
    NoteType: string;
    Notefilename: string;
    Sitename: string;
    Absoluteurl: string;
    RadioClient: string;
    hideDialog: boolean;
    DepartmentItems: string[];
}
