import * as React from 'react';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import styles from './PaperlessApproval.module.scss';
import { IPaperlessApprovalProps } from './IPaperlessApprovalProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
//import { PeoplePicker } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext, PeoplePicker,PrincipalType} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { CIState } from "../Model/CIState";
import { default as pnp, ItemAddResult, File, sp, Web } from "sp-pnp-js";
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/components/Button';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
//import { CurrentUser } from '@pnp/sp/src/siteusers';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Button } from 'office-ui-fabric-react/lib/Button';
import { Attachments } from 'sp-pnp-js/lib/graph/attachments';
import * as jQuery from 'jquery';
import * as $ from "jquery";
import { SiteUser } from 'sp-pnp-js/lib/sharepoint/siteusers';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { ListItemPicker } from '@pnp/spfx-controls-react/lib/listItemPicker';

//require('../css/custom.css');
//require('/sites/EasyApprovalUATNew/SiteAssets/css/styles.css');
SPComponentLoader.loadCss('../SiteAssets/css/styles.css'); 
//SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css');
// SPComponentLoader.loadCss('https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css');

const Delete: any = require('../images/Delete.png');
const Video: any = require('../images/Video.png');
const Logo: any = require('../images/Logo.png');
const Annex: any = require('../images/Upload-Annex.png');
const NoteAtt: any = require('../images/Upload-Note.png');

export default class PNoteForms extends React.Component<IPaperlessApprovalProps, CIState> {
  constructor(props : any) {
    super(props);
    SPComponentLoader.loadCss('https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css');

    this.handleTitle = this.handleTitle.bind(this);
    this.handleDesc = this.handleDesc.bind(this);
    this._onRenderFooterContent = this._onRenderFooterContent.bind(this);
    this.createItem = this.createItem.bind(this);
    this._getManager = this._getManager.bind(this);
    this._getCCPeople = this._getCCPeople.bind(this);
    this._getReceivedFrom = this._getReceivedFrom.bind(this);
    this.DeleteApprover = this.DeleteApprover.bind(this);

    //this.handleKeyUp = this.handleKeyUp.bind(this);

    //  this.setButtonsEventHandlers();
    this.state = {
      name: "",
      description: "",
      selectedItems: [],
      hideDialog: true,
      showPanel: false,
      dpselectedItem: undefined,
      dpselectedItems: [],
      dropdownOptions: [],
      disableToggle: false,
      defaultChecked: false,
      termKey: undefined,
      userManagerIDs: [],
      pplPickerType: "",
      status: "",
      isChecked: false,
      required: "This is required",
      onSubmission: false,
      termnCond: false,
      ManagerEmail: [],
      seqno: "",
      attachments: [],
      Note: [],
      AttachType: '',
      Appstatus: '',
      MgrName: '',
      files: null,
      UserID: 0,
      UserEmail: '',
      iframeDialog: true,
      ImgUrl: '',
      CurrentItemId: 0,
      RecpEmail: [],
      RecpID: [],
      RecpName: [],
      NoteType: '',
      Notefilename: '',
      Sitename: '',
      Absoluteurl: '',
      ccEmail: [],
      ccIDS: [],
      ccName: [],
      AppSeqNo: 0,
      RecommSeqNo: 0,
      ccSelectedItems: [],
      InwarddocketnoSet: '',
      Outwarddocketno: [],
      OutwarddocketnoSet: '',
      RadioClient: '',
      controllerName: '',
      controllerPPId: 0,
      RestrictedEmails: [],
      RestrictedEmailsMsg: [],
      DepartmentItems:[],
      FinNotes:[],
      DOPItems:[],
      selectedDate: undefined,
      checklist: '',
      status2: '',
      items: [],
      savedData: [],
      vettingobservation : '',
      Checklistlabel: '',
      Title:'' 
    };
  }
    
  addItem = () => {
    const { checklist, status, items } = this.state;
  
    // Trim whitespace and validate fields
    if (!checklist.trim()) {
      alert("Please enter a checklist item.");
      return;
    }
    if (!status) {
      alert("Please select a status.");
      return;
    }
  
    // Add new item to the list
    const newItems = [...items, { 
      id: items.length + 1, 
      checklist: checklist.trim(), 
      status: status,
      itemid : 0 
    }];
  
    // Update state and reset fields
    this.setState({ items: newItems, checklist: '', status: '' });
  };  
  

  deleteItem = (id: number) => {
    const filteredItems = this.state.items.filter(item => item.id !== id);
    this.setState({ items: filteredItems });
  };

  
  // saveTableData = () => {
  //   this.setState({ savedData: [...this.state.items] });
  //   console.log('Saved Data:', this.state.items);
  // }
  
  public render(): React.ReactElement<IPaperlessApprovalProps> {
    //debugger;
    const { dpselectedItem, dpselectedItems } = this.state;
    const { name, description } = this.state;
    pnp.setup({
      spfxContext: this.props.context
    });
    const peoplePickerContext: IPeoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient
    };        
    const statusOptions: IDropdownOption[] = [
      { key: 'Checked', text: 'Checked' },
      { key: 'NotChecked', text: 'NotChecked' },
    ];

    return (
      <form >
        <div className={styles.paperlessApproval}>
          <div className={styles.container}>            
            <div className={styles.formrow}>
              <div id="divHeadingNew" style={{textAlign:"center",backgroundColor:"#0c78b8",color:"white",display:"block", fontSize:'18px'}}>
                <h3 className={styles.heading}>Note Form </h3>
              </div>
               <div hidden id="divHeadingSubmit" className="ms-Grid-col ms-u-sm10 block" style={{ display: "none" }}>
                <h3 className={styles.heading} style={{ fontSize: "18px", textAlign: "center", color: "white", top: "5px" }}>Note Form</h3>
              </div>
            </div>


            <div className={styles.panel}>                        
            <div className='row pt-2 pb-1 m-0' style={{width:"100%",backgroundColor:"#50B4E6", color:'#fff', justifyItems:'center'}}>
               <div className='col-md-1 col-lg-2 col-sm-4'>
                  <label className='control-form-label'><b>Requester</b></label>
               </div>
               <div className='col-md-2 col-lg-2 col-sm-8' id="tdName" style={{borderRight:'1px solid #fff'}}>                
               </div>

               <div className='col-md-1 col-lg-2 col-sm-4' >
                  <label className='control-form-label'><b>Status</b></label>
               </div>
               <div className='col-md-2 col-lg-2 col-sm-8' style={{borderRight:'1px solid #fff'}}> 
               New               
               </div>
               <div className='col-md-2 col-lg-2 col-sm-4'>
                  <label className='control-form-label'><b>Creation Date</b></label>
               </div>
               <div className='col-md-2 col-lg-2 col-sm-8' id="tdDate">                
               </div>
               <div className='col-md-8 col-lg-8 col-sm-12' style={{display:"none"}}  id="tdFY">

               </div>
            </div>

            <hr/>             
            <div className={styles.formrow + " " + "form-group row"}>
            <div  className='col-md-3'>
              <label className={styles.lbl + " " + styles.Reqdlabel}>Note For</label>
              </div>
              <div className='col-md-9'>
                <input  type="text" title="Enter Note For" placeholder="Enter Note For" id="txtNote"  className='form-control form-control-sm'/>                
              </div>             
            </div>

            <div className={styles.formrow + " " + "form-group row"}>
            <div  className='col-md-3'>
              <label className={styles.lbl + " " + styles.Reqdlabel}>Subject</label>
              </div>
              <div className='col-md-9'>
                <input  type="text" title="Enter Subject" placeholder="Enter Subject" id="txtSubject"  className='form-control form-control-sm'/>                
              </div>             
            </div>

            <div className={styles.formrow + " " + "form-group row"}>
            <div  className='col-md-3'>
              <label className={styles.lbl + " " + styles.Reqdlabel}>Purpose</label>
              </div>
              <div className='col-md-9'>
                <input  type="text" title="Enter Purpose" placeholder="Enter Purpose" id="txtPurpose"  className='form-control form-control-sm'/>                
              </div>             
            </div>
            
            <div className={styles.formrow + " " + "form-group row"}>
            <div  className='col-md-3'>
              <label className={styles.lbl}>Product Name</label>
              </div>
              <div className='col-md-9'>
                <input  type="text" title="Enter Product Name" placeholder="Enter Product Name" id="txtReturn"  className='form-control form-control-sm'/>                
              </div>             
            </div>
            
            <div className={styles.formrow + " " + "form-group row"}>
            <div  className='col-md-3'>
            <label className={styles.lbl + " " + styles.Reqdlabel}>Department Ownership</label>
            </div>
              <div className='col-md-9'>
                <select className='form-control form-control-sm' id="ddlDeptOwnership" title="Select Department" placeholder="Select Department">
                  <option>Select</option>
                </select>
              </div>
              <br />
            </div>              

            <div className={styles.formrow + " " + "form-group row"}>
            <div  className='col-md-3'>
              <label className={styles.lbl}>Referred Guidelines</label>
              </div>
              <div className='col-md-9'>
                <input  type="text" title="Enter Referred Guidelines" placeholder="Enter Referred Guidelines" id="txtGuidelines"  className='form-control form-control-sm'/>                
               </div>             
            </div> 

            <div className={styles.formrow + " " + "form-group row"}>
            <div className="col-md-3">
            <label className={styles.lbl}>{this.state.vettingobservation}</label>
            </div>
            <div className="col-md-9">
            <textarea 
            title={`Enter ${this.state.vettingobservation}`}
            placeholder={`Enter ${this.state.vettingobservation}`}
            id="txtVetting" 
            className="form-control form-control-sm"
            rows={4}
            style={{}}             
            />
            </div>
            </div>


            <div>
            {/* Checklist Input Field with given styles */}
            <div className={styles.formrow + " " + "form-group row"}>
            <div className="col-md-3">
            <label className={styles.lbl}>{this.state.Checklistlabel}</label>
            </div>
            <div className="col-md-9">
            <input
            type="text"
            title={`Enter ${this.state.Checklistlabel}`}
            placeholder={`Enter ${this.state.Checklistlabel}`}
            value={this.state.checklist}
            onChange={(e) => this.setState({ checklist: e.target.value || '' })}
            className="form-control form-control-sm"
            />
            </div>
            </div>

            
            <div className={styles.formrow + " " + "form-group row"}>
            <div className="col-md-3">
            <label className={styles.lbl}>Status</label>
            </div>
            <div className="col-md-9">
            <Dropdown
            selectedKey={this.state.status}
            onChange={(e, option) => this.setState({ status: option?.key as string })}
            options={statusOptions}
            styles={{
            root: { width: '100%' },
            dropdown: { backgroundColor: '#fff' },
            title: { borderRadius: '5px' },
            }}
            />
            </div>
            </div>

            {/* Add Button */}
            <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: '10px' , paddingRight: '35px'}}>
            <PrimaryButton iconProps={{ iconName: 'Add' }} onClick={this.addItem} />
            </div>
            <br />            
            
            {/* Checklist Table */}
            {this.state.items.length > 0 && (
            <div style={{ overflowX: 'auto', width: '100%' }}>
            <table className="table table-bordered" style={{ width: "100%", textAlign: "center" }}>
            <thead style={{ backgroundColor: "#f4f4f4", fontWeight: "bold" }}>
            <tr>
            <th style={{ width: "10%", padding: "8px" }}>Sr No.</th>
            <th style={{ width: "55%", padding: "8px", wordWrap: "break-word" }}>{this.state.Checklistlabel}</th>
            <th style={{ width: "20%", padding: "8px" }}>Status</th>
            <th style={{ width: "15%", padding: "8px" }}>Action</th>
            </tr>
            </thead>
            <tbody>
            {this.state.items.map((item, index) => (
            <tr key={item.id}>
            <td style={{ padding: "8px" }}>{index + 1}</td>
            <td style={{ padding: "8px", wordWrap: "break-word" }}>{item.checklist}</td>
            <td style={{ padding: "8px" }}>{item.status}</td>
            <td style={{ padding: "8px" }}>
            <PrimaryButton
            iconProps={{ iconName: 'Delete' }}
            text="Remove"
            onClick={() => this.deleteItem(item.id)}
            styles={{
            root: { backgroundColor: "#d9534f", color: "#fff", borderRadius: "5px" },
            rootHovered: { backgroundColor: "#c9302c" },
            }}
            />
            </td>
            </tr>
            ))}
            </tbody>
            </table>
            </div>
            )}
            </div>            
            

            {/* <div className={styles.formrow + " " + "form-group row"}>
            <div className='col-md-3'>
            <label className={styles.lbl + " " + styles.Reqdlabel}>Due Date</label>
            </div>
            <div className='col-md-3'>
            <DatePicker
            placeholder="Select a date..."
            value={this.state.selectedDate}
            onSelectDate={(date) => this.setState({ selectedDate: date || undefined })}
            isMonthPickerVisible={false}  
            formatDate={(date) => date ? `${date.getDate()}/${date.getMonth() + 1}/${date.getFullYear()}` : ''} 
            showGoToToday={false}  
            calloutProps={{
            doNotLayer: true,  
            styles: { root: { maxWidth: '250px', padding: '5px' } } 
            }}
            styles={{
            root: { width: '200px' },  
            textField: { width: '200px', fontSize: '14px' }  
            }}
            />
            </div>
            </div> 

            <div className={styles.formrow + " " + "form-group row"}>
            <div  className='col-md-3'>
              <label className={styles.lbl + " " + styles.Reqdlabel}>Place</label>
              </div>
              <div className='col-md-9'>
                <input  type="text" title="Enter Place" placeholder="Enter Place" id="txtPlace"  className='form-control form-control-sm'/>                
              </div>             
            </div>*/}

            <div className={styles.formrow + " " + "form-group row"}>
            <div  className='col-md-3'>
              <label className={styles.lbl + " " + styles.Reqdlabel}>Note Type</label>
              </div>
              <div className='col-md-9'>
                <select  className='form-control form-control-sm' id="ddlSource" title="Select Note Type" placeholder="Select Note Type" onChange={() => this.SelectSource()}>
                  <option>Select</option>
                  <option>Financial</option>
                  <option>Non-Financial</option>
                </select>
              </div>
              <br />
            </div>
           
              <div className='FinancialClass' style={{ display: "none" }}>
              <div className={styles.formrow + " " + "form-group row"}>
              <div  className='col-md-3'>
              <label className={styles.lbl + " " + styles.Reqdlabel}>Type of Financial Note</label>
              </div>
              <div className='col-md-9'>
                <select  className='form-control form-control-sm' id="ddlFinNote" placeholder="Select Financial Note" title="Select Financial Note">
                  <option>Select</option>
                </select>
              </div>  
              </div>
                     
            </div>
            <div className='FinancialClass' style={{ display: "none" }}>
            <div className={styles.formrow + " " + "form-group row "}>
            <div className='col-md-3'>
              <label className={styles.lbl + " " + styles.Reqdlabel}>Amount</label>
              </div>
              <div className='col-md-9'>
                <input type="number" id="Amount"  className='form-control form-control-sm'></input>
              </div>   
              </div>           
            </div>
            {/* <div className='FinancialClass' style={{ display: "none" }}>
            <div className={styles.formrow + " " + "form-group row"} >
            <div className='col-md-3'>
              <label className={styles.lbl + " " + styles.Reqdlabel}>DOP Details</label>
             </div>
              <div className='col-md-9'>
                <select  className='form-control form-control-sm' id="ddlDOP" placeholder="Select Delegation of Power" title="Select DOP">
                  <option>Select</option>
                </select>
              </div>
              </div>             
            </div> */}

             {/*<div className={styles.formrow + " " + "form-group row"} id="divClient" style={{ display: "" }}>
            <div className='col-md-3 pr-0'>
              <label className={styles.lbl + " " + styles.Reqdlabel}>Do you want to add Client?</label>
                   </div>
              <div className='col-md-9'>
                <label className="custom-radio">
                  <input id="CYes" name="radioAttach" onChange={() => this.Radibtnchangeevent("radioAttach", "CYes")} value="CYes" type="radio" />
                  <span className="custom-control-indicator" style={{ padding: "2px" }}></span>
                  <span className={"custom-control-description"}>Yes</span>
                </label>
                <label className="custom-radio" style={{ padding: "8px" }}>
                  <input id="CNo" name="radioAttach" onChange={() => this.Radibtnchangeevent("radioAttach", "CNo")} value="CNo" type="radio" />
                  <span className="custom-control-indicator" style={{ padding: "2px" }}></span>
                  <span className={"custom-control-description"}>No</span>
                </label>
              </div>
             
            </div>
            <div id="divClientName" style={{ display: "none" }}>
            <div className={styles.formrow + " " + "form-group row"} >
            <div  className='col-md-3'>
              <label className={styles.lbl + " " + styles.Reqdlabel}>Client Name/Vendor Name</label>
             </div>
              <div  className='col-md-9'>
                <input type="text" title="Enter Client/Vendor Name" placeholder="Enter Client Name" id="txtClient"  className='form-control form-control-sm' />
              </div>              
            </div>
            </div>
            <div className={styles.formrow + " " + "form-group row"}>
           <div className='col-md-3'>
              <label className={styles.lbl + " " + styles.Reqdlabel}>Is an Exception / Deviation ?</label>
              </div>
              <div className='col-md-9'>
                <label className="custom-radio">
                  <input id="ExcYes" name="radioExc" onChange={() => this.Radibtnchangeevent("radioExc", "ExcYes")} value="ConfYes" type="radio" />
                  <span className="custom-control-indicator" style={{ padding: "2px" }}></span>
                  <span className={"custom-control-description"}>Yes</span>
                </label>
                <label className="custom-radio" style={{ padding: "8px" }}>
                  <input id="ExcNo" name="radioExc" onChange={() => this.Radibtnchangeevent("radioExc", "ExcNo")} value="ConfNo" type="radio" />
                  <span className="custom-control-indicator" style={{ padding: "2px" }}></span>
                  <span className={"custom-control-description"}>No</span>
                </label>
              </div>
              <br></br>
              <input type="text" id="txtExceptional" style={{ display: "none" }}  className='form-control form-control-sm'></input>
            </div>
              
              <div  id="divConfidential" style={{ display: "none" }}>
            <div className={styles.formrow + " " + "form-group row"}>
            <div className='col-md-3'>
              <label className={styles.lbl + " " + styles.Reqdlabel}>Is it a Confidential Note?</label>
              </div>
              <div className='col-md-9'>
                <label className="custom-radio">
                  <input id="ConfYes" name="radioConf" onChange={() => this.Radibtnchangeevent("radioConf", "ConfYes")} value="ConfYes" type="radio" />
                  <span className="custom-control-indicator" style={{ padding: "2px" }}></span>
                  <span className={"custom-control-description"}>Yes</span>
                </label>
                <label className="custom-radio" style={{ padding: "8px" }}>
                  <input id="ConfNo" name="radioConf" onChange={() => this.Radibtnchangeevent("radioConf", "ConfNo")} value="ConfNo" type="radio" />
                  <span className="custom-control-indicator" style={{ padding: "2px" }}></span>
                  <span className={"custom-control-description"}>No</span>
                </label>
              </div>
              <br></br>
              <input type="text" id="txtConfidential" style={{ display: "none" }}  className='form-control form-control-sm'></input>
            </div>
              </div> */}

            <div className={styles.formrow + " " + "form-group row"} style={{ display: "none" }}>
              <div className={styles.lbl}>Comments</div>
              <div className={styles.Vcolumn}>
                <textarea id="txtComments"   className='form-control form-control-sm'/>
              </div>
              <br />
            </div>
            <div className={styles.container} style={{padding:'0px 8px'}}>
              <div className={styles.formrow}>
              <h3 className={"text-left"} style={{ backgroundColor: "#50B4E6", fontSize: "16px", padding:'5px 10px', color:'#fff', width:'100%' }}>Recommender Details
                    <span style={{ position: "relative", marginLeft: "10px", color: "Red", fontSize: "14px", fontStyle: "italic" }}>*Note: Max.10 Recommenders can be added.</span>
                  </h3>

                  <table className={styles.tbl} id="tblMain" style={{ width: "100%" }}>
                    <tr>
                      <td style={{ width: "15%", paddingLeft:'10px', fontWeight:700}}>Recommender</td>
                      <td style={{ width: "70%" }} id="RecommenderPPtd">                        
                        <PeoplePicker
                        context={peoplePickerContext}
                        //titleText="People Picker"
                        personSelectionLimit={100}
                        groupName={""} 
                        showtooltip={true}
                        required={true}
                        disabled={false}
                        searchTextLimit={5}
                        ensureUser={true}
                        onChange={this._getReceivedFrom}
                        showHiddenInUI={false}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        defaultSelectedUsers= {this.state.RecpEmail}
                        errorMessageClassName={styles.hideElementManager}
                        />
                      </td>
                      <td style={{ width: "10%" }}><PrimaryButton style={{ width: "80pt", borderRadius: "5%", backgroundColor: "#f00", display: "none" }} text="Add Recommender" onClick={() => { this.AddRecommender(); }} /></td>
                    </tr>
                    {this.state.dpselectedItems ? this.state.dpselectedItems.map((data) => {
                      return data;
                    }) : null}


                  </table>
              </div>
            
              <hr/>
              <div className={styles.formrow}>
              <h3 className={styles.Reqdlabel + " " + "text-left"} style={{ backgroundColor: "#50B4E6", fontSize: "16px", color:'#fff', padding:'5px 10px' }}>Approver Details
                    <span style={{ position: "relative", marginLeft: "10px", color: "Red", fontSize: "14px", fontStyle: "italic" }}>*Note: Max.10 Approvers can be added.</span>
                  </h3>

                  <div className={styles.lbl}>
                  <table className={styles.tbl} id="tblMain1" style={{ width: "100%" }}>
                    <tr>
                      <td style={{ width: "15%", paddingLeft:'10px', fontWeight:700}}>Approver</td>
                      <td style={{ width: "70%" }} id="ApproverPPtd">                        
                        <PeoplePicker
                        context={peoplePickerContext}
                        //titleText="People Picker"
                        personSelectionLimit={100}
                        groupName={""} 
                        showtooltip={true}
                        required={true}
                        disabled={false}
                        searchTextLimit={5}
                        onChange={this._getManager}
                        showHiddenInUI={false}
                        ensureUser={true}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        defaultSelectedUsers= {this.state.ManagerEmail}
                        errorMessageClassName={styles.hideElementManager}
                        />
                      </td>
                      <td style={{ width: "10%" }}><PrimaryButton style={{ width: "80pt", borderRadius: "5%", backgroundColor: "#50B4E6", display: "none" }} id="btnAddApprover" text="Add Approver" onClick={() => { this.AddApprover(); }} /></td>
                    </tr>
                    {this.state.selectedItems ? this.state.selectedItems.map((data) => {
                      return data;
                    }) : null}


                  </table>
                </div>
              </div>
            
              <hr/>
              <div className={styles.formrow + " " + "form-group FinancialClass"} style={{ display: "none" }}>
              <h3 className={styles.Reqdlabel + " " + "text-left"} style={{ backgroundColor: "#50B4E6", fontSize: "16px" , color:'#fff', padding:'5px 10px', width:'100%'}}>Controller Details
                    <span style={{ position: "relative", marginLeft: "10px", color: "Red", fontSize: "14px", fontStyle: "italic" }}>*Note: Only 1 Controller can be added.</span>
                  </h3>
              </div>
              <div className={styles.formrow + " " + "form-group FinancialClass"} style={{ display: "none" }}>
                <div>
                  <table className={styles.tbl} id="tblMain1" style={{ width: "100%" }}>
                    <tr>
                      <td style={{ width: "15%", paddingLeft:'10px', fontWeight:700 }}>Controller</td>
                      <td style={{ width: "70%" }} id="ControllerPPtd">                        
                        <PeoplePicker
                        context={peoplePickerContext}
                        //titleText="People Picker"
                        personSelectionLimit={1}
                        groupName={""} 
                        showtooltip={true}
                        required={false}
                        disabled={false}
                        placeholder={"Person Name or Email address"}
                        searchTextLimit={5}
                        onChange={this._getCCPeople}
                        showHiddenInUI={false}
                        ensureUser={true}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        defaultSelectedUsers= {this.state.ccEmail}
                        errorMessageClassName={styles.hideElementManager}
                        />
                      </td>
                      <td style={{ width: "10%" }}><label style={{ display: "none" }} id="lblController"></label>
                      <PrimaryButton style={{ width: "80pt", borderRadius: "5%", backgroundColor: "#50B4E6", display: "none" }} 
                      text="Add Controller" id="AddControllerBtn" onClick={() => { this.AddController(); }} />
                      </td>
                    </tr>
                    {this.state.selectedItems ? this.state.ccSelectedItems.map((data) => {
                      return data;
                    }) : null}


                  </table>
                  <br />
                </div>
              </div>

              <div className={styles.formrow + " " + "form-group"}>
                
                  <h3 className="text-left" style={{ backgroundColor: "#50b4e6", color:'#fff', fontSize: "16px", padding:'5px 10px', width:'100%' }}>Attachments</h3>
               
              </div>

              <div className={styles.formrow + " " + "form-group row"} id="divNote" style={{ display: "" }}>
                <div className={styles.lbl + " " + styles.Tcolumn}>Note</div>
                <div className={styles.Vcolumn} >
                  <a id="NoteFile" href=""></a><span id="NoteDel" style={{ display: "none" }}><img src={Delete} style={{ width: "10pt", height: "10pt", position: "absolute" }} onClick={() => this.DeleteNote(this.state.OutwarddocketnoSet)}></img></span>
                </div>
              </div>

              <div className={styles.formrow + " " + "form-group row"} id="divAttach" style={{ display: "" }}>

                <div className={styles.lbl + " " + styles.Reqdlabel + " " + styles.Tcolumn}>
                  <a href="#"><img src={NoteAtt} style={{ height: "16pt", marginLeft: "10px" }} onClick={() => { this.UploadAttach('Note'); }}></img></a>
                  <br></br>
                  <label>(Note: Only pdf file can be attached)</label>
                </div>
                <div className={styles.Vcolumn}>
                  {this.state.Note.map((vals) => {
                    let filename = vals.split("/")[1];
                    return (<span style={{ position: "relative", padding: "5px" }}><a href={this.state.Absoluteurl + "/ChecklistAttach/" + vals}>{filename}</a><img src={Delete} style={{ width: "10pt", height: "10pt", position: "absolute" }} onClick={() => this.DeleteNote(vals)}></img> </span>);

                  })}

                </div>
                <div hidden className="ms-Grid-col ms-u-sm12 block hide" id="divAttachButton" style={{ backgroundColor: "white", display: "none" }}>
                  <input type='file' style={{}} id='fileUploadInput' required={true} name='myfile' multiple onChange={this.AttachLib} />
                </div>
              </div>
              <br />
              <div className={styles.formrow + " " + "form-group row"} style={{ margin: "0px" }}>

                <div className={styles.lbl + " " + styles.Tcolumn}>
                  <a href="#"><img src={Annex} style={{ height: "16pt" }} onClick={() => { this.UploadAttach('Annexures'); }}></img></a>
                  <br></br>
                  <small style={{color:'#f00'}}>(image,.pdf,.doc,.docx,.xlsx,.eml)</small>
                  <br></br>
                  <label>*Max 20 Annexures</label>
                </div>
                <div className={styles.Vcolumn}>
                  {this.state.attachments.map((vals) => {
                    let filename = vals.split("/")[1];
                    return (<span style={{ position: "relative", padding: "5px" }}><a href={this.state.Absoluteurl + "/ChecklistAnnexures/" + vals}>{filename}</a><img src={Delete} style={{ width: "10pt", height: "10pt", position: "absolute" }} onClick={() => this.DeleteAttachment(vals)}></img> </span>);

                  })}

                </div>

              </div>
            </div>
          </div>

          <div className={styles.container} style={{ marginTop: "5px" }}>

            <div className={styles.overlay} id="overlay" style={{ display: "none" }} >
              <span className={styles.overlayContent} style={{ textAlign: "center" }}>Please Wait!!!</span>
            </div>
            <br></br>
            <div className={styles.formrow + " " + "form-group row"} style={{margin: "0px", paddingLeft:'10px' }}>             
              <div id="btnCreate" style={{ display: "block", marginLeft:'10px' }} >
                <PrimaryButton className='btn' style={{ width: "25pt", borderRadius: "5%", backgroundColor: "#0c78b8", color:'#fff'  }} text="Submit" onClick={() => { this.validateForm(); }} /> 
                </div>

              <div id="btnDraft" style={{borderRadius: "5px", marginLeft:'10px' }} >
                <PrimaryButton className='btn' style={{ width: "25pt", borderRadius: "5%", backgroundColor: "#0c78b8", color:'#fff'  }} text="Save Draft" onClick={() => { this.SaveDraftNew(); }} />
              </div>

              <div id="btnCancel" style={{ display: "block", marginLeft:'10px' }}>
                <DefaultButton className='btn' style={{ width: "25pt", borderRadius: "5%", backgroundColor: "#f00", color:'#fff' }} text="Cancel" onClick={() => { this.cancel(); }} />
              </div>
              <div id="btnClose" style={{ display: "none", width: "25pt", borderRadius: "50%" , marginLeft:'10px'}}>
                <DefaultButton className='btn' style={{ width: "25pt", borderRadius: "5%", backgroundColor: "#f00" , color:'#fff' }} text="Close" onClick={() => { this.cancel(); }} />
              </div>
             
            </div>

            
            <div>

              <Panel
                isOpen={this.state.showPanel}
                type={PanelType.smallFixedFar}
                onDismiss={this._onClosePanel}
                isFooterAtBottom={false}
                headerText="Are you sure you want to submit this request?"
                closeButtonAriaLabel="Close"
                onRenderFooterContent={this._onRenderFooterContent}
              ><span>Please check the details filled along with attachment and click on Confirm button to submit request.</span>
              </Panel>
            </div>


            <Dialog
              hidden={this.state.hideDialog}
              onDismiss={this._closeDialog}
              dialogContentProps={{
                type: DialogType.largeHeader,
                title: 'Request has been Submitted Successfully',
                subText: ""
              }}
              modalProps={{
                titleAriaId: 'myLabelId',
                subtitleAriaId: 'mySubTextId',
                isBlocking: false,
                containerClassName: 'ms-dialogMainOverride'
              }}>
              <div dangerouslySetInnerHTML={{ __html: this.state.status }} />
              <DialogFooter>
                <PrimaryButton onClick={() => this.gotoHomePage()} text="Okay" />
              </DialogFooter>
            </Dialog>

          </div>
          </div>
        </div>
      </form>
    );
  }
  /* -- Starting All Functions-- */

  /*-- For Upload Attachment Popup--*/
  public UploadAttach(AttType: string) {
    //debugger;
    this.setState({ AttachType: AttType });
    setTimeout(() => {
      // document.getElementById('fileUploadInput').click();
      let overlay = document.getElementById('fileUploadInput');
      if (overlay) {    
        overlay.click();
      }
    }, 1500);

  }
  /*--End--*/

  /*-- For Updating Attachments State Change--*/
  public handleChange(files : any) {
    this.setState({
      files: files
    });
  }
  /*-- End Function--*/

  /*--For on(show) and off(hide) please wait overlay while page load--*/
  private on() {
    let ht = window.innerHeight;
    // document.getElementById('overlay').style.height = ht.toString() + "px";
    // document.getElementById("overlay").style.display = "block";
    let overlay = document.getElementById('overlay');
    if (overlay) {    
      overlay.style.height = ht.toString() + "px";
    }
    let overlay2 = document.getElementById('overlay');
    if (overlay2) {    
      overlay2.style.display = "block";
    }
  }
  private off() {
    // document.getElementById("overlay").style.display = "none";
    let overlay = document.getElementById('overlay');
    if (overlay) {    
      overlay.style.display = "none";
    }
  }
  /*--End--*/

  /*-- On Load Function--*/
  public componentDidMount() {
    //debugger;
    var reacthandler = this;


    //pnp.sp.web.currentUser.get().then((r: CurrentUser) => {  //To get current user details from site  
    pnp.sp.web.currentUser.get().then((r) => {
      //  console.log(r);
      let sitename = r['odata.id'].split("/_api")[0];
      let absoluteurl = sitename.split("com")[1] + "/Main";
      this.setState({ Absoluteurl: absoluteurl });
      this.setState({ Sitename: sitename });
      const uname = r['UserPrincipalName'].split('@')[0];
      let username = r['Title'];
      // document.getElementById("tdName").innerText = username;
      let overlay = document.getElementById('tdName');
      if (overlay) {    
        overlay.innerText = username;
      }
      this.setState({ name: username });
      this.setState({ UserID: r['Id'] });
      let CurrUserEmail = r['LoginName'].split("|")[2];
      this.setState({ UserEmail: CurrUserEmail });
      /*-- To generate random string for sequence number--*/
      const text = new Array();
      const possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
      for (let i = 0; i < 6; i++) {
        text[i] = possible.charAt(Math.floor(Math.random() * possible.length));
      }
      this.setState({ seqno: uname + "-" + text.join("") });
      /*-- --*/
    });
    /*-- for current date --*/
    let newDate = new Date();
    let date = newDate.getDate().toString();
    let month = (newDate.getMonth() + 1).toString();
    let year = newDate.getFullYear().toString();

    if (month.toString().length == 1) { month = "0" + month.toString(); }
    if (date.toString().length == 1) { date = "0" + date.toString(); }

    let fullDate = date + "-" + month + "-" + year;
    // document.getElementById("tdDate").innerText = fullDate;
    let overlay = document.getElementById('tdDate');
    if (overlay) {    
      overlay.innerText = fullDate;
    }
    /*--End--*/

    /*-- To get details from masters(lists) --*/
    this.getDepartments();
    this.getFinNotes();
    //this.getDOP();
    this.getFY();
    this.getRestrictedEmails();
    this.fetchCounterData();
    /*--End--*/

  }
  /*-- End Function--*/

  /*-- To get details from Departments master for Department dropdown --*/
  private getDepartments() {
    //debugger;
    pnp.sp.site.rootWeb.lists.getByTitle('Departments').items.select("ID,Title,Dept_Alias").orderBy("ID asc").getAll().then((items: any[]) => {
      //debugger;
      this.setState({DepartmentItems: items });
      let links: string = '';
      for (let i = 0; i < items.length; i++) {
        links += "<option value='" + items[i].Dept_Alias + "'>" + items[i].Title + "</option>";
      }
      jQuery('select[id="ddlDeptOwnership"]').append(links);

    });
  }
  /*--End--*/

  /*-- To get details from FYMaster master --*/
  private getFY() {
    //debugger;
    pnp.sp.site.rootWeb.lists.getByTitle('FYMaster').items.select("ID,Title,Active").orderBy("ID asc").getAll().then((items: any[]) => {
      //debugger;
      let links: string = '';
      for (let i = 0; i < items.length; i++) {
        if (items[i].Active == 'Yes') {
          jQuery('#tdFY').text(items[i].Title);
        }

      }
    });
  }
  /*--End--*/

  /*-- To get details from FinNotes master for Type of Financial Note --*/
  private getFinNotes() {
    //debugger;
    pnp.sp.site.rootWeb.lists.getByTitle('FinNotes').items.select("ID,Title").orderBy("ID asc").getAll().then((items: any[]) => {
      //debugger;
      // console.log(items);
      this.setState({FinNotes: items });
      let links: string = '';
      for (let i = 0; i < items.length; i++) {
        links += "<option value='" + items[i].Title + "'>" + items[i].Title + "</option>";
      }
      jQuery('select[id="ddlFinNote"]').append(links);

    });
  }
  /*--End--*/

  /*-- To get details from DOP master for DOP Details --*/
  // private getDOP() {
  //   //debugger;
  //   pnp.sp.site.rootWeb.lists.getByTitle('DOP').items.select("ID,Title").orderBy("ID asc").getAll().then((items: any[]) => {
  //     //debugger;
  //     this.setState({DOPItems: items });
  //     let links: string = '';
  //     for (let i = 0; i < items.length; i++) {
  //       links += "<option value='" + items[i].Title + "'>" + items[i].Title + "</option>";
  //     }
  //     jQuery('select[id="ddlDOP"]').append(links);

  //   });
  // }
  /*--End--*/

  /*-- To Update Recommanders in Approvals list--*/
  private AddRecommender() {
    //debugger;
    let seqno = this.state.RecommSeqNo + 1;
    let MgrID = this.state.RecpID;
    let userid = this.state.UserID;
    let TotalRecomm = this.state.dpselectedItems;
    let restricedEmails = this.state.RestrictedEmails;
    let restricedEmailsMsg = this.state.RestrictedEmailsMsg;
    if (this.state.RecpName[0] == '') {
      alert('Kindly select username!');
      // $('#RecommenderPPtd >div>div>div>div>div>div>div>input').focus();
      $('#RecommenderPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    else if (TotalRecomm.length == 10) {
      alert('Only 10 Recommenders can be added!');
      // $('#RecommenderPPtd >div>div>div>div>div>div>div>span>div>button>div>i').click();
      $('#RecommenderPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
      // $('#RecommenderPPtd >div>div>div>div>div>div>div>input').focus();
      $('#RecommenderPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    /*
    else if(this.state.RecpEmail[0].toLowerCase()=='arun.mehta@sbicaps.com'){
      alert('Mr. Arun Mehta (MD and CEO) is on mandatory leave from Feb 22,2021 to Mar 05,2021, please select another recommender');
      $('#RecommenderPPtd >div>div>div>div>div>div>div>span>div>button>div>i').click();
      $('#RecommenderPPtd >div>div>div>div>div>div>div>input').focus();
    }
    */
    else if (restricedEmails.indexOf(this.state.RecpEmail[0].toLowerCase()) >= 0) {
      debugger;
      let indx=restricedEmails.indexOf(this.state.RecpEmail[0].toLowerCase());
      let msg = restricedEmailsMsg[indx];
      alert(msg);
      //alert(this.state.RecpEmail[0] + ' cannot be added, kindly select proper name id');
      $('#RecommenderPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
      $('#RecommenderPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    else if (userid == MgrID[0]) {
      alert('Requester cannot be recommender!');
      $('#RecommenderPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
      $('#RecommenderPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    else {

      let mgrEmail = this.state.RecpEmail[0];
      this.checkRecommender(mgrEmail).then((len) => {

        if (len == 0) {
          this.checkApprover(mgrEmail).then((len1) => {

            if (len1 == 0) {
              let SeqNo = this.state.seqno;
              //debugger;
              let web = new Web('Main');

              web.lists.getByTitle('ApprovalsChecklist').items.add({
                Title: this.state.seqno,
                Status: 'Pending',
                Seq: seqno,
                ApproverId: this.state.RecpID[0],
                AppID: this.state.RecpID[0],
                AppName: this.state.RecpName[0],
                AppEmail: this.state.RecpEmail[0]
              }).then((iar: ItemAddResult) => {
                this.setState({ RecommSeqNo: seqno });
                console.log(iar.data.ID);
                $("#RecommenderPPtd .ms-PickerItem-removeButton").trigger("click");
                this.retrieveRecommenders();
                $('#RecommenderPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
              });
            }
            else {
              alert('Approver cannot be Recommender!');
              $('#RecommenderPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");

              return;
            }
          });
        }
        else {
          alert('Recommender has already been added!');
          $('#RecommenderPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");


          return;


        }

      });

    }
  }
  /*--End--*/

  /*-- To Update Approvers in FApprovals list--*/
  private AddApprover() {
    //debugger;
    let seqno = this.state.AppSeqNo + 1;
    let MgrID = this.state.userManagerIDs;
    let userid = this.state.UserID;
    let TotalApp = this.state.selectedItems;
    let controllerflag = "";
    let restricedEmails = this.state.RestrictedEmails;
    let restricedEmailsMsg = this.state.RestrictedEmailsMsg;
    if (jQuery('#ddlDepartment option:selected').val() == "TIG") {
      controllerflag = "Yes";
    }
    if (this.state.MgrName == '') {
      alert('Kindly select username!');

      $('#ApproverPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    /*
    else if(this.state.ManagerEmail[0].toLowerCase()=='arun.mehta@sbicaps.com'){
      alert('Mr. Arun Mehta (MD and CEO) is on mandatory leave from Feb 22,2021 to Mar 05,2021, please select another approver');
      $('#ApproverPPtd >div>div>div>div>div>div>div>span>div>button>div>i').click();
      $('#ApproverPPtd >div>div>div>div>div>div>div>input').focus();
    }
    */
    else if (restricedEmails.indexOf(this.state.ManagerEmail[0].toLowerCase()) >= 0) {
      debugger;
      let indx=restricedEmails.indexOf(this.state.ManagerEmail[0].toLowerCase());
      let msg = restricedEmailsMsg[indx];
      alert(msg);
      //alert(this.state.ManagerEmail[0] + ' cannot be added, kindly select proper name id');
      $('#ApproverPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
      $('#ApproverPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    else if (TotalApp.length == 10) {
      alert('Only 10 Approvers can be added!');
      $('#ApproverPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
      $('#ApproverPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    else if (userid == MgrID[0] && controllerflag != 'Yes') {
      alert('Requester cannot be approver!');
      $('#ApproverPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
      $('#ApproverPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    else {

      let mgrEmail = this.state.ManagerEmail[0];
      console.log(this.state.userManagerIDs[0]);
      console.log(this.state.MgrName);

      this.checkApprover(mgrEmail).then((len) => {
        if (len == 0) {
          this.checkRecommender(mgrEmail).then((len1) => {

            if (len1 == 0) {
              let SeqNo = this.state.seqno;
              let web = new Web('Main');
              //debugger;
              web.lists.getByTitle('FApprovalsChecklist').items.add({
                Title: this.state.seqno,
                Status: 'Pending',
                Seq: seqno,
                ApproverId: this.state.userManagerIDs[0],
                AppID: this.state.userManagerIDs[0],
                AppName: this.state.MgrName,
                AppEmail: this.state.ManagerEmail[0]
              }).then((iar: ItemAddResult) => {
                this.setState({ AppSeqNo: seqno });
                console.log(iar.data.ID);
                // this.retrieveApprovers();
                // $('#ApproverPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");

                $("#ApproverPPtd .ms-PickerItem-removeButton").trigger("click");
                this.retrieveApprovers();
                $('#ApproverPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
              });
            }
            else {
              alert('Recommender cannot be Approver!');
              $('#ApproverPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
              return;
            }
          });

        }

        else {
          alert('Approver has already been added!');
          $('#ApproverPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
          return;


        }

      });

    }

  }
  /*--End--*/
  /*Added on 29/03/2025 */
  fetchCounterData = async () => {
    const num = await this.getCounter();
    if(num[3] == 'Compliance')
    {this.setState({ vettingobservation: 'Vetting Observation' });
    this.setState({Checklistlabel: 'Checklist'});}
    else{this.setState({ vettingobservation: 'Remarks' });
    this.setState({Checklistlabel: 'Annexure'});}    
  };

  // Add Controller before submission
  /*-- To Update Controller in CApprovals list--*/
  private AddController() {
    //debugger;
    let seqno = 1;
    let MgrID = this.state.ccIDS;
    let userid = this.state.UserID;
    let Controllers = this.state.ccSelectedItems;
    let restricedEmails = this.state.RestrictedEmails;
    let restricedEmailsMsg = this.state.RestrictedEmailsMsg;

    if (this.state.ccName[0] == '') {
      alert('Kindly select username!');
      $('#ControllerPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    else if (Controllers.length > 0) {
      alert('Only 1 Controller can be added!');
      $('#ControllerPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
      $('#ControllerPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    /*
    else if(this.state.ccEmail[0].toLowerCase()=='arun.mehta@sbicaps.com'){
      alert('Mr. Arun Mehta (MD and CEO) is on mandatory leave from Feb 22,2021 to Mar 05,2021, please select another approver');
      $('#ControllerPPtd >div>div>div>div>div>div>div>span>div>button>div>i').click();
      $('#ControllerPPtd >div>div>div>div>div>div>div>input').trigger("focus");
    }
    */
    else if (restricedEmails.indexOf(this.state.ccEmail[0].toLowerCase()) >= 0) {
      debugger;
      let indx=restricedEmails.indexOf(this.state.ccEmail[0].toLowerCase());
      let msg = restricedEmailsMsg[indx];
      alert(msg);
      //alert(this.state.ccEmail[0] + ' cannot be added, kindly select proper name id');
      $('#ControllerPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
      $('#ControllerPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    else if (userid == MgrID[0]) {
      alert('Requester cannot be Controller!');
      $('#ControllerPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
      $('#ControllerPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    else {

      let mgrEmail = this.state.ccEmail[0];
      this.setState({ controllerName: this.state.ccName[0] });
      this.setState({ controllerPPId: this.state.ccIDS[0] });
      this.checkApprover(mgrEmail).then((len) => {
        if (len == 0) {
          this.checkRecommender(mgrEmail).then((len1) => {

            if (len1 == 0) {
              let SeqNo = this.state.seqno;
              let web = new Web('Main');
              //debugger;
              web.lists.getByTitle('CApprovalsChecklist').items.add({
                Title: this.state.seqno,
                Status: 'Pending',
                Seq: seqno,
                ApproverId: this.state.ccIDS[0],
                AppID: this.state.ccIDS[0],
                AppName: this.state.ccName[0],
                AppEmail: this.state.ccEmail[0]
              }).then((iar: ItemAddResult) => {
                this.setState({ AppSeqNo: seqno });
                console.log(iar.data.ID);
                // this.retrieveController();
                // $('#ControllerPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");

                //$("#ControllerPPtd .ms-PickerItem-removeButton").trigger("click");
                this.retrieveController();
                $('#ControllerPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
              });
            }
            else {
              alert('Recommender cannot be Controller!');
              $('#ControllerPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
              return;
            }
          });

        }

        else {
          alert('Approver has already been added!');
          $('#ControllerPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
          return;


        }

      });

    }
  }
  /*--End--*/

  /*-- To Check adding approver present in FApprovals list or not--*/
  private checkApprover(appemail: string): Promise<number> {
    //debugger;
    let title = this.state.seqno;
    let len = 0;
    let web = new Web('Main');
    return web.lists.getByTitle('FApprovalsChecklist').items.select("ID,Title,AppName,AppEmail").filter("Title eq '" + title + "'").orderBy("Seq asc").getAll().then((items: any[]) => {

      for (let i = 0; i < items.length; i++) {
        if (items[i].AppEmail == appemail) {
          len = 1;
        }
      }

      return Promise.resolve(len);
    });

  }
  /*--End--*/

  /*-- To Check adding recommender present in Approvals list or not--*/
  private checkRecommender(appemail: string): Promise<number> {
    //debugger;
    let title = this.state.seqno;
    let len = 0;
    let web = new Web('Main');
    return web.lists.getByTitle('ApprovalsChecklist').items.select("ID,Title,AppName,AppEmail").filter("Title eq '" + title + "'").orderBy("Seq asc").getAll().then((items: any[]) => {

      for (let i = 0; i < items.length; i++) {
        if (items[i].AppEmail == appemail) {
          len = 1;
        }
      }

      return Promise.resolve(len);
    });

  }
  /*--End--*/

  /*-- To retrieve approvers from FApprovals List--*/
  private retrieveApprovers() {
    let title = this.state.seqno;
    // let data = [];
    let data: any[] = []; 
    let web = new Web('Main');
    web.lists.getByTitle('FApprovalsChecklist').items.select("ID,Title,AppName").filter("Title eq '" + title + "' ").orderBy("Seq asc").getAll().then((items: any[]) => {
      //debugger;
      if (items.length > 0) {
        for (let i = 0; i < items.length; i++) {
          data.push(<tr><td style={{paddingLeft:'10px'}}>{i + 1}</td><td style={{paddingLeft:'10px'}}>{items[i].AppName}</td><td><button className='btn btn-sm' onClick={() => { this.DeleteApprover(items[i].ID); }}><Icon style={{color:'#f00'}} iconName="delete" /></button></td></tr>);
        }
      }

    }).then(() => {
      this.setState({ selectedItems: data });
    });

  }
  /*--End--*/

  /*-- To retrieve recommanders from Approvals List--*/
  private retrieveRecommenders() {
    let title = this.state.seqno;
    // let data = [];
    let data: any[] = [];
    let web = new Web('Main');
    web.lists.getByTitle('ApprovalsChecklist').items.select("ID,Title,AppName").filter("Title eq '" + title + "' ").orderBy("Seq asc").getAll().then((items: any[]) => {
      //debugger;
      if (items.length > 0) {
        for (let i = 0; i < items.length; i++) {
          data.push(<tr><td style={{paddingLeft:'10px'}}>{i + 1}</td><td style={{paddingLeft:'10px'}}>{items[i].AppName}</td><td><button className='btn btn-sm' onClick={() => { this.DeleteRecommender(items[i].ID); }}><Icon style={{color:'#f00'}} iconName="delete" /></button></td></tr>);
        }
      }

    }).then(() => {
      this.setState({ dpselectedItems: data});
    });

  }
  /*--End--*/

  /*-- To retrieve controller from CApprovals List--*/
  private retrieveController() {
    console.log("Seqno:" +this.state.seqno+" \nControllerID" +this.state.ccIDS);
    let title = this.state.seqno;
    // let data = [];
    let data: any[] = [];
    let ControllerID = this.state.ccIDS;
    let web = new Web('Main');
    web.lists.getByTitle('CApprovalsChecklist').items.select("ID,Title,AppName").filter("Title eq '" + title + "' and AppID eq " + ControllerID[0]).orderBy("Seq asc").getAll().then((items: any[]) => {
      //debugger;
      if (items.length > 0) {
        for (let i = 0; i < items.length; i++) {
          data.push(<tr><td style={{paddingLeft:'10px'}}>{i + 1}</td><td style={{paddingLeft:'10px'}}>{items[i].AppName}</td><td><button  className='btn btn-sm' onClick={() => { this.DeleteController(items[i].ID); }}><Icon style={{color:'#f00'}} iconName="delete" /></button></td></tr>);
        }
      }

    }).then(() => {
      this.setState({ ccSelectedItems: data });
    });

  }
  /*--End--*/

  /*-- To Delete approvers in FApprovals List--*/
  public DeleteApprover(uid: number, event?: React.MouseEvent<HTMLButtonElement>): void {
    //debugger;
    event?.preventDefault();
    let web = new Web('Main');

    let list = web.lists.getByTitle('FApprovalsChecklist');
    list.items.getById(uid).delete().then(() => {
      console.log('List Item Deleted');
      this.retrieveApprovers();
    });

  }
  /*--End--*/

  /*-- To Delete controller in CApprovals List--*/
  public DeleteController(uid: number, event?: React.MouseEvent<HTMLButtonElement>): void {
    //debugger;
    event?.preventDefault();
    let web = new Web('Main');

    let list = web.lists.getByTitle('CApprovalsChecklist');
    list.items.getById(uid).delete().then(() => {
      console.log('List Item Deleted');
      this.retrieveController();
      this.setState({ ccSelectedItems: [] });
      this.setState({ controllerName: '' });
      this.setState({ controllerPPId: 0 });

    });

  }
  /*--End--*/

  /*-- To Delete recommender in Approvals List--*/
  public DeleteRecommender(uid: number, event?: React.MouseEvent<HTMLButtonElement>): void {
    //debugger;
    event?.preventDefault();
    let web = new Web('Main');

    let list = web.lists.getByTitle('ApprovalsChecklist');
    list.items.getById(uid).delete().then(() => {
      console.log('List Item Deleted');
      this.retrieveRecommenders();
    });

  }

  /*--End--*/

  /*-- To get first approver in FApprovals List(to set current approver while submit)--*/
  private retrieveFirstApprover(): Promise<any[]> {
    let title = this.state.seqno;
    // let approverID = [];
    let approverID: any[] = [];
    let web = new Web('Main');
    return web.lists.getByTitle('FApprovalsChecklist').items.select("ID,Title,AppName,Approver/ID,Approver/Title").filter("Title eq '" + title + "'").expand("Approver").orderBy("ID asc").getAll().then((items: any[]) => {
      //debugger;
      this.setState({ MgrName: items[0].Approver.Title });
      approverID[0] = items[0].Approver.ID;
      approverID[1] = items[0].ID;
      return approverID;

    });
  }
  /*--End--*/

  /*-- To update  first approver in FApprovals List--*/
  private updateFirstApprover(uid: number): Promise<any[]> {
    let web = new Web('Main');
    return web.lists.getByTitle('FApprovalsChecklist').items.getById(uid).update({
      Status: 'Submitted'
    }).then(() => {
      console.log('Approver updated');
      return Promise.resolve(['Done']);

    });

  }
  /*--End--*/

  /*-- To get first recommander in Approvals List(to set current approver while submit)--*/
  private retrieveFirstRecommender(): Promise<any[]> {
    let title = this.state.seqno;
    // let approverID = [];
    let approverID: any[] = [];
    let web = new Web('Main');
    return web.lists.getByTitle('ApprovalsChecklist').items.select("ID,Title,AppName,Approver/ID,Approver/Title").filter("Title eq '" + title + "'").expand("Approver").orderBy("ID asc").getAll().then((items: any[]) => {
      //debugger;
      this.setState({ MgrName: items[0].Approver.Title });
      approverID[0] = items[0].Approver.ID;
      approverID[1] = items[0].ID;
      return approverID;

    });
  }
  /*--End--*/

  /*-- To update  first recommander in Approvals List--*/
  private updateFirstRecommender(uid: number): Promise<any[]> {
    let web = new Web('Main');
    return web.lists.getByTitle('ApprovalsChecklist').items.getById(uid).update({
      Status: 'Submitted'
    }).then(() => {
      console.log('Approver updated');
      return Promise.resolve(['Done']);
    });
  }
  /*--End--*/

  /*-- To add work flow histori in WFHistory list--*/
  private AddWFHistory(): Promise<any[]> {
    let dt = new Date();
    let mnth = (dt.getMonth() + 1).toString();
    let dat = dt.getDate().toString();
    let hrs = dt.getHours().toString();
    let mins = dt.getMinutes().toString();
    let secs = dt.getSeconds().toString();
    if (mnth.length == 1) { mnth = '0' + mnth; } if (dat.length == 1) { dat = '0' + dat; } if (hrs.length == 1) { hrs = '0' + hrs; } if (mins.length == 1) { mins = '0' + mins; } if (secs.length == 1) { secs = '0' + secs; }
    let createDate = dat + "-" + mnth + "-" + dt.getFullYear() + " " + hrs + ":" + mins + ":" + secs;
    let log = 'Submitted to ' + this.state.MgrName + ' by ' + this.state.name + ' on ' + createDate;
    //debugger;
    let web = new Web('Main');
    return web.lists.getByTitle('WFHistory').items.add({
      Title: this.state.seqno,
      AuditLog: log,
      Currapprover: this.state.MgrName,
      FormName: 'Note',
      ActionDateTime: createDate
    }).then((iar: ItemAddResult) => {
      console.log('History Log Created!');
      return Promise.resolve(['Done']);

    });

  }
  /*--End--*/

  /*-- Department change function--*/
  private ChangeDepartment() {
    let dept = jQuery('#ddlDepartment option:selected').val();
    //if(dept=='HRD'){   //Commented by Surendra at 28/1/2022 : As per Sagar mail, divConfidential is available for all department 
    if (dept != "Select") {
      let overlay = document.getElementById('divConfidential');
      if (overlay) {
        overlay.style.display = 'block';
      }
    }
    else {
      // document.getElementById('divConfidential').style.display = 'none';
      let overlay = document.getElementById('divConfidential');
      if (overlay) {    
        overlay.style.display = 'none';
      }
    }

  }
  /*--End--*/

  /*-- Note Type change function--*/
  private SelectSource() {
    let source = jQuery('#ddlSource option:selected').val();
    if (source == 'Financial') {
      jQuery('.FinancialClass').css('display', 'block');
    }
    else {
      jQuery('.FinancialClass').css('display', 'none');
    }

  }
  /*--End--*/

  /*-- To save name,email and id for controller people picker--*/
  private _getCCPeople(items: any[]) {//debugger;
    this.state.ccIDS.length = 0;
    let Recpid = [];
    let Recpname = [];
    let Recpemail = [];

    for (let item in items) {
      Recpid.push(items[item].id);
      Recpname.push(items[item].text);
      Recpemail.push(items[item].loginName.split("|")[2]);
      // alert(items[item].id);
    }
    this.setState({ ccName: Recpname });
    this.setState({ ccIDS: Recpid });
    this.setState({ ccEmail: Recpemail });
    $('#lblController').text(Recpid[0]);
    setTimeout(() => {
      if (this.state.ccIDS.length == 1) { this.AddController(); }
    }, 1000);
  }
  /*--End--*/

  /*-- To save name,email and id for recommander people picker--*/
  private _getReceivedFrom(items: any[]) {debugger;
    this.state.RecpID.length = 0;
    let Recpid = [];
    let Recpname = [];
    let Recpemail = [];
    if (items.length > 0) {
      this.setState({ isChecked: true });
      for (let item in items) {
        Recpid.push(items[item].id);
        Recpname.push(items[item].text);
        Recpemail.push(items[item].loginName.split("|")[2]);

      }

      this.setState({ RecpID: Recpid });
      this.setState({ RecpName: Recpname });
      this.setState({ RecpEmail: Recpemail });
      setTimeout(() => {
        if (items.length > 0) {
          this.AddRecommender();
        }
      }, 1000);
    } // Ending If of items.length

  }
  /*--End--*/

  /*-- To save name,email and id for approver people picker--*/
  private _getManager(items: any[]) {
    //debugger;
    this.state.userManagerIDs.length = 0;
    let tempuserMngArr = [];
    let MgrEmail = [];
    let MgrName = '';
    for (let item in items) {
      tempuserMngArr.push(items[item].id);
      MgrName = items[item].text;
      MgrEmail.push(items[item].loginName.split("|")[2]);
    }
    this.setState({ userManagerIDs: tempuserMngArr });
    this.setState({ ManagerEmail: MgrEmail });
    this.setState({ MgrName: MgrName });

    setTimeout(() => {
      if (items.length > 0) {
        this.AddApprover();
      }
    }, 1000);
  }
  /*--End--*/

  /*-- On Submission display panel--*/
  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <PrimaryButton id="Createbutton" onClick={this.createItem} style={{ marginRight: '5px', width: "25pt" }}>
          Confirm
        </PrimaryButton>
        < PrimaryButton id="Cancelbutton" style={{ marginLeft: '5px', width: "25pt" }} onClick={this._onClosePanel}>Cancel</PrimaryButton>
      </div>
    );
  }
  /*-- End Function--*/

  /*-- cancel button logic --*/
  private cancel = () => {
    this.setState({ showPanel: false });
    // self.close();
    const query = window.location.search.split('uid=')[1];
    let uid = 0;
    if (query != undefined) { uid = parseInt(query); }
    if (uid == 0) {
      window.location.replace(this.props.siteUrl);
    }
    else {
      window.location.replace(this.props.siteUrl);

    }
  }
  /*--End --*/
  /*--close panel --*/
  private _onClosePanel = () => {
    this.setState({ showPanel: false });

  }
  /*--End --*/

  /*-- Redirecting page logic --*/
  private redirect() {
    let sitename = this.state.Sitename;
    const query = window.location.search.split('uid=')[1];
    let uid = 0;
    if (query != undefined) { uid = parseInt(query); }
    if (uid == 0) {
      window.location.replace(sitename);
    }
    else {
      setTimeout(() => {
        window.location.replace(sitename);
      }, 3000);
    }
  }
  /*-- End --*/
  /*-- Show Panel Logic--*/
  private _onShowPanel = () => {
    this.setState({ showPanel: true });
  }
  /*-- End Function--*/

  /*-- Set Title Function--*/
  private handleTitle(value: string): void {
    return this.setState({
      name: value
    });
  }
  /*-- End Function--*/
  /*-- Set dscription state Function--*/
  private handleDesc(value: string): void {
    return this.setState({
      description: value
    });
  }
  /*-- End Function--*/

  /*--Close dialog function--*/
  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }
  /*--End--*/

  /*--show dialog logic--*/
  private _showDialog = (status: string): void => {
    this.setState({ hideDialog: false });
    this.setState({ status: status });
  }
  /*--End--*/
  /**
   * Form Validation On Submission
   */
  /*--End--*/
  /*--Form Submit validation--*/
  private validateForm(): void {
    this.setState({ onSubmission: true });
    let allowCreate = true;
    
    const getValue = (selector: string): string => String(jQuery(selector).val() || '');
    const focusElement = (selector: string): void => {
      jQuery(selector).focus();
  };
    const showAlert = (message: string, selector: string): void => {
        alert(message);
        focusElement(selector);
    };
    
    const Notefor = getValue('#txtNote');
    const Subject = getValue('#txtSubject');
    const Purpose = getValue('#txtPurpose');
    const Productname = getValue('#txtReturn');
    const ddlDeptOwnership = getValue('#ddlDeptOwnership');
    const referredguidelines = getValue('#txtGuidelines');
    const Vetting = getValue('#txtVetting');
    const Financial = getValue('#ddlSource');
    const FinType = getValue('#ddlFinNote');
    const Amount = getValue('#Amount');
    //const DOP = getValue('#ddlDOP');
    const Approvers = this.state.selectedItems;
    const filename = this.state.Notefilename;
    const Checklisttable = this.state.items;
    const regx = /^[A-Za-z0-9 !@#$()_.-]+$/;
    
    if (!Notefor) { showAlert('Kindly enter Note For!', '#txtNote'); return; }
    if (!Subject) { showAlert('Kindly enter Subject!', '#txtSubject'); return; }
    if (Subject.indexOf('http://') > -1 || Subject.indexOf('https://') > -1) {
        showAlert('Kindly do not enter URLs in Subject!', '#txtSubject'); return;
    }
    // if (!regx.test(Subject)) { showAlert('Subject contains special characters!', '#txtSubject'); return; }
    if (Subject.length > 250) { showAlert('Max 250 chars are allowed in Subject!', '#txtSubject'); return; }
    if (!Purpose) { showAlert('Kindly enter Purpose!', '#txtPurpose'); return; }
    //if (!Productname) { showAlert('Kindly enter Product Name!', '#txtReturn'); return; }
    if (ddlDeptOwnership === 'Select') { showAlert('Kindly select the Department Ownership!', '#ddlDeptOwnership'); return; }
    //if (!referredguidelines) { showAlert('Kindly enter the Referred Guidelines!', '#txtGuidelines'); return; }
    //if (!Vetting) { showAlert('Kindly enter Vetting Observation!', '#txtVetting'); return; }
    //if (Checklisttable.length === 0) { showAlert('Kindly enter Checklist!', '#ddlDeptOwnership'); return; }
    
    if (Financial === 'Select') { showAlert('Kindly select the Note Type!', '#ddlSource'); return; }
    if (Financial !== 'Financial' && Financial !== 'Non-Financial') { showAlert('Invalid Note Type selected!', '#ddlSource'); return; }
    if (Financial === 'Financial' && FinType === 'Select') { showAlert('Kindly select the Financial Note Type!', '#ddlFinNote'); return; }
    if (Financial === 'Financial' && (isNaN(Number(Amount)) || !Amount || Amount === '0')) {
        showAlert('Kindly enter a valid Amount!', '#Amount'); return;
    }
    //if (Financial === 'Financial' && DOP === 'Select') { showAlert('Kindly select the DOP details!', '#ddlDOP'); return; }
    
    if (Approvers.length === 0) { showAlert('Kindly select at least one Approver!', '#btnAddApprover'); return; }
    if (!filename) { showAlert('Kindly select at least one Main Note!', '#ddlTemplate'); return; }
    
    // If all validations pass
    this._onShowPanel();
}

  /*--End--*/
  /*--Save Draft function--*/
  private SaveDraft(): void {

    //debugger;
    jQuery('#btnDraft').remove();
    jQuery('#Createbutton').remove();
    jQuery('#Cancelbutton').remove();
    let dept = jQuery('#ddlDepartment option:selected').text();
    let OrginatingDept = jQuery('#ddlRecpDepartment option:selected').text();
    let deptAlias = jQuery('#ddlDepartment option:selected').val();

    let title = '';
    if (deptAlias != "Select") {
      this.getCounter().then((countVal) => {
        let DeptGroupID = parseInt(countVal[2]);
        // title = "Note/" + deptAlias + "/Draft";
        title = "Note/" + countVal[3] + "/Draft";
      });

    }
    else {
      title = deptAlias + ":112/Note/Draft";
    }

  }
  /*--End--*/

  /*--Record Submit function to update lists--*/
  private createItem(): void {
    //debugger;
    this._onClosePanel();
    this.on();
    jQuery('#btnDraft').remove();
    jQuery('#Createbutton').remove();
    jQuery('#Cancelbutton').remove();
    let FY = jQuery('#tdFY').text();
    //let dept = jQuery('#ddlDepartment option:selected').text();
    //let deptAlias = jQuery('#ddlDepartment option:selected').val();
    let counter = 0;
    let uid = 0;
    let Financial = jQuery('#ddlSource option:selected').text();
    let FinType = jQuery('#ddlFinNote').val();
    //let DOP = jQuery('#ddlDOP').val();
    let Amount = jQuery('#Amount').val();
    let Exceptional = jQuery('#txtExceptional').val();
    let Confidential = jQuery('#txtConfidential').val();

    //added on 16/02/2025
    let Notefor = jQuery('#txtNote').val();
    let Purpose = jQuery('#txtPurpose').val();
    let ReturnName = jQuery('#txtReturn').val();
    let DeptOwnership = jQuery('#ddlDeptOwnership option:selected').text();
    let VettingObservation = jQuery('#txtVetting').val();
    let RefferedGuidlines = jQuery('#txtGuidelines').val();
    // let DueDate = this.state.selectedDate;
    // let Place = jQuery('#txtPlace').val();
    let Checklisttable = this.state.items;
    if (Financial == 'Financial') {
      Financial = String(FinType);
    }
    if (Amount == '') {
      Amount = 0;
    }
    let Recommenders = this.state.dpselectedItems.length;

    let filename = this.state.Notefilename;

    //this.getCounter(String(deptAlias)).then((countVal) => {
      this.getCounter().then((countVal) => {
      counter = parseInt(countVal[0]);
      uid = parseInt(countVal[1]);
      let deptAlias = countVal[3];
      let dept = countVal[3];
      let DeptGroupID = parseInt(countVal[2]);
      let Subj = jQuery('#txtSubject').val();
      let Comment = jQuery('#txtComments').val();      
      let client = jQuery('#txtClient').val();
      let requester = this.state.UserID;
      let dt = new Date();
      let mnth = ("0" + ((dt.getMonth() + 1).toString())).slice(-2);
      let dat = ("0" + (dt.getDate().toString())).slice(-2);
      let fulldate = dat + mnth + dt.getFullYear().toString();
      let title = "Note/" + deptAlias + "/" + fulldate + "/" + counter.toString();
      // let ControllerID = $('#lblController').text();
      let ControllerID = 0;
      var checklistId : number ;
      // if (ControllerID == '') {ControllerID = String(this.state.ccIDS[0]); }
      // else { ControllerID = String(parseInt($('#lblController').text())); }

      if(this.state.ccIDS[0] != undefined){ControllerID = parseInt($('#lblController').text());}
      
      
      this.setState({ attachments: [] });
      if (Recommenders > 0) {
        this.retrieveFirstRecommender().then((Appid) => {          
          var approverID = Appid[0];           
          var AppItemid = Appid[1];
          let web = new Web('Main');
          let Approvers : Number[] = [];
          Approvers.push(approverID);
          console.log("SeqNo: "+this.state.seqno);
          web.lists.getByTitle('ChecklistNote').items.add({
            Title: title,
            SeqNo: this.state.seqno,
            Subject: Subj,
            Department: dept,
            Comments: Comment,
            Exceptional: Exceptional,
            Confidential: Confidential,
            CurApproverId: approverID,
            ApproversId: { results: Approvers },
            NotifyId: approverID,
            Amount: Amount,
            RequesterId: requester,
            NoteFilename: filename,
            NoteType: Financial,
            //DOP: DOP,
            DeptAlias: deptAlias,
            ClientName: client,
            Migrate: "",
            FY: FY,
            DeptGroupId: DeptGroupID,
            ControllerId: ControllerID,
            Status: "Submitted to Recommender#1",
            StatusNo: 1,
            Notefor : Notefor,
            Purpose : Purpose,
            ReturnName : ReturnName,
            DeptOwnership : DeptOwnership,
            RefferedGuidlines:RefferedGuidlines,
            VettingObservation:VettingObservation
            //DueDate : DueDate,
            //Place : Place
          }).then((iar: ItemAddResult) => {
            console.log(iar.data.ID);
            let id = iar.data.ID;
            checklistId = iar.data.ID;
            pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.add({
              Title: title,
              Subject: Subj,
              Department: dept,
              NoteType: Financial,
              Exceptional: Exceptional,
              Confidential: Confidential,
              SeqNo: this.state.seqno,
              PID: id,
              FY: FY,
              DeptAlias: deptAlias,
              CurApproverTxt: this.state.MgrName,
              ClientName: client,
              DeptGroupId: DeptGroupID,
              CurApproverId: approverID,
              RequesterId: requester,
              NoteFilename: filename,
              Sitename: 'Main',
              Status: "Submitted to Recommender#1",
              StatusNo: 1,
              Notefor : Notefor,
              Purpose : Purpose,
              ReturnName : ReturnName,
              DeptOwnership : DeptOwnership,
              RefferedGuidlines:RefferedGuidlines,
              VettingObservation:VettingObservation
              // DueDate : DueDate,
              // Place : Place,              
            })
              .then((iar1: ItemAddResult) => {
                let WFweb = new Web('WF');
                WFweb.lists.getByTitle('ChecklistNoteNotifications').items.add({
                  Title: title,
                  SeqNo: this.state.seqno,
                  Subject: Subj,
                  Department: dept,
                  Comments: Comment,
                  CurApproverId: approverID,
                  ApproversId: { results: Approvers },
                  NotifyId: approverID,
                  Amount: Amount,
                  RequesterId: requester,
                  NoteFilename: filename,
                  NoteType: Financial,
                  //DOP: DOP,
                  DeptAlias: deptAlias,
                  ClientName: client,
                  Migrate: "",
                  FY: FY,
                  MainRecID: id,
                  DeptGroupId: DeptGroupID,
                  ControllerId: ControllerID,
                  Status: "Submitted to Recommender#1",
                  StatusNo: 1,
                  Notefor : Notefor,
                  Purpose : Purpose,
                  ReturnName : ReturnName,
                  DeptOwnership : DeptOwnership,
                  RefferedGuidlines:RefferedGuidlines,
                  VettingObservation:VettingObservation
                  // DueDate : DueDate,
                  // Place : Place                  
                }).then((iar: ItemAddResult) =>{
                  for(var i=0;i<Checklisttable.length;i++)
                  {
                    pnp.sp.site.rootWeb.lists.getByTitle('Checklist').items.add({
                    Title: title,
                    SeqNo: this.state.seqno,
                    AppId: checklistId ,
                    Checklist:Checklisttable[i].checklist,
                    Status:Checklisttable[i].status
                    });
                  }                                
                }).then(() => {
                  this.setCounter(uid, counter).then(() => {
                    this.updateFirstRecommender(Number(AppItemid)).then(() => {
                      this.AddWFHistory().then(() => {
                        this.redirect();
                      });
                    });
                  });
                });
              });
          });
        });

      } else {
        this.retrieveFirstApprover().then((Appid) => {
          let approverID = Appid[0];
          let AppItemid = Appid[1];
          let web = new Web('Main');
          let Approvers : Number[] = [];
          Approvers.push(approverID);
          // web.lists.getByTitle('Notes').items.add({
            web.lists.getByTitle('ChecklistNote').items.add({
            Title: title,
            SeqNo: this.state.seqno,
            Subject: Subj,
            Department: dept,
            Comments: Comment,
            //DOP: DOP,
            Exceptional: Exceptional,
            Confidential: Confidential,
            CurApproverId: approverID,
            NotifyId: approverID,
            ApproversId: { results: Approvers },
            Amount: Amount,
            RequesterId: requester,
            NoteFilename: filename,
            NoteType: Financial,
            DeptAlias: deptAlias,
            ClientName: client,
            Migrate: "",
            FY: FY,
            DeptGroupId: DeptGroupID,
            ControllerId: ControllerID,
            Status: "Submitted to Approver#1",
            StatusNo: 6,
            Notefor : Notefor,
            Purpose : Purpose,
            ReturnName : ReturnName,
            DeptOwnership : DeptOwnership,
            RefferedGuidlines:RefferedGuidlines,
            VettingObservation:VettingObservation
            // DueDate : DueDate,
            // Place : Place,                        
          }).then((iar: ItemAddResult) => {
            console.log(iar.data.ID);
            let id = iar.data.ID;
            checklistId = iar.data.ID;
            pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.add({
              Title: title,
              Subject: Subj,
              Department: dept,
              SeqNo: this.state.seqno,
              NoteType: Financial,
              PID: id,
              FY: FY,
              ClientName: client,
              Exceptional: Exceptional,
              Confidential: Confidential,
              CurApproverTxt: this.state.MgrName,
              DeptGroupId: DeptGroupID,
              CurApproverId: approverID,
              RequesterId: requester,
              NoteFilename: filename,
              Sitename: this.state.Sitename,
              Status: "Submitted to Approver#1",
              StatusNo: 6,
              Notefor : Notefor,
              Purpose : Purpose,
              ReturnName : ReturnName,
              DeptOwnership : DeptOwnership,
              RefferedGuidlines:RefferedGuidlines,
              VettingObservation:VettingObservation
              // DueDate : DueDate,
              // Place : Place
            }).then((iar1: ItemAddResult) => {
              let WFweb = new Web('WF');
              WFweb.lists.getByTitle('ChecklistNoteNotifications').items.add({
                Title: title,
                SeqNo: this.state.seqno,
                Subject: Subj,
                Department: dept,
                Comments: Comment,
                //DOP: DOP,
                Exceptional: Exceptional,
                Confidential: Confidential,
                CurApproverId: approverID,
                NotifyId: approverID,
                ApproversId: { results: Approvers },
                Amount: Amount,
                RequesterId: requester,
                NoteFilename: filename,
                NoteType: Financial,
                DeptAlias: deptAlias,
                ClientName: client,
                Migrate: "",
                MainRecID: id,
                FY: FY,
                DeptGroupId: DeptGroupID,
                ControllerId: ControllerID,
                Status: "Submitted to Approver#1",
                StatusNo: 6,
                Notefor : Notefor,
                Purpose : Purpose,
                ReturnName : ReturnName,
                DeptOwnership : DeptOwnership,
                RefferedGuidlines:RefferedGuidlines,
                VettingObservation:VettingObservation
                // DueDate : DueDate,
                // Place : Place
              }).then((iar: ItemAddResult) =>{
                for(var i=0;i<Checklisttable.length;i++)
                {
                  pnp.sp.site.rootWeb.lists.getByTitle('Checklist').items.add({
                  Title: title,
                  SeqNo: this.state.seqno,
                  AppId: checklistId ,
                  Checklist:Checklisttable[i].checklist,
                  Status:Checklisttable[i].status
                  });
                }                                
              }).then(() => {
                this.setCounter(uid, counter).then(() => {
                  this.updateFirstApprover(AppItemid).then(() => {
                    this.AddWFHistory().then(() => {
                      this.redirect();
                    });
                  });
                });
              });
            });
          });
        });
      }


    });
  }  
  /*--End--*/

  /*--get counter from Department List--*/
  //commented on 16/02/2025
  // private getCounter(dept: string): Promise<any[]> {
  //   let num : Number[] = [];
  //   return pnp.sp.site.rootWeb.lists.getByTitle('Departments').items.select("ID,Title,Dept_Alias,Counter,GroupID").filter("Dept_Alias eq '" + dept + "'").orderBy("ID asc").getAll().then((items: any[]) => {
  //     num[0] = parseInt(items[0].Counter) + 1;
  //     num[1] = items[0].ID;
  //     num[2] = items[0].GroupID;
  //     return num;
  //   });

  // }

  
  private getCounter(): Promise<any[]> {
      let num : Number[] = [];
      return pnp.sp.site.rootWeb.lists.getByTitle('Counter').items.select("ID,Title,NoteId,MemoCounter,Department,GroupID").orderBy("ID asc").getAll().then((items: any[]) => {
        num[0] = parseInt(items[0].NoteId) + 1;
        num[1] = items[0].ID;
        num[2] = items[0].GroupID;
        num[3] = items[0].Department;
        return num;
      });  
  }
  /*--End--*/

  /*--Update increment count value in department list--*/
  // private setCounter(uid: number, counter: number): Promise<any[]> {
  //   return pnp.sp.site.rootWeb.lists.getByTitle('Departments').items.getById(uid).update({
  //     Counter: counter
  //   }).then(() => {
  //     console.log('Counter updated');
  //     return Promise.resolve(['Done']);

  //   });

  // }

  private setCounter(uid: number, counter: number): Promise<any[]> {
      return pnp.sp.site.rootWeb.lists.getByTitle('Counter').items.getById(uid).update({
        NoteId: counter
      }).then(() => {
        console.log('Counter updated');
        return Promise.resolve(['Done']);
  
      });
  
    }

    private SaveDraftNew(): void {
      //debugger;
      this._onClosePanel();
      this.on();
      jQuery('#btnDraft').remove();    
      jQuery('#Createbutton').remove();
      jQuery('#Cancelbutton').remove();
      let FY = jQuery('#tdFY').text();
      //let dept = jQuery('#ddlDepartment option:selected').text();
      //let deptAlias = jQuery('#ddlDepartment option:selected').val();
      let counter = 0;
      let uid = 0;
      let Financial = jQuery('#ddlSource option:selected').text();
      let FinType = jQuery('#ddlFinNote').val();
      //let DOP = jQuery('#ddlDOP').val();
      let Amount = jQuery('#Amount').val();
      let Exceptional = jQuery('#txtExceptional').val();
      let Confidential = jQuery('#txtConfidential').val();
    
      //added on 16/02/2025
      let Notefor = jQuery('#txtNote').val();
      let Purpose = jQuery('#txtPurpose').val();
      let ReturnName = jQuery('#txtReturn').val();
      let DeptOwnership = jQuery('#ddlDeptOwnership option:selected').text();
      let VettingObservation = jQuery('#txtVetting').val();
      let RefferedGuidlines = jQuery('#txtGuidelines').val();
      // let DueDate = this.state.selectedDate;
      // let Place = jQuery('#txtPlace').val();
      let Checklisttable = this.state.items;
      if (Financial == 'Financial') {
        Financial = String(FinType);
      }
      if (Amount == '') {
        Amount = 0;
      }
      let Recommenders = this.state.dpselectedItems.length;
    
      let filename = this.state.Notefilename;
    
      //this.getCounter(String(deptAlias)).then((countVal) => {
        this.getCounter().then((countVal) => {
        counter = parseInt(countVal[0]);
        uid = parseInt(countVal[1]);
        let deptAlias = countVal[3];
        let dept = countVal[3];
        let DeptGroupID = parseInt(countVal[2]);
        let Subj = jQuery('#txtSubject').val();
        let Comment = jQuery('#txtComments').val();      
        let client = jQuery('#txtClient').val();
        let requester = this.state.UserID;
        let dt = new Date();
        let mnth = ("0" + ((dt.getMonth() + 1).toString())).slice(-2);
        let dat = ("0" + (dt.getDate().toString())).slice(-2);
        let fulldate = dat + mnth + dt.getFullYear().toString();
        let title = "Note/" + deptAlias + "/" + fulldate + "/" + counter.toString();
        // let ControllerID = $('#lblController').text();
        let ControllerID = 0;
        var checklistId : number ;
        // if (ControllerID == '') {ControllerID = String(this.state.ccIDS[0]); }
        // else { ControllerID = String(parseInt($('#lblController').text())); }
    
        if(this.state.ccIDS[0] != undefined){ControllerID = parseInt($('#lblController').text());}
        
        
        this.setState({ attachments: [] });                               
            let web = new Web('Main');
            let Approvers : Number[] = [];            
            console.log("SeqNo: "+this.state.seqno);
            web.lists.getByTitle('ChecklistNote').items.add({
              Title: title,
              SeqNo: this.state.seqno,
              Subject: Subj,
              Department: dept,
              Comments: Comment,
              Exceptional: Exceptional,
              Confidential: Confidential,              
              ApproversId: { results: Approvers },              
              Amount: Amount,
              RequesterId: requester,
              NoteFilename: filename,
              NoteType: Financial,
              //DOP: DOP,
              DeptAlias: deptAlias,
              ClientName: client,
              Migrate: "",
              FY: FY,
              DeptGroupId: DeptGroupID,
              ControllerId: ControllerID,
              Status: "Draft",
              StatusNo: 12,
              Notefor : Notefor,
              Purpose : Purpose,
              ReturnName : ReturnName,
              DeptOwnership : DeptOwnership,
              RefferedGuidlines:RefferedGuidlines,
              VettingObservation:VettingObservation
              
            }).then((iar: ItemAddResult) => {
              console.log(iar.data.ID);
              let id = iar.data.ID;
              checklistId = iar.data.ID;
              pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.add({
                Title: title,
                Subject: Subj,
                Department: dept,
                NoteType: Financial,
                Exceptional: Exceptional,
                Confidential: Confidential,
                SeqNo: this.state.seqno,
                PID: id,
                FY: FY,
                DeptAlias: deptAlias,
                CurApproverTxt: this.state.MgrName,
                ClientName: client,
                DeptGroupId: DeptGroupID,                
                RequesterId: requester,
                NoteFilename: filename,
                Sitename: 'Main',
                Status: "Draft",
                StatusNo: 12,
                Notefor : Notefor,
                Purpose : Purpose,
                ReturnName : ReturnName,
                DeptOwnership : DeptOwnership,
                RefferedGuidlines:RefferedGuidlines,
                VettingObservation:VettingObservation                              
              }).then((iar: ItemAddResult) =>{
                    for(var i=0;i<Checklisttable.length;i++)
                    {
                      pnp.sp.site.rootWeb.lists.getByTitle('Checklist').items.add({
                      Title: title,
                      SeqNo: this.state.seqno,
                      AppId: checklistId ,
                      Checklist:Checklisttable[i].checklist,
                      Status:Checklisttable[i].status
                      });
                    }                                
                  }).then(() => {
                    this.setCounter(uid, counter).then(() => {                      
                        this.AddWFHistory().then(() => {
                          this.redirect();
                        });
                      });                    
                  });
                });
            });                       
    }
  /*--End--*/

  /*-- Redirect to home Page--*/
  private gotoHomePage(): void {
    window.location.replace(this.props.siteUrl);
  }
  /*-- End Function--*/

  /*--Delete Attachment in NoteAnnexures library for Annexures attachment--*/
  public DeleteAttachment(vals : String): void {
    //debugger;
    this.setState({
      attachments: []
    });
    let sitename = this.state.Absoluteurl;
    let web = new Web('Main');
    let url = sitename + '/ChecklistAnnexures/' + vals;
    let fldr = vals.split("/")[0];
    let fldURL = sitename + '/ChecklistAnnexures/' + fldr;
    web.getFileByServerRelativeUrl(url).recycle().then(data => {
      console.log("File Deleted " + vals);
      web.getFolderByServerRelativeUrl(fldURL).files.get().then((result) => {
        let links: any[] = [];

        for (let i = 0; i < result.length; i++) {
          links.push(fldr + "/" + result[i].Name);

        }



        // console.log(links);
        this.setState({ attachments: links });
        // document.getElementById("fileUploadInput").nodeValue = null;        
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
      });

    });

  }
  /*--End--*/

  /*--Delete Attachment in NoteAttach library for Note attachment--*/
  public DeleteNote(vals:String): void {
    //debugger;
    this.setState({
      Note: []
    });
    let sitename = this.state.Absoluteurl;
    let url = sitename + '/ChecklistAttach/' + vals;
    let fldr = vals.split("/")[0];
    let fldURL = sitename + '/ChecklistAttach/' + fldr;
    let web = new Web('Main');
    web.getFileByServerRelativeUrl(url).recycle().then(data => {
      console.log("File Deleted " + vals);
      web.getFolderByServerRelativeUrl(fldURL).files.get().then((result) => {
        let links: any[] = [];

        for (let i = 0; i < result.length; i++) {
          links.push(fldr + "/" + result[i].Name);

        }
        this.setState({ Note: links });
        this.setState({ Notefilename: "" });
        // // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
        // document.getElementById("NoteDel").style.display = "none";
        let ddlDepartment2 = document.getElementById('ddlTemplate');
        if (ddlDepartment2) {
          ddlDepartment2.nodeValue = null;
          ddlDepartment2.style.display = "none";
        }
        jQuery('#NoteFile').text('');
      });

    });

  }
  /*--End--*/

  /*--Adding attachments in Document Library function--*/
  public AttachLib = (event : any) => {
    //debugger;
    var uploadFlag = true;
    //in case of multiple files,iterate or else upload the first file.
    let file = event.target.files[0];
    let filesize = file.size / 1048576;
    // let fileExtn1=file.name.split(".")[1].toLowerCase();
    var n = (file.name.length - file.name.lastIndexOf("."));
    //let fileExtn=file.name.substr(file.name.length-(n-1)).toLowerCase();
    let fileExtn = file.name.substr((file.name.lastIndexOf('.') + 1)).toLowerCase();
    let fileSplit = file.name.split(".");
    let fileType = this.state.AttachType;
    let PermissibleExtns = ['pdf'];
    let listName = 'ChecklistAttach';
    let NoteCount = this.state.Note.length;
    let notetype = this.state.NoteType;
    let TotalAnnexures = this.state.attachments.length;
    let fileTest = file.name.substring(0, (file.name.length - n));
    console.log(fileTest);
    let match = new RegExp('[~#%\&{}+.\|]|\\.\\.|^\\.|\\.$').test(fileTest);

    if (fileType != 'Note') {
      PermissibleExtns = ['png', 'jpeg', 'jpg', 'gif', 'pdf', 'doc', 'docx', 'xls', 'xlsx', '.eml'];
      listName = 'ChecklistAnnexures';
    }
    else {
      PermissibleExtns = ['pdf'];
    }


    if (fileSplit.length > 2) {
      alert('Alert-Selected file double extension is not allowed!');
      // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
      return false;
    }
    else if (match) {
      alert('Invalid file name. The name of the attached file contains invalid characters!');
      // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
      return false;
    }
    else if (file.name.split(".")[0].length > 150) {
      alert('Invalid file name. file names cannot be more than 150 chars!');
      // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
      return false;
    }
    else if (PermissibleExtns.indexOf(fileExtn.toLowerCase()) == -1) {
      alert('Alert-Selected file type is not allowed!');
      // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
      return false;
    }
    else if (filesize > 10) {
      alert('Alert-File size is more than permissible limit!');
      // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
      return false;
    }
    else if (fileType == 'Note' && NoteCount == 1) {
      alert('Alert-Only 1 Note is allowed!');
      // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
      return false;
    }
    else if (fileType != 'Note' && TotalAnnexures == 20) {
      alert('Alert-Only 20 Annexures can be uploaded!');
      // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
      return false;
    }
    else {
      this.on();
      if (file != undefined || file != null) {
        let SeqNo = this.state.seqno;
        let web = new Web('Main');
        web.getFolderByServerRelativePath(listName).folders.add(SeqNo).then(data => {
          console.log("Folder is created at " + data.data.ServerRelativeUrl);
          //assuming that the name of document library is Documents, change as per your requirement, 
          //this will add the file in root folder of the document library, if you have a folder named test, replace it as "/Documents/test"

          web.getFolderByServerRelativePath(listName + "/" + SeqNo).files.add(file.name, file, true).then((result) => {
            console.log(file.name + " uploaded successfully!");
            let links: any[] = [];

            if (fileType == 'Note') {
              this.setState({ Notefilename: file.name });
              links = this.state.Note;
              links.push(SeqNo + "/" + file.name);
              this.setState({ Note: links });
            } else {
              links = this.state.attachments;
              links.push(SeqNo + "/" + file.name);
              this.setState({ attachments: links });
            }
            console.log(links);

            // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
          });
        }).catch(data => {
          console.log(data);
          uploadFlag = false;

        });

      }
      else {
        uploadFlag = false;
      }

    }
    this.off();
    event.preventDefault();
    // return uploadFlag;
  }
  /*--End File Attach Function--*/

  /*--Is it a Confidential Note, Do you want to add Client and Is it Exceptional radio button change function */
  private Radibtnchangeevent(name : string, value : string) {//debugger;
    //debugger;

    if (name == "radioAttach") {
      this.setState({ RadioClient: value });
      if (value == 'CYes') {
        jQuery('#divClientName').show();
      }
      else {
        jQuery('#divClientName').hide();
      }
    }

    if (name == "radioConf") {
      if (value == 'ConfYes') {
        jQuery('#txtConfidential').val('Yes');
      }
      else {
        jQuery('#txtConfidential').val('No');
      }
    }

    if (name == "radioExc") {
      if (value == 'ExcYes') {
        jQuery('#txtExceptional').val('Yes');
      }
      else {
        jQuery('#txtExceptional').val('No');
      }
    }

  }
  /*--End--*/

  /*-- To get details from Restricted Emails master --*/
  private getRestrictedEmails() {
    //debugger;
    pnp.sp.site.rootWeb.lists.getByTitle('RestrictedEmails').items.select("ID,Title,AlertMessage").orderBy("ID asc").getAll().then((items: any[]) => {
      debugger;
      let links: string[]=[];
      let msg: string[]=[];
      for (let i = 0; i < items.length; i++) {
        //links += items[i].Title;
        //msg += items[i].AlertMessage;
        links.push(items[i].Title);
        msg.push(items[i].AlertMessage);
      }
      this.setState({ RestrictedEmails: links });
      this.setState({ RestrictedEmailsMsg: msg });
    });
  }
  /*--End All Functions--*/

  /*--KeyUp */
  private handleKeyUp(event : any) {
    let regx = /^[A-Za-z0-9 _.-]+$/;
    const keyValue = event.key;
    if (regx.test(keyValue))
    {
      //event.key;
      return true;
    }
    else
    {
      return false;
    }
  }
  /*--End--*/
}

