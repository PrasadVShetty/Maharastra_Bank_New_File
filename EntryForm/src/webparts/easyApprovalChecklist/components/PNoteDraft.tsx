import * as React from 'react';
import styles from './PaperlessApproval.module.scss';
import { IPaperlessApprovalProps } from './IPaperlessApprovalProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
//import { PeoplePicker } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext, PeoplePicker,PrincipalType} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { CIState } from "../Model/CIState";
import { default as pnp, ItemAddResult, File,sp,Web } from "sp-pnp-js";
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/components/Button';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
//import { CurrentUser } from '@pnp/sp/src/siteusers'; 
import { Button } from 'office-ui-fabric-react/lib/Button';
import { Attachments } from 'sp-pnp-js/lib/graph/attachments';
import * as jQuery from 'jquery';
import * as $ from "jquery";
import { SPComponentLoader } from '@microsoft/sp-loader';  
import { ListItemPicker } from '@pnp/spfx-controls-react/lib/listItemPicker';
SPComponentLoader.loadCss('../SiteAssets/css/styles.css');
require('../css/custom.css');
SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css');  
var RefioOutwardObj = [];
var RefioInwardObj = [];
const Delete: any = require('../images/Delete.png');
const Video: any = require('../images/Video.png');
const Logo:any=require('../images/Logo.png');
const Annex:any=require('../images/Upload-Annex.png');
const NoteAtt:any=require('../images/Upload-Note.png');

export default class PNoteForms extends React.Component<IPaperlessApprovalProps, CIState> {
  constructor(props : any) {
  super(props);
  this.handleTitle = this.handleTitle.bind(this);
  this.handleDesc = this.handleDesc.bind(this);
   this._onRenderFooterContent = this._onRenderFooterContent.bind(this);
  this.createItem = this.createItem.bind(this);
    this._getManager = this._getManager.bind(this);  
  this._getReceivedFrom = this._getReceivedFrom.bind(this); 
  this._getCCPeople = this._getCCPeople.bind(this); 
  this.DeleteApprover=this.DeleteApprover.bind(this);
   
  //  this.setButtonsEventHandlers();
     this.state = {
      name: "",
      description: "",
      selectedItems: [],
      hideDialog: true,
      showPanel: false,
      dpselectedItem: undefined,
      dpselectedItems: [],
      dropdownOptions:[],
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
      ManagerEmail:[],
      seqno:"",
      attachments:[],
      Note:[],
      AttachType:'',
      Appstatus:'',
      MgrName:'',
      files:null,
      UserID:0,
      UserEmail:'',
      iframeDialog:true,
      ImgUrl:'',
      CurrentItemId:0,
      RecpEmail:[],
      RecpID:[],
      RecpName:[],
      NoteType:'',
      Notefilename:'',
      Sitename:'',
      Absoluteurl:'',
      ccEmail:[],
      ccIDS:[],
      ccName:[],
      AppSeqNo:0,
      RecommSeqNo:0,
      ccSelectedItems: [],
      InwarddocketnoSet:'',
      Outwarddocketno:[],
      OutwarddocketnoSet:'',  
      RadioClient:'',
      controllerName:'',
      controllerPPId:0,
      RestrictedEmails:[],
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
      itemid:0 
    }];
  
    // Update state and reset fields
    this.setState({ items: newItems, checklist: '', status: '' });
  };  

  addItemChecklist =(uid : string) => {
    var newitems2 : any [] ;    
    debugger;
    pnp.sp.site.rootWeb.lists.getByTitle("Checklist").items.select("ID,Title,SeqNo,Checklist,AppId,Status").filter(`AppId eq ${uid}`).orderBy("ID", true).get().then((items: any[]) => {
      debugger;
      if(items.length > 0)
      {      
        newitems2 = items.map((item, index) => ({ id: index, checklist: item.Checklist, status: item.Status, itemid: item.ID}));        
        this.setState({ items: newitems2});
      }
    });
  }
  
  deleteItem = async (id: number, itemid: number) => {
    console.log(itemid);
    if(itemid > 0)
    {
      const recycleResult = await sp.web.lists .getByTitle("Checklist") .items.getById(itemid) .recycle();
      if (recycleResult) {
        console.log(`Successfully recycled item with ID: ${itemid}. Recycle Bin ID: ${recycleResult}`);
        const filteredItems = this.state.items.filter(item => item.id !== itemid);
        this.setState({ items: filteredItems });
      } else {
        console.error(`Recycling failed for item with ID: ${itemid}`);
      }
      const filteredItems = this.state.items.filter(item => item.id !== id);
      this.setState({ items: filteredItems });
    } 
    else
    {
      const filteredItems = this.state.items.filter(item => item.id !== id);
      this.setState({ items: filteredItems });
    }   
  };

  public render(): React.ReactElement<IPaperlessApprovalProps> {
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
      <div id="divHeadingNew" style={{display:"block",backgroundColor:"#0c78b8", textAlign:'center', color:'#fff'}}>
      <h3 className={styles.heading}>Note Form </h3> 

      </div>

      <div hidden id="divHeadingSubmit" style={{display:"none",backgroundColor:"#0c78b8", textAlign:'center', color:'#fff'}}>
      <h3  className={styles.heading}>Note Form </h3> 
      </div>

      </div>

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
                <select className='form-control form-control-sm' id="ddlDepartment" title="Select Department" placeholder="Select Department">
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
            onClick={() => this.deleteItem(item.id,item.itemid)}
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
                        ensureUser={true}
                        onChange={this._getManager}
                        showHiddenInUI={false}                        
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

                      <div className={styles.container} style={{marginTop:"5px"}}>
                     
        
          <div className={styles.overlay} id="overlay" style={{display:"none"}} >
              <span className={styles.overlayContent} style={{textAlign:"center"}}>Please Wait!!!</span>
       </div>
       <br></br>
       <div  className={styles.formrow+" "+"form-group row"} style={{backgroundColor:"cornsilk",borderRadius:"5px",margin:"0px"}}>
       <hr></hr>
            
            <div className="ms-Grid-col ms-u-sm3 block" id="btnCreate" style={{display:"block"}} > 
            <PrimaryButton style={{width:"25pt",borderRadius:"5%",backgroundColor:"#50B4E6"}} text="Submit" onClick={() => { this.validateForm(); }} /> </div>
           
           <div className="ms-Grid-col ms-u-sm3 block" id="btnDraft" style={{display:"block"}} >
            <PrimaryButton style={{ width: "25pt", borderRadius: "5%", backgroundColor: "#50B4E6", color:'#fff'}} text="Save Draft" onClick={() => { this.SaveDraftNew(); }} />
          </div>            
                   
            <div className="ms-Grid-col ms-u-sm3 block" id="btnCancel" style={{display:"block"}}>
              <PrimaryButton style={{width:"25pt",borderRadius:"5%",backgroundColor:"#50B4E6"}} text="Cancel" onClick={() => { this.cancel(); }} />
                      </div>
            <div className="ms-Grid-col ms-u-sm3 block" id="btnClose" style={{display:"none",width:"25pt",borderRadius:"50%"}}>
               <PrimaryButton style={{width:"25pt",borderRadius:"5%",backgroundColor:"#50B4E6"}} text="Close" onClick={() => { this.cancel(); }} />
            </div>
                 <hr></hr>
            </div>
           
            <br></br>
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
    </form>
    );
  }
/* -- Starting All Functions-- */

/*-- For Upload Attachment Popup--*/
  public UploadAttach(AttType:string){
    debugger;
    this.setState({AttachType:AttType});
    setTimeout(()=>{
      let fileUploadInput = document.getElementById('fileUploadInput');
      if(fileUploadInput)
      {fileUploadInput.click();}      
    },1500);
   
  }
  /*-- Set state on file change--*/
public  handleChange(files : any){
 this.setState({
   files: files
 });
}
/*-- End Function --*/

/*--For on(show) and off(hide) please wait overlay while page load--*/
private on() {
let ht=window.innerHeight;
  let overlay = document.getElementById('overlay');
  if(overlay){
    overlay.style.height=ht.toString()+"px";
    overlay.style.display = "block";
  }
  
}
private off() {
  let overlay = document.getElementById("overlay");
  if(overlay)
  {overlay.style.display = "none";}  
}
/*--End Function--*/

/*--Form On Load Function--*/
public  componentDidMount(){
    var reacthandler=this;
    var qstrnew : string[] = window.location.search.includes('pid=') ? window.location.search.split('pid=') : window.location.search.split('Pid=');
    // get Currnt User's details
    //pnp.sp.web.currentUser.get().then((r: CurrentUser) => {  //To get current user details from site 
    pnp.sp.web.currentUser.get().then((r) => {
    debugger;
    let sitename=r['odata.id'].split("/_api")[0];
    let absoluteurl=sitename.split("com")[1]+"/Main";
    this.setState({Absoluteurl:absoluteurl});
    this.setState({Sitename:sitename});
    const uname=r['UserPrincipalName'].split('@')[0];
    let username=r['Title'];
    let tdName = document.getElementById("tdName");
    if(tdName){tdName.innerText=username;}
    this.setState({name:username});
    this.setState({UserID: r['Id'] });
    let CurrUserEmail=r['LoginName'].split("|")[2];
    this.setState({UserEmail:CurrUserEmail});
    this.on();
    let qstr=window.location.search.includes('pid=') ? window.location.search.split('pid=') : window.location.search.split('Pid=');  
    let uid=0;
    if(qstr.length>1){uid= parseInt(qstr[1]);}
    this.setFields(uid);

    });
    /*-- for current date --*/
    let newDate = new Date();
    let date = newDate.getDate().toString();
    let month = (newDate.getMonth() + 1).toString();
    let year = newDate.getFullYear().toString();

    if(month.toString().length==1){month="0"+month.toString();}
    if(date.toString().length==1){date="0"+date.toString();}

    let fullDate=date+"-"+month+"-"+year;
    let tdDate = document.getElementById("tdDate");
    if(tdDate){tdDate.innerText=fullDate;}
    /*--End--*/

    /*-- To get details from masters(lists) --*/
    // this.setFin();
    this.getFinNotes();
    this.getDOP();
    this.getRestrictedEmails();
    this.fetchCounterData();
    this.getDepartments();
    this.addItemChecklist(qstrnew[1]);
    /*--End--*/
 
                     
 }

private SaveDraftNew(): void {
debugger;
this._onClosePanel();
this.on();
jQuery('#btnDraft').remove();    
jQuery('#Createbutton').remove();
jQuery('#Cancelbutton').remove();
let FY = jQuery('#tdFY').text();       
let Financial = jQuery('#ddlSource option:selected').text();
let FinType = jQuery('#ddlFinNote').val();       
let Amount = jQuery('#Amount').val();
let Exceptional = jQuery('#txtExceptional').val();
let Confidential = jQuery('#txtConfidential').val();            
let Notefor = jQuery('#txtNote').val();
let Purpose = jQuery('#txtPurpose').val();
let ReturnName = jQuery('#txtReturn').val();
let DeptOwnership = jQuery('#ddlDepartment option:selected').text();
let VettingObservation = jQuery('#txtVetting').val();
let RefferedGuidlines = jQuery('#txtGuidelines').val();       
let Checklisttable = this.state.items;
let qstr=window.location.search.includes('pid=') ? window.location.search.split('pid=') : window.location.search.split('Pid=');  
let uid=0;
if(qstr.length>1){uid= parseInt(qstr[1]);}
if (Financial == 'Financial') {
  Financial = String(FinType);
}
if (Amount == '') {
  Amount = 0;
}
let Recommenders = this.state.dpselectedItems.length;

let filename = this.state.Notefilename;           
let Subj = jQuery('#txtSubject').val();
let Comment = jQuery('#txtComments').val();      
let client = jQuery('#txtClient').val();      
let ControllerID = 0;
var checklistId : number ;         

if(this.state.ccIDS[0] != undefined){ControllerID = parseInt($('#lblController').text());}


this.setState({ attachments: [] });                               
let web = new Web('Main');
let Approvers : Number[] = [];            
console.log("SeqNo: "+this.state.seqno);
  web.lists.getByTitle('ChecklistNote').items.getById(uid).update({              
    SeqNo: this.state.seqno,
    Subject: Subj,              
    Comments: Comment,
    Exceptional: Exceptional,
    Confidential: Confidential,              
    ApproversId: { results: Approvers },              
    Amount: Amount,               
    NoteFilename: filename,
    NoteType: Financial,              
    ClientName: client,
    Migrate: "",
    FY: FY,              
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
    pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.getById(uid).update({                
    Subject: Subj,                 
    NoteType: Financial,
    Exceptional: Exceptional,
    Confidential: Confidential,
    SeqNo: this.state.seqno,
    PID: id,
    FY: FY,                 
    CurApproverTxt: this.state.MgrName,
    ClientName: client,                                                  
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
        if(Checklisttable[i].itemid == 0)
        {
          pnp.sp.site.rootWeb.lists.getByTitle('Checklist').items.add({
          Title: this.state.Title,
          SeqNo: this.state.seqno,
          AppId: checklistId ,
          Checklist:Checklisttable[i].checklist,
          Status:Checklisttable[i].status
          });
        }
      }                                
      }).then(() => {                                    
      this.AddWFHistory().then(() => {
      this.redirect();
      });
    });                                       
  });                                    
}

 /*--To get saved data from Notes list and update to current form fields--*/
 private setFields(uid:number){
  debugger;
  let web=new Web('Main');  
  let fldr='';
  web.lists.getByTitle('ChecklistNote').items.select("ID,Title,Department,Created,SeqNo,Status,StatusNo,Comments,Subject,NoteType,NoteFilename,DeptAlias,Amount,ClientName,Confidential,Requester/Title,Requester/EMail,Requester/Name,Requester/ID,CurApprover/EMail,CurApprover/Title,CurApprover/ID,CurApprover/Name,ReturnedBy/ID,ReturnedBy/Title,ReferredBy/ID,ReferredBy/Title,ReferredTo/ID,ReferredTo/Title,Controller/ID,Controller/EMail,Controller/Title,Approvers/ID,DOP,WorkflowFlag,Modified,RefCount,Notefor,Purpose,ReturnName,DeptOwnership,DueDate,Place,RefferedGuidlines,VettingObservation").expand('Requester,CurApprover,ReturnedBy,Controller,ReferredBy,ReferredTo,Approvers').filter('ID eq '+uid).orderBy("ID asc").getAll().then((items: any[]) => {
           if(items[0].SeqNo!=null){
        this.setState({seqno:items[0].SeqNo});
       }
   
   this.setState({NoteType:items[0].NoteType});
   this.setState({Notefilename:items[0].NoteFilename});
   if(items[0].Title!=null){
   this.setState({Appstatus:items[0].Title});
   }
    
    $("#txtSubject").val(items[0].Subject);
    $('#divDepartment').val(items[0].Department);
    $('#ddlDepartment').val(items[0].Department);
    //$('#divTitle').text(items[0].Title);
   this.setState({Title:items[0].Title});
   //$("#ddlDepartment option:contains(" + items[0].Department + ")").attr('selected', 'selected');
   $("#ddlDOP option:contains(" + items[0].DOP + ")").attr('selected', 'selected');
   let NoteType=items[0].NoteType;
   if(NoteType=='Non-Financial'){
    $("#ddlSource option:contains(" + NoteType + ")").attr('selected', 'selected');
 
   }else{
    $("#ddlSource option:contains(Financial)").first().attr('selected', 'selected');
    $('.FinancialClass').css('display','block');
    $("#ddlFinNote option:contains(" + NoteType + ")").attr('selected', 'selected');
     $("#Amount").val(items[0].Amount);
   }
   let Confidential=items[0].Confidential;
   let client=items[0].ClientName;
   if(items[0].deptAlias=='HRD'){
    let divConfidential = document.getElementById('divConfidential')
    if(divConfidential){divConfidential.style.display='block';}
    
   }


   if(client!=null){
      //$('#CYes').attr('checked',true);
      $('#CYes').prop('checked',true);      
      $('#divClientName').css('display','block');
      $('#txtClient').val(client);
      this.setState({RadioClient:'CYes'});
   }
   else{
    this.setState({RadioClient:'CNo'});
    //$('#CNo').attr('checked',true);
    $('#CNo').prop('checked',true);
   }
   $('#txtConfidential').val(Confidential);
   if(Confidential=='Yes'){
    //$('#ConfYes').attr('checked',true);
    $('#ConfYes').prop('checked',true);
       }
   else{
    //$('#ConfNo').attr('checked',true);
    $('#ConfNo').prop('checked',true);
   }
   let curapprover=0;
  if(items[0].CurApprover!=undefined){
    curapprover=items[0].CurApprover.ID;
  }
 
  // if(this.state.UserID!=curapprover && curapprover>0){
  //  let btnCreate = document.getElementById("btnCreate");
  //  if(btnCreate){btnCreate.style.display='none';}
  //  // document.getElementById("btnDraft").style.display='none';     
  // }

  let requesterid = items[0].Requester.ID;
  if(this.state.UserID!=requesterid && requesterid>0){
     let btnCreate = document.getElementById("btnCreate");
     if(btnCreate){btnCreate.style.display='none';}     
  }
  

   //added on 16/02/2025
   let tdPurpose = document.getElementById("txtPurpose") as HTMLInputElement;;
   if (tdPurpose) tdPurpose.value = items[0].Purpose;
   let tdNotefor = document.getElementById("txtNote") as HTMLInputElement;;
   if (tdNotefor) tdNotefor.value = items[0].Notefor;
   let tdReturnName = document.getElementById("txtReturn") as HTMLInputElement;;
   if (tdReturnName) tdReturnName.value = items[0].ReturnName;
   let tdDeptOwnership = document.getElementById("ddlDepartment") as HTMLInputElement;;
   if (tdDeptOwnership) tdDeptOwnership.value = items[0].DeptOwnership;
   let tdreffered = document.getElementById("txtGuidelines") as HTMLInputElement;;
   if (tdreffered) tdreffered.value = items[0].RefferedGuidlines;
   let tdvetting = document.getElementById("txtVetting") as HTMLInputElement;;
   if (tdvetting) tdvetting.value = items[0].VettingObservation;

  
jQuery('#txtComments').val(items[0].Comments);

// Retrieve All masters and Child lists Records
this.retrieveRecommenders();
this.retrieveController();
this.retrieveApprovers();
this.getMainNote();    
this.getAnnexures();   
this.off();

  });

 }
 /*--End--*/
 /*--get attachments for Notes--*/
private getMainNote(){
  let web=new Web('Main');  
  let fldURL='ChecklistAttach/'+this.state.seqno;
  web.getFolderByServerRelativeUrl(fldURL).files.get().then((result) => {
    let links:any[]=[];
 
    for(let i=0;i<result.length;i++){
      links.push(this.state.seqno+"/"+result[i].Name);

    }
    this.setState({ Note: links});
  });
   
}
/*--End--*/
/*--get attachments for Annexures--*/
private getAnnexures(){
  let web=new Web('Main');  
  let fldURL='ChecklistAnnexures/'+this.state.seqno;
  web.getFolderByServerRelativeUrl(fldURL).files.get().then((result) => {
    let links:any[]=[];
 
    for(let i=0;i<result.length;i++){
      links.push(this.state.seqno+"/"+result[i].Name);

    }
    
    this.setState({ attachments: links});
   
});
}
/*--End--*/

/*-- To get details from Departments master for Department dropdown --*/
  private getDepartments(){
    debugger;
    pnp.sp.site.rootWeb.lists.getByTitle('Departments').items.select("ID,Title,Dept_Alias").orderBy("ID asc").getAll().then((items: any[]) => {
      debugger;
      // console.log(items);
      let links:string='';
               for(let i=0;i<items.length;i++){
              links+= "<option value='" + items[i].Title + "'>" + items[i].Title + "</option>";
              }
          jQuery('select[id="ddlDepartment"]').append(links);
 
  });
  }
  /*--End--*/
  /*-- To get details from FYMaster master --*/
  private getFY(){
    debugger;
    pnp.sp.site.rootWeb.lists.getByTitle('FYMaster').items.select("ID,Title,Active").orderBy("ID asc").getAll().then((items: any[]) => {
      debugger;
      // console.log(items);
      let links:string='';
               for(let i=0;i<items.length;i++){
                 if(items[i].Active=='Yes'){
                  jQuery('#tdFY').text(items[i].Title);
                 }
              
              }
          
 
  });
  }
/*--End--*/ 
  /*-- To get details from FinNotes master for Type of Financial Note --*/
  private getFinNotes(){
    debugger;
    pnp.sp.site.rootWeb.lists.getByTitle('FinNotes').items.select("ID,Title").orderBy("ID asc").getAll().then((items: any[]) => {
      debugger;
       // console.log(items);
      let links:string='';
         for(let i=0;i<items.length;i++){
                    links+= "<option value='" + items[i].Title + "'>" + items[i].Title + "</option>";
               }
     jQuery('select[id="ddlFinNote"]').append(links);
  
  });
  }
  /*--End--*/ 
  /*-- To set Note Type dropdown --*/
 private setFin(){
    let links:string='';
    links+= "<option value='Financial'>Financial</option>";
    links+= "<option value='Non-Financial'>Non-Financial</option>";
    jQuery('select[id="ddlSource"]').append(links);
  }
   /*--End--*/
  /*-- To get details from DOP master for DOP Details --*/
  private getDOP(){
    debugger;
    pnp.sp.site.rootWeb.lists.getByTitle('DOP').items.select("ID,Title").orderBy("ID asc").getAll().then((items: any[]) => {
      debugger;
      let links:string='';
         for(let i=0;i<items.length;i++){
                    links+= "<option value='" + items[i].Title + "'>" + items[i].Title + "</option>";
               }
     jQuery('select[id="ddlDOP"]').append(links);
  
  });
  }
    /*--End--*/
  /*-- To Update Recommanders in Approvals list--*/
  private AddRecommender(){
    debugger;
    let seqno= this.state.RecommSeqNo+1;
    let MgrID=this.state.RecpID;
    let userid=this.state.UserID;
    let TotalRecomm=this.state.dpselectedItems;
    let restricedEmails=this.state.RestrictedEmails;
    let restricedEmailsMsg=this.state.RestrictedEmailsMsg;
    if(this.state.RecpName[0]==''){
        alert('Kindly select username!');
            $('#RecommenderPPtd >div>div>div>div>div>div>div>input').trigger("focus");
        return;
    }
    else if(TotalRecomm.length==10){
      alert('Only 10 Recommenders can be added!');
      $('#RecommenderPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
      $('#RecommenderPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    else if(restricedEmails.indexOf(this.state.RecpEmail[0].toLowerCase())>=0){
      let indx=restricedEmails.indexOf(this.state.RecpEmail[0].toLowerCase());
      let msg = restricedEmailsMsg[indx];
      alert(msg);
      //alert(this.state.RecpEmail[0] +' cannot be added, kindly select proper name id');
      $('#RecommenderPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
      $('#RecommenderPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    else if(userid==MgrID[0]){
      alert('Requester cannot be recommender!');
      $('#RecommenderPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
      $('#RecommenderPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    else{
         let mgrEmail=this.state.RecpEmail[0];
                   this.checkRecommender(mgrEmail).then((len)=>{
                    if(len==0 ){   
                       this.checkApprover(mgrEmail).then((len1)=>{
                    if(len1==0 ){   
                    let SeqNo=this.state.seqno;
                          debugger;
                  let web=new Web('Main');
                 
                  web.lists.getByTitle('ApprovalsChecklist').items.add({
                         Title:this.state.seqno,
                         Status:'Pending',
                            Seq:seqno,
                           ApproverId: this.state.RecpID[0],
                           AppID:this.state.RecpID[0],
                         AppName:this.state.RecpName[0],
                         AppEmail:this.state.RecpEmail[0]             
                     }).then((iar:ItemAddResult) => {
                       this.setState({ RecommSeqNo:seqno});
                       console.log(iar.data.ID);
                       $("#RecommenderPPtd .ms-PickerItem-removeButton").trigger("click");
                         this.retrieveRecommenders();
                       $('#RecommenderPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
                                        });
                   }
                  else{
                    alert('Approver cannot be Recommender!');
                    $('#RecommenderPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
      
                    return;
                  }
                 });
                    }
                  else{
                   alert('Recommender has already been added!');
               $('#RecommenderPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
          
      
                return;
           
           
                    }
           
               });
      
  }
  }
  /*--End--*/

  /*-- To Update Approvers in FApprovals list--*/
  private AddApprover(){
    debugger;
    let seqno= this.state.AppSeqNo+1;
    let MgrID=this.state.userManagerIDs;
    let userid=this.state.UserID;
    let TotalApp=this.state.selectedItems;
    let controllerflag="";
    let restricedEmails=this.state.RestrictedEmails;
    let restricedEmailsMsg=this.state.RestrictedEmailsMsg;
    if(jQuery('#ddlDepartment option:selected').val()=="TIG")  {
   controllerflag = "Yes";
    }
    if(this.state.MgrName==''){
        alert('Kindly select username!');
        
      $('#ApproverPPtd >div>div>div>div>div>div>div>input').trigger("focus");
        return;
    }
    else if(restricedEmails.indexOf(this.state.ManagerEmail[0].toLowerCase())>=0){
      let indx=restricedEmails.indexOf(this.state.ManagerEmail[0].toLowerCase());
      let msg = restricedEmailsMsg[indx];
      alert(msg);
      //alert(this.state.ManagerEmail[0] +' cannot be added, kindly select proper name id');
          $('#ApproverPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
     $('#ApproverPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    else if(TotalApp.length==10){
      alert('Only 10 Approvers can be added!');
        $('#ApproverPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
      $('#ApproverPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    else if(userid==MgrID[0] && controllerflag != 'Yes'){
      alert('Requester cannot be approver!');
      $('#ApproverPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
      $('#ApproverPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    else{
      
      let mgrEmail=this.state.ManagerEmail[0];
     console.log(this.state.userManagerIDs[0]);
          console.log(this.state.MgrName);
    
         this.checkApprover(mgrEmail).then((len)=>{
          if(len==0  ){
            this.checkRecommender(mgrEmail).then((len1)=>{
         
              if(len1==0 ){   
         let SeqNo=this.state.seqno;
         let web=new Web('Main');
            debugger;
            web.lists.getByTitle('FApprovalsChecklist').items.add({
              Title:this.state.seqno,
              Status:'Pending',
                 Seq:seqno,
                ApproverId: this.state.userManagerIDs[0],
                AppID:this.state.userManagerIDs[0],
              AppName:this.state.MgrName,
              AppEmail:this.state.ManagerEmail[0]             
          }).then((iar: ItemAddResult) => {
            this.setState({ AppSeqNo:seqno});
            console.log(iar.data.ID);
               this.retrieveApprovers();
               $("#ApproverPPtd .ms-PickerItem-removeButton").trigger("click");
            $('#ApproverPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
                    });
        }
        else{
          alert('Recommender cannot be Approver!');
          $('#ApproverPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
             return;
        }
      });
  
          }
         
          else{
            alert('Approver has already been added!');
            $('#ApproverPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
             return;
  
  
          }
  
      });
    
  }
  }
  
  /*--End--*/
  // Add Controller before submission
  /*-- To Update Controller in CApprovals list--*/
  private AddController(){
    debugger;
    let seqno= 1;
    let MgrID=this.state.ccIDS;
    let userid=this.state.UserID;
    let Controllers=this.state.ccSelectedItems;
    let restricedEmails=this.state.RestrictedEmails;
    let restricedEmailsMsg=this.state.RestrictedEmailsMsg;
  
    if(this.state.ccName[0]==''){
        alert('Kindly select username!');
        //jQuery('input[aria-label="People Picker"]').focus();
        $('#ControllerPPtd >div>div>div>div>div>div>div>input').trigger("focus");
        return;
    }
    else if(Controllers.length>0){
      alert('Only 1 Controller can be added!');
      $('#ControllerPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
      $('#ControllerPPtd >div>div>div>div>div>div>div>input').trigger("focus");
  //      $('#selected-items-id__59 > div>button>div>i').click();
      return;
    }
    else if(restricedEmails.indexOf(this.state.ccEmail[0].toLowerCase() )>=0){
      let indx=restricedEmails.indexOf(this.state.ccEmail[0].toLowerCase());
      let msg = restricedEmailsMsg[indx];
      alert(msg);
      //alert(this.state.ccEmail[0] +' cannot be added, kindly select proper name id');
      $('#ControllerPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    else if(userid==MgrID[0]){
      alert('Requester cannot be Controller!');
      $('#ControllerPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
      $('#ControllerPPtd >div>div>div>div>div>div>div>input').trigger("focus");
      return;
    }
    else{
     
      let mgrEmail=this.state.ccEmail[0];
     // console.log(this.state.userManagerIDs[0]);
        //  console.log(this.state.MgrName);
        this.setState({ controllerName:this.state.ccName[0] });
        this.setState({ controllerPPId:this.state.ccIDS[0] });
         this.checkApprover(mgrEmail).then((len)=>{
          if(len==0  ){
            this.checkRecommender(mgrEmail).then((len1)=>{
         
              if(len1==0 ){   
         let SeqNo=this.state.seqno;
         let web=new Web('Main');
            debugger;
            web.lists.getByTitle('CApprovalsChecklist').items.add({
            Title:this.state.seqno,
            Status:'Pending',
            Seq:seqno,
            // LikedById: {results:[this.state.userManagerIDs[0]]},
            // Views: 1,
            ApproverId: this.state.ccIDS[0],
            AppID:this.state.ccIDS[0],
            AppName:this.state.ccName[0],
            AppEmail:this.state.ccEmail[0]             
          }).then((iar: ItemAddResult) => {
            this.setState({ AppSeqNo:seqno});
            console.log(iar.data.ID);
          //  jQuery('i[data-icon-name="Cancel"]').click();
            this.retrieveController();
            $('#ControllerPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
            //$('#selected-items-id__59 > div>button>div>i').click();
          });
        }
        else{
              alert('Recommender cannot be Controller!');
              $('#ControllerPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
             return;
        }
      });
  
          }
         
          else{
            alert('Approver has already been added!');
            $('#ControllerPPtd >div>div>div>div>div>div>div>span>div>button>div>i').trigger("click");
             return;
  
  
          }
  
      });
    
  }
  // $('i[data-icon-name="Cancel"]').click();
  }
  /*--End--*/
/*-- To Check adding approver present in FApprovals list or not--*/
  private checkApprover(appemail:string):Promise<number>{
    debugger;
     let title=this.state.seqno;
    let len=0;
    let web=new Web('Main');
      return web.lists.getByTitle('FApprovalsChecklist').items.select("ID,Title,AppName,AppEmail").filter("Title eq '"+title+"'").orderBy("Seq asc").getAll().then((items: any[]) => {
       
    for(let i=0;i<items.length;i++){
      if(items[i].AppEmail==appemail){
        len=1;
      }
    }
   
      return Promise.resolve(len); 
    });
   
    }
/*--End--*/
/*-- To Check adding recommender present in Approvals list or not--*/
    private checkRecommender(appemail:string):Promise<number>{
      debugger;
       let title=this.state.seqno;
      let len=0;
      let web=new Web('Main');
      return web.lists.getByTitle('ApprovalsChecklist').items.select("ID,Title,AppName,AppEmail").filter("Title eq '"+title+"'").orderBy("Seq asc").getAll().then((items: any[]) => {
         
      for(let i=0;i<items.length;i++){
        if(items[i].AppEmail==appemail){
          len=1;
        }
      }
     
        return Promise.resolve(len); 
      });
     
      }

   

  /*--End--*/
/*-- To retrieve approvers from FApprovals List--*/   
  private retrieveApprovers(){
    let title=this.state.seqno;
    // let data=[];
    let data: any[] = [];
    let web=new Web('Main');
     web.lists.getByTitle('FApprovalsChecklist').items.select("ID,Title,AppName").filter("Title eq '"+title+"' ").orderBy("Seq asc").getAll().then((items: any[]) => {
      debugger;
      if(items.length>0){
        for(let i=0;i<items.length;i++){
           data.push(<tr><td>{i+1}</td><td>{items[i].AppName}</td><td><button className='btn' onClick={()=>{this.DeleteApprover(items[i].ID);}}>Delete</button></td></tr>);
       }
      }

    }).then(()=> {
      this.setState({selectedItems:data});
    });

  }
/*--End--*/
/*-- To retrieve recommanders from Approvals List--*/ 
  private retrieveRecommenders(){
    let title=this.state.seqno;
    // let data=[];
    let data: any[] = [];
    let web=new Web('Main');
    web.lists.getByTitle('ApprovalsChecklist').items.select("ID,Title,AppName").filter("Title eq '"+title+"' ").orderBy("Seq asc").getAll().then((items: any[]) => {
      debugger;
      if(items.length>0){
        for(let i=0;i<items.length;i++){
           data.push(<tr><td>{i+1}</td><td>{items[i].AppName}</td><td><button className='btn' onClick={()=>{this.DeleteRecommender(items[i].ID);}}>Delete</button></td></tr>);
       }
      }

    }).then(()=> {
      this.setState({dpselectedItems:data});
    });

  }
/*--End--*/
/*-- To retrieve controller from CApprovals List--*/
  private retrieveController(){
    let title=this.state.seqno;
    // let data=[];
    let data: any[] = [];
    let ControllerID=this.state.ccIDS;
    let web=new Web('Main');
     web.lists.getByTitle('CApprovalsChecklist').items.select("ID,Title,AppName").filter("Title eq '"+title+"'").orderBy("Seq asc").getAll().then((items: any[]) => {
      debugger;
      if(items.length>0){
        for(let i=0;i<items.length;i++){
           data.push(<tr><td>{i+1}</td><td>{items[i].AppName}</td><td><button className='btn' onClick={()=>{this.DeleteController(items[i].ID);}}>Delete</button></td></tr>);
       }
      }

    }).then(()=> {
      this.setState({ccSelectedItems:data});
    });

  }
/*--End--*/
/*-- To Delete approvers in FApprovals List--*/
  public DeleteApprover(uid: number, event?: React.MouseEvent<HTMLButtonElement>):void{
    debugger;
    event?.preventDefault();
    let web=new Web('Main');
     
    let list =web.lists.getByTitle('FApprovalsChecklist');
    list.items.getById(uid).delete().then(() => {console.log('List Item Deleted');
    this.retrieveApprovers();
  });

  }
   /*--End--*/
/*-- To Delete controller in CApprovals List--*/
  public DeleteController(uid: number, event?: React.MouseEvent<HTMLButtonElement>):void{
    debugger;
    event?.preventDefault();
    let web=new Web('Main');
     
    let list =web.lists.getByTitle('CApprovalsChecklist');
    list.items.getById(uid).delete().then(() => {console.log('List Item Deleted');
    this.retrieveController();
    this.setState({ccSelectedItems:[]});
   
  });

  }
    /*--End--*/
/*-- To Delete recommender in Approvals List--*/
  public DeleteRecommender(uid: number, event?: React.MouseEvent<HTMLButtonElement>):void{
    debugger;
    event?.preventDefault();
    let web=new Web('Main');
     
    let list =web.lists.getByTitle('ApprovalsChecklist');
    list.items.getById(uid).delete().then(() => {console.log('List Item Deleted');
    this.retrieveRecommenders();
  });

  }
  /*--End--*/
/*-- To get first approver in FApprovals List(to set current approver while submit)--*/

  private retrieveFirstApprover():Promise<any[]>{
    let title=this.state.seqno;
    // let approverID=[];
    let approverID: any[] = [];
    let web=new Web('Main');
      return web.lists.getByTitle('FApprovalsChecklist').items.select("ID,Title,AppName,Approver/ID,Approver/Title").filter("Title eq '"+title+"'").expand("Approver").orderBy("ID asc").getAll().then((items: any[]) => {
      debugger;
      this.setState({MgrName:items[0].Approver.Title});
      approverID[0]=items[0].Approver.ID;
      approverID[1]=items[0].ID;
      return approverID;
     
    }); 
  //return data;
  }
  /*--End--*/
/*-- To update  first approver in FApprovals List--*/
  private updateFirstApprover(uid:number):Promise<any[]>{
    let web=new Web('Main');
        return web.lists.getByTitle('FApprovalsChecklist').items.getById(uid).update({
       Status: 'Submitted'
     }).then(() => {
       console.log('Approver updated');
         return Promise.resolve(['Done']); 
       
   });
   
    }
/*--End--*/
/*-- To get first recommander in Approvals List(to set current approver while submit)--*/
    private retrieveFirstRecommender():Promise<any[]>{
      let title=this.state.seqno;
      // let approverID=[];
      let approverID: any[] = [];
      let web=new Web('Main');
      return web.lists.getByTitle('ApprovalsChecklist').items.select("ID,Title,AppName,Approver/ID,Approver/Title").filter("Title eq '"+title+"'").expand("Approver").orderBy("ID asc").getAll().then((items: any[]) => {
        debugger;
        this.setState({MgrName:items[0].Approver.Title});
        approverID[0]=items[0].Approver.ID;
        approverID[1]=items[0].ID;
        return approverID;
       
      });
     
     
  //    return data;
    }
     /*--End--*/
/*-- To update  first recommander in Approvals List--*/
    private updateFirstRecommender(uid:number):Promise<any[]>{
      let web=new Web('Main');
      return web.lists.getByTitle('ApprovalsChecklist').items.getById(uid).update({
         Status: 'Submitted'
       }).then(() => {
         console.log('Approver updated');
           return Promise.resolve(['Done']); 
         
     });
     
      }
/*--End--*/
/*-- To add work flow history in WFHistory list--*/
    private AddWFHistory():Promise<any[]>{
      let dt=new Date();
      let mnth=(dt.getMonth()+1).toString();
      let dat=dt.getDate().toString();
      let hrs=dt.getHours().toString();
      let mins=dt.getMinutes().toString();
      let secs=dt.getSeconds().toString();
      if(mnth.length==1 ){mnth='0'+mnth;} if(dat.length==1 ){dat='0'+dat;}if(hrs.length==1 ){hrs='0'+hrs;}if(mins.length==1 ){mins='0'+mins;}if(secs.length==1 ){secs='0'+secs;}
      let createDate=dat+"-"+mnth+"-"+dt.getFullYear()+" "+hrs+":"+mins+":"+secs;
      let log='Submitted to '+this.state.MgrName+' by '+this.state.name+' on '+createDate;
      debugger;
      let web=new Web('Main');
      return web.lists.getByTitle('WFHistory').items.add({
        Title:this.state.seqno,
        AuditLog:log,
        Currapprover:this.state.MgrName,
        FormName:'Note',
        ActionDateTime:createDate       
    }).then((iar: ItemAddResult) => {
      console.log('History Log Created!');
      return Promise.resolve(['Done']);
     
    });
  
    }
 /*--End--*/
 /*-- Note Type change function--*/   
private SelectSource(){
  let source=jQuery('#ddlSource option:selected').val();
  if(source=='Financial'){
  jQuery('.FinancialClass').css('display','block');
    }
  else{
    jQuery('.FinancialClass').css('display','none');
  }

}
/*--End--*/
 /*-- To save name,email and id for controller people picker--*/
private _getCCPeople(items: any[]) {debugger;
  this.state.ccIDS.length = 0;
  let Recpid = [];
  let Recpname=[];
  let Recpemail=[];
 
  for (let item in items) {
    Recpid.push(items[item].id);
    Recpname.push(items[item].text);
    Recpemail.push(items[item].loginName.split("|")[2]);
    // alert(items[item].id);
  }
  this.setState({ ccName:Recpname });
  this.setState({ccIDS:Recpid });
   this.setState({ccEmail:Recpemail});
   $('#lblController').text(Recpid[0]);
   setTimeout(()=>{
    if(this.state.ccIDS.length==1)
    {    this.AddController();}
       
  },1000);
}
/*--End--*/
/*-- To save name,email and id for recommander people picker--*/
 private _getReceivedFrom(items: any[]) {debugger;
   this.state.RecpID.length = 0;
  let Recpid = [];
  let Recpname=[];
  let Recpemail=[];
  if(items.length>0){
    this.setState({isChecked:true});
    for (let item in items) {
      Recpid.push(items[item].id);
      Recpname.push(items[item].text);
      Recpemail.push(items[item].loginName.split("|")[2]);
      // alert(items[item].id);
    }
     
    this.setState({RecpID:Recpid });
    this.setState({ RecpName:Recpname });
    this.setState({RecpEmail:Recpemail});
    setTimeout(() => {
      if(items.length>0){
      this.AddRecommender();}
    }, 1000);
} // Ending If of items.length
  
}
/*--End--*/
/*-- To save name,email and id for approver people picker--*/
 private _getManager(items: any[]) {
   debugger;
   this.state.userManagerIDs.length = 0;
   let tempuserMngArr = [];
   let MgrEmail=[];
   let MgrName='';
   for (let item in items) {
     tempuserMngArr.push(items[item].id);
     MgrName=items[item].text;
     MgrEmail.push(items[item].loginName.split("|")[2]);
     // alert(items[item].id);
   }
   this.setState({ userManagerIDs: tempuserMngArr });
   this.setState({ManagerEmail:MgrEmail});
   this.setState({MgrName:MgrName});
   setTimeout(() => {
    if(items.length>0){
    this.AddApprover();}
  }, 1000);
 }
/*--End--*/
    
 /*-- Panel on Submission --*/    
private _onRenderFooterContent = (): JSX.Element => {
   return (
     <div>
       <PrimaryButton id="Createbutton" onClick={this.createItem} style={{ marginRight: '5px', width:"25pt"}}>
         Confirm
       </PrimaryButton>
       < PrimaryButton id="Cancelbutton" style={{ marginLeft: '5px', width:"25pt"}} onClick={this._onClosePanel}>Cancel</PrimaryButton>
     </div>
   );
 }
 /*-- End Function --*/

 /*-- cancel button logic --*/
 private cancel = () => {
   this.setState({ showPanel: false });
    // self.close();
    const query = window.location.search.split('Pid=')[1];
    let uid=0;
    if( query!=undefined){uid=parseInt(query); }
    if(uid==0){

    window.location.replace(this.props.siteUrl);
    }
    else{
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
 private redirect(){
   let sitename=this.state.Sitename;
  const query = window.location.search.split('Pid=')[1];
  let uid=0;
  if( query!=undefined){uid=parseInt(query); }
  if(uid==0){
  window.location.replace(sitename);
  }
  else{
    setTimeout(() => {
      window.location.replace(sitename);
     // self.close();
     }, 3000);
  }
 }
/*-- End --*/
 
/*--Show Panel function--*/
private _onShowPanel = () => {
   this.setState({ showPanel: true });
 }

 /*--Set Name state function--*/
 private handleTitle(value: string): void {
   return this.setState({
     name: value
   });
 }
/*--End function--*/

 /*--Set Description state function--*/
 private handleDesc(value: string): void {
   return this.setState({
     description: value
   });
 }
/*--End function--*/
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
  const ddlDeptOwnership = getValue('#ddlDepartment');
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
  if (ddlDeptOwnership === 'Select') { showAlert('Kindly select the Department Ownership!', '#ddlDepartment'); return; }
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

 fetchCounterData = async () => {
     const num = await this.getCounter();
     if(num[3] == 'Compliance')
     {this.setState({ vettingobservation: 'Vetting Observation' });
     this.setState({Checklistlabel: 'Checklist'});}
     else{this.setState({ vettingobservation: 'Remarks' });
     this.setState({Checklistlabel: 'Annexure'});}    
   };
 
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

/*--Record Submit function to update lists--*/
private createItem(): void {
  debugger;
  this._onClosePanel();
  this.on();
  jQuery('#Createbutton').remove();
  jQuery('#Cancelbutton').remove();
  //let title=jQuery('#divTitle').text();
  let dept=jQuery('#divDepartment').text();
  let qstr=window.location.search.includes('pid=') ? window.location.search.split('pid=') : window.location.search.split('Pid=');  
  let uid=0;
  if(qstr.length>1){uid= parseInt(qstr[1]);}
  let Financial=jQuery('#ddlSource option:selected').text();
  let FinType=jQuery('#ddlFinNote').val();
  let DOP=jQuery('#ddlDOP').val();
  let Amount=jQuery('#Amount').val();
  let Confidential=jQuery('#txtConfidential').val();
  if(Financial=='Financial'){
    Financial=String(FinType);
  }
  if(Amount==''){
    Amount=0;
  }
  let Recommenders=this.state.dpselectedItems.length;
 
  let filename=this.state.Notefilename;
  
  //added on 16/02/2025
  let Notefor = jQuery('#txtNote').val();
  let Purpose = jQuery('#txtPurpose').val();
  let ReturnName = jQuery('#txtReturn').val();
  let DeptOwnership = jQuery('#ddlDepartment option:selected').text();
  let VettingObservation = jQuery('#txtVetting').val();
  let RefferedGuidlines = jQuery('#txtGuidelines').val();
  let Checklisttable = this.state.items;
  var checklistId : number ;
  //

    
  let Subj=jQuery('#txtSubject').val();
  let Comment=jQuery('#txtComments').val();
  let client=jQuery('#txtClient').val();
  let requester=this.state.UserID;     
  // let ControllerID=$('#lblController').text();
  // if(ControllerID==''){ControllerID=String(this.state.ccIDS[0]);}
  let ControllerID = 0;
  var checklistId : number ;        
  if(this.state.ccIDS[0] != undefined){ControllerID = parseInt($('#lblController').text());}

  if(Recommenders>0){
    this.retrieveFirstRecommender().then((Appid)=>{
      var approverID: Number[] = [];
      approverID=Appid[0];
      var AppItemid: Number[] = [];
      AppItemid=Appid[1];      
      let web=new Web('Main');    
      let Approvers: any[] = [];
       Approvers.push(approverID);        
      web.lists.getByTitle('ChecklistNote').items.getById(uid).update({
        Subject:Subj,
        Comments:Comment,
        Confidential:Confidential,
        CurApproverId:approverID,
        NotifyId:approverID,
        ApproversId:{results:Approvers},
        Amount:Amount,
        NoteFilename:filename,
        NoteType:Financial,
        DOP:DOP,
        ClientName:client,
        Migrate:"",
        ControllerId:ControllerID,
        Status:"Submitted to Recommender#1",
        StatusNo:1,
        Notefor : Notefor,
        Purpose : Purpose,
        ReturnName : ReturnName,
        DeptOwnership : DeptOwnership,
        RefferedGuidlines:RefferedGuidlines,
        VettingObservation:VettingObservation
          }).then((iar: ItemAddResult) => {
            console.log(iar.data.ID);
            let id=iar.data.ID;
            pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.select('Title','ID','PID').filter("PID eq "+uid ).get().then(r=>{
              let Approverid=r[0].ID;
              pnp.sp.site.rootWeb.lists.getByTitle('ChecklistNote').items.getById(Approverid).update({
              Subject:Subj,
              NoteType:Financial,
              Confidential:Confidential,
              CurApproverTxt:this.state.MgrName,
              ClientName:client,
              CurApproverId:approverID,
              NoteFilename:filename,
              Sitename:'Main',
              Status:"Submitted to Recommender#1",
              StatusNo:1,
              Notefor : Notefor,
              Purpose : Purpose,
              ReturnName : ReturnName,
              DeptOwnership : DeptOwnership,
              RefferedGuidlines:RefferedGuidlines,
              VettingObservation:VettingObservation
             }).then((iar1: ItemAddResult) => {
              let WFweb=new Web('WF');  
              WFweb.lists.getByTitle('ChecklistNoteNotifications').items.add({
                Title:this.state.Title,
                SeqNo:this.state.seqno,
                Subject:Subj,
                Department:dept,
                Comments:Comment,
                Confidential:Confidential,
                CurApproverId:approverID,
                NotifyId:approverID,
                ApproversId:{results:Approvers},
                Amount:Amount,
                MainRecID:uid,
                RequesterId:requester,
                NoteFilename:filename,
                NoteType:Financial,
                DOP:DOP,
                ClientName:client,
                Migrate:"",
                ControllerId:this.state.ccIDS[0],
                Status:"Submitted to Recommender#1",
                StatusNo:1,
                Notefor : Notefor,
                Purpose : Purpose,
                ReturnName : ReturnName,
                DeptOwnership : DeptOwnership,
                RefferedGuidlines:RefferedGuidlines,
                VettingObservation:VettingObservation
                  }).then((iar2: ItemAddResult) =>{
                    checklistId = iar1.data.ID;
                  for(var i=0;i<Checklisttable.length;i++)
                  {
                    if(Checklisttable[i].itemid == 0)
                    {
                      pnp.sp.site.rootWeb.lists.getByTitle('Checklist').items.add({
                      Title: this.state.Title,
                      SeqNo: this.state.seqno,
                      AppId: uid ,
                      Checklist:Checklisttable[i].checklist,
                      Status:Checklisttable[i].status
                      });
                    }
                  }                                
                }).then(() => {
                    this.updateFirstRecommender(Number(AppItemid)).then(()=>{
                    this.AddWFHistory().then(()=>{
                    this.redirect();
                });
              });
            });
          });
        });
      });
    }); 

  }else{
    this.retrieveFirstApprover().then((Appid)=>{
      // let approverID=Appid[0];
      // let AppItemid=Appid[1];
      
      var approverID=Appid[0];      
      var AppItemid=Appid[1]; 
      let web=new Web('Main');  
      let Approvers:Number[] = [];
      Approvers.push(approverID);           
      web.lists.getByTitle('ChecklistNote').items.getById(uid).update({
        Subject:Subj,
        Comments:Comment,
        CurApproverId:approverID,
        NotifyId:approverID,
        ApproversId:{results:Approvers},
        Amount:Amount,
        NoteFilename:filename,
        NoteType:Financial,
        ClientName:client,
        Migrate:"",
        ControllerId:this.state.ccIDS[0],
        Status:"Submitted to Approver#1",
        StatusNo:6,
        Notefor : Notefor,
        Purpose : Purpose,
        ReturnName : ReturnName,
        DeptOwnership : DeptOwnership,
        RefferedGuidlines:RefferedGuidlines,
        VettingObservation:VettingObservation
          }).then((iar: ItemAddResult) => {
            console.log(iar.data.ID);
            let id=iar.data.ID;
            pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.select('Title','ID','PID').filter("PID eq "+uid ).get().then(Appr=>{
              let Appruverid=Appr[0].ID;
              pnp.sp.site.rootWeb.lists.getByTitle('ChecklistNote').items.getById(Appruverid).update({
              Subject:Subj,
              NoteType:Financial,
              CurApproverTxt:this.state.MgrName,
              CurApproverId:approverID,
              NoteFilename:filename,
              Status:"Submitted to Approver#1",                 
              StatusNo:6,
              Notefor : Notefor,
              Purpose : Purpose,
              ReturnName : ReturnName,
              DeptOwnership : DeptOwnership,
              RefferedGuidlines:RefferedGuidlines,
              VettingObservation:VettingObservation
             }).then((iar1: ItemAddResult) => {
              let WFweb=new Web('WF');  
              WFweb.lists.getByTitle('ChecklistNoteNotifications').items.add({
              Title:this.state.Title,
              SeqNo:this.state.seqno,
              MainRecID:uid,
              Subject:Subj,
              Department:dept,
              RequesterId:requester,
              Comments:Comment,
              CurApproverId:approverID,
              NotifyId:approverID,
              ApproversId:{results:Approvers},
              Amount:Amount,
              NoteFilename:filename,
              NoteType:Financial,
              ClientName:client,
              Migrate:"",
              ControllerId:this.state.ccIDS[0],
              Status:"Submitted to Approver#1",
              StatusNo:6,
              Notefor : Notefor,
              Purpose : Purpose,
              ReturnName : ReturnName,
              DeptOwnership : DeptOwnership,
              RefferedGuidlines:RefferedGuidlines,
              VettingObservation:VettingObservation
              }).then((iar2: ItemAddResult) =>{
                checklistId = iar1.data.ID;
                for(var i=0;i<Checklisttable.length;i++)
                {
                  if(Checklisttable[i].itemid == 0){
                    pnp.sp.site.rootWeb.lists.getByTitle('Checklist').items.add({
                    Title: this.state.Title,
                    SeqNo: this.state.seqno,
                    AppId: checklistId ,
                    Checklist:Checklisttable[i].checklist,
                    Status:Checklisttable[i].status
                    });
                  }
                }                                
            }).then(() => {
                this.updateFirstApprover(Number(AppItemid)).then(()=>{
                this.AddWFHistory().then(()=>{
                this.redirect();
                });
              });
            });
          });
        });
      });
    }); 
  }
}
/*--End Function--*/
 
/*--Redirect to home Page--*/
 private gotoHomePage(): void {
// self.close();
     window.location.replace(this.props.siteUrl);
 }


   /*--Delete Attachment in ChecklistAnnexures library for Annexures attachment--*/
   public DeleteAttachment(vals : string):void{
     debugger;
     this.setState({
       attachments:[]
     });
     let sitename=this.state.Absoluteurl;
     let web=new Web('Main'); 
     let url=sitename+'/ChecklistAnnexures/'+vals;
     let fldr=vals.split("/")[0];
     let fldURL=sitename+'/ChecklistAnnexures/'+fldr;
     web.getFileByServerRelativeUrl(url).recycle().then(data=> {  
       console.log("File Deleted " + vals) ;
       web.getFolderByServerRelativeUrl(fldURL).files.get().then((result) => {
         let links:any[]=[];
      
         for(let i=0;i<result.length;i++){
           links.push(fldr+"/"+result[i].Name);

         }
        
         this.setState({ attachments: links});
        //document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
        
     });
     
     });

   }
 /*--End--*/

/*--Delete Attachment in ChecklistAttach library for Note attachment--*/
   public DeleteNote(vals : string):void{
    debugger;
    this.setState({
      Note:[]
    });
    let sitename=this.state.Absoluteurl;
    let url=sitename+'/ChecklistAttach/'+vals;
    let fldr=vals.split("/")[0];
    let fldURL=sitename+'/ChecklistAttach/'+fldr;
    let web=new Web('Main');           
    web.getFileByServerRelativeUrl(url).recycle().then(data=> {  
      console.log("File Deleted " + vals) ;
      web.getFolderByServerRelativeUrl(fldURL).files.get().then((result) => {
        let links:any[]=[];
     
        for(let i=0;i<result.length;i++){
          links.push(fldr+"/"+result[i].Name);

        }
       
        this.setState({ Note: links});
        this.setState({Notefilename:""});
        // // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
        // document.getElementById("NoteDel").style.display="none";
        let fileUploadInput = document.getElementById('fileUploadInput');
        if (fileUploadInput) {
          fileUploadInput.nodeValue = null;
        }
        let NoteDel =document.getElementById("NoteDel");
        if(NoteDel){NoteDel.style.display="none";}          
    jQuery('#NoteFile').text('');
    });
    
    });

  }
  /*--End--*/

  /*--Adding attachments in Document Library function--*/
   public  AttachLib=(event : any)=> {
     debugger;
        var uploadFlag=true;
    //in case of multiple files,iterate or else upload the first file.
     // let file = fileUpload.files[0];
    let file = event.target.files[0];
    let filesize=file.size/1048576;
    var n = (file.name.length-file.name.lastIndexOf("."));
    //let fileExtn=file.name.substr(file.name.length-(n-1)).toLowerCase();
    let fileExtn=file.name.substr((file.name.lastIndexOf('.') + 1)).toLowerCase();
    let fileSplit=file.name.split(".");
    let fileType=this.state.AttachType;
    let PermissibleExtns=['pdf'];
    let listName='ChecklistAttach';
    let NoteCount=this.state.Note.length;
    let TotalAnnexures=this.state.attachments.length;
    let notetype=this.state.NoteType;
    let fileTest=file.name.substring(0,(file.name.length-n));
  console.log(fileTest);
  let match = new RegExp('[~#%\&{}+.\|]|\\.\\.|^\\.|\\.$').test(fileTest);

     if(fileType!='Note'){
       PermissibleExtns=['png','jpeg','jpg','gif','pdf','doc','docx','xls','xlsx','.eml'];
       listName='ChecklistAnnexures';
     }
     else {
      PermissibleExtns=['pdf'];
           }
     
     
      if(fileSplit.length>2)
      {
        alert('Alert-Selected file double extension is not allowed!');
        // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
        return false;
      }
      else if(match)
      {
      alert('Invalid file name. The name of the attached file contains invalid characters!');
      // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
      return false;

     }else if(file.name.split(".")[0].length >150){
      alert('Invalid file name. file names cannot be more than 150 chars!');
      // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
      return false;
     }
     else if(PermissibleExtns.indexOf(fileExtn.toLowerCase())==-1){
       alert('Alert-Selected file type is not allowed!');
       // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
       return false;
     }
     else if(  filesize>10 ){
       alert('Alert-File size is more than permissible limit!');
       // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
       return false;
     }
     else if(fileType=='Note' && NoteCount==1){
      alert('Alert-Only 1 Note is allowed!');
      // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
      return false;
     }
     else if(fileType!='Note' && TotalAnnexures==20){
      alert('Alert-Only 20 Annexures can be uploaded!');
      // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
      return false;
     }
     else{
     if (file!=undefined || file!=null){
           let SeqNo=this.state.seqno;
            let web=new Web('Main');           
                web.getFolderByServerRelativePath(listName).folders.add(SeqNo).then(data=> {  
         console.log("Folder is created at " + data.data.ServerRelativeUrl) ;
     //assuming that the name of document library is Documents, change as per your requirement, 
     //this will add the file in root folder of the document library, if you have a folder named test, replace it as "/Documents/test"
     
   web.getFolderByServerRelativePath(listName+"/"+SeqNo).files.add(file.name, file, true).then((result) => {
        console.log(file.name + " uploaded successfully!");
        let links:any[]=[];
        
        if(fileType=='Note'){
          this.setState({Notefilename:file.name});
          links=this.state.Note;
          links.push(SeqNo+"/"+file.name);
          this.setState({ Note: links});
        }else{
        links=this.state.attachments;
        links.push(SeqNo+"/"+file.name);
        this.setState({ attachments: links});
      }
        console.log(links);
        
        // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
   });
     }).catch(data=>{  
     console.log(data);  
       uploadFlag=false;

     });  
 
   }
   else{
     uploadFlag=false;
   }

 }
 
 event.preventDefault();
  // return uploadFlag;
   }
/*--End--*/
  /*--Is it a Confidential Note and Do you want to add Client radio button change function */
  
   private Radibtnchangeevent(name : string ,value : string){debugger;
  debugger;
  
    if(name == "radioAttach"){
      this.setState({RadioClient:value});
      if(value=='CYes'){
        jQuery('#divClientName').show();
             }
      else{
        jQuery('#divClientName').hide();
      }

      
          }

          if(name == "radioConf"){

            if(value=='ConfYes'){
              jQuery('#txtConfidential').val('Yes');
                   }
            else{
              jQuery('#txtConfidential').val('No');
            }
      
            
                }
                 
    }
/*--End Function for Radio Btn Change--*/
 
   
 /*-- To get details from Restricted Emails master --*/
 private getRestrictedEmails(){
  debugger;
  pnp.sp.site.rootWeb.lists.getByTitle('RestrictedEmails').items.select("ID,Title,AlertMessage").orderBy("ID asc").getAll().then((items: any[]) => {
    debugger;
    let links: string[]=[];
    let msg: string[]=[];
    for (let i = 0; i < items.length; i++) {
      links += items[i].Title;
      msg += items[i].AlertMessage;
    }
    this.setState({ RestrictedEmails: links });
    this.setState({ RestrictedEmailsMsg: msg });

});
}
/*--End All Functions--*/
    
}
