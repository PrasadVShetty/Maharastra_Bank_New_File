import * as React from 'react';
import styles from './PaperlessApprovalEdit.module.scss';
import { IPaperlessApprovalProps } from './IPaperlessApprovalProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
//import { PeoplePicker } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext, PeoplePicker,PrincipalType} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { EditState } from "../Model/EditState";
import { default as pnp, ItemAddResult, File, Web } from "sp-pnp-js";
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/components/Button';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
//import { CurrentUser } from '@pnp/sp/src/siteusers';
import { Icon } from '@fluentui/react';
import { Button } from 'office-ui-fabric-react/lib/Button';
import { Attachments } from 'sp-pnp-js/lib/graph/attachments';
import { IFrameDialog } from "@pnp/spfx-controls-react/lib/IFrameDialog";
import { IFramePanel } from "@pnp/spfx-controls-react/lib/IFramePanel";
import { SiteUser } from 'sp-pnp-js/lib/sharepoint/siteusers';
import * as jQuery from 'jquery';
import * as $ from 'jquery';
import { SPComponentLoader } from '@microsoft/sp-loader';
SPComponentLoader.loadCss('/sites/EasyApproval/SiteAssets/css/styles.css');
require('../css/custom.css');
// SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css');  

const Delete: any = require('../Images/Delete.png');
const Video: any = require('../Images/Video.png');
const Annex: any = require('../Images/Upload-Annex.png');
const NoteAtt: any = require('../Images/Upload-Note.png');
const Logo: any = require('../Images/Logo.png');
const Expand: any = require('../Images/Expand.jpg');
const Collapse: any = require('../Images/Collapse.jpg');

export default class PNoteFormsEdit extends React.Component<IPaperlessApprovalProps, EditState> {
  constructor(props : any) {
    super(props);
    this.handleTitle = this.handleTitle.bind(this);
    this.handleDesc = this.handleDesc.bind(this);
    this._onCheckboxChange = this._onCheckboxChange.bind(this);
    this._onRenderFooterContent = this._onRenderFooterContent.bind(this);
    this.createItem = this.createItem.bind(this);
    this._getManager = this._getManager.bind(this);
    this._getRecommender = this._getRecommender.bind(this);
    this._getAdmin = this._getAdmin.bind(this);
    this._getCCPeople = this._getCCPeople.bind(this);
    this.DeleteComments = this.DeleteComments.bind(this);

    //  this.setButtonsEventHandlers();
    this.state = {
      name: "",
      description: "",
      CommentsLog: [],
      WFHistoryLog: [],
      selectedItems: [],
      hideDialog: true,
      showPanel: false,
      dpselectedItem: undefined,
      dpselectedItems: [],
      RecomselectedItems: [],
      RecomNewselectedItems: [],
      ReferselectedItems: [],
      ControlselectedItems: [],
      dropdownOptions: [],
      disableToggle: false,
      defaultChecked: false,
      termKey: undefined,
      pplPickerType: "",
      status: "",
      statusno: 0,
      isChecked: false,
      required: "This is required",
      onSubmission: false,
      termnCond: false,
      ManagerEmail: [],
      seqno: "",
      attachments: [],
      Note: [],
      AppAttachments: [],
      AttachType: '',
      Appstatus: '',
      MgrName: '',
      files: null,
      UserID: 0,
      UserEmail: '',
      iframeDialog: true,
      ImgUrl: '',
      ReturnedByID: 0,
      ReturnedByName: '',
      ReferredByID: 0,
      ReferredByName: '',
      ReqID: 0,
      ReqName: '',
      pplTo: 0,
      To: '',
      NoteType: '',
      Notefilename: '',
      Sitename: '',
      Absoluteurl: '',
      CurrApproverEmail: '',
      CurrAppID: 0,
      AdminFlag: '',
      userManagerIDs: [],
      ccIDS: [],
      ccName: '',
      ccEmail: '',
      ModifiedDate: null,
      MarkIDs: [],
      MarkName: [],
      MarkEmails: [],
      MarkItems: [],
      AllApprovers: [],
      RecpID: [],
      RecpName: [],
      RecpEmail: [],
      ReferredCasesCount: 0,
      ReferredCasesLastCount: 0,
      Charsleft: 2000,
      RestrictedEmails: [],
      ChecklistselectedItems:[]
    };
  }
  public render(): React.ReactElement<IPaperlessApprovalProps> {
    debugger;
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
    let qstr = window.location.search.split('uid=');
    let uid = 0;
    if (qstr.length > 1) { uid = parseInt(qstr[1]); }

    return (
      <form >
    <div className={styles.paperlessApprovalEdit} style={{ display: "none" }} id="divMain">
    <div className={styles.editcontainer}>
    
    <div className={styles.formrow}>
    <div id="divHeadingNew" style={{ display: "block", backgroundColor: "#0c78b8", color: '#fff' }}>
    <h3 style={{ fontSize: "18px", textAlign: "center", color: "white", padding: '5px 0px' }}>Note Workflow Form </h3>
    </div>
    </div>
    
    <div className={styles.row} style={{ position: 'relative' }}>
              <div className={styles.frame + " " + styles.column + " " + styles.mobiledivleft} id="divContent">
                <br></br>
                <div id="divMainComments" style={{ display: "none" }} className={styles.formrow + " " + "form-group row"}>
                  <h3 className="text-left" style={{ backgroundColor: "#50B4E6", fontSize: "16px" }}>Note Comments</h3>

                  <div className={styles.lbl + " " + styles.Mcolumn}>
                    <table className={styles.tbl} id="tblMain" style={{ width: "100%" }}>
                      <tr>
                        <th style={{ width: "10%" }}>Page#</th>
                        <th style={{ width: "10%" }}>Doc Reference</th>
                        <th style={{ width: "70%" }}>Comments
                          <span style={{ position: "relative", marginLeft: "10px", color: "Red", fontSize: "14px", fontStyle: "italic" }}>*Note: Max.2000 Chars.</span>
                        </th>
                        <th style={{ width: "10%" }}>Action</th>
                      </tr>
                      <tr>
                        <td><input type="text" id="txtPage"></input></td>
                        <td><input type="text" id="txtRef"></input></td>
                        <td >

                          <textarea rows={3} cols={200} style={{ height: "120px", width: "100%" }} className="notes" id="txtComments"></textarea>
                        </td>
                        <td ><PrimaryButton id="btnComm" style={{ width: "250t", fontSize: "12pt", borderRadius: "5%", backgroundColor: "#50B4E6" }} text="Add Comments" onClick={() => { this.AddComments(); }} /></td>
                      </tr>
                      {this.state.selectedItems ? this.state.selectedItems.map((data) => {
                        console.log(data);
                        return data;
                      }) : null}


                    </table>
                  </div>
                </div>
    <div className={styles.row + ' ' + styles.panelsection} id="GeneralCollapse" style={{ backgroundColor: "#50B4E6", display: "none" }}>
    {/* <img src={Expand} onClick={() => { this.Expand('General'); }}></img> */}
    <Icon iconName='CalculatorAddition' onClick={() => { this.Expand('General'); }} />
    <span>General Section</span>
    </div>
    <div className={styles.row + ' ' + styles.expandpanel} id="GeneralExpand" style={{ display: "block" }}>
    <div className={styles.panelbody}>
    <Icon iconName='StatusCircleBlock2' onClick={() => { this.Collapse('General'); }} />
    {/* <img src={Collapse} onClick={() => { this.Collapse('General'); }} /> */}
    <span> General Section </span>
    </div>
    {/* <img src={Collapse} onClick={() => { this.Collapse('General'); }}></img> */}
    {/* <span style={{ backgroundColor: "#50B4E6" }}> General Section</span> */}
    <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Tcolumn}>Request#</div>
    <div id="tdFileRefNo" style={{ display: "none" }}></div>
    <div className={styles.Vcolumn} id="tdTitle">

    </div>
    </div>
    <br />
    <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Tcolumn}>Requester</div>
    <div className={styles.Vcolumn} id="tdName">

    </div>
    </div>
    <br />
    <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Tcolumn}>Request Date</div>
    <div className={styles.Vcolumn} id="tdDate">

    </div>
    </div>
    <br />
    <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Tcolumn}>Status</div>
    <div className={styles.Vcolumn} id="tdStatus">

    </div>
    </div>
    <br />
    <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Tcolumn}>Current Approver</div>
    <div className={styles.Vcolumn} id="tdCurrApprover">

    </div>
    </div>
    <br />
    <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Tcolumn}>Department</div>
    <div className={styles.Vcolumn} id="divDepartment">
    </div>
    </div>
    {/*<br />
    <br />
    <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Tcolumn}>Note For</div>
    <div className={styles.Vcolumn} id="txtNote">
    </div>
    </div>
    <br />
    <br />
    <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Tcolumn}>Purpose</div>
    <div className={styles.Vcolumn} id="txtPurpose">
    </div>
    </div>
    <br />
    <br />
    <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Tcolumn}>Return Name</div>
    <div className={styles.Vcolumn} id="txtReturn">
    </div>
    </div>
    <br />
    <br />*/}
    <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Tcolumn}>Dept Ownership</div>
    <div className={styles.Vcolumn} id="ddlDeptOwnership">
    </div>
    </div>
    <br />
    <br />
    {/* <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Tcolumn}>Due Date</div>
    <div className={styles.Vcolumn} id="txtDueDate">
    </div>
    </div>
    <br />
    <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Tcolumn}>Place</div>
    <div className={styles.Vcolumn} id="txtPlace">
    </div>
    </div>
    <br />
    <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Tcolumn}>Subject</div>
    <div className={styles.Vcolumn} id="divSubject" >

    </div>
    </div>
    <br /> */}

<div className={styles.formrow + " " + "form-group"}>
    <div className={styles.lbl + " " + styles.tableresponsive}>
        <table className={styles.tbl} id="tblMain100" style={{ width: "100%" }}>
            <thead>
                <tr>
                    <th>SNo</th>
                    <th style={{ width: "40%" }}>Checklist</th>
                    <th>Status</th>  
                </tr>
            </thead>
            <tbody>
                {this.state.ChecklistselectedItems && this.state.ChecklistselectedItems.length > 0 ? (
                    this.state.ChecklistselectedItems.map((data, index) => (
                        <tr key={data.ID}>
                            <td>{index + 1}</td>
                            <td>{data.Checklist}</td>
                            <td>{data.Status}</td>
                        </tr>
                    ))
                ) : (
                    <tr>
                        <td colSpan={3}>No data available</td>
                    </tr>
                )}
            </tbody>
        </table>
    </div>
</div>

    <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Tcolumn}>Note Type</div>
    <div className={styles.Vcolumn} id="divNoteType" >

    </div>
    </div>
    <br />
    <div className={styles.formrow + " " + "form-group row"} id="RowdivAmount" style={{ display: "none" }} >
    <div className={styles.lbl + " " + styles.Tcolumn}>Amount</div>
    <div className={styles.Vcolumn} id="divAmount" >

    </div>
    <br />
    </div>
    <div className={styles.formrow + " " + "form-group row"} style={{ display: "" }} >
    <div className={styles.lbl + " " + styles.Tcolumn}>DOP Details</div>
    <div className={styles.Vcolumn} id="divDOP" >

    </div>
    <br />
    </div>

    <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Tcolumn}>Client Name/Vendor Name</div>
    <div className={styles.Vcolumn} id="divClient" >

    </div>
    <br />
    </div>
    <div className={styles.formrow + " " + "form-group row"} id="divConf" style={{ display: "none" }}>
    <div className={styles.lbl + " " + styles.Tcolumn}>Confidential</div>
    <div className={styles.Vcolumn} id="divConfidential" >

    </div>
    <br />
    </div>

    <div className={styles.formrow + " " + "form-group row"} style={{ display: "none" }}>
    <div className={styles.lbl + " " + styles.Tcolumn}>Comments</div>
    <div className={styles.Vcolumn} id="divComments">
    </div>
    </div>
    </div>
    <div className={styles.row + ' ' + styles.panelsection} id="RecommenderCollapse" style={{ backgroundColor: "#50B4E6", display: "block", fontSize: "16px" }}>
    {/* <img src={Expand} onClick={() => { this.Expand('Recommender'); }}></img> */}
    <Icon iconName='CalculatorAddition' onClick={() => { this.Expand('Recommender'); }} />
    <span>Recommenders Section</span>
    </div>

    <div className={styles.row + ' ' + styles.expandpanel} id="RecommenderExpand" style={{ display: "none", fontSize: "16px" }}>
              <div className={styles.panelbody}>
                <Icon iconName='StatusCircleBlock2' onClick={() => { this.Collapse('Recommender'); }} />
                {/* <img src={Collapse} onClick={() => { this.Collapse('General'); }} /> */}
                <span> Recommenders Section </span>
              </div>
              {/* <img src={Collapse} onClick={() => { this.Collapse('Recommender'); }}></img>
              <span style={{ backgroundColor: "#50B4E6" }}> Recommenders Section</span> */}

              <div className={styles.formrow + " " + "form-group row"} style={{ display: "none" }} id="divAddNewRecommender" >
                <div className={styles.lbl + " " + styles.Tcolumn}>
                  Do you want add new Recommender?
                </div>
                <div className={styles.Vcolumn}>
                  <select id="ddlApprover" onChange={() => this.showDiv('Recomm')} >
                    <option value="Yes">Yes</option>
                    <option value="No" selected>No</option>
                  </select>
                </div>
              </div>
              <div className={styles.formrow + " " + "form-group row"} style={{ display: "none" }} id="divAddRecomm">
                <div className={styles.lbl + " " + styles.tableresponsive}>
                  <table className={styles.tbl} id="tblMain" style={{ width: "100%" }}>
                    <tr>
                      <td style={{ width: "15%" }}>Recommender</td>
                      <td style={{ width: "70%" }} id="RecommenderPPtd">
                        {/* <PeoplePicker context={this.props.context}
                          peoplePickerCntrlclassName={styles.picker}
                          titleText=""
                          tooltipMessage={"Enter email address!"}
                          placeholder={"Person Name or Email address"}
                          groupName={""} // Leave this blank in case you want to filter from all users
                          showtooltip={true}
                          isRequired={false}
                          ensureUser={true}
                          disabled={false}
                          selectedItems={this._getRecommender}
                          defaultSelectedUsers={this.state.RecpEmail}
                          errorMessageClassName={styles.hideElementManager}
                        /> */}
                        <PeoplePicker
                        context={peoplePickerContext}                            
                        personSelectionLimit={100}
                        groupName={""} 
                        showtooltip={true}
                        required={true}
                        disabled={false}
                        searchTextLimit={5}
                        onChange={this._getRecommender}
                        showHiddenInUI={false}
                        ensureUser={true}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        defaultSelectedUsers= {this.state.RecpEmail}
                        errorMessageClassName={styles.hideElementManager}
                        />
                      </td>
                      <td style={{ width: "15%" }}><PrimaryButton style={{ width: "80pt", borderRadius: "5%", backgroundColor: "#50B4E6" }} text="Add Recommender" onClick={() => { this.AddRecommender(); }} /></td>
                    </tr>
                    {this.state.RecomNewselectedItems ? this.state.RecomNewselectedItems.map((data) => {
                      return data;
                    }) : null}


                  </table>
                </div>
              </div>
              <div className={styles.formrow + " " + "form-group"}>
                <div className={styles.lbl + " " + styles.tableresponsive}>
                  <table className={styles.tbl} id="tblMain1" style={{ width: "100%" }}>
                    <tr>
                      <th>SNo</th>
                      <th style={{ width: "40%" }}>Recommender</th>
                      <th>Status</th>
                      <th>Action Date</th>
                    </tr>

                    {this.state.RecomselectedItems ? this.state.RecomselectedItems.map((data) => {
                      console.log(data);
                      return data;
                    }) : null}


                  </table>
                </div>
              </div>
    </div>
    <div className={styles.row + ' ' + styles.panelsection} id="ApproverCollapse" style={{ backgroundColor: "#50B4E6", display: "block", fontSize: "16px" }}>
    {/* <img src={Expand} onClick={() => { this.Expand('Approver'); }}></img> */}
    <Icon iconName='CalculatorAddition' onClick={() => { this.Expand('Approver'); }} />
    <span>Approvers Section</span>
    </div>
    <div className={styles.row + ' ' + styles.expandpanel} id="ApproverExpand" style={{ display: "none", fontSize: "16px" }}>
    {/* <img src={Collapse} onClick={() => { this.Collapse('Approver'); }}></img> */}
    <div className={styles.panelbody}>
    <Icon iconName='StatusCircleBlock2' onClick={() => { this.Collapse('Approver'); }} />
    <span> Approvers Section</span>
    </div>
    <div className={styles.formrow + " " + "form-group"}>
    <div className={styles.lbl + " " + styles.tableresponsive}>

    <table className={styles.tbl} id="tblMain1" style={{ width: "100%" }}>
    <tr>
      <th style={{ width: "10%" }}>SNo</th>
      <th style={{ width: "40%" }}>Approver</th>
      <th style={{ width: "25%" }}>Status</th>
      <th style={{ width: "25%" }}>Action Date</th>
    </tr>

    {this.state.dpselectedItems ? this.state.dpselectedItems.map((data) => {
      console.log(data);
      return data;
    }) : null}


    </table>
    </div>
    </div>
    </div> 
    <div className={styles.row + ' ' + styles.panelsection} id="ControllerCollapse" style={{ backgroundColor: "#50B4E6", display: "none", fontSize: "16px" }}>
              {/* <img src={Expand} onClick={() => { this.Expand('Controller'); }}></img> */}
              <Icon iconName='CalculatorAddition' onClick={() => { this.Expand('Controller'); }} />
              <span>Controller Section</span>
    </div>
    <div className={styles.row + ' ' + styles.expandpanel} id="ControllerExpand" style={{ display: "none", fontSize: "16px" }}>
    <div className={styles.panelbody}>
    <Icon iconName='StatusCircleBlock2' onClick={() => { this.Collapse('Controller'); }} />
    {/* <img src={Collapse} onClick={() => { this.Collapse('Controller'); }}></img> */}
    <span> Controller Section</span>
    </div>
    <div className={styles.formrow + " " + "form-group"}>
    <div className={styles.lbl + " " + styles.tableresponsive}>
      <table className={styles.tbl} id="tblMain1" style={{ width: "100%" }}>
        <tr>
          <th>SNo</th>
          <th>Approver</th>
          <th>Status</th>
          <th>Action Date</th>
        </tr>

        {this.state.ControlselectedItems ? this.state.ControlselectedItems.map((data) => {
          console.log(data);
          return data;
        }) : null}


      </table>
    </div>
    </div>
    </div>  

    <div className={styles.row + ' ' + styles.panelsection} id="RefererCollapse" style={{ backgroundColor: "#50B4E6", display: "block", fontSize: "16px" }}>
              <Icon iconName='CalculatorAddition' onClick={() => { this.Expand('Referer'); }} />
              {/* <img src={Expand} onClick={() => { this.Expand('Referer'); }}></img> */}
              <span>Referrers Section</span>
    </div>
    <div className={styles.row + ' ' + styles.expandpanel} id="RefererExpand" style={{ display: "none", fontSize: "16px" }}>
    <div className={styles.panelbody}>
    {/* <img src={Collapse} onClick={() => { this.Collapse('Referer'); }}></img> */}
    <Icon iconName='StatusCircleBlock2' onClick={() => { this.Collapse('Referer'); }} />
    <span> Referrers Section</span>
    </div>
    <div className={styles.formrow + " " + "form-group"}>
    <div className={styles.lbl + " " + styles.tableresponsive}>
      <table className={styles.tbl} id="tblMain1" style={{ width: "100%" }}>
        <tr>
          <th>SNo</th>
          <th>Referred By</th>
          <th>Referred To</th>
          <th>Status</th>
          <th>Action Date</th>
        </tr>

        {this.state.ReferselectedItems ? this.state.ReferselectedItems.map((data) => {
          console.log(data);
          return data;
        }) : null}


      </table>
    </div>
    </div>
    </div>
    <div className={styles.row + ' ' + styles.panelsection} id="ReferCollapse" style={{ backgroundColor: "#50B4E6", display: "none", fontSize: "16px" }}>
    <Icon iconName='CalculatorAddition' onClick={() => { this.Expand('Refer'); }} />
    {/* <img src={Expand} onClick={() => { this.Expand('Refer'); }}></img> */}
    <span>Seek more Information</span>
    </div>
    <div className={styles.row + ' ' + styles.expandpanel} id="ReferExpand" style={{ display: "none" }}>
    <div className={styles.panelbody}>
    {/* <img src={Collapse} onClick={() => { this.Collapse('Refer'); }}></img> */}
    <Icon iconName='StatusCircleBlock2' onClick={() => { this.Collapse('Refer'); }} />
    <span style={{ backgroundColor: "#50B4E6", fontSize: "16px" }}> Seek more Information</span>
    </div>

    <h3 className="text-left" style={{ backgroundColor: "#50B4E6", fontSize: "16px", display: "none" }}>Refer/Return</h3>
    <div className={styles.formrow + " " + "form-group row"}  >
    <div className={styles.lbl + " " + styles.Tcolumn}>
      Do you want to Seek More Informarion?
    </div>
    <div className={styles.Vcolumn}>
      <select id="ddlRefer" onChange={() => this.showDiv('Refer')} >
        <option value="Yes">Yes</option>
        <option value="No" selected>No</option>
      </select>
    </div>
    </div>

    <div className={styles.formrow + " " + "form-group row"} id="divRefer" style={{ display: "none" }}>
    <div className={styles.lbl + " " + styles.Tcolumn} id="divRefApprover">
      Select Name
    </div>
    <div className={"ms-Grid-col ms-u-sm8 block"} id="ReferPPtd">
      {/*<div  className={styles.Vcolumn}>*/}
      {/* <PeoplePicker context={this.props.context}
        titleText=" "
        personSelectionLimit={1}
        groupName={""} // Leave this blank in case you want to filter from all users
        showtooltip={false}
        isRequired={false}
        ensureUser={true}
        disabled={false}
        selectedItems={this._getManager}
        defaultSelectedUsers={this.state.ManagerEmail}
        errorMessageClassName={styles.hideElementManager}
      /> */}
      <PeoplePicker
      context={peoplePickerContext}                            
      personSelectionLimit={1}
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
    </div>
    </div>
    <div className={styles.formrow + " " + "form-group row"} style={{ display: "none" }}>
    <div className={styles.lbl + " " + styles.Tcolumn}>
      Do you want to Return?
    </div>
    <div className={styles.Vcolumn}>
      <select id="ddlReturn" onChange={() => this.showDiv('Return')} >
        <option value="Yes">Yes</option>
        <option value="No" selected>No</option>
      </select>
    </div>
    </div>
    <div className={styles.formrow + " " + "form-group row"} id="divReturn" style={{ display: "none" }}>
    <div className={styles.lbl + " " + styles.Tcolumn}>
      Select the Name
    </div>
    <div className={styles.Vcolumn}>
      <select id="ddlReturnTo" >
        <option >Select</option>
      </select>
    </div>
    </div>

    </div>  

    <div className={styles.row + ' ' + styles.panelsection} id="CommentsCollapse" style={{ backgroundColor: "#50B4E6", display: "block", fontSize: "16px" }}>
              <Icon iconName='CalculatorAddition' onClick={() => { this.Expand('Comments'); }} />
              {/* <img src={Expand} onClick={() => { this.Expand('Comments'); }}></img> */}
              <span>Comments Log</span>
    </div>
    <div className={styles.row + ' ' + styles.expandpanel} id="CommentsExpand" style={{ display: "none", fontSize: "16px" }}>
    <div className={styles.panelbody}>
    <Icon iconName='StatusCircleBlock2' onClick={() => { this.Collapse('Comments'); }} />
    {/* <img src={Collapse} onClick={() => { this.Collapse('Comments'); }}></img> */}
    <span> Comments Log</span>
    </div>
    <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Mcolumn}>
      <table className={styles.tbl} id="tblComments" style={{ width: "100%" }}>
        <tr>
          <th style={{ width: "10%" }}>SNo</th>
          <th style={{ width: "60%" }}>Comments</th>
          <th style={{ width: "20%" }}>Comments By</th>
        </tr>
        {this.state.selectedItems ? this.state.CommentsLog.map((data) => {
          console.log(data);
          return data;
        }) : null}


      </table>
    </div>
    </div>
    </div> 

    <div className={styles.row + ' ' + styles.panelsection} id="AnnexureCollapse" style={{ backgroundColor: "#50B4E6", display: "block", fontSize: "16px" }}>
              <Icon iconName='CalculatorAddition' onClick={() => { this.Expand('Annexure'); }} />
              {/* <img src={Expand} onClick={() => { this.Expand('Annexure'); }}></img> */}
              <span>Attach NoteAnnexures</span>
    </div>
    <div className={styles.row + ' ' + styles.expandpanel} id="AnnexureExpand" style={{ display: "none" }}>
    <div className={styles.panelbody}>
    <Icon iconName='StatusCircleBlock2' onClick={() => { this.Collapse('Annexure'); }} />
    {/* <img src={Collapse} onClick={() => { this.Collapse('Annexure'); }}></img> */}
    <span>Attach NoteAnnexures</span>
    </div>
    <div className={styles.formrow + " " + "form-group row"} style={{ margin: "0px" }}>

    <div className={styles.lbl + " " + styles.Tcolumn}>
      <a href="#"><img src={Annex} style={{ height: "16pt", marginLeft: "10px" }} onClick={() => { this.UploadAttach(); }}></img></a>
    </div>
    <div className={styles.Vcolumn}>
      {this.state.AppAttachments.map((vals) => {
        let filename = vals.split("/")[1];
        return (<span style={{ position: "relative", padding: "10px" }}><a href={this.props.siteUrl + "/Main/NoteAnnexures/" + vals}>{filename}</a><img src={Delete} style={{ width: "10pt", height: "10pt", position: "absolute" }} onClick={() => this.DeleteAttachment(vals)}></img> </span>);

      })}

    </div>

    </div>
    </div>   
      
    <div className={styles.row + ' ' + styles.panelsection} id="WorkflowCollapse" style={{ backgroundColor: "#50B4E6", display: "block", fontSize: "16px" }}>
              {/* <img src={Expand} onClick={() => { this.Expand('Workflow'); }}></img> */}
              <Icon iconName='CalculatorAddition' onClick={() => { this.Expand('Workflow'); }} />
              <span>Workflow Log</span>
    </div>
    <div className={styles.row + ' ' + styles.expandpanel} id="WorkflowExpand" style={{ display: "none" }}>
    {/* <img src={Collapse} onClick={() => { this.Collapse('Workflow'); }}></img> */}
    <div className={styles.panelbody}>
    <Icon iconName='StatusCircleBlock2' onClick={() => { this.Collapse('Workflow'); }} />
    <span>Workflow Log</span>
    </div>
    <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Mcolumn}>
      <table className={styles.tbl} id="tblHistory" style={{ width: "100%" }}>
        {this.state.selectedItems ? this.state.WFHistoryLog.map((data) => {
          console.log(data);
          return data;
        }) : null}


      </table>
    </div>
    </div>
    </div>

    <div className={styles.row + ' ' + styles.panelsection} id="AttachmentCollapse" style={{ backgroundColor: "#50B4E6", display: "block", fontSize: "16px" }}>
              <Icon iconName='CalculatorAddition' onClick={() => { this.Expand('Attachment'); }} />
              {/* <img src={Expand} onClick={() => { this.Expand('Attachment'); }}></img> */}
              <span>Attachments</span>
    </div>
    <div className={styles.row + ' ' + styles.expandpanel} id="AttachmentExpand" style={{ display: "none" }}>
    <div className={styles.panelbody}>
    <Icon iconName='StatusCircleBlock2' onClick={() => { this.Collapse('Attachment'); }} />
    {/* <img src={Collapse} onClick={() => { this.Collapse('MarkInfo'); }}></img> */}
    <span>Attachments</span>
    </div>
    {/* <img src={Collapse} onClick={() => { this.Collapse('Attachment'); }}></img>
    <span style={{ backgroundColor: "#50B4E6", fontSize: "16px" }}>Attachments</span> */}

    <div className={styles.formrow + " " + "form-group row"}>
    <div hidden className={styles.Mcolumn} id="divAttach" style={{}}>
      <label id="lblAttach" style={{}} className="ms-Label">Attach Note</label>
    </div>
    <div hidden className="ms-Grid-col ms-u-sm12 block hide" id="divAttachButton" style={{ backgroundColor: "white" }}>
      <input type='file' style={{}} id='fileUploadInput' required={true} name='myfile' multiple onChange={this.AttachLib} />
    </div>

    <div className={styles.lbl + " " + styles.Tcolumn}>
      Main Note
    </div>
    <div className={styles.Vcolumn} style={{ backgroundColor: "white" }}>{this.state.Note.map((vals) => {
      return (<span style={{ position: "relative" }}><a href={"javascript:void(window.open('" + vals + "'))"}>{this.state.Notefilename}</a></span>);
    })}</div>

    <div className={styles.lbl + " " + styles.Mcolumn}>
      NoteAnnexures
    </div>

    <div className={styles.Mcolumn} style={{ backgroundColor: "white" }}>
      <table className={styles.tbl} id="tblNoteAnnexures" style={{ width: "100%" }}>
        <tr>
          <th style={{ width: "10%" }}>SNo</th>
          <th style={{ width: "15%" }}>Attachment</th>
          <th style={{ width: "45%" }}>Attached By</th>
          <th style={{ width: "30%" }}>Date</th>
        </tr>
        {this.state.selectedItems ? this.state.attachments.map((Attdata) => {
          return Attdata;
        }) : null}


      </table>
    </div>
    </div>
    </div>

    <br></br>
    <div id="divAppComments" style={{ display: "none" }} className={styles.formrow + " " + "form-group row"}>
    <h3 className="text-left" style={{ backgroundColor: "#50B4E6", fontSize: "16px" }}>Approval Comments</h3>

    <div className={styles.lbl + " " + styles.Mcolumn}>
    <table className={styles.tbl} id="tblAppMain" style={{ width: "100%" }}>
      <tr>
        <td><span style={{ position: "relative", marginLeft: "10px", color: "Red", fontSize: "14px", fontStyle: "italic" }}>*Note: Max.2000 Valid Chars.</span></td>
      </tr>
      <tr>
        <td >
          <p style={{ color: "#50B4E6", fontSize: "12px", fontVariant: "normal" }}>Characters Left: {this.state.Charsleft}</p>
          <textarea rows={3} cols={200} style={{ height: "150px", width: "100%" }} className="notes" onChange={this.CheckComments.bind(this)} id="txtAppComments"></textarea>
        </td>
      </tr>
    </table>
    </div>
    </div>
    <div className={styles.row + ' ' + styles.panelsection} id="MarkInfoCollapse" style={{ backgroundColor: "#50B4E6", display: "none", fontSize: "16px" }}>
    {/* <img src={Expand} onClick={() => { this.Expand('MarkInfo'); }}></img> */}
    <Icon iconName='CalculatorAddition' onClick={() => { this.Expand('MarkInfo'); }} />
    <span>Mark For Information</span>
    </div>
    <div className={styles.row + ' ' + styles.expandpanel} id="MarkInfoExpand" style={{ display: "none" }}>
    <div className={styles.panel}>
    <Icon iconName='StatusCircleBlock2' onClick={() => { this.Collapse('MarkInfo'); }} />
    {/* <img src={Collapse} onClick={() => { this.Collapse('MarkInfo'); }}></img> */}
    <span style={{ backgroundColor: "#50B4E6", fontSize: "16px" }}>Mark For Information</span>
    </div>
    <div className={styles.formrow + " " + "form-group row"}>
    <div className={styles.lbl + " " + styles.Mcolumn}>
      <span style={{ position: "relative", marginLeft: "10px", color: "Red", fontSize: "14px", fontStyle: "italic" }}>*Note: Max.10 recipients can be added.</span>
      <table className={styles.tbl} id="tblMark" style={{ width: "100%" }}>
        <tr id="trAddMark" style={{ display: "none" }}>
          <td style={{ width: "15%" }}>Recipient</td>
          <td style={{ width: "70%" }}>
            {/* <PeoplePicker context={this.props.context}
              peoplePickerCntrlclassName={styles.picker}
              titleText={""}
              personSelectionLimit={1}
              tooltipMessage={"Type and select from suggested names"}
              placeholder={"Person Name or Email address"}
              groupName={""} // Leave this blank in case you want to filter from all users
              showtooltip={true}
              isRequired={false}
              ensureUser={true}
              disabled={false}
              selectedItems={this._getCCPeople}
              defaultSelectedUsers={this.state.MarkEmails}
              errorMessageClassName={styles.hideElementManager}
            /> */}
            <PeoplePicker
                    context={peoplePickerContext}
                    titleText="People Picker"
                    personSelectionLimit={1}
                    groupName={""} 
                    showtooltip={true}
                    required={true}
                    disabled={false}
                    searchTextLimit={5}
                    onChange={this._getCCPeople}
                    showHiddenInUI={false}
                    ensureUser={true}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                    defaultSelectedUsers= {this.state.MarkEmails}
                    errorMessageClassName={styles.hideElementManager}
                    />
          </td>
          <td style={{ width: "15%" }}><PrimaryButton style={{ width: "80pt", borderRadius: "5%", backgroundColor: "#50B4E6" }} id="btnAddMarkForInfo" text="Add Recipient" onClick={() => { this.AddMarkforInfo(); }} /></td>
        </tr>
      </table>
      <table className={styles.tbl} id="tblMark1" style={{ width: "100%" }}>
        {this.state.MarkItems ? this.state.MarkItems.map((data) => {
          return data;
        }) : null}


      </table>
    </div>
    </div>
    <hr></hr>
    </div>

    <div className={styles.formrow + " " + "form-group row"} id="divAdmin" style={{ display: "none" }}>
              <div className={styles.lbl + " " + styles.Tcolumn} >
                Change Current Approver
              </div>
              <div className={styles.Vcolumn}>
                {/* <PeoplePicker context={this.props.context}
                  titleText=" "
                  personSelectionLimit={1}
                  groupName={""} // Leave this blank in case you want to filter from all users
                  showtooltip={false}
                  isRequired={false}
                  ensureUser={true}
                  disabled={false}
                  selectedItems={this._getAdmin}
                  defaultSelectedUsers={this.state.ManagerEmail}
                  errorMessageClassName={styles.hideElementManager}
                /> */}
                <PeoplePicker
                    context={peoplePickerContext}
                    titleText="People Picker"
                    personSelectionLimit={100}
                    groupName={""} 
                    showtooltip={true}
                    required={true}
                    disabled={false}
                    searchTextLimit={5}
                    onChange={this._getAdmin}
                    showHiddenInUI={false}
                    ensureUser={true}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                    defaultSelectedUsers= {this.state.ManagerEmail}
                    errorMessageClassName={styles.hideElementManager}
                    />
              </div>
    </div>     
    </div>

    <div className={styles.frame + " " + styles.column + " " + styles.mobiledivright} id="divFrame">
    <div style={{ display: "none" }} id="divSeqNo"></div>
    <div className={styles.row} id="IframeAttachmentExpand" style={{ backgroundColor: "blanchedalmond", display: "block", fontSize: "16px" }}><img src={Expand} onClick={() => { this.ExpandIframe(); }}></img><span>Full Screen</span></div>
    <div className={styles.row} id="IframeAttachmentCollapse" style={{ backgroundColor: "blanchedalmond", display: "none", fontSize: "16px" }}><img src={Collapse} onClick={() => { this.CollapseIframe(); }}></img><span>Exit Full Screen</span></div>
    <iframe width="100%" id="NoteAttchiframe" height="100%" src={this.state.ImgUrl}></iframe>
       
    </div>
    </div>   
    </div>  
    <div className={styles.formrow + " " + "form-group"} style={{marginTop:'60px'}} >
      <div className={styles.container} >
        <div className={styles.overlay} id="overlay" style={{ display: "none" }} >
          <span className={styles.overlayContent} style={{ textAlign: "center" }}>Please Wait!!!</span>
        </div>
      </div>
      <div className='clearfix clear'></div>
      <div className={styles.formrow + " " + "form-group row"} style={{ marginLeft: "10px", marginRight: "10px" }}>

        <div id="btnApprove" style={{ display: "none", paddingRight: '10px' }} >
          <PrimaryButton className={styles.button} style={{ borderRadius: "5%", backgroundColor: "#50B4E6" }} text="Approve" onClick={() => { this.validateForm(); }} />
        </div>
        <div id="btnCancel" style={{ display: "none", paddingRight: '10px' }}>
          <PrimaryButton className={styles.button} style={{ borderRadius: "5%", backgroundColor: "#f00" }} text="Reject" onClick={() => { this.Rejected(); }} />
        </div>

        <div style={{ display: 'flex' }}>
          <PrimaryButton className={styles.button} id="btnChangeApprover" style={{ display: "none", borderRadius: "5%", backgroundColor: "#50B4E6", paddingRight: '10px' }} text="Change Approver" onClick={() => { this.ChangeApprover(); }} />
          <PrimaryButton className={styles.button} id="btnRefer" style={{ display: "none", borderRadius: "5%", backgroundColor: "#50B4E6", paddingRight: '10px' }} text="Submit" onClick={() => { this.referred(); }} />
          <PrimaryButton className={styles.button} id="btnReturn" style={{ display: "none", borderRadius: "5%", backgroundColor: "#50B4E6", paddingRight: '10px' }} text="Return" onClick={() => { this.returned(); }} />
        </div>
        <div style={{ display: 'flex' }}>
          <PrimaryButton className={styles.button} id="btnReturnBack" style={{ display: "none", borderRadius: "5%", backgroundColor: "#50B4E6", paddingRight: '10px' }} text="Submit" onClick={() => { this.returnBack(); }} />
          <PrimaryButton className={styles.button} id="btnReferBack" style={{ display: "none", borderRadius: "5%", backgroundColor: "#50B4E6", paddingRight: '10px' }} text="Refer Back" onClick={() => { this.referBack(); }} />
          <PrimaryButton className={styles.button} id="btnCallBack" style={{ display: "none", borderRadius: "5%", backgroundColor: "#50B4E6", paddingRight: '10px' }} text="Call Back" onClick={() => { this.CallBack(); }} />
        </div>
        <div id="btnClose" style={{ display: "none", paddingRight: '10px' }}>
          <PrimaryButton className={styles.button} style={{ borderRadius: "5%", backgroundColor: "#f00" }} text="Close" onClick={() => { this.cancel(); }} />
        </div>
        <div id="btnCancelled" style={{ display: "none", paddingRight: '10px' }} >
          <PrimaryButton className={styles.button} style={{ borderRadius: "5%", backgroundColor: "#f00" }} text="Cancel" onClick={() => { this.Cancelled(); }} />
        </div>
        {/* <a href={this.state.Sitename+"/SiteAssets/NoteDetailsWeb/NoteDetailsHistory.aspx/?uid="+uid} target="_blank" >Note Details History</a> */}
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

        <hr></hr>
      </div>


      <div style={{ display: "none" }}>
        <button id="btnAllComments" onClick={() => { this.retrieveAllComments(); }}></button>
      </div>
      </div>
    </div>            
    </form>
    );
  }

  // Starting Code and Functions
  /*-- For Upload Attachment Popup--*/
  public UploadAttach() {
    //document.getElementById('fileUploadInput').click();
    let fileUploadInput = document.getElementById("fileUploadInput");
      if (fileUploadInput) {
      (fileUploadInput as HTMLInputElement).value = ''; // Clear the file input
      }
  }
  /*-- End Upload Attach Function--*/

  /*-- Set the state on change of Files --*/
  public handleChange(files : any) {
    this.setState({
      files: files
    });
  }
  /*-- End Function--*/

  /*-- On Load Function--*/
  public componentDidMount() {

    var reacthandler = this;
    // get Logged-in user's details
    // pnp.sp.web.currentUser.get().then((r: CurrentUser) => { //To get current user details from site 
    pnp.sp.web.currentUser.get().then((r) => {
      debugger;
      let sitename = r['odata.id'].split("/_api")[0];
      this.setState({ Sitename: sitename });
      const uname = r['UserPrincipalName'].split('@')[0];
      let username = r['Title'];
      this.setState({ UserID: r['Id'] });
      this.setState({ name: username });
      let CurrUserEmail = r['LoginName'].split("|")[2];
      this.setState({ UserEmail: CurrUserEmail });
    }).then(() => {
      this.on();
      let qstr = window.location.search.split('uid=');
      let uid = 0;
      if (qstr.length > 1) { uid = parseInt(qstr[1]); }
      this.getRestrictedEmails();
      this.setFields(uid);
    });
    // End Get Current User PNP call

    setTimeout(() => {
      this.off();
    }, 5000);

  }
  // End On Load Function


  


  
  /*--Add comments in Commentslog list logic --*/
  public AddComments() {
    debugger;
    let comment = String(jQuery('#txtComments').val()).trim();
    let page = String(jQuery('#txtPage').val()).trim();
    let ref = String(jQuery('#txtRef').val()).trim();

    if (page == "") {
      alert('Kindly enter Page#!');
      // document.getElementById("txtPage").focus();
      let fileUploadInput = document.getElementById("txtPage");
      if (fileUploadInput) {
      (fileUploadInput as HTMLInputElement).focus();
      }
      return;
    }
    else if (ref == "") {
      alert('Kindly enter Document Reference!');
      // document.getElementById("txtRef").focus();
      const txtRefElement = document.getElementById("txtRef");
      if (txtRefElement) {
          txtRefElement.focus();
      }
      return;
    }
    else if (comment == "") {
      alert('Kindly add Comments!');
      // document.getElementById("txtComments").focus();
      const txtRefElement = document.getElementById("txtComments");
      if (txtRefElement) {
          txtRefElement.focus();
      }
      return;
    }

    else if (comment.length > 2000) {
      alert('Max. 2000 characters are allowed!');
      // document.getElementById("txtComments").focus();
      const txtRefElement = document.getElementById("txtComments");
      if (txtRefElement) {
          txtRefElement.focus();
      }
      return;
    }
    else {
      let SeqNo = this.state.seqno;

      debugger;
      let web = new Web('Main');
      web.lists.getByTitle("CommentsLog").items.add({
        Title: this.state.seqno,
        Page: page,
        Docref: ref,
        Comments: comment,
        Appname: this.state.name,
        Appemail: this.state.UserEmail
      }).then((iar: ItemAddResult) => {
        console.log(iar.data.ID);
        // document.getElementById("txtComments").innerText = '';
        // document.getElementById("txtPage").innerText = '';
        // document.getElementById("txtRef").innerText = '';

        const txtComments = document.getElementById("txtComments");
        const txtPage = document.getElementById("txtPage");
        const txtRef = document.getElementById("txtRef");

        if (txtComments) {
            txtComments.innerText = '';
        }

        if (txtPage) {
            txtPage.innerText = '';
        }

        if (txtRef) {
            txtRef.innerText = '';
        }

        this.retrieveAllComments();
      });
    }
  }  
  /*-- End Comments Log Function --*/

  /*-- Retrieve comments from commentslog List --*/
  public retrieveComments() {
    let title = this.state.seqno;
    // let data = [];
    let data: any[] = [];
    let userEmail = this.state.UserEmail;
    let web = new Web('Main');
    web.lists.getByTitle('CommentsLog').items.select("ID,Title,Page,Docref,Comments,Appname,Appemail,Created").filter("Title eq '" + title + "' and Appemail eq '" + userEmail + "'").orderBy("ID asc").getAll().then((items: any[]) => {
      debugger;
      let tbldata = '';

      let ModifiedDate = this.state.ModifiedDate;
      for (let i = 0; i < items.length; i++) {
        let CurrDate = new Date(items[i].Created);
        if(ModifiedDate)
        {
          if (CurrDate > ModifiedDate) {
            data.push(<tr><td>{items[i].Page}</td><td>{items[i].Docref}</td><td>{items[i].Comments}</td><td><button onClick={() => { this.DeleteComments(items[i].ID); }}>Delete</button></td></tr>);
          }
        }
        
      }
    }).then(() => {
      this.setState({ selectedItems: data });
    });
  }
  // End Function

  /*-- Retrieve All comments from commentslog List --*/
  public retrieveAllComments() {
    let title = this.state.seqno;
    // let data = [];
    let data: any[] = [];
    let web = new Web('Main');
    web.lists.getByTitle('CommentsLog').items.select("ID,Title,Page,Docref,Comments,Appname,Appemail,Created").filter("Title eq '" + title + "'").orderBy("Created", false).get().then((items: any[]) => {
      debugger;
      let tbldata = '';
      let createDate = '';
      for (let i = 0; i < items.length; i++) {
        let dt = new Date(items[i].Created);
        let mnth = (dt.getMonth() + 1).toString();
        let dat = dt.getDate().toString();
        let hrs = dt.getHours().toString();
        let mins = dt.getMinutes().toString();
        if (mnth.length == 1) { mnth = '0' + mnth; } if (dat.length == 1) { dat = '0' + dat; } if (hrs.length == 1) { hrs = '0' + hrs; } if (mins.length == 1) { mins = '0' + mins; }
        createDate = dat + "-" + mnth + "-" + dt.getFullYear() + " " + hrs + ":" + mins;
        data.push(<tr><td>{i + 1}</td><td>{items[i].Comments}</td><td>{items[i].Appname}<br></br>{createDate}</td></tr>);
        // data.push(<tr><td>{items[i].Page}</td><td>{items[i].Docref}</td><td>{items[i].Comments}</td><td>{items[i].Appname}</td></tr>);
      }


    }).then(() => {
      this.setState({ CommentsLog: data });
    });

  }
  /*--End Function--*/

  /*--Delete Comments in Commentlog List--*/
  public DeleteComments(uid: number, event?: React.MouseEvent<HTMLButtonElement>): void {
    debugger;
    event?.preventDefault();
    let web = new Web('Main');
    let list = web.lists.getByTitle("CommentsLog");
    list.items.getById(uid).delete().then(_ => {
      console.log('List Item Deleted');
      this.retrieveComments();
    });

  }
  /*--End--*/
  /*--Retrieve all attached approvers from FApprovals list--*/
  private retrieveApprovers() {
    let title = this.state.seqno;
    // let data = [];
    let data: any[] = [];
    let web = new Web('Main');
    web.lists.getByTitle('FApprovalsChecklist').items.select("ID,Title,AppName,Status,AppEmail,Created,Modified,Seq").filter("Title eq '" + title + "'").orderBy("Seq asc").getAll().then((items: any[]) => {
      debugger;

      for (let i = 0; i < items.length; i++) {

        let createDate = 'NA';
        if (items[i].Status.trim() != 'Pending') {
          let dt = new Date(items[i].Modified);
          let mnth = (dt.getMonth() + 1).toString();
          let dat = dt.getDate().toString();
          let hrs = dt.getHours().toString();
          let mins = dt.getMinutes().toString();
          if (mnth.length == 1) { mnth = '0' + mnth; } if (dat.length == 1) { dat = '0' + dat; } if (hrs.length == 1) { hrs = '0' + hrs; } if (mins.length == 1) { mins = '0' + mins; }
          createDate = dat + "-" + mnth + "-" + dt.getFullYear() + " " + hrs + ":" + mins;
        }

        data.push(<tr><td>{i + 1}</td><td>{items[i].AppName}</td><td>{items[i].Status}</td><td>{createDate}</td></tr>);
      }
    }).then(() => {
      this.setState({ dpselectedItems: data });
    });
  }
  /*--End Function--*/

  /*--Retrieve all attached recommenders from FApprovals list--*/
  private retrieveRecommenders() {
    let title = this.state.seqno;
    // let data = [];
    let data: any[] = [];
    let web = new Web('Main');
    web.lists.getByTitle('ApprovalsChecklist').items.select("ID,Title,AppName,Status,AppEmail,Created,Modified,Seq").filter("Title eq '" + title + "'").orderBy("Seq asc").getAll().then((items: any[]) => {
      debugger;

      for (let i = 0; i < items.length; i++) {
        if (items[i].Status == 'Approved' || items[i].Status == 'Returned Back' || items[i].Status == 'Referred Back') {
          let Return = "<option value='" + items[i].AppEmail + "'>" + items[i].AppName + "</option>";
          jQuery('select[id="ddlReturnTo"]').append(Return);
        }
        let createDate = 'NA';
        if (items[i].Status.trim() != 'Pending') {
          let dt = new Date(items[i].Modified);
          let mnth = (dt.getMonth() + 1).toString();
          let dat = dt.getDate().toString();
          let hrs = dt.getHours().toString();
          let mins = dt.getMinutes().toString();
          if (mnth.length == 1) { mnth = '0' + mnth; } if (dat.length == 1) { dat = '0' + dat; } if (hrs.length == 1) { hrs = '0' + hrs; } if (mins.length == 1) { mins = '0' + mins; }
          createDate = dat + "-" + mnth + "-" + dt.getFullYear() + " " + hrs + ":" + mins;
        }

        data.push(<tr><td>{i + 1}</td><td>{items[i].AppName}</td><td>{items[i].Status}</td><td>{createDate}</td></tr>);
      }

    }).then(() => {
      this.setState({ RecomselectedItems: data });
    });

  }
  /*--End Function--*/

  /*--Following function is used when last Recommender is Approving--*/
  private retrieveApprover(): Promise<any[]> {
    let web = new Web('Main');
    let title = this.state.seqno;
    let approverID : Number[] = [];
    return web.lists.getByTitle('FApprovalsChecklist').items.select("ID,Title,AppName,AppEmail,Approver/ID,Approver/EMail").filter("Title eq '" + title + "'").expand("Approver").orderBy("ID asc").getAll().then((items: any[]) => {
      debugger;
      let indx = 0;
      if (items.length > 0) {
        approverID[0] = items[0].Approver.ID;
        approverID[1] = items[0].ID;
        approverID[2] = 1;
        approverID[3] = items[0].Approver.EMail;
        this.updateNextApprover(items[0].ID);
        this.setState({ MgrName: items[0].AppName });
      }

      return approverID;

    });

  }
  /*--End--*/

  /*--Following function is used to get first approver--*/
  private retrieveFirstApprover(): Promise<any[]> {
    let title = this.state.seqno;
    let web = new Web('Main');
    let approverID : Number[] = [];
    return web.lists.getByTitle('FApprovalsChecklist').items.select("ID,Title,AppName,AppEmail,Approver/ID,Approver/EMail").filter("Title eq '" + title + "'").expand("Approver").orderBy("ID asc").getAll().then((items: any[]) => {
      debugger;
      let indx = 0;

      if (items.length > 0) {
        for (let i = 0; i < items.length; i++) {
          // let userid = this.state.UserEmail;
          let userid = this.state.UserID;
          // if (items[i].AppEmail == userid) {
            if (items[i].Approver.ID == userid) {
            indx = i + 1;
            this.updateFirstApprover(items[i].ID);

          }
          else if (indx == i && i != 0) {
            approverID[0] = items[i].Approver.ID;
            approverID[1] = items[i].ID;
            approverID[2] = i + 1;
            approverID[3] = items[i].Approver.EMail;
            this.updateNextApprover(items[i].ID);
            this.setState({ MgrName: items[i].AppName });
          }
          if (indx == items.length) {
            approverID[0] = 999;
            approverID[1] = 999;
            this.setState({ MgrName: '' });
          }
        }
      } else {
        approverID[0] = 999;
        approverID[1] = 999;
        this.setState({ MgrName: '' });


      }
      if (approverID.length == 0) {
        approverID[0] = 999;
        approverID[1] = 999;
        this.setState({ MgrName: '' });
      }

      return approverID;

    });

  }
  /*--End--*/

  /*--Following function is used to get first recommender--*/
  private retrieveFirstRecommender(): Promise<any[]> {
    let title = this.state.seqno;
    let approverID : Number[] = [];
    let web = new Web('Main');
    return web.lists.getByTitle('ApprovalsChecklist').items.select("ID,Title,AppName,AppEmail,Approver/ID,Approver/EMail").filter("Title eq '" + title + "'").expand("Approver").orderBy("ID asc").getAll().then((items: any[]) => {
      debugger;
      let indx = 0;
      for (let i = 0; i < items.length; i++) {
        if (items.length > 0) {
          // let userid = this.state.UserEmail;
          let userid = this.state.UserEmail;
          if (items[i].AppEmail == userid) {
            indx = i + 1;
            this.updateFirstRecommender(items[i].ID);

          }
          else if (indx == i && i != 0) {
            approverID[0] = items[i].Approver.ID;
            approverID[1] = items[i].ID;
            approverID[2] = i + 1;
            approverID[3] = items[i].Approver.EMail;
            this.updateNextRecommender(items[i].ID);
            this.setState({ MgrName: items[i].AppName });
          }
          if (indx == items.length) {
            approverID[0] = 999;
            approverID[1] = 999;
            this.setState({ MgrName: '' });
          }
        } else {
          approverID[0] = 999;
          approverID[1] = 999;
          this.setState({ MgrName: '' });
        }

      }

      return approverID;

    });
  }
  /*--End--*/

  /*-- To update  first approver in FApprovals List--*/
  private updateFirstApprover(uid: number): Promise<any[]> {
    let web = new Web('Main');
    return web.lists.getByTitle('FApprovalsChecklist').items.getById(uid).update({
      Status: 'Approved'
    }).then(() => {
      console.log('Approver updated');
      return Promise.resolve(['Done']);

    });

  }
  /*--End--*/

  /*-- To update  first recommender status in FApprovals List(Approved person record)--*/
  private updateFirstRecommender(uid: number): Promise<any[]> {
    let web = new Web('Main');
    return web.lists.getByTitle('ApprovalsChecklist').items.getById(uid).update({
      Status: 'Approved'
    }).then(() => {
      console.log('Approver updated');
      return Promise.resolve(['Done']);

    });

  }
  /*--End--*/

  /*-- To update  next approver status in FApprovals List--*/
  private updateNextApprover(uid: number): Promise<any[]> {
    let web = new Web('Main');
    return web.lists.getByTitle('FApprovalsChecklist').items.getById(uid).update({
      Status: 'Submitted'
    }).then(() => {
      console.log('Approver updated');
      return Promise.resolve(['Done']);

    });

  }
  /*--End--*/

  /*-- To update  next recommender status in FApprovals List--*/
  private updateNextRecommender(uid: number): Promise<any[]> {
    let web = new Web('Main');
    return web.lists.getByTitle('ApprovalsChecklist').items.getById(uid).update({
      Status: 'Submitted'
    }).then(() => {
      console.log('Approver updated');
      return Promise.resolve(['Done']);

    });

  }
  /*--End--*/

  /*-- Old code to update recommender--*/
  private updateAllRecommenders(uid: number, n: number, i: number): Promise<any[]> {
    debugger;
    let web = new Web('Main');
    if (n == 0 && i == 0) {
      return web.lists.getByTitle('ApprovalsChecklist').items.getById(uid).update({
        Status: 'Submitted'
      }).then(() => {
        console.log('Approver updated');
        return Promise.resolve(['Done']);
      });
    }
    else if (n == 1 && i == 0) {
      return web.lists.getByTitle('ApprovalsChecklist').items.getById(uid).update({
        Status: 'Approved'
      }).then(() => {
        console.log('Approver updated');
        return Promise.resolve(['Done']);
      });
    }
    else if (n == 1 && i == 1) {
      return web.lists.getByTitle('ApprovalsChecklist').items.getById(uid).update({
        Status: 'Submitted'
      }).then(() => {
        console.log('Approver updated');
        return Promise.resolve(['Done']);
      });
    }
    else {
      return web.lists.getByTitle('ApprovalsChecklist').items.getById(uid).update({
        Status: 'Pending'
      }).then(() => {
        console.log('Approver updated');
        return Promise.resolve(['Done']);

      });
    }


  }
  /*--End--*/

  /*-- update approvers status--*/
  private updateAllApprovers(uid: number, n: number, i: number): Promise<any[]> {
    debugger;
    let web = new Web('Main');
    if (n == 0 && i == 0) {
      return web.lists.getByTitle('FApprovalsChecklist').items.getById(uid).update({
        Status: 'Submitted'
      }).then(() => {
        console.log('Approver updated');
        return Promise.resolve(['Done']);
      });
    }
    else if (n == 1 && i == 0) {
      return web.lists.getByTitle('FApprovalsChecklist').items.getById(uid).update({
        Status: 'Approved'
      }).then(() => {
        console.log('Approver updated');
        return Promise.resolve(['Done']);
      });
    }
    else if (n == 1 && i == 1) {
      return web.lists.getByTitle('FApprovalsChecklist').items.getById(uid).update({
        Status: 'Submitted'
      }).then(() => {
        console.log('Approver updated');
        return Promise.resolve(['Done']);
      });
    }
    else {
      return web.lists.getByTitle('FApprovalsChecklist').items.getById(uid).update({
        Status: 'Pending'
      }).then(() => {
        console.log('Approver updated');
        return Promise.resolve(['Done']);

      });
    }


  }
  /*--End--*/

  /*-- Old code to get Dept Alias--*/
  private getDepartmentAlias(deptname: string): Promise<any[]> {
    debugger;
    let web = new Web('Main');
    return web.lists.getByTitle('Departments').items.select("ID,Title,Dept_Alias,GroupName").filter("Title eq '" + deptname + "'").orderBy("ID asc").getAll().then((items: any[]) => {
      debugger;
      let dept = items[0].GroupName;
      this.setState({ description: dept });
      return Promise.resolve([dept]);
    });
  }
  // End Function

  /*--To get saved data from Notes list and update to current form fields--*/
  private setFields(uid: number) {
    debugger;
    let fldr = '';
    let web = new Web('Main');
    pnp.sp.site.rootWeb.lists.getByTitle("Checklist").items.select("ID,Title,SeqNo,Checklist,AppId,Status").filter(`AppId eq ${uid}`).orderBy("ID", true).get().then((items: any[]) => {
          if (items.length > 0) {
              this.setState({ ChecklistselectedItems: items });
          } else {
              this.setState({ ChecklistselectedItems: [] });
          }
      }).then(()=>{web.lists.getByTitle("ChecklistNote").items.select("ID,Title,Department,Created,SeqNo,Status,StatusNo,Comments,Subject,NoteType,NoteFilename,DeptAlias,Amount,ClientName,Confidential,Requester/Title,Requester/EMail,Requester/Name,Requester/ID,CurApprover/EMail,CurApprover/Title,CurApprover/ID,CurApprover/Name,ReturnedBy/ID,ReturnedBy/Title,ReferredBy/ID,ReferredBy/Title,ReferredTo/ID,ReferredTo/Title,Controller/ID,Controller/EMail,Controller/Title,Approvers/ID,DOP,WorkflowFlag,Modified,RefCount,Notefor,Purpose,ReturnName,DeptOwnership,DueDate,Place").expand('Requester,CurApprover,ReturnedBy,Controller,ReferredBy,ReferredTo,Approvers').filter('ID eq ' + uid).orderBy("ID asc").getAll().then((items: any[]) => {
      debugger;
      console.log(items);
      let statusno = items[0].StatusNo;
      this.setState({ statusno: statusno });
      if (statusno == 0) {

        //window.top.location.replace(this.state.Sitename + '/SitePages/NoteApproval.aspx/?Pid=' + uid);
        if (window.top) {
          window.top.location.replace(this.state.Sitename + '/SitePages/NoteApproval.aspx/?Pid=' + uid);
        }
      }
      else {
        this.setState({ ReferredCasesCount: items[0].RefCount });
        this.setState({ ReferredCasesLastCount: items[0].RefCount });

        //document.getElementById('divMain').style.display = 'block';
        const divMain = document.getElementById("divMain");
        if (divMain) {
          divMain.style.display = 'block';
        }
        let relativeurl = this.state.Sitename.split(".com")[1];
        let fldrUrl = 'NoteAttach/' + items[0].SeqNo;
        let fldrUrl1 = 'NoteAnnexures/' + items[0].SeqNo;
        fldr = items[0].Title;
        let title = items[0].Title;
        this.setState({ Notefilename: items[0].NoteFilename });
        this.setState({ NoteType: items[0].NoteType });
        //document.getElementById("divNoteType").innerText = items[0].NoteType;
        const divNoteType = document.getElementById("divNoteType");
        if (divNoteType) {
          divNoteType.innerText = items[0].NoteType;
        }
        let modDate = new Date(items[0].Modified);
        this.setState({ ModifiedDate: modDate });
        console.log('Modified Date---' + this.state.ModifiedDate);

        let DeptAlias = items[0].DeptAlias;
        if (DeptAlias == 'HRD') {
          //document.getElementById('divConf').style.display = 'block';
          const divConf = document.getElementById("divConf");
          if (divConf) {
            divConf.innerText = '';
          }
        }
        this.setState({ description: DeptAlias });

        this.setState({ To: items[0].To });

        //document.getElementById("tdTitle").innerText = title;
        let tdTitle1 = document.getElementById("tdTitle");
        if (tdTitle1) {
          tdTitle1.innerText = title;
        }
        // document.getElementById("tdName").innerText = items[0].Requester.Title;
        let tdName1 = document.getElementById("tdName");
        if (tdName1) {
          tdName1.innerText = items[0].Requester.Title;
        }
        this.setState({ ReqID: items[0].Requester.ID });
        this.setState({ ReqName: items[0].Requester.Title });
        let ReqEmail = items[0].Requester.Name.split("|")[2];
        let CAppEmail = '';
        let CAppID = 0;
        if (items[0].CurApprover) {
          if (items[0].CurApprover.Name != null) {
            CAppEmail = items[0].CurApprover.Name.split("|")[2];
            CAppID = items[0].CurApprover.ID;
            this.setState({ CurrApproverEmail: CAppEmail });
            this.setState({ CurrAppID: items[0].CurApprover.ID });
            let tdCurrApprover = document.getElementById("tdCurrApprover");
            if(tdCurrApprover)
            {tdCurrApprover.innerText = items[0].CurApprover.Title;}            
          }
        }

        let Return = "<option value='" + ReqEmail + "'>" + items[0].Requester.Title + "</option>";

        jQuery('select[id="ddlReturnTo"]').append(Return);
        // document.getElementById("tdStatus").innerText = items[0].Status;

        // document.getElementById("divSubject").innerText = items[0].Subject;
        // document.getElementById("divComments").innerText = items[0].Comments;
        // let seqno = items[0].SeqNo;
        // document.getElementById("divSeqNo").innerText = items[0].SeqNo;

        // let dt = items[0].Created.toString().split("T")[0].split("-");
        // let date = dt[2]; let mnth = dt[1];
        // if (date.length == 1) { date = "0" + date; }
        // if (mnth.length == 1) { mnth = "0" + mnth; }
        // document.getElementById("tdDate").innerText = date + "-" + mnth + "-" + dt[0];
        // let dept = items[0].Department;
        // let Client = items[0].ClientName;
        // let Amount = items[0].Amount;
        // document.getElementById("divDepartment").innerText = dept;
        // document.getElementById("divClient").innerText = Client;
        // document.getElementById("divConfidential").innerText = items[0].Confidential;
        // document.getElementById("divDOP").innerText = items[0].DOP;
        // document.getElementById("divAmount").innerText = Amount;
        // if (Amount > 0) {
        //   document.getElementById('RowdivAmount').style.display = 'block';
        // }

        let tdStatus = document.getElementById("tdStatus");
        if (tdStatus) tdStatus.innerText = items[0].Status;

        //added on 16/02/2025
        // let tdPurpose = document.getElementById("txtPurpose");
        // if (tdPurpose) tdPurpose.innerText = items[0].Purpose;
        // let tdNotefor = document.getElementById("txtNote");
        // if (tdNotefor) tdNotefor.innerText = items[0].Notefor;
        // let tdReturnName = document.getElementById("txtReturn");
        // if (tdReturnName) tdReturnName.innerText = items[0].ReturnName;
        // let tdDeptOwnership = document.getElementById("ddlDeptOwnership");
        // if (tdDeptOwnership) tdDeptOwnership.innerText = items[0].DeptOwnership;
        // let tdDueDate = document.getElementById("txtDueDate");
        // if (tdDueDate) tdDueDate.innerText = items[0].DueDate;
        // let tdPlace = document.getElementById("txtPlace");
        // if (tdPlace) tdPlace.innerText = items[0].Place;
        

        //end
        let divSubject = document.getElementById("divSubject");
        if (divSubject) divSubject.innerText = items[0].Subject;

        let divComments = document.getElementById("divComments");
        if (divComments) divComments.innerText = items[0].Comments;

        let seqno = items[0].SeqNo;
        let divSeqNo = document.getElementById("divSeqNo");
        if (divSeqNo) divSeqNo.innerText = seqno;

        let dt = items[0].Created.toString().split("T")[0].split("-");
        let date = dt[2], mnth = dt[1];
        if (date.length == 1) { date = "0" + date; }
        if (mnth.length == 1) { mnth = "0" + mnth; }

        let tdDate = document.getElementById("tdDate");
        if (tdDate) tdDate.innerText = date + "-" + mnth + "-" + dt[0];

        let dept = items[0].Department;
        let Client = items[0].ClientName;
        let Amount = items[0].Amount;

        let divDepartment = document.getElementById("divDepartment");
        if (divDepartment) divDepartment.innerText = dept;

        let divClient = document.getElementById("divClient");
        if (divClient) divClient.innerText = Client;

        let divConfidential = document.getElementById("divConfidential");
        if (divConfidential) divConfidential.innerText = items[0].Confidential;

        let divDOP = document.getElementById("divDOP");
        if (divDOP) divDOP.innerText = items[0].DOP;

        let divAmount = document.getElementById("divAmount");
        if (divAmount) divAmount.innerText = Amount;

        let rowDivAmount = document.getElementById('RowdivAmount');
        if (rowDivAmount && Amount > 0) rowDivAmount.style.display = 'block';

        this.setState({ seqno: items[0].SeqNo });


        if (items[0].Controller) {
          if (items[0].Controller.ID != null) {
            //document.getElementById('ControllerCollapse').style.display = 'block';
            let ControllerCollapse = document.getElementById('ControllerCollapse');
            if(ControllerCollapse)
            {ControllerCollapse.style.display = 'block';}
            
            let Controller = [];
            Controller.push(items[0].Controller.ID);
            this.setState({ ccIDS: Controller });
            this.setState({ ccName: items[0].Controller.Title });
            this.setState({ ccEmail: items[0].Controller.EMail });
          }
        }
        if (items[0].Approvers) {
          if (items[0].Approvers.length > 0) {
            // let Approvers = [];
            let Approvers: any[] = [];
            $.each(items[0].Approvers, (index, value) => {
              Approvers.push(value.ID);
            });
            this.setState({ AllApprovers: Approvers });
          }
        }

        web.getFolderByServerRelativeUrl(fldrUrl).files.get().then((result) => {
          let links: any[] = [];
          let Notelinks: any[] = [];
          let Applinks: any[] = [];
          for (let i = 0; i < result.length; i++) {
            links.push(seqno + "/" + result[i].Name);
            //  Applinks.push(seqno+"/PDF/"+result[i].Name);
          }

          let url = '';
          let host = window.location.hostname;

          // if ((items[0].Status != 'Approved' || items[0].Status != 'Rejected') && (CAppEmail == this.state.UserEmail)) {
            if ((items[0].Status != 'Approved' || items[0].Status != 'Rejected') && (CAppID == this.state.UserID)) {
            url = this.state.Sitename + "/SiteAssets/web/Editor.aspx?file=" + this.state.Sitename + "/Main/NoteAttach/" + links[0];
            Notelinks.push(url);

          }
          else {
            url = this.state.Sitename + "/SiteAssets/web/Eviewer.aspx?file=" + this.state.Sitename + "/Main/NoteAttach/" + links[0];
            Notelinks.push(url);
          }
          this.setState({ ImgUrl: url });
          console.log(links);
          this.setState({ Note: Notelinks });

        });

        // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
        this.retrieveApprovers();
        this.retrieveRecommenders();
        this.retrieveRefers();
        this.retrieveController();
        this.retrieveAllComments();
        this.retrieveHistory();

        let Alinks: any[] = [];
        web.getFolderByServerRelativeUrl(fldrUrl1).files.select('*,Author/Title').expand('Author').orderBy("TimeCreated", false).get().then((result) => {
          for (let i = 0; i < result.length; i++) {

            let Attdt = new Date(result[i].TimeCreated);
            let Attmnth = (Attdt.getMonth() + 1).toString();
            let dat = Attdt.getDate().toString();
            let hrs = Attdt.getHours().toString();
            let mins = Attdt.getMinutes().toString();
            let secs = Attdt.getSeconds().toString();
            if (Attmnth.length == 1) { Attmnth = '0' + Attmnth; } if (dat.length == 1) { dat = '0' + dat; } if (hrs.length == 1) { hrs = '0' + hrs; } if (mins.length == 1) { mins = '0' + mins; } if (secs.length == 1) { secs = '0' + secs; }
            let createDate = dat + "-" + Attmnth + "-" + Attdt.getFullYear() + " " + hrs + ":" + mins + ":" + secs;

            let Atturl = <a href={"javascript:void(window.open('" + this.props.siteUrl + "/Main/NoteAnnexures/" + seqno + "/" + result[i].Name + "'))"}>{result[i].Name}</a>;
            Alinks.push(<tr><td>{i + 1}</td><td>{Atturl}</td><td>{result[i].Author.Title}</td><td>{createDate}</td></tr>);
          }
        });

        setTimeout(() => {
          this.setState({ attachments: Alinks });
        }, 2000);

        this.setState({ Appstatus: items[0].Status });


        let ownersGroup : String[] = [];
        let sitename = this.state.Sitename;
        if (sitename.indexOf('CapApprovals') != -1) {
          ownersGroup[0] = "CapApprovals Owners";
          ownersGroup[1] = "CapApprovals " + DeptAlias;
        }
        else { ownersGroup[0] = "EasyApproval Owners"; ownersGroup[1] = "EasyApproval " + DeptAlias; }
        pnp.sp.web.siteUsers.getById(this.state.UserID).groups.get().then((grps: any) => {
          for (let i = 0; i < grps.length; i++) {
            if (ownersGroup.indexOf(grps[i].Title) != -1) {
              this.setState({ AdminFlag: 'Yes' });
              break;
            }
          }

          if (items[0].Status == 'Approved' || items[0].Status == 'Closed') {
            // document.getElementById('btnApprove').style.display = 'none';
            // document.getElementById('btnCancel').style.display = 'none';
            // document.getElementById('ReferCollapse').style.display = 'none';
            // document.getElementById('AnnexureCollapse').style.display = 'none';
            // document.getElementById('AnnexureExpand').style.display = 'none';            
            // document.getElementById('divMainComments').style.display = 'none';
            // document.getElementById('divAppComments').style.display = 'none';
            // document.getElementById('divAttach').style.display = 'none';
            // document.getElementById('lblAttach').style.display = 'none';
            // document.getElementById('divAttachButton').style.display = 'none';
            // document.getElementById('fileUploadInput').style.display = 'none';
            // document.getElementById('btnClose').style.display = 'block';
            // document.getElementById('MarkInfoCollapse').style.display = 'block';
            // Get elements by their IDs
            let btnApprove = document.getElementById('btnApprove');
            let btnCancel = document.getElementById('btnCancel');
            let referCollapse = document.getElementById('ReferCollapse');
            let annexureCollapse = document.getElementById('AnnexureCollapse');
            let annexureExpand = document.getElementById('AnnexureExpand');
            let divMainComments = document.getElementById('divMainComments');
            let divAppComments = document.getElementById('divAppComments');
            let divAttach = document.getElementById('divAttach');
            let lblAttach = document.getElementById('lblAttach');
            let divAttachButton = document.getElementById('divAttachButton');
            let fileUploadInput = document.getElementById('fileUploadInput');
            let btnClose = document.getElementById('btnClose');
            let markInfoCollapse = document.getElementById('MarkInfoCollapse');

            // Apply style changes only if elements exist
            if (btnApprove) btnApprove.style.display = 'none';
            if (btnCancel) btnCancel.style.display = 'none';
            if (referCollapse) referCollapse.style.display = 'none';
            if (annexureCollapse) annexureCollapse.style.display = 'none';
            if (annexureExpand) annexureExpand.style.display = 'none';
            if (divMainComments) divMainComments.style.display = 'none';
            if (divAppComments) divAppComments.style.display = 'none';
            if (divAttach) divAttach.style.display = 'none';
            if (lblAttach) lblAttach.style.display = 'none';
            if (divAttachButton) divAttachButton.style.display = 'none';
            if (fileUploadInput) fileUploadInput.style.display = 'none';
            if (btnClose) btnClose.style.display = 'block';
            if (markInfoCollapse) markInfoCollapse.style.display = 'block';

            //trAddMark

            if (this.state.ReqID == this.state.UserID) {
              // document.getElementById('trAddMark').style.display = 'block';
              let trAddMark = document.getElementById('trAddMark');
              if(trAddMark)
              {
                trAddMark.style.display = 'block';
              }
            }
            this.retrieveMarkForInfo();
          }
          else {
            // if ((CAppEmail == this.state.UserEmail && statusno == 1) || (CAppEmail == this.state.UserEmail && statusno == 6)) {
              if ((CAppID == this.state.UserID && statusno == 1) || (CAppID == this.state.UserID && statusno == 6)) {
              let divAddNewRecommender = document.getElementById('divAddNewRecommender');
              if (statusno == 1) { if(divAddNewRecommender){divAddNewRecommender.style.display = 'block';} }
              // document.getElementById('btnApprove').style.display = 'block';
              // document.getElementById('btnCancel').style.display = 'block';
              // document.getElementById('ReferCollapse').style.display = 'block';
              // document.getElementById('divAppComments').style.display = 'block';
              // document.getElementById('GeneralCollapse').style.display = 'block';
              // document.getElementById('GeneralExpand').style.display = 'none';
              // document.getElementById('btnClose').style.display = 'block';
              
              let btnApprove = document.getElementById('btnApprove');
              let btnCancel = document.getElementById('btnCancel');
              let referCollapse = document.getElementById('ReferCollapse');
              let divAppComments = document.getElementById('divAppComments');
              let generalCollapse = document.getElementById('GeneralCollapse');
              let generalExpand = document.getElementById('GeneralExpand');
              let btnClose = document.getElementById('btnClose');

              // Apply style changes only if elements exist
              if (btnApprove) btnApprove.style.display = 'block';
              if (btnCancel) btnCancel.style.display = 'block';
              if (referCollapse) referCollapse.style.display = 'block';
              if (divAppComments) divAppComments.style.display = 'block';
              if (generalCollapse) generalCollapse.style.display = 'block';
              if (generalExpand) generalExpand.style.display = 'none';
              if (btnClose) btnClose.style.display = 'block';

            }

            // else if (CAppEmail == this.state.UserEmail && statusno == 3) {
              else if (CAppID == this.state.UserID && statusno == 3) {
              this.setState({ ReturnedByID: items[0].ReturnedBy.ID });
              this.setState({ ReturnedByName: items[0].ReturnedBy.Title });
                // document.getElementById('btnApprove').style.display = 'none';
                // document.getElementById('btnCancel').style.display = 'none';
                // document.getElementById('ReferCollapse').style.display = 'none';
                // document.getElementById('divAppComments').style.display = 'block';
                // document.getElementById('GeneralCollapse').style.display = 'block';
                // document.getElementById('GeneralExpand').style.display = 'none';
                // document.getElementById('btnReturnBack').style.display = 'block';
                // document.getElementById('btnClose').style.display = 'block';
                
                let btnApprove = document.getElementById('btnApprove');
                let btnCancel = document.getElementById('btnCancel');
                let referCollapse = document.getElementById('ReferCollapse');
                let divAppComments = document.getElementById('divAppComments');
                let generalCollapse = document.getElementById('GeneralCollapse');
                let generalExpand = document.getElementById('GeneralExpand');
                let btnReturnBack = document.getElementById('btnReturnBack');
                let btnClose = document.getElementById('btnClose');

                // Apply style changes only if elements exist
                if (btnApprove) btnApprove.style.display = 'none';
                if (btnCancel) btnCancel.style.display = 'none';
                if (referCollapse) referCollapse.style.display = 'none';
                if (divAppComments) divAppComments.style.display = 'block';
                if (generalCollapse) generalCollapse.style.display = 'block';
                if (generalExpand) generalExpand.style.display = 'none';
                if (btnReturnBack) btnReturnBack.style.display = 'block';
                if (btnClose) btnClose.style.display = 'block';
            }
            // else if (CAppEmail == this.state.UserEmail && statusno == 4) {
              else if (CAppID == this.state.UserID && statusno == 4) {
              debugger;
              this.setState({ ReferredByID: items[0].ReferredBy.ID });
              this.setState({ ReferredByName: items[0].ReferredBy.Title });
              // document.getElementById('btnApprove').style.display = 'none';
              // document.getElementById('btnCancel').style.display = 'none';
              // document.getElementById('ReferCollapse').style.display = 'none';
              // document.getElementById('divAppComments').style.display = 'block';
              // document.getElementById('GeneralCollapse').style.display = 'block';
              // document.getElementById('GeneralExpand').style.display = 'none';
              // document.getElementById('btnReferBack').style.display = 'block';
              // document.getElementById('btnClose').style.display = 'block';            
              let btnApprove = document.getElementById('btnApprove');
              let btnCancel = document.getElementById('btnCancel');
              let referCollapse = document.getElementById('ReferCollapse');
              let divAppComments = document.getElementById('divAppComments');
              let generalCollapse = document.getElementById('GeneralCollapse');
              let generalExpand = document.getElementById('GeneralExpand');
              let btnReferBack = document.getElementById('btnReferBack');
              let btnClose = document.getElementById('btnClose');

              // Apply style changes only if elements are found
              if (btnApprove) btnApprove.style.display = 'none';
              if (btnCancel) btnCancel.style.display = 'none';
              if (referCollapse) referCollapse.style.display = 'none';
              if (divAppComments) divAppComments.style.display = 'block';
              if (generalCollapse) generalCollapse.style.display = 'block';
              if (generalExpand) generalExpand.style.display = 'none';
              if (btnReferBack) btnReferBack.style.display = 'block';
              if (btnClose) btnClose.style.display = 'block';

              // Check the max Referee's count
              if (items[0].RefCount >= 3) {
              }
              else {
                // document.getElementById('ReferCollapse').style.display = 'block';
                let referCollapse = document.getElementById('ReferCollapse');
                if (referCollapse) {
                    referCollapse.style.display = 'block';
                }
              }
              let web1 = new Web('Main');
              web1.lists.getByTitle('RApprovalsChecklist').items.select('Title', 'ID', 'AppEmail').filter("Title eq '" + seqno + "' and Status eq 'Referred'").getAll().then((RApprovalsItems: any[]) => {

                if (RApprovalsItems.length == 0) {
                  this.setState({ ReferredCasesLastCount: 0 });
                }

              });

            }
            else if (items[0].WorkflowFlag == null && this.state.ReqID == this.state.UserID) {
              // document.getElementById('btnApprove').style.display = 'none';
              // document.getElementById('btnCancel').style.display = 'none';
              // document.getElementById('ReferCollapse').style.display = 'none';
              // document.getElementById('divMainComments').style.display = 'none';
              // document.getElementById('divAppComments').style.display = 'none';
              // document.getElementById('GeneralCollapse').style.display = 'none';
              // document.getElementById('GeneralExpand').style.display = 'block';
              // //document.getElementById('btnCallBack').style.display='block';
              // document.getElementById('btnClose').style.display = 'block';
              // document.getElementById('AnnexureCollapse').style.display = 'none';
              // if (items[0].Status != "Cancelled" && items[0].Status != "Recalled Back" && items[0].Status != "Cancel" && items[0].Status != 'Rejected' && items[0].Status != 'Referred Back' && statusno != 10 && statusno != 4 && statusno != 3 && items[0].Status != "Returned" && items[0].Status != "Returned Back" && items[0].Status != "Referred" && items[0].Status != "Called Back" && items[0].Status != "Approved") {
              //   document.getElementById('btnCallBack').style.display = 'block';
              // }
              
              let btnApprove = document.getElementById('btnApprove');
              let btnCancel = document.getElementById('btnCancel');
              let referCollapse = document.getElementById('ReferCollapse');
              let divMainComments = document.getElementById('divMainComments');
              let divAppComments = document.getElementById('divAppComments');
              let generalCollapse = document.getElementById('GeneralCollapse');
              let generalExpand = document.getElementById('GeneralExpand');
              let btnClose = document.getElementById('btnClose');
              let annexureCollapse = document.getElementById('AnnexureCollapse');
              let btnCallBack = document.getElementById('btnCallBack');

              if (btnApprove) btnApprove.style.display = 'none';
              if (btnCancel) btnCancel.style.display = 'none';
              if (referCollapse) referCollapse.style.display = 'none';
              if (divMainComments) divMainComments.style.display = 'none';
              if (divAppComments) divAppComments.style.display = 'none';
              if (generalCollapse) generalCollapse.style.display = 'none';
              if (generalExpand) generalExpand.style.display = 'block';
              if (btnClose) btnClose.style.display = 'block';
              if (annexureCollapse) annexureCollapse.style.display = 'none';

              // Check if the status doesn't match the exclusion conditions, and show btnCallBack if necessary
              if (items[0] && items[0].Status !== "Cancelled" && items[0].Status !== "Recalled Back" &&
                  items[0].Status !== "Cancel" && items[0].Status !== "Rejected" &&
                  items[0].Status !== "Referred Back" && statusno !== 10 && statusno !== 4 &&
                  statusno !== 3 && items[0].Status !== "Returned" && items[0].Status !== "Returned Back" &&
                  items[0].Status !== "Referred" && items[0].Status !== "Called Back" && items[0].Status !== "Approved") {
                  if (btnCallBack) btnCallBack.style.display = 'block';
              }
            }
            else {
              if (this.state.AdminFlag == 'Yes' && (statusno == 1 || statusno == 4 || statusno == 6)) {
                // Check if elements exist before modifying their styles
                let divAdmin = document.getElementById('divAdmin');
                let btnChangeApprover = document.getElementById('btnChangeApprover');
                let btnCancelled = document.getElementById('btnCancelled');
                let divAppComments = document.getElementById('divAppComments');
                let annexureCollapse = document.getElementById('AnnexureCollapse');

                if (divAdmin) divAdmin.style.display = 'block';
                if (btnChangeApprover) btnChangeApprover.style.display = 'block';
                if (btnCancelled) btnCancelled.style.display = 'block';
                if (divAppComments) divAppComments.style.display = 'block';
                if (annexureCollapse) annexureCollapse.style.display = 'none';

              }
              // document.getElementById('AnnexureCollapse').style.display = 'none';
              // document.getElementById('AnnexureExpand').style.display = 'none';
              // document.getElementById('btnApprove').style.display = 'none';
              // document.getElementById('btnCancel').style.display = 'none';
              // document.getElementById('ReferCollapse').style.display = 'none';
              // document.getElementById('divMainComments').style.display = 'none';
              // document.getElementById('GeneralCollapse').style.display = 'none';
              // document.getElementById('GeneralExpand').style.display = 'block';
              // document.getElementById('btnClose').style.display = 'block';
              
              let annexureCollapse = document.getElementById('AnnexureCollapse');
              let annexureExpand = document.getElementById('AnnexureExpand');
              let btnApprove = document.getElementById('btnApprove');
              let btnCancel = document.getElementById('btnCancel');
              let referCollapse = document.getElementById('ReferCollapse');
              let divMainComments = document.getElementById('divMainComments');
              let generalCollapse = document.getElementById('GeneralCollapse');
              let generalExpand = document.getElementById('GeneralExpand');
              let btnClose = document.getElementById('btnClose');

              if (annexureCollapse) annexureCollapse.style.display = 'none';
              if (annexureExpand) annexureExpand.style.display = 'none';
              if (btnApprove) btnApprove.style.display = 'none';
              if (btnCancel) btnCancel.style.display = 'none';
              if (referCollapse) referCollapse.style.display = 'none';
              if (divMainComments) divMainComments.style.display = 'none';
              if (generalCollapse) generalCollapse.style.display = 'none';
              if (generalExpand) generalExpand.style.display = 'block';
              if (btnClose) btnClose.style.display = 'block';

            }
          }


        });
        // let ht = document.getElementById('divContent').clientHeight + 20;
        // console.log(document.getElementById('divFrame').clientHeight);
        // console.log(ht);
        // document.getElementById('divFrame').style.height = ht + "px";
        let divContent = document.getElementById('divContent');
        let divFrame = document.getElementById('divFrame');

        if (divContent && divFrame) {
            let ht = divContent.clientHeight + 20;
            console.log(divFrame.clientHeight);
            console.log(ht);
            divFrame.style.height = ht + "px";
        }

      } // Ending else condition of statusno =0
    })
  });
  }
  /*--End Set Fields Function --*/

  /*-- To save name,email and id for refer people picker--*/
  /*private _getManager(items: any[]) {
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
  }*/

  /*--End--*/

  //Added for get manager

  //Restricted Emails
  // private getRestrictedEmails() {
  //   debugger;
  //   pnp.sp.site.rootWeb.lists.getByTitle('RestrictedEmails').items.select("ID,Title").orderBy("ID asc").getAll().then((items: any[]) => {
  //     debugger;
  //     let links: string[];
  //     for (let i = 0; i < items.length; i++) {
  //       links += items[i].Title;
  //     }
  //     this.setState({ RestrictedEmails: links });

  //   });
  // }

  private getRestrictedEmails() {
    debugger;
    pnp.sp.site.rootWeb.lists.getByTitle('RestrictedEmails').items.select("ID,Title").orderBy("ID asc").getAll().then((items: any[]) => {
        debugger;
        let links: string[] = []; // Initialize the array before using it
        for (let i = 0; i < items.length; i++) {
            links.push(items[i].Title); // Use push to add items to the array
        }
        this.setState({ RestrictedEmails: links });
    });
  }


  /*-- To save name,email and id for refer people picker--*/
  private _getManager(items: any[]) {
    debugger;
    this.state.userManagerIDs.length = 0;
    let tempuserMngArr = [];
    let MgrEmail = [];
    let MgrName = '';
    let restricedEmails = this.state.RestrictedEmails;
    for (let item in items) {

      tempuserMngArr.push(items[item].id);
      MgrName = items[item].text;
      MgrEmail.push(items[item].loginName.split("|")[2]);
      // alert(items[item].id);
    }
    if (MgrEmail.length > 0) {

      if (restricedEmails.indexOf(MgrEmail[0].toLowerCase()) >= 0) {
        alert(MgrEmail[0] + ' cannot be added, kindly select proper name id');
        setTimeout(() => {
          $('#ReferPPtd >div>div>div>div>div>div>div>span>div>button>div>i').click();
          $('#ReferPPtd >div>div>div>div>div>div>div>input').focus();

          return;

        }, 500);
      } else {
        this.setState({ userManagerIDs: tempuserMngArr });
        this.setState({ ManagerEmail: MgrEmail });
        this.setState({ MgrName: MgrName });
      }
    }

  }

  /*-- To save name,email and id for Change Current Approver people picker--*/
  private _getAdmin(items: any[]) {
    debugger;
    this.state.userManagerIDs.length = 0;
    let tempuserMngArr = [];
    let MgrEmail = [];
    let MgrName = '';
    for (let item in items) {
      tempuserMngArr.push(items[item].id);
      MgrName = items[item].text;
      MgrEmail.push(items[item].loginName.split("|")[2]);
      // alert(items[item].id);
    }
    this.setState({ userManagerIDs: tempuserMngArr });
    this.setState({ ManagerEmail: MgrEmail });
    this.setState({ MgrName: MgrName });
  }
  /*--End--*/
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

  private _log(str: string): () => void {
    return (): void => {
      console.log(str);
    };
  }
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
      //  self.close();


    }
  }
  /*--End --*/

  /*--Function for On-Close Panel --*/
  private _onClosePanel = () => {
    this.setState({ showPanel: false });

  }
  /*--End Function--*/

  /*--Function for onShow panel--*/
  private _onShowPanel = () => {
    this.setState({ showPanel: true });
  }
  /*--End Function--*/

  /*--Function to set Title --*/
  private handleTitle(value: string): void {
    return this.setState({
      name: value
    });
  }
  /*--End --*/
  /*--Function to set Description State --*/
  private handleDesc(value: string): void {
    return this.setState({
      description: value
    });
  }
  /*--End --*/

  /*--Function for CheckBox Change --*/
  private _onCheckboxChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean): void {
    console.log(`The option has been changed to ${isChecked}.`);
    this.setState({ termnCond: (isChecked) ? true : false });
  }
  /*--End --*/

  /*--Function for CloseDaialog--*/
  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }
  /*--End --*/

  /*--Function For Show Dialog --*/
  private _showDialog = (status: string): void => {
    this.setState({ hideDialog: false });
    this.setState({ status: status });
  }
  /*--End --*/

  /**
   * Form Validation on Submit
   */
  private validateForm(): void {
    let allowCreate: boolean = true;
    this.setState({ onSubmission: true });

    if (allowCreate) {
      this._onShowPanel();
    }
  }

  /*-- Redirecting page logic --*/
  private redirect() {
    const query = window.location.search.split('uid=')[1];
    let uid = 0;
    if (query != undefined) { uid = parseInt(query); }
    let deptAlias = this.state.description;
    let homeURL = this.props.siteUrl.split(deptAlias)[0];
    if (uid == 0) {
      window.location.replace(homeURL);
    }
    else {
      setTimeout(() => {
        window.location.replace(homeURL);
        // self.close();
      }, 3000);
    }
  }
  /*--old code--*/
  private createItem(): void {
    debugger;
    jQuery('#Createbutton').remove();
    jQuery('#Cancelbutton').remove();
    let web = new Web('Main');
    this._onClosePanel();
    // let Comments=this.state.selectedItems;
    // let Comments = String(jQuery('#txtAppComments').val()).trim();
    let Comments: string = '';
    let Comments2 = document.getElementById("txtAppComments") as HTMLInputElement | null;
    if (Comments2) {
        Comments = Comments2.value.trim();
    }

    if (Comments.length > 0 && Comments.length < 2000) {

      this.on();

      const query = window.location.search.split('uid=')[1];
      let uid = 0;
      if (query != undefined) { uid = parseInt(query); }
      let statusno = this.state.statusno;

      if (statusno == 1) {
        this.AddAppComments().then(() => {
          this.retrieveFirstRecommender().then((Appid) => {

            if (Appid[0] != 999) {
              let approverID = Appid[0];
              let AppItemid = Appid[1];
              let AppCount = Appid[2];
              let AppEmail = Appid[3];
              let CurrUser = this.state.name;
              let Appname = this.state.MgrName;
              let Approvers = this.state.AllApprovers;
              Approvers.push(approverID);
              web.lists.getByTitle("ChecklistNote").items.getById(uid).update({
                // Description: this.state.description,
                CurApproverId: approverID,
                NotifyId: approverID,
                ApproversId: { results: Approvers },
                Migrate: "",
                Status: "Submitted to " + Appname + " (Recommender#" + AppCount.toString() + ")",
                WorkflowFlag: "Triggered",
                Comments: Comments,
                StatusNo: 1
              }).then(() => {
                pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.select('Title', 'ID', 'PID').filter("PID eq " + uid).get().then(r => {
                  let Approverid = r[0].ID;
                  pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.getById(Approverid).update({
                    CurApproverId: approverID,
                    CurApproverTxt: AppEmail,
                    Status: "Submitted to " + Appname + " (Recommender#" + AppCount.toString() + ")"
                  }).then(() => {

                    let statuslog = 'Submitted';
                    let Notifstatus = "1-Submitted to " + Appname + " (Recommender#" + AppCount.toString() + ")";
                    this.AddWFHistory(statuslog).then(() => {
                      this.AddNotesNotifications(Notifstatus, approverID).then(() => {
                        this.dummyHistory().then(() => {
                          this.redirect();
                        });
                      });
                    });
                  });
                });
              });


            } // If Condition for Next Recommender
            else {
              this.retrieveApprover().then((Approverid) => {
                let approverID = Approverid[0];
                let AppItemid = Approverid[1];
                let AppCount = Approverid[2];
                let AppEmail = Approverid[3];
                let CurrUser = this.state.name;
                let Appname = this.state.MgrName;
                let Approvers = this.state.AllApprovers;
                Approvers.push(approverID);
                web.lists.getByTitle("ChecklistNote").items.getById(uid).update({
                  // Description: this.state.description,
                  CurApproverId: approverID,
                  NotifyId: approverID,
                  ApproversId: { results: Approvers },
                  WorkflowFlag: "Triggered",
                  Migrate: "",
                  Comments: Comments,
                  Status: "Submitted to " + Appname + " (Approver#" + AppCount.toString() + ")",
                  StatusNo: 6
                }).then(() => {
                  pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.select('Title', 'ID', 'PID').filter("PID eq " + uid).get().then(r => {
                    let Apprvrid = r[0].ID;
                    pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.getById(Apprvrid).update({
                      CurApproverId: approverID,
                      CurApproverTxt: AppEmail,
                      Status: "Submitted to " + Appname + " (Approver#" + AppCount.toString() + ")"
                    }).then(() => {

                      let statuslog = 'Submitted';
                      let Notifstatus = "6-Submitted to " + Appname + " (Approver#" + AppCount.toString() + ")";
                      this.AddWFHistory(statuslog).then(() => {
                        this.AddNotesNotifications(Notifstatus, approverID).then(() => {
                          this.dummyHistory().then(() => {
                            this.redirect();
                          });
                        });
                      });
                    });
                  });
                });
              });
            } // Else if last Recommender

          });
        }); // Eding Add Comments
      } // Ending if of statusno=1
      else {
        this.Approve();
      }
      //    this._onClosePanel();
      //   this._showDialog("Submitting Request");

    }
    else {

      alert('Comments are mandatory, it cannot contain more than 2000 chars!');
      // document.getElementById("txtAppComments").focus();
      const txtAppComments = document.getElementById("txtAppComments");
      if (txtAppComments) {
          txtAppComments.focus();
      }
    }


  }
  /*--End--*/

  /*--Approver button click logic--*/
  /**
   **/
  private async Approve() {
    debugger;
    let web = new Web('Main'); // Initialize SharePoint Web instance
    let Controller = this.state.ccName;
    const query = window.location.search.split('uid=')[1]; // Extract UID from URL
    let Comments: string = '';
    let Comments2 = document.getElementById("txtAppComments") as HTMLInputElement | null;
    if (Comments2) {
        Comments = Comments2.value.trim();
    }

    let uid = query ? parseInt(query) : 0; // Parse UID from query string

    try {
        await this.AddAppComments(); // Add approval comments
        let Appid = await this.retrieveFirstApprover(); // Retrieve next approver details

        if (Appid[0] !== 999) { // If there is a next approver
            debugger;
            let [approverID, AppItemid, AppCount, AppEmail] = Appid;
            let CurrUser = this.state.name;
            let Appname = this.state.MgrName;
            let Approvers = [...this.state.AllApprovers, approverID];

            // Update the Notes list with new approver details
            await web.lists.getByTitle("ChecklistNote").items.getById(uid).update({
                CurApproverId: approverID,
                ApproversId: { results: Approvers },
                NotifyId: approverID,
                Migrate: "",
                Comments: Comments,
                Status: `Submitted to ${Appname} (Approver#${AppCount})`,
                WorkflowFlag: "Triggered",
                StatusNo: 6
            });

            let r = await pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.select('Title', 'ID', 'PID').filter(`PID eq ${uid}`).get();
            let Approverid = r[0].ID;
            await pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.getById(Approverid).update({
                CurApproverId: approverID,
                CurApproverTxt: AppEmail,
                StatusNo: 6,
                WorkflowFlag: "Triggered",
                Status: `Submitted to ${Appname} (Approver#${AppCount})`
            });

            let Notifstatus = `6-Submitted to ${Appname} (Approver#${AppCount})`;
            await this.AddWFHistory('Submitted');
            await this.AddNotesNotifications(Notifstatus, approverID);
            await this.dummyHistory();
            this.redirect();
        } else { // If no next approver, check for Controller
            debugger;
            let ccID = this.state.ccIDS;
            let ccEmail = this.state.ccEmail;
            let curApprover = this.state.UserID;

            if (Controller && ccID[0] !== curApprover) { // Assign Controller as approver
                let Approvers = [...this.state.AllApprovers, ccID[0]];
                await web.lists.getByTitle("ChecklistNote").items.getById(uid).update({
                    CurApproverId: ccID[0],
                    ApproversId: { results: Approvers },
                    NotifyId: ccID[0],
                    Migrate: "",
                    Comments: Comments,
                    Status: `Submitted to ${Controller} (Controller)`,
                    StatusNo: 6
                });

                let r = await pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.select('Title', 'ID', 'PID').filter(`PID eq ${uid}`).get();
                let Approverid = r[0].ID;
                await pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.getById(Approverid).update({
                    CurApproverId: ccID[0],
                    CurApproverTxt: ccEmail,
                    StatusNo: 6,
                    Status: `Submitted to ${Controller} (Controller)`
                });

                let Notifstatus = `6-Submitted to ${Controller} (Controller)`;
                await this.AddWFHistory('Submitted');
                await this.AddNotesNotifications(Notifstatus, ccID[0]);
                await this.dummyHistory();
                this.redirect();
            } else { // Final Approval Step
                debugger;
                await web.lists.getByTitle("ChecklistNote").items.getById(uid).update({
                    Migrate: "",
                    NotifyId: -1,
                    CurApproverId: -1,
                    Status: "Approved",
                    Comments: Comments,
                    WorkflowFlag: "Triggered",
                    StatusNo: 7
                });

                let r = await pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.select('Title', 'ID', 'PID').filter(`PID eq ${uid}`).get();
                let Approverid = r[0].ID;
                await pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.getById(Approverid).update({
                    CurApproverId: -1,
                    CurApproverTxt: '',
                    Status: "Approved",
                    StatusNo: 7
                });

                let title = document.getElementById("tdTitle")?.innerText || "";
                let dept = this.state.description;
                let subject = document.getElementById("divSubject")?.innerText || "";

                await pnp.sp.site.rootWeb.lists.getByTitle('WatermarkPDF').items.add({
                    Title: title,
                    Status: 'Approved',
                    Department: dept,
                    Subject: subject,
                    SeqNo: this.state.seqno,
                    PID: uid.toString(),
                    Flag: 'Pending',
                    NoteAttachLib: 'NoteAttach'
                });

                if (ccID[0] === curApprover) {
                    let seqno = this.state.seqno;
                    let AppID = await web.lists.getByTitle('CApprovalsChecklist').items.select('Title', 'ID', 'AppEmail').filter(`Title eq '${seqno}'`).get();
                    let NoteAppID = AppID[0].ID;
                    await web.lists.getByTitle('CApprovalsChecklist').items.getById(NoteAppID).update({ Status: 'Approved' });
                }

                let Notifstatus = "7-Approved";
                await this.AddWFHistory('Approved');
                await this.AddNotesNotifications(Notifstatus, -1);
                await this.dummyHistory();
                this.redirect();
            }
        }
    } catch (error) {
        console.error("Error in approval process:", error);
    }
}

  /*--End--*/

  /*--Cancel button click logic--*/
  private Cancelled() {
    debugger;
    let web = new Web('Main');
    this._onClosePanel();
    // let Comments = String(jQuery('#txtAppComments').val()).trim();
    let Comments: string = '';
    let Comments2 = document.getElementById("txtAppComments") as HTMLInputElement | null;
    if (Comments2) {
        Comments = Comments2.value.trim();
    }
    if (confirm('Are you sure you want to cancel the Note?')) {

      if (Comments.length > 0 && Comments.length < 2000) {
        this.on();
        const query = window.location.search.split('uid=')[1];
        let uid = 0;
        this.AddAppComments().then(() => {
          if (query != undefined) { uid = parseInt(query); }
          web.lists.getByTitle("ChecklistNote").items.getById(uid).update({
            Migrate: "Cancel",
            NotifyId: -1,
            CurApproverId: -1,
            Status: "Cancelled",
            StatusNo: 11
          }).then(() => {
            pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.select('Title', 'ID', 'PID').filter("PID eq " + uid).get().then(r => {
              let Approverid = r[0].ID;
              pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.getById(Approverid).update({
                CurApproverId: -1,
                Status: "Cancelled",
                StatusNo: 11
              }).then(() => {
                let statuslog = 'Cancelled';
                let Notifstatus = "11-Cancelled";
                this.AddWFHistory(statuslog).then(() => {
                  this.AddNotesNotifications(Notifstatus, -1).then(() => {
                    this.dummyHistory().then(() => {
                      this.redirect();
                    });
                  });
                });
              });
            });
          });
        });
      }
      else {
        alert('Comments are mandatory, it cannot contain more than 2000 chars!');
        //document.getElementById("txtAppComments").focus();
        const txtAppComments = document.getElementById("txtAppComments");
        if (txtAppComments) {
            txtAppComments.focus();
        }
      }
    }

  }
  /*--End--*/

  /*--Reject button click logic--*/
  private Rejected() {
    debugger;
    let web = new Web('Main');
    this._onClosePanel();
    // let Comments = String(jQuery('#txtAppComments').val()).trim();
    let Comments: string = '';
    let Comments2 = document.getElementById("txtAppComments") as HTMLInputElement | null;
    if (Comments2) {
        Comments = Comments2.value.trim();
    }
    if (confirm('Are you sure you want to reject the Note?')) {

      if (Comments.length > 0 && Comments.length < 2000) {
        this.on();
        let curStatusNo = this.state.statusno;
        let listname = '';
        let Controller = this.state.ccIDS;
        if (curStatusNo == 1) {
          listname = 'ApprovalsChecklist';
        }
        else if (Controller.length > 0 && curStatusNo == 6) {
          if (this.state.UserID == Controller[0]) {
            listname = 'CApprovalsChecklist';
          }
          else {
            listname = 'FApprovalsChecklist';
          }

        }
        else {
          listname = 'FApprovalsChecklist';
        }
        let seqno = this.state.seqno;
        let userID = this.state.UserID;
        let UserEmail = this.state.UserEmail;        
        const query = window.location.search.split('uid=')[1];
        let uid = 0;
        if (query != undefined) { uid = parseInt(query); }

        this.AddAppComments().then(() => {
          web.lists.getByTitle("ChecklistNote").items.getById(uid).update({
            Migrate: "",
            NotifyId: -1,
            Comments: Comments,
            CurApproverId: -1,
            WorkflowFlag: "Triggered",
            Status: "Rejected",
            StatusNo: 10
          }).then(() => {
            pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.select('Title', 'ID', 'PID').filter("PID eq " + uid).get().then(r => {
              let Approverid = r[0].ID;
              pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.getById(Approverid).update({
                // ApproversId:{results:AllAppIDs},
                CurApproverId: -1,
                Status: "Rejected",
                StatusNo: 10
              }).then(() => {
                // web.lists.getByTitle(listname).items.select('Title', 'ID', 'AppEmail').filter("Title eq '" + seqno + "' and AppEmail eq '" + UserEmail + "'").get().then(AppID => {
                  web.lists.getByTitle(listname).items.select("ID,Title,AppName,AppEmail,Approver/ID,Approver/EMail").expand("Approver").filter("Title eq '" + seqno + "' and ApproverId eq '" + userID + "'").get().then(AppID => {
                  let NoteAppID = AppID[0].ID;
                  web.lists.getByTitle(listname).items.getById(NoteAppID).update({
                    Status: 'Rejected'
                  }).then(() => {

                    // let title = document.getElementById("tdTitle").innerText;
                    const title = document.getElementById("tdTitle");
                    if (title) {
                      title.focus();
                    }
                    let dept = this.state.description;
                    // let subject = document.getElementById("divSubject").innerText;
                    const subject = document.getElementById("divSubject");
                    if (subject) {
                      subject.focus();
                    }
                    pnp.sp.site.rootWeb.lists.getByTitle('WatermarkPDF').items.add({
                      Title: title,
                      Status: 'Rejected',
                      Department: dept,
                      Subject: subject,
                      SeqNo: this.state.seqno,
                      PID: uid.toString(),
                      Flag: 'Pending',
                      NoteAttachLib: 'NoteAttach'

                    }).then(() => {
                      let statuslog = 'Rejected';
                      let Notifstatus = "10-Rejected";
                      this.AddWFHistory(statuslog).then(() => {
                        this.AddNotesNotifications(Notifstatus, -1).then(() => {
                          this.dummyHistory().then(() => {
                            this.redirect();
                          });
                        });
                      });
                    });
                  });
                });
              });
            });
          });
        });
      }
      else {
        alert('Comments are mandatory, it cannot contain more than 2000 chars!');
        // document.getElementById("txtAppComments").focus();
        const txtAppComments = document.getElementById("txtAppComments");
        if (txtAppComments) {
            txtAppComments.focus();
        }
      }
    }

  }
  /*--End--*/

  /*--call back button click logic--*/
  private CallBack() {
    this.on();
    let web = new Web('Main');
    let UserEmail = this.state.UserEmail;
    let userID = this.state.UserID;
    const query = window.location.search.split('uid=')[1];
    let uid = 0;
    if (query != undefined) { uid = parseInt(query); }
    web.lists.getByTitle("ChecklistNote").items.getById(uid).update({
      Migrate: "",
      CurApproverId: userID,
      Status: "Called Back",
      StatusNo: 0
    }).then(() => {
      pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.select('Title', 'ID', 'PID').filter("PID eq " + uid).get().then(r => {
        let Approverid = r[0].ID;
        pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.getById(Approverid).update({
          CurApproverId: userID,
          Status: "Called Back",
          StatusNo: 0
        }).then(() => {
          let statuslog = 'Called Back';
          let Notifstatus = "0-Called Back";
          this.AddWFHistory(statuslog).then(() => {
            this.AddNotesNotifications(Notifstatus, userID).then(() => {
              this.dummyHistory().then(() => {
                this.redirect();
              });
            });
          });
        });
      });
    });
  }
  /*--End--*/

  /*--Reffered functionality--*/
  private referred() {
    debugger;
    let web = new Web('Main');
    // let Comments = String(jQuery('#txtAppComments').val()).trim();
    let Comments: string = '';
    let Comments2 = document.getElementById("txtAppComments") as HTMLInputElement | null;
    if (Comments2) {
        Comments = Comments2.value.trim();
    }
    let Controller = this.state.ccIDS;
    if (Comments.length > 0 && Comments.length < 2000) {
      let RefCount = this.state.ReferredCasesCount + 1;
      let listname = 'FApprovalsChecklist';
      let curStatusNo = this.state.statusno;
      if (curStatusNo == 1) {
        listname = 'ApprovalsChecklist';
      }
      else if (Controller.length > 0 && curStatusNo == 6) {
        if (this.state.UserID == Controller[0]) {
          listname = 'CApprovalsChecklist';
        }

      }
      else if (curStatusNo == 4) {
        listname = 'RApprovalsChecklist';
      }
      else {
        listname = 'FApprovalsChecklist';
      }
      let seqno = this.state.seqno;
      let currUser = this.state.UserID;
      let RequesterID = this.state.ReqID;
      let ReturnVal = this.state.userManagerIDs;
      let ReferEmail = this.state.ManagerEmail;
      let ReferName = this.state.MgrName;
      if (ReturnVal.length == 0) {
        alert('Kindly select the name!');
        return false;

      }
      else {
        debugger;
        let mgrEmail = ReferEmail[0];
        this._onClosePanel();
        this.on();
        let UserEmail = this.state.UserEmail;
        let userID = this.state.UserID;
        const query = window.location.search.split('uid=')[1];
        let uid = 0;
        if (query != undefined) { uid = parseInt(query); }
        this.AddAppComments().then(() => {
          web.lists.getByTitle("ChecklistNote").items.getById(uid).update({
            Migrate: "",
            Comments: Comments,
            ReferredById: currUser,
            CurApproverId: ReturnVal[0],
            NotifyId: ReturnVal[0],
            ReferredToId: ReturnVal[0],
            Status: "Referred",
            WorkflowFlag: "Triggered",
            StatusNo: 4,
            RefCount: RefCount,
          }).then(() => {
            pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.select('Title', 'ID', 'PID').filter("PID eq " + uid).get().then(r => {
              let Approverid = r[0].ID;
              pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.getById(Approverid).update({
                Status: "Referred",
                CurApproverId: ReturnVal[0],
                CurApproverTxt: mgrEmail,
                StatusNo: 4
              }).then(() => {
                // web.lists.getByTitle(listname).items.select('Title', 'ID', 'AppEmail').filter("Title eq '" + seqno + "' and AppEmail eq '" + UserEmail + "'").get().then(AppID => {
                  web.lists.getByTitle(listname).items.select("ID,Title,AppName,AppEmail,Approver/ID,Approver/EMail").expand("Approver").filter("Title eq '" + seqno + "' and ApproverId eq '" + userID + "'").get().then(AppID => {
                  let NoteAppID = AppID[0].ID;
                  web.lists.getByTitle(listname).items.getById(NoteAppID).update({
                    Status: 'Referred'
                  }).then(() => {
                    web.lists.getByTitle('RApprovalsChecklist').items.add({
                      Title: this.state.seqno,
                      Status: 'Pending',
                      Seq: 1,
                      ApproverId: ReturnVal[0],
                      AppID: ReturnVal[0],
                      AppName: ReferName,
                      AppEmail: ReferEmail[0]
                    }).then(() => {
                      let statuslog = 'Referred';
                      let Notifstatus = "4-Referred";
                      this.AddWFHistory(statuslog).then(() => {
                        this.AddNotesNotifications(Notifstatus, ReturnVal[0]).then(() => {
                          this.dummyHistory().then(() => {
                            this.redirect();
                          });
                        });
                      });
                    });
                  });
                });
              });
            });
          });
        });

      }
    }
    else {
      this._onClosePanel();
      alert('Comments are mandatory, it cannot contain more than 2000 chars!');
      // document.getElementById("txtAppComments").focus();
      const txtAppComments = document.getElementById("txtAppComments");
      if (txtAppComments) {
          txtAppComments.focus();
      }
    }


  }
  /*--End--*/
  /*--referback functionality--*/
  private referBack() {
    debugger;
    let web = new Web('Main');
    //  let Comments=this.state.selectedItems;
    // let Comments = String(jQuery('#txtAppComments').val()).trim();
    let Comments: string = '';
    let Comments2 = document.getElementById("txtAppComments") as HTMLInputElement | null;
    if (Comments2) {
        Comments = Comments2.value.trim();
    }

    if (Comments.length > 0 && Comments.length < 2000) {
      let listname = 'FApprovalsChecklist';
      let curStatusNo = this.state.statusno;

      this._onClosePanel();
      this.on();
      let ReturnById = this.state.ReferredByID;
      let req = this.state.ReqID;
      // let curruser=this.state.UserID;
      console.log(ReturnById);

      debugger;
      let seqno = this.state.seqno;
      let UserEmail = this.state.UserEmail;
      let userid = this.state.UserID;
      const query = window.location.search.split('uid=')[1];
      let uid = 0;
      if (query != undefined) { uid = parseInt(query); }
      this.AddAppComments().then(() => {
        this.checkRefApprover(ReturnById).then((App) => {
          let statusno = 6;
          let StatusN = "Referred Back";
          if (App == 'Approver') { statusno = 1; listname = 'ApprovalsChecklist'; }
          else {
            if (this.state.ccIDS.length > 0) {
              if (this.state.ccIDS[0] == ReturnById) {
                listname = 'CApprovalsChecklist';
              }
            }
          }
          let RefCount = 0;
          if (this.state.ReferredCasesLastCount > 0) {
            statusno = 4;
            StatusN = "Referred";
            listname = "RApprovalsChecklist";
            RefCount = 3;
          } else {

          }
          let ApproverId = 0;
          let createdId = 0;
          web.lists.getByTitle(listname).items.select('Title', 'ID', 'Seq', 'Approver/ID', 'Approver/Title', 'Author/ID', 'Author/Title').expand('Approver', 'Author').filter("Title eq '" + seqno + "' and Status eq 'Referred'").orderBy('ID', false).get().then(refbackapp => {
            debugger;

            if (refbackapp.length > 0) {
              ApproverId = refbackapp[0].Approver.ID;
              createdId = refbackapp[0].Author.ID;
            }

          }).then(() => {
            web.lists.getByTitle("ChecklistNote").items.getById(uid).update({
              Migrate: "",
              Comments: Comments,
              CurApproverId: ApproverId,
              NotifyId: ApproverId,
              Status: StatusN,
              StatusNo: statusno,
              RefCount: RefCount,
              ReferredById: createdId,
              ReferredToId: ApproverId,
              WorkflowFlag: "",
            }).then(() => {
              pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.select('Title', 'ID', 'PID').filter("PID eq " + uid).get().then(r => {
                let Approverid = r[0].ID;
                pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.getById(Approverid).update({
                  CurApproverId: ApproverId,
                  Status: StatusN,
                  StatusNo: statusno
                }).then(() => {
                  web.lists.getByTitle("RApprovalsChecklist").items.select('Title', 'ID', 'Seq', 'Approver/ID').expand('Approver').filter("Title eq '" + seqno + "' and Approver/ID eq " + userid + " and Status eq 'Pending'").get().then(rapp => {
                    let Refid = rapp[0].ID;
                    web.lists.getByTitle('RApprovalsChecklist').items.getById(Refid).update({
                      // ApproverId:ReturnById,
                      Status: "Referred Back",
                      Comments: "Referred Back",
                    }).then(() => {
                      let statuslog = 'Referred Back';
                      web.lists.getByTitle(listname).items.select('Title', 'ID', 'Seq', 'Approver/ID').expand('Approver').filter("Title eq '" + seqno + "' and Status eq 'Referred'").orderBy('ID', false).get().then(refbackapp => {
                        let Refbackid = refbackapp[0].ID;
                        web.lists.getByTitle(listname).items.getById(Refbackid).update({
                          // ApproverId:ReturnById,
                          Status: "Pending",
                          Comments: "Pending",
                        }).then(() => {

                          let Notifstatus = statusno.toString() + "-Referred Back";
                          this.AddWFHistory(statuslog).then(() => {
                            this.AddNotesNotifications(Notifstatus, ApproverId).then(() => {
                              this.dummyHistory().then(() => {
                                this.redirect();
                              });
                            });
                          });
                        });
                      });
                    });

                  });
                });
              });

            });

          });


        });
      });

    } // If for Comments Validation
    else {
      this._onClosePanel();
      alert('Comments are mandatory, it cannot contain more than 2000 chars!');
      //document.getElementById("txtAppComments").focus();
      const txtAppComments = document.getElementById("txtAppComments");
      if (txtAppComments) {
          txtAppComments.focus();
      }
    }


  }
  /*--End--*/

  /*--returned logic--*/
  private returned() {
    debugger;
    let web = new Web('Main');
    // let Comments = String(jQuery('#txtAppComments').val()).trim();
    let Comments: string = '';
    let Comments2 = document.getElementById("txtAppComments") as HTMLInputElement | null;
    if (Comments2) {
        Comments = Comments2.value.trim();
    }
    let Controller = this.state.ccIDS;
    if (confirm('Are you sure you want to return the Note?')) {
      if (Comments.length > 0 && Comments.length < 2000) {
        let listname = 'FApprovalsChecklist';
        let curStatusNo = this.state.statusno;
        if (curStatusNo == 1) {
          listname = 'ApprovalsChecklist';
        }
        else if (Controller.length > 0 && curStatusNo == 6) {
          if (this.state.UserID == Controller[0]) {
            listname = 'CApprovalsChecklist';
          }

        }
        else {
          listname = 'FApprovalsChecklist';
        }
        let ReturnedTo = jQuery('#ddlReturnTo option:selected').text();
        if (ReturnedTo == 'Select') {
          alert('Please select the name!');
          jQuery('#ddlReturnTo').focus();
          return false;
        }
        else {
          this._onClosePanel();
          this.on();
          let ReturnVal = 'i:0#.f|membership|' + jQuery('#ddlReturnTo option:selected').val();
          let UserID = 0;
          let req = this.state.UserID;
          console.log(UserID);
          pnp.sp.web.siteUsers.getByLoginName(ReturnVal).get().then(result => {
            console.log(result);
            UserID = result.Id;
          }).then(() => {
            debugger;
            let seqno = this.state.seqno;
            let userID = this.state.UserID;
            let UserEmail = this.state.UserEmail;
            const query = window.location.search.split('uid=')[1];
            let uid = 0;
            if (query != undefined) { uid = parseInt(query); }
            web.lists.getByTitle("ChecklistNote").items.getById(uid).update({
              WorkflowFlag: "Triggered",
              Migrate: "",
              Comments: Comments,
              ReturnedById: req,
              CurApproverId: UserID,
              NotifyId: UserID,
              ReturnedToId: UserID,
              Status: "Returned",
              StatusNo: 3
            }).then(() => {
              pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.select('Title', 'ID', 'PID').filter("PID eq " + uid).get().then(r => {
                let Approverid = r[0].ID;
                pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.getById(Approverid).update({
                  CurApproverId: UserID,
                  Status: "Returned",
                }).then(() => {
                  // web.lists.getByTitle(listname).items.select('Title', 'ID', 'AppEmail').filter("Title eq '" + seqno + "' and AppEmail eq '" + UserEmail + "'").get().then(AppID => {
                    web.lists.getByTitle(listname).items.select("ID,Title,AppName,AppEmail,Approver/ID,Approver/EMail").expand("Approver").filter("Title eq '" + seqno + "' and AppEmail eq '" + userID + "'").get().then(AppID => {
                    let NoteAppID = AppID[0].ID;
                    web.lists.getByTitle(listname).items.getById(NoteAppID).update({
                      Status: 'Returned'
                    }).then(() => {
                      this.AddAppComments().then(() => {
                        let statuslog = 'Returned';
                        let Notifstatus = "3-Returned";
                        this.AddWFHistory(statuslog).then(() => {
                          this.AddNotesNotifications(Notifstatus, UserID).then(() => {
                            this.dummyHistory().then(() => {
                              this.redirect();
                            });
                          });
                        });
                      });
                    });
                  });
                });
              });
            });
          });
        }
      }
      else {
        this._onClosePanel();
        alert('Comments are mandatory, it cannot contain more than 2000 chars!');
        // document.getElementById("txtAppComments").focus();
        const txtAppComments = document.getElementById("txtAppComments");
        if (txtAppComments) {
            txtAppComments.focus();
        }
      }
    }
  }
  /*--End--*/

  /*--return back logic--*/
  private returnBack() {
    debugger;
    let web = new Web('Main');
    // let Comments = String(jQuery('#txtAppComments').val()).trim();
    let Comments: string = '';
    let Comments2 = document.getElementById("txtAppComments") as HTMLInputElement | null;
    if (Comments2) {
        Comments = Comments2.value.trim();
    }

    if (Comments.length > 0 && Comments.length < 2000) {
      let listname = 'FApprovalsChecklist';
      this._onClosePanel();
      this.on();
      let ReturnById = this.state.ReturnedByID;
      let req = this.state.ReqID;
      let curruser = this.state.UserID;
      console.log(ReturnById);

      debugger;
      let seqno = this.state.seqno;
      let UserEmail = this.state.UserEmail;
      const query = window.location.search.split('uid=')[1];
      let uid = 0;
      if (query != undefined) { uid = parseInt(query); }
      this.AddAppComments().then(() => {
        this.checkRefApprover(ReturnById).then((App) => {
          let statusno = 6;
          if (App == 'Approver') { statusno = 1; listname = 'ApprovalsChecklist'; }
          else {
            if (this.state.ccIDS.length > 0) {
              if (this.state.ccIDS[0] == ReturnById) { listname = 'CApprovalsChecklist'; }
            }
          }

          web.lists.getByTitle("ChecklistNote").items.getById(uid).update({
            Migrate: "",
            Comments: Comments,
            CurApproverId: ReturnById,
            NotifyId: ReturnById,
            Status: "Returned Back",
            StatusNo: statusno
          }).then(() => {
            pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.select('Title', 'ID', 'PID').filter("PID eq " + uid).get().then(r => {
              let Approverid = r[0].ID;
              pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.getById(Approverid).update({
                CurApproverId: ReturnById,
                Status: "Returned Back",
                StatusNo: statusno
              }).then(() => {
                web.lists.getByTitle(listname).items.select('Title', 'ID', 'Seq', 'Approver/ID').expand('Approver').filter("Title eq '" + seqno + "' and Approver/ID eq " + ReturnById).get().then(rapp => {
                  let Refid = rapp[0].ID;
                  web.lists.getByTitle(listname).items.getById(Refid).update({
                    Status: "Pending",
                    Comments: "Pending",
                  }).then(() => {
                    if (curruser != req) {
                      web.lists.getByTitle('ApprovalsChecklist').items.select('Title', 'ID', 'Seq', 'Approver/ID').expand('Approver').filter("Title eq '" + seqno + "' and Approver/ID eq " + curruser).get().then(retapp => {
                        let Returnid = retapp[0].ID;
                        web.lists.getByTitle('ApprovalsChecklist').items.getById(Returnid).update({
                          Status: "Returned Back",
                          Comments: "Returned Back",
                        }).then(() => {
                          let statuslog = 'Returned Back';
                          let Notifstatus = statusno.toString() + "-Returned Back";
                          this.AddWFHistory(statuslog).then(() => {
                            this.AddNotesNotifications(Notifstatus, ReturnById).then(() => {
                              this.dummyHistory().then(() => {
                                this.redirect();
                              });
                            });
                          });
                        });
                      });

                    } else {
                      let statuslog = 'Returned Back';
                      let Notifstatus = statusno.toString() + "-Returned Back";
                      this.AddWFHistory(statuslog).then(() => {
                        this.AddNotesNotifications(Notifstatus, ReturnById).then(() => {
                          this.dummyHistory().then(() => {
                            this.redirect();
                          });
                        });
                      });
                    }

                  });
                });
              });
            });

          });

        });
      });

    }
    else {
      this._onClosePanel();
      alert('Comments are mandatory, it cannot contain more than 2000 chars!');
      // document.getElementById("txtAppComments").focus();
      const txtAppComments = document.getElementById("txtAppComments");
      if (txtAppComments) {
          txtAppComments.focus();
      }
    }


  }
  /*--End--*/

  /*--change approver functionality--*/
  private ChangeApprover() {
    debugger;
    let web = new Web('Main');
    // let Comments = String(jQuery('#txtAppComments').val()).trim();
    let Comments: string = '';
    let Comments2 = document.getElementById("txtAppComments") as HTMLInputElement | null;
    if (Comments2) {
        Comments = Comments2.value.trim();
    }
    if (Comments.length > 0 && Comments.length < 2000) {
      let listname = 'ApprovalsChecklist';
      let curStatusNo = this.state.statusno;
      let NewApproverID = this.state.userManagerIDs;
      let NewApproverName = this.state.MgrName;
      let NewApproverEmail = this.state.ManagerEmail[0];
      let Approvers = this.state.AllApprovers;

      let currUser = this.state.UserID;
      let ReqID = this.state.ReqID;
      let CurrApproverEmail = this.state.CurrApproverEmail.toLowerCase();
      let ccEmail = this.state.ccEmail.toLowerCase();

      if (curStatusNo == 1) {
        listname = 'ApprovalsChecklist';
      }
      else if (curStatusNo == 4) {
        listname = 'RApprovalsChecklist';
      }
      else if (curStatusNo == 6 && ccEmail == CurrApproverEmail) {
        listname = 'CApprovalsChecklist';
      }
      else {
        listname = 'FApprovalsChecklist';
      }

      debugger;
      if (this.state.MgrName == '') {
        alert('Kindly select username!');
        jQuery('input[aria-label="People Picker"]').focus();
        return;
      }
      else if (ReqID == NewApproverID[0]) {
        alert('Requester cannot be approver!');
        jQuery('input[aria-label="People Picker"]').focus();
        return;
      }

      else {
        if (curStatusNo != 4) {
          Approvers.push(NewApproverID[0]);
          let CurApproverID = this.state.CurrAppID;
          Approvers = jQuery.grep(Approvers, (value) => {
            return value != CurApproverID;
          });
        }
        this.checkApprover(NewApproverEmail).then((len) => {
          if (len == 0) {
            this.checkRecommender(NewApproverEmail).then((Rlen) => {
              if (Rlen == 0) {
                this._onClosePanel();
                this.on();
                let seqno = this.state.seqno;
                let UserEmail = this.state.UserEmail;
                const query = window.location.search.split('uid=')[1];
                let uid = 0;
                if (query != undefined) { uid = parseInt(query); }
                this.AddAppComments().then(() => {
                  pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.select('Title', 'ID', 'PID').filter("PID eq " + uid).get().then(r => {
                    let Approverid = r[0].ID;
                    pnp.sp.site.rootWeb.lists.getByTitle("ChecklistNote").items.getById(Approverid).update({
                      CurApproverId: NewApproverID[0],
                      Status: "Submitted to " + NewApproverName,
                      CurApproverTxt: NewApproverEmail
                    }).then(() => {
                      web.lists.getByTitle(listname).items.select('Title', 'ID', 'AppEmail', 'Approver/EMail').expand('Approver').filter("Title eq '" + seqno + "' and Approver/EMail eq '" + CurrApproverEmail + "'").get().then(AppID => {
                        let NoteAppID = AppID[0].ID;
                        web.lists.getByTitle(listname).items.getById(NoteAppID).update({
                          Status: 'Pending',
                          ApproverId: NewApproverID[0],
                          AppEmail: NewApproverEmail,
                          AppID: NewApproverID[0],
                          AppName: NewApproverName
                        }).then(() => {
                          if (curStatusNo == 6 && ccEmail == CurrApproverEmail) {
                            web.lists.getByTitle("ChecklistNote").items.getById(uid).update({
                              // Description: this.state.description,
                              Migrate: "",
                              ApproversId: { results: Approvers },
                              CurApproverId: NewApproverID[0],
                              Comments: Comments,
                              ControllerId: NewApproverID[0],
                              NotifyId: NewApproverID[0],
                              Status: "Submitted to " + NewApproverName,
                            }).then(() => {
                              let statuslog = 'ChangeApprover';
                              let Notifstatus = curStatusNo.toString() + "-Submitted to " + NewApproverName;
                              this.AddWFHistory(statuslog).then(() => {
                                this.AddNotesNotifications(Notifstatus, NewApproverID[0]).then(() => {
                                  this.dummyHistory().then(() => {

                                    this.redirect();
                                  });
                                });
                              });
                            });
                          } // If Condition for Controller
                          else if (curStatusNo == 4) {
                            web.lists.getByTitle("ChecklistNote").items.getById(uid).update({
                              Migrate: "",
                              // ApproversId:{results:Approvers},
                              CurApproverId: NewApproverID[0],
                              Comments: Comments,
                              ReferredToId: NewApproverID[0],
                              NotifyId: NewApproverID[0],
                              Status: "Referred to " + NewApproverName,
                            }).then(() => {
                              let statuslog = 'ChangeApprover';
                              let Notifstatus = curStatusNo.toString() + "-Referred to " + NewApproverName;
                              this.AddWFHistory(statuslog).then(() => {
                                this.AddNotesNotifications(Notifstatus, NewApproverID[0]).then(() => {
                                  this.dummyHistory().then(() => {

                                    this.redirect();
                                  });
                                });
                              });
                            });
                          } // If Condition for Referree
                          else {
                            web.lists.getByTitle("ChecklistNote").items.getById(uid).update({
                              // Description: this.state.description,
                              Migrate: "",
                              ApproversId: { results: Approvers },
                              CurApproverId: NewApproverID[0],
                              Comments: Comments,
                              NotifyId: NewApproverID[0],
                              Status: "Submitted to " + NewApproverName,
                            }).then(() => {
                              let statuslog = 'ChangeApprover';
                              let Notifstatus = curStatusNo.toString() + "-Submitted to " + NewApproverName;
                              this.AddWFHistory(statuslog).then(() => {
                                this.AddNotesNotifications(Notifstatus, NewApproverID[0]).then(() => {
                                  this.dummyHistory().then(() => {

                                    this.redirect();
                                  });
                                });
                              });
                            });
                          } // Else Condition for Controller

                        });
                      });
                    });
                  });
                });
              }
              else {
                alert('Recommender has already been added!');
                jQuery('input[aria-label="People Picker"]').focus();
                return;
              }
            });
          }
          else {
            alert('Approver has already been added!');
            jQuery('input[aria-label="People Picker"]').focus();
            return;

          }

        });

      }
    }
    else {

      alert('Comments are mandatory, it cannot contain more than 2000 chars!');
      //document.getElementById("txtAppComments").focus();
      const txtAppComments = document.getElementById("txtAppComments");
      if (txtAppComments) {
          txtAppComments.focus();
      }
    }

  }
  /*--End--*/

  /*--save comments in  Commentslog list--*/
  private AddAppComments(): Promise<any[]> {
    debugger;

    let SeqNo = this.state.seqno;
    // let comment = String(jQuery('#txtAppComments').val()).trim();
    let Comments: string = '';
    let Comments2 = document.getElementById("txtAppComments") as HTMLInputElement | null;
    if (Comments2) {
        Comments = Comments2.value.trim();
    }
    debugger;
    let web = new Web('Main');
    return web.lists.getByTitle("CommentsLog").items.add({
      Title: this.state.seqno,
      Page: '0',
      Docref: '0',
      Comments: Comments,
      Appname: this.state.name,
      Appemail: this.state.UserEmail
    }).then((iar: ItemAddResult) => {
      console.log(iar.data.ID);
      return Promise.resolve(['Done']);
    });

  }
  /*--End--*/

  /*--Save workflow history in wfhistory list--*/
  private AddWFHistory(status: string): Promise<any[]> {
    debugger;
    let web = new Web('Main');
    let dt = new Date();
    let mnth = (dt.getMonth() + 1).toString();
    let dat = dt.getDate().toString();
    let hrs = dt.getHours().toString();
    let mins = dt.getMinutes().toString();
    let secs = dt.getSeconds().toString();
    if (mnth.length == 1) { mnth = '0' + mnth; } if (dat.length == 1) { dat = '0' + dat; } if (hrs.length == 1) { hrs = '0' + hrs; } if (mins.length == 1) { mins = '0' + mins; } if (secs.length == 1) { secs = '0' + secs; }
    let createDate = dat + "-" + mnth + "-" + dt.getFullYear() + " " + hrs + ":" + mins + ":" + secs;

    // let dat= dt.getDate()+"-"+((dt.getMonth())+1)+"-"+dt.getFullYear()+" "+dt.getHours()+":"+dt.getMinutes()+":"+dt.getSeconds();
    let log = '';
    //if(status=='Submitted'){
    //  log='Submitted to '+this.state.MgrName+' by '+this.state.name+' on '+createDate;
    //}
    if (status == 'Approved' || status == 'Submitted') {
      log = 'Approved by ' + this.state.name + ' on ' + createDate;
    }
    else if (status == 'Rejected') {
      log = 'Rejected by ' + this.state.name + ' on ' + createDate;
    }
    else if (status == 'Cancelled') {
      log = 'Cancelled by ' + this.state.name + ' on ' + createDate;
    }

    else if (status == 'Returned') {
      let ReturnVal = jQuery('#ddlReturnTo option:selected').text();
      log = 'Returned by ' + this.state.name + ' to ' + ReturnVal + ' on ' + createDate;
    }
    else if (status == 'Returned Back') {
      var ReturnedTo = this.state.ReturnedByName;
      log = 'Returned back by ' + this.state.name + ' to ' + ReturnedTo + ' on ' + createDate;
    }
    else if (status == 'Referred') {
      var ReferredTo = this.state.MgrName;
      log = 'Referred by ' + this.state.name + ' to ' + ReferredTo + ' on ' + createDate;
    }
    else if (status == 'Referred Back') {
      var ReferredBy = this.state.ReferredByName;
      log = 'Referred back by ' + this.state.name + ' to ' + ReferredBy + ' on ' + createDate;
    }
    else if (status == 'ChangeApprover') {
      let NewApp = this.state.MgrName;
      // let OldApp = document.getElementById("tdCurrApprover").innerText;
      let OldApp2 = document.getElementById("tdCurrApprover");
      let OldApp : string = '';
      if(OldApp2){OldApp = OldApp2.innerText}
      log = 'Approver changed from ' + OldApp + ' to ' + NewApp + ' by ' + this.state.name + ' on ' + createDate;
    }
    else if (status == 'Called Back') {
      log = 'Called back by ' + this.state.name + ' on ' + createDate;
    }
    else {
      log = 'Approved by ' + this.state.name + ' on ' + createDate;
    }

    // let seqno = document.getElementById("divSeqNo").innerText;
    let seqno2 = document.getElementById("divSeqNo");
    let seqno : string = '';
    if(seqno2){seqno = seqno2.innerText}
    
    return web.lists.getByTitle("WFHistory").items.add({
      Title: seqno,
      AuditLog: log,
      Currapprover: this.state.name,
      FormName: 'Note',
      ActionDateTime: createDate
    }).then((iar: ItemAddResult) => {
      console.log('History Log Created!');
      return Promise.resolve(['Done']);

    });

  }

  /*--End--*/

  /*--Dummy Call to retrieve History--*/
  public dummyHistory(): Promise<any[]> {
    let title = this.state.seqno;
    let data = [];
    let web = new Web('Main');
    return web.lists.getByTitle('WFHistory').items.select("ID,Title,AuditLog").filter("Title eq '" + title + "'").orderBy("Created", false).getAll().then((items: any[]) => {
      debugger;
      return Promise.resolve(items);

    });

  }
  /*--End--*/
  /*--retrieve workflow history--*/
  public retrieveHistory() {
    let web = new Web('Main');
    let title = this.state.seqno;
    // let data = [];
    let data: any[] = [];

    web.lists.getByTitle('WFHistory').items.select("ID,Title,AuditLog,Modified").filter("Title eq '" + title + "'").orderBy("Modified", false).get().then((items: any[]) => {
      debugger;
      let tbldata = '';

      for (let i = 0; i < items.length; i++) {
        data.push(<tr><td>{items[i].AuditLog}</td></tr>);
      }


    }).then(() => {
      this.setState({ WFHistoryLog: data });
    });

  }
  /*--End--*/
  /*--Redirect to Home Page--*/
  private gotoHomePage(): void {
    window.location.replace(this.state.Sitename);
  }

  /*--End--*/

  /*--Delete attachments for Annexures--*/
  public DeleteAttachment(vals : string): void {
    debugger;
    let web = new Web('Main');
    this.setState({ AppAttachments: [] });
    let sitename = this.state.Sitename.split(".com")[1];
    let url = sitename + '/Main/NoteAnnexures/' + vals;
    let fldr = vals.split("/")[0];
    let fldURL = sitename + '/Main/NoteAnnexures/' + fldr;
    web.getFileByServerRelativeUrl(url).recycle().then(data => {
      console.log("File Deleted " + vals);
      let userid = this.state.UserID;
      web.getFolderByServerRelativeUrl(fldURL).files.select('*,listItemAllFields').expand('listItemAllFields').get().then((result) => {
        console.log(result);
        let links: any[] = [];
        for (let i = 0; i < result.length; i++) {
          if (result[i].ListItemAllFields.AuthorId == userid) {
            links.push(fldr + "/" + result[i].Name);
          }
        }

        this.setState({ AppAttachments: links });
        // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
      });



    });

  }
  /*--End--*/

  /*--Add attachments for annexures--*/
  public AttachLib = (event : any) => {
    debugger;
    let web = new Web('Main');
    var uploadFlag = true;
    //in case of multiple files,iterate or else upload the first file.
    // let file = fileUpload.files[0];
    let file = event.target.files[0];
    let filesize = file.size / 1048576;
    var n = (file.name.length - file.name.lastIndexOf("."));
    //let fileExtn=file.name.substr(file.name.length-(n-1)).toLowerCase();
    let fileExtn = file.name.substr((file.name.lastIndexOf('.') + 1)).toLowerCase();
    let fileSplit = file.name.split(".");
    let PermissibleExtns = ['png', 'jpeg', 'jpg', 'gif', 'pdf', 'doc', 'docx', 'xls', 'xlsx', 'eml'];
    let fileTest = file.name.substring(0, (file.name.length - n));
    console.log(fileTest);
    let match = new RegExp('[~#%\&{}+.\|]|\\.\\.|^\\.|\\.$').test(fileTest);

    if (fileSplit.length > 2) {
      alert('Alert-Selected file double extension is not allowed!');
      // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
      return false;
    }
    else if (PermissibleExtns.indexOf(fileExtn) == -1) {
      alert('Alert-Selected file type is not allowed!');
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
    else if ((fileExtn == 'mp4' && filesize > 10) || (fileExtn != 'mp4' && filesize > 10)) {
      alert('Alert-File size is more than permissible limit (Max 5MB is allowed) !');
      // document.getElementById("fileUploadInput").nodeValue = null;
        let ddlDepartment = document.getElementById('fileUploadInput');
        if (ddlDepartment) {
          ddlDepartment.nodeValue = null;
        }
      return false;
    }
    else {
      if (file != undefined || file != null) {
        let SeqNo = this.state.seqno;
        web.folders.getByName('NoteAnnexures').folders.add(SeqNo).then(data => {
          console.log("Folder is created at " + data.data.ServerRelativeUrl);
          //assuming that the name of document library is Documents, change as per your requirement, 
          //this will add the file in root folder of the document library, if you have a folder named test, replace it as "/Documents/test"

          web.getFolderByServerRelativeUrl("NoteAnnexures/" + SeqNo).files.add(file.name, file, true).then((result) => {
            console.log(file.name + " uploaded successfully!");
            let links: any[] = [];
            links = this.state.AppAttachments;
            links.push(SeqNo + "/" + file.name);
            console.log(links);
            this.setState({ AppAttachments: links });
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
    // return uploadFlag;
  }

  /*--End--*/
  /*--For Expand(+)--*/
  private Expand(str: string) {
    let expand = str + 'Expand';
    let collapse = str + 'Collapse';
    const expandElement = document.getElementById(expand);
    const collapseElement = document.getElementById(collapse);
    const divContent = document.getElementById('divContent');
    const divFrame = document.getElementById('divFrame');
    // document.getElementById(expand).style.display = 'block';
    // document.getElementById(collapse).style.display = 'none';
    // let ht = document.getElementById('divContent').clientHeight + 20;
    // console.log(document.getElementById('divFrame').clientHeight);
    // console.log(ht);
    // document.getElementById('divFrame').style.height = ht + "px";    

    var ht : Number = 0;
    if(collapseElement){collapseElement.style.display = 'none';};
    if(expandElement){expandElement.style.display = 'block';}
    if(divContent){ht = divContent.clientHeight + 20;};
    if(divFrame){divFrame.style.height = ht + "px";}
  }

  /*--End--*/

  /*--For collapse(-)--*/
  private Collapse(str: string) {
    let expand = str + 'Expand';
    let collapse = str + 'Collapse';
    const expandElement = document.getElementById(expand);
    const collapseElement = document.getElementById(collapse);
    const divContent = document.getElementById('divContent');
    const divFrame = document.getElementById('divFrame');
    // document.getElementById(collapse).style.display = 'block';
    // document.getElementById(expand).style.display = 'none';
    //let ht = document.getElementById('divContent').clientHeight + 20;
    // console.log(document.getElementById('divFrame').clientHeight);
    // console.log(ht);
    // document.getElementById('divFrame').style.height = ht + "px";
    var ht : Number = 0;
    if(collapseElement){collapseElement.style.display = 'block';};
    if(expandElement){expandElement.style.display = 'none';}
    if(divContent){ht = divContent.clientHeight + 20;};
    if(divFrame){divFrame.style.height = ht + "px";}
  }  

  /*--End--*/
  /*--to show seek information--*/
  private showDiv(str: string) {
    let ApproverVal = jQuery('#ddlApprover option:selected').val();
    let Recomm = this.state.RecomNewselectedItems.length;
    if (Recomm > 0) {
      $("#ddlRefer").val('No');
      $("#ddlReturn").val('No');
      $("#ddlApprover").val('Yes');
      alert('Kindly remove Recommender/s');
      return false;
    }
    else {
      if (str == 'Refer') {
        let ReferVal = jQuery('#ddlRefer option:selected').val();
        if (ReferVal == 'Yes') {
          // document.getElementById("divRefer").style.display = 'block';
          // document.getElementById("divReturn").style.display = 'none';
          // document.getElementById("btnRefer").style.display = 'block';
          // document.getElementById("btnApprove").style.display = 'none';
          // document.getElementById("btnReturn").style.display = 'none';
          // document.getElementById("btnCancel").style.display = 'none';
          // document.getElementById("divAddRecomm").style.display = 'none';
          const divRefer = document.getElementById("divRefer");
          if (divRefer) {
              divRefer.style.display = 'block';
          }

          const divReturn = document.getElementById("divReturn");
          if (divReturn) {
              divReturn.style.display = 'none';
          }

          const btnRefer = document.getElementById("btnRefer");
          if (btnRefer) {
              btnRefer.style.display = 'block';
          }

          const btnApprove = document.getElementById("btnApprove");
          if (btnApprove) {
              btnApprove.style.display = 'none';
          }

          const btnReturn = document.getElementById("btnReturn");
          if (btnReturn) {
              btnReturn.style.display = 'none';
          }

          const btnCancel = document.getElementById("btnCancel");
          if (btnCancel) {
              btnCancel.style.display = 'none';
          }

          const divAddRecomm = document.getElementById("divAddRecomm");
          if (divAddRecomm) {
              divAddRecomm.style.display = 'none';
          }

        }
        else {
          // document.getElementById("divRefer").style.display = 'none';
          // document.getElementById("btnRefer").style.display = 'none';
          // document.getElementById("btnApprove").style.display = 'block';
          // document.getElementById("btnReturn").style.display = 'none';
          // document.getElementById("btnCancel").style.display = 'block';
          const divRefer = document.getElementById("divRefer");
          if (divRefer) {
              divRefer.style.display = 'none';
          }

          const btnRefer = document.getElementById("btnRefer");
          if (btnRefer) {
              btnRefer.style.display = 'none';
          }

          const btnApprove = document.getElementById("btnApprove");
          if (btnApprove) {
              btnApprove.style.display = 'block';
          }

          const btnReturn = document.getElementById("btnReturn");
          if (btnReturn) {
              btnReturn.style.display = 'none';
          }

          const btnCancel = document.getElementById("btnCancel");
          if (btnCancel) {
              btnCancel.style.display = 'block';
          }
        }
      }
      else if (str == 'Recomm') {
        // document.getElementById("divRefer").style.display = 'none';
        // document.getElementById("divReturn").style.display = 'none';
        // document.getElementById("divAddRecomm").style.display = 'block';
        // document.getElementById("divRefer").style.display = 'none';
        // document.getElementById("divReturn").style.display = 'none';
        // document.getElementById("btnRefer").style.display = 'none';
        // document.getElementById("btnApprove").style.display = 'block';
        // document.getElementById("btnReturn").style.display = 'none';
        // document.getElementById("btnCancel").style.display = 'block';

        const divRefer = document.getElementById("divRefer");
        if (divRefer) {
            divRefer.style.display = 'none';
        }

        const divReturn = document.getElementById("divReturn");
        if (divReturn) {
            divReturn.style.display = 'none';
        }

        const divAddRecomm = document.getElementById("divAddRecomm");
        if (divAddRecomm) {
            divAddRecomm.style.display = 'block';
        }

        const btnRefer = document.getElementById("btnRefer");
        if (btnRefer) {
            btnRefer.style.display = 'none';
        }

        const btnApprove = document.getElementById("btnApprove");
        if (btnApprove) {
            btnApprove.style.display = 'block';
        }

        const btnReturn = document.getElementById("btnReturn");
        if (btnReturn) {
            btnReturn.style.display = 'none';
        }

        const btnCancel = document.getElementById("btnCancel");
        if (btnCancel) {
            btnCancel.style.display = 'block';
        }

      }
      else {
        let ReturnVal = jQuery('#ddlReturn option:selected').val();
        if (ReturnVal == 'Yes') {
          // document.getElementById("divRefer").style.display = 'none';
          // document.getElementById("divReturn").style.display = 'block';
          // document.getElementById("btnRefer").style.display = 'none';
          // document.getElementById("btnApprove").style.display = 'none';
          // document.getElementById("btnReturn").style.display = 'block';
          // document.getElementById("btnCancel").style.display = 'none';
          // document.getElementById("divAddRecomm").style.display = 'none';

          const divRefer = document.getElementById("divRefer");
          if (divRefer) {
              divRefer.style.display = 'none';
          }

          const divReturn = document.getElementById("divReturn");
          if (divReturn) {
              divReturn.style.display = 'block';
          }

          const btnRefer = document.getElementById("btnRefer");
          if (btnRefer) {
              btnRefer.style.display = 'none';
          }

          const btnApprove = document.getElementById("btnApprove");
          if (btnApprove) {
              btnApprove.style.display = 'none';
          }

          const btnReturn = document.getElementById("btnReturn");
          if (btnReturn) {
              btnReturn.style.display = 'block';
          }

          const btnCancel = document.getElementById("btnCancel");
          if (btnCancel) {
              btnCancel.style.display = 'none';
          }

          const divAddRecomm = document.getElementById("divAddRecomm");
          if (divAddRecomm) {
              divAddRecomm.style.display = 'none';
          }
        }
        else {
          // document.getElementById("divReturn").style.display = 'none';
          // document.getElementById("btnRefer").style.display = 'none';
          // document.getElementById("btnApprove").style.display = 'block';
          // document.getElementById("btnReturn").style.display = 'none';
          // document.getElementById("btnCancel").style.display = 'block';
          // document.getElementById("divAddRecomm").style.display = 'none';

          const divReturn = document.getElementById("divReturn");
          if (divReturn) {
              divReturn.style.display = 'none';
          }

          const btnRefer = document.getElementById("btnRefer");
          if (btnRefer) {
              btnRefer.style.display = 'none';
          }

          const btnApprove = document.getElementById("btnApprove");
          if (btnApprove) {
              btnApprove.style.display = 'block';
          }

          const btnReturn = document.getElementById("btnReturn");
          if (btnReturn) {
              btnReturn.style.display = 'none';
          }

          const btnCancel = document.getElementById("btnCancel");
          if (btnCancel) {
              btnCancel.style.display = 'block';
          }

          const divAddRecomm = document.getElementById("divAddRecomm");
          if (divAddRecomm) {
              divAddRecomm.style.display = 'none';
          }
        }
      }
    }
    debugger;
    if (this.state.statusno == 4) {
      debugger;
      const btnApprove = document.getElementById("btnApprove");
      if (btnApprove) {
          btnApprove.style.display = 'none';
      }

      const btnCancel = document.getElementById("btnCancel");
      if (btnCancel) {
          btnCancel.style.display = 'none';
      }

      const btnReferBack = document.getElementById("btnReferBack");
      if (btnReferBack) {
          if (jQuery('#ddlRefer option:selected').val() == "Yes") {
              btnReferBack.style.display = 'none';
          } else if (jQuery('#ddlRefer option:selected').val() == "No") {
              btnReferBack.style.display = 'block';
          }
      }
    }

  }
  /*--End--*/

  /*--For on(show) and off(hide) please wait overlay while page load--*/
    // private on() {
    //   let ht = window.innerHeight;
    //   document.getElementById('overlay').style.height = ht.toString() + "px";
    //   document.getElementById("overlay").style.display = "block";
    // }
    // private off() {
    //   document.getElementById("overlay").style.display = "none";
    // }

  private on() {
    let ht = window.innerHeight;
    const overlay = document.getElementById('overlay');
    if (overlay) {
        overlay.style.height = ht.toString() + "px";
        overlay.style.display = "block";
    }
  }
  
  private off() {
      const overlay = document.getElementById("overlay");
      if (overlay) {
          overlay.style.display = "none";
      }
  }
  
  /*-- --*/


  /*--To check adding persons already approver or not--*/
  private checkApprover(appemail: string): Promise<number> {
    debugger;
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

  /*--To check adding persons already recommander or not*/
  private checkRecommender(appemail: string): Promise<number> {
    debugger;
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

  /*--To check adding persons already referrer or not*/
  private checkRefApprover(appID: number): Promise<string> {
    debugger;
    let title = this.state.seqno;
    let App = '';
    let web = new Web('Main');
    return web.lists.getByTitle('ApprovalsChecklist').items.select("ID,Title,AppName,AppEmail,Approver/ID").expand('Approver').filter("Title eq '" + title + "' and Approver/ID eq " + appID).orderBy("Seq asc").getAll().then((items: any[]) => {

      if (items.length > 0) {
        App = 'Approver';
      }

      return Promise.resolve(App);
    });

  }
  // End Function 

  /*--Retrieve all referer details--*/
  private retrieveRefers() {
    let title = this.state.seqno;
    // let data = [];
    let data: any[] = []; 
    let web = new Web('Main');
    web.lists.getByTitle('RApprovalsChecklist').items.select("ID,Title,AppName,Status,AppEmail,Created,Modified,Author/Title,Seq").expand('Author').filter("Title eq '" + title + "'").orderBy("Seq asc").getAll().then((items: any[]) => {
      debugger;

      for (let i = 0; i < items.length; i++) {

        let createDate = 'NA';
        if (items[i].Status.trim() != 'Pending') {
          let dt = new Date(items[i].Modified);
          let mnth = (dt.getMonth() + 1).toString();
          let dat = dt.getDate().toString();
          let hrs = dt.getHours().toString();
          let mins = dt.getMinutes().toString();
          if (mnth.length == 1) { mnth = '0' + mnth; } if (dat.length == 1) { dat = '0' + dat; } if (hrs.length == 1) { hrs = '0' + hrs; } if (mins.length == 1) { mins = '0' + mins; }
          createDate = dat + "-" + mnth + "-" + dt.getFullYear() + " " + hrs + ":" + mins;
        }

        data.push(<tr><td>{i + 1}</td><td>{items[i].Author.Title}</td><td>{items[i].AppName}</td><td>{items[i].Status}</td><td>{createDate}</td></tr>);
      }
    }).then(() => {
      this.setState({ ReferselectedItems: data });
    });
  }
  /*--End--*/

  /*--Retrieve all controller details--*/
  private retrieveController() {
    let title = this.state.seqno;
    // let data = [];
    let data: any[] = []; 
    let web = new Web('Main');
    web.lists.getByTitle('CApprovalsChecklist').items.select("ID,Title,AppName,Status,AppEmail,Created,Modified,Author/Title,Seq").expand('Author').filter("Title eq '" + title + "'").orderBy("Seq asc").getAll().then((items: any[]) => {
      debugger;

      for (let i = 0; i < items.length; i++) {

        let createDate = 'NA';
        if (items[i].Status.trim() != 'Pending') {
          let dt = new Date(items[i].Modified);
          let mnth = (dt.getMonth() + 1).toString();
          let dat = dt.getDate().toString();
          let hrs = dt.getHours().toString();
          let mins = dt.getMinutes().toString();
          if (mnth.length == 1) { mnth = '0' + mnth; } if (dat.length == 1) { dat = '0' + dat; } if (hrs.length == 1) { hrs = '0' + hrs; } if (mins.length == 1) { mins = '0' + mins; }
          createDate = dat + "-" + mnth + "-" + dt.getFullYear() + " " + hrs + ":" + mins;
        }

        data.push(<tr><td>{i + 1}</td><td>{items[i].AppName}</td><td>{items[i].Status}</td><td>{createDate}</td></tr>);
      }

    }).then(() => {
      this.setState({ ControlselectedItems: data });
    });

  }
  // End Function

  // Add Referee before submission
  private AddRefer() {
    debugger;
    let seqno = 1;
    let RequesterID = this.state.ReqID;
    let ReturnVal = this.state.userManagerIDs;
    let ReferEmail = this.state.ManagerEmail;
    let ReferName = this.state.MgrName;
    let currUser = this.state.UserID;
    if (ReturnVal.length == 0) {
      alert('Kindly select username!');
      jQuery('input[aria-label="People Picker"]').focus();
      return;
    }
    else if (ReturnVal[0] == RequesterID) {
      alert('Requester cannot be Referred!');
      jQuery('input[aria-label="People Picker"]').focus();
      return;
    }
    else {

      let mgrEmail = ReferEmail[0];

      this.checkApprover(mgrEmail).then((len) => {
        if (len == 0) {
          this.checkRecommender(mgrEmail).then((len1) => {
            if (len1 == 0) {
              let SeqNo = this.state.seqno;
              let web = new Web('Main');
              debugger;
              web.lists.getByTitle('RApprovalsChecklist').items.add({
                Title: this.state.seqno,
                Status: 'Pending',
                Seq: seqno,
                // LikedById: {results:[this.state.userManagerIDs[0]]},
                // Views: 1,
                ApproverId: ReturnVal[0],
                AppID: ReturnVal[0],
                AppName: ReferName,
                AppEmail: ReferEmail[0]
              }).then((iar: ItemAddResult) => {
                console.log(iar.data.ID);
              });
            }
            else {
              alert('Recommender cannot be Referred!');
              return;
            }
          });

        }

        else {
          alert('Approver cannot be Referred!');
          return;


        }

      });

    }
  }
  //  End Function

  /*-- To Add Mark for Info Recipients--*/
  private AddMarkforInfo() {
    debugger;
    let seqno = 1;
    let MgrID = this.state.MarkIDs;
    let userid = this.state.UserID;
    let AllRecipients = this.state.MarkItems;
    if (this.state.MarkName[0] == '') {
      alert('Kindly select username!');
      jQuery('input[aria-label="People Picker"]').focus();
      return;
    }
    else if (AllRecipients.length == 10) {
      alert('Maximum 10 recipients can be added!');
      jQuery('input[aria-label="People Picker"]').focus();
      return;
    }
    else if (userid == MgrID[0]) {
      alert('Requester cannot be Marked For Info!');
      jQuery('input[aria-label="People Picker"]').focus();
      return;
    }
    else {
      // let title = document.getElementById("tdTitle").innerText;
      const tdTitle = document.getElementById("tdTitle");
      let title = '';
      if (tdTitle) {
          title = tdTitle.innerText;
      }
      let SeqNo = this.state.seqno;
      let web = new Web('Main');
      debugger;
      web.lists.getByTitle('MarkRecipients').items.add({
        Title: title,
        SeqNo: SeqNo,
        RecipientId: this.state.MarkIDs[0],
        RequesterId: this.state.UserID,
      }).then((iar: ItemAddResult) => {
        console.log(iar.data.ID);
        this.retrieveMarkForInfo();
      });



    }

  }
  /*-- Ending Add Mark for Info Recipients--*/

  /*-- To Retrieve Mark for Info Recipients--*/
  private retrieveMarkForInfo() {
    let title = this.state.seqno;
    // let data = [];
    let data: any[] = []; 
    let web = new Web('Main');
    web.lists.getByTitle('MarkRecipients').items.select("ID,Title,Recipient/Title,Requester/Title,Created").expand('Recipient,Requester').filter("SeqNo eq '" + title + "' ").orderBy("ID asc").getAll().then((items: any[]) => {
      debugger;
      if (this.state.ReqID == this.state.UserID) {
        data.push(<tr><th>SNo</th><th>Recipient</th><th>Marked On</th><th>Action</th></tr>);
      }
      else {
        data.push(<tr><th>SNo</th><th>Recipient</th><th>Marked On</th></tr>);
      }

      if (items.length > 0) {

        for (let i = 0; i < items.length; i++) {
          let dt = new Date(items[i].Created);
          let mnth = (dt.getMonth() + 1).toString();
          let dat = dt.getDate().toString();
          let hrs = dt.getHours().toString();
          let mins = dt.getMinutes().toString();
          if (mnth.length == 1) { mnth = '0' + mnth; } if (dat.length == 1) { dat = '0' + dat; } if (hrs.length == 1) { hrs = '0' + hrs; } if (mins.length == 1) { mins = '0' + mins; }
          let createDate = dat + "-" + mnth + "-" + dt.getFullYear() + " " + hrs + ":" + mins;
          if (this.state.ReqID == this.state.UserID) {
            if (items[i].Recipient != undefined) {
              data.push(<tr><td>{i + 1}</td><td>{items[i].Recipient.Title}</td><td>{createDate}</td><td><button onClick={() => { this.DeleteMark(items[i].ID); }}>Delete</button></td></tr>);
            }

          }

          else {
            if (items[i].Recipient != undefined) {
              data.push(<tr><td>{i + 1}</td><td>{items[i].Recipient.Title}</td><td>{createDate}</td></tr>);
            }

          }

        }
      }

    }).then(() => {
      this.setState({ MarkItems: data });
    });
  }
  /*-- Ending Retrieval of Mark for Info Recipients--*/

  /*-- To Delete Mark for Info Recipients--*/
  public DeleteMark(uid: number, event?: React.MouseEvent<HTMLButtonElement>): void {
    debugger;
    event?.preventDefault();
    let web = new Web('Main');

    let list = web.lists.getByTitle('MarkRecipients');
    list.items.getById(uid).delete().then(() => {
      console.log('List Item Deleted');
      this.retrieveMarkForInfo();
    });

  }
  /*-- Ending Delete Mark for Info Recipients--*/

  /*-- To save name,email and id for controller people picker--*/
  private _getCCPeople(items: any[]) {
    debugger;
    this.state.MarkIDs.length = 0;
    let Recpid = [];
    let Recpname = [];
    let Recpemail = [];

    for (let item in items) {
      Recpid.push(items[item].id);
      Recpname.push(items[item].text);
      Recpemail.push(items[item].loginName.split("|")[2]);
    }
    this.setState({ MarkName: Recpname });
    this.setState({ MarkIDs: Recpid });
    this.setState({ MarkEmails: Recpemail });
  }
  /*--End--*/

  /*-- To save name,email and id for recommander people picker--*/
  private _getRecommender(items: any[]) {
    debugger;
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
        // alert(items[item].id);
      }

      this.setState({ RecpID: Recpid });
      this.setState({ RecpName: Recpname });
      this.setState({ RecpEmail: Recpemail });

    } // Ending If of items.length

  }
  /*--End--*/

  /*-- To Update Recommanders in Approvals list--*/
  private AddRecommender() {
    debugger;
    let seqno = this.state.RecomselectedItems.length + this.state.RecomNewselectedItems.length + 1;
    let MgrID = this.state.RecpID;
    let userid = this.state.UserID;
    let ReqID = this.state.ReqID;
    let ControllerID = this.state.ccIDS;
    let ControlID = 0;
    if (ControllerID.length > 0) {
      ControlID = this.state.ccIDS[0];
    }
    let TotalRecomm = this.state.RecomselectedItems.length + this.state.RecomNewselectedItems.length;
    if (this.state.RecpName.length == 0) {
      alert('Kindly select username!');
      jQuery('input[aria-label="People Picker"]').focus();
      return;
    }
    else if (TotalRecomm == 10) {
      alert('Only 10 Recommenders can be added!');
      jQuery('input[aria-label="People Picker"]').focus();
      return;
    }
    else if (ControllerID.length > 0 && ControlID == MgrID[0]) {
      alert('Controller cannot be recommender!');
      jQuery('input[aria-label="People Picker"]').focus();
      return;
    }
    else if (ReqID == MgrID[0]) {
      alert('Requester cannot be recommender!');
      jQuery('input[aria-label="People Picker"]').focus();
      return;
    }
    else {
      let mgrEmail = this.state.RecpEmail[0];
      this.checkRecommender(mgrEmail).then((len) => {
        if (len == 0) {
          this.checkApprover(mgrEmail).then((len1) => {
            if (len1 == 0) {
              let SeqNo = this.state.seqno;
              debugger;
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
                //              console.log(iar.data.ID);
                this.setState({ RecpID: [] });
                this.setState({ RecpName: [] });
                this.setState({ RecpEmail: [] });
                this.retrieveRecommendersNew();
              });
            }
            else {
              alert('Approver cannot be Recommender!');
              return;
            }
          });
        }
        else {
          alert('Recommender has already been added!');
          return;


        }

      });

    }
  }
  // End Function

  /*-- To retieve all Recommenders in Approvals list--*/
  private retrieveRecommendersNew() {
    debugger;
    let title = this.state.seqno;
    // let data = [];
    let data: any[] = []; 
    let web = new Web('Main');
    let userID = this.state.UserID;
    web.lists.getByTitle('ApprovalsChecklist').items.select("ID,Title,Status,Modified,AppName,Author/ID").expand('Author').filter("Title eq '" + title + "'").orderBy("Seq asc").getAll().then((items: any[]) => {
      debugger;
      if (items.length > 0) {
        for (let i = 0; i < items.length; i++) {
          if (items[i].Author.ID == userID) {
            data.push(<tr><td>{i + 1}</td><td>{items[i].AppName}</td><td>{items[i].Status}</td><td><button onClick={() => { this.DeleteRecommender(items[i].ID); }}>Delete</button></td></tr>);
          } else {

            let createDate = 'NA';
            if (items[i].Status.trim() != 'Pending') {
              let dt = new Date(items[i].Modified);
              let mnth = (dt.getMonth() + 1).toString();
              let dat = dt.getDate().toString();
              let hrs = dt.getHours().toString();
              let mins = dt.getMinutes().toString();
              if (mnth.length == 1) { mnth = '0' + mnth; } if (dat.length == 1) { dat = '0' + dat; } if (hrs.length == 1) { hrs = '0' + hrs; } if (mins.length == 1) { mins = '0' + mins; }
              createDate = dat + "-" + mnth + "-" + dt.getFullYear() + " " + hrs + ":" + mins;
            }

            data.push(<tr><td>{i + 1}</td><td>{items[i].AppName}</td><td>{items[i].Status}</td><td>{createDate}</td></tr>);

          }

        }
      }

    }).then(() => {
      this.setState({ RecomselectedItems: data });
    });

  }
  // End Function

  /*-- To delete Recommanders in Approvals list--*/
  public DeleteRecommender(uid: number, event?: React.MouseEvent<HTMLButtonElement>): void {
    debugger;
    event?.preventDefault();
    let web = new Web('Main');
    let list = web.lists.getByTitle('ApprovalsChecklist');
    list.items.getById(uid).delete().then(() => {
      console.log('List Item Deleted');
      this.retrieveRecommendersNew();
      $("#ddlRefer").val('No');
      $("#ddlReturn").val('No');

    });
  }
  /*--End Function--*/

  /*-For Comments -*/
  // public CheckComments(event) {
  //   let input = event.target.value;
  //   let maxlimit = 2000;
  //   this.setState({ Charsleft: maxlimit - input.length });
  // }
  
  public CheckComments(event: React.ChangeEvent<HTMLTextAreaElement>) {
    let inputValue = event.target.value; // Get only the text value
    let maxlimit = 2000;
  
    this.setState({ Charsleft: Math.max(0, maxlimit - inputValue.length) });
  }  

  /*--End--*/
  private ExpandIframe() {
    $("#divContent").hide();
    $("#divFrame").css({ "width": "100%" });
    // document.getElementById("IframeAttachmentCollapse").style.display='block';
    // document.getElementById("IframeAttachmentExpand").style.display='none';
    $("#IframeAttachmentCollapse").show();
    $("#IframeAttachmentExpand").hide();
    //let ht=document.getElementById('divContent').clientHeight+20;
    //console.log(document.getElementById('divFrame').clientHeight);
    //console.log(ht);
    //document.getElementById('divFrame').style.height=ht+"px";
  }

  // private CollapseIframe() {
  //   $("#divContent").show();
  //   $("#divContent").css({ "width": "50%" });
  //   $("#divFrame").css({ "width": "50%" });
  //   document.getElementById("IframeAttachmentExpand").style.display = 'block';
  //   document.getElementById("IframeAttachmentCollapse").style.display = 'none';
  //   $("#IframeAttachmentExpand").show();
  //   $("#IframeAttachmentCollapses").hide();
  //   let ht = document.getElementById('divContent').clientHeight + 20;
  //   console.log(document.getElementById('divFrame').clientHeight);
  //   console.log(ht);
  //   document.getElementById('divFrame').style.height = ht + "px";
  // }

  private CollapseIframe() {
    const divContent = document.getElementById("divContent");
    const divFrame = document.getElementById("divFrame");
    const iframeAttachmentExpand = document.getElementById("IframeAttachmentExpand");
    const iframeAttachmentCollapse = document.getElementById("IframeAttachmentCollapse");

    // Check if the elements exist before modifying them
    if (divContent && divFrame && iframeAttachmentExpand && iframeAttachmentCollapse) {
        $("#divContent").show();
        $("#divContent").css({ "width": "50%" });
        $("#divFrame").css({ "width": "50%" });

        iframeAttachmentExpand.style.display = 'block';
        iframeAttachmentCollapse.style.display = 'none';

        $("#IframeAttachmentExpand").show();
        $("#IframeAttachmentCollapses").hide();

        let ht = divContent.clientHeight + 20;
        console.log(divFrame.clientHeight);
        console.log(ht);
        divFrame.style.height = ht + "px";
    } 
}

  /*--End--*/
  /*--Save Email Notifications in NotesNotifications list--*/
  private AddNotesNotifications(status: string, approverID: number): Promise<any[]> {
    debugger;
    let web = new Web('WF');
    let statusno = parseInt(status.split("-")[0]);
    let statustxt = status.split("-")[1];
    let title = jQuery('#tdTitle').text();
    let dept = jQuery('#divDepartment').text();
    let Subj = jQuery('#divSubject').text();
    let Financial = jQuery('#divNoteType').text();
    let Amount = jQuery('#divAmount').text();
    let client = jQuery('#divClient').text();
    // let Comments = String(jQuery('#txtAppComments').val()).trim();
    let Comments: string = '';
    let Comments2 = document.getElementById("txtAppComments") as HTMLInputElement | null;
    if (Comments2) {
        Comments = Comments2.value.trim();
    }
    let NotifyId = approverID;
    if (statustxt == 'Called Back') {
      NotifyId = this.state.CurrAppID;
    }
    // let approverID=this.state.UserID;

    let returnedto = 0;
    if (statustxt == 'Returned') {
      returnedto = approverID;
    }
    let requester = this.state.ReqID;
    let ControllerID = this.state.ccIDS;
    let WorkflowFlag = "Triggered";
    if (statustxt == "Referred Back") {
      WorkflowFlag = "";
    }
    let qstr = window.location.search.split('uid=');

    let uid = 0;
    if (qstr.length > 1) { uid = parseInt(qstr[1]); }
    // let seqno = document.getElementById("divSeqNo").innerText;
    const divSeqNo = document.getElementById("divSeqNo");
    let seqno = '';

    if (divSeqNo) {
        seqno = divSeqNo.innerText;
    }
    return web.lists.getByTitle("ChecklistNoteNotifications").items.add({
      Title: title,
      SeqNo: seqno,
      Subject: Subj,
      Department: dept,
      Comments: Comments,
      ReturnedToId: returnedto,
      CurApproverId: approverID,
      NotifyId: NotifyId,
      Amount: Amount,
      RequesterId: requester,
      NoteType: Financial,
      WorkflowFlag: WorkflowFlag,
      ClientName: client,
      Migrate: "",
      MainRecID: uid,
      ControllerId: ControllerID[0],
      ReturnedById: this.state.ReturnedByID,
      Status: statustxt,
      StatusNo: statusno
    }).then((iar: ItemAddResult) => {
      console.log('Notifications Record Created!');
      return Promise.resolve(['Done']);
    });

  }

  /*--End--*/
}
