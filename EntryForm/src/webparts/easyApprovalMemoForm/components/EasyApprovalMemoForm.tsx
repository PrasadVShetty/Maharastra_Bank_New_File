import * as React from 'react';
import styles from './EasyApprovalMemoForm.module.scss';
import { IEasyApprovalMemoFormProps } from './IEasyApprovalMemoFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IPeoplePickerContext, PeoplePicker,PrincipalType} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { CIState } from "../Model/MemoNewState";
import { default as pnp, ItemAddResult, File,sp, Web } from "sp-pnp-js";
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/components/Button';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
//import { CurrentUser } from '@pnp/sp/src/siteusers'; 
import * as jQuery from 'jquery';
import * as $ from "jquery";
import { SPComponentLoader } from '@microsoft/sp-loader';
require('../css/custom.css');
SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css'); 
const Delete: any = require('../images/Delete.png');
const Video: any = require('../images/Video.png');
const Logo:any=require('../images/Logo.png');
const Annex:any=require('../images/Upload-Annex.png');
const NoteAtt:any=require('../images/Upload-Note.png');

export default class EasyApprovalMemoForm extends React.Component<IEasyApprovalMemoFormProps, CIState> {
  constructor(props : any) {
    super(props);
    this.handleTitle = this.handleTitle.bind(this);
    this.handleDesc = this.handleDesc.bind(this);
  this._onRenderFooterContent = this._onRenderFooterContent.bind(this);
    this.createItem = this.createItem.bind(this);
    this._getManager = this._getManager.bind(this);
    
    //  this.setButtonsEventHandlers();
       this.state = {
        selectedItems: [],
        name: '', 
        description: '', 
         pplPickerType:'',
        userManagerIDs: [],
        status: '',
        hideDialog: true,
        showPanel:false,
         onSubmission:false,
          ManagerEmail:[],
          seqno: '',
        attachments:[],
        Note:[],
        AttachType:'',
         MgrName:'',
        files:[],
        UserID:0,
        UserEmail:'',
        ImgUrl:'',
        CurrentItemId:0,
         NoteType:'',
        Notefilename:'',
        Sitename:'',
        Absoluteurl:'',
        RadioClient:'',
        DepartmentItems:[]
      };
    }

  private _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);
  }

  public render(): React.ReactElement<IEasyApprovalMemoFormProps> {
    const {  selectedItems } = this.state;
    const { name, description } = this.state;
    pnp.setup({
      spfxContext: this.props.context
    });
    const peoplePickerContext: IPeoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient
    };
    return (
      <form >
      <div className={styles.easyApprovalMemoForm}>
        <div className={styles.container}>

          <div>
          <div id="divHeadingNew" style={{display:"block",backgroundColor:"#0c78b8", textAlign:'center', color:'#fff'}}>
          <h3 className={styles.heading}>Memo Form</h3>             
          </div>        
          <div hidden id="divHeadingSubmit" style={{display:"none", backgroundColor:"#0c78b8", textAlign:'center', color:'#fff'}}>
          <h3 className={styles.heading}>Memo Form</h3> 
          </div>
          </div>

          <div className={styles.panel}>   
          <div className='row pt-2 pb-1 m-0' style={{width:"100%",backgroundColor:"#50B4E6", color:'#fff', justifyItems:'center'}}>
               <div className='col-md-1 col-lg-2 col-sm-4'>
                  <label className='control-form-label'><b>Requester</b></label>
               </div>
               <div className='col-md-2 col-lg-2 col-sm-8' id="tdName" style={{borderRight:'1px solid #fff'}} >                
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
          <div className={styles.formrow+" "+"form-group row"}>
          <div className='col-md-2 col-lg-2 col-sm-4'>
        <label className={styles.lbl+" "+styles.Reqdlabel}>Department</label>
         </div>
         <div className='col-md-10 col-lg-10 col-sm-8'>
          <select id="ddlDepartment" title="Select Department" placeholder="Select Department" className='form-control form-control-sm'>
              <option>Select</option></select>
         </div> 
      
      </div>
     
      <div className={styles.formrow+" "+"form-group row"}>
      <div className='col-md-2 col-lg-2 col-sm-4'>
          <label className={styles.lbl+" "+styles.Reqdlabel}>Subject</label>
            </div>
          <div className='col-md-10 col-lg-10 col-sm-8'>
              <input type="text" title="Enter Subject" placeholder="Enter Subject"  id="txtSubject" className='form-control form-control-sm'/>
          </div>
          
     </div>
     
     <div id="divClient">
     <div  className={styles.formrow+" "+"form-group row"}>
     <div className='col-md-2 col-lg-2 col-sm-4 pr-0'>
          <label className={styles.lbl+" "+styles.Reqdlabel}>Do you want to add Client?</label>
   </div>
         <div className='col-md-10 col-lg-10 col-sm-8'>
                      <label className="custom-radio">
      <input id="CYes" name="radioAttach" onChange={() => this.Radibtnchangeevent("radioAttach","CYes")} value="CYes" type="radio" />
      <span className="custom-control-indicator" style={{padding:"2px"}}></span>
      <span className={"custom-control-description"}>Yes</span>
    </label>
    <label className="custom-radio" style={{padding:"8px"}}>
      <input id="CNo" name="radioAttach" onChange={() =>this.Radibtnchangeevent("radioAttach","CNo")} value="CNo" type="radio" />
      <span className="custom-control-indicator" style={{padding:"2px"}}></span>
      <span className={"custom-control-description"}>No</span>
    </label>
      </div>
                    
      </div>
      </div>

              <div  id="divClientName" style={{display:"none"}}>                      
            <div  className={styles.formrow+" "+"form-group row"} >
              <div className='col-md-2 col-lg-2 col-sm-4'>
             <label className={styles.lbl+" "+styles.Reqdlabel}>Client Name/Vendor Name</label>
             </div>
             <div className='col-md-10 col-lg-10 col-sm-8'>
             <input type="text" title="Enter Client/Vendor Name" placeholder="Enter Client Name"  id="txtClient" className='form-control form-control-sm'/>
            </div>
            </div> 
           </div>
           
            <div className={styles.formrow+" "+"form-group row"}>
            <div className='col-md-2 col-lg-2 col-sm-4'>
          <label className={styles.lbl}>Comments</label>
          </div>
          <div className='col-md-10 col-lg-10 col-sm-8'>
              <textarea  id="txtComments" className='form-control form-control-sm'/>
          </div>
         
     </div>
     
                <div className={styles.formrow+" "+"form-group row"}>
                  
                <div className='col-md-2 col-lg-2 col-sm-4'>
                 <label className={styles.lbl}>Select Recipient</label>
                 </div>

          <div className='col-md-10 col-lg-10 col-sm-8'>
               <PeoplePicker
                context={peoplePickerContext}
                // titleText="People Picker"
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
            </div>      
        </div>
         
           <div className={styles.formrow+" "+"form-group"}>
           <h3 style={{backgroundColor:"#50B4E6",fontSize:"16px", width:'100%', padding:'5px 10px', color:'#fff'}}>Attachments</h3>
            </div>
           <div  className={styles.formrow+" "+"form-group pl-2"} id="divAttach" style={{display:""}}>            
            <div className={styles.lbl+" "+styles.Reqdlabel}>
             <a href="#"><img src={NoteAtt} style={{height:"16pt"}} onClick={() => { this.UploadAttach('Note'); }}></img></a>
              </div>
          <ul>
          {this.state.Note.map((vals)=>{
                let filename=vals.split("/")[1];
                                return (<li style={{position:"relative",padding:"5px"}}><a href={this.state.Absoluteurl+"/MemoAttachments/"+vals}>{filename}</a><img src={Delete} style={{width:"10pt",height:"10pt",position:"absolute" }} onClick={() => this.DeleteNote(vals)}/> </li>);
                      
              })}        
          </ul>
            <div hidden className="ms-Grid-col ms-u-sm12 block hide" id="divAttachButton" style={{backgroundColor:"white", display:"none"}}>
            <input type='file' style={{}} id='fileUploadInput' required={true} name='myfile' multiple onChange= {this.AttachLib}/>
            {/* <input type='file' style={{}} id='fileUploadInput' required={true} name='myfile' multiple onChange= {(e) => this.AttachLib(e)}/> */}
             </div>
             </div>
             <br/>
             <div  className={styles.formrow+" "+"form-group pl-2"}>
                  
          <div className={styles.lbl}>
          <a href="#"><img src={Annex} style={{height:"16pt"}} onClick={() => { this.UploadAttach('Annexures'); }}></img></a>
          <br></br>
          <small style={{color:'#f00'}}>(image,.pdf,.doc,.docx,.xlsx)</small>
          </div>
            <ul>
          {this.state.attachments.map((vals)=>{
                let filename=String(vals).split("/")[1];
              return (<li style={{position:"relative",padding:"5px"}}><a href={this.state.Absoluteurl+"/MemoAttachments/"+vals}>{filename}</a><img src={Delete} style={{width:"10pt",height:"10pt",position:"absolute" }} onClick={() => this.DeleteAttachment(String(vals))}></img> </li>);
                      
              })}
        
          </ul>
         
          </div>
          
          <div className={styles.overlay} id="overlay" style={{display:"none"}} >
              <span className={styles.overlayContent} style={{textAlign:"center"}}>Please Wait!!!</span>
       </div>
      
       <div  className={styles.formrow+" "+"form-group row ml-3"}>       
            <div id="btnCreate" className='ml-2' > 
            <PrimaryButton className='btn' style={{width:"25pt",backgroundColor:"#0c78b8", color:'#fff'}} text="Submit" onClick={() => { this.validateForm(); }} /> </div>
                                          
            <div id="btnCancel" className='pl-2'>
              <DefaultButton className='btn' style={{width:"25pt",backgroundColor:"#f00", color:'#fff'}} text="Cancel" onClick={() => { this.cancel(); }} />
                      </div>
            <div  id="btnClose" className='pl-2' style={{display:"none",width:"25pt"}}>
               <DefaultButton className='btn' style={{width:"25pt", backgroundColor:"#f00", color:'#fff'}} text="Close" onClick={() => { this.cancel(); }} />
            </div>
            <div id="" style={{display:"none",width:"25pt",borderRadius:"50%"}}>

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
  public UploadAttach(AttType:string){
    //debugger;
    this.setState({AttachType:AttType});
    setTimeout(()=>{
      //alert();
      // document.getElementById('fileUploadInput').click();
      let fileUploadInput = document.getElementById("fileUploadInput");
      if (fileUploadInput) {
      (fileUploadInput as HTMLInputElement).click(); // Clear the file input
      }
    },1500);
   
  }
public  handleChange(files : any){
 this.setState({
   files: files
 });
}


private on() {
let ht=window.innerHeight;
  // document.getElementById('overlay').style.height=ht.toString()+"px";
  // document.getElementById("overlay").style.display = "block";
  let overlay = document.getElementById('overlay');

  if (overlay) {
    overlay.style.height = ht.toString() + "px";
    overlay.style.display = "block";
  }
}

private off() {
  // document.getElementById("overlay").style.display = "none";
  let overlay = document.getElementById('overlay');

  if (overlay) {    
    overlay.style.display = "none";
  }
}
public  componentDidMount(){
  
    var reacthandler=this;

//debugger;
 //pnp.sp.web.currentUser.get().then((r: CurrentUser) => {  
 pnp.sp.web.currentUser.get().then((r) => {
   //debugger;
 //  console.log(r);
 let sitename=r['odata.id'].split("/_api")[0];
 let absoluteurl=sitename.split("com")[1]+"/Main";
 this.setState({Absoluteurl:absoluteurl});
this.setState({Sitename:sitename});
     const uname=r['UserPrincipalName'].split('@')[0];
     let username=r['Title'];
    //  document.getElementById("tdName").innerText=username;
    let overlay = document.getElementById('overlay');

    if (overlay) {    
      overlay.innerText = username;
    }
     this.setState({name:username});
     this.setState({UserID: r['Id'] });
  let CurrUserEmail=r['LoginName'].split("|")[2];
  this.setState({UserEmail:CurrUserEmail});
   const text = new Array();
   const possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
   for (let i = 0; i < 6; i++) {
       text[i] = possible.charAt(Math.floor(Math.random() * possible.length));
   }
   this.setState({ seqno: uname+"-"+text.join("") });
  // alert(uname+"-"+text.join(""));
 });

 let newDate = new Date();
let date = newDate.getDate().toString();
let month = (newDate.getMonth() + 1).toString();
let year = newDate.getFullYear().toString();

if(month.toString().length==1){month="0"+month.toString();}
if(date.toString().length==1){date="0"+date.toString();}

let fullDate=date+"-"+month+"-"+year;
// document.getElementById("tdDate").innerText=fullDate;
let overlay = document.getElementById('tdDate');

    if (overlay) {    
      overlay.innerText = fullDate;
    }

// this.getUserGroups().then((uid)=>{
  this.getDepartments();
    this.getFY();
// });

                     
 }

 
  private getDepartments(){
    //debugger;
    pnp.sp.site.rootWeb.lists.getByTitle('Departments').items.select("ID,Title,Dept_Alias").orderBy("ID asc").getAll().then((items: any[]) => {
      //debugger;
      // console.log(items);
      this.setState({DepartmentItems: items });
      let links:string='';
               for(let i=0;i<items.length;i++){
              links+= "<option value='" + items[i].Dept_Alias + "'>" + items[i].Title + "</option>";
              }
          jQuery('select[id="ddlDepartment"]').append(links);
 
  });
  }
  private getFY(){
    //debugger;
    pnp.sp.site.rootWeb.lists.getByTitle('FYMaster').items.select("ID,Title,Active").orderBy("ID asc").getAll().then((items: any[]) => {
      //debugger;
      // console.log(items);
      let links:string='';
               for(let i=0;i<items.length;i++){
                 if(items[i].Active=='Yes'){
                  jQuery('#tdFY').text(items[i].Title);
                 }
              
              }
          
 
  });
  }

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
    //debugger;
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

  private _getManager(items: any[]) {
    //debugger;
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
  }
 
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
  private cancel = () => {
    this.setState({ showPanel: false });
     // self.close();
     const query = window.location.search.split('uid=')[1];
     let uid=0;
     if( query!=undefined){uid=parseInt(query); }
     if(uid==0){
       window.location.replace(this.props.siteUrl);
     }
     else{
       window.location.replace(this.props.siteUrl);
    
     }
  }
  private _onClosePanel = () => {
    this.setState({ showPanel: false });
     // self.close();
        // self.close();
      
     
  }
 
  private redirect(){
    let sitename=this.state.Sitename;
   const query = window.location.search.split('uid=')[1];
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
 
  private _onShowPanel = () => {
    this.setState({ showPanel: true });
  }
  private handleTitle(value: string): void {
    return this.setState({
      name: value
    });
  }
 
  private handleDesc(value: string): void {
    return this.setState({
      description: value
    });
  }

  private validateForm(): void {
    //debugger;
   
    let allowCreate: boolean = true;
    this.setState({ onSubmission: true });
    
   let Department=jQuery('#ddlDepartment option:selected').val();
   let Client = $('#txtClient').val();
   let Comment=jQuery('#txtComments').val();
   let Approvers=this.state.userManagerIDs;
   let filename=this.state.Notefilename;
   let notetype=this.state.NoteType.toLowerCase();
   let ClientCheck=this.state.RadioClient;
   let Subject=jQuery('#txtSubject').val();
   let depItems :any[]= this.state.DepartmentItems;
  
   let regex = /^[A-Za-z0-9 !@#$()_.-]+$/;// /^[A-Za-z0-9\, ]+$/;
   let isValid = regex.test(String(Subject));
   if(Department=='Select'){
     alert('Kindly select the Department!');
    //  document.getElementById('ddlDepartment').focus();
    let ddlDepartment = document.getElementById('ddlDepartment');
      if (ddlDepartment) {
        ddlDepartment.focus();
      }
      allowCreate = false;
       return;
   }
   else{
      let new_arr = $.grep(depItems, (n, i)=>{ // just use arr
        return n.Dept_Alias == Department;        
      });

      if(new_arr.length<=0)
      {
        alert('The selected Deparmatent is not available. Kindly select the proper Department!');
        // document.getElementById('ddlDepartment').focus();
        let ddlDepartment = document.getElementById('ddlDepartment');
      if (ddlDepartment) {
        ddlDepartment.focus();
      }
        allowCreate = false;
        return;
      }
    }

    if(Subject==''){
     alert('Kindly enter Subject!');
    //  document.getElementById('txtSubject').focus();
     let ddlDepartment = document.getElementById('txtSubject');
      if (ddlDepartment) {
        ddlDepartment.focus();
      }
      allowCreate = false;
       return;
   }
   else if(typeof Subject === "string" && Subject.length > 250){
     alert('Max 250 chars are allowed in Subject!');
    //  document.getElementById('txtSubject').focus();
     let ddlDepartment = document.getElementById('txtSubject');
      if (ddlDepartment) {
        ddlDepartment.focus();
      }
      allowCreate = false;
       return;
   }
   else if (!isValid) {
     alert("Subject contains Special Characters.");
    //  document.getElementById('txtSubject').focus();
    let ddlDepartment = document.getElementById('txtSubject');
      if (ddlDepartment) {
        ddlDepartment.focus();
      }
     allowCreate = false;
      return;
 }
 
   
   else if(ClientCheck=='' ){
     alert('Kindly Select if client name is required!');
    //  document.getElementById('CYes').focus();
    let ddlDepartment = document.getElementById('CYes');
      if (ddlDepartment) {
        ddlDepartment.focus();
      }
      allowCreate = false;
       return;
   }
   else if( ClientCheck=='CYes' && String(Client).trim()=='' ){
     alert('Kindly enter client name!');
    //  document.getElementById('txtClient').focus();
    let ddlDepartment = document.getElementById('txtClient');
      if (ddlDepartment) {
        ddlDepartment.focus();
      }
      allowCreate = false;
       return;
   }
   else if (ClientCheck == 'CYes' && String(Client).trim() != '' && !regex.test(String(Client))) {
    alert('Client Name/Vendor Name contains Special Characters!');
    // document.getElementById('txtClient').focus();
    let ddlDepartment = document.getElementById('txtClient');
      if (ddlDepartment) {
        ddlDepartment.focus();
      }
    allowCreate = false;
    return;
  }
  else if (String(Comment).trim() != '' && !regex.test(String(Comment))) {
    alert('Comment contains Special Characters!');
    // document.getElementById('txtComments').focus();
    let ddlDepartment = document.getElementById('txtComments');
      if (ddlDepartment) {
        ddlDepartment.focus();
      }
    allowCreate = false;
    return;
  }
  else if(Approvers.length==0){
     alert('Kindly select at least 1 Recipient!');
        allowCreate = false;
      return;
   }
   else if( filename==''){
     alert('Kindly select at least 1 Main Note!');
    // document.getElementById('ddlTemplate').focus();
       allowCreate = false;
       return;
   }
   else 
   {allowCreate=true ;
      this._onShowPanel();
    }
    
  }

  private createItem(): void {
    //debugger;
    this._onClosePanel();
    this.on();
    jQuery('#Createbutton').remove();
    jQuery('#Cancelbutton').remove();
    let FY=jQuery('#tdFY').text();
    // let dept=jQuery('#ddlDepartment option:selected').text();
    // let deptAlias=jQuery('#ddlDepartment option:selected').val();
    let counter=0;
    let uid=0;
      
    let filename=this.state.Notefilename;
   
     this.getCounter().then((countVal)=>{
     counter=parseInt(countVal[0]);
     uid=parseInt(countVal[1]);
     let DeptGroupID=parseInt(countVal[2]);
     let dept = countVal[3];
     let deptAlias = countVal[3];
     // let SeqNo=this.state.seqno;
     let Subj=jQuery('#txtSubject').val();
     let Comment=jQuery('#txtComments').val();
     //  let RefIWLetter=jQuery('#txtRefIWLetter').val();
     
     let client=jQuery('#txtClient').val();
     let requester=this.state.UserID;     
     let dt=new Date();
     let mnth=(dt.getMonth()+1).toString();
     let dat=dt.getDate().toString();
     let fulldate=dat+mnth+dt.getFullYear().toString();
    let title="Memo/"+deptAlias+"/"+fulldate+"/"+counter.toString();
  
    let hrs=dt.getHours().toString();
    let mins=dt.getMinutes().toString();
    let secs=dt.getSeconds().toString();
    if(mnth.length==1 ){mnth='0'+mnth;} if(dat.length==1 ){dat='0'+dat;}if(hrs.length==1 ){hrs='0'+hrs;}if(mins.length==1 ){mins='0'+mins;}if(secs.length==1 ){secs='0'+secs;}
    let createDate=dat+"-"+mnth+"-"+dt.getFullYear()+" "+hrs+":"+mins+":"+secs;
    let log='Submitted to '+this.state.MgrName+' by '+this.state.name+' on '+createDate;
   
        let approverID= this.state.userManagerIDs[0];
           let web=new Web('Main');           
        web.lists.getByTitle('MemoWorkflow').items.add({
          Title:title,
          SeqNo:this.state.seqno,
          Subject:Subj,
          Department:dept,
          Comments:Comment,
            RecipientId:approverID,
             RequesterId:requester,
          FileName:filename,
          DeptAlias:deptAlias,
          ClientName:client,
          Migrate:"",
          FY:FY,
          DeptGroupId:DeptGroupID,
          AuditLog:log
              }).then((iar: ItemAddResult) => {
              console.log(iar.data.ID);
              let id=iar.data.ID;
              pnp.sp.site.rootWeb.lists.getByTitle("MemoWorkflow").items.add({
                Title:title,
                SeqNo:this.state.seqno,
                Subject:Subj,
                PID:id,
                Department:dept,
                Comments:Comment,
                  RecipientId:approverID,
                   RequesterId:requester,
                FileName:filename,
                DeptAlias:deptAlias,
                ClientName:client,
                Migrate:"",
                FY:FY,
                DeptGroupId:DeptGroupID,
                AuditLog:log
               }).then((iar1: ItemAddResult) => {
        console.log(iar1.data.ID);
        this.setCounter(uid,counter).then(()=>{
                          this.redirect();
                    });
              });
          });
        });
      
  }
  
  //  private getCounter(dept:string):Promise<any[]>{
  //   //  let num=[];
  //   let num: number[] = []; 
  //    return pnp.sp.site.rootWeb.lists.getByTitle('Departments').items.select("ID,Title,Dept_Alias,MemoCounter,GroupID").filter("Dept_Alias eq '"+dept+"'").orderBy("ID asc").getAll().then((items: any[]) => {
  //     //debugger;
  //      // console.log(items);    
  //      num[0]=parseInt(items[0].MemoCounter)+1;
  //      num[1]=items[0].ID;
  //      num[2]=items[0].GroupID;
  //      return num;
  //    });    
  //  }

  private getCounter():Promise<any[]>{
  //  let num=[];
  let num: number[] = []; 
    //return pnp.sp.site.rootWeb.lists.getByTitle('Departments').items.select("ID,Title,Dept_Alias,MemoCounter,GroupID").filter("Dept_Alias eq '"+dept+"'").orderBy("ID asc").getAll().then((items: any[]) => {
      return pnp.sp.site.rootWeb.lists.getByTitle('Counter').items.select("ID,Title,NoteId,MemoCounter,Department").orderBy("ID asc").getAll().then((items: any[]) => {
    //debugger;
      // console.log(items);    
      num[0]=parseInt(items[0].MemoCounter)+1;
      num[1]=items[0].ID;
      num[2]=items[0].GroupID;
      num[3]=items[0].Department;
      return num;
    });
  }

  //  private setCounter(uid:number,counter:number):Promise<any[]>{
  //  return pnp.sp.site.rootWeb.lists.getByTitle('Departments').items.getById(uid).update({
  //     MemoCounter: counter
  //   }).then(() => {
  //     console.log('Counter updated');
  //       return Promise.resolve(['Done']); 
      
  //   });
  
  //  }
  
  private setCounter(uid:number,counter:number):Promise<any[]>{
    return pnp.sp.site.rootWeb.lists.getByTitle('Counter').items.getById(uid).update({
      MemoCounter: counter
    }).then(() => {
      console.log('Counter updated');
        return Promise.resolve(['Done']); 
      
    });

  }

   private gotoHomePage(): void {
  // self.close();
       window.location.replace(this.props.siteUrl);
   }
  
  
    
     public DeleteAttachment(vals:string):void{
       //debugger;
       this.setState({
         attachments:[]
       });
       let sitename=this.state.Absoluteurl;
       let web=new Web('Main'); 
       let url=sitename+'/MemoAttachments/'+vals;
       let fldr=vals.split("/")[0];
       let fldURL=sitename+'/MemoAttachments/'+fldr;
       web.getFileByServerRelativeUrl(url).recycle().then(data=> {  
         console.log("File Deleted " + vals) ;
         web.getFolderByServerRelativeUrl(fldURL).files.get().then((result) => {
           let links:any[]=[];
        
           for(let i=0;i<result.length;i++){
             links.push(fldr+"/"+result[i].Name);
  
           }
          
    
          
            // console.log(links);
           this.setState({ attachments: links});
          //  document.getElementById("fileUploadInput").nodeValue=null;
           let ddlDepartment = document.getElementById('ddlDepartment');
            if (ddlDepartment) {
              ddlDepartment.nodeValue = null;
            }
       });
       
       });
  
     }
  
     public DeleteNote(vals:string):void{
      //debugger;
      this.setState({
        Note:[]
      });
      let sitename=this.state.Absoluteurl;
      let url=sitename+'/MemoAttachments/'+vals;
      let fldr=vals.split("/")[0];
      let fldURL=sitename+'/MemoAttachments/'+fldr;
      let web=new Web('Main');           
      web.getFileByServerRelativeUrl(url).recycle().then(data=> {  
        console.log("File Deleted " + vals) ;
        web.getFolderByServerRelativeUrl(fldURL).files.get().then((result) => {
          let links:any[]=[];
       
          for(let i=0;i<result.length;i++){
            links.push(fldr+"/"+result[i].Name);
  
          }
         
   
         
           // console.log(links);
          this.setState({ Note: links});
          this.setState({Notefilename:""});
          // document.getElementById("fileUploadInput").nodeValue=null;
          // document.getElementById("NoteDel").style.display="none";
      jQuery('#NoteFile').text('');
      });
      
      });
  
    }
     public  AttachLib=(event : any)=> {
       //debugger;
          var uploadFlag=true;
      //in case of multiple files,iterate or else upload the first file.
       // let file = fileUpload.files[0];
       let file = event.target.files[0];
       let filesize=file.size/1048576;
       //let fileExtn=file.name.split(".")[1].toLowerCase();
       let fileExtn=file.name.substr((file.name.lastIndexOf('.') + 1)).toLowerCase();
       let fileSplit=file.name.split(".");
       let fileType=this.state.AttachType;
       let PermissibleExtns=['pdf'];
       let listName='MemoAttachments';
      let NoteCount=this.state.Note.length;
      let notetype=this.state.NoteType;
      let SeqNo=this.state.seqno;
      let web=new Web('Main');     
      let fldr='';
      let match = (new RegExp('[~#%\&{}+\|]|\\.\\.|^\\.|\\.$')).test(file.name.split(".")[0]);
       if(fileType!='Note'){
         PermissibleExtns=['png','jpeg','jpg','gif','pdf','doc','docx','xls','xlsx'];
         fldr=SeqNo+'-Annex';
       }
       else {
         fldr=SeqNo;
        PermissibleExtns=['pdf'];
             }
       
      
       if(fileSplit.length>2)
       {
        alert('Alert-Selected file double extension is not allowed!');
        // document.getElementById("fileUploadInput").nodeValue=null;
        let ddlDepartment = document.getElementById('fileUploadInput');
      if (ddlDepartment) {
        ddlDepartment.nodeValue = null;
      }
        return false;
       }
       else if(match)
       {
        alert('Invalid file name. The name of the attached file contains invalid characters!');
        // document.getElementById("fileUploadInput").nodeValue=null;
        let ddlDepartment = document.getElementById('fileUploadInput');
      if (ddlDepartment) {
        ddlDepartment.nodeValue = null;
      }
        return false;
  
       }else if(file.name.split(".")[0].length >150){
        alert('Invalid file name. file names cannot be more than 150 chars!');
        // document.getElementById("fileUploadInput").nodeValue=null;
        let ddlDepartment = document.getElementById('fileUploadInput');
      if (ddlDepartment) {
        ddlDepartment.nodeValue = null;
      }
        return false;
       }
       else if(PermissibleExtns.indexOf(fileExtn.toLowerCase())==-1){
         alert('Alert-Selected file type is not allowed!');
        //  document.getElementById("fileUploadInput").nodeValue=null;
        let ddlDepartment = document.getElementById('fileUploadInput');
      if (ddlDepartment) {
        ddlDepartment.nodeValue = null;
      }
         return false;
       }
       else if(  filesize>5 ){
         alert('Alert-File size is more than permissible limit!');
        //  document.getElementById("fileUploadInput").nodeValue=null;
        let ddlDepartment = document.getElementById('fileUploadInput');
      if (ddlDepartment) {
        ddlDepartment.nodeValue = null;
      }
         return false;
       }
       else if(fileType=='Note' && NoteCount==1){
        alert('Alert-Only 1 Note is allowed!');
        // document.getElementById("fileUploadInput").nodeValue=null;
        let ddlDepartment = document.getElementById('fileUploadInput');
      if (ddlDepartment) {
        ddlDepartment.nodeValue = null;
      }
        return false;
       }
       else{
       if (file!=undefined || file!=null){
                  
                  web.getFolderByServerRelativePath(listName).folders.add(fldr).then(data=> {  
           console.log("Folder is created at " + data.data.ServerRelativeUrl) ;
       //assuming that the name of document library is Documents, change as per your requirement, 
       //this will add the file in root folder of the document library, if you have a folder named test, replace it as "/Documents/test"
       
     web.getFolderByServerRelativePath(listName+"/"+fldr).files.add(file.name, file, true).then((result) => {
          console.log(file.name + " uploaded successfully!");
          let links:any[]=[];
          
          if(fileType=='Note'){
            this.setState({Notefilename:file.name});
            links=this.state.Note;
            links.push(SeqNo+"/"+file.name);
            this.setState({ Note: links});
          }else{
          links=this.state.attachments;
          links.push(SeqNo+"-Annex/"+file.name);
          this.setState({ attachments: links});
        }
          console.log(links);
          
          // document.getElementById("fileUploadInput").nodeValue=null;
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
  
    
     private Radibtnchangeevent(name : string,value : string){
    //debugger;
    
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
      private _closeDialog = (): void => {
        this.setState({ hideDialog: true });
      }
}
