import * as React from 'react';
import styles from './EasyApprovalNotesTat.module.scss';
import { IEasyApprovalNotesTatProps } from './IEasyApprovalNotesTatProps';
//import { escape, round } from '@microsoft/sp-lodash-subset';
import { IReactPnpResponsiveDataTableState } from './DataTableState';  
import { SPComponentLoader } from '@microsoft/sp-loader';  
import { PrimaryButton} from 'office-ui-fabric-react/lib/components/Button';
import * as $ from 'jquery';  
import * as pnp from "sp-pnp-js";
import 'jszip/dist/jszip';  
import 'pdfmake/build/pdfmake';  
import 'datatables.net';  
import 'datatables.net-responsive';  
import 'datatables.net-buttons';  
import 'datatables.net-buttons/js/buttons.html5';  
import 'datatables.net-buttons/js/buttons.print';  
//import { string } from 'prop-types';
import { Web } from 'sp-pnp-js';
SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/bootstrap@4.6.2/dist/css/bootstrap.min.css'); 
SPComponentLoader.loadCss('/sites/EasyApproval/SiteAssets/css/styles.css');  
SPComponentLoader.loadCss('https://cdn.datatables.net/responsive/2.2.3/css/responsive.bootstrap.min.css');    
SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.20/css/jquery.dataTables.min.css');  
SPComponentLoader.loadCss('https://cdn.datatables.net/buttons/1.6.0/css/buttons.dataTables.min.css');  
SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.53/vfs_fonts.js')  ;
require ('../assets/styles.css');
var sSearchtext='Search :';  
var sInfotext = 'Showing _START_ to _END_ of _TOTAL_ entries';  
var   sZeroRecordsText='No data available in table';  
var sinfoFilteredText="(filtered from _MAX_ total records)";  
//var   placeholderkeyword="Keyword";  
var lengthMenutxt="Show _MENU_ entries";  
var firstpage="First";  
var Lastpage="Last";  
var Nextpage="Next";  
var Previouspage="Previous"; 

export default class EasyApprovalNotesTAT extends React.Component<IEasyApprovalNotesTatProps, IReactPnpResponsiveDataTableState> {
  constructor(props: IEasyApprovalNotesTatProps, state: IReactPnpResponsiveDataTableState) {  
    debugger;
    super(props);  
    this.state = {  
      Sitename:'',
      Projectstatus: [{ Title: "", Description: "", id: "", Requester: "", Created: "" }]  
  
    };  
    this.fetchdatas = this.fetchdatas.bind(this);
  }  
 public componentDidMount(){  
   debugger;
   this.getDepartments();
   this.getFinNotes();
   this.getDOP();
   this.getFY();
   // this.fetchdatas();  
  }  

  private fetchdatas() 
  {  
    debugger;
    $('#SpfxDatatable').DataTable().destroy();
    //var ret = true;
    var nospecial=/[*|\":<>[\]{}`\\()';@&$%~!#^?+=]/;

    //var reqno = $('#txtRef').val().trim('');
    var reqno = ($('#txtRef').val() as string | undefined)?.toString().trim() || '';
    if (nospecial.test(reqno)) 
    { 
      alert('Note contains invalid characters!');
      return false;
    }

    //var req = $('#txtRequester').val().trim('');
    var req = ($('#txtRequester').val() as string | undefined)?.toString().trim() || '';
    if (nospecial.test(req)) 
    { 
      alert('Requester contains invalid characters!');
      return false;
    }

    // var client = $('#txtClient').val().trim('');
    var client = ($('#txtClient').val() as string | undefined)?.toString().trim() || ''; 
    if (nospecial.test(client)) 
    { 
      alert('Client Name contains invalid characters!');
      return false;
    }
    
    //var subject = $('#txtSubject').val().trim('');
    var subject = ($('#txtSubject').val() as string | undefined)?.toString().trim() || '';
    if (nospecial.test(subject)) 
    { 
      alert('Subject contains invalid characters!');
      return false;
    }
   
    // var Financial = $('#ddlFinancial').val().trim('');
	  // var dept = $('#ddlDepartment').val().trim('');
    // let  sdt= $('#txtFromDate').val().trim('');   
    // let dop=$('#ddlDOP').val().trim('');
    // let fy=$('#ddlFY').val().trim('');

    var Financial = ($('#ddlFinancial').val() as string | undefined)?.toString().trim() || '';
    var dept = ($('#ddlDepartment').val() as string | undefined)?.toString().trim() || '';
    var sdt = ($('#txtFromDate').val() as string | undefined)?.toString().trim() || '';
    var dop = ($('#ddlDOP').val() as string | undefined)?.toString().trim() || '';
    var fy = ($('#ddlFY').val() as string | undefined)?.toString().trim() || '';


    // let approver=$('#txtApprover').val().trim(''); 
    var approver = ($('#txtApprover').val() as string | undefined)?.toString().trim() || '';
    if (nospecial.test(approver)) 
    { 
      alert('Approver contains invalid characters!');
      return false;
    }
    
    let GMTSdate='';

    if(sdt!=""){
      let Sdate:Date=new Date(sdt.substring(0,4)+"-"+sdt.substring(5,7)+"-"+sdt.substring(8,10));
         GMTSdate=Sdate.getFullYear().toString()+"-"+(Sdate.getMonth()+1).toString()+"-"+Sdate.getDate().toString();
    }

    let   GMTTodate='';
    // var tdt = $('#txtToDate').val().trim('');  
    var tdt = ($('#txtToDate').val() as string | undefined)?.toString().trim() || ''; 
    if(tdt!=""){
      var Todate=new Date(tdt.substring(0,4)+"-"+tdt.substring(5,7)+"-"+tdt.substring(8,10));
      GMTTodate=Todate.getFullYear().toString()+"-"+(Todate.getMonth()+1).toString()+"-"+Todate.getDate().toString();
    }

    var n=0;
    let str ="";	
	  if(reqno != ""){
		  str = str + "<Contains><FieldRef Name='Title' /><Value Type='Text'>"+reqno+"</Value></Contains>--";
      n=n+1;
	  }
    if(req != ""){
      str = str + "<Contains><FieldRef Name='Requester' /><Value Type='Lookup'>"+req+"</Value></Contains>--";	
      n=n+1;
    }
    if(client != ""){
      str = str + "<Contains><FieldRef Name='ClientName' /><Value Type='Text'>"+client+"</Value></Contains>--";	
      n=n+1;
    }
    //if(status != ""){
      str = str + "<Contains><FieldRef Name='Status' /><Value Type='Text'>Approved</Value></Contains>--";	
      n=n+1;
    //}
    if(subject != ""){
      str = str + "<Contains><FieldRef Name='Subject' /><Value Type='Text'>"+subject+"</Value></Contains>--";	
      n=n+1;
    }
    if(dept != "Select"){
      str = str + "<Eq><FieldRef Name='Department' /><Value Type='Text'>"+dept+"</Value></Eq>--";	
      n=n+1;
    }
    if(Financial != "Select"){
      str = str + "<Eq><FieldRef Name='NoteType' /><Value Type='Text'>"+Financial+"</Value></Eq>--";	
      n=n+1;
    }
    if(dop != "Select"){
      str = str + "<Eq><FieldRef Name='DOP' /><Value Type='Text'>"+dop+"</Value></Eq>--";	
      n=n+1;
    }
    if(fy != "Select"){
      str = str + "<Eq><FieldRef Name='FY' /><Value Type='Text'>"+fy+"</Value></Eq>--";	
      n=n+1;
    }
    if(approver != ""){
      //str = str + "<Contains><FieldRef Name='CurApprover' /><Value Type='Lookup'>"+approver+"</Value></Contains>--";	
      str = str + "<Contains><FieldRef Name='CurApprover' /><Value Type='Text'>"+approver+"</Value></Contains>--";
      n=n+1;
    }
    if(sdt != ""){
      str = str + "<Geq><FieldRef Name='Created' /><Value IncludeTimeValue='TRUE' Type='DateTime'>"+GMTSdate+"</Value></Geq>--";	
      n=n+1;
    }
    if(tdt != ""){
      str = str + "<Leq><FieldRef Name='Created' /><Value IncludeTimeValue='TRUE' Type='DateTime'>"+GMTTodate+"</Value></Leq>--";	
      n=n+1;
    }

    let finalstr="";
    let Sstr=str.split("--");
    if(n==0){
      alert('Kindly select any 1 parameter');
    }
    else if(n==1){
      str=str.split("--")[0];
      finalstr="<Query><Where>"+Sstr+"</Where></Query>";
    }
    else if(n==2){
      finalstr="<Query><Where><And>"+Sstr[0]+Sstr[1]+"</And></Where></Query>";
    }
    else if(n==3){
      finalstr="<Query><Where><And>"+Sstr[0]+"<And>"+Sstr[1]+Sstr[2]+"</And></And></Where></Query>";
    }
    else if(n==4){
      finalstr="<Query><Where><And>"+Sstr[0]+"<And>"+Sstr[1]+"<And>"+Sstr[2]+Sstr[3]+"</And></And></And></Where></Query>";
    }
    else if(n==5){
      finalstr="<Query><Where><And>"+Sstr[0]+"<And>"+Sstr[1]+"<And>"+Sstr[2]+"<And>"+Sstr[3]+Sstr[4]+"</And></And></And></And></Where></Query>";
    }
    else if(n==6){
      finalstr="<Query><Where><And>"+Sstr[0]+"<And>"+Sstr[1]+"<And>"+Sstr[2]+"<And>"+Sstr[3]+"<And>"+Sstr[4]+Sstr[5]+"</And></And></And></And></And></Where></Query>";
    }
    else if(n==7){
      finalstr="<Query><Where><And>"+Sstr[0]+"<And>"+Sstr[1]+"<And>"+Sstr[2]+"<And>"+Sstr[3]+"<And>"+Sstr[4]+"<And>"+Sstr[5]+Sstr[6]+"</And></And></And></And></And></And></Where></Query>";
    }
    else if(n==8){
      finalstr="<Query><Where><And>"+Sstr[0]+"<And>"+Sstr[1]+"<And>"+Sstr[2]+"<And>"+Sstr[3]+"<And>"+Sstr[4]+"<And>"+Sstr[5]+"<And>"+Sstr[6]+Sstr[7]+"</And></And></And></And></And></And></And></Where></Query>";
    }
    else if(n==9){
      finalstr="<Query><Where><And>"+Sstr[0]+"<And>"+Sstr[1]+"<And>"+Sstr[2]+"<And>"+Sstr[3]+"<And>"+Sstr[4]+"<And>"+Sstr[5]+"<And>"+Sstr[6]+"<And>"+Sstr[7]+Sstr[8]+"</And></And></And></And></And></And></And></And></Where></Query>";
    }

    const xml = `<View>      <ViewFields>
                            <FieldRef Name='ID' />
                            <FieldRef Name='Title' />
                            <FieldRef Name='Requester' />
                            <FieldRef Name='Subject' />
                            <FieldRef Name='Department' />
                            <FieldRef Name='ClientName' />
                            <FieldRef Name='Status' />
                            <FieldRef Name='NoteType' />
                            <FieldRef Name='PID' />
                            <FieldRef Name='Created' />
                            </ViewFields>`+finalstr+`
                        <RowLimit>5000</RowLimit>
                    </View>`;
    const q: any = {
      ViewXml: xml
    };

    let FetchProjectDetails: any[] = [];  
    let web = new Web('Main');
    web.lists.getByTitle('Notes').getItemsByCAMLQuery(q,'FieldValuesAsText').then((r: any[]) => 
    {
      console.log(r);
      debugger;
      for (let i = 0; i < r.length; i++) 
      {  
        var fdate=this.formatDate(r[i].Created);
        var mdate=this.formatDate(r[i].Modified);
        let Mod=new Date(r[i].Modified);
        let Creat=new Date(r[i].Created);
    
        let days= Math.round((Mod.valueOf()-Creat.valueOf())/1000/60/60/24);
        FetchProjectDetails.push({  
          //Title: r[i].Title,  
          Title: (r[i].Title!=undefined?r[i].Title:""),
          Subject: r[i].Subject,  
          id: r[i].Id,  
          Status:r[i].Status,
          Requester: r[i].FieldValuesAsText.Requester, 
          Department: r[i].Department, 
          PID:r[i].PID, 
          Client:r[i].ClientName,
          Financial:r[i].NoteType,
          Created:fdate,
          Modified:mdate,
          Days:days
        });  
      }  

      this.setState({ Projectstatus: FetchProjectDetails });
      this.setState({Sitename:this.props.context.pageContext.web.absoluteUrl});
      //document.getElementById("Searchresults").style.display="block";
      const searchResults = document.getElementById("Searchresults");
      if (searchResults) {
          searchResults.style.display = "block";
      }
      $.extend( $.fn.dataTable.defaults, {  
        responsive: {
          details: {
              type: 'column',
              target: 'tr'
          }
        }
      });  
  
      $("#SpfxDatatable").DataTable({  
        "info": true,  
        "pagingType": 'full_numbers',  
        dom: 'lBfrtip',
        scrollX: true,
        buttons: 
        [  
          {
            extend: 'csv',
            text:"Export to CSV",
            title:'My Requests',
            className: 'buttonexcel',
          },                   
        ],             
        "order": [],  
        "language": 
        {  
          "infoEmpty":sInfotext,  
          "info":sInfotext,  
          "zeroRecords":sZeroRecordsText,  
          "infoFiltered":sinfoFilteredText,  
          "lengthMenu": lengthMenutxt,  
          "search":sSearchtext,  
          "paginate": 
          {  
            "first": firstpage,  
            "last": Lastpage,  
            "next": Nextpage,  
            "previous": Previouspage  
          }      
        }      
      });  
    });
  }
  
  // public formatDate(InputDate){
  //   //debugger;
  //   var dt  = InputDate.split("T");
  //   var dt1=dt[0].split("-");
  //   var dateOutput=dt1[2]+"/"+dt1[1]+"/"+dt1[0];
   
  //   return dateOutput;
  // }
  public formatDate(InputDate: string): string {    
    var dt = InputDate.split("T");
    var dt1 = dt[0].split("-");
    var dateOutput = dt1[2] + "/" + dt1[1] + "/" + dt1[0];
    return dateOutput;
  }
  
  public render(): React.ReactElement<IEasyApprovalNotesTatProps> 
  {
    debugger;
    return (
      <div className={styles.easyApprovalNotesTAT}>
        <div className={styles.container}>
        <h2 id="Submitted" style={{textAlign:"center",backgroundColor:"#0c78b8",color:"white",display:"block",  padding:'5px 0px', fontSize:'18px'}}>Note TAT Parameters</h2>
         
    <div className='row form-group ml-0 mr-0'>
    <div className='col-md-1 col-lg-2 col-lg-02 col-sm-4'>
        <label>Note#</label>
    </div>
    <div className='col-md-3 col-lg-2 col-sm-8'>
    <input type="Text" className="form-control form-control-sm" id="txtRef"/>
     </div>
     <div className='col-md-1 col-lg-2 col-lg-02 col-sm-4'>
        <label>Requester</label>
    </div>
    <div className='col-md-3 col-lg-2 col-sm-8'>
    <input type="Text" className="form-control form-control-sm" id="txtRequester"  />
     </div>

     <div className='col-md-1 col-lg-2 col-lg-02 col-sm-4'>
        <label>Department</label>
    </div>
    <div className='col-md-3 col-lg-2 col-sm-8'>
    <select id="ddlDepartment" className="form-control form-control-sm" >
                <option>Select</option>
              </select>
     </div>
       </div>   

     <div className='row form-group ml-0 mr-0'>
    <div className='col-md-1 col-lg-2 col-lg-02 col-sm-4'>
        <label>Client Name</label>
    </div>
    <div className='col-md-3 col-lg-2 col-sm-8'>
    <input type="Text" className="form-control form-control-sm" id="txtClient" />
     </div>
     <div className='col-md-1 col-lg-2 col-lg-02 col-sm-4'>
        <label>From Date</label>
    </div>
    <div className='col-md-3 col-lg-2 col-sm-8'>
    <input type="date" className="form-control form-control-sm" id="txtFromDate" />
     </div>
     <div className='col-md-1 col-lg-2 col-lg-02 col-sm-4'>
        <label>To Date</label>
    </div>
    <div className='col-md-3 col-lg-2 col-sm-8'>
    <input type="date" className="form-control form-control-sm" id="txtToDate" />
     </div>

    </div>     

    <div className='row form-group ml-0 mr-0'>
    <div className='col-md-1 col-lg-2 col-lg-02 col-sm-4'>
        <label>Subject</label>
    </div>
    <div className='col-md-3 col-lg-2 col-sm-8'>
    <input type="Text" className="form-control form-control-sm" id="txtSubject" />
     </div>
     <div className='col-md-1 col-lg-2 col-lg-02 col-sm-4'>
        <label>Financial</label>
    </div>
    <div className='col-md-3 col-lg-2 col-sm-8'>
    <select id="ddlFinancial" className="form-control form-control-sm">
                <option>Select</option>
                <option>Non-Financial</option>
              </select>
     </div>
     <div className='col-md-1 col-lg-2 col-lg-02 col-sm-4'>
        <label>FY</label>
    </div>
    <div className='col-md-3 col-lg-2 col-sm-8'>
    <select id="ddlFY" className="form-control form-control-sm">
                <option>Select</option>
              </select> 
     </div>
    </div>    
        
    <div className='row form-group ml-0 mr-0'>
    <div className='col-md-1 col-lg-2 col-lg-02 col-sm-4'>
        <label>Approver</label>
    </div>
    <div className='col-md-3 col-lg-2 col-sm-8'>
    <input type="Text" className="form-control form-control-sm" id="txtApprover" />
     </div>
     <div className='col-md-1 col-lg-2 col-lg-02 col-sm-4'>
        <label>DOP</label>
    </div>
    <div className='col-md-3 col-lg-2 col-sm-8'>
    <select id="ddlDOP" className="form-control form-control-sm">
                <option>Select</option>
              </select>
     </div>
     
    </div> 

        <div className='form-group'>          
        <div style={{display:"block",textAlign:"center"}}>
          <PrimaryButton className='btn bnt-primary' style={{width:"100px", backgroundColor:"#2380db", color:'#fff'}} text="Search" onClick={() => { this.fetchdatas(); }} /> 
          </div>
          </div> 
         
          <div  className="ms-u-sm12 block" style={{display:"block"}}><br></br></div>
         
          <div  className="ms-u-sm12 block" style={{display:"block"}}></div>
          <div className=" ms-u-sm12 block" id="Searchresults" style={{display:"none"}}>
            <h2 id="Submitted" style={{textAlign:"center",backgroundColor:"#FF6633",color:"white",display:"block", padding:'5px 0px', fontSize:'18px'}}>Note TAT Results</h2>
            <div className='pl-2 pr-2'>
            <table className='table table-striped table-bordered row-border stripe' id='SpfxDatatable'>  
              <thead>  
                <tr>  
                  <th>Title</th>  
                  <th>Requester</th>
                  <th>Department</th>  
                  <th>Client</th>
                  <th style={{maxWidth:'200px'}}>Subject</th>
                  <th>Status</th>  
                  <th>Created</th>  
                  <th style={{maxWidth:'120px'}}>Approved On</th>
                  <th>TAT</th>
                  <th>ID</th> 
                </tr>  
              </thead>  
              <tbody id='SpfxDatatableBody'>  
                {
                  this.state.Projectstatus && this.state.Projectstatus.map((item: any, i: number) => 
                  {  
                    return [  
                      <tr key={i}>  
                        <td><a href={"javascript:window.open('"+this.state.Sitename+"/SitePages/NoteApproval.aspx/?uid="+item.id+"','','width=600,height=700')"} >{item.Title}</a></td>  
                        <td>{item.Requester}</td>  
                        <td>{item.Department}</td>
                        <td>{item.Client}</td>
                        <td>{item.Subject}</td>
                        <td>{item.Status}</td>  
                        <td>{item.Created}</td> 
                        <td>{item.Modified}</td>
                        <td>{item.Days}</td> 
                        <td>{item.id}</td>  
                      </tr> 
                    ];  
                  })
                }  
              </tbody>  
            </table>  
            </div>
          </div>
          
        </div>  

      </div> 
    ); 
  }  
  
  private getDepartments()
  {
    debugger;
    pnp.sp.web.lists.getByTitle('Departments').items.select("ID,Title,Dept_Alias").orderBy("ID asc").getAll().then((items: any[]) => 
    {
      debugger;
      console.log(items);
      let links:string='';
      for(let i=0;i<items.length;i++)
      {
        links+= "<option value='" + items[i].Title + "'>" + items[i].Title + "</option>";
      }
      $('select[id="ddlDepartment"]').append(links);
    });
  }

  private getFinNotes()
  {
    debugger;
    pnp.sp.web.lists.getByTitle('FinNotes').items.select("ID,Title").orderBy("Title asc").getAll().then((items: any[]) => 
    {
      debugger;
      console.log(items);
      let links:string='';
      for(let i=0;i<items.length;i++)
      {
        links+= "<option value='" + items[i].Title + "'>" + items[i].Title + "</option>";
      }
      $('select[id="ddlFinancial"]').append(links);
    });
  }

  private getFY()
  {
    debugger;
    pnp.sp.web.lists.getByTitle('FYMaster').items.select("ID,Title").orderBy("Title asc").getAll().then((items: any[]) => 
    {
      debugger;
      console.log(items);
      let links:string='';
      for(let i=0;i<items.length;i++)
      {
        links+= "<option value='" + items[i].Title + "'>" + items[i].Title + "</option>";
      }
      $('select[id="ddlFY"]').append(links);
    });
  }

  private getDOP()
  {
    debugger;
    pnp.sp.web.lists.getByTitle('DOP').items.select("ID,Title").orderBy("Title asc").getAll().then((items: any[]) => 
    {
      debugger;
      console.log(items);
      let links:string='';
      for(let i=0;i<items.length;i++)
      {
        links+= "<option value='" + items[i].Title + "'>" + items[i].Title + "</option>";
      }
      $('select[id="ddlDOP"]').append(links);
    });
  }

  public componentWillMount(){

  }
  public componentDidUpdate() {  
   
  }
  
}
