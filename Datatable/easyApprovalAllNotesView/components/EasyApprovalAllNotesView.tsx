import * as React from 'react';
import _styles from './EasyApprovalAllNotesView.module.scss';
import { IEasyApprovalAllNotesViewProps } from './IEasyApprovalAllNotesViewProps';
import { IReactPnpResponsiveDataTableState } from './DataTableState';  
//import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';  
//import { sp } from '@pnp/sp/presets/all'; 
import * as $ from 'jquery';  
//import { sp} from '@pnp/sp';  
import * as pnp from "sp-pnp-js";
 
import 'pdfmake/build/pdfmake';  
import 'datatables.net';  
import 'datatables.net-responsive';  
import 'datatables.net-buttons';  
import 'datatables.net-buttons/js/buttons.html5';  
import 'datatables.net-buttons/js/buttons.print';  
//import { string } from 'prop-types';
//import { JSONParser } from '@pnp/odata';
SPComponentLoader.loadCss('https://cdn.datatables.net/responsive/2.2.3/css/responsive.bootstrap.min.css');
SPComponentLoader.loadCss('/sites/EasyApproval/SiteAssets/css/styles.css');     
SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.20/css/jquery.dataTables.min.css');  
SPComponentLoader.loadCss('https://cdn.datatables.net/buttons/1.6.0/css/buttons.dataTables.min.css');  
SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.53/vfs_fonts.js')  ;
require ('../assets/styles.css');
 
var sSearchtext='Search :';  
var sInfotext = 'Showing _START_ to _END_ of _TOTAL_ entries';  
var sZeroRecordsText='No data available in table';  
var sinfoFilteredText="(filtered from _MAX_ total records)";  
//var placeholderkeyword="Keyword";  
var lengthMenutxt="Show _MENU_ entries";  
var firstpage="First";  
var Lastpage="Last";  
var Nextpage="Next";  
var Previouspage="Previous";

export default class EasyApprovalAllNotesView extends React.Component<IEasyApprovalAllNotesViewProps, IReactPnpResponsiveDataTableState> {
  constructor(props: IEasyApprovalAllNotesViewProps, state: IReactPnpResponsiveDataTableState) {  
    //debugger;
    super(props);  
    this.state = {  
      Sitename:'',
      Projectstatus: [{ Title: "", Description: "", id: "", Requester: "", Created: "" }] ,
      Projectstatus1:[],
      TempData:[]
    };  
    this.MNfetchdatas = this.MNfetchdatas.bind(this);   
  }  
  public componentDidMount(){  
   //debugger;
  //  pnp.sp.web.currentUser.get().then((r: CurrentUser) => { 
    pnp.sp.web.currentUser.get().then((r) => {  
    //debugger;
  
   let CurrUserEmail=r['Email'];
    this.MNfetchdatas(CurrUserEmail);  
  });
  }  
  private MNfetchdatas(CurrUserEmail : any) {  
    debugger;
    // let web = new Web(this.props.context.pageContext.web.absoluteUrl);  
     // let web=pnp.sp.web;
    console.log(CurrUserEmail);
    //const list2 =  pnp.sp.web.lists.getByTitle("Notes");  
    let FetchProjectDetails :any[]= [];  
    //let WebpartDesc=this.props.description;
    debugger;
  pnp.sp.web.lists.getByTitle('Notes').items.select('ID,Title,Status,Department,Subject,PID,Created,ClientName,Modified,Requester/Title,Sitename,CurApprover/EMail,CurApprover/Title').expand('Requester,CurApprover').orderBy("Modified",false).top(5000).get().then(r => {  
      for (let i = 0; i < r.length; i++) {  
              var fdate=this.formatDate(r[i].Created);
            var Mdate=this.formatDate(r[i].Modified);
            var fdate1=this.formatDate1(r[i].Created);
            var Mdate1=this.formatDate1(r[i].Modified);
            let CApprover="";
            if(r[i].CurApprover!=undefined){CApprover=r[i].CurApprover.Title;}

            /*if(r[i].Title==undefined)
            {
              console.log('id : ' + r[i].Id);
              console.log('Subject : ' + r[i].Subject);
              console.log('Sitename : ' + r[i].Sitename);
              console.log('ClientName : ' + r[i].ClientName);
              console.log('Status : ' + r[i].Status);
              console.log('Department : ' + r[i].Department);
              console.log('Requester.Title : ' + r[i].Requester.Title);
              console.log('PID : ' + r[i].PID);
            }*/
           FetchProjectDetails.push({  
          //Title: r[i].Title,  
          Title: (r[i].Title!=undefined?r[i].Title:""),  
          Subject:r[i].Subject,
           id: r[i].Id,  
          Sitename:r[i].Sitename,
          ClientName:r[i].ClientName,
          Status:r[i].Status,
          Department:r[i].Department,
          Requester: r[i].Requester.Title, 
          PID:r[i].PID, 
          Created:fdate,
          CApprover:CApprover, 
          HCreated:fdate1,
          Modified:Mdate,
          HModified:Mdate1
        });  
         
      }  
    
      this.setState({ Projectstatus: FetchProjectDetails });
      this.setState({Sitename:this.props.context.pageContext.web.absoluteUrl});
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

  // public formatDate1(InputDate){
  //   //debugger;
  //   var dt  = InputDate.split("T");
  //    return dt[0];
  // }

  public formatDate1(InputDate: string): string {
    var dt  = InputDate.split("T");
    return dt[0];
  }

  public render(): React.ReactElement<IEasyApprovalAllNotesViewProps> {
    //debugger;
    return (
      <div>  
        <h2 id="Submitted" style={{textAlign:"center",backgroundColor:"#0c78b8",color:"white",display:"block",  padding:'5px 0px', fontSize:'18px'}}>View - All Notes </h2>
        <div className='table-reponsive' style={{overflowX:"auto",width: "100%"}}>               
            <table className='table table-striped table-bordered row-border stripe' id='MNSpfxDatatable'>  
          <thead>  
            <tr>  
              <th>Title </th>  
              <th style={{maxWidth:'150px'}}>Requester</th>
              <th style={{maxWidth:'150px'}}>Department</th>
              <th style={{maxWidth:'150px'}}>Subject</th>
              <th>Client</th>  
               <th>Status</th>   
               <th>Curr Approver</th>              
               <th style={{maxWidth:'120px'}}>Modified Date</th>
              <th>Created</th>  
              <th>ID</th> 
                  </tr>  
          </thead>  
          <tbody id='MNSpfxDatatableBody'>  
            {this.state.Projectstatus && this.state.Projectstatus.map((item, i) => {  
              return [  
                  <tr key={i}>  
                    <td><a href={this.state.Sitename+"/SitePages/NoteApproval.aspx/?uid="+item.PID} >{item.Title}</a></td>  
                    <td style={{maxWidth:'200px'}}>{item.Requester}</td>
                    <td style={{maxWidth:'150px'}}>{item.Department}</td>
                     <td style={{maxWidth:'200px'}}>{item.Subject}</td> 
                    <td>{item.ClientName}</td>
                    <td>{item.Status}</td>  
                    <td>{item.CApprover}</td> 
                     <td><span style={{display:"none"}}>{item.HModified}</span>{item.Modified}</td> 
                    <td><span style={{display:"none"}}>{item.HCreated}</span>{item.Created}</td>  
                    <td>{item.id}</td>  
                  </tr> 
              ];  
            })}  
          </tbody>  
        </table> 
        </div> 
      </div>  
    );  
  }  
  
  public componentWillMount(){
  //debugger;
  }
  public componentDidUpdate() {  
        //debugger;
    $.extend( $.fn.dataTable.defaults, {  
    //   responsive: {
    //     details: {
    //         type: 'column',
    //         target: 'tr'
    //     }
    // }
    });  
    debugger;
  // setTimeout(() => {
      $("#MNSpfxDatatable").DataTable( {  
        
      "info": true,  
      destroy: true,
      retrieve: true,
      //scrollX: true,      
      "pagingType": 'full_numbers',  
      dom: 'lBfrtip',
      buttons: [  
      {extend: 'csv',
      text:"Export to CSV",
      className: 'buttonexcel',
      title:'My Requests'
      },                   
      ],             
      "order": [],  
      "language": {  
      "infoEmpty":sInfotext,  
      "info":sInfotext,  
      "zeroRecords":sZeroRecordsText,  
      "infoFiltered":sinfoFilteredText,  
      "lengthMenu": lengthMenutxt,  
      "search":sSearchtext,  
      "paginate": {  
      "first": firstpage,  
      "last": Lastpage,  
      "next": Nextpage,  
      "previous": Previouspage  
      }      
      }      
    });  
 //   }, 1000);  
  } 
}
