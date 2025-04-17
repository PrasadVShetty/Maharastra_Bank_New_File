import * as React from 'react';
import _styles from './EasyApprovalNoteInProgress.module.scss';
import { IReactPnpResponsiveDataTableState } from './DataTableState';  
import { IEasyApprovalNoteInProgressProps } from './IEasyApprovalNoteInProgressProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';  
//import { CurrentUser } from 'sp-pnp-js/lib/sharepoint/siteusers'; 
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
SPComponentLoader.loadCss('https://cdn.datatables.net/responsive/2.2.3/css/responsive.bootstrap.min.css');    
SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.20/css/jquery.dataTables.min.css');  
SPComponentLoader.loadCss('https://cdn.datatables.net/buttons/1.6.0/css/buttons.dataTables.min.css');  
SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.53/vfs_fonts.js')  ;
SPComponentLoader.loadCss('../SiteAssets/css/styles.css');  
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

export default class EasyApprovalNoteInProgress extends React.Component<IEasyApprovalNoteInProgressProps, IReactPnpResponsiveDataTableState> {
  constructor(props: IEasyApprovalNoteInProgressProps, state: IReactPnpResponsiveDataTableState) {  
    super(props);  
    this.state = {  
      Sitename:'',
      Projectstatus: [{ Title: "", Description: "", id: "", Requester: "", Created: "" }] 
      
    };  
    this.Ifetchdatas = this.Ifetchdatas.bind(this);   
  }  
 public componentDidMount(){  
   debugger;
   pnp.sp.web.currentUser.get().then((r) => {
          debugger
          let CurrUserEmail = r.Email;
          console.log(CurrUserEmail);    
          this.Ifetchdatas(CurrUserEmail);
        });
  }  
  private Ifetchdatas(CurrUserEmail: any) {  
    debugger;
    // let web = new Web(this.props.context.pageContext.web.absoluteUrl);  
     // let web=pnp.sp.web;
 
    //const list2 =  pnp.sp.web.lists.getByTitle("Notes");  
    let IFetchProjectDetails: any[] = [];      
    //let WebpartDesc=this.props.description;
    let qstr=window.location.search.split('status=');
    let filterText="Status ne 'Approved' and Status ne 'Rejected'";
    if(qstr[1]=='Inprogress'){
      filterText="Status ne 'Approved' and Status ne 'Rejected'";
      $('#ISubmitted').text('View - In-Progress Notes 1');
    }
    else if(qstr[1]=='Approved'){
      filterText="Status eq 'Approved'";
      $('#ISubmitted').text('View - Approved Notes');
    }
    else if(qstr[1]=='Rejected'){
      filterText="Status eq 'Rejected'";
      $('#ISubmitted').text('View - Rejected Notes');
    }
    else{
      filterText="Status ne 'Approved' and Status ne 'Rejected'";
      $('#ISubmitted').text('View - In-Progress Notes ');
    }
    
    // 
  //  if(WebpartDesc=='Approved'){filterText='Status eq \'Approved\'';}
  pnp.sp.web.lists.getByTitle('ChecklistNote').items.select('ID,Title,Status,Subject,Department,PID,Created,ClientName,Modified,Requester/Title,Sitename,CurApprover/EMail,CurApprover/Title').expand('Requester,CurApprover').filter(filterText).orderBy("Modified",false).top(5000).get().then(r => {  
      for (let i = 0; i < r.length; i++) {  
            var fdate=this.formatDate(r[i].Created);
            var Mdate=this.formatDate(r[i].Modified);
            var fdate1=this.formatDate1(r[i].Created);
            var Mdate1=this.formatDate1(r[i].Modified);
             
          // if(r[i].Id == 195)
          // {
          //   console.log(r[i].Id);
          // }

           IFetchProjectDetails.push({  
          Title: r[i].Title,  
          Subject:r[i].Subject,
           id: r[i].Id,  
          Sitename:r[i].Sitename,
          Status:r[i].Status,
          Department:r[i].Department,
          ClientName:r[i].ClientName,
          Requester: r[i].Requester.Title, 
          //CApprover: r[i].CurApprover.Title, 
          CApprover: (r[i].CurApprover != undefined?r[i].CurApprover.Title:''), 
          PID:r[i].PID, 
          Created:fdate,
          HCreated:fdate1,
          Modified:Mdate,
          HModified:Mdate1
        });  
      }  
      this.setState({ Projectstatus: IFetchProjectDetails });  
      this.setState({Sitename:this.props.context.pageContext.web.absoluteUrl});
    });  
  }  
  // public formatDate(InputDate){
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
  //   var dt  = InputDate.split("T");
  //    return dt[0];
  // }

  public formatDate1(InputDate: string): string {
    var dt  = InputDate.split("T");
    return dt[0];
  }

  public render(): React.ReactElement<IEasyApprovalNoteInProgressProps> {
    return (
      <div className={_styles.NotesInProgress +' '+ 'table-responsive'}>  
        <h2 id="ISubmitted" style={{textAlign:"center",backgroundColor:"#0c78b8",color:"white",display:"block", padding:'5px 0px', fontSize:'18px'}}>In-Progress Notes</h2>
          <table className='table table-striped table-bordered row-border stripe' id='ISpfxDatatable'>  
          <thead>  
          <tr>  
          <th>Title </th>  
          <th>Requester</th>
          <th style={{maxWidth:'150px'}}>Department</th>
          <th>Subject</th>
          <th>Client</th>  
          <th>Status</th>  
          <th>Curr Approver</th>
          <th>Created</th>  
          <th>ID</th> 
          </tr>  
          </thead>  
          <tbody id='ISpfxDatatableBody'>  
            {this.state.Projectstatus && this.state.Projectstatus.map((item, i) => {  
              return [  
                  <tr key={i}>  
                    <td><a href={this.state.Sitename+"/SitePages/Checklist.aspx/?uid="+item.PID} >{item.Title}</a></td>  
                    <td>{item.Requester}</td>
                    <td style={{maxWidth:'150px'}}>{item.Department}</td>
                    <td>{item.Subject}</td> 
                    <td>{item.ClientName}</td>
                    <td>{item.Status}</td>  
                    <td>{item.CApprover}</td> 
                    <td><span style={{display:"none"}}>{item.HCreated}</span>{item.Created}</td>  
                    <td>{item.id}</td>  
                  </tr> 
              ];  
            })}  
          </tbody>  
        </table>  
      </div>  
    );  
  }  
  
  public componentWillMount(){

  }
  public componentDidUpdate() {  
        
    $.extend( $.fn.dataTable.defaults, {  
    //   responsive: {
    //     details: {
    //         type: 'column',
    //         target: 'tr'
    //     }
    // }
    } );  
    $("#ISpfxDatatable").DataTable( {  
    "info": true,  
    destroy: true,
    retrieve: true,
    //   scrollX: true,
    "pagingType": 'full_numbers',  
    dom: 'lBfrtip',

    buttons: [  
    {extend: 'csv',
    //class: "fa fa-pencil",
    className: 'buttonexcel',
    text:"Export to CSV ",
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
  }  
}
