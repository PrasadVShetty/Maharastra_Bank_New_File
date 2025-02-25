import * as React from 'react';
import _styles from './EasyApprovalMemoView.module.scss';
import { IEasyApprovalMemoViewProps } from './IEasyApprovalMemoViewProps';
import { IReactPnpResponsiveDataTableState } from './DataTableState';  
//import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';  
//import { SiteUser } from '@pnp/sp/site-users'; 
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
//import { CurrentUser } from 'sp-pnp-js/lib/sharepoint/siteusers';
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
//var lengthMenutxt="Show _MENU_ entries";  
var firstpage="First";  
var Lastpage="Last";  
var Nextpage="Next";  
var Previouspage="Previous";

export default class EasyApprovalMemoView extends React.Component<IEasyApprovalMemoViewProps, IReactPnpResponsiveDataTableState> {
  constructor(props: IEasyApprovalMemoViewProps, state: IReactPnpResponsiveDataTableState) {  
    super(props);  
    this.state = {  
      Sitename:'',
      Projectstatus: [{ Title: "", Description: "", id: "", Requester: "", Created: "" }] 
      
    };  
    this.Memofetchdatas = this.Memofetchdatas.bind(this);   
  }  
 public componentDidMount(){  
   debugger;
   pnp.sp.web.currentUser.get().then((r) => {
       debugger
       let CurrUserEmail = r.Email;
       console.log(CurrUserEmail);    
       this.Memofetchdatas(CurrUserEmail);
     });   
  }  
  private Memofetchdatas(CurrUserEmail: any) {  
    debugger;
    // let web = new Web(this.props.context.pageContext.web.absoluteUrl);  
     // let web=pnp.sp.web;
 
    //const list2 =  pnp.sp.web.lists.getByTitle("Notes");  
    let FetchProjectDetails: any[] = [];
    //let WebpartDesc=this.props.description;
   
  pnp.sp.web.lists.getByTitle('MemoWorkflow').items.select('ID,Title,Subject,PID,Created,ClientName,Requester/Title,Recipient/Title').expand('Requester,Recipient').orderBy("Modified",false).top(5000).get().then(r => {  
      for (let i = 0; i < r.length; i++) {  
            var fdate=this.formatDate(r[i].Created);
                      var fdate1=this.formatDate1(r[i].Created);
                     
           FetchProjectDetails.push({  
          Title: r[i].Title,  
          Subject:r[i].Subject,
           id: r[i].Id,  
           ClientName:r[i].ClientName,
          Requester: r[i].Requester.Title, 
          Recipient:r[i].Recipient.Title,
          PID:r[i].PID, 
          Created:fdate,
          HCreated:fdate1,
         
        });  
      }  
      this.setState({ Projectstatus: FetchProjectDetails });  
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

  public render(): React.ReactElement<IEasyApprovalMemoViewProps> {
    return (
      <div>  
        <h2 id="Submitted" style={{textAlign:"center",backgroundColor:"#0c78b8",color:"white",display:"block",  padding:'5px 0px', fontSize:'18px'}}>View - All Memos </h2>
            <table className='table table-striped table-bordered row-border stripe' id='MemoSpfxDatatable'>  
          <thead>  
            <tr>  
              <th>Title</th>  
              <th>Requester</th>
              <th>Subject</th>
              <th>Client</th>  
               <th>Recipient</th>                 
               <th>Created</th>  
              <th>ID</th> 
                  </tr>  
          </thead>  
          <tbody id='MemoSpfxDatatableBody'>  
            {this.state.Projectstatus && this.state.Projectstatus.map((item, i) => {  
              return [  
                  <tr key={i}>  
                    <td><a href={this.state.Sitename+"/SitePages/Memo.aspx/?uid="+item.PID} >{item.Title}</a></td>  
                    <td>{item.Requester}</td>
                     <td>{item.Subject}</td> 
                    <td>{item.ClientName}</td>
                           <td>{item.Recipient}</td> 
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
  $("#MemoSpfxDatatable").DataTable( {  
          "info": true,  
          destroy: true,
          retrieve: true,
          pageLength : 10,
          //scrollX: true,
          lengthMenu: [[5, 10], [5, 10]],
          dom: 'lBfrtip',

          buttons: [  
          {extend: 'csv',
          text:"Export to CSV",
          title:'My Requests',
          className: 'buttonexcel',
          },                   
          ],             
          "order": [],  
          "language": {  
          "infoEmpty":sInfotext,  
          "info":sInfotext,  
          "zeroRecords":sZeroRecordsText,  
          "infoFiltered":sinfoFilteredText,  
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
