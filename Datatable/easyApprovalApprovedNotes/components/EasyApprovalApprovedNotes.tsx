import * as React from 'react';
import _styles from './EasyApprovalApprovedNotes.module.scss';
import { IEasyApprovalApprovedNotesProps } from './IEasyApprovalApprovedNotesProps';
import { IReactPnpResponsiveDataTableState } from './DataTableState';  
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
SPComponentLoader.loadCss('/sites/EasyApproval/SiteAssets/css/styles.css');  
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
export default class EasyApprovalApprovedNotes extends React.Component<IEasyApprovalApprovedNotesProps, IReactPnpResponsiveDataTableState> {
  constructor(props: IEasyApprovalApprovedNotesProps, state: IReactPnpResponsiveDataTableState) {  
    super(props);  
    this.state = {  
      Sitename:'',
      Projectstatus: [{ Title: "", Description: "", id: "", Requester: "", Created: "" }] 
      
    };  
    this.MNfetchdatas = this.MNfetchdatas.bind(this);   
  }  
 public componentDidMount(){  
   debugger;
   pnp.sp.web.currentUser.get().then((r) => {
       debugger
       let CurrUserEmail = r.Email;
       console.log(CurrUserEmail);    
       this.MNfetchdatas(CurrUserEmail);
     });
  }  
  private MNfetchdatas(CurrUserEmail: any) {  
    debugger;
    // let web = new Web(this.props.context.pageContext.web.absoluteUrl);  
     // let web=pnp.sp.web;
 
    //const list2 =  pnp.sp.web.lists.getByTitle("Notes");  
    let FetchProjectDetails: any[] = [];
    //let WebpartDesc=this.props.description;
     //let filterText="Status eq 'Approved'";
       
    // 
  //  if(WebpartDesc=='Approved'){filterText='Status eq \'Approved\'';}
  //pnp.sp.web.lists.getByTitle('Notes').items.select('ID,Title,Status,Subject,PID,Department,Created,ClientName,Modified,Requester/Title,Sitename,CurApprover/EMail').expand('Requester,CurApprover').filter(filterText).orderBy("Modified",false).top(5000).get().then(r => {  
  /*pnp.sp.web.lists.getByTitle('Notes').items.select('ID,Title,Status,Subject,PID,Department,Created,ClientName,Modified,Requester/Title,Sitename,CurApprover/EMail').expand('Requester,CurApprover').filter(filterText).orderBy("Modified",false).top(5000).get().then(r => {  
      for (let i = 0; i < r.length; i++) {  
            var fdate=this.formatDate(r[i].Created);
            var Mdate=this.formatDate(r[i].Modified);
            var fdate1=this.formatDate1(r[i].Created);
            var Mdate1=this.formatDate1(r[i].Modified);
             
           FetchProjectDetails.push({  
          Title: r[i].Title,  
          Subject:r[i].Subject,
           id: r[i].Id,  
          Sitename:r[i].Sitename,
          ClientName:r[i].ClientName,
          Department:r[i].Department,
          Requester: r[i].Requester.Title, 
          PID:r[i].PID, 
          Created:fdate,
          HCreated:fdate1,
          Modified:Mdate,
          HModified:Mdate1
        });  
      }  
      this.setState({ Projectstatus: FetchProjectDetails });  
      this.setState({Sitename:this.props.context.pageContext.web.absoluteUrl});
    });  */
    interface RequestItem {
      Id: number;
      Title: string;
      Status: string;
      Subject: string;
      PID: string;
      Created: string; // Assuming these are strings, but adjust if they are Date objects
      Modified: string;
      ClientName: string;
      Sitename: string;
      Requester: { Title: string; EMail: string };
      CurApprover: { EMail: string; Title: string };
      Department: string;
    }
    
    let SiteUrl = this.props.siteUrl;
let URL = SiteUrl + "/_api/Web/Lists/GetByTitle('Notes')/items?$select=ID,Title,Status,Subject,PID,Created,ClientName,Modified,Requester/Title,Requester/EMail,Sitename,CurApprover/EMail,CurApprover/Title&$expand=Requester,CurApprover&$orderby=Modified desc&$top=5000";
console.log(URL);

this.AddRowsCollection(URL).then((respMyRequests: RequestItem[]) => {
  var finalData = respMyRequests.filter((el: RequestItem) => el.Status == 'Approved');
  
  for (let i = 0; i < finalData.length; i++) {
    var fdate = this.formatDate(finalData[i].Created);
    var Mdate = this.formatDate(finalData[i].Modified);
    var fdate1 = this.formatDate1(finalData[i].Created);
    var Mdate1 = this.formatDate1(finalData[i].Modified);

    FetchProjectDetails.push({
      Title: finalData[i].Title,
      Subject: finalData[i].Subject,
      id: finalData[i].Id,
      Sitename: finalData[i].Sitename,
      ClientName: finalData[i].ClientName,
      Department: finalData[i].Department,
      Requester: finalData[i].Requester.Title,
      PID: finalData[i].PID,
      Created: fdate,
      HCreated: fdate1,
      Modified: Mdate,
      HModified: Mdate1
    });
  }
  this.setState({ Projectstatus: FetchProjectDetails });
  this.setState({ Sitename: this.props.context.pageContext.web.absoluteUrl });
});

    
  }
  
  public AddRowsCollection(urlForAllItems: string) // , successCallback, errorCallback)
  {
    var deferred = $.Deferred();
    //var response = response || [];
    let response: any[] = [];     
    response = response || [];

    function getListItemsRecursively() {
      
      $.ajax(
        {
          url: urlForAllItems,
          type: "GET",
          headers:
          {
            "Accept": "application/json;odata=verbose",
          },
          //success: function (data) {
          success: ((data)=> {
            response = response.concat(data.d.results);

            if (data.d.__next) {
              urlForAllItems = data.d.__next;
              getListItemsRecursively();
            }
            else {
              deferred.resolve(response);
            }
          }),
          //error: function (err) {
          error: ((err)=> {
            deferred.reject(err);
          }),
        });
    }

    getListItemsRecursively();

    return deferred.promise();
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

  public render(): React.ReactElement<IEasyApprovalApprovedNotesProps> {
    return (
      <div>  
        <h2 id="Submitted" style={{textAlign:"center",backgroundColor:"#0c78b8",color:"white",display:"block",  padding:'5px 0px', fontSize:'18px'}}>View - All Approved Notes </h2>
       <div className='table-responsive1'>
       <table className='table table-striped table-bordered row-border stripe' id='MNSpfxDatatable'>  
          <thead>  
            <tr>  
              <th>Title</th>  
              <th>Requester</th>
              <th>Department</th>
              <th>Subject</th>
              <th>Client</th>  
              <th>Created</th>  
              <th>ID</th> 
                  </tr>  
          </thead>  
          <tbody id='MNSpfxDatatableBody'>  
            {this.state.Projectstatus && this.state.Projectstatus.map((item, i) => {  
              return [  
                  <tr key={i}>  
                    <td><a href={this.state.Sitename+"/SitePages/NoteApproval.aspx/?uid="+item.PID} >{item.Title}</a></td>  
                    <td>{item.Requester}</td>
                    <td>{item.Department}</td>
                     <td>{item.Subject}</td> 
                    <td>{item.ClientName}</td>
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
