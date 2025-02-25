import * as React from 'react';
import _styles from './EasyApprovalMyPendingNotesView.module.scss';
import { IReactPnpResponsiveDataTableState } from './DataTableState';
import { IEasyApprovalMyPendingNotesViewProps } from './IEasyApprovalMyPendingNotesViewProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
//import { CurrentUser } from 'sp-pnp-js/lib/sharepoint/siteusers';
import * as $ from 'jquery';
//import { sp } from '@pnp/sp';
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
SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.53/vfs_fonts.js');
require ('../assets/styles.css');
var sSearchtext = 'Search :';
var sInfotext = 'Showing _START_ to _END_ of _TOTAL_ entries';
var sZeroRecordsText = 'No data available in table';
var sinfoFilteredText = "(filtered from _MAX_ total records)";
//var placeholderkeyword = "Keyword";
//var lengthMenutxt = "Show _MENU_ entries";
var firstpage = "First";
var Lastpage = "Last";
var Nextpage = "Next";
var Previouspage = "Previous";

export default class EasyApprovalMyPendingNotesView extends React.Component<IEasyApprovalMyPendingNotesViewProps, IReactPnpResponsiveDataTableState> {
  constructor(props: IEasyApprovalMyPendingNotesViewProps, state: IReactPnpResponsiveDataTableState) {
    debugger;
    super(props);
    this.state = {
      Sitename: '',
      Projectstatus: [{ Title: "", Description: "", id: "", Requester: "", Created: "" }]

    };
    this.fetchdatas = this.fetchdatas.bind(this);
  }
  public componentDidMount() {
    debugger;
    pnp.sp.web.currentUser.get().then((r) => {
        debugger
        let CurrUserEmail = r.Email;
        console.log(CurrUserEmail);    
        this.fetchdatas(CurrUserEmail);
      });
  }

  private fetchdatas(CurrUserEmail: any) {
    debugger;    
    let FetchProjectDetails: any[] = [];    
    // let SiteUrl = this.props.siteUrl;
    // let URL = SiteUrl + "/_api/Web/Lists/GetByTitle('Notes')/items?$select=ID,Title,Status,Subject,PID,Created,ClientName,Modified,Requester/Title,Requester/EMail,Sitename,CurApprover/EMail,CurApprover/Title&$expand=Requester,CurApprover&$orderby=Modified desc&$top=5000";
    // console.log(URL);
    // this.AddRowsCollectionPen(URL).then(respPenRequests => {
    //   var finalPenData = [];
    //   finalPenData = respPenRequests.filter(el => el.Status != 'Approved' && el.Status != 'Rejected' && el.CurApprover.EMail == CurrUserEmail);
    //   $('#spanPending').text(finalPenData.length);
    //   for (let i = 0; i < finalPenData.length; i++) {
    //     var fdate = this.formatDate(finalPenData[i].Created);
    //     var Mdate = this.formatDate(finalPenData[i].Modified);
    //     var fdate1 = this.formatDate1(finalPenData[i].Created);
    //     var Mdate1 = this.formatDate1(finalPenData[i].Modified);

    //     FetchProjectDetails.push({
    //       Title: finalPenData[i].Title,
    //       Subject: finalPenData[i].Subject,
    //       id: finalPenData[i].Id,
    //       Sitename: finalPenData[i].Sitename,
    //       ClientName: finalPenData[i].ClientName,
    //       Requester: finalPenData[i].Requester.Title,
    //       PID: finalPenData[i].PID,
    //       Status: finalPenData[i].Status,
    //       Created: fdate,
    //       HCreated: fdate1,
    //       Modified: Mdate,
    //       HModified: Mdate1
    //     });
    //   }
    //   this.setState({ Projectstatus: FetchProjectDetails });
    //   this.setState({ Sitename: this.props.context.pageContext.web.absoluteUrl });
    // });

    interface RequestItem {
      Id: number;
      Title: string;
      Status: string;
      Subject: string;
      PID: string;
      Created: string; // Assuming these are strings, update if they are Date objects
      Modified: string;
      ClientName: string;
      Sitename: string;
      Requester: { Title: string; EMail: string };
      CurApprover: { EMail: string; Title: string };
    }
    
    let SiteUrl = this.props.siteUrl;
    let URL = SiteUrl + "/_api/Web/Lists/GetByTitle('Notes')/items?$select=ID,Title,Status,Subject,PID,Created,ClientName,Modified,Requester/Title,Requester/EMail,Sitename,CurApprover/EMail,CurApprover/Title&$expand=Requester,CurApprover&$orderby=Modified desc&$top=5000";
    
    console.log(URL);
    
    this.AddRowsCollectionPen(URL).then((respPenRequests: RequestItem[]) => {
      var finalPenData = respPenRequests.filter(
        (el: RequestItem) => el.Status !== 'Approved' && el.Status !== 'Rejected' && el.CurApprover.EMail === CurrUserEmail
      );
    
      $('#spanPending').text(finalPenData.length);
    
      for (let i = 0; i < finalPenData.length; i++) {
        var fdate = this.formatDate(finalPenData[i].Created);
        var Mdate = this.formatDate(finalPenData[i].Modified);
        var fdate1 = this.formatDate1(finalPenData[i].Created);
        var Mdate1 = this.formatDate1(finalPenData[i].Modified);
    
        FetchProjectDetails.push({
          Title: finalPenData[i].Title,
          Subject: finalPenData[i].Subject,
          id: finalPenData[i].Id,
          Sitename: finalPenData[i].Sitename,
          ClientName: finalPenData[i].ClientName,
          Requester: finalPenData[i].Requester.Title,
          PID: finalPenData[i].PID,
          Status: finalPenData[i].Status,
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

  public AddRowsCollectionPen(urlForAllItems: string) // , successCallback, errorCallback)
  {
    var deferred = $.Deferred();
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

  // public formatDate(InputDate) {
  //   var dt = InputDate.split("T");
  //   var dt1 = dt[0].split("-");
  //   var dateOutput = dt1[2] + "/" + dt1[1] + "/" + dt1[0];

  //   return dateOutput;
  // }

  public formatDate(InputDate: string): string {        
    var dt = InputDate.split("T");
    var dt1 = dt[0].split("-");
    var dateOutput = dt1[2] + "/" + dt1[1] + "/" + dt1[0];
    return dateOutput;
  }

  // public formatDate1(InputDate) {
  //   var dt = InputDate.split("T");
  //   return dt[0];
  // }

  public formatDate1(InputDate: string): string {
    var dt  = InputDate.split("T");
    return dt[0];
  }

  public render(): React.ReactElement<IEasyApprovalMyPendingNotesViewProps> {
    return (
      <div>
        <h2 id="Submitted" style={{textAlign:"center",backgroundColor:"#0c78b8",color:"white",display:"block",  padding:'5px 0px', fontSize:'18px'}}>View - My Pending Notes <span id="spanPending" style={{ marginLeft: "10px", padding: "0px 10px", backgroundColor: "red", borderRadius: "50%", fontSize: "16px", fontWeight: "bold", width:'20px', height:'20px' }}></span></h2>
        <div className='table-responsive'>
        <table className='table table-striped table-bordered row-border stripe' id='SpfxDatatable'>
          <thead>
            <tr>
              <th>Title</th>
              <th>Requester</th>
              <th>Subject</th>
              <th>Client</th>
              <th>Status</th>
              <th>Received Date</th>
              <th>Created</th>
              <th>ID</th>
            </tr>
          </thead>
          <tbody id='SpfxDatatableBody'>
            {this.state.Projectstatus && this.state.Projectstatus.map((item, i) => {
              return [
                <tr key={i}>
                  <td><a href={this.state.Sitename + "/SitePages/Checklist.aspx/?uid=" + item.PID} >{item.Title}</a></td>
                  <td>{item.Requester}</td>
                  <td>{item.Subject}</td>
                  <td>{item.ClientName}</td>
                  <td>{item.Status}</td>
                  <td><span style={{ display: "none" }}>{item.HModified}</span>{item.Modified}</td>
                  <td><span style={{ display: "none" }}>{item.HCreated}</span>{item.Created}</td>
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

  public componentWillMount() {

  }
  public componentDidUpdate() {

    $.extend($.fn.dataTable.defaults, {
      // responsive: {
      //   details: {
      //     type: 'column',
      //     target: 'tr'
      //   }
      // }
    });
    $("#SpfxDatatable").DataTable({
      "info": true,
      pageLength: 5,
      lengthMenu: [[5, 10], [5, 10]],
      destroy: true,
      retrieve: true,
      //scrollX: true,
      dom: 'lBfrtip',
      buttons: [
        {
          extend: 'csv',
          text: "Export to CSV",
          title: 'My Requests',
          className: 'buttonexcel',
        },

      ],
      "order": [],
      "language": {
        "infoEmpty": sInfotext,
        "info": sInfotext,
        "zeroRecords": sZeroRecordsText,
        "infoFiltered": sinfoFilteredText,
        "search": sSearchtext,
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
