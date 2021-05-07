import * as React from 'react';
import styles from './CctvInternalLegalMgr.module.scss';
import { ICctvInternalLegalMgrProps } from './ICctvInternalLegalMgrProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import * as moment from 'moment';
import { CCTVInternalList }  from './../../Interfaces/ICCTVInternal';
import { Items, sp } from '@pnp/sp/presets/all';
import "@pnp/sp/folders";
import { SPComponentLoader } from '@microsoft/sp-loader';

declare var arrLang: any;
declare var lang:string;

const errormsgStyle = {
  color: 'red',
};
const displayStyle = {
  display: 'none',
};
let groups: any[] = [];
var Listname = "CCTV_Internal_Incident";
var LogsListname = "CCTVInternalIncidentLogs";
var DocumentLibraryname = "CCTVInternalFootage";
const url : any = new URL(window.location.href);
const ItemID= url.searchParams.get("ItemID");
var StatusID:any;
var RequestTitle:any;
var ITManagerGroupID:any;
var LegalManagerGroupID:any;
var IsLegalTeamMember:number;
export default class CctvInternalLegalMgr extends React.Component<ICctvInternalLegalMgrProps, {}> {
  public constructor(props: ICctvInternalLegalMgrProps){    
    super(props);    
    this.state ={    
      items:[],
     
    };  
    this.PageLoad();
  } 
  private ITMgrTaskUrl:string="/Pages/cctv/CCTVLegalMgr.aspx?ItemID="+ItemID;
  private LegalMgrTaskUrl:string="/Pages/cctv/CCTVLegalMgr.aspx?ItemID="+ItemID;
  private _externalJsUrl: string = "https://diyarunitedcompany.sharepoint.com/sites/TEC/Style%20Library/TEC/JS/CustomJs.js";

  // adding customjs file before render
  public onInit(): Promise<void> {

    let scriptTag: HTMLScriptElement = document.createElement("script");
    scriptTag.src = this._externalJsUrl;
    scriptTag.type = "text/javascript";
    document.getElementsByTagName("head")[0].appendChild(scriptTag);
  
    //console.log(isinnovateteamMember);
    return Promise.resolve<void>();
  }
  public render(): React.ReactElement<ICctvInternalLegalMgrProps> {
     // let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";    
    SPComponentLoader.loadCss(this._externalJsUrl);
    var weburl=this.props.weburl;
    var langcode=this.props.pagecultureId;
    
    lang=langcode=="en-US"?"en":"ar";
    return (
      
      <div>
        <div>
        <section className={"inner-page-cont"}>
           
           <div className={"Inner-page-title"}>
               <h2 className={"page-heading"}>TABS</h2>
           </div>
           <div className={"container-fluid"} id="Suggestion_Tabs">
                <ul className={"tabs"}>
                  <li className={"active"} key={"tab1"}><a rel={"tab1"}>Request Details</a></li>
                  <li  key={"tab2"}><a rel={"tab2"}>Legal Manager</a></li>
                  <li>IT Manager</li>
                  <li>SSS</li>
                </ul>
                <div className={"tab_container"} style={{ padding: 20}}>
                  <h3 className={"d_active tab_drawer_heading"}  key={"tab1"}>EM</h3>
                  <div id="tab1" className="tab_content">
                            
                      <div className={"row gray-box"}>
                        <div className={"col-md-12"}>
                            
                          <div className={"col-lg-4  mb-2"}>
                            <label id="lbl_empName" className={"form-label"}><b>Request Number</b><span  style={errormsgStyle}>*</span>:</label>
                            <label id="lbl_req_number" className={"form-label"}></label>
                          </div>
                          <div className={"col-lg-4  mb-2"}>
                            <label id="lbl_empName" className={"form-label"}><b>Employee Name</b><span  style={errormsgStyle}>*</span>:</label>
                            <label id="lbl_empName_val" className={"form-label"}></label>
                          </div>
                          <div className="col-lg-4 mb-2">
                            <label id="lbl_empid" className={"form-label"}><b>Employee ID</b><span  style={errormsgStyle}>*</span>:</label>
                            <label id="lbl_empId_val" className={"form-label"}></label>
                          </div>
                          <div className={"col-lg-4  mb-2"}>
                            <label id="lbl_status" className={"form-label"}><b>Status</b><span  style={errormsgStyle}>*</span>:</label>
                            <label id="lbl_status_val" className={"form-label"}></label>
                          </div>
                          <div className={"col-lg-4  mb-2"}>  
                            <label id="lbl_empDept" className={"form-label"}><b>Department</b><span  style={errormsgStyle}>*</span>:</label>
                            <label id="lbl_empDept_val" className={"form-label"}></label>
                          </div>
                          <div className={"col-lg-4  mb-2"}>
                            <label id="lbl_empEmail" className={"form-label"}><b>Email</b><span  style={errormsgStyle}>*</span>:</label>
                            <label id="lbl_empEmail_val" className={"form-label"}></label>
                          </div>
                          <div className={"col-lg-4  mb-2"}>
                            <label id="lbl_empMobile" className={"form-label"}><b>Mobile/Telephone</b><span  style={errormsgStyle}>*</span>:</label>
                            <label id="lbl_empMobile_val" className={"form-label"}></label>
                          </div>

                          <div className="col-lg-4 mb-2">
                            <label id="lbl_date_request" className={"form-label"}><b>Date of Request</b><span  style={errormsgStyle}>*</span>:</label>
                            <label id="lbl_dt_req_val" className={"form-label"}></label>
                          </div>
                          
                          <div className={"col-lg-4  mb-2"}>  
                            <label id="lbl_dt_inc" className={"form-label"}><b>Date of Incident</b><span  style={errormsgStyle}>*</span>:</label>
                            <label id="lbl_dt_inc_val" className={"form-label"}></label>
                          </div>
                          <div className={"col-lg-4  mb-2"}>
                            <label id="lbl_time_of_dt" className={"form-label"}><b>Date of Time</b><span  style={errormsgStyle}>*</span>:</label>
                            <label id="lbl_time_dt_val" className={"form-label"}></label>
                          </div>
                          <div className={"col-lg-4  mb-2"}>
                            <label id="lbl_rsn_req" className={"form-label"}><b>Reason for Request</b><span  style={errormsgStyle}>*</span>:</label>
                            <label id="lbl_rsn_req_val" className={"form-label"}></label>
                          </div>
                       </div>
                      </div>   
                    </div>
                  
                  <h3 className={"tab_drawer_heading"}  key={"tab2"}>EM</h3>
                  <div id="tab2" className="tab_content">
                            
                      <div className={"row gray-box"}>
                          <div className={"col-md-12"}>
                                <div id="div_status_NeedmoreInfo" style={displayStyle}>
                                  <h2>Required More Infomation</h2>
                                  <div  id="div_rb_list">
                                  <div className={"col-lg-12  mb-2"}>   
                                            <label id="lbl_info" className={"form-label"}><b>Information</b></label>
                                  </div> 
                                  <div className={"col-lg-12 mb-2 vleft"}>
                                              <input type="radio" id="rb_Available" checked name="MoreInfo" className={"form-control"} value="1"/>
                                              <label id="lbl_rb_available" className={"form-label"}>Available</label>
                                              <input type="radio" id="rb_NotAvailable" name="MoreInfo" className={"form-control"} value="4"/>
                                              <label  id="lbl_rb_notavailable" className={"form-label"}>Not Available</label> 
                                  </div>
                                  </div>
                                  <div className={"col-lg-4  mb-2"}>  
                                    <label id="lbl_more_info" className={"form-label"}><b>More Infomation</b><span  style={errormsgStyle}>*</span>:</label>
                                    <textarea id="txt_more_info" className={"form-control"} name="InnovateTeamCommnents"></textarea>
                                    <label id="lbl_more_info_err" className={"form-label"} style={errormsgStyle}></label>
                                  </div>
                                  <div className="col-lg-4">
                                  
                                  <button className={"red-btn shadow-sm  mr-3"} id="btnUpdate"   onClick={this.UpdateMoreInfo.bind(this)}>Update</button>
                                  <button className={"red-btn shadow-sm  mr-3"} id="btnCancel"  onClick={(e) => {
                                              e.preventDefault();
                                              window.location.href=weburl+"/Pages/cctv/MyTasks.aspx";
                                              }}>Cancel</button> 
                                  </div> 
                                
                                </div>
                                <div id="div_status_acknowledgment" style={displayStyle}>
                                  <h2>Required Acknowledgment</h2>
                                  <div className="col-lg-4">
                                      <label id="lbl_footage_url" className={"form-label"}><b>Footage Link</b></label>
                                      <div id="div_anchorhtml"></div>         
                                  </div> 
                                  <div className="col-lg-4">
                                  <button className={"red-btn shadow-sm  mr-3"} id="btnReceived"   onClick={this.Acknowledge.bind(this)}>Received</button>
                                  </div> 
                               </div>
                        </div>
                      </div>
                   </div>
                    
                  </div>
                  <h3 className="tab_drawer_heading" >Tab 3</h3>
                  <div id="tab3" className="tab_content">
                          <h2>Tab 3 content</h2>
                            <p>Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.</p>
                  </div>
                  <h3 className="tab_drawer_heading">Tab 4</h3>
                  <div id="tab4" className="tab_content"  >
                  <h2 >Tab 4 content</h2>
                    <p >Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.</p>
                  </div>
                </div> 
                <div id="lbl_unauthorized_err" style={{display:"none", color:"red"}}  >You are not member in LegalManager group.</div>
          </section>
        </div>
        <div>
      </div>
    </div>
        
    );

  }
  private PageLoad()
  {
    this._checkUserInGroup();
    this.getListData();
    this.GetITMgrGroupID("ITManager");
    this.GetLegalMgrGroupID("LegalManager");
  }
  private getRelatedDocuments(reqTitle:string){
    this.props.spHttpClient.get(`${this.props.siteurl}/_api/web/lists/getbytitle('${DocumentLibraryname}')/items?$select=*,Request/ID,Folder/ServerRelativeUrl,FileLeafRef,ID,FileSystemObjectType,BaseName,Modified&$expand=Request&$filter=Request/ID%20eq%20${ItemID}`, SPHttpClient.configurations.v1)      
    .then(response=>{
      return response.json()
      .then((items: any): void => {
          
          var htmlSnippet="";
          if(items["value"].length>0){
            items["value"].forEach(element => {
                htmlSnippet+='<div><a class="footage" href="'+element.ServerRedirectedEmbedUrl+'" target="_blank">'+element.FileLeafRef+'<a></div>';
              });
          }
          else{
            htmlSnippet+="<div><a class='footage' href='#'>No Attachments<a></div>";
          }
          $('#div_anchorhtml').html(htmlSnippet);
        });
      });
     
  }
  private Acknowledge(event){
    sp.site.rootWeb.lists.getByTitle(Listname).items.getById(ItemID).update({
      StatusId:9,//Received & Making empty assigned and task url 
      AssignedToId:LegalManagerGroupID,
      TaskUrl: {
        "__metadata": { type: "SP.FieldUrlValue" },
        Description: "LegalManager  Task Url",
        Url:this.props.weburl+this.LegalMgrTaskUrl,
      },
    }).then(r=>{
      this.updateLogs(ItemID,9, " ");
      alert("Ackowledgment Updated Successfully");
      window.location.href=this.props.weburl+"/Pages/cctv/MyTasks.aspx";
    }).catch(function(err) {  
      console.log(err);  
    });
    event.preventDefault();
  }
  private async _checkUserInGroup()
  {
    let groups1 = await sp.web.currentUser.groups();

    if(groups1.length>0){
      for(var i=0;i<groups1.length;i++){
        groups.push(groups1[i].Title);
      }
      //console.log(groups);
    }
    if(groups.length>0)
    {
      IsLegalTeamMember=$.inArray( "LegalManager", groups ) ;
    }
    if(IsLegalTeamMember<0){
      alert("You are Not Authorized user");
      $("#div_status_NeedmoreInfo").hide();
      $("#div_status_NeedmoreInfo").hide();
      $('#lbl_unauthorized_err').show();
      //window.location.href=this.props.weburl;
    }
    else{
      //
    }
  } 
  private  GetLegalMgrGroupID(groupname:string)
  {
   
       sp.web.siteGroups.getByName(groupname).get().then(function(result) {  
        LegalManagerGroupID=result.Id;
      }).catch(function(err) {  
      console.log(err);  
    });  
  
  }
  private  GetITMgrGroupID(groupname:string)
  {
   
       sp.web.siteGroups.getByName(groupname).get().then(function(result) {  
        ITManagerGroupID=result.Id;
      }).catch(function(err) {  
      console.log(err);  
    });  
  
  }
  private UpdateMoreInfo(event){
   // if(moreinfoId==1){
      var selectedchoice=$("input[name=MoreInfo]:checked").val();
      var more_info_val=$("#txt_more_info").val();
      if(more_info_val==""){
        $("#lbl_more_info_err").text("More Information is mandatory");
      }
      else{
        $("#lbl_more_info_err").text(" ");
        sp.site.rootWeb.lists.getByTitle(Listname).items.getById(ItemID).update({
          MoreInfofromLegalMgr:more_info_val,
          StatusId:selectedchoice,
          AssignedToId:ITManagerGroupID,
          TaskUrl: {
            "__metadata": { type: "SP.FieldUrlValue" },
            Description: "ITManager Task Url",
            Url:this.props.weburl+this.ITMgrTaskUrl,
          },
        }).then(r=>{
          this.updateLogs(ItemID,selectedchoice,more_info_val);
          alert("Task completed successfully");
          window.location.href=this.props.weburl+"/Pages/cctv/MyTasks.aspx";
        }).catch(function(err) {  
          console.log(err);  
      });
      }
   // }
   event.preventDefault();

  }
  private updateLogs(itemid,stsid,stsComments) {
    sp.site.rootWeb.lists.getByTitle(LogsListname).items.add({
      StatusId: stsid,
      RequestId:itemid,
      Title:"CCTV_Int_Req_00"+itemid,
      Comments:stsComments,
    }).then(r=>{
      console.log("added data to history list");

    }).catch(function(err) {  
      console.log(err);  
    });
  }
  
  private getListData() {
    var URL = "";
    URL =`${this.props.siteurl}/_api/web/lists/getbytitle('${Listname}')/items?$select=*,Status/ID,Status/Title,Departments/ID,Departments/Title&$expand=Departments,Status&$filter=ID eq `+ItemID;
    this.props.spHttpClient
      .get(URL, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json().then((items: any): void => {
          
            let listItems: CCTVInternalList[] = items.value;
            listItems.forEach((item: CCTVInternalList) => {
            StatusID=item.Status.ID;
            RequestTitle=item.Title;
            var momentObj = moment(item.Created);           
            var formatCreatedDate=momentObj.format('DD-MM-YYYY');
            var IncidentTimeMoment=moment(item.DateOfIncident);
            var formatIncidentDate=IncidentTimeMoment.format('DD-MM-YYYY');
            var leg_mgr_more_info=item.MoreInfofromLegalMgr!=null?item.MoreInfofromLegalMgr:"";
          
            $("#lbl_status_val").html(item.Status.Title);
            $("#lbl_req_number").html(item.Title);
            $("#lbl_empName_val").html(item.EmployeeName);
            $("#lbl_empId_val").html(item.EmployeeID);
            $("#lbl_empDept_val").html(item.Departments.Title);
            $("#lbl_empEmail_val").html(item.EmailAdress);

            $("#lbl_empMobile_val").html(item.MobileNumber);
            $("#lbl_dt_req_val").html(formatCreatedDate);
            $("#lbl_dt_inc_val").html(formatIncidentDate);

            $("#lbl_time_dt_val").html(item.TimeOfIncident);
            $("#lbl_rsn_req_val").html(item.ReasonForRequest);
              if(StatusID==3 && IsLegalTeamMember>=0)//Needmoreinfo
              {
                $("#div_status_NeedmoreInfo").show();
                $("#div_status_acknowledgment").hide();
                $("#div_rb_list").show();
              }
              else if(StatusID==7 && IsLegalTeamMember>=0)//acknowledge
              {
                $("#div_status_acknowledgment").show();
                $("#div_rb_list").hide();
                $("#div_status_NeedmoreInfo").hide();
              }
              else if(StatusID==9)// status received
              {
                $("#div_status_acknowledgment").show();
                $("#div_rb_list").hide();
                $("#btnReceived").hide();
              }
              else if(StatusID>3 && StatusID!=7)
              {
                $("#div_status_NeedmoreInfo").show();
                $('#txt_more_info').val(leg_mgr_more_info);
                $('#txt_more_info').prop('disabled', true); 
                $("#div_rb_list").hide();
                $("#btnUpdate").hide();
                $("#btnCancel").hide();
              }
             
          });
          this.getRelatedDocuments(RequestTitle);
        }).catch(function(err) {  
          console.log(err);  
        });
      });
  }
  
}
