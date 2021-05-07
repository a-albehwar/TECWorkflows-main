import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { Items, sp } from '@pnp/sp/presets/all';
import styles from './CctvInternalLegalManagerWebPart.module.scss';
import * as strings from 'CctvInternalLegalManagerWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import * as moment from 'moment';
import { CCTVInternalList }  from './../Interfaces/ICCTVInternal';

export interface ICctvInternalLegalManagerWebPartProps {
  description: string;
}
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
let anchorhtml: string ='';
export default class CctvInternalLegalManagerWebPart extends BaseClientSideWebPart<ICctvInternalLegalManagerWebPartProps> {
  private _externalJsUrl: string = "https://tecq8.sharepoint.com/sites/IntranetDev/Style%20Library/TEC/JS/CustomJs.js";

  private ITMgrTaskUrl:string="/Pages/TecPages/cctv/ITManagerAction.aspx?ItemID="+ItemID;
  private LegalMgrTaskUrl:string="/Pages/TecPages/cctv/CCTVLegalMgr.aspx?ItemID="+ItemID;
 
  // adding customjs file before render
  public onInit(): Promise<void> {

    let scriptTag: HTMLScriptElement = document.createElement("script");
    scriptTag.src = this._externalJsUrl;
    scriptTag.type = "text/javascript";
    document.getElementsByTagName("head")[0].appendChild(scriptTag);
   // isinnovateteamMember=this._checkUserInGroup("1-SuggestionsBoxDepartment");
    //console.log(isinnovateteamMember);
    return Promise.resolve<void>();
  }
  public render(): void {
   // var weburl=this.props.weburl;
   // var langcode=this.props.pagecultureId;
    
    this.domElement.innerHTML = `
    <section class="inner-page-cont">
           
         <div class="Inner-page-title">
             <h2 class="page-heading">TABS</h2>
         </div>
         <div class="container-fluid mt-5">
                <ul class="tabs">
                  <li rel="tab1">Request Details</li>
                  <li class="active" rel="tab2">Legal Manager</li>
                  <li rel="tab3">IT Manager</li>
                  <li rel="tab4">SSS</li>
                </ul>
                <div class="tab_container">
                  <h3 class="tab_drawer_heading" rel="tab1">Tab 1</h3>
                  <div id="tab1" class="tab_content" style="display: none;">
                  <div class="row gray-box">
                  <div class="col-md-12">
                    <h4>Employee and Incident Details</h4>
                    <div class="col-lg-4  mb-2"}>
                      <label id="lbl_empName" class="form-label"><b>Request Number</b><span>*</span>:</label>
                      <label id="lbl_req_number" class="form-label"></label>
                    </div>
                    <div class="col-lg-4  mb-2">
                      <label id="lbl_empName" class="form-label"><b>Employee Name</b><span >*</span>:</label>
                      <label id="lbl_empName_val" class="form-label"></label>
                    </div>
                    <div class="col-lg-4 mb-2">
                      <label id="lbl_empid" class="form-label"><b>Employee ID</b><span >*</span>:</label>
                      <label id="lbl_empId_val" class="form-label"></label>
                    </div>
                    <div class="col-lg-4  mb-2">
                      <label id="lbl_status" class="form-label"><b>Status</b><span >*</span>:</label>
                      <label id="lbl_status_val" class="form-label"></label>
                    </div>
                    <div class="col-lg-4  mb-2">  
                      <label id="lbl_empDept" class="form-label"><b>Department</b><span *</span>:</label>
                      <label id="lbl_empDept_val" class="form-label"></label>
                    </div>
                    <div class="col-lg-4  mb-2">
                      <label id="lbl_empEmail" class="form-label"><b>Email</b><span  >*</span>:</label>
                      <label id="lbl_empEmail_val" class="form-label"></label>
                    </div>
                    <div class="col-lg-4  mb-2">
                      <label id="lbl_empMobile" class="form-label"><b>Mobile/Telephone</b><span >*</span>:</label>
                      <label id="lbl_empMobile_val" class="form-label"></label>
                    </div>

                    <div class="col-lg-4 mb-2">
                      <label id="lbl_date_request" class="form-label"><b>Date of Request</b><span >*</span>:</label>
                      <label id="lbl_dt_req_val" class="form-label"></label>
                    </div>
                    
                    <div class="col-lg-4  mb-2">  
                      <label id="lbl_dt_inc" class="form-label"><b>Date of Incident</b><span  >*</span>:</label>
                      <label id="lbl_dt_inc_val" class="form-label"></label>
                    </div>
                    <div class="col-lg-4  mb-2">
                      <label id="lbl_time_of_dt" class="form-label"><b>Date of Time</b><span  >*</span>:</label>
                      <label id="lbl_time_dt_val" class="form-label"></label>
                    </div>
                    <div class="col-lg-4  mb-2">
                      <label id="lbl_rsn_req" class="form-label"><b>Reason for Request</b><span  >*</span>:</label>
                      <label id="lbl_rsn_req_val" class="form-label"></label>
                    </div>    
                 </div>
                </div>   
                  </div>
                  <!-- #tab1 -->
                  <h3 class="d_active tab_drawer_heading" rel="tab2">Tab 2</h3>
                  <div id="tab2" class="tab_content"  style="display:block">
                  <div class="row gray-box">
                          <div class="col-md-12">
                                
                                <div id="div_status_NeedmoreInfo" style="display:none">
                                  <h4>Required More Infomation</h4>
                                  <div  id="div_rb_list">
                                  <div class="col-lg-12  mb-2">   
                                            <label id="lbl_info" class="form-label"><b>Information</b></label>
                                  </div> 
                                  <div class="col-lg-12 mb-2 vleft">
                                              <input type="radio" id="rb_Available" checked name="MoreInfo" class="form-control" value="1"/>
                                              <label id="lbl_rb_available" class="form-label">Available</label>
                                              <input type="radio" id="rb_NotAvailable" name="MoreInfo" class="form-control" value="4"/>
                                              <label  id="lbl_rb_notavailable" class="form-label">Not Available</label> 
                                  </div>
                                  </div>
                                  <div class="col-lg-12  mb-2">  
                                    <label id="lbl_more_info" class="form-label"><b>More Infomation</b><span  style="color:red">*</span>:</label>
                                    <textarea id="txt_more_info" style="height:auto !important" rows="5" cols="5" class="form-control" name="InnovateTeamCommnents"></textarea>
                                    <label id="lbl_more_info_err" class="form-label" style="color:red"></label>
                                  </div>
                                  <div class="col-lg-12">
                                  <button class="red-btn shadow-sm  mr-3" id="btnUpdate">Update</button>
                                  
                                  </div> 
                                
                                </div>
                                <div id="div_status_acknowledgment" style="display:none">
                                  <h4>Required Acknowledgment</h4>
                                  <div class="col-lg-4">
                                      <label id="lbl_footage_url" class="form-label"><b>Footage Link</b></label>
                                      <div id="div_anchorhtml" style="margin: 10px 0px 20px 0px;"></div>         
                                  </div> 
                                  <div class="col-lg-4">
                                  <button class="red-btn shadow-sm  mr-3 mt-4" id="btnReceived">Received</button>
                                  </div> 
                               </div>
                        </div>
                      </div>
                  </div>
                  <!-- #tab2 -->
                  <h3 class="tab_drawer_heading" rel="tab3">Tab 3</h3>
                  <div id="tab3" class="tab_content">
                    <div class="row gray-box">
                        <div class="col-lg-12">
                          <h4>IT Manager Action</h4>
                          <label id="lblITManagerComments" class="form-label"><b>IT Manager Comments</b></label>
                          <textarea disabled="disabled" style="height:auto !important" rows="5" cols="5" id="txtITManagerComments" class="form-input" name="txtITManagerComments"></textarea>
                          <span class="error-msg" style="display:none; color:red">* Required</span>
                      </div>
                     </div>    
                  </div>
                  <!-- #tab3 -->
                  <h3 class="tab_drawer_heading" rel="tab4">Tab 4</h3>
                  <div id="tab4" class="tab_content">
                      <div class="row gray-box">
                        <div class="col-lg-12">
                        <h4>SSS Action</h4>
                         
                          <div class="col-lg-12">
                            <label id="lblFootageDetails" class="form-label"><b>Footage Details</b></label>
                            <div id="spanFootage" " style="margin: 10px 0px 20px 0px;" class="form-label footage"></div>   
                          </div>
                          <div class="col-lg-12  ">
                                <label id="lblSSSComments" class="form-label"><b>SSS Comments</b></label>
                                <textarea disabled="disabled" style="height:auto !important" rows="5" cols="5" id="txtSSSComments" class="form-input" name="txtSSSComments"></textarea>
                          </div>
                          
                          </div> 
                        </div>
                  </div>
                  <!-- #tab4 --> 
                </div>
                <!-- .tab_container -->
            </div>

         </section> 
    `;
    
    this.PageLoad();
    this.setButtonsEventHandlers();
  }

  private PageLoad()
  {
    this._checkUserInGroup();
    
    this.GetITMgrGroupID("ITManager");
    this.GetLegalMgrGroupID("LegalManager");
   
  }
  private setButtonsEventHandlers(): void {
    const webPart: CctvInternalLegalManagerWebPart = this;
    // this.domElement.querySelector('#btnCancel').addEventListener('click', (e) => { 
    //   e.preventDefault();
    //   window.location.href=this.context.pageContext.web.absoluteUrl+"/Pages/TecPages/cctv/MyTasks.aspx";
    //  }); 

        // if(StatusID==3 && IsLegalTeamMember>=0)//Needmoreinfo
        // {
        //   this.domElement.querySelector('#btnUpdate').addEventListener('click', (e) => { 
        //     e.preventDefault();
            
        //     webPart.UpdateMoreInfo();
        //     });
        // }
        // else if(StatusID==7 && IsLegalTeamMember>=0)//acknowledge
        // {
        //   this.domElement.querySelector('#btnReceived').addEventListener('click', (e) => { 
        //     e.preventDefault();
        //     webPart.Acknowledge();
        //    });
        // }
        // else if(StatusID==9)// status received
        // {
          
        // }
        // else if(StatusID>3 && StatusID!=7)
        // {
        // }
     
    this.domElement.querySelector('#btnUpdate').addEventListener('click', (e) => { 
      e.preventDefault();
      
      webPart.UpdateMoreInfo();
     });

     this.domElement.querySelector('#btnReceived').addEventListener('click', (e) => { 
      e.preventDefault();
      webPart.Acknowledge();
     });
    
  }
  private Acknowledge(){
    sp.site.rootWeb.lists.getByTitle(Listname).items.getById(ItemID).update({
      StatusId:9,//Received & Making empty assigned and task url 
      AssignedToId:LegalManagerGroupID,
      TaskUrl: {
        "__metadata": { type: "SP.FieldUrlValue" },
        Description: "LegalManager  Task Url",
        Url:this.context.pageContext.web.absoluteUrl+this.LegalMgrTaskUrl,
      },
    }).then(r=>{
      this.updateLogs(ItemID,9, " ");
      alert("Task Completed Successfully");
      window.location.href=this.context.pageContext.web.absoluteUrl+"/Pages/TecPages/cctv/MyTasks.aspx";
    }).catch(function(err) {  
      console.log(err);  
    });
   // event.preventDefault();
  }
  private  GetITMgrGroupID(groupname:string)
  {
    sp.site.rootWeb.siteGroups.getByName(groupname).get().then(function(result) {  
        ITManagerGroupID=result.Id;
      }).catch(function(err) {  
      console.log(err);  
    });  
  
  }

  private  GetLegalMgrGroupID(groupname:string)
  {
    sp.site.rootWeb.siteGroups.getByName(groupname).get().then(function(result) {  
        LegalManagerGroupID=result.Id;
      }).catch(function(err) {  
      console.log(err);  
    });  
  
  }

  private UpdateMoreInfo(){
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
             Url:this.context.pageContext.web.absoluteUrl+this.ITMgrTaskUrl,
           },
         }).then(r=>{
           this.updateLogs(ItemID,selectedchoice,more_info_val);
           alert("Task completed successfully");
           window.location.href=this.context.pageContext.web.absoluteUrl+"/Pages/TecPages/cctv/MyTasks.aspx";
         }).catch(function(err) {  
           console.log(err);  
       });
       }
    // }
    //event.preventDefault();
 
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
  
  private async _checkUserInGroup()
  {
    let groups1 = await  sp.site.rootWeb.currentUser.groups();

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
      $('#lbl_unauthorized_err').show();
      //window.location.href=this.props.weburl;
    }
    else{
      this.getListData();    }
  } 

  private getListData() {
    var URL = "";
    URL =`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${Listname}')/items?$select=*,Status/ID,Status/Title,Departments/ID,Departments/Title&$expand=Departments,Status&$filter=ID eq `+ItemID;
    this.context.spHttpClient
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
            var it_mgr_Comments=item.ITManagerComments!=null?item.ITManagerComments:"";
            var sss_Comments=item.SSSComments!=null?item.SSSComments:"";
            /*if(items.value[0].AttachmentFiles.length>0){
              for(var i=0;i<items.value[0].AttachmentFiles.length;i++){
                var anchorfileURL=this.context.pageContext.site.absoluteUrl+"/Lists/"+Listname+"/Attachments/"+ItemID+"/"+items.value[0].AttachmentFiles[i].FileNameAsPath.DecodedUrl+"?web=1";
                //console.log(anchorfileURL);
                anchorhtml+='<a href="'+anchorfileURL+'">'+items.value[0].AttachmentFiles[i].FileName+'</a><br>';
              
              }
             }
             */
            $("#tab1").hide();
            $("#tab2").show();
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
            $("#txtITManagerComments").val(it_mgr_Comments);
            $("#txtSSSComments").val(sss_Comments);
            
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
                //$("#btnCancel").hide();
              }
              else if(StatusID==1){
                //  // disabled radio button 
                //   $("input[name='footageAvailable']").each(function(i) {
                //     $(this).attr('disabled', 'disabled');
                // });
                // hide radio button div 
                $("#div_rb_list").hide();
                // showing once more info available completed by legal mgr
                $('#txt_more_info').val(leg_mgr_more_info);
                $('#txt_more_info').prop('disabled', true); 
                $("#div_status_NeedmoreInfo").show();
                $("#btnUpdate").hide();
              }
             
          });
          this.getRelatedDocuments(RequestTitle);
        }).catch(function(err) {  
          console.log(err);  
        });
      });
  }
  
  private getRelatedDocuments(reqTitle:string){
    this.context.spHttpClient.get(`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${DocumentLibraryname}')/items?$select=*,Request/ID,Folder/ServerRelativeUrl,FileLeafRef,ID,FileSystemObjectType,BaseName,Modified&$expand=Request&$filter=Request/ID%20eq%20${ItemID}`, SPHttpClient.configurations.v1)      
    .then(response=>{
      return response.json()
      .then((items: any): void => {
          
          var htmlSnippet="";
          if(items["value"].length>0){
            items["value"].forEach(element => {
                var anchorDocURL=this.context.pageContext.site.absoluteUrl+"/CCTVInternalFootage/"+reqTitle+"/"+element.FileLeafRef;
                htmlSnippet+='<div><a class="footage" href="'+anchorDocURL+'" target="_blank">'+element.FileLeafRef+'<a></div>';
              });
          }
          else{
            htmlSnippet+="<div><a class='footage' href='#'>No Attachments<a></div>";
          }
          $('#div_anchorhtml').html(htmlSnippet);
          $('#spanFootage').html(htmlSnippet);
        });
      });
     
  }

  /*protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
