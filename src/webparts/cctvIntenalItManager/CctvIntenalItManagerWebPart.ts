import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CctvIntenalItManagerWebPart.module.scss';
import * as strings from 'CctvIntenalItManagerWebPartStrings';
import { sp } from "@pnp/sp/presets/all";
import * as $ from 'jquery';
import {  SPHttpClient ,SPHttpClientResponse } from '@microsoft/sp-http'; 
import { CCTVInternalList, CommonLinks }  from './../Interfaces/ICCTVInternal';
import * as moment from 'moment';


export interface ICctvIntenalItManagerWebPartProps {
  description: string;
}

let groups: any[] = [];
var IsITManagerUser:number;
var AssignedITGroupID:any;
var AssignedSSSGroupID:any;
var AssignedLegalGroupID:any;

export default class CctvIntenalItManagerWebPart extends BaseClientSideWebPart<ICctvIntenalItManagerWebPartProps> {

  items: any;

  private masterCCTVRequestList: string = "CCTV_Internal_Incident";
  private LogsListname: string = "CCTVInternalIncidentLogs";
  private CCTVFootageDocLibrary:string="CCTVInternalFootage";

  private ITManagerCommentsField: string='ITManagerComments';
  private StatusField:string = 'Status';

  private reqItemID:number;
  private reqItemTitle:string;
  private CurrentStatusId: number;

  private StatusCodeApproveId: number = 2;
  private StatusCodeRejectId:number = 3;

  private StatusCodeFinalApproveId:number = 7;
  private StatusCodeFinalRejectId:number = 8;

  private AssignedToGroupITManager:string='ITManager';
  private AssignedToGroupLegalManager:string="LegalManager";
  private AssignedToGroupSSS='System Security Specialist';

  private FootageAvailableText:string="Available";
  private FootageNotAvailableText:string="Not Available";
  private FootageAvailableValue:string="1";
  private FootageNotAvailableValue:string="0";

  private SSSActionUrl:string="/Pages/TecPages/cctv/SSSAction.aspx?ItemID=";
  private SSSAction: string='SSS Action';
  private LegalMgrActionUrl:string="/Pages/TecPages/cctv/CCTVLegalMgr.aspx?ItemID=";
  private LegalAction: string='Legal Manager Action';
  private ITMgrActionUrl:string="/Pages/TecPages/cctv/ITManagerAction.aspx?ItemID=";
  private ITManagerAction: string='IT Manager Action';
  private CCTVInternalTaskUrl:string="/Lists/CCTV_Internal_Incident?ItemID=";

  private MyPendingTaskUrl:string="/Pages/TecPages/cctv/Mytasks.aspx";

  private _externalJsUrl: string = "https://tecq8.sharepoint.com/sites/IntranetDev/Style%20Library/TEC/JS/CustomJs.js";
  //private _externalJsUrl: string = "https://diyarunitedcompany.sharepoint.com/sites/TEC/Style%20Library/TEC/JS/CustomJs.js";

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
    this.domElement.innerHTML = `
    <section class="inner-page-cont">
    <div class="Inner-page-title">
        <h2 class="page-heading">TABS</h2>
    </div>
    <div class="container-fluid mt-5">
        <ul class="tabs">
            <li rel="tab1">Request Details</li>
            <li rel="tab2">Legal Manager</li>
            <li class="active" rel="tab3">IT Manager</li>
            <li rel="tab4">SSS Action</li>
        </ul>
        <div class="tab_container">
            <h3 class="tab_drawer_heading" rel="tab1">Tab 1</h3>
            <div id="tab1" class="tab_content">
                <div class="row gray-box">
                    <div class="col-md-12">

                        <h4>Request and Incident Details</h4>

                        <div class="col-lg-4  mb-2">
                            <label id="lblReqtitle" class="form-label"><b>Request Title</b></label>
                            <input type="text" id="txtRequestTitle" class="form-input" name="txtRequestTitle" aria-disabled="true" disabled="disabled">
                        </div>
                        <div class="col-lg-4  mb-2">
                            <label id="lblReqDate" class="form-label"><b>Date of Request</b></label>
                            <input type="text" id="txtRequestDate" class="form-input" name="txtRequestDate" aria-disabled="true" disabled="disabled">
                        </div>
                        <div class="col-lg-4  mb-2">
                            <label id="lblIncidentDate" class="form-label"><b>Date of Incident</b></label>
                            <input type="text" id="txtIncidentDate" class="form-input" name="txtIncidentDate" disabled="disabled">
                        </div>
                        <div class="col-lg-4  mb-2">
                            <label id="lblRequestRequest" class="form-label">Reason for Request</label>
                            <textarea style="height:auto !important" rows="5" cols="5" id="txtReqReason" class="form-input" name="txtReqReason" disabled="disabled"></textarea>
                        </div>

                        <h4>Employee Details</h4>
                        <div class="col-lg-4  mb-2">
                            <label id="lblEmpName" class="form-label">Employee Name</label>
                            <input type="text" id="txtEmpName" class="form-input" name="txtEmpName" aria-disabled="true" disabled="disabled">
                        </div>
                        <div class="col-lg-4  mb-2">
                            <label id="lblDept" class="form-label">Department</label>
                            <input type="text" id="txtDept" class="form-input" name="txtDept" aria-disabled="true" disabled="disabled">
                        </div>
                        <div class="col-lg-4  mb-2">
                            <label id="lblEmpID" class="form-label">Employee ID</label>
                            <input type="text" id="txtEmpID" class="form-input" name="txtEmpID" aria-disabled="true" disabled="disabled">
                        </div>
                        <div class="col-lg-4  mb-2">
                            <label id="lblAddress" class="form-label">Email Address</label>
                            <!--<textarea style="height:auto !important" rows="5" cols="5" id="txtAddress" class="form-input" name="txtAddress"   aria-disabled="true"></textarea>-->
                            <input type="text" id="txtAddress" class="form-input" name="txtAddress" aria-disabled="true" disabled="disabled">
                        </div>

                        <div class="col-lg-4  mb-2">
                            <label id="lblTitle" class="form-label">Mobile/Tel No</label>
                            <input type="text" id="txtMobileNumber" class="form-input" name="txtMobileNumber" aria-disabled="true" disabled="disabled">
                        </div>


                    </div>
                </div>
            </div>
            <!-- #tab1 -->
            <h3 class="tab_drawer_heading" rel="tab2">Tab 2</h3>
            <div id="tab2" class="tab_content">
                <div class="row gray-box">
                    <div class="col-md-12">
                        <div id="div_legal_Action" style=displayStyle>
                            <h4>Legal Manager Action</h4>

                            <div id="dvMoreInfoLegal">
                                <div class="col-lg-4  mb-2">
                                    <label id="lblMoreInfoLegal" class="form-label">More Information From Legal</label>
                                    <textarea style="height:auto !important" rows="5" cols="5" id="txtMoreInfoLegal" class="form-input" name="txtMoreInfoLegal" placeholder="" disabled="disabled"></textarea>
                                </div>
                            </div>

                        </div>

                    </div>
                </div>
            </div>
            <!-- #tab2 -->
            <h3 class="d_active tab_drawer_heading" rel="tab3">IT Manager</h3>
            <div id="tab3" class="tab_content">
                <div class="row gray-box">
                    <div class="col-md-12">


                        <h4>IT Manager</h4>
                        <div id="dvMoreInfoLegal_it" style="display:none;">
                                <div class="col-lg-4  mb-2">
                                    <label class="form-label"><b>More Information From Legal</b></label>
                                    <textarea style="height:auto !important" rows="5" cols="5" id="txtMoreInfoLegal_it" class="form-input" name="txtMoreInfoLegal" placeholder="" disabled="disabled"></textarea>
                                </div>
                            </div>
                        <div id="footageAvailableIT" style="display:none">
                            <div class="col-lg-4 mb-2">
                                <label class="form-label"><b>Footage Status</b></label>
                            </div>
                            <div class="col-lg-4 mb-2 vleft">
                                <div>
                                <input disabled="disabled" type="radio" id="rd_it1" name="footageAvailable_it" value="0" class="form-control"><label for="0" class="form-label">Not Available</label>
                                <input disabled="disabled" type="radio" id="rd_it2" name="footageAvailable_it" value="1" class="form-control"> <label for="1" class="form-label">Available</label>
                                <span class="error-msg" style="display:none; color:red">* Required</span>
                                </div>
                            </div>


                            <div id="dvFootageDetails_it" style="display:none;">
                                <div class="col-lg-4  mb-2">
                                    <label class="form-label"><b>Footage Details</b></label>
                                    <div id="spanFootage_it" class="form-label footage"></div>
                                </div>
                            </div><br />

                            <div class="col-lg-4  mb-2">
                            <label class="form-label"><b>SSS Comments</b></label>
                            <textarea disabled="disabled" style="height:auto !important" rows="5" cols="5" id="txtSSSComments_it" class="form-input" name="txtSSSComments_it"></textarea>
                        </div>
                        
                        </div>
                        
                        <div id="ManagerApproval1">
                            <div class="col-lg-4  mb-2">
                                <label id="lblITManagerComments" class="form-label"><b>IT Manager Comments</b></label>
                                <textarea style="height:auto !important" rows="5" cols="5" id="txtITManagerComments" class="form-input" name="txtITManagerComments"></textarea>
                                <span class="error-msg" style="display:none; color:red">* Required</span>
                            </div>
                        </div>
                        <div class="col-lg-4">
                            <button class="red-btn red-btn-effect shadow-sm  mt-4" id="btnApprove"> <span>Approve</span></button>
                            <button class="red-btn red-btn-effect shadow-sm  mt-4" id="btnReject"> <span>Need Clarification</span></button>
                            <button class="red-btn red-btn-effect shadow-sm  mt-4" id="btnCancel" style="display:none;"> <span>Cancel</span></button>
                        </div>
                        <br />
                    </div>
                </div>


            </div>
            <!-- #tab3 -->
            <h3 class="tab_drawer_heading" rel="tab4">Tab 4</h3>
            <div id="tab4" class="tab_content">


                <div class="dvMain" id="dvMain">


                    <div class="row gray-box">
                        <div class="col-md-12">
                            <div id="div_status_NeedmoreInfo" style="displayStyle">
                                <h4>SSS Action</h4><br />
                                <div>
                                    <div class="col-lg-4 mb-2">
                                        <label id="lblFootageStatus" class="form-label"><b>Footage Status</b></label>
                                    </div>
                                    <div class="col-lg-4 mb-2 vleft">
                                        <div id="dvFootageRdBtnList">
                                            <input disabled="disabled" type="radio" id="rd1" name="footageAvailable" value="0" class="form-control"><label for="0" class="form-label">Not Available</label>
                                            <input disabled="disabled" type="radio" id="rd2" name="footageAvailable" value="1" class="form-control"> <label for="1" class="form-label">Available</label>
                                            <span class="error-msg" style="display:none; color:red">* Required</span>
                                        </div>
                                    </div>
                                </div>

                                <div id="dvFootageDetails" style="display:none;">
                                    <div class="col-lg-4  mb-2">
                                        <label id="lblFootageDetails" class="form-label"><b>Footage Details</b></label>
                                        <div id="spanFootage" class="form-label footage"></div>
                                    </div>
                                </div><br />
                                <div class="col-lg-4  mb-2">
                                    <label id="lblSSSComments" class="form-label"><b>SSS Comments</b></label>
                                    <textarea disabled="disabled" style="height:auto !important" rows="5" cols="5" id="txtSSSComments" class="form-input" name="txtSSSComments"></textarea>

                                </div>


                            </div>
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
      
  }

  private PageLoad():void{
    
    this._checkUserInGroup();
    const url : any = new URL(window.location.href);
    const cctvItemID= url.searchParams.get("ItemID");
    this.reqItemID=cctvItemID;
    this._setButtonEventHandlers();
    
    this.GetGroupIDFromGroupName(this.AssignedToGroupITManager);
    this.GetGroupIDFromGroupName(this.AssignedToGroupLegalManager);
    this.GetGroupIDFromGroupName(this.AssignedToGroupSSS);

    this.GetItemDetails();
    
  }

//   public render():void
//   {
//       //this.domElement.innerHTML+="";
//       //this._setButtonEventHandlers();
//   }
  private _setButtonEventHandlers(): void {
    const webPart: CctvIntenalItManagerWebPart = this;
       this.domElement.querySelector('#btnApprove').addEventListener('click', (e) => {
                  
        e.preventDefault();
       if(this.ValidateFields())
       {
           if(this.CurrentStatusId==1)
           {
            this.UpdateItemDetails(this.StatusCodeApproveId,AssignedSSSGroupID,this.SSSActionUrl,this.SSSAction);
           }
           if(this.CurrentStatusId==5)
           {
            this.UpdateItemDetails(this.StatusCodeFinalApproveId,AssignedLegalGroupID, this.LegalMgrActionUrl,this.LegalAction);
           }
       }
    });

    this.domElement.querySelector('#btnReject').addEventListener('click', (e) => {
      e.preventDefault();
        if(this.ValidateFields())
        {
            if(this.CurrentStatusId==1)
            {
                this.UpdateItemDetails(this.StatusCodeRejectId,AssignedLegalGroupID, this.LegalMgrActionUrl,this.LegalAction);
            }
            if(this.CurrentStatusId==5)
            {
                this.UpdateItemDetails(this.StatusCodeFinalRejectId,AssignedSSSGroupID,this.SSSActionUrl,this.SSSAction);
            }
        }
    });
    this.domElement.querySelector('#btnCancel').addEventListener('click', (e) => {
      e.preventDefault();
      window.location.href=this.context.pageContext.web.serverRelativeUrl+ this.MyPendingTaskUrl;
        
    });
}
  private GetItemDetails(){
              this.context.spHttpClient.get(`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${this.masterCCTVRequestList}')/items?$select=*,Status/ID,Status/Title,Departments/ID,Departments/Title&$expand=Departments,Status&$filter=ID%20eq%20${this.reqItemID}`, SPHttpClient.configurations.v1)      
              .then(response=>{

                return response.json()
                .then((items: any): void => {
                    let listItems: CCTVInternalList[] = items["value"];
                    this.reqItemTitle=listItems[0].Title;
                    $("#txtRequestTitle").val(listItems[0].Title);
                    $("#txtEmpName").val(listItems[0].EmployeeName);
                    $("#txtEmpID").val(listItems[0].EmployeeID);
                    $('#txtAddress').val(listItems[0].EmailAdress);
                    $('#txtMobileNumber').val(listItems[0].MobileNumber);
                    $('#txtDept').val(listItems[0].Departments.Title);
                    //let datevalue: Date = listItems[0].Created;
                    var momentCreatedObj = moment(listItems[0].Created);           
                    var formatCreatedDate=momentCreatedObj.format('DD-MM-YYYY');
                    formatCreatedDate=formatCreatedDate+" "+listItems[0].TimeOfIncident;
                    //let createdDate: string =formatCreatedDate.toString();
                    $('#txtRequestDate').val(formatCreatedDate);
                    var momentIncidentObj = moment(listItems[0].DateOfIncident);           
                    var formatIncidentDate=momentIncidentObj.format('DD-MM-YYYY');
                    $('#txtIncidentDate').val(formatIncidentDate);
                    $('#txtReqReason').val(listItems[0].ReasonForRequest);
                    if(listItems[0].ITManagerComments)
                    {
                        $('#txtITManagerComments').val(listItems[0].ITManagerComments);
                    }
                    if(listItems[0].MoreInfofromLegalMgr)
                    {
                      $('#txtMoreInfoLegal').val(listItems[0].MoreInfofromLegalMgr);
                      $('#txtMoreInfoLegal_it').val(listItems[0].MoreInfofromLegalMgr);
                      $('#dvMoreInfoLegal').show();
                    }
                    else
                    {
                      //$('#dvMoreInfoLegal').hide();
                    }
                    if(listItems[0].SSSComments)
                    {
                      $('#txtSSSComments').val(listItems[0].SSSComments);
                      $('#txtSSSComments_it').val(listItems[0].SSSComments);
                    }
                    if(listItems[0].FootageAvailable)
                    {
                      var footageVal=listItems[0].FootageAvailable;
                      if(footageVal==this.FootageAvailableText)
                      {
                        $("#rd2").prop("checked", true);
                        $("#rd_it2").prop("checked", true);
                      }
                      else if(footageVal==this.FootageNotAvailableText)
                      {
                        $("#rd1").prop("checked", true);
                        $("#rd_it1").prop("checked", true);
                      }
                    }

                    this.CurrentStatusId =listItems[0].Status.ID;
                    if(this.CurrentStatusId==1)
                    {
                        $('#dvFootageDetails').css("display", "none");
                        $('#dvFootageDetails_it').css("display", "none");
                        $('#footageAvailableIT').css("display","none");
                        $('#btnReject').html('<span>Need Clarification</span>');

                        $('#txtITManagerComments').val("");

                        if(listItems[0].MoreInfofromLegalMgr)
                        {
                          $('#dvMoreInfoLegal_it').show();
                        }
                    }
                    else if(this.CurrentStatusId==5)
                    {
                        $('#dvFootageDetails').css("display", "block");
                        $('#dvFootageDetails_it').css("display", "block");
                        $('#btnReject').html('<span>Reject</span>');
                        $('#footageAvailableIT').css("display","block");
                        //load all documents here..
                        this.LoadAllDocuments();
                        $('#txtITManagerComments').val("");

                    }
                    else
                    {
                        //hide the text buttons..
                        $('#dvMain :input').prop("disabled", true);
                        $('#btnReject').hide();
                        $('#btnApprove').hide();

                        $('#dvFootageDetails').css("display", "block");
                        //load all documents here..
                        this.LoadAllDocuments();
                    }

                    $("#tab1").hide();
                    $("#tab3").show();
                    
            });
        });
    }
  private UpdateItemDetails(StatusIdValue,GroupIdVal, TaskUrlLink, TaskUrlTitle){
    var reviewComments = $('#txtITManagerComments').val();
    var taskUrl=this.context.pageContext.web.serverRelativeUrl +TaskUrlLink+this.reqItemID;
    //console.log(taskUrl);
        
              sp.site.rootWeb.lists.getByTitle(this.masterCCTVRequestList).items.getById(this.reqItemID).update({
                ITManagerComments: reviewComments,
                StatusId:StatusIdValue,
                AssignedToId: GroupIdVal,
                TaskUrl: {
                  "__metadata": { type: "SP.FieldUrlValue" },
                  Description: TaskUrlTitle,
                  Url: taskUrl,
                },
              }).then(r=>{
                this.AddItemstoCCTVLogs(this.reqItemTitle,StatusIdValue,this.reqItemID,reviewComments);
                
                //noty.ShowNotificationWithRedirection("success", "message", "https://diyarunitedcompany.sharepoint.com/sites/TEC");

              }).catch(function(err) {  
                console.log(err);  
              });
        // }
        // else{
        //   //$("#lbl_Innovate_first_comm_err").text(arrLang[lang]['SuggestionBox']['CommentMandatory']);
        // }
  }

  private ValidateFields()
  {
    var isValid = true;
    //check the comments..
    if ($("#txtITManagerComments").val() == "" || $("#txtITManagerComments").val() == undefined) {
        $("#txtITManagerComments").next("span").css("display", "block");

        isValid = false;
    }
    else {
        $("#txtITManagerComments").next("span").css("display", "none");

    }
    return isValid;
}
  
private LoadAllDocuments(){
    this.context.spHttpClient.get(`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${this.CCTVFootageDocLibrary}')/items?$select=*,Request/ID,Folder/ServerRelativeUrl,FileLeafRef,ID,FileSystemObjectType,BaseName,Modified&$expand=Request&$filter=Request/ID%20eq%20${this.reqItemID}`, SPHttpClient.configurations.v1)      
    .then(response=>{
        console.log(response);

      return response.json()
      .then((items: any): void => {
          //let listItems: CCTVInternalList[] = items["value"];
          
          //$("#txtRequestTitle").val(listItems[0].Title);
          var htmlSnippet="";
          if(items["value"].length>0)
          {
            items["value"].forEach(element => {
              var docUrl=this.context.pageContext.site.absoluteUrl+"/"+this.CCTVFootageDocLibrary+"/"+this.reqItemTitle+"/"+element.FileLeafRef;
                htmlSnippet+='<div><a class="footage" href="'+docUrl+'" target="_blank">'+element.FileLeafRef+'<a></div>';
                  //console.log(element.ServerRedirectedEmbedUrl+" - "+element.FileLeafRef);
                  
                  //console.log(element.)
              });

          }

          $('#spanFootage').html(htmlSnippet);
          $('#spanFootage_it').html(htmlSnippet);
 
          
  });
});
}
  private AddItemstoCCTVLogs(TitleValue,StatusIdValue,RequestIdValue,CommentsValue)
  {
    sp.site.rootWeb.lists.getByTitle(this.LogsListname).items.add({
        Title: TitleValue,
        StatusId: StatusIdValue,
        RequestId:RequestIdValue,
        Comments:CommentsValue
      }).then(r=>{
        //this.updateLogs(vsid,6,$("#Innovate_First_Comments").val());
        console.log("Log Updated Successfully.");
        alert("Task completed successfully.");
        window.location.href=this.context.pageContext.web.serverRelativeUrl+ this.MyPendingTaskUrl;
        //alert(arrLang[lang]['SuggestionBox']['SuccessClosed']);
        //window.location.href="https://tecq8.sharepoint.com/sites/IntranetDev/"+lang+"/Pages/TecPages/EmployeeSuggestions/AllSuggestions.aspx";

      }).catch(function(err) {  
        console.log(err);  
      });

  }


  private GetGroupIDFromGroupName(GroupName)
  {
    if(this.AssignedToGroupITManager==GroupName)
        {
            sp.site.rootWeb.siteGroups.getByName(GroupName).get().then(function(result) {          
                AssignedITGroupID=result.Id;
            }).catch(function(err) {  
                console.log(err);  
              }); 
        }
       
        else if(this.AssignedToGroupLegalManager==GroupName)
        {
            
          sp.site.rootWeb.siteGroups.getByName(GroupName).get().then(function(result) {          
                AssignedLegalGroupID=result.Id;
            }).catch(function(err) {  
                console.log(err);  
              });
        }
        else if(this.AssignedToGroupSSS== GroupName)
        {
            
          sp.site.rootWeb.siteGroups.getByName(GroupName).get().then(function(result) {          
                AssignedSSSGroupID=result.Id;
            }).catch(function(err) {  
                console.log(err);  
              });
        }
       
  
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
      IsITManagerUser=$.inArray( this.AssignedToGroupITManager, groups ) ;
    }
    if(IsITManagerUser<0){
      alert("You are Not Authorized to take action for this task.");
      $('#btnApprove').hide();
      $('#btnReject').hide();
      $('#btnCancel').hide();
      //$('#dvMain').hide();
      //$('#dvHistoryMain').hide();
      
      $('#tab3 :input').prop("disabled", true);
      //window.location.href=this.props.weburl;
    }
    else{
      //
    }
  } 
}