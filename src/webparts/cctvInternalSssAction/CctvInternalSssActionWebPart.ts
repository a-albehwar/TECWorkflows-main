import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CctvInternalSssActionWebPart.module.scss';
import * as strings from 'CctvInternalSssActionWebPartStrings';


import { sp } from "@pnp/sp/presets/all";
import * as $ from 'jquery';
import {  SPHttpClient , SPHttpClientResponse } from '@microsoft/sp-http'; 
import { CCTVInternalList, CommonLinks }  from './../Interfaces/ICCTVInternal';
import * as moment from 'moment';

export interface ICctvInternalSssActionWebPartProps {
  description: string;

  
}
var ITManagerGroupID:any;
let groups: any[] = [];
var IsSSSUser:number;

export default class CctvInternalSssActionWebPart extends BaseClientSideWebPart<ICctvInternalSssActionWebPartProps> {

  private reqItemID:number;
  private reqItemTitle:string;
  private CurrentStatusId: number;

  private AssignedToID:number;
  private TaskUrl:string;
  private TaskUrlDescription:string;

  items: any;

  private StatusCodeApproveId:number = 5;
  private StatusCodeRejectId:number= 6;

  private footageAvailable: string;

  private masterCCTVRequestList: string = "CCTV_Internal_Incident";
  private LogsListname: string = "CCTVInternalIncidentLogs";
  private CCTVFootageDocLibrary:string="CCTVInternalFootage";

  private ITManagerCommentsField: string='ITManagerComments';
  private StatusField:string = 'Status';

  private AssignedToGroupITManager:string='ITManager';
  private AssignedToGroupLegalManager:string="LegalManager";
  private AssignedToGroupSSSTeam:string="System Security Specialist";

  private FootageAvailableText:string="Available";
  private FootageNotAvailableText:string="Not Available";

  private SSSActionUrl:string="/Pages/TecPages/cctv/TecPages/SSSAction.aspx";
  private LegalMgrActionUrl:string="/Pages/TecPages/cctv/CCTVLegalMgr.aspx";
  private CCTVInternalTaskUrl:string="/Pages/TecPages/cctv/ITManagerAction.aspx?ItemID=";

  private MyPendingTaskUrl:string="/Pages/TecPages/cctv/MyTasks.aspx";

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
                  <li  rel="tab1">Request Details</li>
                  <li rel="tab2">Legal Manager</li>
                  <li rel="tab3">IT Manager</li>
                  <li class="active" rel="tab4">SSS Action</li>
                </ul>
                <div   class="tab_container">
                  <h3 class="tab_drawer_heading" rel="tab1">Tab 1</h3>
                  <div id="tab1" class="tab_content">
                  <div class="row gray-box">
                  <div class="col-md-12">

                  <h4>Request and Incident Details</h4>
              
                  <div class="col-lg-4  mb-2">
                      <label id="lblReqtitle" class="form-label">Request Title</label>
                      <input type="text" id="txtRequestTitle" class="form-input" name="txtRequestTitle" aria-disabled="true" disabled="disabled">
                  </div>
                  <div class="col-lg-4  mb-2">
                      <label id="lblReqDate" class="form-label">Date of Request</label>
                      <input type="text" id="txtRequestDate" class="form-input" name="txtRequestDate" aria-disabled="true" disabled="disabled">
                  </div>
                  <div class="col-lg-4  mb-2">
                      <label id="lblIncidentDate" class="form-label">Date of Incident</label>
                      <input type="text" id="txtIncidentDate" class="form-input" name="txtIncidentDate" disabled="disabled">
                  </div>
                  <div class="col-lg-4  mb-2">
                      <label id="lblRequestRequest" class="form-label">Reason for Request</label>
                      <textarea style="height:auto !important" rows="5" cols="5" id="txtReqReason" class="form-input" name="txtReqReason" placeholder="" disabled="disabled"></textarea>
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
                      <!--<textarea style="height:auto !important" rows="5" cols="5" id="txtAddress" class="form-input" name="txtAddress"  placeholder="" aria-disabled="true"></textarea>-->
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
                  <h3 class="tab_drawer_heading" rel="tab3">IT Manager</h3>
                  <div id="tab3" class="tab_content">
                  <div class="row gray-box">
                          <div class="col-md-12">
                  <h4>IT Manager</h4>
                  <div id="ManagerApproval1">
                  <div class="col-lg-4  mb-2">
                      <label id="lblITManagerComments" class="form-label"><b>IT Manager Comments</b></label>
                      <textarea class="form-input" style="height:auto !important" rows="5" cols="5" id="txtITManagerComments" name="txtITManagerComments" placeholder="" disabled="disabled"></textarea>
                  </div>
              </div>
              </div></div>
                  </div>
                  <!-- #tab4 -->
                  <h3 class="d_active tab_drawer_heading" rel="tab4">Tab 4</h3>
                  <div id="tab4" class="tab_content">
                  
                  
                  <div class="dvMain" id="dvMain">
                  
                  
                  <div class="row gray-box">
                          <div class="col-md-12">
                                <div id="div_status_NeedmoreInfo" style="displayStyle">
                                <h4>SSS Action</h4><br/>
                  <div>
                  <div class="col-lg-4 mb-2">
                      <label id="lblFootageStatus" class="form-label"><b>Footage Status</b></label>
                      </div>
                      <div class="col-lg-4 mb-2 vleft">
                      <input  type="radio" id="rd1" name="footageAvailable" value="0" class="form-control"><label for="0" class="form-label">Not Available</label>
                      <input  type="radio" id="rd2" name="footageAvailable" value="1" class="form-control"> <label for="1" class="form-label">Available</label>
                      <span class="error-msg" style="display:none; color:red">* Required</span>
                  </div>
                  </div>
                  <div class="col-lg-4  mb-2" id="dvFootageUpload">
                      <label id="lblupload" class="form-label"><b>Upload Footage</b></label>
                      <div id="spanFootage" class="form-label footage" style="display:none;"></div><br/>
                      <input type="file" id="flUpload" class="form-input" />
                      
                      <span class="error-msg" style="display:none; color:red">* Required</span>
                  </div><br/>
                  <div class="col-lg-4  mb-2">
                      <label id="lblSSSComments" class="form-label"><b>SSS Comments</b></label>
                      <textarea style="height:auto !important" rows="5" cols="5" id="txtSSSComments" class="form-input" name="txtSSSComments" placeholder=""></textarea>
                      <span class="error-msg" style="display:none; color:red">* Required</span>
                  </div>
                  <div class="col-lg-4">
                      <button class="red-btn red-btn-effect shadow-sm  mt-4" id="btnSubmit"><span>Submit</span></button> 
                      <button class="red-btn red-btn-effect shadow-sm  mt-4" id="btnCancel"> <span>Cancel</span></button>
                  </div>
                  <br />
                  </div></div></div>
              </div>
                  </div>
                  <!-- #tab4 --> 
                </div>
                <!-- .tab_container -->
            </div>

         </section> 


     `;

//this._setButtonEventHandlers();
this.PageLoad();
  
    }

  

  

  private PageLoad():void{
    
    this._checkUserInGroup();
    const url : any = new URL(window.location.href);
    const cctvItemID= url.searchParams.get("ItemID");
    this.reqItemID=cctvItemID;
    //let lkstsid= sp.site.rootWeb.lists.getByTitle("SuggestionsBox").items.getById(cctvItemID).select( "*, AssignedDepartment/ID").expand("AssignedDepartment").get();
    this._setButtonEventHandlers();
    
    
    this.GetGroupIDFromGroupName(this.AssignedToGroupITManager);
    this.GetItemDetails();
    
  }

  
  

  private _setButtonEventHandlers(): void {
    const webPart: CctvInternalSssActionWebPart = this;
       this.domElement.querySelector('#btnSubmit').addEventListener('click', (e) => {
        e.preventDefault();
        //var footageAvailable = '';
        //footage not available..
        
         //this.GetGroupIDFromGroupName(this.AssignedToGroupITManager);
         if(this.ValidateFields())
         {
          if ($("input[name='footageAvailable']:checked").val() == 1) {
            //CheckAndCreateFolder();
            //this.footageAvailable = 'Available';
            //this.UpdateItemDetails(this.StatusCodeApproveId, this.footageAvailable);
            this.UploadFiles('Available',this.StatusCodeApproveId);
        }
         //footage  available..
        else if ($("input[name='footageAvailable']:checked").val() == 0) {
          let input = <HTMLInputElement>document.getElementById("flUpload");    
            let file = input.files[0];   
            if (file!=undefined || file!=null){
              this.UploadFiles('Not Available',this.StatusCodeRejectId);
            }
            else
            {
              this.footageAvailable = 'Not Available';
              this.UpdateItemDetails(this.StatusCodeRejectId, this.footageAvailable);

            }                       
        }
      }       
        
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
          var momentCreatedObj = moment(listItems[0].Created);           
          var formatCreatedDate=momentCreatedObj.format('DD-MM-YYYY');
          $('#txtRequestDate').val(formatCreatedDate);
          var momentIncidentObj = moment(listItems[0].DateOfIncident);           
          var formatIncidentDate=momentIncidentObj.format('DD-MM-YYYY');
          formatIncidentDate=formatIncidentDate+"  "+listItems[0].TimeOfIncident;
          $('#txtIncidentDate').val(formatIncidentDate);
          $('#txtReqReason').val(listItems[0].ReasonForRequest);
          if(listItems[0].ITManagerComments)
          {
              $('#txtITManagerComments').val(listItems[0].ITManagerComments);
          }
          if(listItems[0].MoreInfofromLegalMgr)
          {
            $('#txtMoreInfoLegal').val(listItems[0].MoreInfofromLegalMgr);
            //$('#dvMoreInfoLegal').show();
          }
          if(listItems[0].SSSComments)
          {
            $('#txtSSSComments').val(listItems[0].SSSComments);
          }

          if(listItems[0].FootageAvailable)
                    {
                      var footageVal=listItems[0].FootageAvailable;
                      if(footageVal==this.FootageAvailableText)
                      {
                        $("#rd2").prop("checked", true);
                      }
                      else if(footageVal==this.FootageNotAvailableText)
                      {
                        $("#rd1").prop("checked", true);
                      }
                    }

          this.CurrentStatusId =listItems[0].Status.ID;
          //retrieve request initiated: upload documents..
          if(this.CurrentStatusId==2)
          {
              //$('#dvFootageDetails').css("display", "none");
              
              

          }
          else if(this.CurrentStatusId==8)
          {
              //$('#dvFootageDetails').css("display", "block");
              
              //load all documents here..
              this.LoadAllDocuments();
              $('#flUpload').show();
              $('#spanFootage').show();
              $('#txtSSSComments').val('');

          }
          else
          {
              //hide the text buttons..
              $('#dvMain :input').prop("disabled", true);
              this.LoadAllDocuments();
              $('#flUpload').hide();
              $('#spanFootage').show();
              $('#btnSubmit').hide();
              $('#btnCancel').hide();
          }

          $("#tab1").hide();
          $("#tab4").show();
          
  });
});

}


private ValidateFields()
  {    
    var isValid = true;
    if ($("input[name='footageAvailable']:checked").val() == undefined) {
      $('#rd1').nextAll('span:first').css("display", "block");
      isValid = false;
    }
    else if ($("input[name='footageAvailable']:checked").val() == 1) {
      $('#rd1').nextAll('span:first').css("display", "none");
            //check if upload is empty..            
            let input = <HTMLInputElement>document.getElementById("flUpload");    
            let file = input.files[0];   
            if (file==undefined || file==null){      
              $("#flUpload").next("span").css("display", "block");      
              isValid = false;    
            }    
            else    
            {    
              $("#flUpload").next("span").css("display", "none");    
            }
            // if ($('#flUpload')[0].files.length == 0) {
            //     $("#flUpload").next("span").css("display", "block");
            //     isValid = false;
            // }
            // else {
            //     $("#flUpload").next("span").css("display", "none");

            // }
        }
        //check the comments..
        else
        {
            //$("input[name='footageAvailable']").next("span").css("display", "none");
            $('#rd1').nextAll('span:first').css("display", "none");

            if ($("#txtSSSComments").val() == "" || $("#txtSSSComments").val() == undefined) {
                $("#txtSSSComments").next("span").css("display", "block");

                isValid = false;
            }
            else {
                $("#txtSSSComments").next("span").css("display", "none");

            }
          }
        return isValid;
      }

  private UpdateItemDetails(StatusIdValue,FootageAvailableValue){
    var reviewComments = $('#txtSSSComments').val();
        console.log(reviewComments);
        var serverUrl= this.context.pageContext.web.serverRelativeUrl + this.CCTVInternalTaskUrl + this.reqItemID;
        
              sp.site.rootWeb.lists.getByTitle(this.masterCCTVRequestList).items.getById(this.reqItemID).update({
                SSSComments: reviewComments,
                StatusId:StatusIdValue,
                FootageAvailable:FootageAvailableValue,
                AssignedToId: ITManagerGroupID,
                TaskUrl: {
                  "__metadata": { type: "SP.FieldUrlValue" },
                  Description: "IT Manager Action LInk",
                  Url: serverUrl,
                },
              }).then(r=>{

                this.AddItemstoCCTVLogs(this.reqItemTitle,StatusIdValue,this.reqItemID,reviewComments);
                
              }).catch(function(err) {  
                console.log(err);  
              });
        // }
        // else{
        //   //$("#lbl_Innovate_first_comm_err").text(arrLang[lang]['SuggestionBox']['CommentMandatory']);
        // }
  }

  private AddItemstoCCTVLogs(TitleValue,StatusIdValue,RequestIdValue,CommentsValue)
  {
    sp.site.rootWeb.lists.getByTitle(this.LogsListname).items.add({
        Title: TitleValue,
        StatusId: StatusIdValue,
        RequestId: RequestIdValue,
        Comments: CommentsValue
      }).then(r=>{
        //this.updateLogs(vsid,6,$("#Innovate_First_Comments").val());
        console.log("Log Updated Successfully.");
        alert("Task completed successfully.");
        window.location.href= this.context.pageContext.web.serverRelativeUrl+this.MyPendingTaskUrl;

      }).catch(function(err) {  
        console.log(err);  
      });

  }

  private UploadFiles(FootageAvailable,StatusIdValue) {
    let input = <HTMLInputElement>document.getElementById("flUpload");
    let file = input.files[0];
   // var files = document.getElementById('deptfile');
   
    if (file!=undefined || file!=null){
      this.CheckAndCreateFolder();
    //var folderUrl=this.context.pageContext.site.serverRelativeUrl+"/"+this.CCTVFootageDocLibrary+"/"+itemid;
    var folderUrl=this.context.pageContext.site.serverRelativeUrl+"/"+this.CCTVFootageDocLibrary+"/"+this.reqItemTitle;
    sp.site.rootWeb.getFolderByServerRelativeUrl(folderUrl).files.add(file.name, file, true).then((result) => {
      console.log(file.name + " upload successfully!");
        result.file.listItemAllFields.get().then((listItemAllFields) => {
           // get the item id of the file and then update the columns(properties)
          sp.site.rootWeb.lists.getByTitle(this.CCTVFootageDocLibrary).items.getById(listItemAllFields.Id).update({
                      //Title: "My New Title",
                      RequestId:this.reqItemID
          }).then(r=>{
                      console.log(file.name + " properties updated successfully!");
                      this.footageAvailable=FootageAvailable;
                      this.UpdateItemDetails(StatusIdValue,this.footageAvailable);
          });           
      }); 
  }).catch(err => {
      console.log("Error While uploading file...");
      alert(err);
    });
   
    // sp.site.rootWeb.getFolderByServerRelativeUrl(folderUrl).files.add(file.name, file, true).then((result) => {
    //     console.log(file.name + " upload successfully!");
    //       result.file.listItemAllFields.get().then((listItemAllFields) => {
      
    //         sp.site.rootWeb.lists.getByTitle(this.CCTVFootageDocLibrary).items.getById(listItemAllFields.Id).update({
    //                     //Title: 'My New Title',
    //                     RequestId:this.reqItemID,
    //         })
    //         .then(r=>{
    //                     console.log(file.name + " properties updated successfully!");
    //                     this.footageAvailable="Available";
    //                     this.UpdateItemDetails(this.StatusCodeApproveId,this.footageAvailable);
    //                     //alert(arrLang[lang]['SuggestionBox']['SuccessApproved']);
    //           //window.location.href="https://tecq8.sharepoint.com/sites/IntranetDev/"+lang+"/Pages/TecPages/EmployeeSuggestions/AllSuggestions.aspx";
    //         });           
    //     }); 
    // }).catch(err => {
    //   console.log("Error While uploading file...");
    //   alert(err);
    // });
  }
}

  private UploadFiles1() {
    
    let input = <HTMLInputElement>document.getElementById("flUpload");
    let file = input.files[0];
    if (file!=undefined || file!=null){
      
      //create a folder if yes create...
      this.CheckAndCreateFolder();
      var folderUrl=this.context.pageContext.site.serverRelativeUrl+"/"+this.CCTVFootageDocLibrary+"/"+this.reqItemTitle;
      console.log(folderUrl);
 
    //assuming that the name of document library is Documents, change as per your requirement, 
    //this will add the file in root folder of the document library, if you have a folder named test, replace it as "/Documents/test"
    //sp.site.rootWeb.getFolderByServerRelativeUrl("/sites/IntranetDev/SuggestionBoxDocuments").files.add(file.name, file, true).then((result) => {
    sp.site.rootWeb.getFolderByServerRelativeUrl(folderUrl).files.add(file.name, file, true).then((result) => {


      result.file.getItem().then(item => {  
        item.update({  
            //Title: this.reqItemTitle,
            RequestId: this.reqItemID,  
        }).then((myupdate) => {  

          //this.footageAvailable="Available";
          //this.UpdateItemDetails(this.StatusCodeApproveId,this.footageAvailable);
          console.log(file.name + "file properties updated successfully!");
           
        });  
    }); 
        //console.log(file.name + " upload successfully!");
        // console.log(result);
        // result.file.listItemAllFields.get().then((listItemAllFields) => {
        //   console.log(listItemAllFields);
        //      // get the item id of the file and then update the columns(properties)
        //      sp.site.rootWeb.lists.getByTitle(this.CCTVFootageDocLibrary).items.getById(listItemAllFields.Id).update({
        //                 //Title: this.reqItemTitle,
        //                 RequestId: this.reqItemID,
        //               }).then(r=>{
        //                 this.footageAvailable="Available";
        //                 //this.UpdateItemDetails(this.StatusCodeApproveId,this.footageAvailable);
        //                 console.log(file.name + "file properties updated successfully!");
        //               }); 
        //             }); 


                  }).catch(err => {
                    console.log("Error While uploading file...");
                  });
                }
              }
            

  private CheckAndCreateFolder()
  {
    var folderUrl=this.context.pageContext.site.serverRelativeUrl+"/"+this.CCTVFootageDocLibrary+"/"+ this.reqItemTitle;
    sp.site.rootWeb.getFolderByServerRelativeUrl(folderUrl).select('Exists').get().then(data => {
      
        console.log(data.Exists);

      if(data.Exists)
      {
      
      }
      else{
        sp.site.rootWeb.lists.getByTitle(this.CCTVFootageDocLibrary).rootFolder.folders.add(this.reqItemTitle)
        .then(data => {
          console.log("Created Folder successfully.");
        }).catch(err => {
          console.log("Error while creating folder");
        });
      }
     
    }).catch(err => {
        console.log("Error While fetching Folder");
        
    });
  
  
}

  private GetGroupIDFromGroupName(GroupName)
  {
   
    sp.site.rootWeb.siteGroups.getByName(GroupName).get().then(function(result) {  
        ITManagerGroupID=result.Id;
      }).catch(function(err) {  
      console.log(err);  
    });  
  
  }

  // private  CheckUserInGroup()
  // {
  //   let groups1 =   sp.site.rootWeb.currentUser.groups();
  //   var IsLegalTeamMember:number;
  //   if(groups1.length>0){
  //     for(var i=0;i<groups1.length;i++){
  //       groups.push(groups1[i].Title);
  //     }
  //     //console.log(groups);
  //   }
  //   if(groups.length>0)
  //   {
  //     IsLegalTeamMember=$.inArray( "LegalManager", groups ) ;
  //   }
  //   if(IsLegalTeamMember<0){
  //     alert("You are Not Authorized user");
  //     $("#div_status_NeedmoreInfo").hide();
  //     $("#div_status_NeedmoreInfo").hide();
  //     $('#lbl_unauthorized_err').show();
  //     //window.location.href=this.props.weburl;
  //   }
  //   else{
  //     //
  //   }
  // } 

  private async _checkUserInGroup()
  {
    let groups1 = await  sp.site.rootWeb.currentUser.groups();
    console.log( sp.site.rootWeb.currentUser.groups());

    if(groups1.length>0){
      for(var i=0;i<groups1.length;i++){
        groups.push(groups1[i].Title);
      }
      //console.log(groups);
    }
    if(groups.length>0)
    {
      IsSSSUser=$.inArray(this.AssignedToGroupSSSTeam, groups);
    }
    if(IsSSSUser<0){
      alert("You are Not Authorized user");
      //$('.tab_container').hide();
      //$('#dvHistoryMain').hide();
      // $("#div_status_NeedmoreInfo").hide();
      // $("#div_status_NeedmoreInfo").hide();
      // $('#lbl_unauthorized_err').show();
      //window.location.href=this.props.weburl;
      $('#tab4 :input').prop("disabled", true);
      $('#btnSubmit').hide();
      $('btnCancel').hide();
    }
    else{
      //
    }
  } 


  private LoadAllDocuments(){
    this.context.spHttpClient.get(`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${this.CCTVFootageDocLibrary}')/items?$select=*,Request/ID,Folder/ServerRelativeUrl,FileLeafRef,ID,FileSystemObjectType,BaseName,Modified&$expand=Request&$filter=Request/ID%20eq%20${this.reqItemID}`, SPHttpClient.configurations.v1)
        .then(response => {
            console.log(response);

            return response.json()
                .then((items: any): void => {
                    //let listItems: CCTVInternalList[] = items["value"];

                    //$("#txtRequestTitle").val(listItems[0].Title);
                    var htmlSnippet = "";
                    if (items["value"].length > 0) {
                        items["value"].forEach(element => {
                            var docUrl = this.context.pageContext.site.absoluteUrl + "/" + this.CCTVFootageDocLibrary + "/" + this.reqItemTitle + "/" + element.FileLeafRef;
                            htmlSnippet += '<div><a class="footage" href="' + docUrl + '" target="_blank">' + element.FileLeafRef + '<a></div>';
                        });
                    }

                    $('#spanFootage').html(htmlSnippet);                      

                });
        });
}

}
