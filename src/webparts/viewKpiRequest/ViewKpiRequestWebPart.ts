import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { Items, sp } from '@pnp/sp/presets/all';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { IFieldInfo } from "@pnp/sp/fields";

import * as moment from 'moment';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import * as $ from 'jquery';
import { KPIReportRequestItem, TECDepartments,KPIValueItem, KPIDocItem } from './../Interfaces/IKPIRequest';

import styles from './ViewKpiRequestWebPart.module.scss';
import * as strings from 'ViewKpiRequestWebPartStrings';

export interface IViewKpiRequestWebPartProps {
  description: string;
}

let groups: any[] = [];
var fileInfos = [];

let KPIRequestItemGlobal:KPIReportRequestItem;

var KPIListName="KPI Reporting Request";
var DepartmentList:string="TECDepartments";
var KPIConfigList="KPIFields";
var KPIValueList="KPIValues";
var KPIReportLibrary="KPIPerformanceReports";

var KPIAnalystTeam="KPIAnalyst";
var KPIOwnerGroupName:string;

var Status_populateData:number=1;
var Status_reviewData:number=2;
var Status_generateData:number=3;
var Status_rejected:number=4;
var Status_readyToView:number=5;

var IsKPIAnalystMember:number=-1;
var IsKPIOwnerMember:number=-1;
var IsUserMember:number;

const url : any = new URL(window.location.href);
const ItemID= url.searchParams.get("ItemID");


export default class ViewKpiRequestWebPart extends BaseClientSideWebPart<IViewKpiRequestWebPartProps> {



private CurrentStatusId:number;

//private RequestCompleteUrl='/Pages/TecPages/common/RequestComplete1.aspx';
private RequestCompleteUrl='/Pages/TecPages/common/RequestComplete1.aspx?PN=WF/KPIRequsts';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="inner-page-cont" style="margin-top:-30px;"> 
    <div class="container-fluid mt-5">
        
            <div class="row user-info col-md-10 mx-auto col-12 pl-0 pr-0 m-0">
                <h3 class="mb-4 col-12">Request Details</h3>
                <div class="col-md-4 col-12 mb-4">
                    <p>Request Title</p>
                    <input type="text" class="form-control " name="" id="txtRequestTitle" disabled>
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Pre Set Date</p>
                    <input type="text" id="txtPreferenceDate" class="form-input" disabled />
                </div>
               
                <div class="col-md-4 col-12 mb-4">
                    <p>Department</p>
                    <select name="ddlDepartment" id="ddlDepartment" class="form-input" disabled></select>
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Time Period</p>
                    <select name="ddlTimePeriod" id="ddlTimePeriod" class="form-input" disabled></select>
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Period</p>
                    <input type="text" id="txtPeriod" class="form-input" disabled>
                </div>
                <div class="col-md-4 col-12 mb-4">
                <p>Status</p>
                <input type="text" class="form-control" name="" id="txtStatus" disabled>
            </div>

            </div>
        
    </div>
    <div class="container-fluid mt-5">
        <div class="col-md-10 mx-auto col-12">
        <div class="row user-info" id="dvApproverAction">
        <h3 class="mb-4 col-12">Approver Details</h3>

        <div class="col-md-12 col-12 mb-4" id="dvKPIOwner"><div class="col-md-11 col-12 mb-4">
            <p>Data</p>
            <div id="dvExport" style="display:none;"><button class="red-btn red-btn-effect shadow-sm mt-4" id="btnExport"> <span>Export to Excel</span></button></div>
            <div id="dvRepeatKPI">
            <span id="span-error" style="color: red;display: none;">*Please fill all the missing Values and Comments.</span>
            </div>
        </div><div class="col-md-6 col-12 mb-4">
            <p>KPI Owner Comments</p>
            <textarea id="txtKPIOwnerComments" class="form-input" style="height:auto!important" rows="3" cols="5"></textarea>
    </div>

  </div>
          
          <div class="col-md-12 col-12 mb-4" id="dvKPIAnalyst">
          <div class="col-md-6 col-12 mb-4" style="display:none;">
              <p>KPI Analyst Comments</p>
              <textarea id="txtKPIAnalystComments" class="form-input" style="height:auto!important" rows="3" cols="5"></textarea>
  </div>

<div class="col-md-4 col-12 mb-4" id="dvReport">
            <p>Performance Management Report</p>
            <!--<input type="file" class="form-control" name="" id="flUpload">-->
            <div class="input-group" id="dvUploadReport">
                        <input type="text" name="filename" class="form-control" id="file_input" readonly="" placeholder="No file selected">
                        <span class="input-group-btn">
                            <div class="btn file-btn custom-file-uploader">
                            <input type="file" className="form-control" id="flUpload"/>
                                Select a file
                            </div>
                        </span>
                    </div>
            <span class="error-msg" id="uploadMsg" style="color: red;"></span>
            <div id="dvPerformanceReport" style="display:none;"></div>
        </div></div>
        

    </div>
    </div>


    <div class="container-fluid mt-5" id="dvButtonMain">
        <div class="col-md-10 mx-auto col-12">
            <div class="row">
                <div class=" col-12 btnright">
                    <button class="red-btn red-btn-effect shadow-sm mt-4" id="btnSubmit"> <span>Submit</span></button>
                    <button class="red-btn red-btn-effect shadow-sm mt-4" id="btnCancel"> <span>Cancel</span></button>
                </div>
                </div>
            </div>
        </div>
</section>
      `;

      this.LoadDepartments();
      this.LoadChoiceColumn("TimePeriod","#ddlTimePeriod");
      //this.getListData();
      this._checkUserInAnalystGroup();
      this._setButtonEventHandlers();
  }



  private _setButtonEventHandlers(): void{
    const webpart:ViewKpiRequestWebPart=this;
    this.domElement.querySelector('#btnSubmit').addEventListener('click', (e)=>{
      e.preventDefault();
        
      if(this.CurrentStatusId== Status_populateData || this.CurrentStatusId== Status_rejected)
      {
        if(this.ValidatePopulateData())
        {
            this.GetKPIValuesAndUpdate();
        }
        else
          {
            alert("Sorry, Please check your form where some data is not in a valid format.");
          } 
    }
     else if(this.CurrentStatusId== Status_generateData)
      {
          //validate fields..
          if(this.ValidateUploadFields())
          {
            this.CheckAndCreateFolder();
              //this.UpdateAndUploadDocument();
          } 
          else
          {
            alert("Sorry, Please check your form where some data is not in a valid format.");
          }     
      }
      
    });

      this.domElement.querySelector('#btnCancel').addEventListener('click', (e) => {
      e.preventDefault();
      window.location.href=this.context.pageContext.web.serverRelativeUrl;
    });

    this.domElement.querySelector('#flUpload').addEventListener('change', (e) => { 
      //e.preventDefault();
      this.blob();
     });
  }


  private LoadDepartments():void {
  sp.site.rootWeb.lists.getByTitle(DepartmentList).items.select("Title","ID","KPIOwnerId").get()
  .then(function (data) {
    $("#ddlDepartment").append('<option value="0">Please select a department</option>');
    for (var k in data) {
      
      $("#ddlDepartment").append('<option value="' + data[k].ID + '">' +data[k].Title + '</option>');
    }
  
  if(KPIRequestItemGlobal)
  {
    $("#ddlDepartment").val(KPIRequestItemGlobal.Department.ID);
  }
  });
  }


private LoadChoiceColumn(ChoiceColumnName,ControlName)
  {
    var control=document.getElementById(ControlName);
    sp.site.rootWeb.lists.getByTitle(KPIListName).fields.getByInternalNameOrTitle(ChoiceColumnName).get().then((fieldData)=>
    
    {
      if(fieldData['Choices'].length>0)
      {
        $(ControlName).append('<option value="0">Please select..</option>');
        fieldData['Choices'].forEach(element => {
             $(ControlName).append("<option value=\"" +element+ "\">" +element + "</option>");
        });
        if(KPIRequestItemGlobal)        
        {
          $(ControlName).val(KPIRequestItemGlobal[ChoiceColumnName]);
        }
      }

    });
     
  }

  private async _checkUserInKPIOwnerGroup(GroupName)
  {
    //console.log(GroupName);
    let groups1 = await  sp.site.rootWeb.currentUser.groups();

    if(groups1.length>0){
      for(var i=0;i<groups1.length;i++){
        groups.push(groups1[i].Title);
      }
    }
    if(groups.length>0)
    {
      
        IsKPIOwnerMember=$.inArray( GroupName, groups );
      
    }

    
    this.LoadBasedOnPermissions();


    
  } 

  private async _checkUserInAnalystGroup()
  {
    let groups1 = await  sp.site.rootWeb.currentUser.groups();

    if(groups1.length>0){
      for(var i=0;i<groups1.length;i++){
        groups.push(groups1[i].Title);
      }
    }
    if(groups.length>0)
    {
     
        IsKPIAnalystMember=$.inArray( KPIAnalystTeam, groups );
        this.getListData();
    }
    
  } 

  private getListData() {
    var URL = "";
    URL =`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${KPIListName}')/items?$select=*,Author/Name,Author/Title,Department/ID,Department/Title,Status/Title,Status/ID,KPIOwner/ID,KPIOwner/Title&$expand=Author/Id,Department,Status,KPIOwner&$filter=ID eq `+ItemID;
    this.context.spHttpClient
      .get(URL, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json().then((items: any): void => {
                       

            let listItems: KPIReportRequestItem[] = items.value;
            listItems.forEach((item: KPIReportRequestItem) => {
            // check is login user item created user or marcomm team member
            KPIRequestItemGlobal=item;
            this.CurrentStatusId=item.Status.ID; 
            this.LoadHtmlControls(item);
            this.GetDepartmentById(item.Department.ID);
            //if item is populate data or need review -> task assigned is KPIOwner..
            this.LoadKPIValues(item.ID);
            //this._checkUserInGroup(KPIAnalystTeam);

          });
          //this.getRelatedDocuments(RequestTitle);
        }).catch(function(err) {  
          console.log(err);  
        });
      });
  }

  private LoadHtmlControls(kpiRequestItem)
  {
    let kpiItem:KPIReportRequestItem=kpiRequestItem;

    var momentObj = moment(kpiRequestItem.PreSetDate);           
    var formatPreSetDate=momentObj.format('DD-MM-YYYY');
    $('#txtRequestTitle').val(kpiRequestItem.Title);
    $('#txtPreferenceDate').val(formatPreSetDate);
    $('#txtStatus').val(kpiRequestItem.Status.Title);
    $('#ddlDepartment').val(kpiRequestItem.Department.ID);
    $('#ddlTimePeriod').val(kpiRequestItem.TimePeriod);
    $('#txtPeriod').val(kpiRequestItem.Period);

    $('#txtKPIOwnerComments').val(kpiItem.KPIOwnerComments);
    $('#txtKPIAnalystComments').val(kpiItem.AnalystComments);
    if(kpiItem.Status.ID==Status_readyToView)
    {
      this.LoadReports(kpiItem.ID);
      $('#dvUploadReport').hide();
    }
    else
    {
      //$('#dvUploadReport').hide();
    }

    
  }

  private LoadReports(ReqItemID)
  {
    var URL = "";
    URL =`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${KPIReportLibrary}')/items?$select=*,KPIReport/ID,KPIReport/Title,KPIReport/Period,File/ServerRelativeUrl,File/Name&$expand=KPIReport,File&$filter=KPIReportId eq `+ReqItemID;
    this.context.spHttpClient
      .get(URL, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json().then((items: any): void => {
          var htmlString:string='';
            items.value.forEach((item: KPIDocItem) => {
             htmlString+='<a href="'+item.File.ServerRelativeUrl+'">'+item.File.Name+'<a>';

          });

          $('#dvPerformanceReport').append(htmlString);
          
          //this.getRelatedDocuments(RequestTitle);
        }).catch(function(err) {  
          console.log(err);  
        });
      });
    
  }

  private GetDepartmentById(DepartmentId) {
    //console.log("Indide department ID "+DepartmentId);
    var URL = "";
    URL =`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${DepartmentList}')/items?$select=*,KPIOwner/ID,KPIOwner/Title&$expand=KPIOwner/Id&$filter=ID eq `+DepartmentId;
    this.context.spHttpClient
      .get(URL, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json().then((items: any): void => {
            //let listItems: KPIReportRequestItem[] = items.value;
            items.value.forEach((item: TECDepartments) => {
            // check is login user item created user or marcomm team member
            
            //console.log(item.KPIOwner.Title);
            //this._checkUserInGroup(item.KPIOwner.Title);
            this._checkUserInKPIOwnerGroup(item.KPIOwner.Title);

          });
          
          //this.getRelatedDocuments(RequestTitle);
        }).catch(function(err) {  
          console.log(err);  
        });
      });
  }

  private LoadKPIValues(ItemID)
  {
    //Populate Table for KPITable..
    var URL = "";
    URL =`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${KPIValueList}')/items?$select=*,Author/Id,KPIReport/Id,KPIReport/Title&$expand=Author/Id,KPIReport&$filter=KPIReport eq `+ItemID;
    this.context.spHttpClient
      .get(URL, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json().then((items: any): void => {
            //var loginuserid=this.context.pageContext.legacyPageContext["userId"];  

            let listItems: KPIValueItem[] = items.value;
            var htmlTable:string = '<table id="tblKPIValues" class="table table-bordered table-hover footable"/>';
            htmlTable +='<thead><th>Sl No</th><th>Component Name</th><th>System</th><th>Value</th><th>Comments</th></thead>';
            htmlTable+='<tbody>';
            var count:number=1;
            listItems.forEach((item: KPIValueItem) => {
            // check is login user item created user or marcomm team member
            htmlTable+='<td>'+count+'</td>'+'<td>'+item.Title+'</td>'+'<td>'+item.System+'</td>';
            var itemValue=item.Value?item.Value:"";
            var itemComments=item.Comments?item.Comments:"";
            htmlTable+='<td><input type="text"  value="' + itemValue + '" id="txtValue' + count + '"/></td>';
            htmlTable+='<td><input type="text" value="' + itemComments + '" id="txtComments' + count + '"/><input type="hidden" class="hdnKPI" value="' + item.ID + '" id="hdnID' + count + '"/></td>';
            htmlTable+='';
            htmlTable+='</tr>';
            count++;

          });
          htmlTable+='</tbody>';
          htmlTable+='</table>';

          $('#dvRepeatKPI').append(htmlTable);
          if(Status_readyToView==KPIRequestItemGlobal.Status.ID || Status_reviewData==KPIRequestItemGlobal.Status.ID|| Status_generateData==KPIRequestItemGlobal.Status.ID)
          {
            $('#dvRepeatKPI input').prop("disabled",true);
          }

          //this.getRelatedDocuments(RequestTitle);
        }).catch(function(err) {  
          console.log(err);  
        });
      });
  }

  private ValidatePopulateData()
  {
    var isValid = true;

    $("#tblKPIValues tr").each(function () {
      var self = $(this);
      var CompNameValue = self.find("td:eq(1)").text().trim();

      var Value = self.find("td").eq(3).find(":text").val();
      var CommentsValue = self.find("td").eq(4).find(":text").val();

      if (CompNameValue) {
          if (!Value) {
              $('#span-error').css("display", "block");
              isValid = false;
              return isValid;
          }
          if (!CommentsValue) {
              $('#span-error').css("display", "block");
              isValid = false;
              return isValid;
          }
      }
  });

    return isValid;
  }

  private ValidateUploadFields() {
    var isValid = true;
    
    if(fileInfos.length==0)    
          {
            $('#uploadMsg').text("Report is mandatory, must be xlsx,pdf,docx,pptx format and max allowed size is 2.0 MB");      
          isValid = false;    
        }    
        else    
        {  
          $('#uploadMsg').text("");  
        }
    
    
    return isValid;

}


private UpdateAndUploadDocument()
{
  let input = <HTMLInputElement>document.getElementById("flUpload");
    let file = input.files[0];
   // var files = document.getElementById('deptfile');
   
    if (file!=undefined || file!=null){
     // this.CheckAndCreateFolder();
    //var folderUrl=this.context.pageContext.site.serverRelativeUrl+"/"+this.CCTVFootageDocLibrary+"/"+itemid;
    var folderUrl=this.context.pageContext.site.serverRelativeUrl+"/"+KPIReportLibrary+"/"+ KPIRequestItemGlobal.Title;
    sp.site.rootWeb.getFolderByServerRelativeUrl(folderUrl).files.add(file.name, file, true).then((result) => {
      console.log(file.name + " upload successfully!");
        result.file.listItemAllFields.get().then((listItemAllFields) => {
           // get the item id of the file and then update the columns(properties)
          sp.site.rootWeb.lists.getByTitle(KPIReportLibrary).items.getById(listItemAllFields.Id).update({
                      //Title: "My New Title",
                      KPIReportId:KPIRequestItemGlobal.ID,
          }).then(r=>{
                      console.log(file.name + " properties updated successfully!");
                      this.UpdateItemUpdateDetails(Status_readyToView);
          });           
      }); 
  }).catch(err => {
      console.log("Error While uploading file...");
      alert(err);
    });
  }

}

private CheckAndCreateFolder()
{
  var folderUrl=this.context.pageContext.site.serverRelativeUrl+"/"+KPIReportLibrary +"/"+ KPIRequestItemGlobal.Title;
  sp.site.rootWeb.getFolderByServerRelativeUrl(folderUrl).select('Exists').get().then(data => {
    
      //console.log(data.Exists);

    if(data.Exists)
    {
      this.UpdateAndUploadDocument();
    
    }
    else{
      sp.site.rootWeb.lists.getByTitle(KPIReportLibrary).rootFolder.folders.add(KPIRequestItemGlobal.Title)
      .then(data => {
        console.log("Created Folder successfully.");
        this.UpdateAndUploadDocument();
      }).catch(err => {
        console.log("Error while creating folder");
      });
    }
   
  }).catch(err => {
      console.log("Error While fetching Folder");
      
  });


}


private UpdateItemUpdateDetails(StatusIdValue){
  
      
      
      
            sp.site.rootWeb.lists.getByTitle(KPIListName).items.getById(ItemID).update({
              
              StatusId:StatusIdValue,              
              // AssignedToId: ITManagerGroupID,
              // TaskUrl: {
              //   "__metadata": { type: "SP.FieldUrlValue" },
              //   Description: "IT Manager Action LInk",
              //   Url: serverUrl,
              // },
            }).then(r=>{

              alert("Thank You! The request was updated successfully.");
              window.location.href=this.context.pageContext.web.serverRelativeUrl+this.RequestCompleteUrl;

              //this.AddItemstoCCTVLogs(this.reqItemTitle,StatusIdValue,this.reqItemID,reviewComments);
              
            }).catch(function(err) {  
              console.log(err);  
            });
      // }
      // else{
      //   //$("#lbl_Innovate_first_comm_err").text(arrLang[lang]['SuggestionBox']['CommentMandatory']);
      // }
}

private UpdateItemDetails(StatusIdValue){
  var kpiOwnerCommentsVal = $('#txtKPIOwnerComments').val();
            sp.site.rootWeb.lists.getByTitle(KPIListName).items.getById(ItemID).update({
              
              StatusId:StatusIdValue,  
              KPIOwnerComments:kpiOwnerCommentsVal,            
              
            }).then(r=>{
              alert("Thank You! The request was updated successfully.");
              window.location.href=this.context.pageContext.web.serverRelativeUrl+this.RequestCompleteUrl;
              
            }).catch(function(err) {  
              console.log(err);  
            });
}


private GetKPIValuesAndUpdate()
{
  var count=$("#tblKPIValues tr").length;
  var requestUrl=this.context.pageContext.web.serverRelativeUrl+this.RequestCompleteUrl;
  var kpiOwnerCommentsVal = $('#txtKPIOwnerComments').val();
  var countVal:number=1;
  $("#tblKPIValues tr").each(function (i) {
    //countVal++;
    var self = $(this);
    var CompNameValue = self.find("td:eq(1)").text().trim();
    //var SystemValue = self.find("td:eq(2)").text().trim();
    var kpiValue = self.find("td").eq(3).find(":text").val();
    var CommentsValue = self.find("td").eq(4).find(":text").val();
    var IdValue = self.find("td").eq(4).find(":hidden").val();
    
    if (CompNameValue) {
      
      
      sp.site.rootWeb.lists.getByTitle(KPIValueList).items.getById(parseInt(IdValue.toString())).update({
        Value:kpiValue,
        Comments:CommentsValue,          
      }).then(r=>{
        countVal++;

        //this.UpdateItemDetails(Status_reviewData);
        if(count == countVal)
        {
           console.log("Inside update");
            sp.site.rootWeb.lists.getByTitle(KPIListName).items.getById(ItemID).update({
              
              StatusId:Status_reviewData,  
              KPIOwnerComments:kpiOwnerCommentsVal,            
              
            }).then(r=>{
              alert("Thank You! The request was updated successfully.");
              window.location.href=requestUrl;
              
            }).catch(function(err) {  
              console.log(err);  
            });
        }       
      }).catch(function(err) {  
          console.log(err);  
        });
    }
});
// if(count==countVal)
// {
//   this.UpdateItemDetails(Status_reviewData);

// }
// else
// {
//   console.log("Error Ocuured while updating KPI Values - GetKPIValuesAndUpdate()");
// }
}


private LoadBasedOnPermissions()
{
  var loginuserid=this.context.pageContext.legacyPageContext["userId"]; 
  if(loginuserid== KPIRequestItemGlobal.AuthorId || IsKPIAnalystMember>=0 || IsKPIOwnerMember>=0)
            {
              if(this.CurrentStatusId== Status_populateData || this.CurrentStatusId==Status_rejected)
            {
              $('#dvKPIAnalyst').hide();
              if(IsKPIOwnerMember>=0){
                $('#dvButtonMain').show();
                
              }
              else if(IsKPIAnalystMember>=0)
              {
                //disable all controls and hide buttons.
                $('#dvButtonMain').hide();
                $('#dvKPIOwner input').prop("disabled",true);              
                $('#dvKPIOwner textarea').prop("disabled",true);
              }

            }
            else if(this.CurrentStatusId==Status_reviewData)
            {
              $('#dvKPIOwner input').prop("disabled",true);
              $('#dvKPIOwner textarea').prop("disabled",true);
              $('#dvKPIAnalyst').hide();
              //$('#dvExport').show();
              $('#dvButtonMain').hide();
              if(IsKPIAnalystMember>=0)
              {
                
              }
              else if(IsKPIOwnerMember>=0){
                $('#dvKPIAnalyst input').prop("disabled",true);
                $('#dvKPIAnalyst textarea').prop("disabled",true);
                
              }
              
            }
            else if(this.CurrentStatusId==Status_generateData)
            {
              //$('#dvExport').show();
              $('#dvKPIOwner input').prop("disabled",true);
              $('#dvKPIOwner textarea').prop("disabled",true);
              if(IsKPIAnalystMember>=0)
              {
                //disable all controls and show buttons.
                $('#dvButtonMain').show();
                $('#dvKPIAnalyst input').prop("disabled",false);
                $('#dvKPIAnalyst textarea').prop("disabled",false);
                $('#txtKPIAnalystComments').parent().css("display","none");
              }
              else if(IsKPIOwnerMember>=0){
                //hide respective div s and hide buttons
                $('#dvButtonMain').hide();
                $('#dvReport').hide();
                $('#dvKPIAnalyst input').prop("disabled",true);
                $('#dvKPIAnalyst textarea').prop("disabled",true);
                
              }             
            }
            
            else if(this.CurrentStatusId==Status_readyToView)
            {
              
              //both kpi owner and kpi analyst read only view...
              if(IsKPIOwnerMember>=0 || IsKPIAnalystMember>=0){
                // If KPI Owner.
                //show all div s and hide buttons
                $('#dvButtonMain').hide();
                //$('#dvExport').show();
                $('#dvApproverAction input').prop("disabled",true);
                $('#dvApproverAction textarea').prop("disabled",true);
                $('#dvKPIOwner input').prop("disabled",true);
                $('#flUpload').hide();
                $('#dvPerformanceReport').show();                
              }
              
            }
          }
          else{
              // if unauthorized user
              alert("You don't have access,please contact administrator for more info.");
              window.location.href=this.context.pageContext.web.absoluteUrl;
            }
}

private blob() {
  //Get the File Upload control id
  let input = <HTMLInputElement>document.getElementById("flUpload");
  var fileNameHtml="";
  var isValidaformat=true;
  var isValidSize=true;
  var fileCount = input.files.length;
  if(fileCount>0){
  for (var i = 0; i < fileCount; i++) {
     var fileName = input.files[i].name;
     var ext=fileName.replace(/^.*\./, '');
     //Xlsx, PDF, Docx, Pptx
     if ($.inArray(ext, ['xlsx','pdf','docx','pptx']) == -1){
      //$("#spFootage").text("snapshot must be xlsx,pdf,docx,pptx format.");
      
      //$("#file_input").val(fileName);
      fileNameHtml+=fileName+";";
      fileInfos.length=0;
      isValidaformat=false;
     }
    else
     {
      
      var filesize=input.files[0].size;
      const kb = Math.round((filesize / 1024));
      if(kb<=2048){//  checking file must be 2 mb less than
      
      //$("#file_input").val(fileName);
      fileNameHtml+=fileName+";";
      var file = input.files[i];
      var reader = new FileReader();
      reader.onload = (function(file) {
          return function(e) {
            
                fileInfos.push({
                  "name": file.name,
                  "content": e.target.result
                  });
               
                }
          })(file);
      reader.readAsArrayBuffer(file);
      }
      else{
        //$("#spFootage").append("Max allowed footage size is 2.0 MB"); 
        
        fileInfos.length=0;
        isValidSize=false;
      }
    }
  }
  
  if(!isValidSize && !isValidaformat)
  {
    //show error msg for size and format
    $("#uploadMsg").append("Report must be xlsx,pdf,docx,pptx format and Max allowed report size is 2.0 MB"); 
    fileInfos.length=0;
  }
  else if(!isValidaformat)
  {
    //show error msg for format
    $("#uploadMsg").text("Report must be xlsx,pdf,docx,pptx format.");
    fileInfos.length=0;

  }
  else if(!isValidSize)
  {
    //show error msg for size
    $("#uploadMsg").text("Max allowed report size is 2.0 MB"); 
    fileInfos.length=0;

  }
  else
  {
    $("#uploadMsg").text("");
    $("#file_input").val(fileNameHtml);
  }
  //End of for loop
  }
  else{
    $("#file_input").val("No file Selected");
  }
}

}
