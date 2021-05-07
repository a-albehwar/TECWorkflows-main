import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { Items, sp } from '@pnp/sp/presets/all';
import { SPComponentLoader } from '@microsoft/sp-loader';

import * as moment from 'moment';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import * as $ from 'jquery';
import { AttendanceItem } from './../Interfaces/IAttendance';

import styles from './ViewAttendanceRequestWebPart.module.scss';
import * as strings from 'ViewAttendanceRequestWebPartStrings';

const url : any = new URL(window.location.href);
const ItemID= url.searchParams.get("ItemID");

export interface IViewAttendanceRequestWebPartProps {
  description: string;
}

let groups: any[] = [];
var fileInfos = [];

export default class ViewAttendanceRequestWebPart extends BaseClientSideWebPart<IViewAttendanceRequestWebPartProps> {


  private MasterAttendanceList:string='Attendance Process Request';
  private DepartmentList:string='LK_Departments';
  private AttendanceSnapshotLibrary='AttendanceSnapshot';

  private Status_Request_Initiated:number=1;
  private Status_SnapshotAvailable:number=2;
  private Status_SnapshotNotAvailable:number=3;
  private Status_SnapshotMatched:number=4;
  private Status_SnapshotNotMatched:number=5;
  private Status_SnapshotVerified:number=6;



  private SSS_GroupName:string='System Security Specialist';
  private SSS_GroupId:number;

  private Is_SSS_Group:number;

  private HROfficer_GroupName='HR Officer';
  private Is_HROfficer_Group:number;

  private HRManager_GroupName='HR Manager';
  private Is_HRManager_Group:number;

  private IT_Manager_GroupName='IT Manager';
  private Is_ITManager_Group:number;

  private CreatedUserId:number;
  private CurrentStatusId:number;

  private Snapshot_Available='Available';
  private Snapshot_NotAvailable='Not Available';

  private RequestCompleteUrl='/Pages/TecPages/common/RequestComplete.aspx';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="inner-page-cont" style="margin-top:-30px;">
    <div class="container-fluid mt-5">
        
            <div class="row user-info col-md-10 mx-auto col-12 pl-0 pr-0 m-0">
                <h3 class="mb-4 col-12">Employee Details</h3>
                
                <div class="col-md-4 col-12 mb-4">
                    <p>Full Name</p>
                    <input type="text" id="txtEmployeeName" class="form-input" disabled />
                </div>

                <div class="col-md-4 col-12 mb-4">
                    <p>Employee ID</p>
                    <input type="text" id="txtEmployeeID" class="form-input" disabled />
                </div>


                <div class="col-md-4 col-12 mb-4">
                    <p>Department</p>
                    
                    <input type="text" id="txtDepartment" class="form-input" disabled />
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Mobile/Telephone Number</p>
                    <input type="text" id="txtNumber" class="form-input" disabled />
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Email Address</p>
                    <input type="text" id="txtEmail" class="form-input" disabled />

                </div>
                
                <div class="col-md-4 col-12 mb-4">
                    <p>Request Created Date</p>
                    <input type="text" id="txtCreatedDate" class="form-input" disabled />
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Status</p>
                    <input type="text" id="txtStatus" class="form-input" disabled />
                </div>
                
                <h3 class="mb-4 col-12">Request Details</h3>
                <div class="col-md-4 col-12 mb-4">
                    <p>Request Title</p>
                    <input type="text" id="txtRequestTitle" class="form-input" disabled />
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Date and Time of Request</p>
                    <input type="text" id="txtDateOfRequest" class="form-input" disabled />
                    
                </div>

                <div class="col-md-4 col-12 mb-4">
                    <p>Duration (Minutes)</p>
                    <input type="text" id="txtTimeOfAbsence" class="form-input" disabled />
                </div>
                <div class="col-md-6 col-12 mb-4">
                    <p>Reason for Request</p>
                    <textarea id="txtReasonForRequest" class="form-input" style="height:auto!important" rows="3" cols="5" disabled></textarea>
                </div>

            </div>
        
    </div>
    <div class="container-fluid mt-5" id="dvSnapshotDetails">
        <div class="col-md-10 mx-auto col-12">
            <div class="row user-info">
                <h3 class="mb-4 col-12">Snapshot</h3>
                
                <div class="col-md-6 col-12 mb-4" id='dvSnapshot'>
                    <!--<input type="file" id="flUpload" class="form-input" multiple />-->
                    <p>Upload snapshot<span  style="color:red">*</span></p>
                    <div class="input-group">
                        <input type="text" name="filename" class="form-control" id="file_input" readonly="" placeholder="No file selected">
                        <span class="input-group-btn">
                            <div class="btn file-btn custom-file-uploader">
                            <input type="file" className="form-control" accept="image/x-png,image/gif,image/jpeg,image/jpg" id="flUpload" multiple/>
                                Select one/multiple files
                            </div>
                        </span>
                    </div>
                    <span id='spFootage' class="form-label" style="color:red;"></span>
                    </div> 
                    <div class="col-md-8 col-12 mb-4" id="dvUpload">

                    </div>
                  
                <!--<div class="col-md-4 col-12 mb-4 alert alert-info" id="dvHyperlink" style="display:none;">
                  Close the task from  <b><a href="https://emea.flow.microsoft.com/manage/environments/Default-d4c2002c-5752-43e7-854d-fe00ca0a181c/approvals/received" target="_blank">Approval Center 
                  </a></b>after submitting the form here.
                </div>
                <h3 class="mb-4 col-12"></h3>
                <div class="col-md-6 col-12 mb-4">
                    <p>SSS Comments</p>
                    <textarea id="txtSSSComments" class="form-input inputDisabled" style="height:auto!important" rows="3" cols="5"></textarea>
                    <span class="err-msg" style="color:red;display: none;">* Required</span>
                </div>-->           
            </div>
        </div>
    </div>
    <div class="container-fluid mt-5" id="dvButtonMain" style="display:none;">
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

      //this.GetMasterListItem();
      this._checkUserInGroup();
      this._setButtonEventHandlers();
  }


  private GetMasterListItem() {
    var URL = "";
    URL =`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${this.MasterAttendanceList}')/items?$select=*,Author/Name,Author/Title,Department/ID,Department/Title,Status/Title,Status/ID&$expand=Author/Id,Department,Status&$filter=ID eq `+ItemID;
    this.context.spHttpClient
      .get(URL, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json().then((items: any): void => {
          
                       

            let listItems: AttendanceItem[] = items.value;
            listItems.forEach((item: AttendanceItem) => {

            var loginuserid=this.context.pageContext.legacyPageContext["userId"];
            this.CreatedUserId=item.AuthorId;
            this.CurrentStatusId=item.Status.ID;
            
            if(loginuserid==this.CreatedUserId || this.Is_SSS_Group>=0 || this.Is_ITManager_Group>=0|| this.Is_HRManager_Group>=0|| this.Is_HROfficer_Group>=0)
            {
              this.LoadHtmlControls(item);
            }
            else
            {
              alert("You are not an authorised user to view this task.");
              window.location.href=this.context.pageContext.web.serverRelativeUrl;

            }

          });
          //this.getRelatedDocuments(RequestTitle);
        }).catch(function(err) {  
          console.log(err);  
        });
      });
  }

  private LoadHtmlControls(MasterItem)
  {
      let attendanceItem:AttendanceItem=MasterItem;
      $('#txtRequestTitle').val(attendanceItem.Title);
      $('#txtEmployeeName').val(attendanceItem.Author.Title);
      $('#txtEmployeeID').val(attendanceItem.EmployeeID);
      $('#txtNumber').val(attendanceItem.ContactNumber);
      $('#txtDepartment').val(attendanceItem.Department.Title);
      $('#txtEmail').val(attendanceItem.Email);
      var timeofreq=attendanceItem.TimeofRequest !=null?attendanceItem.TimeofRequest:"";
      var momentObj=moment(attendanceItem.DateofRequest);
      var formattedReqDate=momentObj.format('DD-MM-YYYY');
      $('#txtDateOfRequest').val(formattedReqDate +" "+timeofreq);
      var momentObjCreated=moment(attendanceItem.Created);
      var formattedCreatdDate=momentObjCreated.format('DD-MM-YYYY hh:mm a');
      $('#txtCreatedDate').val(formattedCreatdDate);
      $('#txtTimeOfAbsence').val(attendanceItem.TimeofAbsence);
      $('#txtReasonForRequest').val(attendanceItem.ReasonForRequest);
      //$('#txtSSSComments').val(attendanceItem.SSSComments);
      $('#txtStatus').val(attendanceItem.Status.Title);
      // if(attendanceItem.Status==this.sta)
      // {
       
      //   $('#dvSnapshot').show();
      // }
      // else if(attendanceItem.SnapshotAvailable==this.Snapshot_NotAvailable)
      // {
      //   $('#decision2').attr('checked','true');
      //   $('#dvSnapshot').hide();
      // }
      

      if(attendanceItem.Status.ID == this.Status_Request_Initiated && this.Is_SSS_Group>=0)
      {
          //new request and upload attachments.
          $('#flUpload').show();
          $('#dvSnapshot').show();
          $('#dvButtonMain').show();
      }
      else if(attendanceItem.Status.ID == this.Status_SnapshotNotMatched && this.Is_SSS_Group>=0)
      {
        $('#flUpload').show();
        $('#dvSnapshot').show();
        $('#dvUpload').show();
        $('#dvButtonMain').show();
        //load html for attachments..
        this.LoadAllDocuments(false);
      }
      else
      {
        $('#flUpload').hide();
        $('#dvSnapshot').hide();
        $('#dvUpload').show();
        $('.inputDisabled').prop("disabled",true);
        //$('.inputDisabled textarea').prop("disabled",true);
        //load html for attachments.
        this.LoadAllDocuments(true);
        
      }

  }

  private ValidateFields()
  {
      var isValid=true;
      //let input = <HTMLInputElement>document.getElementById("flUpload");    
        //let file = input.files[0];   
        //if (file==undefined || file==null){  
          if(fileInfos.length==0)    
          {
            $('#spFootage').text("Snapshot is mandatory, must be jpg,jpeg,png,gif format and max allowed size is 10.0 MB");      
          isValid = false;    
        }    
        else    
        {  
          $('#spFootage').text("");  
        }

      return isValid;
  }

  private _setButtonEventHandlers(): void {
    const webPart: ViewAttendanceRequestWebPart = this;
       
    this.domElement.querySelector('#btnSubmit').addEventListener('click', (e) => {
        e.preventDefault();
        
         if(this.ValidateFields())
         {
           if(this.CurrentStatusId==this.Status_Request_Initiated)
          {
            this.CheckAndCreateFolder();
            //this.UploadFiles();
          }
          else if(this.CurrentStatusId==this.Status_SnapshotNotMatched)
          {
            this.DisableDocumentsAndUpload();
          }
            
            
         
          }
         
        
    });
    
    this.domElement.querySelector('#flUpload').addEventListener('change', (e) => { 
      //e.preventDefault();
      this.blob();
     });

  }

  private UploadFiles()
  {
    let input = <HTMLInputElement>document.getElementById("flUpload");
    var updateCount:number=0;
    var fileCount=input.files.length;   
    var folderUrl=this.context.pageContext.site.serverRelativeUrl+"/"+this.AttendanceSnapshotLibrary+"/"+ItemID;
    //this.CheckAndCreateFolder();
    //if(this.CheckAndCreateFolder())
    {
        for(var i=0;i<fileCount;i++)
        {
            var file=input.files[i];
            sp.site.rootWeb.getFolderByServerRelativeUrl(folderUrl).files.add(file.name, file, true).then((result) => {
                console.log(file.name + " upload successfully!");
                result.file.listItemAllFields.get().then((listItemAllFields) => {
                    // get the item id of the file and then update the columns(properties)
                    sp.site.rootWeb.lists.getByTitle(this.AttendanceSnapshotLibrary).items.getById(listItemAllFields.Id).update({
                      
                      RequestId:ItemID,
                      IsActive:true,
                    }).then(r=>{
                        console.log(file.name + " properties updated successfully!");
                        updateCount++;
                        if(updateCount== fileCount)
                        {
                          
                          alert("Thank you ! The request was updated successfully");
                          window.location.href=this.context.pageContext.web.serverRelativeUrl+this.RequestCompleteUrl;
                        }
                    });           
                }); 
            }).catch(err => {
                console.log("Error While uploading file...");
                //alert(err);
                
            });
        }
        
    }
  }


  

  private  CheckAndCreateFolder()
  {
      //var folderExists=false;
    var folderUrl=this.context.pageContext.site.serverRelativeUrl+"/"+this.AttendanceSnapshotLibrary+"/"+ ItemID;
    sp.site.rootWeb.getFolderByServerRelativeUrl(folderUrl).select('Exists').get().then(data => {
      
      console.log(data.Exists);
      if(data.Exists)
      {
        //folderExists=true;
        //return folderExists;
      
      }
      else{
        sp.site.rootWeb.lists.getByTitle(this.AttendanceSnapshotLibrary).rootFolder.folders.add(ItemID)
        .then(data => {
        //folderExists=true;
        //return folderExists;

          console.log("Created Folder successfully.");
        }).catch(err => {
          console.log("Error while creating folder");
          //return folderExists;
        });
      }
      this.UploadFiles();
     
    }).catch(err => {
        //folderExists=false;
        console.log("Error While fetching Folder");
        
    });
    //return folderExists;
}

private LoadAllDocuments(UpdateStatus){
  //this.context.spHttpClient.get(`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${this.AttendanceSnapshotLibrary}')/items?$select=*,Request/ID,Folder/ServerRelativeUrl,FileLeafRef,ID,FileSystemObjectType,BaseName,Modified&$expand=Request&$filter=Request/ID%20eq%20${ItemID}`, SPHttpClient.configurations.v1)
  this.context.spHttpClient.get(`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${this.AttendanceSnapshotLibrary}')/items?$select=*,Request/ID,Folder/ServerRelativeUrl,FileLeafRef,ID,FileSystemObjectType,BaseName,Modified&$expand=Request&$filter=(Request/ID%20eq%20${ItemID}) and (IsActive eq 1)`, SPHttpClient.configurations.v1)
      .then(response => {
          return response.json()
              .then((items: any): void => {
                  var htmlSnippet = "";
                  if (items["value"].length > 0) {
                      items["value"].forEach(element => {
                          var docUrl = this.context.pageContext.site.absoluteUrl + "/" + this.AttendanceSnapshotLibrary + "/" + ItemID + "/" + element.FileLeafRef;
                          htmlSnippet += '<div><a class="footage" href="' + docUrl + '" target="_blank">' + element.FileLeafRef + '<a></div>';
                      });
                  }
                  else if(UpdateStatus)
                  {
                    $('#dvSnapshotDetails').hide();
                  }
                  $('#dvUpload').html(htmlSnippet);                     

              });
      });
}


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
      this.Is_ITManager_Group=$.inArray(this.IT_Manager_GroupName, groups);
      this.Is_SSS_Group=$.inArray(this.SSS_GroupName, groups);
      this.Is_HRManager_Group=$.inArray(this.HRManager_GroupName, groups);
      this.Is_HROfficer_Group=$.inArray(this.HROfficer_GroupName, groups);
    }
    this.GetMasterListItem();
    
  } 

  private DisableDocumentsAndUpload()
  {
    let docLibrary=sp.site.rootWeb.lists.getByTitle(this.AttendanceSnapshotLibrary);

    sp.site.rootWeb.lists.getByTitle(this.AttendanceSnapshotLibrary).items.select("IsActive,Request/ID,ID").expand("Request").filter("Request/ID eq '"+ItemID+"' and IsActive eq 1").getAll().then(r=>{
      if(r.length>0)
      {
        r.forEach(element => {
          docLibrary.items.getById(element.ID).update(
            {
              IsActive:false,
            }
          ).then(r=>{
            //alert("Inside then");

          }).catch(function(err) {  
            console.log(err);  
         
          }).then(r=>{
            this.UploadFiles();

          });

        });
        //alert("updated");
        
        

      }
      else
      {
        this.UploadFiles();
      }

      //window.location.href=this.context.pageContext.web.absoluteUrl+"/Pages/TecPages/common/RequestComplete.aspx";
    }).catch(function(err) {  
      console.log(err);  
    });
    //docLibrary.items.query.
    // console.log(`${this.context.pageContext.site.absoluteUrl}'/_api/web/lists/getbytitle('${this.AttendanceSnapshotLibrary}')/items?$select=*,Request/ID,Folder/ServerRelativeUrl,FileLeafRef,ID,FileSystemObjectType,BaseName,Modified&$expand=Request&$filter=(Request/ID%20eq%20${ItemID})%20and%20(IsActive eq 1)`);
    // this.context.spHttpClient.get(`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${this.AttendanceSnapshotLibrary}')/items?$select=*,Request/ID,Folder/ServerRelativeUrl,FileLeafRef,ID,FileSystemObjectType,BaseName,Modified&$expand=Request&$filter=(Request/ID%20eq%20${ItemID})%20and%20(IsActive eq 1)`, SPHttpClient.configurations.v1)
    //   .then(response => {
    //       return response.json()
    //           .then((items: any): void => {
                  
    //               if (items["value"].length > 0) {
    //                   items["value"].forEach(element => {
    //                     docLibrary.items.getById(element.Id).update({
    //                       IsActive:false,
                        
    //                     });
                        
    //                   });
    //                   //this.UploadFiles();
    //                   alert("updated");
    //               }
                  
                                 

    //           });
    //   });


  }

  private blob() {
    //Get the File Upload control id
    let input = <HTMLInputElement>document.getElementById("flUpload");
    var fileNameHtml="";
    var isValidaformat=true;
    var isValidSize=true;
    var fileCount = input.files.length;
    if(fileCount>0 && fileCount<=3){
      for (var i = 0; i < fileCount; i++) {
       var fileName = input.files[i].name;
       var ext=fileName.replace(/^.*\./, '');
       if ($.inArray(ext, ['gif','png','jpg','jpeg']) == -1){
        //$("#spFootage").text("snapshot must be jpg,jpeg,png,gif format.");
        
        //$("#file_input").val(fileName);
        fileNameHtml+=fileName+";";
        fileInfos.length=0;
        isValidaformat=false;
       }
        else 
        {
          
          var filesize=input.files[0].size;
          const kb = Math.round((filesize / 1024));
          if(kb<=10240){//  checking file must be 10 mb less than
          
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
            //$("#spFootage").append("Max allowed footage size is 10.0 MB"); 
            
            fileInfos.length=0;
            isValidSize=false;
          }
        }
      
      }
    
    
    
    
    
        if(!isValidSize && !isValidaformat)
        {
          //show error msg for size and format
         
          $("#spFootage").append("Snapshot must be jpg,jpeg,png,gif format and Max allowed size is 10.0 MB"); 
          fileInfos.length=0;
        }
        else if(!isValidaformat)
        {
          //show error msg for format
          
          $("#spFootage").text("Snapshot must be jpg,jpeg,png,gif format.");
          fileInfos.length=0;

        }
        else if(!isValidSize)
        {
          //show error msg for size
          $("#spFootage").text("Max allowed screenshot size is 10.0 MB"); 
          fileInfos.length=0;

        }
        else
        {
          $("#spFootage").text("");
          $("#file_input").val(fileNameHtml);
        }
    //End of for loop
    }
    else if(fileCount>3){
      $("#spFootage").text("Max allowed files are 3."); 
    }
    else{
      $("#file_input").val("No file Selected");
    }
  }
  

}
