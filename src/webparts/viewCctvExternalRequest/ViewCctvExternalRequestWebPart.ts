import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import styles from './ViewCctvExternalRequestWebPart.module.scss';
import * as strings from 'ViewCctvExternalRequestWebPartStrings';
import { sp } from '@pnp/sp/presets/all';
import * as moment from 'moment';

import { CCTVExternalList } from './../Interfaces/ICCTVExternal';
export interface IViewCctvExternalRequestWebPartProps {
  description: string;
}
  let groups: any[] = [];
  let extensions_arr:any[]=['gif','png','jpg','jpeg'];
  var Listname = "CCTV_External_Incident";
  var ITMgrGroupName="ITManager";
  var LegalMgrGroupName="LegalManager";
  var SSSGroupName="System Security Specialist";
  var CEOGroupName="CEO";
  var cctvfootageLibrary="CCTVExternalFootage";
  const url : any = new URL(window.location.href);
  const ItemID= url.searchParams.get("ItemID");
  var StatusID:any;
  var ItemContentTypeName:string;
  var IsITTeamMember:number;
  var IsLegalTeamMember:number;
  var IsSSSTeamMember:number;
  var IsCEOTeamMember:number;
  var fileInfos = [];
  var No_of_Attachments:number;
  var fileCount:number;
  var type_footage_val:string="Screenshot";
export default class ViewCctvExternalRequestWebPart extends BaseClientSideWebPart<IViewCctvExternalRequestWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <section class="inner-page-cont" style="margin-top: -60px">

      <div class="container-fluid mt-5">
        
            <div class="row user-info col-md-10 mx-auto col-12 pl-0 pr-0 m-0">
               <h3 class="mb-4 col-12">Requester Details</h3>
                <div class="col-md-4 col-12 mb-4">
                    <p>Requester Name</p>
                    <input type="text" id="lbl_emp_name" class="form-input" name="txtEmpName" disabled="disabled">
                    
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Email</p>
                    <input type="text" id="lbl_emp_email" class="form-input" name="txtEmpMail" disabled="disabled">
                    
                </div>
                <div class="col-md-4 col-12 mb-4">
                  <p>Mobile/Telephone Number</p>
                  <input type="text" id="lbl_emp_mobile" class="form-input" name="txtEmpMobile" disabled="disabled">
                  
                </div>
                <h3 class="mb-4 col-12">Incident Details</h3>
                <div class="col-md-4 col-12 mb-4">
                    <p>Incident Date & Time (From)</p>
                    <input type="text" id="lbl_date_time_req" class="form-input" name="txtEmpMobile" disabled="disabled">

                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Incident Date & Time (To)</p>
                    <input type="text" id="lbl_date_time_req_to" class="form-input" disabled="disabled">

                </div>
                <div class="col-md-4 col-12 mb-4">
                  <p>Location/Facility of incident</p>
                  <input type="text" id="lbl_location_facility" class="form-input" name="lbl_location_facility" disabled="disabled">
                  
                </div>
                <div class="col-12 mb-4" id="div_attachment" style="display:none">
                    <p>Attachments</p>

                    <table class="attachment_sec" id="attchment_tbl"></table>
                </div>
                
                <div class="col-md-12 col-12 mb-4">
                    <p>Reason for Request</p>
                    <textarea style="height:auto !important" rows="5" cols="5" id="txtReqReason" class="form-input" name="txtReqReason" disabled="disabled"></textarea>
                </div>  
                <h3 class="mb-4 col-12">Request Details</h3>
                <div class="col-md-4 col-12 mb-4">
                    <p>Request Number</p>
                    <input type="text" id="lbl_req_num" class="form-input" name="txtReqNum" disabled="disabled">
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Requested By</p>
                    <input type="text" id="lbl_req_by" class="form-input" name="txtEmpName" disabled="disabled">
                    
                </div>
                 <div class="col-md-4 col-12 mb-4">
                    <p>Date of Request</p>
                    <input type="text" id="lbl_date_of_req" class="form-input" name="txtEmpMobile" disabled="disabled">
                    
                </div>
                <div class="col-md-6 col-12 mb-4">
                    <p>Status</p>
                    <select name="lbl_status" id="lbl_status" class="form-control" disabled="disabled"></select>
                    
                </div>
            </div>
        
      </div>
      
      <div class="container-fluid mt-5" id="div_row_sss" style="display:none">
          <div class="col-md-10 mx-auto col-12">
              <div class="row user-info">
                  <h3 class="mb-4 col-12">APPROVER ACTION</h3>
                  <div class="col-md-12 col-12 mb-4">
                  <p>Type of Footage</p>
                    <div class="vleft">
                      <input type="radio" id="rb_screenshot" name="Type_Footage" class="form-control" value="Screenshot" checked>
                      <label id="lbl_screenshot" class="form-label">Screenshot</label>
                      <input type="radio" id="rb_footage" name="Type_Footage" class="form-control" value="Footage">
                      <label id="lbl_footage" class="form-label">Footage</label>
                    </div>
                  <label id="lbl_type-footage" class="form-label mb-4" style="color: red;"></label></br>
                </div>
                  <div class="col-md-12 mb-4" id="div_sss_fileupload">
                      <p id="p_pleaseupload">Please upload Screenshot<span style="color:red">*</span></p>
                      <div class="input-group col-md-4 ">
                        <input type="text" name="filename" class="form-control" id="file_input" readonly="" placeholder="No file selected">
                        <span class="input-group-btn">
                            <div class="btn file-btn custom-file-uploader">
                            <input type="file" className="form-control" id="sssfile" multiple/>
                                Select a file
                            </div>
                        </span>
                      </div>
                      <label id="lbl_sss_File_err" class="form-label mb-4" style="color: red;"></label></br>
                      <b>Note :</b><i id="italic_attach_note">The allowed file types are jpg,jpeg,png,gif and the max allowed file size is 10.0 MB</i>                                
                  </div>
              </div>
          </div>
      </div>

      <div class="container-fluid mt-5"  style="display:none" id="div_row_buttons">
        <div class="col-md-10 mx-auto col-12">
            <div class="row">
              <div class=" col-12 btnright">
                    <button class="red-btn shadow-sm  mt-4" id="btn_Submit"><span>Submit</span></button>
                    <button class="red-btn shadow-sm  mt-4" id="btn_Cancel"><span>Cancel</span></button>
                </div>
            </div>
         </div>
      </div>     
    </section>   
    `;
    this.PageLoad();
  }
  private PageLoad()
  {

    this._checkUserInGroup();
    this.setButtonsEventHandlers();
   
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
      IsITTeamMember=$.inArray( ITMgrGroupName, groups );
      IsLegalTeamMember=$.inArray( LegalMgrGroupName, groups );
      IsSSSTeamMember=$.inArray( SSSGroupName, groups );
      IsCEOTeamMember=$.inArray(CEOGroupName,groups);
      this.getListData();
    } 
  } 
  private setButtonsEventHandlers(): void 
  {
    const webPart: ViewCctvExternalRequestWebPart = this;
    
    this.domElement.querySelector('#btn_Submit').addEventListener('click', (e) => { 
      e.preventDefault();
      webPart.UpdateMasterList();
     }); 
     
     this.domElement.querySelector('#btn_Cancel').addEventListener('click', (e) => { 
      e.preventDefault();
      window.location.href=this.context.pageContext.web.absoluteUrl;
     });
     this.domElement.querySelector('#sssfile').addEventListener('change', (e) => { 
      e.preventDefault();
      webPart.blob();
     });
     this.domElement.querySelector('#rb_screenshot').addEventListener('change', (e) => { 
      e.preventDefault();
     this.handleradioClick(e);
     });
     this.domElement.querySelector('#rb_footage').addEventListener('change', (e) => { 
      e.preventDefault();
     this.handleradioClick(e);
     });
  }
  private UpdateMasterList(){
      if(StatusID==2 || StatusID==8){
        // retrieve request intiated and saves SSS data
         if(fileInfos.length>0 && fileInfos.length<=5){
         this.CheckAndCreateFolder();
         }else{
          $("#lbl_sss_File_err").text("Footage/Screenshot is mandatory , must be jpg,jpeg,png,gif format and max files are 5 and allowed size is 20.0 MB");
        }
    }
  }
  private UpdateSSSReview(){
    /* var fileSeqNo=No_of_Attachments+1;
    sp.site.rootWeb.lists.getByTitle(Listname).items.getById(ItemID).attachmentFiles.add(fileSeqNo+"_"+fileInfos[0].name,fileInfos[0].content)
        .then(r=>{
            alert("Thank you ! The request was updated successfully");
            window.location.href=this.context.pageContext.web.absoluteUrl+"/Pages/TecPages/common/RequestComplete.aspx";
          }).catch(function(err) {  
            console.log(err);  
        }); */
    this.CheckAndCreateFolder();
  }
  private blob() {
      //Get the File Upload control id
      $("#lbl_sss_File_err").text("");
      var filebuffsize;
      
      if(type_footage_val=="Footage"){
        extensions_arr=['mp4','mov','wmv','flv','avi'];
        filebuffsize=200;
      }else{
        extensions_arr=['gif','png','jpg','jpeg'];
        filebuffsize=10;
      }
      let input = <HTMLInputElement>document.getElementById("sssfile");
      fileCount = input.files.length;
      if(fileCount>0 && fileCount<=5){
        for (var i = 0; i < fileCount; i++) {
              var fileName = input.files[i].name;
              $("#file_input").val(fileName);
              var ext=fileName.replace(/^.*\./, '');
              if ($.inArray(ext,extensions_arr) == -1){
                $("#lbl_sss_File_err").text(type_footage_val+" must be "+extensions_arr+" format");
                $("#file_input").val(fileName);
              fileInfos.length=0;
            }else{
              $("#lbl_sss_File_err").text("");
              var filesize=input.files[0].size;
              const kb = Math.round((filesize / 1024));
              if(kb<=(filebuffsize*1024)){//  checking file must be 20 mb less than
                $("#lbl_sss_File_err").text("");
                $("#file_input").val(fileName);
                //console.log(fileName);
                var file = input.files[i];
                var reader = new FileReader();
                reader.onload = (function(file) {
                    return function(e) {
                      console.log(file.name);
                      //Push the converted file into array
                          fileInfos.push({
                            "name": file.name,
                        "content": e.target.result
                        });
                      console.log(fileInfos);
                      }
                    })(file);
                  reader.readAsArrayBuffer(file);
                }
                else{
                  $("#lbl_sss_File_err").text("Max allowed "+type_footage_val+" size is "+filebuffsize+"MB"); 
                  fileInfos.length=0;
                }
              }
          }
      //End of for loop
      }
      else if (fileCount>5){
        $("#lbl_sss_File_err").text("Max allowed files are 5"); 
        
      }
      else{
        $("#file_input").val("No file Selected");
      }

  }

  private getListData() {
    let anchorhtml: string ='';
   
    var Url =this.context.pageContext.site.absoluteUrl+`/_api/web/lists/getbytitle('${Listname}')/items?$select=*,Attachments,AttachmentFiles,Author/Name,Author/Title,Status/ID,Status/Title,ContentType/Id,ContentType/Name&$expand=AttachmentFiles,Status,ContentType,Author/Id&$filter=ID eq `+ItemID;
    this.context.spHttpClient
      .get(Url, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json().then((items: any): void => {
         
            let listItems: CCTVExternalList[] = items.value;
            listItems.forEach((item: CCTVExternalList) => {
            
            console.log(item);
            
            //loading all common fields   
             var cctv_inc_status:string;  

             if(item.Status!=null){
             StatusID=item.Status.ID;
             cctv_inc_status=item.Status.Title;
             }

             if(IsSSSTeamMember<0  && IsLegalTeamMember<0 && IsITTeamMember<0 && IsCEOTeamMember<0){
                // if unauthorized user
                alert("You don't have access,please contact administrator for more info.");
                window.location.href=this.context.pageContext.web.absoluteUrl;
              }               

             if(IsSSSTeamMember>=0 && StatusID==2){
               // show sections only to sss retrive request initiated
               $("#div_attachment").show();
               $("#div_row_sss").show();
               $("#div_row_buttons").show();
               this.LoadAllDocuments(true);
             }
             else if((IsITTeamMember>=0 && StatusID==10)||(IsITTeamMember>=0 && StatusID==8)||(IsSSSTeamMember>=0 && StatusID==5)||(IsITTeamMember>=0 && StatusID==5)|| (IsITTeamMember>=0 && StatusID==9) || (IsITTeamMember>=0 && StatusID==7)){
                // show sections to sss footage available and Itmgr footage available
                $("#div_attachment").show();
                this.LoadAllDocuments(false);
             }
            
             // show attachment and approver action if more info required
             if((IsSSSTeamMember>=0 && StatusID==8 )){
                $("#div_attachment").show();
                $("#div_row_sss").show();
                $("#div_row_buttons").show();
                this.LoadAllDocuments(true);
             }
             // show attachment to CEO if cctv information verified
             if(StatusID==7 && IsCEOTeamMember>=0 || StatusID==8 && IsCEOTeamMember>=0 ){
              $("#div_attachment").show();
              this.LoadAllDocuments(false);
             }
             
             //show footage to all once approved by CEO and acknowledged by requestor
             if(StatusID==9 || StatusID==11){
              // cctv footage verified & received
              $("#div_attachment").show();
              this.LoadAllDocuments(false);
            }
             
             var momentObj = moment(item.Created);           
             var formatCreatedDate=momentObj.format('DD-MM-YYYY HH:mm');

             var cctv_req_number=item.Title;
             var cctv_createdBy=item.Author.Title;
             var cctv_emp_name=item.RequesterName!=null?item.RequesterName:"";
             
             var cctv_emp_email=item.EmailAddress!=null?item.EmailAddress:"";
             var cctv_emp_mobile=item.Mobile_x002d_Tel_x0020_No !=null?item.Mobile_x002d_Tel_x0020_No:"";
            
             var cctv_date_of_inc=item.DateOfIncident!=null?item.DateOfIncident:"";
             var IncDatemomentObj = moment(cctv_date_of_inc);           
             var dt_MommentIncDate=IncDatemomentObj.format('DD-MM-YYYY');
             var cctv_time_of_inc=item.TimeOfIncident!=null?item.TimeOfIncident:"";

             var cctv_date_time=dt_MommentIncDate +"  "+cctv_time_of_inc;

             var cctv_date_of_inc_to=item.DateOfIncident_To!=null?item.DateOfIncident_To:"";
             var IncDatemomentObj_to=moment(cctv_date_of_inc_to);
             var dt_MommentIncDate_to=IncDatemomentObj_to.format('DD-MM-YYYY');
             var cctv_time_of_inc_to=item.TimeOfIncident_To!=null?item.TimeOfIncident_To:"";

             var cctv_date_time_to=dt_MommentIncDate_to +"  "+cctv_time_of_inc_to;
             var cctv_inc_reason_for_req=item.ReasonForRequest!=null?item.ReasonForRequest:"";
             var cctv_location_facility=item.LocationFacilityOfIncident!=null? item.LocationFacilityOfIncident:"";

             /* if(items.value[0].AttachmentFiles.length>0){
              //for(var i=0;i<items.value[0].AttachmentFiles.length;i++){
                //get latest file id
                No_of_Attachments=items.value[0].AttachmentFiles.length;
                var actualFileName=items.value[0].AttachmentFiles[No_of_Attachments-1].FileName
                var getlastfilename=actualFileName.split("_")[1];
                var anchorfileURL=this.context.pageContext.site.absoluteUrl+"/Lists/"+Listname+"/Attachments/"+ItemID+"/"+items.value[0].AttachmentFiles[No_of_Attachments-1].FileNameAsPath.DecodedUrl+"?web=1";
                //console.log(anchorfileURL);
                //anchorhtml+='<a href="'+anchorfileURL+'">'+items.value[0].AttachmentFiles[i].FileName+'</a><br>';  
                //get latest one file only
                anchorhtml+='<a href="'+anchorfileURL+'">'+getlastfilename+'</a><br>';  
               // }
             }
             else{
              No_of_Attachments=0;
              anchorhtml+='<a href="#">No attachments</a>';
             } */
             
             $("#lbl_req_by").val(cctv_createdBy);
             $("#lbl_req_num").val(cctv_req_number);
             $("#lbl_emp_name").val(cctv_emp_name);
             $("#lbl_date_of_req").val(formatCreatedDate);
             $("#lbl_location_facility").val(cctv_location_facility);
             $("#lbl_emp_email").val(cctv_emp_email);
             $("#lbl_emp_mobile").val(cctv_emp_mobile);
             $("#txtReqReason").val(cctv_inc_reason_for_req);
             $("#lbl_date_time_req").val(cctv_date_time);
             $("#lbl_date_time_req_to").val(cctv_date_time_to);
             $("#lbl_status").append($('<option></option>').val(cctv_inc_status).html(cctv_inc_status));
             
            
          });
          
         
        }).catch(function(err) {  
          console.log(err);  
        });
      });
  }
  private LoadAllDocuments(isdeleterequired:boolean){
    //this.context.spHttpClient.get(`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${this.AttendanceSnapshotLibrary}')/items?$select=*,Request/ID,Folder/ServerRelativeUrl,FileLeafRef,ID,FileSystemObjectType,BaseName,Modified&$expand=Request&$filter=Request/ID%20eq%20${ItemID}`, SPHttpClient.configurations.v1)
    this.context.spHttpClient.get(`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${cctvfootageLibrary}')/items?$select=*,Request/ID,Folder/ServerRelativeUrl,FileLeafRef,ID,FileSystemObjectType,BaseName,Modified&$expand=Request&$filter=(Request/ID%20eq%20${ItemID}) and (IsActive eq 1)`, SPHttpClient.configurations.v1)
        .then(response => {
            return response.json()
                .then((items: any): void => {
                    var htmlSnippet = "";
                    if (items["value"].length > 0) {
                      if(isdeleterequired==true){
                            items["value"].forEach(element => {
                              var docUrl = this.context.pageContext.site.absoluteUrl + "/" +cctvfootageLibrary+ "/" + ItemID + "/" + element.FileLeafRef;
                              htmlSnippet += '<tr><td><a class="footage" href="' + docUrl + '" target="_blank">' + element.FileLeafRef + '</a></td><td><a href="#" id="anc_'+ element.ID+'" class="deletebtn" onClick="{(e) => {e.preventDefault();alert('+element.ID+');}}"><img src="/sites/IntranetDev/Style%20Library/TEC/images/delete.svg" alt="Delete" /></a></td></tr>';
                          });
                          $("#attchment_tbl").html(htmlSnippet);
                          items["value"].forEach(element => {
                            debugger;
                            this.domElement.querySelector("#anc_"+element.ID).addEventListener('click', (e) => { 
                                e.preventDefault();
                                this.DeleteAttachment(element.ID);
                               }); 
                           });
                      }else
                      { 
                          items["value"].forEach(element => {
                            var docUrl = this.context.pageContext.site.absoluteUrl + "/" +cctvfootageLibrary+ "/" + ItemID + "/" + element.FileLeafRef;
                            htmlSnippet += '<tr><td><a class="footage" href="' + docUrl + '" target="_blank">' + element.FileLeafRef + '</a></td></tr>';
                        });
                        $("#attchment_tbl").html(htmlSnippet);
                      }
                    }
                    else{
                      htmlSnippet+='<tr><td><a href="#">No attachments available</a></td></tr>';
                      $("#attchment_tbl").html(htmlSnippet);
                    }
                    //$('#dvUpload').html(htmlSnippet);                     
                  
                   
                });
        });
  }
  private DeleteAttachment(fileid:number)
  {
    sp.site.rootWeb.lists.getByTitle(cctvfootageLibrary).items.getById(fileid).update({
      IsActive:false,
    }).then(r=>{
        console.log( "File Inactive successfully!");
        this.LoadAllDocuments(true);
    }); 
  }
  private  CheckAndCreateFolder()
  {
      //var folderExists=false;
    var folderUrl=this.context.pageContext.site.serverRelativeUrl+"/"+cctvfootageLibrary+"/"+ ItemID;
    sp.site.rootWeb.getFolderByServerRelativeUrl(folderUrl).select('Exists').get().then(data => {
      
      console.log(data.Exists);
      if(data.Exists)
      {
        //folderExists=true;
        //return folderExists;
      
      }
      else{
        sp.site.rootWeb.lists.getByTitle(cctvfootageLibrary).rootFolder.folders.add(ItemID)
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

  private UploadFiles()
  {
    var updateCount:number=0;
   
    var folderUrl=this.context.pageContext.site.serverRelativeUrl+"/"+cctvfootageLibrary+"/"+ItemID;
   
    {
        for(var i=0;i<fileCount;i++)
        {
            var file=fileInfos[i];
            sp.site.rootWeb.getFolderByServerRelativeUrl(folderUrl).files.add(file.name, file.content, true).then((result) => {
                console.log(file.name + " upload successfully!");
                result.file.listItemAllFields.get().then((listItemAllFields) => {
                    // get the item id of the file and then update the columns(properties)
                    sp.site.rootWeb.lists.getByTitle(cctvfootageLibrary).items.getById(listItemAllFields.Id).update({
                      
                      RequestId:ItemID,
                      IsActive:true,
                    }).then(r=>{
                        console.log(file.name + " properties updated successfully!");
                        updateCount++;
                        if(updateCount== fileCount)
                        {
                          
                          alert("Thank you ! The request was updated successfully");
                          window.location.href=this.context.pageContext.web.absoluteUrl+"/Pages/TecPages/common/RequestComplete.aspx";
                        }
                    });           
                }); 
            }).catch(err => {
                console.log("Error While uploading file...");
            });
        }
        
    }
  }
  private handleradioClick(myRadio)
  {     
        $("#lbl_sss_File_err").text("");
        type_footage_val = myRadio.target.value;
        if(type_footage_val=="Footage")
        {
        $("#p_pleaseupload").html("Please upload Footage<span  style='color:red'>*</span>");
        $("#italic_attach_note").text("The allowed file types are mp4,mov,wmv,flv,avi and the max allowed file size is 200.0 MB");
        }
        else {
        $("#p_pleaseupload").html("Please upload Screenshot<span  style='color:red'>*</span>");
        $("#italic_attach_note").text("The allowed file types are jpg,jpeg,png,gif and the max allowed file size is 10.0 MB");
        }
  }
    /*protected get dataVersion(): Version {
        return Version.parse('1.0');
      } */

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
