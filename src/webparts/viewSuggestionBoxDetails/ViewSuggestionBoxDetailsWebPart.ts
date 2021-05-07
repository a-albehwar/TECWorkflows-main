import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ViewSuggestionBoxDetailsWebPart.module.scss';
import * as strings from 'ViewSuggestionBoxDetailsWebPartStrings';
import { Items, sp } from '@pnp/sp/presets/all';
import * as moment from 'moment';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { SuggestionBoxListCols } from './../Interfaces/ISuggestionBox';

export interface IViewSuggestionBoxDetailsWebPartProps {
  description: string;
}
let groups: any[] = [];
var Listname = "Suggestions Box";

var InnovationGroupName="InnovationTeam";
var assignedDeptName="-SuggestionsBoxDepartment";
var assignedDeptHeadName="-SuggestionsBoxDepartmentHead";
var DeptListname="LK_Departments";
const url : any = new URL(window.location.href);
const ItemID= url.searchParams.get("ItemID");
var StatusID:any;
var ItemContentTypeName:string;
var IsInnovationTeamMember:number;
var fileInfos = [];
var fileCount:number;

export default class ViewSuggestionBoxDetailsWebPart extends BaseClientSideWebPart<IViewSuggestionBoxDetailsWebPartProps> {
  private DocLibraryName:string ="SuggestionBoxDocuments";
  public render(): void {
    this.domElement.innerHTML = `

    <section class="inner-page-cont" style="margin-top: -60px">

    <div class="container-fluid mt-5">
        
            <div class="row user-info col-md-10 mx-auto col-12 pl-0 pr-0 m-0">
                    <h3 class="mb-4 col-12">Employee Details</h3>
                      <div class="col-md-4 col-12 mb-4">
                        <p>Name</p>
                        <input type="text" id="txtEmpName" class="form-input" name="txtEmpName"  disabled="disabled">

                      </div>
                      <div class="col-md-4 col-12 mb-4">
                        <p>Job Title</p>
                        <input type="text" id="txt_jobTitle" class="form-input" name="txt_jobTitle"  disabled="disabled">
                        
                      </div>
                      <div class="col-md-4 col-12 mb-4">
                        <p>Department</p>
                        <input type="text" id="txt_dept" class="form-input" name="txt_dept"  disabled="disabled">
                      </div>
                <h3 class="mb-4 col-12">Suggestion Box Details</h3>
                <div class="col-md-12 col-12 mb-4">
                    <p>Type of Suggestion</p>
                    <div class="vleft">
                      <input type="radio" id="rb_money" name="suggestionType" class="form-control" value="Save Money">
                      <label id="lbl_money" class="form-label">Save Money</label>
                      <input type="radio" id="rb_security" name="suggestionType" class="form-control" value="Improve Safety">
                      <label id="lbl_security" class="form-label">Improve Security</label>
                      <input type="radio" id="rb_efficency" name="suggestionType" class="form-control" value="Improve Efficiency">
                      <label id="lbl_Efficency" class="form-label">Improve Efficiency</label>
                      <input type="radio" id="rb_other" name="suggestionType" class="form-control" value="Other">
                      <label id="lbl_Other" class="form-label">Other</label>
                    </div>
                    <label id="lbl_suggType_val" class="form-label"></label>
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Attachment</p>
                    <label id="lbl_sugg_attach_val" class="form-label"></label>
                </div>
                <div class="col-md-12 col-12 mb-4">
                    <p>Describe the current situation</p>
                    <textarea class="form-input" style="height:auto!important" rows="5" cols="5" disabled="disabled" id="lbl_describe_curr_situ_val"></textarea>
                    
                </div>
                <div class="col-md-12 col-12 mb-4">
                    <p>Describe the suggestion and how it improves the current situation</p>
                    <textarea class="form-input" style="height:auto!important" rows="5" cols="5" disabled="disabled" id="lbl_describe_how_improve_situ_val"></textarea>
                </div>
                <h3 class="mb-4 col-12">Request Details</h3>
                <div class="col-md-4 col-12 mb-4">
                    <p>Requested By</p>
                    <input type="text" class="form-control" name="" id="lbl_sugg_CreatedBy_val" disabled="disabled">
                    
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Created Date</p>
                    <input type="text" class="form-control" name="" id="lbl_sugg_createdDate_val" disabled="disabled">
                  
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Department</p>
                    <select name="lbl_sugg_dept_val" disabled="disabled" class="form-input" id="lbl_sugg_dept_val"></select>
                    
                </div>
                
               
                <div class="col-md-4 col-12 mb-4">
                    <p>Status</p>
                    <select name="lbl_sugg_status_val" disabled="disabled" class="form-input" id="lbl_sugg_status_val"></select>

                </div>
              <!--<div class="col-md-4 col-12 mb-4" style="display:none" id="div_row_innovate_comments">
                    <p>Innovation Team Comments</p>
                    <label id="lbl_innovate_comments_val" class="form-label"></label>
                </div>
                <div class="col-md-4 col-12 mb-4" style="display:none" id="div_row_deptTeam_comments">
                    <p>Deparment Team Comments</p>
                    <label id="lbl_dept_team_comments_val" class="form-label"></label>
                </div>
                <div class="col-md-4 col-12 mb-4" style="display:none"   id="div_row_deptHead_comments">
                    <p>Department Head Comments</p>
                    <label id="lbl_dept_head_val" class="form-label"></label>
                </div>
                -->
            </div>
        
    </div>
     
      <div class="container-fluid mt-5" style="display:none" id="div_row_InnovationTeam">
          <div class="col-md-10 mx-auto col-12">
              <div class="row user-info">
                  <h3 class="mb-4 col-12">APPROVER ACTION</h3>
                  <div class="col-md-4 col-12 mb-4">
                      <p>Department<span style="color:red">*</span></p>
                      <select name="ddl_Dept" id="ddl_Dept" class="form-input">
                      </select>
                      <label id="lbl_Innovate_first_comm_err" class="form-label" style="color: red;"></label>
                  </div>
                 <!-- <div class="col-md-8 col-12 mb-4"  style="display:none" >
                      <p>Comments<span style="color:red">*</span></p>
                      <textarea id="txt_Innovatation_Comments" style="height:auto !important" rows="5" cols="5" class="form-control"></textarea>
                  </div>
                 -->
              </div>
          </div>
      </div>

      <div class="container-fluid mt-5" style="display:none" id="div_row_DeptTeam">
          <div class="col-md-10 mx-auto col-12">
              <div class="row user-info">
                  <h3 class="mb-4 col-12">APPROVER ACTION</h3>
                  <div class="col-md-4 col-12 mb-4">
                      <p>Upload Supporting Document<span style="color:red">*</span></p>
                      <div class="input-group">
                        <input type="text" name="filename" class="form-control" placeholder="No file selected" id="file_input" readonly="">
                        <span class="input-group-btn">
                            <div class="btn file-btn custom-file-uploader">
                            <input type="file" className="form-control" id="deptfile"/>
                                Select a file
                            </div>
                        </span>
                      </div> 
                     
                  </div>
                 <div class="col-md-12">
                    <label id="lbl_DeptTeam_File_err" class="form-label" style="color: red;"></label>                     </br>
                    <b>Note :</b><i>The allowed file types are doc,docx,pdf,ppt and the max allowed file size is 20.0 MB</i>
                  </div>
              </div>
              
          </div>
      </div>
        <!--
          <div class="container-fluid mt-5" style="display:none" id="div_row_DeptHead">
              <div class="col-md-10 mx-auto col-12">
                  <div class="row user-info">
                      <h3 class="mb-4 col-12">APPROVER ACTION</h3>
                      <div class="col-md-12 col-12 mb-4">
                          <p>Comments<span style="color:red">*</span></p>
                          <textarea id="txt_DeptHead_Comments" style="height:auto !important" rows="5" cols="5" class="form-control"></textarea>
                          <label id="lbl_deptHead_comm_err" class="form-label" style="color: red;"></label>
                      </div>
                    
                  </div>
              </div>
          </div>
        -->
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
      IsInnovationTeamMember=$.inArray( InnovationGroupName, groups );
      this.getListData();
    } 
  } 
  private getListData() {
    let anchorhtml: string ='';
   
    var URL = "";
    URL =`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${Listname}')/items?$select=*,Attachments,AttachmentFiles,AssignedDepartment/ID,AssignedDepartment/Title,Author/Name,Author/Title,Suggestion_Status/ID,Suggestion_Status/Title,ContentType/Id,ContentType/Name&$expand=AttachmentFiles,Suggestion_Status,ContentType,AssignedDepartment,Author/Id&$filter=ID eq `+ItemID;
    this.context.spHttpClient
      .get(URL, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json().then((items: any): void => {
          var loginuserid=this.context.pageContext.legacyPageContext["userId"];
            let listItems: SuggestionBoxListCols[] = items.value;
            listItems.forEach((item: SuggestionBoxListCols) => {
            // check is login user innovation team member or assigned dept team member ,assigned dept head
            console.log(item);

            //loading all common fields            
            if(item.Suggestion_Status!=null){
            StatusID=item.Suggestion_Status.ID;
            }
            
            var momentObj = moment(item.Created);           
            var formatCreatedDate=momentObj.format('DD-MM-YYYY HH:mm');
            ItemContentTypeName=item.ContentType.Name;
            var sugg_title=item.Title;
            var sugg_createdBy=item.Author.Title;
            //var sugg_authorID=item.AuthorId;
            var sugg_desc=item.Description;
            var sugg_type=item.Suggestion_Type;
            if(item.Suggestion_Status!=null){
            var sugg_status_title=item.Suggestion_Status.Title;
            var empdept=item.User_Department!=null?item.User_Department:"";
            var empname=item.User_Name!=null?item.User_Name:"";
            var empjobtitle=item.User_JobTitle!=null?item.User_JobTitle:"";
            }
            var deptid:any;
            if(StatusID>1 && item.AssignedDepartment!=null){// Assigned department doesnot exists in suggestion intiated 
            var sugg_assigned_dept_name=item.AssignedDepartment.Title!=null?item.AssignedDepartment.Title:"";
            deptid=item.AssignedDepartment.ID!=null?item.AssignedDepartment.ID:"";
            }
            if(item.AssignedDepartment!=null){
              deptid=item.AssignedDepartment.ID!=null?item.AssignedDepartment.ID:"";
              var sugg_assigned_dept_name=item.AssignedDepartment.Title!=null?item.AssignedDepartment.Title:"";
            }
            var sugg_innovation_comments=item.Innovation_Team_Review!=null?item.Innovation_Team_Review:"";
            var sugg_assigned_dept_team_comments=item.Assigned_Dept_Comments!=null?item.Assigned_Dept_Comments:"";
            var sugg_assigned_dept_head_comments=item.Dept_Head_Comments!=null?item.Dept_Head_Comments:"";
            // load comments section if those are not null
            if(sugg_innovation_comments!=null && sugg_innovation_comments!="" )
            { 
              $("#div_row_innovate_comments").show();
              $("#lbl_innovate_comments_val").html(sugg_innovation_comments);
            }
            if(sugg_assigned_dept_team_comments!=null && sugg_assigned_dept_team_comments!="")
            { 
              $("#div_row_deptTeam_comments").show();
              $("#lbl_dept_team_comments_val").html(sugg_assigned_dept_team_comments);
            }
            if(sugg_assigned_dept_head_comments!=null && sugg_assigned_dept_head_comments!="")
            { 
              $("#div_row_deptHead_comments").show();
              $("#lbl_dept_head_val").html(sugg_assigned_dept_head_comments);
            }
            var IsDeptteamMember=$.inArray( deptid+assignedDeptName, groups );
            var IsDeptHeadMember=$.inArray( deptid+assignedDeptHeadName, groups );
            if(items.value[0].AttachmentFiles.length>0){
              for(var i=0;i<items.value[0].AttachmentFiles.length;i++){
                // list internal name for document not display name
                var anchorfileURL=this.context.pageContext.site.absoluteUrl+"/Lists/SuggestionsBox/Attachments/"+ItemID+"/"+items.value[0].AttachmentFiles[i].FileNameAsPath.DecodedUrl+"?web=1";
                //console.log(anchorfileURL);
                anchorhtml+='<a href="'+anchorfileURL+'">'+items.value[0].AttachmentFiles[i].FileName+'</a><br>';
              
              }
             }
             else{
              anchorhtml+='<a href="#">No attachments</a>';
             }
            //$("#lbl_suggType_val").html(sugg_type);
           // enable radio based on value and disable
            $('input[name="suggestionType"]').prop("disabled", true);
            if(sugg_type=="Save Money" || sugg_type=="Improve Efficiency" || sugg_type =="Improve Safety"){
              $("input[name=suggestionType][value='" + sugg_type + "']").prop('checked', true);
            }
            else{
              $("#lbl_Other").html(sugg_type);
              $("input[name=suggestionType][value='Other']").prop('checked', true);
            }
            $("#lbl_describe_how_improve_situ_val").val(sugg_desc);
            

           // $("#lbl_sugg_dept_val").val(sugg_assigned_dept_name);
            $("#lbl_sugg_dept_val").append($('<option></option>').val(sugg_assigned_dept_name).html(sugg_assigned_dept_name));
            $("#lbl_sugg_createdDate_val").val(formatCreatedDate);

            $("#txtEmpName").val(empname);
            $("#txt_jobTitle").val(empjobtitle);
            $("#txt_dept").val(empdept);

            $("#lbl_describe_curr_situ_val").val(sugg_title);
            $("#lbl_sugg_CreatedBy_val").val(sugg_createdBy);

            $("#lbl_sugg_status_val").val(sugg_status_title);
            $("#lbl_sugg_status_val").append($('<option></option>').val(sugg_status_title).html(sugg_status_title));
            $("#lbl_sugg_attach_val").html(anchorhtml);

            if(IsInnovationTeamMember>=0 && StatusID==1 && deptid == undefined){
              // If Innovation team member
              $("#div_row_InnovationTeam").show();
              $("#div_row_buttons").show();
              
            }
            else if(IsDeptteamMember>=0 && StatusID==2){
              // dept team
              $("#div_row_DeptTeam").show();
              $("#div_row_buttons").show();
            }
            // else if((IsDeptHeadMember>=0 && StatusID==3) || (IsDeptHeadMember>=0 && StatusID==4)){
            //   // dept head - it is required here for comment.
            //   $("#div_row_DeptHead").show();
            //   $("#div_row_buttons").show();
            // }
            else if(loginuserid==item.AuthorId ){
              $("#div_row_DeptTeam").hide();
              $("#div_row_buttons").hide();
            }
            else if(loginuserid!=item.AuthorId && IsDeptHeadMember<0  && IsDeptteamMember<0 && IsInnovationTeamMember<0 ){
              // if unauthorized user
              alert("You don't have access,please contact administrator for more info.");
              window.location.href=this.context.pageContext.web.absoluteUrl;
            }
            
            

          });
          //this.getRelatedDocuments(RequestTitle);
          this.LoadDepartments();
        }).catch(function(err) {  
          console.log(err);  
        });
      });
  }
  private LoadDepartments():void{
   // sp.site.rootWeb.lists.getByTitle(DeptListname).items.select("Title", "ID").orderBy("Title", true).getAll()
   sp.site.rootWeb.lists.getByTitle(DeptListname).items.orderBy('Title', true).get()
    .then(function (data) {
      //console.log(data);
        $('#ddl_Dept').append(`<option value="0">Select Department</option>`);
      for (var k in data) {
        //alert(data[k].Title);
        $("#ddl_Dept").append("<option value=\"" +data[k].ID + "\">" +data[k].Title + "</option>");
      }
     
    });
  }

 
  private setButtonsEventHandlers(): void 
  {
    const webPart: ViewSuggestionBoxDetailsWebPart = this;
    
    this.domElement.querySelector('#btn_Submit').addEventListener('click', (e) => { 
      e.preventDefault();
      webPart.UpdateMasterList();
     }); 
     
     this.domElement.querySelector('#btn_Cancel').addEventListener('click', (e) => { 
      e.preventDefault();
      window.location.href=this.context.pageContext.site.absoluteUrl+"/Pages/TecPages/SearchSB.aspx";
     });
     this.domElement.querySelector('#deptfile').addEventListener('change', (e) => { 
      e.preventDefault();
      webPart.blob();
     });
  }

  private UpdateMasterList(){
    if(StatusID==1){
      let deptid:any=$("#ddl_Dept").val();
     // let innovateRev:any=$("#txt_Innovatation_Comments").val();
      if(deptid!=0 ){
        this.UpdateInnovationReview(deptid);
      }
      else{
        $("#lbl_Innovate_first_comm_err").text("Department is mandatory");
      }
    }
    else if(StatusID==2){
      if(fileInfos.length!=0){
        this.UpdateDeptTeamReview();
       }else{
       $("#lbl_DeptTeam_File_err").text("Supporting document  is mandatory , must be doc,docx,pdf,ppt format and Max allowed document size is 20.0 MB");
       }
    }
    // else if(StatusID==3){
    //   let deptheadcomments:any=$("#txt_DeptHead_Comments").val();
    //   if(deptheadcomments!=""){
    //   this.UpdateDeptHeadReview(deptheadcomments);
    //   }else{
    //     $("#lbl_deptHead_comm_err").text("Comments are mandatory");
    //   }
    // }
  }
  private UpdateInnovationReview(deptid:number){
    sp.site.rootWeb.lists.getByTitle(Listname).items.getById(ItemID).update({
      AssignedDepartmentId: deptid,
      //Innovation_Team_Review: innovcomments,
    }).then(r=>{
      
      alert("Thank you ! The request was updated successfully");
      window.location.href=this.context.pageContext.web.absoluteUrl+"/Pages/TecPages/common/RequestComplete.aspx";
    }).catch(function(err) {  
      console.log(err);  
   });
  }
  private UpdateDeptTeamReview(){
        var prefixFilename="DepatmentTeam_"+fileInfos[0].name;
        sp.site.rootWeb.lists.getByTitle(Listname).items.getById(ItemID).attachmentFiles.add(prefixFilename,fileInfos[0].content)
        .then(r=>{
            alert("Thank you ! The request was updated successfully");
            window.location.href=this.context.pageContext.web.absoluteUrl+"/Pages/TecPages/common/RequestComplete.aspx";
          }).catch(function(err) {  
            console.log(err);  
        });
  }
  private UpdateDeptHeadReview(deptheadComments:string){
    sp.site.rootWeb.lists.getByTitle(Listname).items.getById(ItemID).update({
      Dept_Head_Comments: deptheadComments,
    }).then(r=>{
      alert("Task updated successfully");
      window.location.href=this.context.pageContext.web.absoluteUrl;
    }).catch(function(err) {  
      console.log(err);  
   });
  }

  private blob() {
    //Get the File Upload control id
    let input = <HTMLInputElement>document.getElementById("deptfile");
    fileCount = input.files.length;
    if(fileCount>0){
    for (var i = 0; i < fileCount; i++) {
       var fileName = input.files[i].name;
       var ext=fileName.replace(/^.*\./, '');
       if ($.inArray(ext, ['doc','docx','pdf','ppt']) == -1){
        $("#lbl_DeptTeam_File_err").text("Supporting documents must be doc,docx,pdf,ppt format");
        $("#file_input").val(fileName);
        fileInfos.length=0;
       }else{
        $("#lbl_DeptTeam_File_err").text("");
        var filesize=input.files[0].size;
        const kb = Math.round((filesize / 1024));
        if(kb<=20024){//  checking file must be 20 mb less than
        $("#lbl_DeptTeam_File_err").text("");
        $("#file_input").val(fileName);
        //console.log(fileName);
        var file = input.files[i];
        var reader = new FileReader();
        reader.onload = (function(file) {
            return function(e) {
              //console.log(file.name);
              //Push the converted file into array
                  fileInfos.push({
                    "name": file.name,
                    "content": e.target.result
                    });
                 // console.log(fileInfos);
                  }
            })(file);
        reader.readAsArrayBuffer(file);
        }
        else{
          $("#lbl_DeptTeam_File_err").text("Max allowed document size is 20.0 MB"); 
          fileInfos.length=0;
        }
      }
    }
  }
  /* private blob() {
    //Get the File Upload control id
    let input = <HTMLInputElement>document.getElementById("deptfile");
    fileCount = input.files.length;
    if(fileCount>0){
    for (var i = 0; i < fileCount; i++) {
       var fileName = input.files[i].name;
       $("#file_input").val(fileName);
       console.log(fileName);
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
    //End of for loop
    }
    else{
      $("#file_input").val("No file Selected");
    }*/
  } 
  /*protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  */
    
  
}

