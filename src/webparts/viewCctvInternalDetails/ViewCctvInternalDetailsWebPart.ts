import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ViewCctvInternalDetailsWebPart.module.scss';
import * as strings from 'ViewCctvInternalDetailsWebPartStrings';
import { sp } from '@pnp/sp/presets/all';
import * as moment from 'moment';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { CCTVInternalList } from './../Interfaces/ICCTVInternal';

export interface IViewCctvInternalDetailsWebPartProps {
  description: string;
}
let groups: any[] = [];
var Listname = "CCTV_Internal_Incident";

var ITMgrGroupName="ITManager";
var LegalMgrGroupName="LegalManager";
var SSSGroupName="System Security Specialist";

const url : any = new URL(window.location.href);
const ItemID= url.searchParams.get("ItemID");
var StatusID:any;
var ItemContentTypeName:string;
var IsITTeamMember:number;
var IsLegalTeamMember:number;
var IsSSSTeamMember:number;
var fileInfos = [];
var No_of_Attachments:number;
var fileCount:number;
export default class ViewCctvInternalDetailsWebPart extends BaseClientSideWebPart<IViewCctvInternalDetailsWebPartProps> {
  //siteURL=this.context.pageContext.site.absoluteUrl;
  private valempid:string;
  private valempname:string;
  private valReqReason:string;
  private  valempdept:string;
  private valempmobile:string;
  private valincTimehrs:string;
  private valempmail:string;
  private valincDate:Date;
  private valincDateTo:Date;
  private valincTimeMins:string;
  private valincTimehrsTo:string;
  private valincTimeMinsTo:string;
  public render(): void {
    this.domElement.innerHTML = `
    <section class="inner-page-cont" style="margin-top: -60px">

    <div class="container-fluid mt-5">
        
            <div class="row user-info col-md-10 mx-auto col-12 pl-0 pr-0 m-0">
                <h3 class="mb-4 col-12">EMPLOYEE DETAILS</h3>
                <div class="col-md-4 col-12 mb-4">
                    <p>Full Name</p>
                    <input type="text" id="lbl_emp_name" class="form-input" name="txtEmpName" disabled="disabled">
                    
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Employee ID</p>
                    <input type="text" id="lbl_emp_id" class="form-input" name="txtEmpID" disabled="disabled">
                    
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Department</p>
                    <select name="lbl_emp_dept" id="sel_Dept" class="form-control" disabled="disabled"></select>
                    
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Email</p>
                    <input type="text" id="lbl_emp_email" class="form-input" name="txtEmpMail" disabled="disabled">
                    
                </div>
                <div class="col-md-4 col-12 mb-4">
                  <p>Mobile/Telephone Number</p>
                  <input type="text" id="lbl_emp_mobile" class="form-input" name="txtEmpMobile" disabled="disabled">
                  
                </div>
                
                
                <h3 class="mb-4 col-12">INCIDENT DETAILS</h3>
               
                <div class="col-md-4 col-12 mb-4">
                    <p>Incident Date & Time (From)</p>
                    <input type="text" id="lbl_date_time_req" class="form-input" name="txtEmpMobile" disabled="disabled"  autocomplete="off">
                    <label id="lbl_Inc_date" class="form-label" style="color:red"></label>
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Incident Date & Time (To)</p>
                    <input type="text" id="lbl_date_time_req_to" class="form-input" disabled="disabled"  autocomplete="off">
                    <label id="lbl_Inc_date_to" class="form-label" style="color:red"></label>
                </div>
                <div class="col-md-4 col-12 mb-4" id="div_attachment" style="display:none">
                    <p>Attachments</p>
                    <label id="lbl_attachments" class="form-label"></label>
                </div>
                
                <div class="col-md-12 col-12 mb-4">
                    <p>Reason for Request</p>
                    <textarea style="height:auto !important" rows="5" cols="5" id="txtReqReason" class="form-input" name="txtReqReason" disabled="disabled"></textarea>
                    <label id="lbl_inc_reason" class="form-label" style="color:red"></label>
                </div> 
                <h3 class="mb-4 col-12">REQUEST DETAILS</h3>
                <div class="col-md-4 col-12 mb-4">
                    <p>Request Number</p>
                    <input type="text" id="lbl_req_num" class="form-input" name="lbl_req_num" disabled="disabled">
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
    </div>
    
    <div class="container-fluid mt-5" id="div_row_sss" style="display:none">
          <div class="col-md-10 mx-auto col-12">
              <div class="row user-info">
                  <h3 class="mb-4 col-12">APPROVER ACTION</h3>
                  <div class="col-md-6 col-12 mb-4" id="div_sss_fileupload">
                      <p>Please upload footage screenshot or video<span  style="color:red">*</span></p>
                      <div class="input-group">
                        <input type="text" name="filename" class="form-control" id="file_input" readonly="" placeholder="No file selected">
                        <span class="input-group-btn">
                            <div class="btn file-btn custom-file-uploader">
                            <input type="file" className="form-control" id="sssfile"/>
                                Select a file
                            </div>
                        </span>
                    </div>    
                  </div>
                                                       
                  <div class="col-md-12">
                    <label id="lbl_sss_File_err" class="form-label mb-4" style="color: red;"></label></br>
                    <b>Note :</b><i>The allowed file types are jpg,jpeg,png,gif,mp4,mov,wmv,flv,avi and the max allowed file size is 100.0 MB</i>
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
 
  private setButtonsEventHandlers(): void 
  {
    const webPart: ViewCctvInternalDetailsWebPart = this;
    
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
    
  }
  private blob() {
    //Get the File Upload control id
    let input = <HTMLInputElement>document.getElementById("sssfile");
    fileCount = input.files.length;
    if(fileCount>0){
    for (var i = 0; i < fileCount; i++) {
       var fileName = input.files[i].name;
       var ext=fileName.replace(/^.*\./, '');
       if ($.inArray(ext, ['gif','png','jpg','jpeg','mp4','mov','wmv','flv','avi']) == -1){
        $("#lbl_sss_File_err").text("Footage must be jpg,jpeg,png,gif,mp4,mov,wmv,flv,avi format");
        $("#file_input").val(fileName);
        fileInfos.length=0;
       }else{
        $("#lbl_sss_File_err").text("");
        var filesize=input.files[0].size;
        const kb = Math.round((filesize / 1024));
        if(kb<=100024){//  checking file must be 1000 mb less than
        $("#lbl_sss_File_err").text("");
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
          $("#lbl_sss_File_err").text("Max allowed footage size is 100.0 MB"); 
          fileInfos.length=0;
        }
      }
    }
    //End of for loop
    }
    else{
      $("#file_input").val("No file Selected");
    }
  }
  private UpdateMasterList(){
     if(StatusID==2 || StatusID==8){
     // retrieve request intiated and saves SSS data
      if(fileInfos.length>0){
       this.UpdateSSSReview();
      }else{
        $("#lbl_sss_File_err").text("Footage is mandatory , must be jpg,jpeg,png,gif,mp4,mov,wmv,flv,avi format and max allowed size is 100.0 MB");
      }
    }
    else if (StatusID==3 && IsLegalTeamMember>=0){
      this.UpdateCCTVRequest();
    }
  }
  private UpdateCCTVRequest(){
   
    
      if(this.validations()==true){
      
         /* sp.site.rootWeb.lists.getByTitle(Listname).items.getById(ItemID).update({
          Title:this.valempname,
          DepartmentsId:parseInt(this.valempdept),
          EmployeeID:this.valempid,
          EmailAdress:this.valempmail,
          MobileNumber:this.valempmobile,
          ReasonForRequest:this.valReqReason,
          StatusId:1,
          DateOfIncident:this.valincDate,
          DateOfIncident_To:this.valincDateTo,
          TimeOfIncident:this.valincTimehrs+":"+this.valincTimeMins,
          TimeOfIncident_To:this.valincTimehrsTo+":"+this.valincTimeMinsTo,
          EmployeeName:this.valempname,
          AssignedToId:AssignedITGroupID,
          
        }).then(r=>{
          alert("Thank you! your request has been successfully submitted");
          window.location.href= this.context.pageContext.site.absoluteUrl+"/Pages/TecPages/CCTV-Internal-Requests.aspx";
        }).catch(function(err) {  
          console.log(err);  
          }); 
          */
    }
    else{
      alert("Sorry,Please check your form where some data is not in a valid format.");
    } 
  }

  private validations(){
    var isvalidform=true;

    /*
    this.valempname= document.getElementById('txtEmpName')["value"];
    this.valempdept= document.getElementById('sel_Dept')["value"]
    this.valempid= document.getElementById('txtEmpID')["value"];
    this.valempmail= document.getElementById('txtEmpMail')["value"]; 
    this.valempmobile= document.getElementById('txtEmpMobile')["value"];
    */
    this.valReqReason= document.getElementById('txtReqReason')["value"];

   /*  this.valincTimehrs= document.getElementById('ddlIncidentHours')["value"];
    this.valincTimeMins= document.getElementById('ddlIncidentMins')["value"];
    this.valincTimehrsTo= document.getElementById('ddlIncidentHours_to')["value"];
    this.valincTimeMinsTo= document.getElementById('ddlIncidentMins_to')["value"]; */
    this.valincDate = $('#lbl_date_time_req').datepicker('getDate');
    this.valincDateTo=$('#lbl_date_time_req_to').datepicker('getDate');
  

    if(this.valincDate==null || this.valincDate == undefined)
    {
        $("#lbl_Inc_date").text("Incident Date  (From) is mandatory");
        isvalidform = false;
    }
    else
    {
        $("#lbl_Inc_date").text("");
    }
   if(this.valincDateTo==null || this.valincDateTo == undefined)
    {
        $("#lbl_Inc_date_to").text("Incident Date  (To) is mandatory");
        isvalidform = false;
    }
    else
    {
        $("#lbl_Inc_date_to").text("");
    }

    if(this.valReqReason==""){
      $("#lbl_inc_reason").text("Reason for Request is mandatory");
      isvalidform = false;
    }
    else{
      $("#lbl_inc_reason").text(" ");
    }

   
    /* if(this.valincTimehrs=="HH"){
      $("#lbl_inc_hours").text("Hours are mandatory");
      isvalidform = false;
    }
    else{
      $("#lbl_inc_hours").text(" ");
    }
    
    if(this.valincTimeMins=="MM"){
      $("#lbl_inc_mins").text("Mins are mandatory");
      isvalidform = false;
    }
    else{
      $("#lbl_inc_mins").text(" ");
    }

    if(this.valincTimehrsTo=="HH"){
      $("#lbl_inc_hours_to").text("Hours are mandatory");
      isvalidform = false;
    }
    else{
      $("#lbl_inc_hours_to").text(" ");
    }
    
    if(this.valincTimeMinsTo=="MM"){
      $("#lbl_inc_mins_to").text("Mins are mandatory");
      isvalidform = false;
    }
    else{
      $("#lbl_inc_mins_to").text(" ");
    } */
    var momentobj1=moment(this.valincDate); 
    var momentobj2=moment(this.valincDateTo);
    if((momentobj1.format("dd/mm/yyyy")==momentobj2.format("dd/mm/yyyy"))&&(this.valincTimehrs>this.valincTimehrsTo))
    {
      $("#lbl_inc_hours_to").text("Hours are not less than from Hours");
      isvalidform = false;
    }
    else{
      $("#lbl_inc_hours_to").text(" ");
    }
    if((momentobj1.format("dd/mm/yyyy")==momentobj2.format("dd/mm/yyyy"))&&(this.valincTimehrs==this.valincTimehrsTo)&&(this.valincTimeMins>this.valincTimeMinsTo))
    {
      $("#lbl_inc_mins_to").text("Mins are not less than from Mins");
      isvalidform = false;
    }
    else{
      $("#lbl_inc_mins_to").text(" ");
    } 
    return isvalidform;
  }
  private UpdateSSSReview(){
    var fileSeqNo=No_of_Attachments+1;
          sp.site.rootWeb.lists.getByTitle(Listname).items.getById(ItemID).attachmentFiles.add(fileSeqNo+"_"+fileInfos[0].name,fileInfos[0].content)
          .then(r=>{
            alert("Thank you ! The request was updated successfully");
            window.location.href=this.context.pageContext.web.absoluteUrl+"/Pages/TecPages/common/RequestComplete.aspx";
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
      IsITTeamMember=$.inArray( ITMgrGroupName, groups );
      IsLegalTeamMember=$.inArray( LegalMgrGroupName, groups );
      IsSSSTeamMember=$.inArray( SSSGroupName, groups );
      this.getListData();
    } 
  } 

  private getListData() {
    let anchorhtml: string ='';
   
    var Url =this.context.pageContext.site.absoluteUrl+`/_api/web/lists/getbytitle('${Listname}')/items?$select=*,Attachments,AttachmentFiles,Author/Name,Author/Title,Status/ID,Status/Title,ContentType/Id,ContentType/Name,Departments/ID,Departments/Title&$expand=Departments,AttachmentFiles,Status,ContentType,Author/Id&$filter=ID eq `+ItemID;
    this.context.spHttpClient
      .get(Url, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json().then((items: any): void => {
         
            let listItems: CCTVInternalList[] = items.value;
            listItems.forEach((item: CCTVInternalList) => {
            
            //console.log(item);
            
            //loading all common fields   
             var cctv_inc_status:string;  

             if(item.Status!=null){
             StatusID=item.Status.ID;
             cctv_inc_status=item.Status.Title;
             }

             if(IsSSSTeamMember<0  && IsLegalTeamMember<0 && IsITTeamMember<0 ){
                // if unauthorized user
                alert("You don't have access,please contact administrator for more info.");
                window.location.href=this.context.pageContext.web.absoluteUrl;
              }               

             if(IsSSSTeamMember>=0 && StatusID==2){
               // show sections only to sss retrive request initiated
               $("#div_attachment").show();
               $("#div_row_sss").show();
               $("#div_row_buttons").show();
              
             }
             else if((IsSSSTeamMember>=0 && StatusID==5)||(IsITTeamMember>=0 && StatusID==5)){
                // show sections to sss footage available and Itmgr footage available
                $("#div_attachment").show();
             }
             if(StatusID==7 || StatusID==9){
               // cctv footage verified & received
               $("#div_attachment").show();
             }
             // show attachment and approver action if more info required
             if(IsSSSTeamMember>=0 && StatusID==8){
                $("#div_attachment").show();
                $("#div_row_sss").show();
                $("#div_row_buttons").show();
             }
            
             if(IsLegalTeamMember>=0 && StatusID==3){
              $("#div_attachment").show();
              $("#div_row_buttons").show();
              // making editable form once legal mgr required more info
              $("#lbl_date_time_req").removeAttr('disabled');
              $("#lbl_date_time_req_to").removeAttr('disabled');
              $("#txtReqReason").removeAttr('disabled');
             }
             

             
             
             var momentObj = moment(item.Created);           
             var formatCreatedDate=momentObj.format('DD-MM-YYYY HH:mm');

             var cctv_req_number=item.Title;
             var cctv_createdBy=item.Author.Title;
             var cctv_emp_name=item.EmployeeName!=null?item.EmployeeName:"";
             var cctv_emp_id=item.EmployeeID!=null?item.EmployeeID:"";

             var cctv_emp_dept=item.Departments!=null?item.Departments.Title:"";
             var cctv_emp_email=item.EmailAdress!=null?item.EmailAdress:"";
             var cctv_emp_mobile=item.MobileNumber !=null?item.MobileNumber:"";
            
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
             

             if(items.value[0].AttachmentFiles.length>0){
              //for(var i=0;i<items.value[0].AttachmentFiles.length;i++){
                //get latest file id
                No_of_Attachments=items.value[0].AttachmentFiles.length;
                var actualFileName=items.value[0].AttachmentFiles[No_of_Attachments-1].FileName
                var getlastfilename=actualFileName.split("_")[1];
               // AttachmentFilename=items.value[0].AttachmentFiles[No_of_Attachments-1].FileName;
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
             }
             
            //  var cctv_inc_ITMgrcmmnts=item.ITManagerComments!=null?item.ITManagerComments:"";
            //  var cctv_inc_SSScmmnts=item.SSSComments!=null ?item.SSSComments:"";
            //  var cctv_inc_LegalMgrcmnts=item.LegalManagerComments!=null?item.LegalManagerComments:"";

            //  var cctv_inc_Legal_More_info=item.MoreInfofromLegalMgr!=null?item.MoreInfofromLegalMgr:"";
            //  var cctv_inc_SSS_more_info=item.MoreInfoformSSS!=null?item.MoreInfoformSSS:"";
             
            
             
             
             // load more info sections if those are not null
            //  if(cctv_inc_SSS_more_info!=null && cctv_inc_SSS_more_info!="")
            //  { 
            //    $("#div_row_sss_more_info").show();
            //    $("#lbl_sss_more_info").html(cctv_inc_SSS_more_info);
            //  }
            //  if(cctv_inc_Legal_More_info!=null && cctv_inc_Legal_More_info!="")
            //  { 
            //    $("#div_row_Legal_more_info").show();
            //    $("#lbl_legal_mgr_more_info").html(cctv_inc_Legal_More_info);
            //  }
             
             $("#lbl_req_by").val(cctv_createdBy);
             $("#lbl_req_num").val(cctv_req_number);
             $("#lbl_emp_name").val(cctv_emp_name);
             $("#lbl_date_of_req").val(formatCreatedDate);
             

             $("#lbl_emp_id").val(cctv_emp_id);
             $("#sel_Dept").append($('<option></option>').val(cctv_emp_dept).html(cctv_emp_dept));
             $("#lbl_emp_email").val(cctv_emp_email);
             $("#lbl_emp_mobile").val(cctv_emp_mobile);
             $("#txtReqReason").val(cctv_inc_reason_for_req);
             $("#lbl_date_time_req").val(cctv_date_time);
             $("#lbl_date_time_req_to").val(cctv_date_time_to);
             $("#lbl_status").append($('<option></option>').val(cctv_inc_status).html(cctv_inc_status));
             $("#lbl_attachments").html(anchorhtml);
            
          });
          
          //this.LoadDepartments();
        }).catch(function(err) {  
          console.log(err);  
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
