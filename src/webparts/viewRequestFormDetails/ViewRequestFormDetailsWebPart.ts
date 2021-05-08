import { Version } from '@microsoft/sp-core-library';
import 'jqueryui';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ViewRequestFormDetailsWebPart.module.scss';
import * as strings from 'ViewRequestFormDetailsWebPartStrings';
import { Items, sp } from '@pnp/sp/presets/all';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { IFieldInfo } from "@pnp/sp/fields";

import * as moment from 'moment';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { RequestFormListCols } from './../Interfaces/IRequestForm';
import * as $ from 'jquery';
import { TeachingBubble } from 'office-ui-fabric-react';
import { ICCHistoryLogList } from './../Interfaces/ICCTVInternal';
export interface IViewRequestFormDetailsWebPartProps {
  description: string;
}

export interface IChoiceFieldInfo extends IFieldInfo {
  Choices: string[];
}
var todaydt = new Date();
var fileInfos = [];
var fileCount:number;
var selected_dept:string
var content_title,content_type,content_where_published,content_deadline_date,content_require_biligual,content__asap_details,content_length,content_additional_details,content_sel_bilingual_lang;
var media_title,media_dept_have_budget_for_this,media_which_dept_come_out_from,media_budget_amount,media_new_preferences;
var media_publish_date,media_text_content,media_any_additional_info;
let groups: any[] = [];
let requestFormItem:RequestFormListCols;
//var Listname = "RequestFormProcess";
var Listname="Request Forms";
//var LogsListname = "CCTVInternalIncidentLogs";
//var DocumentLibraryname = "CCTVInternalFootage";
var MarcommGroupName="MarcommTeam";
var MarcommManagerGroupName="MarcommTeamManagers";
const url : any = new URL(window.location.href);
const ItemID= url.searchParams.get("ItemID");
var StatusID:any;
var ItemContentTypeName:string;
var IsMarCommTeamMember:number;
var IsMarCommMgrMember:number;
var loginuserid;
var CT_surveyForm:string ="Survey Form";
var CT_socialMediaFrom:string ="Social Media Form";//"SocialMediaForm";
var CT_photovideo:string ="Photography And Videography Form";//"PhotographyAndVideograhpyForm";
var CT_designAndProductionForm ="Design And Production Form";//"DesignAndProductionForm";
var CT_EventForm ="Events Form";//"EventsForm";
var CT_MediaRequestForm="Media Request Form";//"MediaRequestForm";
var CT_ContentCreationForm="Content Creation Form";//"ContentCreationForm";

var Status_Request_Initiated="Request Initiated";
var Status_Request_Accepted="Request Accepted";
var Status_Request_Declined="Request Declied";
var Status_Rework_Requested="Rework Requested";

var WorkFlowLogsList="WorkflowLogs";

var html_SurveyView:string;
var html_SurveyEdit:string;
var html_VideoPhotoView:string;
var html_VideoPhotoEdit:string;
var html_SocialMediaView:string;
var html_SocialMediaEdit:string;

//private CurrentStatusId:number;

export default class ViewRequestFormDetailsWebPart extends BaseClientSideWebPart<IViewRequestFormDetailsWebPartProps> {

  private CurrentStatus:string;

  private DepartmentList:string="LK_Departments";
  //private CurrentCT:string;

  private tempHtmlString:string;

  private DeptId:string;
  private SurveyType:string;
  private SurveyReportConcluded:string;

  private deptHasBudget:string;
  private SocialMediaType:string;
  private platformsTypes:string;

  
  
  public constructor() {
    super();
    SPComponentLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="inner-page-cont" id="requestFormSection" style="margin-top:-30px">
    <div class="container-fluid mt-5">
        <div  id="dvRequestForm"></div>
        <div id="div_row_marcom_RRI" class="col-md-10 mx-auto col-12">
            <div class="row user-info">
                <div class="col-md-4 col-12 mb-4">
                    <p>Is Request Valid ?</p>
                    <input class="chk_EditmodeField" style="margin-left:50px;" type="checkbox" disabled="disabled" id="chk_isRequestvalid" name="chk_isRequestvalid" >
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Does request meet Criteria ?</p>
                    <input class="chk_EditmodeField" style="margin-left:50px;"  type="checkbox" disabled="disabled" id="chk_DoesMeetCriteria" name="chk_DoesMeetCriteria" >
                </div>
                <div class="col-md-4 col-12 mb-4" id="div_sss_fileupload" style="display:none">
                  <p>Upload Feedback</p>
                    <div class="input-group">
                      <input type="text" name="filename" class="form-control" id="file_input" readonly="" placeholder="No file selected">
                      <span class="input-group-btn"> 
                          <div class="btn file-btn custom-file-uploader">
                          <input type="file" className="form-control" id="sssfile"/>
                              Select a file
                          </div>
                      </span>
                    </div>    
                  <label id="lbl_sss_File_err" class="form-label" style="color: red;"></label>                                     
                </div>
                <div class="row" id="div_criteria_chk_lst">
                  <h3 class="mb-4 col-12">Criteria Check List</h3>
                  <div class="col-md-6 col-12 mb-4">
                      <p>Is the budget originating from MarComm&PR ?</p>
                      <input class="chk_EditmodeField" type="checkbox" style="margin-left:95px;"  id="chk_budget_originating_marcomm" disabled="disabled" name="chk_budget_originating_marcomm" >
                  </div>
                  <div class="col-md-6 col-12 mb-4">
                      <p>Does the request budget exceed 500KD ?</p>
                      <input class="chk_EditmodeField" type="checkbox" style="margin-left:95px;" id="chk_req_bud_exceed" disabled="disabled" name="chk_req_bud_exceed" >
                  </div>
                  <div class="col-md-6 col-12 mb-4">
                      <p>Does the request involve external or customer-facing communications ?</p>
                      <input class="chk_EditmodeField" type="checkbox" style="margin-left:95px;" id="chk_ext_facing_comm" disabled="disabled" name="chk_ext_facing_comm" >
                  </div>
                  <div class="col-md-6 col-12 mb-4">
                      <p>Does request include communication created or published on behalf of the CEO Office ?</p>
                      <input class=" chk_EditmodeField" type="checkbox" style="margin-left:95px;" id="chk_req_created_beh_ceo_office" disabled="disabled" name="chk_req_created_beh_ceo_office" >
                  </div>
                </div>  
            </div>
        </div>


    <div class="container-fluid mt-5" id="maindvButton">
        <div class="col-md-10 mx-auto col-12">
            <div class="row">

                <div class=" col-12 btnright">
                    <button class="red-btn red-btn-effect shadow-sm mt-4" id="btnSubmit" style="display:none;"> <span>Submit</span></button>
                    <button class="red-btn red-btn-effect shadow-sm mt-4" id="btnCancel" style="display:none;"> <span>Cancel</span></button>
                </div>
            </div>
        </div>
    </div>

    
    `;
    this.PageLoad();   
    this._setButtonEventHandlers();
    
    
  }
  private PageLoad()
  {
    this._checkUserInGroup();
  }

  private _setButtonEventHandlers(): void{
    const webpart:ViewRequestFormDetailsWebPart=this;
    this.domElement.querySelector('#btnSubmit').addEventListener('click', (e)=>{
      e.preventDefault();

      if(this.CurrentStatus== Status_Request_Initiated && IsMarCommTeamMember>=0){
        this.UpdateChecklistCriteria();
      }

      
      if(this.CurrentStatus== Status_Rework_Requested && ItemContentTypeName == CT_surveyForm)
      {
        if(this.ValidateSurveyFields())
        {
            this.UpdateSurveyCTItem();
        } else{
          alert("Sorry,Please check your form where some data is not in a valid format.");
          }
      }
     else if(ItemContentTypeName==CT_socialMediaFrom && this.CurrentStatus== Status_Rework_Requested)
      {
          
          if(this.ValidateSocialMediaFields())
          {
              this.UpdateSocialMediaCTItem(); 
          }      
          else{
            alert("Sorry,Please check your form where some data is not in a valid format.");
            }
      }
     else if(ItemContentTypeName==CT_photovideo && this.CurrentStatus== Status_Rework_Requested)
      {
        if(this.ValidatePhotoVideoFormControls()){
          this.UpdatePhotoVideoCTItem();       
        }
        else{
          alert("Sorry,Please check your form where some data is not in a valid format.");
          }
      }
     else if(ItemContentTypeName==CT_designAndProductionForm && this.CurrentStatus== Status_Rework_Requested)
      {
          if(this.ValidateDesignProductFromControls()){
            this.UpdateDesignProdCTItem();       
        }
        else{
          alert("Sorry,Please check your form where some data is not in a valid format.");
          }
      } 
      else if(ItemContentTypeName==CT_EventForm && this.CurrentStatus== Status_Rework_Requested)
      {
        if(this.ValidateEventsControls()){
          this.UpdateEventsCTItem();
        }       
      }
      else if(ItemContentTypeName==CT_MediaRequestForm && this.CurrentStatus== Status_Rework_Requested)
      {
        if(this.validateMediaNewsPaperForm()==true){
          this.UpdateMediaCTItem();       
        }
        else{
          alert("Sorry,Please check your form where some data is not in a valid format.");
          }
      } 
      else if(ItemContentTypeName==CT_ContentCreationForm && this.CurrentStatus== Status_Rework_Requested)
      {
         
          if(this.ValidateContentCreationFrom()==true){
            this.UpdateContentCreationCTItem();  
          }
          else{
            alert("Sorry,Please check your form where some data is not in a valid format.");
            }
              
      }    
     
     });

      this.domElement.querySelector('#btnCancel').addEventListener('click', (e) => {
      e.preventDefault();
      window.location.href=this.context.pageContext.site.absoluteUrl+"/Pages/TecPages/SearchRF.aspx";
    });
    this.domElement.querySelector('#chk_DoesMeetCriteria').addEventListener('change',(e)=>{
      e.preventDefault();
      if($("#chk_DoesMeetCriteria").prop("checked")==true){
        $("#div_criteria_chk_lst").show();
        // clearing all check boxes
        $('#chk_budget_originating_marcomm').prop("checked",false);
        $('#chk_ext_facing_comm').prop("checked",false);
        $('#chk_req_bud_exceed').prop("checked",false);
        $('#chk_req_created_beh_ceo_office').prop("checked",false);
      }
      else{
        $("#div_criteria_chk_lst").hide();
        // clearing all check boxes
        $('#chk_budget_originating_marcomm').prop("checked",false);
        $('#chk_ext_facing_comm').prop("checked",false);
        $('#chk_req_bud_exceed').prop("checked",false);
        $('#chk_req_created_beh_ceo_office').prop("checked",false);
      }
    });//chk_isRequestvalid
    this.domElement.querySelector('#chk_isRequestvalid').addEventListener('change',(e)=>{
      e.preventDefault();
      if($("#chk_isRequestvalid").prop("checked")==true){
        // enabling does meet criteria check box
        $('#chk_DoesMeetCriteria').prop("disabled",false);
        //clearing all check boxes
        $('#chk_DoesMeetCriteria').prop("checked",false);
        $('#chk_budget_originating_marcomm').prop("checked",false);
        $('#chk_ext_facing_comm').prop("checked",false);
        $('#chk_req_bud_exceed').prop("checked",false);
        $('#chk_req_created_beh_ceo_office').prop("checked",false);
      }
      else{
        // hiding criteria check list
        $("#div_criteria_chk_lst").hide();
        // disabling does meet criteria request if isRequestvalid is not checked
        $('#chk_DoesMeetCriteria').prop("disabled",true);
        //clearing all check boxes
        $('#chk_DoesMeetCriteria').prop("checked",false);
        $('#chk_budget_originating_marcomm').prop("checked",false);
        $('#chk_ext_facing_comm').prop("checked",false);
        $('#chk_req_bud_exceed').prop("checked",false);
        $('#chk_req_created_beh_ceo_office').prop("checked",false);
      }
    });
    this.domElement.querySelector('#sssfile').addEventListener('change', (e) => { 
      e.preventDefault();
      webpart.blob();
     });
      
  }
  private validateTextBoxDisplay(e:string,errmsg:string):void{
    //const inputElement = e.target as HTMLInputElement;
     var inputval=$('#'+e).val();
     var inputspan=$('#'+e).next("span");
     if(inputval=="")
     {
       inputspan.css("display", "block");
   
     }
     else
     {
       inputspan.css("display", "none");
     }
   }
  private validateTextBox(e:string,errmsg:string):void{
    //const inputElement = e.target as HTMLInputElement;
     var inputval=$('#'+e).val();
     var inputspan=$('#'+e).next("span");
     if(inputval=="")
     {
       inputspan.text(errmsg);
   
     }
     else
     {
       inputspan.text("");
     }
   }
   private validateDropdown(e:string,errmsg:string):void{
    //const inputElement = e.target as HTMLInputElement;
     var inputval=$('#'+e).val();
     var inputspan=$('#'+e).next("span");
     if(inputval=="0" || inputval=="HH" || inputval=="MM")
     {
       inputspan.text(errmsg);
   
     }
     else
     {
       inputspan.text("");
     }
   }
   private ValMultiDropdown(e:string,errmsg:string):void{
    var inputctrl=$('#'+e+'> option:selected');
    var inputspan=$('#'+e).next("span");
    if(inputctrl.length==0){
        inputspan.text(errmsg);
    }
    else
    {
        if(inputctrl[0].innerText=="---Select---"){
        inputspan.text(errmsg);
        }
        else{
            inputspan.text("");
        }
    }
  }
  private validateDate(e:string,errmsg:string):void{
    //const inputElement = e.target as HTMLInputElement;$('#txtIncidentDate').datepicker('getDate');
     var inputval=$('#'+e).datepicker('getDate');
     var inputspan=$('#'+e).next("span");
     if(inputval==null || inputval==undefined)
     {
       inputspan.text(errmsg);
   
     }
     else
     {
       inputspan.text("");
     }
   }

  private blob() {
    //Get the File Upload control id
    let input = <HTMLInputElement>document.getElementById("sssfile");
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
    }
  }
  private ValidateContentCreationFrom(){
    var result=true;
    content_title=document.getElementById('txtSocialMediaTitle')["value"];
    selected_dept=document.getElementById('ddlSocialDepartment')["value"];
    content_type= document.getElementById('txtContentType')["value"];
    content_where_published=document.getElementById('txtWherePublished')["value"];
    content_deadline_date=$('#calDeadline').datepicker('getDate');
    content_require_biligual=document.getElementById('ddlBilingualRequired')["value"];
    content__asap_details=document.getElementById('txtMoreDetailsforContent')["value"];

    content_additional_details=document.getElementById('txtAdditionalDetails_Social')["value"];
    content_sel_bilingual_lang=$('#ddlSelectLanguage').val();
    content_length=document.getElementById('txtLengthofContent')["value"];
    
    if(content_title==""){
    $("#txtSocialMediaTitle").next("span").text("Request Title is mandatory");
        result=false; 
    }
    else{
        $("#txtSocialMediaTitle").next("span").text(" ");
    } 
    if(selected_dept=="0"){
     $("#ddlDepartment").next("span").text("Department is mandatory");
         result=false; 
     }
     else{
         $("#ddlDepartment").next("span").text(" ");
     } 
    if(content_type=="0"){
        $("#ddl_content_type").next("span").text("Content Type is mandatory");
            result=false; 
        }
        else{
            $("#ddl_content_type").next("span").text(" ");
    }
    if(content_require_biligual=="0"){
        $("#ddlBilingualRequired").next("span").text("Do you require bilingual content ? is mandatory");
            result=false; 
        }
        else{
            $("#ddlBilingualRequired").next("span").text(" ");
    }
    if(content_deadline_date==null){
        $("#calDeadline").next("span").text("Deadline for content creation is mandatory");
     result=false; 
    }
    else{
        $("#calDeadline").next("span").text(" ");
    }
    if(content_where_published==""){
        $("#txtWherePublished").next("span").text("Where will this content be published ? is mandatory");
     result=false; 
    }
    else{
        $("#txtWherePublished").next("span").text(" ");
    }
    if(content__asap_details==""){
        $("#txtMoreDetailsforContent").next("span").text("Please provide as much detail as possible about the content requested  is mandatory");
     result=false; 
    }
    else{
        $("#txtMoreDetailsforContent").next("span").text(" ");
    }
    return result;
  }
  private validateMediaNewsPaperForm(){
    var mediaresult=true;
    selected_dept=document.getElementById('ddlSocialDepartment')["value"];
    media_title=document.getElementById('txtSocialMediaTitle')["value"];
    media_dept_have_budget_for_this=document.getElementById('ddlDeptHaveBudgetSocialMedia')["value"];
    media_which_dept_come_out_from= document.getElementById('txtWhichDeptBudgetSocialMedia')["value"];
    media_budget_amount=document.getElementById('txtBudgetAmtSocialMedia')["value"];
    media_new_preferences=document.getElementById('txtMediaPreferences')["value"]
    media_publish_date=$('#calPublishDate').datepicker('getDate');;
    media_text_content=document.getElementById('txtTextContent')["value"]; 
    media_any_additional_info=document.getElementById('txtAdditionalDetails_Social')["value"]; 

    if(media_title==""){
      $("#txtSocialMediaTitle").next("span").text("Request Title is mandatory");
      mediaresult=false; 
      }
      else{
          $("#txtSocialMediaTitle").next("span").text(" ");
      } 
      if(selected_dept=="0"){
       $("#ddlSocialDepartment").next("span").text("Department is mandatory");
       mediaresult=false; 
       }
       else{
           $("#ddlSocialDepartment").next("span").text(" ");
       } 
      if(media_dept_have_budget_for_this=="0"){
          $("#ddlDeptHaveBudgetSocialMedia").next("span").text("Does the department have a budget for this request ? is mandatory");
          mediaresult=false; 
      }
      else{
          $("#ddlDeptHaveBudgetSocialMedia").next("span").text(" ");
      }
      if(media_which_dept_come_out_from==""){
          $("#txtWhichDeptBudgetSocialMedia").next("span").text("Which department budget will this come out from ? is mandatory");
          mediaresult=false; 
          }
          else{
              $("#txtWhichDeptBudgetSocialMedia").next("span").text(" ");
      }
      if(media_budget_amount==""){
          $("#txtBudgetAmtSocialMedia").next("span").text("Budget amount is mandatory");
          mediaresult=false; 
      }
      else{
          $("#txtBudgetAmtSocialMedia").next("span").text(" ");
      }
      if(media_text_content==""){
          $("#txtTextContent").next("span").text("Text Content is mandatory");
          mediaresult=false; 
      }
      else{
          $("#txtTextContent").next("span").text(" ");
      }
      return mediaresult;
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
      IsMarCommTeamMember=$.inArray( MarcommGroupName, groups );
      IsMarCommMgrMember=$.inArray(MarcommManagerGroupName,groups);
      this.getListData();
    }
    
  } 

  private getListData() {
    var URL = "";
    URL =`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${Listname}')/items?$select=*,Author/Name,Author/Title,ContentType/Id,ContentType/Name,TECDepartment/ID,TECDepartment/Title&$expand=ContentType,Author/Id,TECDepartment&$filter=ID eq `+ItemID;
    this.context.spHttpClient
      .get(URL, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json().then((items: any): void => {
            loginuserid=this.context.pageContext.legacyPageContext["userId"];
            
           

            let listItems: RequestFormListCols[] = items.value;
            listItems.forEach((item: RequestFormListCols) => {
            // check is login user item created user or marcomm team member
            //console.log(item);
            requestFormItem=item;
            this.CurrentStatus=item.Status; 
            //updating check boxes 
            var isValidReq=item.IsRequestValid;
            var isdoesmeetCriteria=item.DoesRequestMeetCriteria;
            var Doestherequestinvolveexternalorc=item.Doestherequestinvolveexternalorc;
            var IsthebudgetoriginatingfromMarCom=item.IsthebudgetoriginatingfromMarCom;
            var Doesrequestincludecommunicationc=item.Doesrequestincludecommunicationc;
            var DoesBudgetExceed500kd=item.Does_x0020_the_x0020_request_x00;
           
           
            ItemContentTypeName=item.ContentType.Name;

            //loading all common fields
            StatusID=item.Status;
            var internalWorkflowStatusVal=item.WorkflowStatus;//
            var momentObj = moment(item.Created);           
            var formatCreatedDate=momentObj.format('DD-MM-YYYY');
            // based on list item content type disabling  divs
            if(ItemContentTypeName==CT_surveyForm)
            {
                this.LoadSurveyEditHtml();
                this.RenderAllDropdowns();
                this.loadScript();
                this.LoadSurveyControls(item);
            }
            else if(ItemContentTypeName==CT_photovideo){
                this.LoadVideoPhotoEditHtml();
                    this.RenderAllDropdowns();
                    this.loadScript();
                  
              this.LoadVideoPhotographyControls(item);
            }
            else if(ItemContentTypeName==CT_socialMediaFrom){
                this.LoadSocialMediaEditHtml();
                    this.RenderAllDropdowns();
                    this.loadScript();
                
              this.LoadSocialMediaControls(item);
            }
            else if(ItemContentTypeName==CT_EventForm){

                this.LoadEventEditHtml();
                    this.RenderAllDropdowns();
                    this.loadScript();
                
              this.LoadEventControls(item);
            }
            else if(ItemContentTypeName==CT_designAndProductionForm){

                this.LoadDesignAndProdEditHtml();
                this.RenderAllDropdowns();
                this.loadScript();
                
              this.LoadDesignAndProductionControls(item);
            }
            else if(ItemContentTypeName==CT_MediaRequestForm){
                this.LoadNewsPaperAndMediaEditHtml();
                    this.RenderAllDropdowns();
                    this.loadScript();
                
              this.LoadMediaFormControls(item);
            }

            else if(ItemContentTypeName==CT_ContentCreationForm){

                this.LoadContentFormEditHtml();
                   // this.RenderAllDropdowns();
                   // this.loadScript();
                
              this.LoadContentCreationFormControls(item);
            }

            // enabling and disabling based on login user and status 

          if(IsMarCommTeamMember>=0){  
             if(this.CurrentStatus==Status_Request_Initiated && internalWorkflowStatusVal=="Marcomm Team In Progress"){ 
                $('#btnSubmit').show();
                $('#maindvButton').show();
                $('.EditmodeField').prop("disabled",true);
                $('.updateField').prop("disabled",true);
                $('.chk_EditmodeField').prop("disabled",false);
                //showing required checkboxes//
                if (isValidReq=="Yes"){
                  $('#chk_isRequestvalid').prop("checked",true);
                }
                if(isdoesmeetCriteria=="Yes"){
                  $('#chk_DoesMeetCriteria').prop("checked",true);
                }
                if(parseInt(item.BudgetAmount)>500){
                  $("#chk_req_bud_exceed").prop("checked",true);
                }
              }
              else{
                  $('#btnSubmit').hide();
                  $('#maindvButton').hide();
                  $('.EditmodeField').prop("disabled",true);
                  $('.updateField').prop("disabled",true);
                  $("#div_criteria_chk_lst").show();
                  $('.chk_EditmodeField').prop("disabled",true);
                  if (isValidReq=="Yes"){
                    $('#chk_isRequestvalid').prop("checked",true);
                  }
                  if(isdoesmeetCriteria=="Yes"){
                    $('#chk_DoesMeetCriteria').prop("checked",true);
                  }
                  if(Doestherequestinvolveexternalorc=="Yes"){
                    $('#chk_ext_facing_comm').prop("checked",true);
                  }
                  if(IsthebudgetoriginatingfromMarCom=="Yes"){
                    $('#chk_budget_originating_marcomm').prop("checked",true);
                  }
                  if(Doesrequestincludecommunicationc=="Yes"){
                    $('#chk_req_created_beh_ceo_office').prop("checked",true);
                  }
                  if(DoesBudgetExceed500kd=="Yes"){
                    $("#chk_req_bud_exceed").prop("checked",true);
                  }
              }
              
          }
          else if(IsMarCommMgrMember>=0 ){
            $('#btnSubmit').hide();
            $('#maindvButton').hide();
            $('.EditmodeField').prop("disabled",true);
            $('.updateField').prop("disabled",true);
            $('.chk_EditmodeField').prop("disabled",true);//
            $("#div_criteria_chk_lst").show();
            if (isValidReq=="Yes"){
              $('#chk_isRequestvalid').prop("checked",true);
            }
            if(isdoesmeetCriteria=="Yes"){
              $('#chk_DoesMeetCriteria').prop("checked",true);
            }
            if(Doestherequestinvolveexternalorc=="Yes"){
              $('#chk_ext_facing_comm').prop("checked",true);
            }
            if(IsthebudgetoriginatingfromMarCom=="Yes"){
              $('#chk_budget_originating_marcomm').prop("checked",true);
            }
            if(Doesrequestincludecommunicationc=="Yes"){
              $('#chk_req_created_beh_ceo_office').prop("checked",true);
            }
            if(DoesBudgetExceed500kd=="Yes"){
              $("#chk_req_bud_exceed").prop("checked",true);
            }
            
          }
          else if(loginuserid==item.AuthorId ){
            if(this.CurrentStatus==Status_Rework_Requested){
                $('#btnSubmit').show();
                $('#maindvButton').show();
                $('.EditmodeField').prop("disabled",false);
                $('.updateField').prop("disabled",false);
                $("#div_row_marcom_RRI").hide();
            }
            else{
                $('#btnSubmit').hide();
                $('#maindvButton').hide();
                $('.EditmodeField').prop("disabled",true);
                $('.updateField').prop("disabled",true);
                $("#div_row_marcom_RRI").hide();
            }
          }
          else if( loginuserid != item.AuthorId && IsMarCommTeamMember<0 && IsMarCommMgrMember<0){
            // if unauthorized user
            alert("You don't have access,please contact administrator for more info.");
            window.location.href=this.context.pageContext.web.absoluteUrl;
          }
          
          });
          //this.getRelatedDocuments(RequestTitle);
        }).catch(function(err) {  
          console.log(err);  
        });
      });
  }

  private LoadSurveyControls(currentItem)
  {
    let listItem: RequestFormListCols = currentItem;
    console.log(listItem);

    this.CurrentStatus=listItem.Status;
    
    var SurveyStartObj = moment(listItem.SurveyStartDate);           
    var formatSurveyStartDate=SurveyStartObj.format('DD-MM-YYYY');
    var SurveyEndObj=moment(listItem.SurveyEndDate);
    var formatSurveyEndDate=SurveyEndObj.format('DD-MM-YYYY');
    
    var purposeOfSurvey=listItem.PurposeOfSurvey?listItem.PurposeOfSurvey.replace(/(<([^>]+)>)/gi, ""):"";   
    var addComments=listItem.AnyAdditionalDetails?listItem.AnyAdditionalDetails.replace(/(<([^>]+)>)/gi, ""):"";    
    
    
    
    if(this.CurrentStatus==Status_Rework_Requested && loginuserid==listItem.AuthorId )
    {
      $('#btnSubmit').show();
      $('.updateField').prop("disabled",false);

      $('#txtSurveyTitle').val(listItem.Title);
      $('#txtStatus').val(listItem.Status);
      $('#txtSurveyCreated').val(listItem.Author.Title);
    
      $('#ddlSurveyDepartment').val(listItem.TECDepartment.ID);
      $('#ddlSurveyType').val(listItem.TypeOfSurvey);
      $('#ddlRequireSurveyReport').val(listItem.DoYouRequireSurveyReport);

      $('#txtAdditionalDetails').val(addComments);
      $('#txtPurposeOfSurvey').val(purposeOfSurvey);
      $('#txtWhoIsSurveyFor').val(listItem.WhoIsSurveyFor);
      $('#calSurveyStartDate').val(formatSurveyStartDate);
      $('#calSurveyEndDate').val(formatSurveyEndDate);

    }
    else{

        $('#maindvButton').hide();
        $('.updateField').prop("disabled",true);

    

    //replace fields..
    $('#txtSurveyTitle').val(listItem.Title);
      $('#txtStatus').val(listItem.Status);
      $('#txtSurveyCreated').val(listItem.Author.Title);
    
      $('#ddlSurveyDepartment').val(listItem.TECDepartment.ID);
      $('#ddlSurveyType').val(listItem.TypeOfSurvey);
      $('#ddlRequireSurveyReport').val(listItem.DoYouRequireSurveyReport);

      $('#txtAdditionalDetails').val(addComments);
      $('#txtPurposeOfSurvey').val(purposeOfSurvey);
      $('#txtWhoIsSurveyFor').val(listItem.WhoIsSurveyFor);
      $('#calSurveyStartDate').val(formatSurveyStartDate);
      $('#calSurveyEndDate').val(formatSurveyEndDate);


    }
    

  }

  private ValidateSurveyFields()
  {
    var isValid = true;
    
    if($('#txtSurveyTitle').val()=="" || $('#txtSurveyTitle').val()==undefined)
    {
        $("#txtSurveyTitle").next("span").css("display", "block");
        isValid = false;
    }
    else
    {
        $("#txtSurveyTitle").next("span").css("display", "none");
    }

    if($('#txtPurposeOfSurvey').val()=="" || $('#txtPurposeOfSurvey').val()==undefined)
    {
        $("#txtPurposeOfSurvey").next("span").css("display", "block");
        isValid = false;
    }
    else
    {
        $("#txtPurposeOfSurvey").next("span").css("display", "none");
    }
    var dateVal = $('#calSurveyStartDate').datepicker('getDate');
        if (dateVal == null || dateVal == undefined) {
            $("#calSurveyStartDate").next("span").css("display", "block");
            isValid = false;
        }
        else {
            $("#calSurveyStartDate").next("span").css("display", "none");
        }
        var dateVal = $('#calSurveyEndDate').datepicker('getDate');
        if (dateVal == null || dateVal == undefined) {
            $("#calSurveyEndDate").next("span").css("display", "block");
            isValid = false;
        }
        else {
            $("#calSurveyEndDate").next("span").css("display", "none");
        }
        //txtWhoIsSurveyFor
        if($('#txtWhoIsSurveyFor').val()=="" || $('#txtWhoIsSurveyFor').val()==undefined)
        {
            $("#txtWhoIsSurveyFor").next("span").css("display", "block");
            isValid = false;
        }    
        else
        {
            $("#txtWhoIsSurveyFor").next("span").css("display", "none");
        }
        return isValid;
    }

private UpdateSurveyCTItem(){   
    var surveyTitle=$('#txtSurveyTitle').val();
    var  deptIDstr=$('#ddlSurveyDepartment').val().toString();
    var surveyType=$('#ddlSurveyType').val().toString();
    var startDate=$('#calSurveyStartDate').datepicker('getDate');
    var endDate=$('#calSurveyEndDate').datepicker('getDate');
    
            sp.site.rootWeb.lists.getByTitle(Listname).items.getById(ItemID).update({
              TECDepartmentId:parseInt(deptIDstr),
              Title:surveyTitle,
              TypeOfSurvey: surveyType,
              PurposeOfSurvey:$('#txtPurposeOfSurvey').val(),
              WhoIsSurveyFor:$('#txtWhoIsSurveyFor').val(),   
              SurveyStartDate:startDate,
              SurveyEndDate:endDate,                
              DoYouRequireSurveyReport:$('#ddlRequireSurveyReport').val(),
              AnyAdditionalDetails:$('#txtAdditionalDetails').val(),

            }).then(r=>{
              this.updateLogsReworkRequested();
            }).catch(function(err) {  
              console.log(err);  
              alert(err);
            });
      
}

  private UpdateSocialMediaCTItem(){   
    var surveyTitle=$('#txtSocialMediaTitle').val();
    var  deptIDstr=$('#ddlSocialDepartment').val().toString();
    var socialMediaTypeVal=$('#ddlSocialMediaType').val();
    var platformVals=$('#ddlPlatformsSocial').val();
    var postDate=$('#calDateOfPost').datepicker('getDate');
    var eventDate=$('#calDateOfEvent').datepicker('getDate');
    var influencerDate=$('#calDateOfInfluencerEngmnt').datepicker('getDate');
            sp.site.rootWeb.lists.getByTitle(Listname).items.getById(ItemID).update({
              TECDepartmentId:parseInt(deptIDstr),
              Title:surveyTitle,
              TypeOfEvent:$('#txtTypeOfEventSocial').val(),
              DoesTheDepartmentHaveaBudgetForT:$('#ddlDeptHaveBudgetSocialMedia').val(),
              WhichDepartmentBudgetWillThisCom:$('#txtWhichDeptBudgetSocialMedia').val(),
              BudgetAmount:$('#txtBudgetAmtSocialMedia').val(),
              DurationOfSponsoredAd:$('#txtDurationOfAd').val(),
              LocationOfEvent:$('#txtLocationOfEventSocial').val(),
              SocialMediaType:{ results: socialMediaTypeVal },
              Platforms:{ results: platformVals },
              DateOfPost:postDate,
              DateOfEvent:eventDate ,
              DateOfInfluencerEngagement:influencerDate,
              AnyAdditionalDetails:$('#txtAdditionalDetails_Social').val(),

             }).then(r=>{

              this.updateLogsReworkRequested();

            }).catch(function(err) {  
              console.log(err);  
            });
      // }
      // else{
      //   //$("#lbl_Innovate_first_comm_err").text(arrLang[lang]['SuggestionBox']['CommentMandatory']);
      // }
}

private UpdateEventsCTItem(){   
    var surveyTitle=$('#txtSocialMediaTitle').val();
    var  deptIDstr=$('#ddlSocialDepartment').val().toString();
    //var socialMediaTypeVal=$('#ddlSocialMediaType').val();
    var requirementVals=$('#ddlRequirements').val();
    //var postDate=$('#calDateOfPost').datepicker('getDate');
    var eventDate=$('#calEventDate').datepicker('getDate');
    //var influencerDate=$('#calDateOfInfluencerEngmnt').datepicker('getDate');
            sp.site.rootWeb.lists.getByTitle(Listname).items.getById(ItemID).update({
              TECDepartmentId:parseInt(deptIDstr),
              Title:surveyTitle,
              EventDateTime:eventDate,
              TimeOfEvent:$('#ddlIncidentHours').val()+":"+$('#ddlIncidentMins').val(),
              DoesTheDepartmentHaveaBudgetForT:$('#ddlDeptHaveBudgetSocialMedia').val(),
              WhichDepartmentBudgetWillThisCom:$('#txtWhichDeptBudgetSocialMedia').val(),
              BudgetAmount:$('#txtBudgetAmtSocialMedia').val(),

              EventDuration:$('#txtDurationOfEvent').val(),
              Location:$('#txtLocationOfEvent').val(),
              TypeOfEvent:$('#txtTypeofEvent').val(),
              Requirements:{ results: requirementVals },
              IfDecorativePleaseSpecify:$('#txtDecorativeElements').val(),
              If_x0020_Other_x0020_Please_x002:$('#txtOthers').val(),
              AnyAdditionalDetails:$('#txtAdditionalDetails_Social').val(),

            }).then(r=>{
              this.updateLogsReworkRequested();
            }).catch(function(err) {  
              console.log(err);  
            });
      // }
      // else{
      //   //$("#lbl_Innovate_first_comm_err").text(arrLang[lang]['SuggestionBox']['CommentMandatory']);
      // }
}

private UpdateDesignProdCTItem(){   
    var surveyTitle=$('#txtSocialMediaTitle').val();
    var  deptIDstr=$('#ddlSocialDepartment').val().toString();
    var typeOfdesignVals=$('#ddlTypeofDesign').val();
    
    var dateofDelivery=$('#calDateofDelivery').datepicker('getDate');
    var installDeliveryDate=$('#calInstallationDeadline').datepicker('getDate');
    //var eventDate=$('#calDateOfEvent').datepicker('getDate');
    //var influencerDate=$('#calDateOfInfluencerEngmnt').datepicker('getDate');
            sp.site.rootWeb.lists.getByTitle(Listname).items.getById(ItemID).update({
              TECDepartmentId:parseInt(deptIDstr),
              Title:surveyTitle,
              TypeOfDesign:{ results: typeOfdesignVals },
              SpecifyDecorativeElements:$('#txtDecorativeElements').val(),
              SpecifyCollateral:$('#txtOtherCollateral').val(),
              Size:$('#txtSize').val(),
              SupportingTextContentLanguage:$('#ddlSupportingText').val(),
              IllustrationReference:$('#txtIllustrationReference').val(),
              DateOfDelivery:dateofDelivery,
              WillYouRequireProduction:$('#ddlRequireProduction').val(),
              DoesTheDepartmentHaveaBudgetForT:$('#ddlSocialMediaDeptHaveBudget').val(),
              WhichDepartmentBudgetWillThisCom:$('#txtSocialMediaDeptBudgetWillCome').val(),              
              BudgetAmount:$('#txtSocialMediaBudgetAmount').val(),
              Quantity:$('#txtQuantity').val(),
              Location:$('#txtLocation').val(),
              InstallationDeadline:installDeliveryDate,             
              AnyAdditionalDetails:$('#txtAdditionalDetails_SocialMedia').val(),

            }).then(r=>{
              this.updateLogsReworkRequested();
            }).catch(function(err) {  
              console.log(err);  
            });
      // }
      // else{
      //   //$("#lbl_Innovate_first_comm_err").text(arrLang[lang]['SuggestionBox']['CommentMandatory']);
      // }
}

private UpdatePhotoVideoCTItem(){   
    var surveyTitle=$('#txtVideoTitle').val();
    var  deptIDstr=$('#ddlVideoDepartment').val().toString();
    var fromShootdate=$('#calDateOfShootFrom').datepicker('getDate');
    var toShootdate=$('#calDateOfShootTo').datepicker('getDate');
    var styleofShoot=$('#ddlStyleOfShoot').val();
    
            sp.site.rootWeb.lists.getByTitle(Listname).items.getById(ItemID).update({
              TECDepartmentId:parseInt(deptIDstr),
              Title:surveyTitle,
            
            DoesTheDepartmentHaveaBudgetForT:$('#ddlDeptHaveBudgetSocialMedia').val(),
            WhichDepartmentBudgetWillThisCom:$('#txtBudgetDept').val(),
            BudgetAmount:$('#txtBudgetAmt').val(),
            TypeOfShoot :$('#ddl_shootType').val(),
            PurposeOfShoot:$('#txtPurposeOfShoot').val(),
            DateOfShoot:fromShootdate,
            DateofShootTo:toShootdate,
            Location:$('#txtLocation').val(),
            StyleOfShoot:{ results:styleofShoot},
            WhereWillThisBePublished:$('#txtWherePublish').val(),
            IsAcastRequired:$('#ddlIsCastRequired').val(),
            AnyAdditionalDetails:$('#txtAdditionalDetails_Video').val(),

            }).then(r=>{
              this.updateLogsReworkRequested();
            }).catch(function(err) {  
              console.log(err);  
            });
      // }
      // else{
      //   //$("#lbl_Innovate_first_comm_err").text(arrLang[lang]['SuggestionBox']['CommentMandatory']);
      // }
}

private UpdateContentCreationCTItem(){   
    var surveyTitle=$('#txtVideoTitle').val();
    var  deptIDstr=$('#ddlSocialDepartment').val().toString();
    var deadlinDate=$('#calDeadline').datepicker('getDate');
    var languagVal=$('#ddlSelectLanguage').val();
    
            sp.site.rootWeb.lists.getByTitle(Listname).items.getById(ItemID).update({
              TECDepartmentId:parseInt(deptIDstr),
              Title:surveyTitle,
              ContentTypeForm:$('#txtContentType').val(),
              WhereWillThisBePublished:$('#txtWherePublished').val(),
              PleaseProvideAllDetailsForConten:$('#txtMoreDetailsforContent').val(),
              LengthOfContent: $('#txtLengthofContent').val(),
              DoYouRequireBilingualContent: $('#ddlBilingualRequired').val(),
              Language:{ results:languagVal},
              DeadLine_x0020_Content_x0020_Cre:deadlinDate,          
              AnyAdditionalDetails:$('#txtAdditionalDetails_Social').val(),

            }).then(r=>{
              this.updateLogsReworkRequested();
            }).catch(function(err) {  
              console.log(err);  
            });
      // }
      // else{
      //   //$("#lbl_Innovate_first_comm_err").text(arrLang[lang]['SuggestionBox']['CommentMandatory']);
      // }
}

private UpdateMediaCTItem(){   
    var surveyTitle=$('#txtSocialMediaTitle').val();
    var  deptIDstr=$('#ddlSocialDepartment').val().toString();   
    var publishDate=$('#calPublishDate').datepicker('getDate');
    //var eventDate=$('#calDateOfEvent').datepicker('getDate');
    //var influencerDate=$('#calDateOfInfluencerEngmnt').datepicker('getDate');
            sp.site.rootWeb.lists.getByTitle(Listname).items.getById(ItemID).update({
              TECDepartmentId:parseInt(deptIDstr),
              Title:surveyTitle,
              AnyAdditionalDetails:$('#txtAdditionalDetails_Social').val(),
              DoesTheDepartmentHaveaBudgetForT:$('#ddlDeptHaveBudgetSocialMedia').val(),
              WhichDepartmentBudgetWillThisCom:$('#txtWhichDeptBudgetSocialMedia').val(),  
              BudgetAmount:$('#txtBudgetAmtSocialMedia').val(),
              TextContent:$('#txtTextContent').val(), 
              NewspaperMediaPlatformPrefrences:$('#txtMediaPreferences').val(),
              PublishDate:publishDate,

            }).then(r=>{
              this.updateLogsReworkRequested();
            }).catch(function(err) {  
              console.log(err);  
            });
  }
      
  private UpdateChecklistCriteria(){
    sp.site.rootWeb.lists.getByTitle(Listname).items.getById(ItemID).update({
      IsRequestValid:($('#chk_isRequestvalid').prop("checked"))==true?"Yes":"No",
      DoesRequestMeetCriteria:($('#chk_DoesMeetCriteria').prop("checked"))==true?"Yes":"No",
      Doestherequestinvolveexternalorc:$('#chk_ext_facing_comm').prop("checked")==true?"Yes":"No",
      IsthebudgetoriginatingfromMarCom:$('#chk_budget_originating_marcomm').prop("checked")==true?"Yes":"No",
      Doesrequestincludecommunicationc:$('#chk_req_created_beh_ceo_office').prop("checked")==true?"Yes":"No",
      Does_x0020_the_x0020_request_x00:$("#chk_req_bud_exceed").prop("checked")==true?"Yes":"No",
      
    }).then(r=>{
      if(fileInfos.length>0){
        r.item.attachmentFiles.add(fileInfos[0].name,fileInfos[0].content)
        .then(r=>{ 
          console.log("upload feedback file uploaded successfully");
        }).catch(function(err) {  
          console.log(err);  
      });
      }
      alert("Thank you ! The request was updated successfully.");
      window.location.href=this.context.pageContext.web.absoluteUrl+"/Pages/TecPages/common/RequestComplete.aspx";
    }).catch(function(err) {  
      console.log(err);  
    });
  }


  private LoadDepartments():void{
      console.log("inside departments");
    sp.site.rootWeb.lists.getByTitle(this.DepartmentList).items.select("Title","ID").get()
    .then(function (data) {
      for (var k in data) {
      // console.log(data[k]);
        $("#ddlSurveyDepartment").append("<option value=\"" +data[k].ID + "\">" +data[k].Title + "</option>");
        $('#ddlVideoDepartment').append("<option value=\"" +data[k].ID + "\">" +data[k].Title + "</option>");
        $('#ddlSocialDepartment').append("<option value=\"" +data[k].ID + "\">" +data[k].Title + "</option>");
      }
      console.log(requestFormItem);
    if(requestFormItem)
    {
        $("#ddlSurveyDepartment").val(requestFormItem.TECDepartment.ID);
        $("#ddlVideoDepartment").val(requestFormItem.TECDepartment.ID);
        $("#ddlSocialDepartment").val(requestFormItem.TECDepartment.ID);
    }
    });
    
    
    console.log("finish departments");
  }

  private LoadChoiceColumn(ChoiceColumnName,ControlName)
  {
    var control=document.getElementById(ControlName);
    sp.site.rootWeb.lists.getByTitle(Listname).fields.getByInternalNameOrTitle(ChoiceColumnName).get().then((fieldData)=>
    
    {
      if(fieldData['Choices'].length>0)
      {
        fieldData['Choices'].forEach(element => {
             $(ControlName).append("<option value=\"" +element+ "\">" +element + "</option>");
        });
        $(ControlName).val(requestFormItem[ChoiceColumnName]);
      }

    });
     
  }

  private LoadVideoPhotographyControls(currentItem)
  {
    let listItem: RequestFormListCols = currentItem;
    console.log(listItem);

    this.CurrentStatus=listItem.Status;    
    //trim..
    var purposeOfShoot=listItem.AnyAdditionalDetails?listItem.PurposeOfShoot.replace(/(<([^>]+)>)/gi, ""):"";  

    var dateFromObj = moment(listItem.DateOfShoot);           
    var formatFromDateDate=dateFromObj.format('DD-MM-YYYY');

    var dateToObj = moment(listItem.DateofShootTo);
    var formatToDateDate=dateToObj.format('DD-MM-YYYY')
    // if(listItem.DateOfShoot)
    // {
    //     var dateFromObj = moment(listItem.DateOfShoot);           
    //     formatFromDateDate=dateFromObj.format('DD-MM-YYYY');
    // }

    var addComments=listItem.AnyAdditionalDetails?listItem.AnyAdditionalDetails.replace(/(<([^>]+)>)/gi, ""):"";   
    
    
      $('#txtVideoTitle').val(listItem.Title);
      $('#txtVideoCreated').val(listItem.Author.Title);
      $('#ddlVideoDepartment').val(listItem.TECDepartment.ID.toString());
      $('#txtStatus').val(listItem.Status);
      $('#ddl_shootType').val(listItem.TypeOfShoot);
      $('#ddlDeptHaveBudgetSocialMedia').val(listItem.DoesTheDepartmentHaveaBudgetForT);
      $('#txtBudgetDept').val(listItem.WhichDepartmentBudgetWillThisCom);
      $('#txtBudgetAmt').val(listItem.BudgetAmount);
      $('#txtPurposeOfShoot').val(purposeOfShoot);
      $('#calDateOfShootFrom').val(formatFromDateDate);   
      $('#calDateOfShootTo').val(formatToDateDate);   
      $('#txtLocation').val(listItem.Location);   
      $('#ddlStyleOfShoot').val(listItem.StyleOfShoot);   
      $('#txtWherePublish').val(listItem.WhereWillThisBePublished);   
      $('#ddlIsCastRequired').val(listItem.IsAcastRequired);
      $('#txtAdditionalDetails_Video').val(addComments);
    
    if(this.CurrentStatus==Status_Rework_Requested)
    {
      $('#btnSubmit').show();
      $("#div_row_photoVideoForm_edit").show();
      $('.updateField').prop("disabled",false);
    }else if(this.CurrentStatus==Status_Request_Initiated && IsMarCommTeamMember>=0) 
    {
      $('#btnSubmit').show();
      $('.updateField').prop("disabled",true);
    }
    else
    {
        $('.updateField').prop("disabled",true);
        $('#maindvButton').hide();
    //   $("#div_row_photoVideoForm").show();
    //   $('#lblVideoTitle').html(listItem.Title);
    //   $('#lblVideoCreated').html(listItem.Author.Title);
    //   $('#lblVideoDepartment').html(listItem.TECDepartment.Title);
    //   $('#lblStatus').html(listItem.Status);
      
    //   $('#lblTypeOfShoot').html(listItem.TypeOfShoot);
    //   $('#lblBudgetAvailable').html(listItem.DoesTheDepartmentHaveaBudgetForT);
    //   $('#lblBudgetDept').html(listItem.WhichDepartmentBudgetWillThisCom);
    //   $('#lblBudgetAmt').html(listItem.BudgetAmount);
    //   $('#lblPurposeOfShoot').html(listItem.PurposeOfShoot);

    //   $('#lblDateFrom').html(formatFromDateDate);
    //   //$('#lblDateTo').html(listItem.TECDepartment.Title);

    //   $('#lblLocation').html(listItem.Location);
    //   $('#lblStyleOfShoot').html(listItem.StyleOfShoot);
    //   $('#lblWherePublish').html(listItem.WhereWillThisBePublished);
    //   $('#lblIsCastRequired').html(listItem.IsAcastRequired);
    //   $('#lblAdditionalDetails_video').html(listItem.AnyAdditionalDetails);
    }
    

  }

  private LoadSocialMediaControls(currentItem)
  {
    let listItem: RequestFormListCols = currentItem;
    console.log(listItem);

      
    var formatFromDateofPost="";
    var formatDateofEvent="";
    var formatDateofInfluencerEngagement="";
    if(listItem.DateOfPost)
    {
      var dateofPostFromObj = moment(listItem.DateOfPost);           
      formatFromDateofPost=dateofPostFromObj.format('DD-MM-YYYY');
    }

    if(listItem.DateOfEvent)
    {
      var dateofPostFromObj = moment(listItem.DateOfEvent);           
      formatDateofEvent=dateofPostFromObj.format('DD-MM-YYYY');
    }
    
    if(listItem.DateOfInfluencerEngagement)
    {
      var dateofPostFromObj = moment(listItem.DateOfInfluencerEngagement);           
      formatDateofInfluencerEngagement=dateofPostFromObj.format('DD-MM-YYYY');
    }
    

    var addComments=listItem.AnyAdditionalDetails?listItem.AnyAdditionalDetails.replace(/(<([^>]+)>)/gi, ""):"";    

    $('#txtStatus').val(listItem.Status);
      $('#txtSocialMediaTitle').val(listItem.Title);
      $('#txtSocialCreated').val(listItem.Author.Title);
      $('#ddlSocialDepartment').val(listItem.TECDepartment.ID.toString());
      $('#ddlDeptHaveBudgetSocialMedia').val(listItem.DoesTheDepartmentHaveaBudgetForT);
      $('#txtWhichDeptBudgetSocialMedia').val(listItem.WhichDepartmentBudgetWillThisCom);
      $('#txtBudgetAmtSocialMedia').val(listItem.BudgetAmount);
      $('#calDateOfPost').val(formatFromDateofPost); 
     
      $('#ddlPlatformsSocial').val(listItem.Platforms);
      $('#txtAdditionalDetails_Social').val(addComments);
    
    if(this.CurrentStatus==Status_Rework_Requested)
    {
      $('#btnSubmit').show();
      //$("#div_row_socialMediaForm_edit").show();
      $('.updateField').prop("disabled",false);
    }
    else if(this.CurrentStatus==Status_Request_Initiated && IsMarCommTeamMember>=0){
      $('#btnSubmit').show();
      $('.updateField').prop("disabled",true);
    }
    else
    {
        $('.updateField').prop("disabled",true);
        $('#maindvButton').hide();
    }
    
    // hiding disabling based on social media type
    var socialmedialen=listItem.SocialMediaType.length;
                if(socialmedialen>=1){
                    var event=false,sponsore=false,influencer=false;
                    
                    for(var i=0;i<socialmedialen;i++){
                        if(listItem.SocialMediaType[i].toString()=="Event Coverage")
                        {
                            event=true;
                           
                        }
                        if(listItem.SocialMediaType[i].toString()=="Paid Post - Sponsored Ad")
                        {
                            sponsore=true;
                           // $("#div_dur_spon").show(); 
                        }
                        if(listItem.SocialMediaType[i].toString()=="Influencer Engagement")
                        {
                            influencer=true;
                           // $("#div_dt_inf_eng").show();
                        }
                        
                    }
                    if(event==true)
                    {
                        $('.hidecls').show();
                        $('#calDateOfEvent').val(formatDateofEvent); 
                        $('#txtTypeOfEventSocial').val(listItem.TypeOfEvent);   
                        $('#txtLocationOfEventSocial').val(listItem.LocationOfEvent);   
                        
                        // $("#calDateOfEvent").datepicker({
                        //     dateFormat: "dd/mm/yy",
                        //     endDate: todaydt,
                        //     changeMonth: true,
                        //     changeYear: true,   
                        // });
                    }
                    else{
                        $('.hidecls').hide(); 
                    }
                    if(sponsore==true){
                        $("#div_dur_spon").show(); 
                        $('#txtDurationOfAd').val(listItem.DurationOfSponsoredAd);
                    }
                    else{
                        $("#div_dur_spon").hide();
                    }
                    if(influencer==true){
                         $("#div_dt_inf_eng").show();
                         $('#calDateOfInfluencerEngmnt').val(formatDateofInfluencerEngagement); 
                        //  $("#calDateOfInfluencerEngmnt").datepicker({
                        //     dateFormat: "dd/mm/yy",
                        //     endDate: todaydt,
                        //     changeMonth: true,
                        //     changeYear: true,   
                        // });
                    }
                    else{
                        $("#div_dt_inf_eng").hide();
                    }
                 }
  }

  private LoadEventControls(currentItem)
  {
    let listItem: RequestFormListCols = currentItem;
    console.log(listItem);
    var formatDateofEvent="";
    
    if(listItem.EventDateTime)
    {
      var eventDateTimeObj = moment(listItem.EventDateTime);           
      formatDateofEvent=eventDateTimeObj.format('DD-MM-YYYY');
    }

    var addComments=listItem.AnyAdditionalDetails?listItem.AnyAdditionalDetails.replace(/(<([^>]+)>)/gi, ""):"";    
    
      $('#txtSocialMediaTitle').val(listItem.Title);
      $('#txtStatus').val(listItem.Status);
      $('#txtSocialCreated').val(listItem.Author.Title);
      $('#ddlSocialDepartment').val(listItem.TECDepartment.ID.toString());
      $('#ddlDeptHaveBudgetSocialMedia').val(listItem.DoesTheDepartmentHaveaBudgetForT);
      $('#txtWhichDeptBudgetSocialMedia').val(listItem.WhichDepartmentBudgetWillThisCom);
      $('#txtBudgetAmtSocialMedia').val(listItem.BudgetAmount);
      $('#calEventDate').val(formatDateofEvent); 
      $('#txtDurationOfEvent').val(listItem.EventDuration); 
      $('#txtLocationOfEvent').val(listItem.Location);
      $('#txtTypeofEvent').val(listItem.TypeOfEvent)
      $('#ddlRequirements').val(listItem.Requirements);   
      $('#txtDecorativeElements').val(listItem.IfDecorativePleaseSpecify);   
      $('#txtOthers').val(listItem.If_x0020_Other_x0020_Please_x002);
      $('#txtAdditionalDetails_Social').val(addComments);
      var timeofevent=listItem.TimeOfEvent!=null?listItem.TimeOfEvent:"";
      if(timeofevent!=""){
      $('#ddlIncidentHours').val(timeofevent.split(":")[0]);
      $('#ddlIncidentMins').val(timeofevent.split(":")[1]);
      }
    
     if(this.CurrentStatus==Status_Rework_Requested)
     {
         $('#btnSubmit').show();
         $('.updateField').prop("disabled",false);
     }
     else if(this.CurrentStatus==Status_Request_Initiated && IsMarCommTeamMember>=0 ){
       $('#btnSubmit').show();
       $('.updateField').prop("disabled",true);
      }
     else
     {
         $('#maindvButton').hide();
         $('.updateField').prop("disabled",true);
      
     } 
    if(listItem.IfDecorativePleaseSpecify!=null){
      $('#txtDecorativeElements').val(listItem.IfDecorativePleaseSpecify);
      $("#div_decorative_ele").show();
    }
    if(listItem.If_x0020_Other_x0020_Please_x002!=null){
      $("#div_other_ele").show();
      $('#txtOthers').val(listItem.If_x0020_Other_x0020_Please_x002);
    }
    
  }
  
  private LoadDesignAndProductionControls(currentItem)
  {
    let listItem: RequestFormListCols = currentItem;
    console.log(listItem);

      
    var formatDateofDelivery="";  
    if(listItem.DateOfDelivery)
    {
      var eventDateTimeObj = moment(listItem.DateOfDelivery);           
      formatDateofDelivery=eventDateTimeObj.format('DD-MM-YYYY');
    }
    var formatInstallationDeadline="";  
    if(listItem.InstallationDeadline)
    {
      var eventDateTimeObj = moment(listItem.InstallationDeadline);           
      formatInstallationDeadline=eventDateTimeObj.format('DD-MM-YYYY');
    }
    
    var addComments=listItem.AnyAdditionalDetails?listItem.AnyAdditionalDetails.replace(/(<([^>]+)>)/gi, ""):"";    
       $('#txtStatus').val(listItem.Status);
       $('#txtSocialMediaTitle').val(listItem.Title);
       $('#txtSocialMediaCreated').val(listItem.Author.Title);
       $('#ddlSocialDepartment').val(listItem.TECDepartment.ID.toString());
        $('#txtAdditionalDetails_SocialMedia').val(addComments);

        $('#ddlTypeofDesign').val(listItem.TypeOfDesign);
        
        if(listItem.SpecifyDecorativeElements!=null){
          $('#txtDecorativeElements').val(listItem.SpecifyDecorativeElements);
          $("#div_decorative_ele").show();
        }
        if(listItem.SpecifyCollateral!=null){
          $("#div_other_ele").show();
          $('#txtOtherCollateral').val(listItem.SpecifyCollateral);
        }
        
        $('#txtSize').val(listItem.Size);
        $('#ddlSupportingText').val(listItem.SupportingTextContentLanguage);
        $('#txtIllustrationReference').val(listItem.IllustrationReference);
        $('#calDateofDelivery').val(formatDateofDelivery);
        $('#ddlRequireProduction').val(listItem.WillYouRequireProduction);
        if(listItem.WillYouRequireProduction=="Yes")
        {
          $("#div_will_req_prod").show();
        } 
        $('#ddlSocialMediaDeptHaveBudget').val(listItem.DoesTheDepartmentHaveaBudgetForT);
        $('#txtSocialMediaDeptBudgetWillCome').val(listItem.WhichDepartmentBudgetWillThisCom);
        $('#txtSocialMediaBudgetAmount').val(listItem.BudgetAmount);
        $('#txtQuantity').val(listItem.Quantity);
        $('#txtLocation').val(listItem.Location);
        $('#txtPreferredMaterial').val(listItem.PrefferedMaterial);
        $('#calInstallationDeadline').val(formatInstallationDeadline);
    
    if(this.CurrentStatus==Status_Rework_Requested)
    {
       $('#btnSubmit').show();
       $('.updateField').prop("disabled",false);     
    
    }
    else if(this.CurrentStatus==Status_Request_Initiated && IsMarCommTeamMember>=0){
      $('#btnSubmit').show();
      $('.updateField').prop("disabled",true);
    }
    else
    {
        $('.updateField').prop("disabled",true); 
        $('#maindvButton').hide();
    }
    

  }

  private LoadMediaFormControls(currentItem)
  {
    let listItem: RequestFormListCols = currentItem;
    console.log(listItem);      
    var formatPublishDate="";
    if(listItem.PublishDate)
    {
      var dateofPostFromObj = moment(listItem.PublishDate);           
      formatPublishDate=dateofPostFromObj.format('DD-MM-YYYY');
    }

    
    var addComments=listItem.AnyAdditionalDetails?listItem.AnyAdditionalDetails.replace(/(<([^>]+)>)/gi, ""):"";  
    var textContntVal=listItem.TextContent?listItem.TextContent.replace(/(<([^>]+)>)/gi, ""):"";  
    
    $('#txtStatus').val(listItem.Status);
       $('#txtSocialMediaTitle').val(listItem.Title);
       $('#txtSocialCreated').val(listItem.Author.Title);
       $('#ddlSocialDepartment').val(listItem.TECDepartment.ID.toString());

        $('#ddlDeptHaveBudgetSocialMedia').val(listItem.DoesTheDepartmentHaveaBudgetForT);
        $('#txtWhichDeptBudgetSocialMedia').val(listItem.WhichDepartmentBudgetWillThisCom);
        $('#txtBudgetAmtSocialMedia').val(listItem.BudgetAmount);
        $('#txtTextContent').val(textContntVal);
        $('#txtMediaPreferences').val(listItem.NewspaperMediaPlatformPrefrences);
        $('#calPublishDate').val(formatPublishDate);
        $('#txtAdditionalDetails_Social').val(addComments);
    
    if(this.CurrentStatus==Status_Rework_Requested)
    {
       $('#btnSubmit').show();
       $('.updateField').prop("disabled",false);
    }
    else if(this.CurrentStatus==Status_Request_Initiated && IsMarCommTeamMember>=0){
      $('#btnSubmit').show();
      $('.updateField').prop("disabled",true);
    }
    else
    {
        $('.updateField').prop("disabled",true);
        $('#maindvButton').hide();
    
    }

  }

  private LoadContentCreationFormControls(currentItem)
  {
    let listItem: RequestFormListCols = currentItem;
    console.log(listItem);      
    //var formatFromDateofPost="";
    var formatDeadline="";
    if(listItem.DeadLine_x0020_Content_x0020_Cre)
    {
      var dateofPostFromObj = moment(listItem.DeadLine_x0020_Content_x0020_Cre);           
      formatDeadline=dateofPostFromObj.format('DD-MM-YYYY');
    }
    var addComments=listItem.AnyAdditionalDetails?listItem.AnyAdditionalDetails.replace(/(<([^>]+)>)/gi, ""):"";    
    var provideAlldetails=listItem.PleaseProvideAllDetailsForConten?listItem.PleaseProvideAllDetailsForConten.replace(/(<([^>]+)>)/gi, ""):""; 

    $('#txtSocialMediaTitle').val(listItem.Title);
       $('#txtStatus').val(listItem.Status);
       $('#txtSocialCreated').val(listItem.Author.Title);
       $('#ddlSocialDepartment').val(listItem.TECDepartment.ID.toString());
       $('#txtContentType').val(listItem.ContentTypeForm);
        $('#txtWherePublished').val(listItem.WhereWillThisBePublished);
        $('#txtMoreDetailsforContent').val(provideAlldetails);
        $('#txtLengthofContent').val(listItem.LengthOfContent);
        $('#ddlBilingualRequired').val(listItem.DoYouRequireBilingualContent);
        $('#ddlSelectLanguage').val(listItem.Language);
        $('#calDeadline').val(formatDeadline);
        $('#txtAdditionalDetails_Social').val(addComments);
        if(listItem.DoYouRequireBilingualContent=="No"){
          $("#div_CC_lang_for_bilingual").show();
        }
    if(this.CurrentStatus==Status_Rework_Requested)
    {
       $('#btnSubmit').show();
       $('.updateField').prop("disabled",false);
    
    }
    else if(this.CurrentStatus==Status_Request_Initiated && IsMarCommTeamMember>=0){
      $('#btnSubmit').show();
      $('.updateField').prop("disabled",true);
    }
    else
    {
        $('.updateField').prop("disabled",true);
        $('#maindvButton').hide();
    }
    

  }

  private LoadSurveyEditHtml()
  {
      html_SurveyEdit=`
      <div class="row user-info col-md-10 mx-auto col-12 pl-0 pr-0 m-0" id="div_row_surveyForm_edit">
    <h3 class="mb-4 col-12">REQUEST DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Requested By</p>
        <input type="text" class="form-input" name="" id="txtSurveyCreated" disabled="disabled">
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Department</p>
        <select name="ddlSurveyDepartment" id="ddlSurveyDepartment" class="form-input updateField"></select>
        <span class="error-msg" style="display:none;color:red"> Department is mandatory</span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Status</p>
        <input type="text" class="form-input" name="" id="txtStatus" disabled="disabled">
    </div>
    <h3 class="mb-4 col-12">SURVEY DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Request Title<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtSurveyTitle">
        <span class="error-msg" style="display:none;color:red"> Request Title is mandatory</span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Type of Survey<span style="color:red">*</span></p>
        <select name="selectdropdown" class="form-input updateField" id="ddlSurveyType"></select>
        <span class="error-msg" style="display:none;color:red">Type of survey is mandatory</span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Survey Start Date<span style="color:red">*</span></p>
        <input type="text" id="calSurveyStartDate"  readonly="readonly" class="form-input updateField"  autocomplete="off">
        <span class="error-msg" style="display:none;color:red">Survey start date is mandatory</span>

    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Survey End Date<span style="color:red">*</span></p>
        <input type="text" id="calSurveyEndDate"  readonly="readonly" class="form-input updateField"  autocomplete="off">
        <span class="error-msg" style="display:none;color:red">Survey end date is mandatory</span>

    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Purpose Of Survey<span style="color:red">*</span></p>
        <textarea class="form-input updateField" style="height:auto!important" rows="5" cols="5" id="txtPurposeOfSurvey"></textarea>
        <span class="error-msg" style="display:none;color:red">Purpose of survey is mandatory</span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Who is the survey for?<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtWhoIsSurveyFor">
        <span class="error-msg" style="display:none;color:red">Who is the survey for? is mandatory</span>
    </div>
    <div class="col-md-6 col-12 mb-4">
        <p>Once concluded, do you require a survey report?</p>
        <select name="selectdropdown" class="form-input updateField" id="ddlRequireSurveyReport"></select>
    </div>
    <div class="col-md-6 col-12 mb-4">
        <p>Any additional field to add</p>
        <textarea class="form-input updateField" style="height:auto!important" rows="5" cols="5" id="txtAdditionalDetails"></textarea>
    </div>
  </div>
      `;
      //$('#requestFormSection').append(html_SurveyEdit);
      $('#dvRequestForm').append(html_SurveyEdit);
      
      this.domElement.querySelector('#txtSurveyTitle').addEventListener('blur', (e) => {this.validateTextBoxDisplay("txtSurveyTitle","Request Title is mandatory") });
      this.domElement.querySelector('#txtWhoIsSurveyFor').addEventListener('blur', (e) => {this.validateTextBoxDisplay("txtWhoIsSurveyFor","Who's for Survey is mandatory") });
      this.domElement.querySelector('#txtPurposeOfSurvey').addEventListener('blur', (e) => {this.validateTextBoxDisplay("txtPurposeOfSurvey","Purpose Of Survey is mandatory") });
  }

  private LoadVideoPhotoEditHtml()
  {
      var htmlRenderEdit=`
      <div class="row user-info col-md-10 mx-auto col-12 pl-0 pr-0 m-0" id="div_row_photoVideoForm_edit">
      <h3 class="mb-4 col-12">REQUEST DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Requested By</p>
        <input type="text" class="form-control" name="" id="txtVideoCreated" disabled="disabled">
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Department<span style="color:red">*</span></p>
        <select name="ddlVideoDepartment" id="ddlVideoDepartment" class="form-input updateField"></select>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Status</p>
        <input type="text" class="form-control" name="" id="txtStatus" disabled="disabled">
    </div>
    <h3 class="mb-4 col-12">PHOTOGRAPHY/VIDEOGRAPHY DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Request Title<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtVideoTitle">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Type of Shoot<span style="color:red">*</span></p>
        <select name="ddl_shootType" id="ddl_shootType" class="form-input updateField"></select>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Does the department have a budget for this request?<span style="color:red">*</span></p>
        <select class="form-input updateField" id="ddlDeptHaveBudgetSocialMedia"></select>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Which department budget will this come out from?<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtBudgetDept">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Budget amount<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtBudgetAmt">
        <span style="color:red"></span>
    </div>
    
    <div class="col-md-4 col-12 mb-4">
        <p>Date of Shoot (From)<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField"  readonly="readonly" id="calDateOfShootFrom"  autocomplete="off">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Date of Shoot (To)<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField"  readonly="readonly" id="calDateOfShootTo"  autocomplete="off">
        <span style="color:red"></span>
    </div>

    <div class="col-md-4 col-12 mb-4">
        <p>Location<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtLocation">
        <span style="color:red"></span>
    </div>
   
    <div class="col-md-4 col-12 mb-4">
        <p>Where will this be published?<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtWherePublish">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Is a cast required?<span style="color:red">*</span></p>
        <select name="ddlStyleOfShoot" class="form-input updateField" id="ddlIsCastRequired"></select>
        <span style="color:red"></span>
    </div>
     <div class="col-md-4 col-12 mb-4">
        <p>Style of Shoot<span style="color:red">*</span></p>
        <select name="ddlStyleOfShoot" multiple="multiple" style="height:70%" class="form-input updateField" id="ddlStyleOfShoot"></select>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Purpose Of Shoot<span style="color:red">*</span></p>
        <textarea class="form-input updateField" style="height:auto!important" rows="5" cols="5" id="txtPurposeOfShoot"></textarea>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Any additional field to add</p>
        <textarea class="form-input updateField" style="height:auto!important" rows="5" cols="5" id="txtAdditionalDetails_Video"></textarea>
    </div>

</div>
      
      `;
      //$('#requestFormSection').append(html_SurveyEdit);
      $('#dvRequestForm').append(htmlRenderEdit);
      console.log("finish edit mode");
      this.domElement.querySelector('#txtVideoTitle').addEventListener('blur', (e) => {this.validateTextBox("txtVideoTitle","Request Title is mandatory") });
      //this.domElement.querySelector('#txtSocialMediaBudgetAmount').addEventListener('blur', (e) => {this.validateTextBox("txtSocialMediaBudgetAmount","Budget amount is mandatory") });
      this.domElement.querySelector('#txtBudgetAmt').addEventListener('blur', (e) => {this.validateTextBox("txtBudgetAmt","Budget amount is mandatory") });
      this.domElement.querySelector('#txtPurposeOfShoot').addEventListener('blur', (e) => {this.validateTextBox("txtPurposeOfShoot","Purpose Of Shoot is mandatory") });
      this.domElement.querySelector('#txtBudgetDept').addEventListener('blur', (e) => {this.validateTextBox("txtBudgetDept","Which department budget will this come out from ? is mandatory") });
      this.domElement.querySelector('#txtLocation').addEventListener('blur', (e) => {this.validateTextBox("txtLocation","Location is mandatory") });
      this.domElement.querySelector('#txtWherePublish').addEventListener('blur', (e) => {this.validateTextBox("txtWherePublish","Where will this be published? is mandatory") });
      
      //txtWherePublish
      this.domElement.querySelector('#ddlVideoDepartment').addEventListener('blur', (e) => {this.validateDropdown("ddlVideoDepartment","Department is mandatory") });
      this.domElement.querySelector('#ddl_shootType').addEventListener('blur', (e) => {this.validateDropdown("ddl_shootType","Type of Shoot is mandatory") });
      this.domElement.querySelector('#ddlDeptHaveBudgetSocialMedia').addEventListener('blur', (e) => {this.validateDropdown("ddlDeptHaveBudgetSocialMedia","Does the department have a budget for this request? is mandatory") });
      this.domElement.querySelector('#ddlIsCastRequired').addEventListener('blur', (e) => {this.validateDropdown("ddlIsCastRequired","Is a cast required ? is mandatory") });
     
      
      this.domElement.querySelector('#calDateOfShootFrom').addEventListener('blur', (e) => {this.validateDate("calDateOfShootFrom","Date of Shoot (From) is mandatory") });
      this.domElement.querySelector('#calDateOfShootTo').addEventListener('blur', (e) => {this.validateDate("calDateOfShootTo","Date of Shoot (To) is mandatory") });
     
      this.domElement.querySelector('#ddlStyleOfShoot').addEventListener('blur', (e) => {this.ValMultiDropdown("ddlStyleOfShoot","Style of shoot is mandatory")});
      
  }

  private LoadSocialMediaEditHtml()
  {
      var htmlRenderEdit=`
      <div class="row user-info col-md-10 mx-auto col-12 pl-0 pr-0 m-0" id="div_row_socialMediaForm_edit">
      <h3 class="mb-4 col-12">REQUEST DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Requested By</p>
        <input type="text" class="form-input" name="" id="txtSocialCreated" disabled="disabled">
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Department<span style="color:red">*</span></p>
        <select name="ddlSocialDepartment" id="ddlSocialDepartment" class="form-input updateField"></select>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Status</p>
        <input type="text" class="form-input" name="" id="txtStatus" disabled="disabled">
    </div>
    <h3 class="mb-4 col-12">SOCIAL MEDIA DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Request Title<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtSocialMediaTitle">
        <span class="error-msg" style="display:none;color:red">Request Title is mandatory</span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Social Media Type<span style="color:red">*</span></p>
        <select multiple="multiple" style="height:70%" name="ddlSocialMediaType" class="form-input updateField" id="ddlSocialMediaType"></select>
    </div>

    <div class="col-md-4 col-12 mb-4">
        <p>Does the department have a budget for this request?<span style="color:red">*</span></p>
        <select name="ddlDeptHaveBudgetSocialMedia" class="form-input updateField" id="ddlDeptHaveBudgetSocialMedia"></select>
    </div>
    <div id="div_dur_spon" class="col-md-4 col-12 mb-4" style="display:none">
        <p>Duration of Sponsored Ad</p>
        <input type="text" class="form-input updateField" name="" id="txtDurationOfAd">
    </div>
    <div id="div_dt_inf_eng"  class="col-md-4 col-12 mb-4"  style="display:none">
        <p>Date of influencer engagement  </p>
        <input type="text"  readonly="readonly"  class="form-input updateField" name="" id="calDateOfInfluencerEngmnt"  autocomplete="off">
    </div>
    
    <div class="col-md-4 col-12 mb-4 hidecls">
        <p>Date of Event</p>
        <input type="text"   readonly="readonly" class="form-input updateField" name="" id="calDateOfEvent" autocomplete="off">
    </div>
    <div class="col-md-4 col-12 mb-4 hidecls">
        <p>Type of Event</p>
        <input type="text" class="form-input updateField" name="" id="txtTypeOfEventSocial">
    </div>
    <div class="col-md-4 col-12 mb-4 hidecls">
        <p>Location of Event</p>
        <input type="text" class="form-input updateField" name="" id="txtLocationOfEventSocial">
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Which department budget will this come out from?<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtWhichDeptBudgetSocialMedia">
        <span class="error-msg" style="display:none;color:red">Which department budget will this come out from? is mandatory</span>
    </div>

    <div class="col-md-4 col-12 mb-4">
        <p> Budget amount<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtBudgetAmtSocialMedia">
        <span class="error-msg" style="display:none;color:red">Budget amount is mandatory</span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Date of post<span style="color:red">*</span></p>
        <input type="text"  readonly="readonly" class="form-input updateField" name="" id="calDateOfPost"  autocomplete="off">
        <span class="error-msg" style="display:none;color:red">Date of post is mandatory</span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Platforms<span style="color:red">*</span></p>
        <select multiple="multiple" name="ddlPlatformsSocial" class="form-input updateField" id="ddlPlatformsSocial" style="height:70%"></select>
        <span class="error-msg" style="display:none;color:red">Platforms are mandatory</span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Any additional field to add</p>
        <textarea class="form-input updateField" style="height:auto!important" rows="5" cols="5" id="txtAdditionalDetails_Social"></textarea>
    </div>

</div>      
      `;
      //$('#requestFormSection').append(html_SurveyEdit);
      $('#dvRequestForm').append(htmlRenderEdit);
      $('.hidecls').hide();

      this.domElement.querySelector('#txtSocialMediaTitle').addEventListener('blur', (e) => {this.validateTextBoxDisplay("txtSocialMediaTitle","Request Title is mandatory") });
      this.domElement.querySelector('#txtWhichDeptBudgetSocialMedia').addEventListener('blur', (e) => {this.validateTextBoxDisplay("txtWhichDeptBudgetSocialMedia","Which department budget will this come out from ? is mandatory") });
      this.domElement.querySelector('#txtBudgetAmtSocialMedia').addEventListener('blur', (e) => {this.validateTextBoxDisplay("txtBudgetAmtSocialMedia","Budget amount is mandatory") });

      //console.log("finish edit mode");
       this.domElement.querySelector('#ddlSocialMediaType').addEventListener('change', (e) => {
                e.preventDefault();
                var length=$("#ddlSocialMediaType :selected").length;
                
                 if(length>=1){
                    var event=false,sponsore=false,influencer=false;
                    
                    for(var i=0;i<length;i++){
                        if($("#ddlSocialMediaType :selected")[i].innerText=="Event Coverage")
                        {
                            event=true;
                           
                        }
                        if($("#ddlSocialMediaType :selected")[i].innerText=="Paid Post - Sponsored Ad")
                        {
                            sponsore=true;
                           // $("#div_dur_spon").show(); 
                        }
                        if($("#ddlSocialMediaType :selected")[i].innerText=="Influencer Engagement")
                        {
                            influencer=true;
                           // $("#div_dt_inf_eng").show();
                        }
                        
                    }
                    if(event==true)
                    {
                        $('.hidecls').show();
                        
                        // $("#calDateOfEvent").datepicker({
                        //     dateFormat: "dd/mm/yy",
                        //     endDate: todaydt,
                        //     changeMonth: true,
                        //     changeYear: true,   
                        // });
                    }
                    else{
                        $('.hidecls').hide(); 
                    }
                    if(sponsore==true){
                        $("#div_dur_spon").show(); 
                    }
                    else{
                        $("#div_dur_spon").hide();
                    }
                    if(influencer==true){
                         $("#div_dt_inf_eng").show();
                        //  $("#calDateOfInfluencerEngmnt").datepicker({
                        //     dateFormat: "dd/mm/yy",
                        //     endDate: todaydt,
                        //     changeMonth: true,
                        //     changeYear: true,   
                        // });
                    }
                    else{
                        $("#div_dt_inf_eng").hide();
                    }
                 }
        }); 
      
  }

  private LoadEventEditHtml()
  {
      var htmlRenderView=`
  <div class="row user-info col-md-10 mx-auto col-12 pl-0 pr-0 m-0" id="div_row_EventForm_edit">
  <h3 class="mb-4 col-12">REQUEST DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Requested By</p>
        <input type="text" class="form-control" name="" id="txtSocialCreated" disabled="disabled">
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Department<span style="color:red">*</span></p>
        <select name="ddlSocialDepartment" id="ddlSocialDepartment" class="form-input updateField"></select>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Status</p>
        <input type="text" class="form-control" name="" id="txtStatus" disabled="disabled">
    </div>
    <h3 class="mb-4 col-12">EVENT DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Request Title<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtSocialMediaTitle">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Date & Time of Event <span style="color:red">*</span></p>
        <input type="text"  readonly="readonly" class="form-input updateField" name=""  readonly="readonly" id="calEventDate"  autocomplete="off"/>
        <span style="color:red"></span>
    </div>
    <div class="col-lg-2">
    <p>HH<span  style="color:red">*</span></p>
          <select name="incidenthours" id="ddlIncidentHours" class="form-input updateField">
              <option value="HH">HH</option>
              <option value="00">00</option>
              <option value="01">01</option>
              <option value="02">02</option>
              <option value="03">03</option>
              <option value="04">04</option>
              <option value="05">05</option>
              <option value="06">06</option>
              <option value="07">07</option>
              <option value="08">08</option>
              <option value="09">09</option>
              <option value="10">10</option>
              <option value="11">11</option>
              <option value="12">12</option>
              <option value="13">13</option>
              <option value="14">14</option>
              <option value="15">15</option>
              <option value="16">16</option>
              <option value="17">17</option>
              <option value="18">18</option>
              <option value="19">19</option>
              <option value="20">20</option>
              <option value="21">21</option>
              <option value="22">22</option>
              <option value="23">23</option>
          </select>
          <span style="color:red"></span>
      </div>
      <div class="col-lg-2"> 
        <p>MM<span  style="color:red">*</span></p>
        <select name="incidentmins" id="ddlIncidentMins" class="form-input updateField">
        <option value="MM">MM</option>
        <option value="00">00</option>
        <option value="01">01</option>
        <option value="02">02</option>
        <option value="03">03</option>
        <option value="04">04</option>
        <option value="05">05</option>
        <option value="06">06</option>
        <option value="07">07</option>
        <option value="08">08</option>
        <option value="09">09</option>
        <option value="10">10</option>
        <option value="11">11</option>
        <option value="12">12</option>
        <option value="13">13</option>
        <option value="14">14</option>
        <option value="15">15</option>
        <option value="16">16</option>
        <option value="17">17</option>
        <option value="18">18</option>
        <option value="19">19</option>
        <option value="20">20</option>
        <option value="21">21</option>
        <option value="22">22</option>
        <option value="23">23</option>
        <option value="24">24</option>
        <option value="25">25</option>
        <option value="26">26</option>
        <option value="27">27</option>
        <option value="28">28</option>
        <option value="29">29</option>
        <option value="30">30</option>
        <option value="31">31</option>
        <option value="32">32</option>
        <option value="33">33</option>
        <option value="34">34</option>
        <option value="35">35</option>
        <option value="36">36</option>
        <option value="37">37</option>
        <option value="38">38</option>
        <option value="39">39</option>
        <option value="40">40</option>
        <option value="41">41</option>
        <option value="42">42</option>
        <option value="43">43</option>
        <option value="44">44</option>
        <option value="45">45</option>
        <option value="46">46</option>
        <option value="47">47</option>
        <option value="48">48</option>
        <option value="49">49</option>
        <option value="50">50</option>
        <option value="51">51</option>
        <option value="52">52</option>
        <option value="53">53</option>
        <option value="54">54</option>
        <option value="55">55</option>
        <option value="56">56</option>
        <option value="57">57</option>
        <option value="58">58</option>
        <option value="59">59</option>

      </select>
      <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Duration of event<span  style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtDurationOfEvent">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Location<span  style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtLocationOfEvent">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Does the department have a budget for this request?<span  style="color:red">*</span></p>
        <select name="ddlDeptHaveBudgetSocialMedia" class="form-input updateField" id="ddlDeptHaveBudgetSocialMedia"></select>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Which department budget will this come out from?<span  style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtWhichDeptBudgetSocialMedia">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Budget amount</p>
        <input type="text" class="form-input updateField" name="" id="txtBudgetAmtSocialMedia">
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Type of Event<span  style="color:red">*</span></p>
        <select name="txtTypeofEvent" class="form-input updateField" id="txtTypeofEvent"></select>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Requirements<span  style="color:red">*</span></p>
        <select multiple="multiple" style="height:65%" name="ddlRequirements" class="form-input updateField" id="ddlRequirements"></select>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4" id="div_decorative_ele" style="display:none">
        <p>If decorative elements please specify</p>
        <input type="text" class="form-input updateField" name="" id="txtDecorativeElements">
    </div>
    <div class="col-md-4 col-12 mb-4" id="div_other_ele" style="display:none">
        <p>if other please specify </p>
        <input type="text" class="form-input updateField" name="" id="txtOthers">

    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Any additional details to add?</p>
        <textarea class="form-input updateField" style="height:auto!important" rows="5" cols="5" id="txtAdditionalDetails_Social"></textarea>
    </div>
</div>      
      `;
      //$('#requestFormSection').append(html_SurveyView);
      $('#dvRequestForm').append(htmlRenderView);
      this.domElement.querySelector('#txtSocialMediaTitle').addEventListener('blur', (e) => {this.validateTextBox("txtSocialMediaTitle","Request Title is mandatory") });
      this.domElement.querySelector('#txtDurationOfEvent').addEventListener('blur', (e) => {this.validateTextBox("txtDurationOfEvent","Duration of event is mandatory") });
      this.domElement.querySelector('#txtLocationOfEvent').addEventListener('blur', (e) => {this.validateTextBox("txtLocationOfEvent","Location of event is mandatory") });
      this.domElement.querySelector('#txtWhichDeptBudgetSocialMedia').addEventListener('blur', (e) => {this.validateTextBox("txtWhichDeptBudgetSocialMedia","Which department budget will this come out from? is mandatory") });

      this.domElement.querySelector('#ddlIncidentHours').addEventListener('blur', (e) => {this.validateDropdown("ddlIncidentHours","Hours of event is mandatory") });
      this.domElement.querySelector('#ddlIncidentMins').addEventListener('blur', (e) => {this.validateDropdown("ddlIncidentMins","Minutes of event is mandatory") });

      this.domElement.querySelector('#ddlRequirements').addEventListener('change', (e) => {
        e.preventDefault();
        var length=$("#ddlRequirements :selected").length;
            
             if(length>=1){
                var dec=false,other=false;
                
                for(var i=0;i<length;i++){
                    if($("#ddlRequirements :selected")[i].innerText=="Other")
                    {
                        other=true;
                       
                    }
                    if($("#ddlRequirements :selected")[i].innerText=="Decorative Elements")
                    {
                        dec=true;
                      
                    }
                   
                }
            }
            if(other==true){
                $("#div_other_ele").show();
            }else{
                $("#div_other_ele").hide();
            }
            if(dec==true){
                $("#div_decorative_ele").show();
            }
            else{
                $("#div_decorative_ele").hide();
            }
      });
  }

  private LoadDesignAndProdEditHtml()
  {
      var htmlRenderView=`
      <div class="row user-info col-md-10 mx-auto col-12 pl-0 pr-0 m-0" id="div_row_designProdForm_edit">
      <h3 class="mb-4 col-12">REQUEST DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Requested By</p>
        <input type="text" class="form-control" id="txtSocialMediaCreated" disabled="disabled" />
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Department<span style="color:red">*</span></p>
        <select id="ddlSocialDepartment" class="form-control updateField"></select>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Status</p>
        <input type="text" class="form-control" id="txtStatus" disabled="disabled" />
    </div>
    <h3 class="mb-4 col-12">DESIGN AND PRODUCTION DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Request Title<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" id="txtSocialMediaTitle" />
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Type of Design<span style="color:red">*</span> </p>
        <select id="ddlTypeofDesign" multiple="multiple" class="form-input updateField" style="height: 65% !important;"></select>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4" style="display:none" id="div_dec_element">
        <p>If decorative elements please specify</p>
        <input type="text" class="form-input updateField" id="txtDecorativeElements" />
    </div>

    <div class="col-md-4 col-12 mb-4" style="display:none" id="div_dec_other">
        <p>If other collateral was selected, please specify</p>
        <input type="text" class="form-input updateField" id="txtOtherCollateral" />
    </div>

    <div class="col-md-4 col-12 mb-4">
        <p>Size</p>
        <input type="text" class="form-input updateField" id="txtSize" />
    </div>

    <div class="col-md-4 col-12 mb-4">
        <p>Supporting Text Content (English & Arabic)</p>
        <select id="ddlSupportingText" class="form-input updateField"></select>
    </div>

    <div class="col-md-4 col-12 mb-4">
        <p>Illustration Reference</p>
        <input type="text" class="form-input updateField" id="txtIllustrationReference" />
    </div>

    <div class="col-md-4 col-12 mb-4">
        <p>Date of Delivery<span style="color:red">*</span></p>
        <input class="form-input updateField" type="text" id="calDateofDelivery"  readonly="readonly"  autocomplete="off"/>
        <span style="color:red"></span>
    </div>

    <div class="col-md-4 col-12 mb-4">
        <p>Will you require production?<span style="color:red">*</span></p>
        <select id="ddlRequireProduction" class="form-input updateField"></select>
        <span style="color:red"></span>
    </div>
    <div class="row" id="div_will_req_prod" style="display:none;margin-left:1px;">
      <div class="col-md-6 col-12 mb-4">
          <p>Does the department have a budget for this request?<span style="color:red">*</span></p>
          <select id="ddlSocialMediaDeptHaveBudget" class="form-input updateField"></select>
          <span style="color:red"></span>
      </div>
      <div class="col-md-6 col-12 mb-4">
          <p>Which department budget will this come out from?<span style="color:red">*</span></p>
          <input class="form-input updateField" type="text" id="txtSocialMediaDeptBudgetWillCome" />
          <span style="color:red"></span>
      </div>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Budget amount<span style="color:red">*</span></p>
        <input class="form-input updateField" type="text" id="txtSocialMediaBudgetAmount" />
        <span style="color:red"></span>
    </div>

    <div class="col-md-4 col-12 mb-4">
        <p>Quantity<span style="color:red">*</span></p>
        <input class="form-input updateField" type="text" id="txtQuantity" />
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Location if applicable  </p>
        <input class="form-input updateField" type="text" id="txtLocation" />
    </div>

    <div class="col-md-4 col-12 mb-4">
        <p>Preffered material</p>
        <input class="form-input updateField" type="text" id="txtPreferredMaterial" />
    </div>

    <div class="col-md-4 col-12 mb-4">
        <p>Installation deadline<span style="color:red">*</span></p>
        <input class="form-input updateField" type="text" id="calInstallationDeadline"  autocomplete="off"/>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Any additional details to add?</p>
        <textarea class="form-input updateField" style="height:auto!important" rows="5" cols="5" id="txtAdditionalDetails_SocialMedia"></textarea>
    </div>
</div>
      `;
      //$('#requestFormSection').append(html_SurveyView);
      
      $('#dvRequestForm').append(htmlRenderView);
      this.domElement.querySelector('#txtSocialMediaTitle').addEventListener('blur', (e) => {this.validateTextBox("txtSocialMediaTitle","Request Title is mandatory") });
      this.domElement.querySelector('#txtSocialMediaBudgetAmount').addEventListener('blur', (e) => {this.validateTextBox("txtSocialMediaBudgetAmount","Budget amount is mandatory") });
      this.domElement.querySelector('#txtQuantity').addEventListener('blur', (e) => {this.validateTextBox("txtQuantity","Quantity is mandatory") });
      this.domElement.querySelector('#txtSocialMediaDeptBudgetWillCome').addEventListener('blur', (e) => {this.validateTextBox("txtSocialMediaDeptBudgetWillCome","Which department budget will this come out from is mandatory") });
      //
      this.domElement.querySelector('#ddlRequireProduction').addEventListener('change', (e) => {
        e.preventDefault();
        if($("#ddlRequireProduction").val()=="Yes"){
            $("#div_will_req_prod").show();
        }
        else{
            $("#div_will_req_prod").hide();
        }
    });
    this.domElement.querySelector('#ddlTypeofDesign').addEventListener('change', (e) => {
      e.preventDefault();

      var length=$("#ddlTypeofDesign :selected").length;
              
                  if(length>=1){
                   var dec=false,other=false;
                  
                     for(var i=0;i<length;i++){
                        if($("#ddlTypeofDesign :selected")[i].innerText=="Decorative Elements")
                        {
                            dec=true;
                            
                        }
                        if($("#ddlTypeofDesign :selected")[i].innerText=="Other Collateral")
                        {
                            other=true;
                          
                        }
                        
                      }
                  }
                  if(dec==true)
                  {
                      $("#div_dec_element").show();  
                    
                  }
                  else{
                      $("#div_dec_element").hide();  
                  }
                  if(other==true){
                      $("#div_dec_other").show();
                  }
                  else{
                      $("#div_dec_other").hide();
                  }
      });
      
  }

  private LoadNewsPaperAndMediaEditHtml()
  {
      var htmlRenderView=`
      <div class="row user-info col-md-10 mx-auto col-12 pl-0 pr-0 m-0" id="div_row_NewspaperForm_edit">
      <h3 class="mb-4 col-12">REQUEST DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Requested By</p>
        <input type="text" class="form-input" name="" id="txtSocialCreated" disabled="disabled">
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Department<span style="color:red">*</span></p>
        <select name="ddlSocialDepartment" id="ddlSocialDepartment" class="form-input updateField"></select>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Status</p>
        <input type="text" class="form-input" name="" id="txtStatus" disabled="disabled">
    </div>
    <h3 class="mb-4 col-12">NEWSPAPER AND MEDIA DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Request Title<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtSocialMediaTitle">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Does the department have a budget for this request?<span style="color:red">*</span></p>
        <select name="ddlDeptHaveBudgetSocialMedia" class="form-input updateField" id="ddlDeptHaveBudgetSocialMedia"></select>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Which department budget will this come out from?<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtWhichDeptBudgetSocialMedia">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Budget amount<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtBudgetAmtSocialMedia">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Newspaper/media platform prefrences</p>
        <input type="text" class="form-input updateField" name="" id="txtMediaPreferences">
    </div>

    <div class="col-md-4 col-12 mb-4">
        <p>Publish date</p>
        <input type="text"  readonly="readonly" class="form-input updateField" name="" id="calPublishDate"  autocomplete="off">
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Text Content <span style="color:red">*</span> </p>
        <textarea class="form-input updateField" style="height:auto!important" rows="5" cols="5" id="txtTextContent"></textarea>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Any additional details to add?</p>
        <textarea class="form-input updateField" style="height:auto!important" rows="5" cols="5" id="txtAdditionalDetails_Social"></textarea>
    </div>
  </div>
      `;
      //$('#requestFormSection').append(html_SurveyView);
      $('#dvRequestForm').append(htmlRenderView);
      this.domElement.querySelector('#txtSocialMediaTitle').addEventListener('blur', (e) => {this.validateTextBox("txtSocialMediaTitle","Request Title is mandatory") });
      this.domElement.querySelector('#txtWhichDeptBudgetSocialMedia').addEventListener('blur', (e) => {this.validateTextBox("txtWhichDeptBudgetSocialMedia","Which department budget will this come out from ? is mandatory") });
      this.domElement.querySelector('#txtBudgetAmtSocialMedia').addEventListener('blur', (e) => {this.validateTextBox("txtBudgetAmtSocialMedia","Budget amount is mandatory") });
      this.domElement.querySelector('#txtTextContent').addEventListener('blur', (e) => {this.validateTextBox("txtTextContent","Text Content is mandatory") });
  }
  

  private LoadContentFormEditHtml()
  {
      var htmlRenderView=`
      <div class="row user-info col-md-10 mx-auto col-12 pl-0 pr-0 m-0" id="div_row_ContentCreation_edit">
      <h3 class="mb-4 col-12">REQUEST DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Requested By</p>
        <input type="text" class="form-control" name="" id="txtSocialCreated" disabled="disabled">
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Department<span style="color:red">*</span></p>
        <select name="ddlSocialDepartment" id="ddlSocialDepartment" class="form-input updateField"></select>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Status</p>
        <input type="text" class="form-control" name="" id="txtStatus" disabled="disabled">
    </div>
    <h3 class="mb-4 col-12">CONTENT CREATION DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Request Title<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtSocialMediaTitle">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Content type<span style="color:red">*</span></p>
        <select id="txtContentType" class="form-input updateField"></select>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Where will this content be published?<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtWherePublished">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Please provide as much detail as possible about the content  requested<span style="color:red">*</span></p>
        <textarea class="form-input updateField" style="height:auto!important" rows="5" cols="5" id="txtMoreDetailsforContent"></textarea>
        <span style="color:red"></span>
    </div>

    <div class="col-md-4 col-12 mb-4">
        <p>Length of content </p>
        <input type="text" class="form-input updateField" name="" id="txtLengthofContent">
        
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Do you require bilingual content <span style="color:red">*</span></p>
        <select id="ddlBilingualRequired" class="form-input updateField"></select>
        <span style="color:red"></span>
    </div>

    <div class="col-md-4 col-12 mb-4"  style="display:none"  id="div_CC_lang_for_bilingual">
        <p>If no, which language do you require the content in</p>
        <select class="form-input updateField" id="ddlSelectLanguage" multiple="multiple" style="height:80%"></select>
    </div>

    <div class="col-md-4 col-12 mb-4">
        <p>Deadline for content creation<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="calDeadline"  readonly="readonly"  autocomplete="off">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Any additional details to add?</p>
        <textarea class="form-input updateField" style="height:auto!important" rows="5" cols="5" id="txtAdditionalDetails_Social"></textarea>
    </div>
</div>
      `;
      //$('#requestFormSection').append(html_SurveyView);
      $('#dvRequestForm').append(htmlRenderView);

      this.domElement.querySelector('#txtSocialMediaTitle').addEventListener('blur', (e) => {this.validateTextBox("txtSocialMediaTitle","Request Title is mandatory") });
      this.domElement.querySelector('#txtWherePublished').addEventListener('blur', (e) => {this.validateTextBox("txtWherePublished","Where will this content be published ? is mandatory") });
      this.domElement.querySelector('#txtMoreDetailsforContent').addEventListener('blur', (e) => {this.validateTextBox("txtMoreDetailsforContent","Please provide as much detail as possible about the content requested is mandatory") });

      this.LoadChoiceColumn("DoYouRequireBilingualContent","#ddlBilingualRequired");
      this.LoadChoiceColumn("Language","#ddlSelectLanguage");
      this.LoadChoiceColumn("ContentTypeForm","#txtContentType");
      this.LoadDepartments();
      this.domElement.querySelector('#ddlBilingualRequired').addEventListener('change', (e) => {
        e.preventDefault();
        if($("#ddlBilingualRequired").val()=="No"){
            $("#div_CC_lang_for_bilingual").show();
        }
        else{
            $("#div_CC_lang_for_bilingual").hide();
        }
    });
    $("#calDeadline").datepicker({
      dateFormat: "dd/mm/yy",
      changeMonth: true,
      changeYear: true,   
      minDate:todaydt,
      onSelect:function(){
        this.focus();
      }
      });
      
  }

  private  RenderAllDropdowns()
  {
    this.LoadDepartments();
    this.LoadChoiceColumn("DoYouRequireSurveyReport","#ddlRequireSurveyReport");
    this.LoadChoiceColumn("TypeOfSurvey","#ddlSurveyType");    
    this.LoadChoiceColumn("TypeOfShoot","#ddl_shootType");
    this.LoadChoiceColumn("IsAcastRequired","#ddlIsCastRequired");
    this.LoadChoiceColumn("StyleOfShoot","#ddlStyleOfShoot");

    this.LoadChoiceColumn("SocialMediaType","#ddlSocialMediaType");
    this.LoadChoiceColumn("DoesTheDepartmentHaveaBudgetForT","#ddlDeptHaveBudgetSocialMedia");
    this.LoadChoiceColumn("Platforms","#ddlPlatformsSocial");

    this.LoadChoiceColumn("TypeOfDesign","#ddlTypeofDesign");
    this.LoadChoiceColumn("SupportingTextContentLanguage","#ddlSupportingText");
    this.LoadChoiceColumn("WillYouRequireProduction","#ddlRequireProduction");
    this.LoadChoiceColumn("DoesTheDepartmentHaveaBudgetForT","#ddlSocialMediaDeptHaveBudget");

    //events
    this.LoadChoiceColumn("Requirements","#ddlRequirements");
    this.LoadChoiceColumn("TypeOfEvent","#txtTypeofEvent");
    //content creation
    this.LoadChoiceColumn("Language","#ddlSelectLanguage");
    this.LoadChoiceColumn("DoYouRequireBilingualContent","#ddlBilingualRequired");

  }

  private loadScript()
  {
    // $('#calSurveyStartDate').calendar({
    //     type: 'date',
    //     //endCalendar: $('#calSurveyEndDate')
    //   });
    //   $('#calSurveyEndDate').calendar({
    //     type: 'date',
    //     //startCalendar: $('#calSurveyStartDate')
    //   });
    $('#calSurveyStartDate').datepicker({
        changeMonth: true,
        changeYear: true,
        dateFormat: "dd-mm-yy",
        minDate:todaydt,
        onSelect: function ()
                {
                  var date2 = $('#calSurveyStartDate').datepicker('getDate');
                  //sets minDate to txt_date_to
                  $('#calSurveyEndDate').datepicker('option', 'minDate', date2);
                    this.focus();
                }
      });
      $('#calSurveyEndDate').datepicker({
        changeMonth: true,
        changeYear: true,
        dateFormat: "dd-mm-yy",
        onSelect: function ()
                {
                    this.focus();
                }
      });

      $('#calDateOfShootFrom').datepicker({
        changeMonth: true,
        changeYear: true,
        dateFormat: "dd-mm-yy",
        minDate:todaydt,
        onSelect: function ()
                {
                  var date2 = $('#calDateOfShootFrom').datepicker('getDate');
                  //sets minDate to txt_date_to
                  $('#calDateOfShootTo').datepicker('option', 'minDate', date2);
                    this.focus();
                }
      });
      $('#calDateOfShootTo').datepicker({
        changeMonth: true,
        changeYear: true,
        dateFormat: "dd-mm-yy",
        onSelect: function ()
                {
                    this.focus();
                }
      });

      $('#calDateOfPost').datepicker({
        changeMonth: true,
        changeYear: true,
        dateFormat: "dd-mm-yy",
        minDate:todaydt,
        onSelect: function ()
                {
                    this.focus();
                }
      });
      $('#calDateOfEvent').datepicker({
        changeMonth: true,
        changeYear: true,
        dateFormat: "dd-mm-yy",
        minDate:todaydt,
        onSelect: function ()
                {
                    this.focus();
                }
      });
      $('#calDateOfInfluencerEngmnt').datepicker({
        changeMonth: true,
        changeYear: true,
        dateFormat: "dd-mm-yy",
        minDate:todaydt,
        onSelect: function ()
                {
                    this.focus();
                }
      });

      //..
      $('#cal_influencerEngagement').datepicker({
        changeMonth: true,
        changeYear: true,
        dateFormat: "dd-mm-yy",
        minDate:todaydt,
        onSelect: function ()
                {
                    this.focus();
                }
      });
      $('#calDateofEvent').datepicker({
        changeMonth: true,
        changeYear: true,
        dateFormat: "dd-mm-yy",
        minDate:todaydt,
        onSelect: function ()
                {
                    this.focus();
                }
      });
      $('#calDateofPost').datepicker({
        changeMonth: true,
        changeYear: true,
        dateFormat: "dd-mm-yy",
        minDate:todaydt,
        onSelect: function ()
                {
                    this.focus();
                }
      });

      //Design
      $('#calDateofDelivery').datepicker({
        changeMonth: true,
        changeYear: true,
        dateFormat: "dd-mm-yy",
        minDate:todaydt,
        onSelect: function ()
                {
                    this.focus();
                }
      });

      //Events
      $('#calEventDate').datepicker({
        changeMonth: true,
        changeYear: true,
        dateFormat: "dd-mm-yy",
        minDate:todaydt,
        onSelect: function ()
                {
                    this.focus();
                }
      });

      //content
      
      $('#calDeadline').datepicker({
        changeMonth: true,
        changeYear: true,
        dateFormat: "dd-mm-yy",
        minDate:todaydt,
        onSelect: function ()
                {
                    this.focus();
                }
      });

      // Media
      $('#calPublishDate').datepicker({
        changeMonth: true,
        changeYear: true,
        dateFormat: "dd-mm-yy",
        minDate:todaydt,
        onSelect: function ()
                {
                    this.focus();
                }
      });

      // design
      $('#calInstallationDeadline').datepicker({
        changeMonth: true,
        changeYear: true,
        dateFormat: "dd-mm-yy",
        minDate:todaydt,
        onSelect: function ()
                {
                    this.focus();
                }
      });

   }

  private ValidateSocialMediaFields()
  {
    var isValid = true;
    
    if($('#txtSocialMediaTitle').val()=="" || $('#txtSocialMediaTitle').val()==undefined)
    {
        $("#txtSocialMediaTitle").next("span").css("display", "block");
        isValid = false;
    }
    else
    {
        $("#txtSocialMediaTitle").next("span").css("display", "none");
    }

    if($('#txtWhichDeptBudgetSocialMedia').val()=="" || $('#txtWhichDeptBudgetSocialMedia').val()==undefined)
    {
        $("#txtWhichDeptBudgetSocialMedia").next("span").css("display", "block");
        isValid = false;
    }
    else
    {
        $("#txtWhichDeptBudgetSocialMedia").next("span").css("display", "none");
    }


    if($('#txtBudgetAmtSocialMedia').val()=="" || $('#txtBudgetAmtSocialMedia').val()==undefined)
    {
        $("#txtBudgetAmtSocialMedia").next("span").css("display", "block");
        isValid = false;
    }
    else
    {
        $("#txtBudgetAmtSocialMedia").next("span").css("display", "none");
    }

    if($('#ddlPlatformsSocial option:selected').length ==0)
    {
        $("#ddlPlatformsSocial").next("span").css("display", "block");
        isValid = false;
    }
    else
    {
        $("#ddlPlatformsSocial").next("span").css("display", "none");
    }




    var dateVal = $('#calDateOfPost').datepicker('getDate');
        if (dateVal == null || dateVal == undefined) {
            $("#calDateOfPost").next("span").css("display", "block");
            isValid = false;
        }
        else {
            $("#calDateOfPost").next("span").css("display", "none");
        }
        
        return isValid;
   }
   private ValidatePhotoVideoFormControls(){
    var photoresult = true;
    if($("#txtVideoTitle").val()==""){
      $("#txtVideoTitle").next("span").text("Request Title is mandatory");
      photoresult=false; 
      }
      else{
          $("#txtVideoTitle").next("span").text(" ");
      } 
      if($("#ddlVideoDepartment").val()=="0"){
       $("#ddlVideoDepartment").next("span").text("Department is mandatory");
       photoresult=false; 
       }
       else{
           $("#ddlVideoDepartment").next("span").text(" ");
       } 
      if($("#ddl_shootType").val()=="0"){
          $("#ddl_shootType").next("span").text("Type of Shoot is mandatory");
          photoresult=false; 
      }
      else{
          $("#ddl_shootType").next("span").text(" ");
      }
      if($("#ddlDeptHaveBudgetSocialMedia").val()=="0"){
          $("#ddlDeptHaveBudgetSocialMedia").next("span").text("Does the department have a budget for this request? is mandatory");
          photoresult=false; 
          }
          else{
              $("#ddlDeptHaveBudgetSocialMedia").next("span").text(" ");
      }
      if($("#txtBudgetAmt").val()==""){
          $("#txtBudgetAmt").next("span").text("Budget amount is mandatory");
          photoresult=false; 
      }
      else{
          $("#txtBudgetAmt").next("span").text(" ");
      }
      var dateValTo = $('#calDateOfShootTo').datepicker('getDate');
      if(dateValTo==null){
          $("#calDateOfShootTo").next("span").text("Date of Shoot (To) is mandatory");
          photoresult=false; 
      }
      else{
          $("#calDateOfShootTo").next("span").text(" ");
      }
      var dateValFrom = $('#calDateOfShootFrom').datepicker('getDate');
      if(dateValFrom==null){
        $("#calDateOfShootFrom").next("span").text("Date of Shoot (From) is mandatory");
        photoresult=false; 
        }
        else{
            $("#calDateOfShootFrom").next("span").text(" ");
        }
        if($("#txtBudgetDept").val()==""){
            $("#txtBudgetDept").next("span").text("Which department budget will this come out from ? is mandatory");
            photoresult=false; 
        }
        else{
            $("#txtBudgetDept").next("span").text(" ");
        }
        if($("#txtLocation").val()==""){
          $("#txtLocation").next("span").text("Location is mandatory");
          photoresult=false; 
          }
          else{
              $("#txtLocation").next("span").text(" ");
          }

          if($("#txtWherePublish").val()==""){
            $("#txtWherePublish").next("span").text("Where will this be published? is mandatory");
            photoresult=false; 
            }
            else{
                $("#txtWherePublish").next("span").text(" ");
            }

            if($("#ddlIsCastRequired").val()=="0"){
            $("#ddlIsCastRequired").next("span").text("Is a cast required ? is mandatory");
            photoresult=false; 
            }
            else{
                $("#ddlIsCastRequired").next("span").text(" ");
            }
            if($("#ddlStyleOfShoot").val()==""){
            $("#ddlStyleOfShoot").next("span").text("Style of shoot is mandatory");
            photoresult=false; 
            }
            else{
                $("#ddlStyleOfShoot").next("span").text(" ");
            }
            if($("#txtPurposeOfShoot").val()==""){
            $("#txtPurposeOfShoot").next("span").text("Purpose Of Shoot is mandatory");
            photoresult=false; 
            }
            else{
                $("#txtPurposeOfShoot").next("span").text(" ");
            }
      return photoresult;
        
   }
   private ValidateDesignProductFromControls(){
    var designResult=true;

          if($("#txtSocialMediaTitle").val()==""){
          $("#txtSocialMediaTitle").next("span").text("Request Title is mandatory");
          designResult=false; 
          }
          else{
              $("#txtSocialMediaTitle").next("span").text(" ");
          } 
          if($("#ddlSocialDepartment").val()=="0"){
          $("#ddlSocialDepartment").next("span").text("Department is mandatory");
          designResult=false; 
          }
          else{
              $("#ddlSocialDepartment").next("span").text(" ");
          } 
          if($("#ddlTypeofDesign").val()==""){
              $("#ddlTypeofDesign").next("span").text("Type of Design is mandatory");
    //.val()
              designResult=false; 
          }
          else{
              $("#ddlTypeofDesign").next("span").text(" ");
          }
          var desi_date_delivery= $("#calDateofDelivery").datepicker('getDate');
          if(desi_date_delivery==null){
          $("#calDateofDelivery").next("span").text("Date of Delivery is mandatory");
          designResult=false; 
          }
          else{
              $("#calDateofDelivery").next("span").text(" ");
          } 
          if($("#ddlRequireProduction").val()=="0"){
              $("#ddlRequireProduction").next("span").text("Will you require production ? is mandatory");
              designResult=false; 
          }
          else{
              $("#ddlRequireProduction").next("span").text(" ");
          }

         if($("#ddlRequireProduction").val()=="Yes" && $("#ddlSocialMediaDeptHaveBudget").val()=="0")
         {
          $("#ddlSocialMediaDeptHaveBudget").next("span").text("Does the department have a budget for this request is mandatory");
          designResult=false; 
         }
         else{
          $("#ddlSocialMediaDeptHaveBudget").next("span").text(" ");
         }
         if($("#ddlRequireProduction").val()=="Yes" && $("#txtSocialMediaDeptBudgetWillCome").val()==""){
          $("#txtSocialMediaDeptBudgetWillCome").next("span").text("Which department budget will this come out from is mandatory");
          designResult=false; 
         }
         else{
          $("#txtSocialMediaDeptBudgetWillCome").next("span").text(" ");
         }

         

         if($("#txtSocialMediaBudgetAmount").val()==""){
          $("#txtSocialMediaBudgetAmount").next("span").text("Budget amount is mandatory");
          designResult=false; 
         }
         else{
          $("#txtSocialMediaBudgetAmount").next("span").text(" ");
         }

         if(Number($("#txtSocialMediaBudgetAmount").val())==NaN)
         {
          $("#txtSocialMediaBudgetAmount").next("span").text("Budget amount must be number");
          designResult=false; 
         }

         if($("#txtQuantity").val()==""){
          $("#txtQuantity").next("span").text("Quantity is mandatory");
          designResult=false; 
         }
         else{
          $("#txtQuantity").next("span").text(" ");
         }
         var desi_installa_deadline= $("#calInstallationDeadline").datepicker('getDate');
         if(desi_installa_deadline==null){
          $("#calInstallationDeadline").next("span").text("Installation Deadline is mandatory");
          designResult=false; 
         }
         else{
          $("#calInstallationDeadline").next("span").text(" ");
         }
      return designResult;
   }
   private ValidateEventsControls(){
    var resultevent=true;
   
    
    if($('#txtSocialMediaTitle').val()=="" || $('#txtSocialMediaTitle').val()==undefined)
    {
        $("#txtSocialMediaTitle").next("span").text("Request title is mandatory");
        resultevent = false;
    }
    else
    {
        $("#txtSocialMediaTitle").next("span").text("");
    }

    if($('#ddlSocialDepartment').val()=="0")
    {
        $("#ddlSocialDepartment").next("span").text("Department is mandatory");
        resultevent = false;
    }
    else
    {
        $("#ddlSocialDepartment").next("span").text();
    }

    var dateVal = $('#calEventDate').datepicker('getDate');
    if(dateVal==null || dateVal == undefined)
    {
        $("#calEventDate").next("span").text("Date of Event is mandatory");
        resultevent = false;
    }
    else
    {
        $("#calEventDate").next("span").text("");
    }

     if($('#txtDurationOfEvent').val()=="")
     {
         $("#txtDurationOfEvent").next("span").text("Duration of event is mandatory");
         resultevent = false;
     }    
     else
     {
         $("#txtDurationOfEvent").next("span").text("");
     }

     if($('#txtTypeofEvent').val()=="0")
     {
         $("#txtTypeofEvent").next("span").text("Type of event is mandatory");
         resultevent = false;
     }    
     else
     {
         $("#txtTypeofEvent").next("span").text("");
     }
     if($('#txtLocationOfEvent').val()=="")
     {
         $("#txtLocationOfEvent").next("span").text("Location of event is mandatory");
         resultevent = false;
     }    
     else
     {
         $("#txtLocationOfEvent").next("span").text("");
     }
     if($('#ddlDeptHaveBudgetSocialMedia').val()=="0")
     {
         $("#ddlDeptHaveBudgetSocialMedia").next("span").text("Does the department have a budget for this request? is mandatory");
         resultevent = false;
     }    
     else
     {
         $("#ddlDeptHaveBudgetSocialMedia").next("span").text("");
     }
     if($('#txtWhichDeptBudgetSocialMedia').val()=="")
     {
         $("#txtWhichDeptBudgetSocialMedia").next("span").text("Which department budget will this come out from? is mandatory");
         resultevent = false;
     }    
     else
     {
         $("#txtWhichDeptBudgetSocialMedia").next("span").text("");
     }
     if($('#ddlIncidentHours').val()=="HH")
     {
         $("#ddlIncidentHours").next("span").text("Hours of event is mandatory");
         resultevent = false;
     }    
     else
     {
         $("#ddlIncidentHours").next("span").text("");
     }
     if($('#ddlIncidentMins').val()=="MM")
     {
         $("#ddlIncidentMins").next("span").text("Mins of event is mandatory");
         resultevent = false;
     }    
     else
     {
         $("#ddlIncidentMins").next("span").text("");
     }
     if($('#ddlRequirements').val()=="")
     {
         $("#ddlRequirements").next("span").text("Requirements are mandatory");
         resultevent = false;
     }    
     else
     {
         $("#ddlRequirements").next("span").text("");
     }
    return resultevent;
   }

   private updateLogsReworkRequested(){
      var current_username=this.context.pageContext.legacyPageContext["userDisplayName"];
      this.context.spHttpClient.get(`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${WorkFlowLogsList}')/items?$select=*&$filter=((ItemID%20eq%20${ItemID})and(Title%20eq%20%27RequestForm%27)and(StatusID%20eq%204))&$top=1&$orderby=Created%20desc`, SPHttpClient.configurations.v1)      
      .then(response=>{
      return response.json()
      .then((items: any): void => {
     
      let listItems: ICCHistoryLogList[] = items["value"];
        if(listItems.length>0){
          var logItemID=listItems[0].ID;
          console.log(logItemID);

          sp.site.rootWeb.lists.getByTitle(WorkFlowLogsList).items.getById(logItemID).update({
            TaskCompletedBy:current_username,
            ApprovedDate:new Date(Date.now()),
            Comments:"---",
          }).then(r=>{
            alert("Thank you ! The request was updated successfully.");
            window.location.href=this.context.pageContext.web.absoluteUrl+"/Pages/TecPages/common/RequestComplete1.aspx?PN=SearchRF";
          }).catch(function(err) {  
            console.log(err);  
          });
        }
      });
      });
  }
}

