import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './NewRequestFormWebPart.module.scss';
import * as strings from 'NewRequestFormWebPartStrings';

import * as moment from 'moment';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { RequestFormListCols } from './../Interfaces/IRequestForm';
import { Items, sp } from '@pnp/sp/presets/all';
import * as $ from 'jquery';
import 'jqueryui';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { DateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-controls-react/lib/DateTimePicker';

import "@pnp/sp/security/web";
import "@pnp/sp/security/list";
import "@pnp/sp/security/item";
import "@pnp/sp/security";
import { IList } from "@pnp/sp/lists";

export interface INewRequestFormWebPartProps {
  description: string;
}

var GroupsforRolesAssignments = [
    { id: 64, groupName: "MarcommTeam"},
    { id: 79, groupName: "MarcommTeamManagers" },
    { id: 29, groupName: "Admins" },
 ];

let requestFormItem:RequestFormListCols;
//var Listname = "RequestFormProcess";
var Listname="Request Forms";

var CT_surveyForm:string ="SurveyForm";
var CT_photovideo:string ="PhotographyAndVideograhpyForm";
var CT_socialMediaFrom:string ="SocialMediaForm";
var CT_designAndProductionForm ="DesignAndProductionForm";
var CT_EventForm ="EventsForm";
var CT_MediaRequestForm="MediaRequestForm";
var CT_ContentCreationForm="ContentCreationForm";

var CT_surveyFormId="0x0100E33D79BCD0ED764F8B4D098910898E5200FB0CE6B88FF61C46B04EF80FD89598E2";
var CT_photovideoId="0x01004B12E34D66606144A8CDE3828F14E2EF00E713B57B2BEA084D89F9B8171AC5C25D3";//0x01004B12E34D66606144A8CDE3828F14E2EF
var CT_socialMediaFromId="0x0100319DBF368E2D484DA37DE200FC8F6A4500FD4AE5D84F7D9C4BBD769F082BEC76E3";
var CT_designAndProductionFormId="0x0100C38E1D3428BC0044B4E577C7783DAC9E00A52BD75BABEB5B469351CCCBB82FE36F";
                                  
var CT_EventFormId="0x0100D11CAAB981FD2443BC2D98C94C5B3DF000484FC50B7E3B5D4DB9B7E01A68C7C6C1";
var CT_MediaRequestFormId="0x01002A966E6C427A724689F499F6D58F2001007C0AA0B92162DB4BB1278F1F66FEFE46";
var CT_ContentCreationFormId="0x0100FB77186557BCAF48A6981530F7CC6DF9004FEEBF3985832F47ADD89267AD7013C7";

var Status_Request_Initiated="Request Initiated";
var todaydt = new Date();

var selected_dept:string,survey_type:string,sur_title:string;
var sur_start_date:Date,sur_end_date:Date;
var sur_purpose:string,sur_whos_for:string,sur_required_report:string,sur_any_additional_info:string;
var media_title,media_dept_have_budget_for_this,media_which_dept_come_out_from,media_budget_amount,media_new_preferences;
var media_publish_date,media_text_content,media_any_additional_info;
var content_title,content_type,content_where_published,content_deadline_date,content_require_biligual,content__asap_details,content_length,content_additional_details,content_sel_bilingual_lang;

var pho_vid_title,pho_vid_type_ofshoot,pho_vid_does_dept_have_bud,pho_vid_which_dept_bud_for_this,pho_bud_amt,pho_vid_shoot_date_from;
var pho_vid_shoot_date_to,pho_vid_shoot_loc,pho_vid_where_will_publish,pho_vid_is_cast_req,pho_vid_style_of_shoot,pho_vid_pur_of_shoot,pho_vid_add_info;

var desi_title,desi_type_of_design,desing_size,design_support_content,desi_illust_ref,desi_date_delivery;
var desi_will_req_prod,desi_does_dept_have_budget,design_which_dept_budget_come,desi_bud_amt,desi_quantity;
var desi_loc,desi_preff_mater,desi_installa_deadline,desi_add_info;


export default class NewRequestFormWebPart extends BaseClientSideWebPart<INewRequestFormWebPartProps> {

  private DepartmentList:string="LK_Departments";

  public constructor() {
    super();
    SPComponentLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="inner-page-cont" id="requestFormSection" style="margin-top:-30px">
    <div class="container-fluid mt-5">
        <div class="col-md-10 mx-auto col-12">

            <div class="row user-info">
                <h3 class="mb-4 col-12" style="color: #3999a7;">REQUEST DETAILS</h3>

                <div class="col-md-5 col-12 mb-4">
                    <p>Select Request Type</p>
                    <select id="ddlContentType" class="form-input">
                        <option value="1">Survey Request</option>
                        <option value="2">Social Media Request</option>
                        <option value="3">Design And Production Request</option>
                        <option value="4">Photography And Videography Request</option>
                        <option value="5">Events Request</option>
                        <option value="6">Media Request</option>
                        <option value="7">Content Creation Request</option>
                    </select>
                    <span style="color:red"></span>
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Department<span style="color:red">*</span></p>
                    <select name="ddlSurveyDepartment" id="ddlDepartment" class="form-input"></select>
                    <span style="color:red"></span>
                </div>
                <div class="col-md-3 col-12 mb-4">
                    <p>Status</p>
                    <input type="text" class="form-control" name="" id="txtStatus" value="Request Initiated" disabled="disabled">
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Requested By</p>
                    <input type="text" class="form-control" name="" id="txtCreated" disabled="disabled">
                </div>
            </div>
            <div id="dvRequestForm" style="display: flex;">
            </div>
        </div>
    </div>
    <div class="container-fluid mt-5">
        <div class="col-md-10 mx-auto col-12">
            <div class="row ">
                <div class=" col-12 btnright">
                    <button class="red-btn red-btn-effect shadow-sm mt-4" id="btnSubmit"> <span>Submit</span></button>
                    <button class="red-btn red-btn-effect shadow-sm mt-4" id="btnCancel"> <span>Cancel</span></button>
                </div>
            </div>
        </div>
    </div>
</section>
      `;

      this._setButtonEventHandlers();
      this.LoadSurveyEditHtml();
      this.LoadDepartments();
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

  
  

  private LoadContentType()  
  {
    //console.log(sp.web.lists.getByTitle(Listname).contentTypes);
  }

  private LoadDepartments():void{
  sp.site.rootWeb.lists.getByTitle(this.DepartmentList).items.select("Title","ID").get()
  .then(function (data) {
    $("#ddlDepartment").append('<option value="0">Select Department</option>');
    for (var k in data) {
     
      $("#ddlDepartment").append("<option value=\"" +data[k].ID + "\">" +data[k].Title + "</option>");
    }
  });
}
private _setButtonEventHandlers(): void{
    const webpart:NewRequestFormWebPart=this;
    this.domElement.querySelector('#btnSubmit').addEventListener('click', (e)=>{
      e.preventDefault();
      var requestValue=$('#ddlContentType').val();
      if(requestValue=="1"){
            if(this.validateSurvey()==true){
                this.SubmitSurveyForm();
            }  
            else{
                alert("Validation errors found.");
            }
       }
       else if(requestValue=="2"){
            if(this.ValidateSocialMediaFields()==true)
            {
                this.SubmitSocialMediaCTItem(); 
            }
            else{
                alert("Validation errors found.");
            } 
       }
       else if(requestValue=="3"){
           if(this.ValidateDesignProdFrom()==true)
           {
            this.SubmitDesignProdItem();
           }
           else{
            alert("Validation errors found.");
            }
       }
       else if(requestValue=="4")
       {
           if(this.validatePhotoVideoForm()==true){
                this.SubmitPhotoCreationItem();
           }else{
                alert("Validation errors found.");
            }
       }
       else if(requestValue=="5")
       {
           if(this.validateEventsForm()==true)
           {
               this.SubmitEventItem();
           }else{
                alert("Validation errors found.");
            }

       }
       else if(requestValue=="6")
       {
           if(this.validateMediaNewsPaperForm()==true)
           {
               this.SubmitMediaNewsPaperCTItem();
           }
           else{
                alert("Validation errors found.");
            }

       }
       else if(requestValue=="7"){
            if(this.ValidateContentCreationFrom()==true){
                this.SubmitContentCreationItem();
            }
            else{
                alert("Validation errors found.");
            }
       }
    });

      this.domElement.querySelector('#btnCancel').addEventListener('click', (e) => {
      e.preventDefault();
      window.location.href=this.context.pageContext.site.absoluteUrl+"/Pages/TecPages/SearchRF.aspx";
    });
    this.domElement.querySelector('#ddlContentType').addEventListener('change', (e) => {
       
        var requestValue=$('#ddlContentType').val();
        if(requestValue=="1"){
        $("#dvRequestForm").empty();
        this.LoadSurveyEditHtml();
        }else if(requestValue=="2"){
            $("#dvRequestForm").empty();
        this.LoadSocialMediaEditHtml();
        }else if(requestValue=="3"){
            $("#dvRequestForm").empty();
        this.LoadDesignAndProdEditHtml();
        }
        else if(requestValue=="4"){
            $("#dvRequestForm").empty();
        this.LoadVideoPhotoEditHtml();
        }
        else if(requestValue=="5"){
            $("#dvRequestForm").empty();
        this.LoadEventEditHtml();
        }
        else if(requestValue=="6"){
            $("#dvRequestForm").empty();
        this.LoadNewsPaperAndMediaEditHtml();
        }
        else if(requestValue=="7"){
            $("#dvRequestForm").empty();
        this.LoadContentFormEditHtml();
        }
      });
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
 
  private SubmitSocialMediaCTItem(){
    var surveyTitle=$('#txtSocialMediaTitle').val();
    var  deptIDstr=$('#ddlDepartment').val().toString();
    var socialMediaTypeVal=$('#ddlSocialMediaType').val();
    var platformVals=$('#ddlPlatformsSocial').val();
    var postDate=$('#calDateOfPost').datepicker('getDate');
    var eventDate=$('#calDateOfEvent').datepicker('getDate');
    var influencerDate=$('#calDateOfInfluencerEngmnt').datepicker('getDate');
            sp.site.rootWeb.lists.getByTitle(Listname).items.add({
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
              ContentTypeId:CT_socialMediaFromId
             }).then(r=>{

                alert("Thank you. The request was submitted successfully.");
                window.location.href=this.context.pageContext.site.absoluteUrl+"/Pages/TecPages/SearchRF.aspx";

            }).catch(function(err) {  
              console.log(err);  
            });
  }
  private LoadVideoPhotoEditHtml()
  {
      var htmlRenderEdit=`
      <div class="row user-info" id="div_row_photoVideoForm_edit">
    <h3 class="mb-4 col-12">PHOTOGRAPHY/VIDEOGRAPHY DETAILS</h3>
     <div class="col-md-4 col-12 mb-4">
        <p>Request Title<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txt_photo_req_title">
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
        <span style="color:red"></span>
    </div>
   
    <div class="col-md-4 col-12 mb-4">
        <p>Budget amount<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtBudgetAmt">
        <span style="color:red"></span>
    </div>
   
    <div class="col-md-4 col-12 mb-4">
        <p>Date of Shoot (From)<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" readonly="readonly" autocomplete="off" id="calDateOfShootFrom">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Date of Shoot (To)<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" readonly="readonly" autocomplete="off" id="calDateOfShootTo">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Which department budget will this come out from?<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtBudgetDept">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p >Location<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtLocation">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p  >Where will this be published?<span style="color:red">*</span></p>
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

      this.domElement.querySelector('#txt_photo_req_title').addEventListener('blur', (e) => {this.validateTextBox("txt_photo_req_title","Request Title is mandatory") });
      //this.domElement.querySelector('#txtSocialMediaBudgetAmount').addEventListener('blur', (e) => {this.validateTextBox("txtSocialMediaBudgetAmount","Budget amount is mandatory") });
      this.domElement.querySelector('#txtBudgetAmt').addEventListener('blur', (e) => {this.validateTextBox("txtBudgetAmt","Budget amount is mandatory") });
      this.domElement.querySelector('#txtPurposeOfShoot').addEventListener('blur', (e) => {this.validateTextBox("txtPurposeOfShoot","Purpose Of Shoot is mandatory") });
      this.domElement.querySelector('#txtBudgetDept').addEventListener('blur', (e) => {this.validateTextBox("txtBudgetDept","Which department budget will this come out from ? is mandatory") });
      this.domElement.querySelector('#txtLocation').addEventListener('blur', (e) => {this.validateTextBox("txtLocation","Location is mandatory") });
      this.domElement.querySelector('#txtWherePublish').addEventListener('blur', (e) => {this.validateTextBox("txtWherePublish","Where will this be published? is mandatory") });
      
      //txtWherePublish
      this.domElement.querySelector('#ddlDepartment').addEventListener('blur', (e) => {this.validateDropdown("ddlDepartment","Department is mandatory") });
      this.domElement.querySelector('#ddl_shootType').addEventListener('blur', (e) => {this.validateDropdown("ddl_shootType","Type of Shoot is mandatory") });
      this.domElement.querySelector('#ddlDeptHaveBudgetSocialMedia').addEventListener('blur', (e) => {this.validateDropdown("ddlDeptHaveBudgetSocialMedia","Does the department have a budget for this request? is mandatory") });
      this.domElement.querySelector('#ddlIsCastRequired').addEventListener('blur', (e) => {this.validateDropdown("ddlIsCastRequired","Is a cast required ? is mandatory") });
     
      
      this.domElement.querySelector('#calDateOfShootFrom').addEventListener('blur', (e) => {this.validateDate("calDateOfShootFrom","Date of Shoot (From) is mandatory") });
      this.domElement.querySelector('#calDateOfShootTo').addEventListener('blur', (e) => {this.validateDate("calDateOfShootTo","Date of Shoot (To) is mandatory") });
     
      this.domElement.querySelector('#ddlStyleOfShoot').addEventListener('blur', (e) => {this.ValMultiDropdown("ddlStyleOfShoot","Style of shoot is mandatory")});


      console.log("finish edit mode");
      this.LoadChoiceColumn("TypeOfShoot","#ddl_shootType");
      this.LoadChoiceColumn("DoesTheDepartmentHaveaBudgetForT","#ddlDeptHaveBudgetSocialMedia");

      this.LoadChoiceColumn("StyleOfShoot","#ddlStyleOfShoot");
      this.LoadChoiceColumn("IsAcastRequired","#ddlIsCastRequired");
         
      var todaydt = new Date();
      $("#calDateOfShootFrom").datepicker({
          dateFormat: "dd/mm/yy",
          endDate: todaydt,
          changeMonth: true,
          changeYear: true,   
          minDate:todaydt,
          onSelect: function (date) {
          //Get selected date 
              var date2 = $('#calDateOfShootFrom').datepicker('getDate');
              //sets minDate to txt_date_to
              $('#calDateOfShootTo').datepicker('option', 'minDate', date2);
             this.focus();
          }
      });
      $('#calDateOfShootTo').datepicker({
          dateFormat: "dd/mm/yy",
          changeMonth: true,
          changeYear: true,
          onSelect:function(){
            this.focus();
          }
      });
  
  }
  private LoadSurveyEditHtml()
  {
      var html_RenderEdit=`
      <div class="row user-info" id="div_row_surveyForm_new">
        <h3 class="mb-4 col-12">SURVEY DETAILS</h3>
        <div class="col-md-4 col-12 mb-4">
            <p>Request Title<span style="color:red">*</span></p>
            <input type="text" class="form-input updateField" name="" id="txtSurveyTitle">
            <span style="color:red"></span>
        </div>      
        <div class="col-md-4 col-12 mb-4">
            <p>Type of Survey<span style="color:red">*</span></p>
            <select name="selectdropdown" class="form-input" id="ddlSurveyType"></select>
            <span style="color:red"></span>
        </div>
        <div class="col-md-4 col-12 mb-4">
            <p>Who is the survey for?<span style="color:red">*</span></p>
            <input type="text" class="form-input" name="" id="txtWhoIsSurveyFor">
            <span style="color:red"></span>
        </div>
        <div class="col-md-4 col-12 mb-4">
            <p>Survey Start Date<span style="color:red">*</span></p>
            <input type="text" id="calSurveyStartDate" readonly="readonly" autocomplete="off" class="form-input">
            <span style="color:red"></span>
        </div>
        <div class="col-md-4 col-12 mb-4">
            <p>Survey End Date<span style="color:red">*</span></p>
            <input type="text" id="calSurveyEndDate" readonly="readonly" autocomplete="off"  class="form-input">
            <span style="color:red"></span>
        </div> 
        <div class="col-md-4 col-12 mb-4">
            <p>Do you require a survey report?</p>
            <select name="selectdropdown" class="form-input" id="ddlRequireSurveyReport"></select>
        </div>
         <div class="col-md-6 col-12 mb-4">
            <p>Purpose Of Survey<span style="color:red">*</span></p>
            <textarea class="form-input" style="height:auto!important" rows="5" cols="5" id="txtPurposeOfSurvey"></textarea>
            <span style="color:red"></span>
        </div>
        <div class="col-md-6 col-12 mb-4">
            <p>Any additional field to add</p>
            <textarea class="form-input" style="height:auto!important" rows="5" cols="5" id="txtAdditionalDetails_Survey"></textarea>
        </div>
    </div>
      `;
        $("#txtCreated").val(this.context.pageContext.user.displayName);
        $('#dvRequestForm').append(html_RenderEdit);      

            this.domElement.querySelector('#txtSurveyTitle').addEventListener('blur', (e) => {this.validateTextBox("txtSurveyTitle","Request Title is mandatory") });
            this.domElement.querySelector('#txtWhoIsSurveyFor').addEventListener('blur', (e) => {this.validateTextBox("txtWhoIsSurveyFor","Who's for Survey is mandatory") });
            this.domElement.querySelector('#txtPurposeOfSurvey').addEventListener('blur', (e) => {this.validateTextBox("txtPurposeOfSurvey","Purpose Of Survey is mandatory") });
            
            this.domElement.querySelector('#ddlSurveyType').addEventListener('blur', (e) => {this.validateDropdown("ddlSurveyType","Survey Type is mandatory") });
            this.domElement.querySelector('#ddlDepartment').addEventListener('blur', (e) => {this.validateDropdown("ddlDepartment","Department is mandatory") });
        
            this.domElement.querySelector('#calSurveyStartDate').addEventListener('blur', (e) => {this.validateDate("calSurveyStartDate","Survey start date is mandatory") });
            this.domElement.querySelector('#calSurveyEndDate').addEventListener('blur', (e) => {this.validateDate("calSurveyEndDate","Survey end date is mandatory") });

            var todaydt = new Date();
            $("#calSurveyStartDate").datepicker({
                dateFormat: "dd/mm/yy",
                endDate: todaydt,
                changeMonth: true,
                changeYear: true,   
                minDate:todaydt,
                onSelect: function (date) {
                //Get selected date 
                    var date2 = $('#calSurveyStartDate').datepicker('getDate');
                    //sets minDate to txt_date_to
                    $('#calSurveyEndDate').datepicker('option', 'minDate', date2);
                   
                   this.focus();
                }
            });
            $('#calSurveyEndDate').datepicker({
                dateFormat: "dd/mm/yy",
                changeMonth: true,
                changeYear: true,
                onSelect:function(){
                    this.focus();
                }
            });
        
      this.LoadChoiceColumn("TypeOfSurvey","#ddlSurveyType");
      this.LoadChoiceColumn("DoYouRequireSurveyReport","#ddlRequireSurveyReport");
  }
  private LoadChoiceColumn(ChoiceColumnName,ControlName)
  {
    var control=document.getElementById(ControlName);
    sp.site.rootWeb.lists.getByTitle(Listname).fields.getByInternalNameOrTitle(ChoiceColumnName).get().then((fieldData)=>
    
    {
      if(fieldData['Choices'].length>0)
      {
        $(ControlName).append("<option value='0'>---Select---</option>");
        fieldData['Choices'].forEach(element => {
             $(ControlName).append("<option value=\"" +element+ "\">" +element + "</option>");
        });
        $(ControlName).val(requestFormItem[ChoiceColumnName]);
      }

    });
     
  }
  private LoadSocialMediaEditHtml()
  {
      var htmlRenderEdit=`
      <div class="row user-info" id="div_row_socialMediaForm_edit">
                    <h3 class="mb-4 col-12">SOCIAL MEDIA DETAILS</h3>
                    <div class="col-md-4 col-12 mb-4">
                        <p  >Request Title<span style="color:red">*</span></p>
                        <input type="text" class="form-input updateField" name="" id="txtSocialMediaTitle">
                        <span style="color:red"></span>
                    </div>
                    <div class="col-md-4 col-12 mb-4">
                        <p  >Social Media Type<span style="color:red">*</span></p>
                        <select multiple="multiple" style="height:60%" name="ddlSocialMediaType" class="form-input updateField" id="ddlSocialMediaType"></select>
                        <span style="color:red"></span>
                    </div>

                    <div class="col-md-4 col-12 mb-4">
                        <p>Does the department have a budget for this request?<span style="color:red">*</span></p>
                        <select name="ddlDeptHaveBudgetSocialMedia" class="form-input updateField" id="ddlDeptHaveBudgetSocialMedia"></select>
                        <span style="color:red"></span>
                    </div>
                    <div id="div_dur_spon" class="col-md-4 col-12 mb-4" style="display:none">
                        <p>Duration of Sponsored Ad</p>
                        <input type="text" class="form-input updateField" name="" id="txtDurationOfAd">
                    </div>
                    <div id="div_dt_inf_eng"  class="col-md-4 col-12 mb-4"  style="display:none">
                        <p>Date of influencer engagement  </p>
                        <input type="text" class="form-input updateField"  readonly="readonly" name="" id="calDateOfInfluencerEngmnt"  autocomplete="off">
                    </div>
                    
                    <div class="col-md-4 col-12 mb-4 hidecls">
                        <p>Date of Event</p>
                        <input type="text"  readonly="readonly" class="form-input updateField" name="" id="calDateOfEvent" autocomplete="off">
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
                        <p>Which department budget will this come out from ?<span style="color:red">*</span></p>
                        <input type="text" class="form-input updateField" name="" id="txtWhichDeptBudgetSocialMedia">
                        <span style="color:red"></span>
                    </div>

                    <div class="col-md-4 col-12 mb-4">
                        <p> Budget amount<span style="color:red">*</span></p>
                        <input type="text" class="form-input updateField" name="" id="txtBudgetAmtSocialMedia">
                        <span style="color:red"></span>
                    </div>
                    <div class="col-md-4 col-12 mb-4">
                        <p>Date of post<span style="color:red">*</span></p>
                        <input type="text" class="form-input updateField" name=""  readonly="readonly" id="calDateOfPost" autocomplete="off">
                        <span style="color:red"></span>
                    </div>
                    <div class="col-md-4 col-12 mb-4">
                        <p>Platforms<span style="color:red">*</span></p>
                        <select multiple="multiple" name="ddlPlatformsSocial" class="form-input updateField" id="ddlPlatformsSocial" style="height:70%"></select>
                        <span style="color:red"></span>
                    </div>
                    <div class="col-md-8 col-12 mb-4">
                        <p>Any additional field to add</p>
                        <textarea class="form-input updateField" style="height:auto!important" rows="5" cols="5" id="txtAdditionalDetails_Social"></textarea>
                    </div>
              
                    

        </div>      
      `;
      //$('#requestFormSection').append(html_SurveyEdit);
      $('#dvRequestForm').append(htmlRenderEdit);
      $('.hidecls').hide();

      this.domElement.querySelector('#txtSocialMediaTitle').addEventListener('blur', (e) => {this.validateTextBox("txtSocialMediaTitle","Request Title is mandatory") });
      this.domElement.querySelector('#txtWhichDeptBudgetSocialMedia').addEventListener('blur', (e) => {this.validateTextBox("txtWhichDeptBudgetSocialMedia","Which department budget will this come out from ? is mandatory") });
      this.domElement.querySelector('#txtBudgetAmtSocialMedia').addEventListener('blur', (e) => {this.validateTextBox("txtBudgetAmtSocialMedia","Budget amount is mandatory") });
      
      //txtWherePublish
      this.domElement.querySelector('#ddlDepartment').addEventListener('blur', (e) => {this.validateDropdown("ddlDepartment","Department is mandatory") });
      this.domElement.querySelector('#ddlDeptHaveBudgetSocialMedia').addEventListener('blur', (e) => {this.validateDropdown("ddlDeptHaveBudgetSocialMedia","Does the department have a budget for this request? is mandatory") });
       
      this.domElement.querySelector('#calDateOfPost').addEventListener('blur', (e) => {this.validateDate("calDateOfPost","Date of post is mandatory") });
  
      this.domElement.querySelector('#ddlSocialMediaType').addEventListener('blur', (e) => {this.ValMultiDropdown("ddlSocialMediaType","Socail media type is mandatory")});
      this.domElement.querySelector('#ddlPlatformsSocial').addEventListener('blur', (e) => {this.ValMultiDropdown("ddlPlatformsSocial","Platforms are mandatory")});



      this.LoadChoiceColumn("DoesTheDepartmentHaveaBudgetForT","#ddlDeptHaveBudgetSocialMedia");
      this.LoadChoiceColumn("SocialMediaType","#ddlSocialMediaType");  
      this.LoadChoiceColumn("Platforms","#ddlPlatformsSocial");
            $("#calDateOfPost").datepicker({
                dateFormat: "dd/mm/yy",
                endDate: todaydt,
                changeMonth: true,
                changeYear: true,   
                minDate:todaydt,
                onSelect: function ()
                {
                    this.focus();
                }
            });
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
                        
                        $("#calDateOfEvent").datepicker({
                            dateFormat: "dd/mm/yy",
                            endDate: todaydt,
                            changeMonth: true,
                            changeYear: true,   
                            minDate:todaydt,
                            onSelect: function ()
                            {
                                this.focus();
                            }
                        });
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
                         $("#calDateOfInfluencerEngmnt").datepicker({
                            dateFormat: "dd/mm/yy",
                            endDate: todaydt,
                            changeMonth: true,
                            changeYear: true,   
                            minDate:todaydt,
                            onSelect: function ()
                            {
                                this.focus();
                            }
                            
                        });
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
      <div class="row user-info" id="div_row_EventForm_edit">
    <h3 class="mb-4 col-12">EVENT DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Request Title<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtSocialMediaTitle">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Date & Time of Event <span style="color:red">*</span></p>
        <input type="text"  readonly="readonly" class="form-input updateField" name="" id="calEventDate">
        <span style="color:red"></span>
    </div>
    <div class="col-lg-2">
    <p>HH<span  style="color:red">*</span></p>
          <select name="incidenthours" id="ddlIncidentHours" class="form-input">
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
        <select name="incidentmins" id="ddlIncidentMins" class="form-input">
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
        <p>Duration of event<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtDurationOfEvent">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Location<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtLocationOfEvent">
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
        <p>Budget amount</p>
        <input type="text" class="form-input updateField" name="" id="txtBudgetAmtSocialMedia">
    </div>

    <div class="col-md-4 col-12 mb-4">
        <p>Type of Event<span style="color:red">*</span></p>
        <select name="ddlTypeofEvent" class="form-input updateField" id="ddlTypeofEvent"></select>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Requirements<span style="color:red">*</span></p>
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
            
            this.domElement.querySelector('#ddlDepartment').addEventListener('blur', (e) => {this.validateDropdown("ddlDepartment","Department is mandatory") });
            this.domElement.querySelector('#ddlDeptHaveBudgetSocialMedia').addEventListener('blur', (e) => {this.validateDropdown("ddlDeptHaveBudgetSocialMedia","Does the department have budget for this request is mandatory") });
            this.domElement.querySelector('#ddlTypeofEvent').addEventListener('blur', (e) => {this.validateDropdown("ddlTypeofEvent","Type of event is mandatory") });
            this.domElement.querySelector('#ddlIncidentHours').addEventListener('blur', (e) => {this.validateDropdown("ddlIncidentHours","Hours of event is mandatory") });
            this.domElement.querySelector('#ddlIncidentMins').addEventListener('blur', (e) => {this.validateDropdown("ddlIncidentMins","Minutes of event is mandatory") });

            this.domElement.querySelector('#calEventDate').addEventListener('blur', (e) => {this.validateDate("calEventDate","Date of event is mandatory") });

            this.domElement.querySelector('#ddlRequirements').addEventListener('blur', (e) => {this.ValMultiDropdown("ddlRequirements","Requirements are mandatory")});

      this.LoadChoiceColumn("Requirements","#ddlRequirements");
      this.LoadChoiceColumn("TypeOfEvent","#ddlTypeofEvent");
      this.LoadChoiceColumn("DoesTheDepartmentHaveaBudgetForT","#ddlDeptHaveBudgetSocialMedia");
      //calEventDate
      $("#calEventDate").datepicker({
        dateFormat: "dd/mm/yy",
        changeMonth: true,
        changeYear: true,   
        minDate:todaydt,
        onSelect: function ()
            {
                this.focus();
            }
        });
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
      <div class="row user-info" id="div_row_designProdForm_edit">
    <h3 class="mb-4 col-12">DESIGN AND PRODUCTION DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Request Title<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" id="txtSocialMediaTitle" />
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Type of Design <span style="color:red">*</span></p>
        <select id="ddlTypeofDesign" multiple="multiple" class="form-input updateField" style="height: 60% !important;"></select>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4" style="display:none" id="div_dec_element">
        <p>If decorative elements, please specify</p>
        <input type="text" class="form-input updateField" id="txtDecorativeElements" />
    </div>

    <div class="col-md-4 col-12 mb-4" style="display:none" id="div_dec_other">
        <p>If other collateral was selected, please specify</p>
        <input type="text" class="form-input updateField" id="txtOtherCollateral" />
    </div>

    <div class="col-md-4 col-12 mb-4" >
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
        <input class="form-input updateField" type="text"  readonly="readonly" id="calDateofDelivery" autocomplete="off"/>
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
            <p >Which department budget will this come out from?<span style="color:red">*</span></p>
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
        <input class="form-input updateField" type="text"  readonly="readonly" id="calInstallationDeadline" autocomplete="off" />
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
            
            this.domElement.querySelector('#ddlDepartment').addEventListener('blur', (e) => {this.validateDropdown("ddlDepartment","Department is mandatory") });
            this.domElement.querySelector('#ddlRequireProduction').addEventListener('blur', (e) => {this.validateDropdown("ddlRequireProduction","Will you require production ? is mandatory") });
            
            this.domElement.querySelector('#calDateofDelivery').addEventListener('blur', (e) => {this.validateDate("calDateofDelivery","Date of delivery is mandatory") });
            this.domElement.querySelector('#calInstallationDeadline').addEventListener('blur', (e) => {this.validateDate("calInstallationDeadline","Installation Deadline is mandatory") });
           
            this.domElement.querySelector('#ddlTypeofDesign').addEventListener('blur', (e) => {this.ValMultiDropdown("ddlTypeofDesign","Type of Design is mandatory")});


      this.LoadChoiceColumn("TypeOfDesign","#ddlTypeofDesign");
      this.LoadChoiceColumn("SupportingTextContentLanguage","#ddlSupportingText");
      this.LoadChoiceColumn("WillYouRequireProduction","#ddlRequireProduction");
      this.LoadChoiceColumn("DoesTheDepartmentHaveaBudgetForT","#ddlSocialMediaDeptHaveBudget");
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
    $("#calDateofDelivery").datepicker({
        dateFormat: "dd/mm/yy",
        endDate: todaydt,
        changeMonth: true,
        changeYear: true,   
        minDate:todaydt,
              onSelect: function ()
            {
                this.focus();
            }
        });
        $("#calInstallationDeadline").datepicker({
            dateFormat: "dd/mm/yy",
            endDate: todaydt,
            changeMonth: true,
            changeYear: true,   
            minDate:todaydt,
                  onSelect: function ()
            {
                this.focus();
            }
        });
  }
  private LoadNewsPaperAndMediaEditHtml()
  {
      var htmlRenderView=`
      <div class="row user-info" id="div_row_NewspaperForm_edit">
    <h3 class="mb-4 col-12">NEWSPAPER AND MEDIA DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Request Title<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtMediaTitle">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Does the department have a budget for this request?<span style="color:red">*</span></p>
        <select name="ddlDeptHaveBudgetNewsMedia" class="form-input updateField" id="ddlDeptHaveBudgetNewsMedia"></select>
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
        <input type="text"  readonly="readonly" class="form-input updateField" name="" id="calPublishDate" autocomplete="off">
    </div>
    <div class="col-md-6 col-12 mb-4">
        <p>Text Content <span style="color:red">*</span></p>
        <textarea class="form-input updateField" style="height:auto!important" rows="5" cols="5" id="txtTextContent"></textarea>
        <span style="color:red"></span>
    </div>
    <div class="col-md-6 col-12 mb-4">
        <p>Any additional details to add?</p>
        <textarea class="form-input updateField" style="height:auto!important" rows="5" cols="5" id="txtAdditionalDetails_Social"></textarea>
    </div>
    </div>
      `;
      //$('#requestFormSection').append(html_SurveyView);
      $('#dvRequestForm').append(htmlRenderView);
            this.domElement.querySelector('#ddlDepartment').addEventListener('blur', (e) => {this.validateDropdown("ddlDepartment","Department is mandatory") });
            this.domElement.querySelector('#ddlDeptHaveBudgetNewsMedia').addEventListener('blur', (e) => {this.validateDropdown("ddlDeptHaveBudgetNewsMedia","Does the department have a budget for this request ? is mandatory") });
            
            this.domElement.querySelector('#txtMediaTitle').addEventListener('blur', (e) => {this.validateTextBox("txtMediaTitle","Request Title is mandatory") });
            this.domElement.querySelector('#txtWhichDeptBudgetSocialMedia').addEventListener('blur', (e) => {this.validateTextBox("txtWhichDeptBudgetSocialMedia","Which department budget will this come out from ? is mandatory") });
            this.domElement.querySelector('#txtBudgetAmtSocialMedia').addEventListener('blur', (e) => {this.validateTextBox("txtBudgetAmtSocialMedia","Budget amount is mandatory") });
            this.domElement.querySelector('#txtTextContent').addEventListener('blur', (e) => {this.validateTextBox("txtTextContent","Text Content is mandatory") });
            
            
      this.LoadChoiceColumn("DoesTheDepartmentHaveaBudgetForT","#ddlDeptHaveBudgetNewsMedia");
      $("#calPublishDate").datepicker({
        dateFormat: "dd/mm/yy",
        endDate: todaydt,
        changeMonth: true,
        changeYear: true,   
        minDate:todaydt,
            onSelect: function ()
            {
                this.focus();
            }
        });
  }

  
  private LoadContentFormEditHtml()
  {
      var htmlRenderView=`
      <div class="row user-info" id="div_row_ContentCreation_edit">
    <h3 class="mb-4 col-12">CONTENT CREATION DETAILS</h3>
    <div class="col-md-4 col-12 mb-4">
        <p>Request Title<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtContentTitle">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Content type<span style="color:red">*</span></p>
        <select id="ddl_content_type" class="form-input updateField"></select>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Where will this content be published ?<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="txtWherePublished">
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Length of content </p>
        <input type="text" class="form-input updateField" name="" id="txtLengthofContent">
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Do you require bilingual content <span style="color:red">*</span> </p>
        <select id="ddlBilingualRequired" class="form-input updateField"></select>
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4">
        <p>Deadline for content creation<span style="color:red">*</span></p>
        <input type="text" class="form-input updateField" name="" id="calDeadline"  autocomplete="off" >
        <span style="color:red"></span>
    </div>
    <div class="col-md-4 col-12 mb-4" style="display:none" id="div_CC_lang_for_bilingual" >
        <p>If no, which language do you require the content in</p>
        <select class="form-input updateField" id="ddlSelectLanguage" multiple="multiple" style="height:68%"></select>
    </div>
     <div class="col-md-6 col-12 mb-4">
        <p>Please provide as much detail as possible about the content  requested<span style="color:red">*</span></p>
        <textarea class="form-input updateField" style="height:auto!important" rows="5" cols="5" id="txtMoreDetailsforContent"></textarea>
        <span style="color:red"></span>
    </div>
    <div class="col-md-6 col-12 mb-4">
        <p>Any additional details to add?</p>
        <textarea class="form-input updateField" style="height:auto!important" rows="5" cols="5" id="txtAdditionalDetails_Content"></textarea>
    </div>
  </div>
      `;
      //$('#requestFormSection').append(html_SurveyView);
      $('#dvRequestForm').append(htmlRenderView);

            this.domElement.querySelector('#txtContentTitle').addEventListener('blur', (e) => {this.validateTextBox("txtContentTitle","Request Title is mandatory") });
            this.domElement.querySelector('#txtWherePublished').addEventListener('blur', (e) => {this.validateTextBox("txtWherePublished","Where will this content be published ? is mandatory") });
            this.domElement.querySelector('#txtMoreDetailsforContent').addEventListener('blur', (e) => {this.validateTextBox("txtMoreDetailsforContent","Please provide as much detail as possible about the content requested is mandatory") });
            
            this.domElement.querySelector('#ddl_content_type').addEventListener('blur', (e) => {this.validateDropdown("ddl_content_type","Content Type is mandatory") });
            this.domElement.querySelector('#ddlDepartment').addEventListener('blur', (e) => {this.validateDropdown("ddlDepartment","Department is mandatory") });
            this.domElement.querySelector('#ddlBilingualRequired').addEventListener('blur', (e) => {this.validateDropdown("ddlBilingualRequired","Department is mandatory") });
  
            this.domElement.querySelector('#calDeadline').addEventListener('blur', (e) => {this.validateDate("calDeadline","Deadline for content creation is mandatory") });
          

      this.LoadChoiceColumn("DoYouRequireBilingualContent","#ddlBilingualRequired");
      this.LoadChoiceColumn("Language","#ddlSelectLanguage");
      this.LoadChoiceColumn("ContentTypeForm","#ddl_content_type");
      $("#calDeadline").datepicker({
        dateFormat: "dd/mm/yy",
        endDate: todaydt,
        changeMonth: true,
        changeYear: true,   
        minDate:todaydt,
            onSelect: function ()
            {
                this.focus();
            }
        });
        this.domElement.querySelector('#ddlBilingualRequired').addEventListener('change', (e) => {
            e.preventDefault();
            if($("#ddlBilingualRequired").val()=="No"){
                $("#div_CC_lang_for_bilingual").show();
            }
            else{
                $("#div_CC_lang_for_bilingual").hide();
            }
        });
  }

  private SubmitSurveyForm(){
    console.log();
    sp.site.rootWeb.lists.getByTitle(Listname).items.add({
        TECDepartmentId:parseInt(selected_dept),
        Title:sur_title,
        TypeOfSurvey: survey_type,
        PurposeOfSurvey:sur_purpose,
        WhoIsSurveyFor:sur_whos_for,
        SurveyStartDate:sur_start_date,
        SurveyEndDate:sur_end_date,                
        DoYouRequireSurveyReport:sur_required_report,
        AnyAdditionalDetails:$('#txtAdditionalDetails_Survey').val(),
        ContentTypeId:CT_surveyFormId,
       }).then(r=>{
        //this.BreakInheritance(r);
        /*  r.item.breakRoleInheritance(false).then(permission => {
            for (var i = 0; i < GroupsforRolesAssignments.length; i++) {
                if(GroupsforRolesAssignments[i].id==29){
                   //r.item.roleAssignments.add(GroupsforRolesAssignments[i].id, 1073741829);// assigning full control to ADMIN
                   this.AddUserToSitePermission(r,GroupsforRolesAssignments[i].id,1073741829);
                }
                else{
                 //r.item.roleAssignments.add(GroupsforRolesAssignments[i].id, 1073741827); // assigning contribute to 
                 this.AddUserToSitePermission(r,GroupsforRolesAssignments[i].id,1073741827);
                }
            }
            r.item.roleAssignments.add(r.data.AuthorId, 1073741827).then(p =>{ // adding contribute permission
            r.item.roleAssignments.remove(r.data.AuthorId,1073741829); // deleting full permission created by user
            });
        }).catch(function(err){
            console.log(err);  
        });
         */
        alert("Thank you. The request was submitted successfully.");
        window.location.href=this.context.pageContext.site.absoluteUrl+"/Pages/TecPages/SearchRF.aspx";
      }).catch(function(err) {  
        console.log(err);  
      });
  }
    private async AddUserToSitePermission(r,userid,roledefid) {
       await r.item.roleAssignments.add(userid, roledefid);
           
        // sp.web.roleAssignments.add(userid,roledefid).then(permission => {
        //     console.log(permission);
        //     // Append the result to your html
        // });
    }
  private validateSurvey(){
       var result=true;
       sur_title=document.getElementById('txtSurveyTitle')["value"];
       selected_dept=document.getElementById('ddlDepartment')["value"];
       survey_type= document.getElementById('ddlSurveyType')["value"];
       sur_start_date= $('#calSurveyStartDate').datepicker('getDate');
      
       sur_end_date=$('#calSurveyEndDate').datepicker('getDate');
       sur_whos_for=document.getElementById('txtWhoIsSurveyFor')["value"];
       sur_required_report=document.getElementById('ddlRequireSurveyReport')["value"];
       sur_purpose=document.getElementById('txtPurposeOfSurvey')["value"];
       if(selected_dept=="0"){
        $("#ddlDepartment").next("span").text("Department is mandatory");
            result=false; 
        }
        else{
            $("#ddlDepartment").next("span").text(" ");
        }
       if(survey_type=="0"){
         $("#ddlSurveyType").next("span").text("Survey Type is mandatory");
         result=false; 
        }
       else{
          $("#ddlSurveyType").next("span").text(" ");
        }
        if(sur_start_date==null){
            $("#calSurveyStartDate").next("span").text("Survey start date is mandatory");
         result=false; 
        }
        else{
            $("#calSurveyStartDate").next("span").text(" ");
        }
        if(sur_end_date==null){
            $("#calSurveyEndDate").next("span").text("Survey end date is mandatory");
         result=false; 
        }
        else{
            $("#calSurveyEndDate").next("span").text(" ");
        }
        if(sur_title==""){
            $("#txtSurveyTitle").next("span").text("Survey title is mandatory");
         result=false; 
        }
        else{
            $("#txtSurveyTitle").next("span").text(" ");
        }
        if(sur_whos_for==""){
            $("#txtWhoIsSurveyFor").next("span").text("Who's for Survey is mandatory");
         result=false; 
        }
        else{
            $("#txtWhoIsSurveyFor").next("span").text(" ");
        }
        if(sur_purpose==""){
            $("#txtPurposeOfSurvey").next("span").text("Purpose of survey is mandatory");
         result=false; 
        }
        else{
            $("#txtPurposeOfSurvey").next("span").text(" ");
        }
     return result;
  }
  private ValidateSocialMediaFields()
  {
    var isValid = true;
    
        if($('#txtSocialMediaTitle').val()=="" || $('#txtSocialMediaTitle').val()==undefined)
        {
            $("#txtSocialMediaTitle").next("span").text("Request title is mandatory");
            isValid = false;
        }
        else
        {
            $("#txtSocialMediaTitle").next("span").text("");
        }
        if($("#ddlDepartment").val()=="0"){
            $("#ddlDepartment").next("span").text("Department is mandatory");
            isValid=false; 
        }
        else{
            $("#ddlDepartment").next("span").text(" ");
        } 

        if($('#txtWhichDeptBudgetSocialMedia').val()=="")
        {
            $("#txtWhichDeptBudgetSocialMedia").next("span").text("Which department budget will this come out from ? is mandatory");
            isValid = false;
        }
        else
        {
            $("#txtWhichDeptBudgetSocialMedia").next("span").text();
        }


        if($('#txtBudgetAmtSocialMedia').val()=="")
        {
            $("#txtBudgetAmtSocialMedia").next("span").text("Budget amount is mandatory");
            isValid = false;
        }
        else
        {
            $("#txtBudgetAmtSocialMedia").next("span").text("");
        }

        if($('#ddlPlatformsSocial option:selected').length ==0 || $('#ddlPlatformsSocial option:selected')[0].innerText=="---Select---")
        {
            $("#ddlPlatformsSocial").next("span").text("Platforms are mandatory");
            isValid = false;
        }
        else
        {
            $("#ddlPlatformsSocial").next("span").text("");//$("#ddlSocialMediaType :selected")[i].innerText=="Event Coverage"
        }
        if($('#ddlSocialMediaType option:selected').length ==0 || $("#ddlSocialMediaType :selected")[0].innerText=="---Select---")
        {
            $("#ddlSocialMediaType").next("span").text("Social media type is mandatory");
            isValid = false;
        }
        else
        {
            $("#ddlSocialMediaType").next("span").text("");
        }

        var dateVal = $('#calDateOfPost').datepicker('getDate');
        if (dateVal == null || dateVal == undefined) {
            $("#calDateOfPost").next("span").text("Date of post is mandatory");
            isValid = false;
        }
        else {
            $("#calDateOfPost").next("span").text("");
        }
        
        
         if($('#ddlDeptHaveBudgetSocialMedia').val()=="0")
         {
             $("#ddlDeptHaveBudgetSocialMedia").next("span").text("Does the department have a budget for this request? is mandatory");
             isValid = false;
         }    
         else
         {
             $("#ddlDeptHaveBudgetSocialMedia").next("span").text("");
         }
        return isValid;
  }
  private ValidateContentCreationFrom(){
    var result=true;
    content_title=document.getElementById('txtContentTitle')["value"];
    selected_dept=document.getElementById('ddlDepartment')["value"];
    content_type= document.getElementById('ddl_content_type')["value"];
    content_where_published=document.getElementById('txtWherePublished')["value"];
    content_deadline_date=$('#calDeadline').datepicker('getDate');
    content_require_biligual=document.getElementById('ddlBilingualRequired')["value"];
    content__asap_details=document.getElementById('txtMoreDetailsforContent')["value"];

    content_additional_details=document.getElementById('txtAdditionalDetails_Content')["value"];
    content_sel_bilingual_lang=$('#ddlSelectLanguage').val();
    content_length=document.getElementById('txtLengthofContent')["value"];
    
    if(content_title==""){
    $("#txtContentTitle").next("span").text("Request Title is mandatory");
        result=false; 
    }
    else{
        $("#txtContentTitle").next("span").text(" ");
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
  private SubmitContentCreationItem(){
    sp.site.rootWeb.lists.getByTitle(Listname).items.add({
        TECDepartmentId:parseInt(selected_dept),
        Title:content_title,
        ContentTypeForm:content_type,
        WhereWillThisBePublished:content_where_published,
        PleaseProvideAllDetailsForConten:content__asap_details,
        LengthOfContent: content_length!=null?content_length:"",
        DoYouRequireBilingualContent: content_require_biligual,
        Language:{ results:content_sel_bilingual_lang},
        DeadLine_x0020_Content_x0020_Cre:content_deadline_date,          
        AnyAdditionalDetails:content_additional_details!=""?content_additional_details:"",
        ContentTypeId:CT_ContentCreationFormId,
      }).then(r=>{

        alert("Thank you. The request was submitted successfully.");
        window.location.href=this.context.pageContext.site.absoluteUrl+"/Pages/TecPages/SearchRF.aspx";

      }).catch(function(err) {  
        console.log(err);  
      });
  }
  private validateMediaNewsPaperForm(){
      var mediaresult=true;
      selected_dept=document.getElementById('ddlDepartment')["value"];
      media_title=document.getElementById('txtMediaTitle')["value"];
      media_dept_have_budget_for_this=document.getElementById('ddlDeptHaveBudgetNewsMedia')["value"];
      media_which_dept_come_out_from= document.getElementById('txtWhichDeptBudgetSocialMedia')["value"];
      media_budget_amount=document.getElementById('txtBudgetAmtSocialMedia')["value"];
      media_new_preferences=document.getElementById('txtMediaPreferences')["value"]
      media_publish_date=$('#calPublishDate').datepicker('getDate');;
      media_text_content=document.getElementById('txtTextContent')["value"]; 
      media_any_additional_info=document.getElementById('txtAdditionalDetails_Social')["value"]; 

      if(media_title==""){
        $("#txtMediaTitle").next("span").text("Request Title is mandatory");
        mediaresult=false; 
        }
        else{
            $("#txtMediaTitle").next("span").text(" ");
        } 
        if(selected_dept=="0"){
         $("#ddlDepartment").next("span").text("Department is mandatory");
         mediaresult=false; 
         }
         else{
             $("#ddlDepartment").next("span").text(" ");
         } 
        if(media_dept_have_budget_for_this=="0"){
            $("#ddlDeptHaveBudgetNewsMedia").next("span").text("Does the department have a budget for this request ? is mandatory");
            mediaresult=false; 
        }
        else{
            $("#ddlDeptHaveBudgetNewsMedia").next("span").text(" ");
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
  private SubmitMediaNewsPaperCTItem(){
    sp.site.rootWeb.lists.getByTitle(Listname).items.add({
        TECDepartmentId:parseInt(selected_dept),
        Title:media_title,
        AnyAdditionalDetails:media_any_additional_info,
        DoesTheDepartmentHaveaBudgetForT:media_dept_have_budget_for_this,
        WhichDepartmentBudgetWillThisCom:media_which_dept_come_out_from,
        BudgetAmount:media_budget_amount,
        TextContent:media_text_content,
        NewspaperMediaPlatformPrefrences:media_new_preferences,
        PublishDate:media_publish_date,
        ContentTypeId:CT_MediaRequestFormId,
        }).then(r=>{
            alert("Thank you. The request was submitted successfully.");
            window.location.href=this.context.pageContext.site.absoluteUrl+"/Pages/TecPages/SearchRF.aspx";
        }).catch(function(err) {  
        console.log(err);  
      });
  }

  private validatePhotoVideoForm(){
    var photoresult=true;
    selected_dept=document.getElementById('ddlDepartment')["value"];
    pho_vid_title=document.getElementById('txt_photo_req_title')["value"];
    pho_vid_type_ofshoot=document.getElementById('ddl_shootType')["value"];
    pho_vid_does_dept_have_bud= document.getElementById('ddlDeptHaveBudgetSocialMedia')["value"];
    pho_bud_amt=document.getElementById('txtBudgetAmt')["value"];
 
    pho_vid_shoot_date_to=$('#calDateOfShootTo').datepicker('getDate');
    pho_vid_shoot_date_from=$('#calDateOfShootFrom').datepicker('getDate');;
    pho_vid_which_dept_bud_for_this=document.getElementById('txtBudgetDept')["value"]; 
    pho_vid_shoot_loc=document.getElementById('txtLocation')["value"]; 
    pho_vid_where_will_publish=document.getElementById('txtWherePublish')["value"]

    pho_vid_is_cast_req=document.getElementById('ddlIsCastRequired')["value"]; 
    pho_vid_style_of_shoot=$("#ddlStyleOfShoot").val();//document.getElementById('ddlStyleOfShoot')["value"]; 
    pho_vid_pur_of_shoot=document.getElementById('txtPurposeOfShoot')["value"];

    pho_vid_add_info=(document.getElementById('txtAdditionalDetails_Video')["value"])!=""?(document.getElementById('txtAdditionalDetails_Video')["value"]):"";
    if(pho_vid_title==""){
      $("#txt_photo_req_title").next("span").text("Request Title is mandatory");
      photoresult=false; 
      }
      else{
          $("#txt_photo_req_title").next("span").text(" ");
      } 
      if(selected_dept=="0"){
       $("#ddlDepartment").next("span").text("Department is mandatory");
       photoresult=false; 
       }
       else{
           $("#ddlDepartment").next("span").text(" ");
       } 
      if(pho_vid_type_ofshoot=="0"){
          $("#ddl_shootType").next("span").text("Type of Shoot is mandatory");
          photoresult=false; 
      }
      else{
          $("#ddl_shootType").next("span").text(" ");
      }
      if(pho_vid_does_dept_have_bud=="0"){
          $("#ddlDeptHaveBudgetSocialMedia").next("span").text("Does the department have a budget for this request? is mandatory");
          photoresult=false; 
          }
          else{
              $("#ddlDeptHaveBudgetSocialMedia").next("span").text(" ");
      }
      if(pho_bud_amt==""){
          $("#txtBudgetAmt").next("span").text("Budget amount is mandatory");
          photoresult=false; 
      }
      else{
          $("#txtBudgetAmt").next("span").text(" ");
      }
      if(pho_vid_shoot_date_to==null){
          $("#calDateOfShootTo").next("span").text("Date of Shoot (To) is mandatory");
          photoresult=false; 
      }
      else{
          $("#calDateOfShootTo").next("span").text(" ");
      }
      if(pho_vid_shoot_date_from==null){
        $("#calDateOfShootFrom").next("span").text("Date of Shoot (From) is mandatory");
        photoresult=false; 
        }
        else{
            $("#calDateOfShootFrom").next("span").text(" ");
        }
        if(pho_vid_which_dept_bud_for_this==""){
            $("#txtBudgetDept").next("span").text("Which department budget will this come out from ? is mandatory");
            photoresult=false; 
        }
        else{
            $("#txtBudgetDept").next("span").text(" ");
        }
        if(pho_vid_shoot_loc==""){
          $("#txtLocation").next("span").text("Location is mandatory");
          photoresult=false; 
          }
          else{
              $("#txtLocation").next("span").text(" ");
          }

          if(pho_vid_where_will_publish==""){
            $("#txtWherePublish").next("span").text("Where will this be published? is mandatory");
            photoresult=false; 
            }
            else{
                $("#txtWherePublish").next("span").text(" ");
            }

            if(pho_vid_is_cast_req=="0"){
            $("#ddlIsCastRequired").next("span").text("Is a cast required ? is mandatory");
            photoresult=false; 
            }
            else{
                $("#ddlIsCastRequired").next("span").text(" ");
            }
            if(pho_vid_style_of_shoot==""){
            $("#ddlStyleOfShoot").next("span").text("Style of shoot is mandatory");
            photoresult=false; 
            }
            else{
                $("#ddlStyleOfShoot").next("span").text(" ");
            }
            if(pho_vid_pur_of_shoot==""){
            $("#txtPurposeOfShoot").next("span").text("Purpose Of Shoot is mandatory");
            photoresult=false; 
            }
            else{
                $("#txtPurposeOfShoot").next("span").text(" ");
            }
      return photoresult;
  }

    private SubmitPhotoCreationItem(){
       // alert(sp.site.rootWeb.toUrl());
            sp.site.rootWeb.lists.getByTitle(Listname).items.add({
                TECDepartmentId:parseInt(selected_dept),
                Title:pho_vid_title,
                
                DoesTheDepartmentHaveaBudgetForT:pho_vid_does_dept_have_bud,
                
                WhichDepartmentBudgetWillThisCom:pho_vid_which_dept_bud_for_this,
                BudgetAmount:pho_bud_amt,
                TypeOfShoot :pho_vid_type_ofshoot,
                PurposeOfShoot:pho_vid_pur_of_shoot,
                DateOfShoot:pho_vid_shoot_date_from,
                DateofShootTo:pho_vid_shoot_date_to,
                Location:pho_vid_shoot_loc,
                StyleOfShoot:{ results:pho_vid_style_of_shoot},
                WhereWillThisBePublished:pho_vid_where_will_publish,
                IsAcastRequired:pho_vid_is_cast_req,
                AnyAdditionalDetails:pho_vid_add_info,
                ContentTypeId:CT_photovideoId
                }).then(r=>{
    
                    alert("Thank you. The request was submitted successfully.");
                    window.location.href=this.context.pageContext.site.absoluteUrl+"/Pages/TecPages/SearchRF.aspx";
    
                }).catch(function(err) {  
                  console.log(err);  
                });
          // }
          // else{
          //   //$("#lbl_Innovate_first_comm_err").text(arrLang[lang]['SuggestionBox']['CommentMandatory']);
          // }
  }

  private ValidateDesignProdFrom()
  {
    var designResult=true;
    selected_dept=document.getElementById('ddlDepartment')["value"];
    desi_title=document.getElementById('txtSocialMediaTitle')["value"];
    desi_type_of_design=$('#ddlTypeofDesign').val();//document.getElementById('ddlTypeofDesign')["value"];
    desing_size= document.getElementById('txtSize')["value"];
    design_support_content=document.getElementById('ddlSupportingText')["value"];
 
    desi_date_delivery=$('#calDateofDelivery').datepicker('getDate');
    desi_installa_deadline=$('#calInstallationDeadline').datepicker('getDate');;
    desi_illust_ref=document.getElementById('txtIllustrationReference')["value"]; 
    desi_will_req_prod=document.getElementById('ddlRequireProduction')["value"]; 
    desi_does_dept_have_budget=document.getElementById('ddlSocialMediaDeptHaveBudget')["value"]

    design_which_dept_budget_come=document.getElementById('txtSocialMediaDeptBudgetWillCome')["value"]; 
    desi_bud_amt=document.getElementById('txtSocialMediaBudgetAmount')["value"]; 
    desi_quantity=document.getElementById('txtQuantity')["value"];

    desi_loc=document.getElementById('txtLocation')["value"]; 
    desi_preff_mater=document.getElementById('txtPreferredMaterial')["value"]; 
    desi_add_info=document.getElementById('txtAdditionalDetails_SocialMedia')["value"];
    if(desi_title==""){
        $("#txtSocialMediaTitle").next("span").text("Request Title is mandatory");
        designResult=false; 
        }
        else{
            $("#txtSocialMediaTitle").next("span").text(" ");
        } 
        if(selected_dept=="0"){
         $("#ddlDepartment").next("span").text("Department is mandatory");
         designResult=false; 
         }
         else{
             $("#ddlDepartment").next("span").text(" ");
         } 
        if(desi_type_of_design==""){
            $("#ddlTypeofDesign").next("span").text("Type of Design is mandatory");
            designResult=false; 
        }
        else{
            $("#ddlTypeofDesign").next("span").text(" ");
        }

        if(desi_date_delivery==null){
         $("#calDateofDelivery").next("span").text("Date of Delivery is mandatory");
         designResult=false; 
         }
         else{
             $("#calDateofDelivery").next("span").text(" ");
         } 
        if(desi_will_req_prod=="0"){
            $("#ddlRequireProduction").next("span").text("Will you require production ? is mandatory");
            designResult=false; 
        }
        else{
            $("#ddlRequireProduction").next("span").text(" ");
        }

        if(desi_date_delivery==null){
            $("#calDateofDelivery").next("span").text("Date of Delivery is mandatory");
            designResult=false; 
            }
            else{
                $("#calDateofDelivery").next("span").text(" ");
            } 
           if(desi_will_req_prod=="0"){
               $("#ddlRequireProduction").next("span").text("Will you require production ? is mandatory");
               designResult=false; 
           }
           else{
               $("#ddlRequireProduction").next("span").text(" ");
           }
           if(desi_will_req_prod=="Yes" && desi_does_dept_have_budget=="0"){
            $("#ddlSocialMediaDeptHaveBudget").next("span").text("Does the department have a budget for this request is mandatory");
            designResult=false; 
           }
           else{
            $("#ddlSocialMediaDeptHaveBudget").next("span").text(" ");
           }
           if(desi_will_req_prod=="Yes" && design_which_dept_budget_come==""){
            $("#txtSocialMediaDeptBudgetWillCome").next("span").text("Which department budget will this come out from is mandatory");
            designResult=false; 
           }
           else{
            $("#txtSocialMediaDeptBudgetWillCome").next("span").text(" ");
           }

           if(desi_bud_amt==""){
            $("#txtSocialMediaBudgetAmount").next("span").text("Budget amount is mandatory");
            designResult=false; 
           }
           else{
            $("#txtSocialMediaBudgetAmount").next("span").text(" ");
           }

           if(desi_bud_amt==""){
            $("#txtSocialMediaBudgetAmount").next("span").text("Budget amount is mandatory");
            designResult=false; 
           }
           else{
            $("#txtSocialMediaBudgetAmount").next("span").text(" ");
           }
        
           if(!isNaN(Number($("#txtSocialMediaBudgetAmount").val())))
         {
          $("#txtSocialMediaBudgetAmount").next("span").text("Budget amount must be number");
          designResult=false; 
         }

           if(desi_quantity==""){
            $("#txtQuantity").next("span").text("Quantity is mandatory");
            designResult=false; 
           }
           else{
            $("#txtQuantity").next("span").text(" ");
           }
           if(desi_installa_deadline==null){
            $("#calInstallationDeadline").next("span").text("Installation Deadline is mandatory");
            designResult=false; 
           }
           else{
            $("#calInstallationDeadline").next("span").text(" ");
           }
    return designResult;
  }
  private SubmitDesignProdItem(){
    sp.site.rootWeb.lists.getByTitle(Listname).items.add({
        TECDepartmentId:parseInt(selected_dept),
        Title:desi_title,
        TypeOfDesign:{ results: desi_type_of_design },
        SpecifyDecorativeElements:$('#txtDecorativeElements').val(),
        SpecifyCollateral:$('#txtOtherCollateral').val(),
        Size:desing_size,
        SupportingTextContentLanguage:design_support_content,
        IllustrationReference:desi_illust_ref,
        DateOfDelivery:desi_date_delivery,
        WillYouRequireProduction:desi_will_req_prod,
        DoesTheDepartmentHaveaBudgetForT:desi_does_dept_have_budget,
        WhichDepartmentBudgetWillThisCom:design_which_dept_budget_come,              
        BudgetAmount:desi_bud_amt,
        Quantity:desi_quantity,
        Location:desi_loc,
        InstallationDeadline:desi_installa_deadline,        
        AnyAdditionalDetails:desi_add_info,
        PrefferedMaterial:desi_preff_mater,
        ContentTypeId:CT_designAndProductionFormId

      }).then(r=>{

        alert("Thank you. The request was submitted successfully.");
        window.location.href=this.context.pageContext.site.absoluteUrl+"/Pages/TecPages/SearchRF.aspx";

      }).catch(function(err) {  
        console.log(err);  
      });
  }
  private validateEventsForm() {
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

        if($('#ddlDepartment').val()=="0")
        {
            $("#ddlDepartment").next("span").text("Department is mandatory");
            resultevent = false;
        }
        else
        {
            $("#ddlDepartment").next("span").text();
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

         if($('#ddlTypeofEvent').val()=="0")
         {
             $("#ddlTypeofEvent").next("span").text("Type of event is mandatory");
             resultevent = false;
         }    
         else
         {
             $("#ddlTypeofEvent").next("span").text("");
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
  private  SubmitEventItem() {
    var eventTitle=$('#txtSocialMediaTitle').val();
    var  deptIDstr=$('#ddlDepartment').val().toString();
    //var socialMediaTypeVal=$('#ddlSocialMediaType').val();
    var requirementVals=$('#ddlRequirements').val();
    //var postDate=$('#calDateOfPost').datepicker('getDate');
    var eventDate=$('#calEventDate').datepicker('getDate');
    //var influencerDate=$('#calDateOfInfluencerEngmnt').datepicker('getDate');
            sp.site.rootWeb.lists.getByTitle(Listname).items.add({
              TECDepartmentId:parseInt(deptIDstr),
              Title:eventTitle,
              EventDateTime:eventDate,
              DoesTheDepartmentHaveaBudgetForT:$('#ddlDeptHaveBudgetSocialMedia').val(),
              WhichDepartmentBudgetWillThisCom:$('#txtWhichDeptBudgetSocialMedia').val(),
              BudgetAmount:$('#txtBudgetAmtSocialMedia').val(),
              TimeOfEvent:$('#ddlIncidentHours').val()+":"+$('#ddlIncidentMins').val(),
              EventDuration:$('#txtDurationOfEvent').val(),
              Location:$('#txtLocationOfEvent').val(),
              TypeOfEvent:$('#ddlTypeofEvent').val(),
              Requirements:{ results: requirementVals },
              IfDecorativePleaseSpecify:$('#txtDecorativeElements').val(),
              If_x0020_Other_x0020_Please_x002:$('#txtOthers').val(),
              AnyAdditionalDetails:$('#txtAdditionalDetails_Social').val(),
              ContentTypeId:CT_EventFormId
            }).then(r=>{

                alert("Thank you. The request was submitted successfully.");
                window.location.href=this.context.pageContext.site.absoluteUrl+"/Pages/TecPages/SearchRF.aspx";

            }).catch(function(err) {  
              console.log(err);  
            });
  }
   private BreakInheritance(r){
    r.item.breakRoleInheritance(false);  
    for (var i = 0; i < GroupsforRolesAssignments.length; i++) {
       if(GroupsforRolesAssignments[i].id!=29){
        r.item.roleAssignments.add(GroupsforRolesAssignments[i].id, 1073741827); // assigning contribute permission to other groups
       }
       else{
        r.item.roleAssignments.add(GroupsforRolesAssignments[i].id, 1073741829);// assigning full control to ADMIN
       }
    }
    r.item.roleAssignments.add(r.data.AuthorId, 1073741827); // adding contribute permission to  created by
    r.item.roleAssignments.remove(r.data.AuthorId,1073741829).then(permission => {
        console.log(permission);
        alert("Thank you. The request was submitted successfully.");
        window.location.href=this.context.pageContext.site.absoluteUrl+"/Pages/TecPages/SearchRF.aspx";
    }); // deleting full permission to created by user 
   }
}
