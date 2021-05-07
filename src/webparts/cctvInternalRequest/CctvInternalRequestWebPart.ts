import { Version } from '@microsoft/sp-core-library';
import { sp } from "@pnp/sp/presets/all";
import { SPComponentLoader } from '@microsoft/sp-loader';
import 'jqueryui';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
//import { Common } from './../Helpers/Common.js';

import styles from './CctvInternalRequestWebPart.module.scss';
import * as strings from 'CctvInternalRequestWebPartStrings';
import * as $ from 'jquery';
export interface ICctvInternalRequestWebPartProps {
  description: string;
}
//require('Common.js');
//require('sppeoplepicker');

import * as moment from 'moment';
import { getThemedContext } from 'office-ui-fabric-react';

declare var arrLang: any;
declare var lang: any;
var AssignedITGroupID:any;
let groups: any[] = [];
var IsLegalTeamMember:number;
var LegalManagerGroupID:any;
export default class CctvInternalRequestWebPart extends BaseClientSideWebPart<ICctvInternalRequestWebPartProps> {
  private AssignedToGroupITManager:string='LegalManager';
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
  private Listname: string = "CCTV_Internal_Incident";
  private EmpDirectoryListName:string="EmployeeDirectory";
  
  private DeptListname:string="LK_Departments";
  private LogsListname: string = "WorkflowLogs";
  private ITMgrActionUrl:string="/Pages/TecPages/cctv/ITManagerAction.aspx?ItemID=";
  private ITManagerAction: string='IT Manager Action';
  public constructor() {
    super();
    SPComponentLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
  }

  public render(): void {
    var lcid= this.context.pageContext.legacyPageContext['currentCultureLCID'];  
    lang=lcid==13313?"ar":"en";
   
    this.domElement.innerHTML = `
    <section class="inner-page-cont" style="margin-top: -60px">

      <div class="container-fluid mt-5">
        
            <div class="row user-info col-md-10 mx-auto col-12 pl-0 pr-0 m-0">
                <h3 class="mb-4 col-12">Employee Details</h3>
                      <div class="col-md-4 col-12 mb-4">
                        <p>Full Name<span  style="color:red">*</span></p>
                        <input type="text" id="txtEmpName" class="form-input" name="txtEmpName" disabled>
                        <label id="lbl_emp_name" class="form-label" style="color:red"></label>
                      </div>
                      <div class="col-md-4 col-12 mb-4">
                        <p>Employee ID<span  style="color:red">*</span></p>
                        <input type="text" id="txtEmpID" class="form-input" name="txtEmpID" disabled>
                        <label id="lbl_emp_id" class="form-label" style="color:red"></label>
                      </div>
                      <div class="col-md-4 col-12 mb-4">
                        <p>Department<span  style="color:red">*</span></p>
                        <select name="department" id="sel_Dept" class="form-input" disabled></select>
                        <label id="lbl_emp_dept" class="form-label" style="color:red"></label>
                      </div>
                      <div class="col-md-4 col-12 mb-4">
                          <p>Email<span  style="color:red">*</span></p>
                          <input type="text" id="txtEmpMail" class="form-input" name="txtEmpMail">
                          <label id="lbl_emp_mail" class="form-label" style="color:red"></label>
                      </div>
                      <div class="col-md-4 col-12 mb-4">
                          <p>Mobile/Telephone Number<span  style="color:red">*</span></p>
                          <input type="text" id="txtEmpMobile" class="form-input" name="txtEmpMobile" disabled>
                          <label id="lbl_emp_mobile" class="form-label" style="color:red"></label>
                      </div>
                      <h3 class="mb-4 col-12">Incident Details</h3>
                      <div class="col-md-4 col-12 mb-4">
                          <p>Incident Date & Time (From)<span  style="color:red">*</span></p>
                          <input type="text" autocomplete="off" id="txtIncidentDate" class="form-input" name="txtIncidentDate" aria-disabled="true">
                          <label id="lbl_Inc_date" class="form-label" style="color:red"></label>
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
                                <label id="lbl_inc_hours" class="form-label" style="color:red"></label>
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
                            <label id="lbl_inc_mins" class="form-label" style="color:red"></label>
                          </div>
                      <div class="col-md-4 col-12 mb-4">
                      </div>
                      <div class="col-md-4 col-12 mb-4">
                        <p>Incident Date & Time (To)<span  style="color:red">*</span></p>
                        <input type="text" autocomplete="off" id="txtIncidentDate_to" class="form-input" name="txtIncidentDate_to" aria-disabled="true">
                        <label id="lbl_Inc_date_to" class="form-label" style="color:red"></label>
                      </div>
                      <div class="col-lg-2">
                          <p>HH<span  style="color:red">*</span></p>
                                <select name="incidenthours" id="ddlIncidentHours_to" class="form-input">
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
                                <label id="lbl_inc_hours_to" class="form-label" style="color:red"></label>
                            </div>
                            <div class="col-lg-2"> 
                            <p>MM<span  style="color:red">*</span></p>
                            <select name="incidentmins_to" id="ddlIncidentMins_to" class="form-input">
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
                          <label id="lbl_inc_mins_to" class="form-label" style="color:red"></label>
                        </div>
                          <div class="col-md-12 col-12 mb-4">
                              <p>Reason for Request<span  style="color:red">*</span></p>
                              <textarea style="height:auto !important" rows="5" cols="5" id="txtReqReason" class="form-input" name="txtReqReason"></textarea>
                              <label id="lbl_inc_reason" class="form-label" style="color:red"></label>
                          </div>
                  </div>
              
          </div>     
          <div class="container-fluid mt-5"  id="div_row_buttons">
            <div class="col-md-10 mx-auto col-12">
                <div class="row">
                  <div class=" col-12 btnright">
                      <button class="red-btn red-btn-effect shadow-sm mt-4" id="btnSubmit"> <span>Submit</span></button>
                      <button class="red-btn red-btn-effect shadow-sm mt-4" id="btnCancle"> <span>Cancel</span></button>
                  </div>
                </div>
            </div>
        </div> 

      </section>
      `;
      this.PageLoad();
      this.setButtonsEventHandlers();
      this._checkUserInGroup();
      this.GetITMgrGroupID("ITManager");
      
  }
  
  private PageLoad():void{
    var todaydt = new Date();
     var momentobj=moment(todaydt);
     var fomattedate=momentobj.format("DD-MM-YYYY");


    $("#txtIncidentDate").datepicker({
        dateFormat: "dd/mm/yy",
        endDate: todaydt,
        changeMonth: true,
        changeYear: true,   
        maxDate:0,
        onSelect: function (date) {
        //Get selected date 
            var date2 = $('#txtIncidentDate').datepicker('getDate');
            //sets minDate to txt_date_to
            $('#txtIncidentDate_to').datepicker('option', 'minDate', date2);
           // $('#calSurveyStartDate').prop("disabled",true);
           this.focus();
        }
    });
    $('#txtIncidentDate_to').datepicker({
        dateFormat: "dd/mm/yy",
        changeMonth: true,
        changeYear: true,
        maxDate:0,
        onSelect:function(){
           // $('#calSurveyEndDate').prop("disabled",true);
           this.focus();
        }
    });

    this.LoadDepartments();
    
  }
  private  GetITMgrGroupID(groupname:string)
  {
    sp.site.rootWeb.siteGroups.getByName(groupname).get().then(function(result) {  
        AssignedITGroupID=result.Id;
      }).catch(function(err) {  
      console.log(err);  
    });  
  
  }
  private validateDropdownToMins()
  {
    var ddlprehrs=$('#ddlIncidentHours').val();
    var ddlcurhrs=$('#ddlIncidentHours_to').val(); 
    var date1 = $('#txtIncidentDate').datepicker('getDate');
    var date2 = $('#txtIncidentDate_to').datepicker('getDate');
    var minsFrom=$('#ddlIncidentMins').val();
    var minsTo=$('#ddlIncidentMins_to').val();
    var momentobj1=moment(date1); 
    var momentobj2=moment(date2);
    if( minsTo=="MM")
    {
      $("#lbl_inc_mins_to").text("Mins are mandatory");
    }
    else if(momentobj1.format("dd-mm-yyyy")==momentobj2.format("dd-mm-yyyy") && ddlprehrs==ddlcurhrs && minsFrom>minsTo){
      $("#lbl_inc_mins_to").text("Mins are not more than from Mins");
    }
    else{
      $("#lbl_inc_mins_to").text("");
    }
  }
  private setButtonsEventHandlers():void {
    //throw new Error('Method not implemented.');
    const webPart: CctvInternalRequestWebPart = this;
    
    this.domElement.querySelector('#btnSubmit').addEventListener('click', (e) => { 
      e.preventDefault();
      webPart.CreateNewCCTVRequest();
     });
     
     this.domElement.querySelector('#btnCancle').addEventListener('click', (e) => { 
      e.preventDefault();
      window.location.href= this.context.pageContext.site.absoluteUrl+"/Pages/TecPages/CCTV-Internal-Requests.aspx";
     });
     //

     this.domElement.querySelector('#txtEmpMail').addEventListener('blur', (e) => { 
      e.preventDefault();
      var empmail=$("#txtEmpMail").val();
      if(empmail!=""){
        if(this.IsEmail(empmail)==false){
          $("#lbl_emp_mail").text("mail must be xyz@tec.com.kw");
        }
        else{
          $("#lbl_emp_mail").text("");
          this.getDataFromEmpDirectoryList(empmail);
        }
      }
      else{
        $("#lbl_emp_mail").text("Email is mandatory");
      }
     });
     this.domElement.querySelector('#txtEmpName').addEventListener('blur', (e) => {this.validateTextBox("txtEmpName","Full Name is mandatory") });
     this.domElement.querySelector('#txtEmpID').addEventListener('blur', (e) => {this.validateTextBox("txtEmpID","Employee ID is mandatory") });
     this.domElement.querySelector('#txtEmpMobile').addEventListener('blur', (e) => {this.validateTextBox("txtEmpMobile","Mobile/Telephone Number is mandatory") });
     this.domElement.querySelector('#txtReqReason').addEventListener('blur', (e) => {this.validateTextBox("txtReqReason","Reason for Request is mandatory") });
     //this.domElement.querySelector('#txtEmpMail').addEventListener('blur', (e) => {this.validateTextBox("txtEmpMail","Email is mandatory") });
     //
     this.domElement.querySelector('#sel_Dept').addEventListener('blur', (e) => {this.validateDropdown("sel_Dept","Department is mandatory") });
     this.domElement.querySelector('#ddlIncidentHours').addEventListener('blur', (e) => {this.validateDropdown("ddlIncidentHours","Hours are mandatory") });
     this.domElement.querySelector('#ddlIncidentMins').addEventListener('blur', (e) => {this.validateDropdown("ddlIncidentMins","Mins are mandatory") });
     this.domElement.querySelector('#ddlIncidentHours_to').addEventListener('blur', (e) => {this.validateDropdownTo("ddlIncidentHours_to","Hours are mandatory","ddlIncidentHours","Hours are not less than from Hours") });
     this.domElement.querySelector('#ddlIncidentMins_to').addEventListener('blur', (e) => {this.validateDropdownToMins() });

     this.domElement.querySelector('#txtIncidentDate').addEventListener('blur', (e) => {this.validateDate("txtIncidentDate","Incident Date  (From) is mandatory") });
     this.domElement.querySelector('#txtIncidentDate_to').addEventListener('blur', (e) => {this.validateDate("txtIncidentDate_to","Incident Date  (To) is mandatory") });
     
  }
  private validateTextBox(e:string,errmsg:string):void{
   //const inputElement = e.target as HTMLInputElement;
    var inputval=$('#'+e).val();
    var inputspan=$('#'+e).next("label");
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
     var inputspan=$('#'+e).next("label");
     if(inputval=="0" || inputval=="HH" || inputval=="MM")
     {
       inputspan.text(errmsg);
   
     }
     else
     {
       inputspan.text("");
     }
   }
   private validateDropdownTo(e:string,errmsg:string,e1:string,errmsg2:string):void{
    //const inputElement = e.target as HTMLInputElement;
     var curinputval=$('#'+e).val();
     var preinputval=$('#'+e1).val();
     var inputspan=$('#'+e).next("label");
     var date1 = $('#txtIncidentDate').datepicker('getDate');
     var date2 = $('#txtIncidentDate_to').datepicker('getDate');
     var momentobj1=moment(date1); 
     var momentobj2=moment(date2);
     if(curinputval=="0" || curinputval=="HH" || curinputval=="MM")
     {
       inputspan.text(errmsg);
   
     }
     else if(momentobj1.format("dd-mm-yyyy")==momentobj2.format("dd-mm-yyyy") && curinputval<preinputval)
     {
       inputspan.text(errmsg2);
     }
     else{
      inputspan.text("");
     }
   }
   private validateDate(e:string,errmsg:string):void{
    //const inputElement = e.target as HTMLInputElement;$('#txtIncidentDate').datepicker('getDate');
     var inputval=$('#'+e).datepicker('getDate');
     var inputspan=$('#'+e).next("label");
     if(inputval==null || inputval==undefined)
     {
       inputspan.text(errmsg);
   
     }
     else
     {
       inputspan.text("");
     }
   }
  private getDataFromEmpDirectoryList(mail){
    sp.site.rootWeb.lists.getByTitle(this.EmpDirectoryListName).items.filter("EmpEmail eq '"+mail+"'").getAll()
    .then(r=>{
      console.log(r);
      if(r.length>0){
        if (r[0].Title!=null){
        $("#txtEmpName").val(r[0].Title).attr('disabled', 'disabled');$
        $("#lbl_emp_name").text("");
        }else{
          $("#txtEmpName").val("").removeAttr('disabled');
        }
        if(r[0].EmpPhone!=null){
          $("#txtEmpMobile").val(r[0].EmpPhone).attr('disabled', 'disabled');
          $("#lbl_emp_mobile").text("");
        }else{
          $("#txtEmpMobile").val(r[0].EmpPhone).removeAttr('disabled');
        }
        if(r[0].EmpDepartment!=null){
          //$("#sel_Dept").val(r[0].EmpDepartment).attr('disabled', 'disabled'); //enable after fixed
          $("#sel_Dept").val(0).removeAttr('disabled');
          $("#lbl_emp_dept").text("");
          
        }else{
          $("#sel_Dept").val(0).removeAttr('disabled');
        }
        if(r[0].TecEmployeeID!=null){
          $("#txtEmpID").val(r[0].TecEmployeeID).attr('disabled', 'disabled');
          $("#lbl_emp_id").text("");
        }else{
          $("#txtEmpID").val(r[0].TecEmployeeID).removeAttr('disabled');
        }
      }
      else{
        $("#txtEmpName").val("").removeAttr('disabled');
        $("#txtEmpMobile").val("").removeAttr('disabled');
        $("#sel_Dept").val(0).removeAttr('disabled');
        $("#txtEmpID").val("").removeAttr('disabled');
      }

    }).catch(function(err) {  
      console.log(err);  
    });
  }

  private IsEmail(email) {
    var regex = /^([a-zA-Z0-9_\.\-\+])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
    if(!regex.test(email)) {
      return false;
    }else{
      return true;
    }
  }
 

  private LoadDepartments():void{
    sp.site.rootWeb.lists.getByTitle(this.DeptListname).items.orderBy('Title', true).get()
    .then(function (data) {
      $('#sel_Dept').append(`<option value="0">Select Department</option>`);
      for (var k in data) {
        $("#sel_Dept").append("<option value=\"" +data[k].ID + "\">" +data[k].Title + "</option>");
      }
    });
  }
  
  private CreateNewCCTVRequest(){
   
    
    if(this.validations()==true){
      
        sp.site.rootWeb.lists.getByTitle(this.Listname).items.add({
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
          sp.site.rootWeb.lists.getByTitle(this.Listname).items.getById(r.data.ID).update({
            Title: "CCTV_Int_Req_00"+r.data.ID,
            TaskUrl: {
              "__metadata": { type: "SP.FieldUrlValue" },
              Description:this.ITManagerAction,
              Url: this.context.pageContext.web.absoluteUrl+this.ITMgrActionUrl+r.data.ID,
            },
          });
          //this.updateLogs(r.data.ID,r.data.AuthorId);
          alert("Thank you! your request has been successfully submitted");
          window.location.href= this.context.pageContext.site.absoluteUrl+"/Pages/TecPages/CCTV-Internal-Requests.aspx";
        }).catch(function(err) {  
          console.log(err);  
          }); 
    }
    else{
      alert("Sorry,Please check your form where some data is not in a valid format.");
    }
  }
  
  private updateLogs(itemid,AuthorID) {
    sp.site.rootWeb.lists.getByTitle(this.LogsListname).items.add({
      Title: "CCTV_Internal_Incident",
      Status: "Pending with IT Manager",
      StatusID:1,
      ItemID:itemid,
      AssignedTo:"IT Manager",
      InitiatedById:AuthorID
    }).then(iar => {
      alert("Thank you ! Your request was submitted Successfully");
      window.location.href= this.context.pageContext.web.absoluteUrl;
      //console.log(iar);
      //this.CheckAndCreateFolder(ITEMID);
    }).catch((error:any) => {
      console.log("Error: ", error);
    });
    // add an item to the list
    
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
      alert("You don't have access to create request,please check with administrator for more info");
      window.location.href=this.context.pageContext.web.absoluteUrl;
    }
    
  } 
 /* protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
 */

 

  private validations(){
        var isvalidform=true;

        this.valempname= document.getElementById('txtEmpName')["value"];
        this.valempdept= document.getElementById('sel_Dept')["value"]
        this.valempid= document.getElementById('txtEmpID')["value"];
        this.valempmail= document.getElementById('txtEmpMail')["value"];
        this.valReqReason= document.getElementById('txtReqReason')["value"];
        this.valempmobile= document.getElementById('txtEmpMobile')["value"];
        this.valincTimehrs= document.getElementById('ddlIncidentHours')["value"];
        this.valincTimeMins= document.getElementById('ddlIncidentMins')["value"];
        this.valincTimehrsTo= document.getElementById('ddlIncidentHours_to')["value"];
        this.valincTimeMinsTo= document.getElementById('ddlIncidentMins_to')["value"];
        this.valincDate = $('#txtIncidentDate').datepicker('getDate');
        this.valincDateTo=$('#txtIncidentDate_to').datepicker('getDate');//
        if($('#txtEmpName').val()=="" || $('#txtEmpName').val()==undefined)
        {
          $("#lbl_emp_name").text("Full Name is mandatory");
            isvalidform = false;
        }
        else
        {
          $("#lbl_emp_name").text("");
        }

        if(this.valempdept=="0")
        {
            $("#lbl_emp_dept").text("Department is mandatory");
            isvalidform = false;
        }
        else
        {
            $("#lbl_emp_dept").text("");
        }
        if(this.valempid==""){
          $("#lbl_emp_id").text("Employee ID is mandatory");
          isvalidform = false;
        }
        else{
          $("#lbl_emp_id").text(" ");
        }

        if(this.valempmail==""){
          $("#lbl_emp_mail").text("Email is mandatory");
          isvalidform = false;
        }
        else{
          $("#lbl_emp_mail").text(" ");
          if(this.IsEmail(this.valempmail)==false){
            $("#lbl_emp_mail").text("mail must be xyz@tec.com.kw");
            isvalidform = false;
          }
        }
   
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

        if(this.valempmobile==""){
          $("#lbl_emp_mobile").text("Mobile/Telephone Number is mandatory");
          isvalidform = false;
        }
        else{
          $("#lbl_emp_mobile").text(" ");
        }
        if(this.valincTimehrs=="HH"){
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
        }
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
