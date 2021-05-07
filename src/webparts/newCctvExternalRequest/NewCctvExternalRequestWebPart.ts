import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './NewCctvExternalRequestWebPart.module.scss';
import * as strings from 'NewCctvExternalRequestWebPartStrings';
import { sp } from "@pnp/sp/presets/all";
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as moment from 'moment';
import * as $ from 'jquery';
import 'jqueryui';

export interface INewCctvExternalRequestWebPartProps {
  description: string;
}
var AssignedITGroupID:any;
var todaydt = new Date();
let groups: any[] = [];
var IsLegalTeamMember:number;

export default class NewCctvExternalRequestWebPart extends BaseClientSideWebPart<INewCctvExternalRequestWebPartProps> {
  private AssignedToGroupITManager:string='LegalManager';
  private EmpDirectoryListName:string="EmployeeDirectory";
  private valempid:string;
  private valempname:string;
  private valReqReason:string;
  private  valempdept:string;
  private valempmobile:string;
  private valincTimehrs:string;
  private valempmail:string;
  private valincDate:Date;
  private valincTimeMins:string;
  private valincTimehrsTo:string;
  private valincTimeMinsTo:string;
  private valincDateTo:Date;
  private valLocationFacility:string;
  private Listname: string = "CCTV_External_Incident";
  
  private DeptListname:string="LK_Departments";
  
  public constructor() {
    super();
    SPComponentLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
  }
  public render(): void {
    this.domElement.innerHTML = `
    <section class="inner-page-cont" style="margin-top: -60px">

      <div class="container-fluid mt-5">
        
            <div class="row user-info col-md-10 mx-auto col-12 pl-0 pr-0 m-0">
                <h3 class="mb-4 col-12">Requester Details</h3>
                      <div class="col-md-4 col-12 mb-4">
                        <p>Requester Name<span  style="color:red">*</span></p>
                        <input type="text" id="txtEmpName" class="form-input" name="txtEmpName">
                        <label id="lbl_emp_name" class="form-label" style="color:red"></label>
                      </div>
                    
                      <div class="col-md-4 col-12 mb-4">
                          <p>Email<span  style="color:red">*</span></p>
                          <input type="text" id="txtEmpMail" class="form-input" name="txtEmpMail">
                          <label id="lbl_emp_mail" class="form-label" style="color:red"></label>
                      </div>
                      <div class="col-md-4 col-12 mb-4">
                          <p>Mobile/Telephone Number<span  style="color:red">*</span></p>
                          <input type="text" id="txtEmpMobile" class="form-input" name="txtEmpMobile">
                          <label id="lbl_emp_mobile" class="form-label" style="color:red"></label>
                      </div>
                      <h3 class="mb-4 col-12">Incident Details</h3>
                     
                      <div class="col-md-4 col-12 mb-4">
                          <p>Incident Date & Time(From)<span  style="color:red">*</span></p>
                          <input type="text" autocomplete="off" id="txtIncidentDateFrom" class="form-input" name="txtIncidentDateFrom"  readonly="readonly" >
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
                          <p>Incident Date & Time(To)<span  style="color:red">*</span></p>
                          <input type="text" autocomplete="off" id="txtIncidentDateTo" class="form-input" name="txtIncidentDateFrom"  readonly="readonly" >
                          <label id="lbl_Inc_date_to" class="form-label" style="color:red"></label>
                      </div>
                      <div class="col-lg-2">
                          <p>HH<span  style="color:red">*</span></p>
                                <select name="incidenthours_To" id="ddlIncidentHours_To" class="form-input">
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
                                <label id="lbl_inc_hours_To" class="form-label" style="color:red"></label>
                            </div>
                            <div class="col-lg-2"> 
                                <p>MM<span  style="color:red">*</span></p>
                                <select name="incidentmins" id="ddlIncidentMins_To" class="form-input">
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
                              <label id="lbl_inc_mins_To" class="form-label" style="color:red"></label>
                            </div>
                          <div class="col-md-4 col-12 mb-4">
                            <p>Location/Facility of incident:<span  style="color:red">*</span></p>
                            <input type="text" id="txtLocation" class="form-input" name="txtLocation">
                            <label id="lbl_location_err" class="form-label" style="color:red"></label>
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
  }
  private PageLoad():void{
  

    this.LoadDepartments();
    
     $("#txtIncidentDateFrom").datepicker({
                dateFormat: "dd/mm/yy",
                endDate: todaydt,
                changeMonth: true,
                changeYear: true,   
                maxDate:0,
                onSelect: function (date) {
                //Get selected date 
                    var date2 = $('#txtIncidentDateFrom').datepicker('getDate');
                    //sets minDate to txt_date_to
                    $('#txtIncidentDateTo').datepicker('option', 'minDate', date2);
                    this.focus();
                }
            });
            $('#txtIncidentDateTo').datepicker({
                dateFormat: "dd/mm/yy",
                changeMonth: true,
                changeYear: true,
                maxDate:0,
                onSelect:function(){
                   // $('#calSurveyEndDate').prop("disabled",true);
                   this.focus();
                }
               
            });

  }

  private setButtonsEventHandlers():void {
    //throw new Error('Method not implemented.');
    const webPart: NewCctvExternalRequestWebPart = this;
    
    this.domElement.querySelector('#btnSubmit').addEventListener('click', (e) => { 
      e.preventDefault();
          if(this.validations()==true){
            webPart.CreateNewCCTVExternalRequest();
          }
          else{
            alert("Sorry,Please check your form where some data is not in a valid format.");
          }
     });
     
     this.domElement.querySelector('#btnCancle').addEventListener('click', (e) => { 
      e.preventDefault();
      window.location.href= this.context.pageContext.site.absoluteUrl+"/Pages/TecPages/CCTV-External-Requests.aspx";
     });

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

     this.domElement.querySelector('#txtEmpName').addEventListener('blur', (e) => {this.validateTextBox("txtEmpName","Requester Name is mandatory") });
     this.domElement.querySelector('#txtLocation').addEventListener('blur', (e) => {this.validateTextBox("txtLocation","Location/Facility of incident is mandatory") });
     this.domElement.querySelector('#txtEmpMobile').addEventListener('blur', (e) => {this.validateTextBox("txtEmpMobile","Mobile/Telephone Number is mandatory") });
     this.domElement.querySelector('#txtReqReason').addEventListener('blur', (e) => {this.validateTextBox("txtReqReason","Reason for Request is mandatory") });
     
     this.domElement.querySelector('#ddlIncidentHours').addEventListener('blur', (e) => {this.validateDropdown("ddlIncidentHours","Hours are mandatory") });
     this.domElement.querySelector('#ddlIncidentMins').addEventListener('blur', (e) => {this.validateDropdown("ddlIncidentMins","Mins are mandatory") });
     this.domElement.querySelector('#ddlIncidentHours_To').addEventListener('blur', (e) => {this.validateDropdownTo("ddlIncidentHours_To","Hours are mandatory","ddlIncidentHours","Hours are not less than from Hours") });
     this.domElement.querySelector('#ddlIncidentMins_To').addEventListener('blur', (e) => {this.validateDropdownToMins() });

     this.domElement.querySelector('#txtIncidentDateFrom').addEventListener('blur', (e) => {this.validateDate("txtIncidentDateFrom","Incident Date  (From) is mandatory") });
     this.domElement.querySelector('#txtIncidentDateTo').addEventListener('blur', (e) => {this.validateDate("txtIncidentDateTo","Incident Date  (To) is mandatory") });
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
    private validateDropdownToMins()
    {
      var ddlprehrs=$('#ddlIncidentHours').val();
      var ddlcurhrs=$('#ddlIncidentHours_To').val(); 
      var date1 = $('#txtIncidentDateFrom').datepicker('getDate');
      var date2 = $('#txtIncidentDateTo').datepicker('getDate');
      var minsFrom=$('#ddlIncidentMins').val();
      var minsTo=$('#ddlIncidentMins_To').val();
      var momentobj1=moment(date1); 
      var momentobj2=moment(date2);
      if( minsTo=="MM")
      {
        $("#lbl_inc_mins_To").text("Mins are mandatory");
      }
      else if(momentobj1.format("dd-mm-yyyy")==momentobj2.format("dd-mm-yyyy") && ddlprehrs==ddlcurhrs && minsFrom>minsTo){
        $("#lbl_inc_mins_To").text("Mins are not more than from Mins");
      }
      else{
        $("#lbl_inc_mins_To").text("");
      }
    }
    private validateDropdownTo(e:string,errmsg:string,e1:string,errmsg2:string):void{
      //const inputElement = e.target as HTMLInputElement;
       var curinputval=$('#'+e).val();
       var preinputval=$('#'+e1).val();
       var inputspan=$('#'+e).next("label");
       var ddlprehrs=$('#ddlIncidentHours').val();
       var ddlcurhrs=$('#ddlIncidentHours_To').val();
       var date1 = $('#txtIncidentDateFrom').datepicker('getDate');
       var date2 = $('#txtIncidentDateTo').datepicker('getDate');
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
 
  private LoadDepartments():void{
    sp.site.rootWeb.lists.getByTitle(this.DeptListname).items.orderBy('Title', true).get()
    .then(function (data) {
      $('#sel_Dept').append(`<option value="0">Select Department</option>`);
      for (var k in data) {
        $("#sel_Dept").append("<option value=\"" +data[k].ID + "\">" +data[k].Title + "</option>");
      }
    });
  }

  private validations(){
        var isvalidform=true;

        this.valempname= document.getElementById('txtEmpName')["value"];
        this.valempmail= document.getElementById('txtEmpMail')["value"];
        this.valReqReason= document.getElementById('txtReqReason')["value"];
        this.valempmobile= document.getElementById('txtEmpMobile')["value"];
        this.valincTimehrs= document.getElementById('ddlIncidentHours')["value"];
        this.valincTimeMins= document.getElementById('ddlIncidentMins')["value"];
        this.valincTimehrsTo= document.getElementById('ddlIncidentHours_To')["value"];
        this.valincTimeMinsTo= document.getElementById('ddlIncidentMins_To')["value"];
        this.valincDate = $('#txtIncidentDateFrom').datepicker('getDate');
        this.valincDateTo=$('#txtIncidentDateTo').datepicker('getDate');
        this.valLocationFacility=document.getElementById('txtLocation')["value"];

        if($('#txtEmpName').val()=="" || $('#txtEmpName').val()==undefined)
        {
          $("#lbl_emp_name").text("Requester Name is mandatory");
            isvalidform = false;
        }
        else
        {
          $("#lbl_emp_name").text("");
        }
        if($('#txtLocation').val()=="" || $('#txtLocation').val()==undefined)
        {
          $("#lbl_location_err").text("Location/Facility of incident is mandatory");
            isvalidform = false;
        }
        else
        {
          $("#lbl_location_err").text("");
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
            $("#lbl_Inc_date").text("Incident Date (From) is mandatory");
            isvalidform = false;
        }
        else
        {
            $("#lbl_Inc_date").text("");
        }
       if(this.valincDateTo==null || this.valincDateTo == undefined)
        {
            $("#lbl_Inc_date_to").text("Incident Date (To) is mandatory");
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
          $("#lbl_inc_hours_To").text("Hours are mandatory");
          isvalidform = false;
        }
        else{
          $("#lbl_inc_hours_To").text(" ");
        }
        
        if(this.valincTimeMinsTo=="MM"){
          $("#lbl_inc_mins_To").text("Mins are mandatory");
          isvalidform = false;
        }
        else{
          $("#lbl_inc_mins_To").text(" ");
        }
        var momentobj1=moment(this.valincDate); 
        var momentobj2=moment(this.valincDateTo);
        if((momentobj1.format("dd/mm/yyyy")==momentobj2.format("dd/mm/yyyy"))&&(this.valincTimehrs>this.valincTimehrsTo))
        {
          $("#lbl_inc_hours_To").text("Hours are not less than from Hours");
          isvalidform = false;
        }
        else{
          $("#lbl_inc_hours_To").text(" ");
        }
        if((momentobj1.format("dd/mm/yyyy")==momentobj2.format("dd/mm/yyyy"))&&(this.valincTimehrs==this.valincTimehrsTo)&&(this.valincTimeMins>this.valincTimeMinsTo))
        {
          $("#lbl_inc_mins_To").text("Mins are not less than from Mins");
          isvalidform = false;
        }
        else{
          $("#lbl_inc_mins_To").text(" ");
        }
      
         
        return isvalidform;
  }

  private IsEmail(email) {
    var regex = /^[\w-\.]+@([tec]{3})+\.+([com]{3})+\.+[kw]{2}$/;
   //var regex=/^[\w-\.]+@([\w-]+\.)+([\w-]+\.)+[\w-]{2}$/;
    if(!regex.test(email)) {
      return false;
    }else{
      return true;
    }
  }

  private CreateNewCCTVExternalRequest(){
      
        sp.site.rootWeb.lists.getByTitle(this.Listname).items.add({
          Title:this.valempname,
          EmailAddress:this.valempmail,
          Mobile_x002d_Tel_x0020_No:this.valempmobile,
          ReasonForRequest:this.valReqReason,
          StatusId:1,
          DateOfIncident:this.valincDate,
          DateOfIncident_To:this.valincDateTo,
          TimeOfIncident:this.valincTimehrs+":"+this.valincTimeMins,
          TimeOfIncident_To:this.valincTimehrsTo+":"+this.valincTimeMinsTo,
          RequesterName:this.valempname,
          LocationFacilityOfIncident:this.valLocationFacility
        }).then(r=>{
          sp.site.rootWeb.lists.getByTitle(this.Listname).items.getById(r.data.ID).update({
            Title: "CCTV_Ext_Req_00"+r.data.ID,
          });
          //this.updateLogs(r.data.ID,r.data.AuthorId);
          alert("Thank you ! Your request was submitted Successfully");""
          window.location.href= this.context.pageContext.site.absoluteUrl+"/Pages/TecPages/CCTV-External-Requests.aspx";
        }).catch(function(err) {  
          console.log(err);  
          }); 
  }

  private getDataFromEmpDirectoryList(mail){
    sp.site.rootWeb.lists.getByTitle(this.EmpDirectoryListName).items.filter("EmpEmail eq '"+mail+"'").getAll()
    .then(r=>{
      console.log(r);
      if(r.length>0){
        $("#txtEmpName").val(r[0].Title!=null?r[0].Title:"");
        //$("#txtEmpID").val(r[0].CivilID!=null?r[0].CivilID:"");
        $("#txtEmpMobile").val(r[0].EmpPhone!=null?r[0].EmpPhone:"");
        //$("#txtEmpMobile").val(r[0].EmpDepartment!=null?r[0].EmpDepartment:0);// check and match employee department with fuad

        //$("#sel_Dept option:contains(" + r[0].EmpDepartment!=null?r[0].EmpDepartment:0 + ")").attr('selected', 'selected');
        // r[0].EmpPhone;
        // r[0].EmpDesignation;
        // r[0].CivilID;
      }

    }).catch(function(err) {  
      console.log(err);  
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
