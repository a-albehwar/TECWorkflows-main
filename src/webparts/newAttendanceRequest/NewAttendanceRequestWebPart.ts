import { Version } from '@microsoft/sp-core-library';
import 'jqueryui';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { Items, sp } from '@pnp/sp/presets/all';
import { SPComponentLoader } from '@microsoft/sp-loader';

import styles from './NewAttendanceRequestWebPart.module.scss';
import * as strings from 'NewAttendanceRequestWebPartStrings';

import * as moment from 'moment';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import * as $ from 'jquery';
import { AttendanceItem } from './../Interfaces/IAttendance';



export interface INewAttendanceRequestWebPartProps {
  description: string;
}

export default class NewAttendanceRequestWebPart extends BaseClientSideWebPart<INewAttendanceRequestWebPartProps> {

  public constructor() {
    super();
    SPComponentLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
  }

  private MasterAttendanceList:string='Attendance Process Request';
  private DepartmentList:string='LK_Departments';

  private Status_Request_Initiated:number=1;

  private SSS_GroupName:string='System Security Specialist';
  private SSS_GroupId:number;

  private MyRequestUrl='/Pages/TecPages/SearchAttendance.aspx';
  private EmpDirectoryListName='EmployeeDirectory';

 

  public render(): void {
    this.domElement.innerHTML = `
    <section class="inner-page-cont" style="margin-top:-30px;">
    <div class="container-fluid mt-5">
        <div class="col-md-10 mx-auto col-12">
            <div class="row user-info">
                <h3 class="mb-4 col-12">Employee Details</h3>
                <div class="col-md-4 col-12 mb-4">
                    <p>Full Name<span style="color:red">*</span></p>
                    <input type="text" id="txtEmployeeName" class="form-input" />
                    <span class="err-msg form-label" style="color:red;display: none;">Employee Name is mandatory</span>
                </div>

                <div class="col-md-4 col-12 mb-4">
                    <p>Employee ID<span style="color:red">*</span></p>
                    <input type="text" id="txtEmployeeID" class="form-input"  />
                    <span class="err-msg form-label" style="color:red;display: none;">Employee ID is mandatory</span>
                </div>


                <div class="col-md-4 col-12 mb-4">
                    <p>Department<span style="color:red">*</span></p>
                    <select name="ddlDepartment" id="ddlDepartment" class="form-input"></select>
                    <span class="err-msg form-label" style="color:red;display: none;">Department is mandatory</span>
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Mobile/Telephone Number<span style="color:red">*</span></p>
                    <input type="text" id="txtNumber" class="form-input" />
                    <span class="err-msg form-label" style="color:red;display: none;">Telephone Number is mandatory</span>
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Email Address<span style="color:red">*</span></p>
                    <input type="text" id="txtEmail" class="form-input" />
                    <span class="err-msg form-label" style="color:red;display: none;">Email Address is mandatory</span>
                    
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Request Created Date</p>
                    <input type="text" id="txtCurrentDate" class="form-input" disabled />
                </div>
                <h3 class="mb-4 col-12">Request Details</h3>

                <div class="col-md-4 col-12 mb-4">
                    <p>Date<span style="color:red">*</span></p>
                    <input type="text" id="calDateOfRequest" class="form-input" autocomplete="off" readonly="readonly"/>
                    <span class="err-msg form-label" id="dateSpan" style="color:red;display: none;">Date is mandatory</span>
                </div>

                <div class="col-md-4 col-12 mb-4">
                  <p>Time<span style="color:red">*</span></p>
                  <input id="TimeOfRequest" class="form-input" type="time" min="00:00" max="24:00">
                  <span class="err-msg form-label" id="datetimeSpan" style="color:red;display: none;">Time is mandatory</span>
                </div>
                
                <div class="col-md-4 col-12 mb-4">
                    <p>Duration (Minutes)<span style="color:red">*</span></p>
                    <select name="incidentmins" id="ddlIncidentMins" class="form-input">
                              <option value="MM">MM</option>
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
                    <span class="err-msg form-label" style="color:red;display: none;">Duration is mandatory</span>
                </div>
                <div class="col-md-6 col-12 mb-4">
                    <p>Reason for Request<span style="color:red">*</span></p>
                    <textarea id="txtReasonForRequest" class="form-input" style="height:auto!important" rows="3" cols="5"></textarea>
                    <span class="err-msg form-label" style="color:red;display: none;">Reason for Request is mandatory</span>
                </div>

            </div>
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

    this._setButtonEventHandlers();
    this.LoadInitialControls();
  }

  private _setButtonEventHandlers(): void{
    const webpart:NewAttendanceRequestWebPart=this;
    this.domElement.querySelector('#btnSubmit').addEventListener('click', (e)=>{
      e.preventDefault();       
        if(this.ValidateFields())
        {
            this.CreateNewAttendanceItem();
        }
        else
        {
          alert("Sorry, Please check your form where some data is not in a valid format.");
        }
      
    });

      this.domElement.querySelector('#btnCancel').addEventListener('click', (e) => {
      e.preventDefault();
      window.location.href=this.context.pageContext.site.serverRelativeUrl+this.MyRequestUrl;
    });


    this.domElement.querySelector('#txtEmail').addEventListener('change', (e) => 
    {this.ValidateTextBox("txtEmail"); });
    this.domElement.querySelector('#txtEmployeeName').addEventListener('change', (e) => 
    {this.ValidateTextBox("txtEmployeeName"); });
    this.domElement.querySelector('#txtEmployeeID').addEventListener('change', (e) => 
    {this.ValidateTextBox("txtEmployeeID"); });
    this.domElement.querySelector('#txtNumber').addEventListener('change', (e) => 
    {this.ValidateTextBox("txtNumber"); });
    this.domElement.querySelector('#txtReasonForRequest').addEventListener('change', (e) => 
    {this.ValidateTextBox("txtReasonForRequest"); });
    this.domElement.querySelector('#TimeOfRequest').addEventListener('blur', (e) => 
    {this.ValidateTextBox("TimeOfRequest"); });
    

    this.domElement.querySelector('#ddlDepartment').addEventListener('change', (e) => 
    {this.ValidateDropDown("ddlDepartment"); });
    
    this.domElement.querySelector('#ddlIncidentMins').addEventListener('change', (e) => 
    {this.ValidateDropDown("ddlIncidentMins"); });
  
    this.domElement.querySelector('#calDateOfRequest').addEventListener('blur', (e) => {this.ValidateDate() });
   

  }

  private LoadDepartments():void {
    //console.log("inside departments");
  sp.site.rootWeb.lists.getByTitle(this.DepartmentList).items.select("Title","ID").get()
  .then(function (data) {
    $("#ddlDepartment").append('<option value="-1">Please select a department</option>');
    for (var k in data) {
      $("#ddlDepartment").append('<option value="' + data[k].ID + '">' +data[k].Title + '</option>');
    } 
    
  
  });
  }

  private LoadInitialControls()
  {
    this.LoadDepartments();
    var currentDate= moment().format('DD-MM-YYYY')
    $('#txtCurrentDate').val(currentDate);
    var emailAddress:string;

    //let web = new Web(this.context.pageContext.site.absoluteUrl);
    let currentUser = sp.site.rootWeb.currentUser.get().then(function(res)
    { 
     emailAddress=res.Email;
    }).then(r=>{
      this.getDataFromEmpDirectoryList(emailAddress);
    });

    // set date of request date
    $('#calDateOfRequest').datepicker({
      changeMonth: true,
      changeYear: true,
      dateFormat: "dd-mm-yy",
      maxDate:0,
      onSelect:function(){
        // $('#calSurveyEndDate').prop("disabled",true);
        this.focus();
     }
      
    });
    this.GetAssignedToUserId();

  }

  private ValidateFields()
  {
    var isValid=true;

    //check department selected..
    if($('#ddlDepartment').val()=="-1")
  {
    $("#ddlDepartment").next("span").css("display", "block");
    isValid=false;
  }
  else
  {
    $("#ddlDepartment").next("span").css("display", "none");
  }

  //check name
  if($('#txtEmployeeName').val()=="" || $('#txtEmployeeName').val()==undefined)
  {
    $("#txtEmployeeName").next("span").css("display", "block");
    isValid=false;
  }
  else
  {
    $("#txtEmployeeName").next("span").css("display", "none");
  }
  //txtEmployeeID = employee id
  if($('#txtEmployeeID').val()=="" || $('#txtEmployeeID').val()==undefined)
  {
    $("#txtEmployeeID").next("span").css("display", "block");
    isValid=false;
  }
  else
  {
    $("#txtEmployeeID").next("span").css("display", "none");
  }

  //txtNumber.. telephone number
  if($('#txtNumber').val()=="" || $('#txtNumber').val()==undefined)
  {
    $("#txtNumber").next("span").css("display", "block");
    isValid=false;
  }
  else
  {
    $("#txtNumber").next("span").css("display", "none");
  }

  //txtEmail.. telephone number
  if($('#txtEmail').val()=="" || $('#txtEmail').val()==undefined)
  {
    $("#txtEmail").next("span").css("display", "block");
    isValid=false;
  }
  else
  {
    $("#txtEmail").next("span").css("display", "none");
  }

  //check department selected..
  if($('#ddlIncidentMins').val()=="MM")
  {
    $("#ddlIncidentMins").next("span").css("display", "block");
    isValid=false;
  }
  else
  {
    $("#ddlIncidentMins").next("span").css("display", "none");
  }

  //txtReasonForRequest

  if($('#txtReasonForRequest').val()=="" || $('#txtReasonForRequest').val()==undefined)
  {
    $("#txtReasonForRequest").next("span").css("display", "block");
    isValid=false;
  }
  else
  {
    $("#txtReasonForRequest").next("span").css("display", "none");
  }

  //Date of Request
  var dateRequest=$('#calDateOfRequest').datepicker('getDate');
  if(dateRequest == null || dateRequest == undefined)
  {
    //$("#calDateOfRequest").next("span").css("display", "block");
    //$('#calDateOfRequest').nextAll('span:first').css("display","block");
    $('#dateSpan').css("display","block");
    isValid=false;
  }
  else
  {
    //$("#calDateOfRequest").next("span").css("display", "none");
    //$('#calDateOfRequest').nextAll('span:first').css("display","none");
    $('#dateSpan').css("display","none");
  }
  if($('#TimeOfRequest').val()==""){
    $('#datetimeSpan').css("display","block");
    isValid=false;
  }else{
    $('#datetimeSpan').css("display","none");
  }
 
    return isValid;
  }



  private CreateNewAttendanceItem()
  {

    var requestDate=$('#calDateOfRequest').datepicker('getDate');
    var momentObj=moment(requestDate);
    var requesTime=$('#TimeOfRequest').val();
    
    //alert(requestTime);
    var formattedRequestDate=momentObj.format('yyyy') +','+momentObj.format('MM') +','+momentObj.format('DD') 
    var requestDateTime = new Date(formattedRequestDate); 
    //alert(requestDateTime);

    sp.site.rootWeb.lists.getByTitle(this.MasterAttendanceList).items.add({
      Title:$('#txtEmployeeName').val(),
      DepartmentId:parseInt($('#ddlDepartment').val().toString()),
      StatusId:this.Status_Request_Initiated,
      EmployeeID:$('#txtEmployeeID').val(),
      ContactNumber:$('#txtNumber').val(),
      ReasonForRequest:$('#txtReasonForRequest').val(),
      Email:$('#txtEmail').val(),
      TimeofAbsence:$('#ddlIncidentMins').val(),
      //DateofRequest:requestDate,
      DateofRequest:requestDateTime,
      AssignedToId:this.SSS_GroupId,
      TimeofRequest:requesTime,
      }).then(r=>{
      sp.site.rootWeb.lists.getByTitle(this.MasterAttendanceList).items.getById(r.data.ID).update({
        Title: "ATT_REQ_" + moment().format('YYYYMMDD') +"_00"+ r.data.ID,
        
      });
      //this.updateLogs(r.data.ID,r.data.AuthorId);
      alert("Thank you. The request was submitted succesfully.");
      window.location.href= this.context.pageContext.site.absoluteUrl+this.MyRequestUrl;
    }).catch(function(err) {  
      console.log(err);  
      }); 

  }

  private async TestReturnValue()
  {
    //var testReturnValue:number= await this._checkUserInAnalystGroup();
    //alert(await this._checkUserInAnalystGroup());

  }

  private async GetAssignedToUserId()
  {
    var grp = await sp.web.siteGroups.getByName(this.SSS_GroupName)();
    this.SSS_GroupId=grp.Id
    //return grp.Id;
    // var memberCount:number=-1;
    // let groups1 = await  sp.site.rootWeb.currentUser.groups();

    // if(groups1.length>0){
    //   for(var i=0;i<groups1.length;i++){
    //     groups.push(groups1[i].Title);
    //   }
    // }
    // if(groups.length>0)
    // {
     
    //   memberCount=$.inArray( "System Security Specialist", groups );
      
        
    // }
    // return memberCount;
    
  }

  private ValidateTextBox(e:string):void{
    
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

   private ValidateDropDown(e:string):void{
    
    var inputval=$('#'+e).val();
    var inputspan=$('#'+e).next("span");
    if(inputval=="-1")
    {
      
      inputspan.css("display", "block");
  
    }
    else
    {
     inputspan.css("display", "none");
    }
  }

  private ValidateDate()
  {
    var dateRequest=$('#calDateOfRequest').datepicker('getDate');
    if(dateRequest == null || dateRequest == undefined)
    {
      $('#dateSpan').css("display","block");
      
    }
    else
    {
      $('#dateSpan').css("display","none");
    }
  }


  private getDataFromEmpDirectoryList(mail){
    sp.site.rootWeb.lists.getByTitle(this.EmpDirectoryListName).items.filter("EmpEmail eq '"+mail+"'").getAll()
    .then(r=>{
      console.log(r);
      if(r.length>0){
            if(r[0].Title!=null){
              $('#txtEmployeeName').val(r[0].Title).attr('readonly','true');
            }
            if(r[0].EmpEmail!=null){
              $('#txtEmail').val(r[0].EmpEmail).attr('readonly','true');
            }
            if(r[0].EmpPhone!=null){
              $("#txtNumber").val(r[0].EmpPhone).attr('readonly','true');
            }
            if(r[0].EmpDepartment!=null){
              //$("#ddlDepartment").val(r[0].EmpDepartment).attr('readonly','true'); // Remove comments once department matched with AD Values.
            }
            //$("#txtEmployeeID").val(r[0].CivilID!=null?r[0].CivilID:"");
      }
 
    }).catch(function(err) {  
      console.log(err);  
    });
  }
}
