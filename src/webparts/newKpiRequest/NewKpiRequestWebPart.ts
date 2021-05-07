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

import styles from './NewKpiRequestWebPart.module.scss';
import * as strings from 'NewKpiRequestWebPartStrings';

import * as moment from 'moment';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import * as $ from 'jquery';
import { KPIReportRequestItem, TECDepartments,KPIValueItem, KPIDocItem } from './../Interfaces/IKPIRequest';

export interface INewKpiRequestWebPartProps {
  description: string;
}

var KPIListName="KPI Reporting Request";
var DepartmentList:string="TECDepartments";
var KPIConfigList="KPIFields";
var KPIValueList="KPIValues";
var KPIReportLibrary="KPIPerformanceReports";

var KPIAnalystTeam="KPIAnalyst";
var KPIOwnerGroupName:string;

var Status_populateData:number=1;

export default class NewKpiRequestWebPart extends BaseClientSideWebPart<INewKpiRequestWebPartProps> {

  public constructor() {
    super();
    SPComponentLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
  }

  private MyRequestUrl='';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="inner-page-cont" style="margin-top:-30px;"> 
    <div class="container-fluid mt-5">
        <div class="col-md-10 mx-auto col-12">
            <div class="row user-info">
                <h3 class="mb-4 col-12">Request Details</h3>
                <div class="col-md-4 col-12 mb-4">
                    <p>Department</p>
                    <select name="ddlDepartment" id="ddlDepartment" class="form-input"></select>
                    <span class="err-msg form-label" style="color:red;display: none;">Department is mandatory</span>
                </div>

                <div class="col-md-4 col-12 mb-4">
                    <p>Pre Set Date</p>
                    <input type="text" id="calPreferenceDate" class="form-input" />
                    <span class="err-msg form-label" style="color:red;display: none;">Pre Set Date is mandatory</span>
                </div>


                <div class="col-md-4 col-12 mb-4">
                    <p>Time Period</p>
                    <select name="ddlTimePeriod" id="ddlTimePeriod" class="form-input"></select>
                    <span class="err-msg form-label" style="color:red;display: none;">Time Period is mandatory</span>
                </div>
                <div class="col-md-4 col-12 mb-4">
                    <p>Period</p>
                    <select name="ddlPeriod" id="ddlPeriod" class="form-input"></select>
                    <span class="err-msg form-label" style="color:red;display: none;">Period is mandatory</span>
                </div>
                <div class="col-md-6 col-12 mb-4">
                    <p>Comments</p>
                    <textarea id="txtRequesterComments" class="form-input" style="height:auto!important" rows="3" cols="5"></textarea>
                </div>

            </div>
        </div>
    </div>  
    <div class="container-fluid mt-5" id="dvButtonMain">
        <div class="col-md-10 mx-auto col-12">
            <div class="row user-info">
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
    this.LoadScripts();
    this._setButtonEventHandlers();
  }


  private _setButtonEventHandlers(): void{
    const webpart:NewKpiRequestWebPart=this;
    this.domElement.querySelector('#btnSubmit').addEventListener('click', (e)=>{
      e.preventDefault();
        
    
        if(this.ValidateFields())
        {
            this.GetDepartmentById();
        }
        else
        {
          alert("Sorry, Please check your form where some data is not in a valid format.");
        }
      
    });

      this.domElement.querySelector('#btnCancel').addEventListener('click', (e) => {
      e.preventDefault();
      window.location.href=this.context.pageContext.web.serverRelativeUrl;
    });

    this.domElement.querySelector('#ddlTimePeriod').addEventListener('change', (e) => {
      //alert("Dropdown changed");
      this.LoadPeriod();
      this.ValidateDropDown("ddlTimePeriod");
    });
    this.domElement.querySelector('#ddlDepartment').addEventListener('change', (e) => 
    {this.ValidateDropDown("ddlDepartment"); });

    this.domElement.querySelector('#ddlPeriod').addEventListener('change', (e) => 
    {this.ValidateDropDown("ddlPeriod"); });

    this.domElement.querySelector('#calPreferenceDate').addEventListener('change', (e) => 
    {
      this.ValidateDate("calPreferenceDate");
    });



  }

  private LoadDepartments():void {
    //console.log("inside departments");
  sp.site.rootWeb.lists.getByTitle(DepartmentList).items.select("Title","ID","KPIOwnerId").get()
  .then(function (data) {
    $("#ddlDepartment").append('<option value="0">Please select a department</option>');
    for (var k in data) {
      $("#ddlDepartment").append('<option value="' + data[k].ID + '">' +data[k].Title + '</option>');
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
        //$(ControlName).val(requestFormItem[ChoiceColumnName]);
      }

    });
     
  }

  private LoadScripts()
  {
      // set pref date
      $('#calPreferenceDate').datepicker({
        changeMonth: true,
        changeYear: true,
        dateFormat: "dd-mm-yy",
        
      });

  }

  private LoadPeriodDropDown()
  {
     var selectedVal=$('#ddlTimePeriod').val();
     $('#ddlPeriod').empty();
     $("#ddlPeriod").append('<option value="0">Please select a period</option>');
     if(selectedVal=="Monthly")
     {
       $("#ddlPeriod").append('<option value="January">January</option>');
       $("#ddlPeriod").append('<option value="February">February</option>');
       $("#ddlPeriod").append('<option value="March">March</option>');
       $("#ddlPeriod").append('<option value="April">April</option>');
       $("#ddlPeriod").append('<option value="May">May</option>');
       $("#ddlPeriod").append('<option value="June">June</option>');
       $("#ddlPeriod").append('<option value="July">July</option>');
       $("#ddlPeriod").append('<option value="August">August</option>');
       $("#ddlPeriod").append('<option value="Septembeer">Septembeer</option>');
       $("#ddlPeriod").append('<option value="October">October</option>');
       $("#ddlPeriod").append('<option value="November">November</option>');
       $("#ddlPeriod").append('<option value="Decemeber">Decemeber</option>');
     }
     if(selectedVal=="Quarterly")
     {
       $("#ddlPeriod").append('<option value="Quarter 1">Quarter 1</option>');
       $("#ddlPeriod").append('<option value="Quarter 2">Quarter 2</option>');
       $("#ddlPeriod").append('<option value="Quarter 3">Quarter 3</option>');
       $("#ddlPeriod").append('<option value="Quarter 4">Quarter 4</option>');

       
     }
     if(selectedVal=="Quarterly")
     {
       

       $("#ddlPeriod").append('<option value="May">May</option>');
       $("#ddlPeriod").append('<option value="June">June</option>');
       $("#ddlPeriod").append('<option value="July">July</option>');
       $("#ddlPeriod").append('<option value="August">August</option>');
       $("#ddlPeriod").append('<option value="Septembeer">Septembeer</option>');
       $("#ddlPeriod").append('<option value="October">October</option>');
       $("#ddlPeriod").append('<option value="November">November</option>');
       $("#ddlPeriod").append('<option value="Decemeber">Decemeber</option>');
     }
  }

  private LoadPeriod() {

    $('#ddlPeriod').empty();//clear second dropdown..
    //$("#ddlPeriod").append("<option value='-1'>Select a Value</option>");

    var selectedValue = $('#ddlTimePeriod').val();
    var selectedText = $('#ddlTimePeriod option:selected').text();

    if (selectedValue != "0") {
        var $periodDropdown = $('#ddlPeriod');

        var periodOptions = {
            'Monthly': ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
            'Quarterly': ['Quarter 1', 'Quarter 2', 'Quarter 3', 'Quarter 4'],
            'Semi-Annual': ['Period 1', 'Period 2'],
            'Annual': ['2021', '2022', '2022'],
        }

        var selectedText = $('#ddlTimePeriod option:selected').text();
        var subPeriod = selectedText, lcns = periodOptions[subPeriod] || [];

        var html = $.map(lcns, function (lcn) {
            return '<option value="' + lcn + '">' + lcn + '</option>'
        }).join('');


        $periodDropdown.html('<option value="0">Select a Value</option>');
        $periodDropdown.append(html);
    }

}

private ValidateFields()
{
  
  var isValid=true;

  if($('#ddlDepartment').val()=="0")
  {
    $("#ddlDepartment").next("span").css("display", "block");
    isValid=false;
  }
  else
  {
    $("#ddlDepartment").next("span").css("display", "none");
  }
  var dateVal = $('#calPreferenceDate').datepicker('getDate');
  if (dateVal == null || dateVal == undefined) {
    $("#calPreferenceDate").next("span").css("display", "block");
    isValid = false;
  }
  else {
    $("#calPreferenceDate").next("span").css("display", "none");
  }

  if ($("#ddlTimePeriod").val() == "0") {
    $("#ddlTimePeriod").next("span").css("display", "block");
    //.error-msg
    isValid = false;
}
else {
    $("#ddlTimePeriod").next("span").css("display", "none");
}


if ($("#ddlPeriod").val() == "0") {
    $("#ddlPeriod").next("span").css("display", "block");
    //.error-msg
    isValid = false;
}
else {
    $("#ddlPeriod").next("span").css("display", "none");
}
  return isValid;
}

private CreateNewKPIRequest(KPIOwnerVal)
{
  
  var preSetDate=$('#calPreferenceDate').datepicker('getDate');

  sp.site.rootWeb.lists.getByTitle(KPIListName).items.add({
    Title:'KPI_Request',
    DepartmentId:parseInt($('#ddlDepartment').val().toString()),
    StatusId:Status_populateData,
    TimePeriod:$('#ddlTimePeriod').val(),
    Period:$('#ddlPeriod').val(),
    PreSetDate:preSetDate,
    RequesterComments:$('#txtRequesterComments').val(),
    KPIOwnerId:KPIOwnerVal,
    
    
  }).then(r=>{
    sp.site.rootWeb.lists.getByTitle(KPIListName).items.getById(r.data.ID).update({
      Title: "KPI_Req_" + moment().format('YYYYMM') + r.data.ID,
      
    });
    alert("Thank You! Your request was submitted successfully.");
    window.location.href= this.context.pageContext.web.absoluteUrl+this.MyRequestUrl;
  }).catch(function(err) {  
    console.log(err);  
    }); 


}


private GetDepartmentById() {
  var deptIdVal=parseInt($('#ddlDepartment').val().toString());
  var URL = "";
  URL =`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${DepartmentList}')/items?$select=*,KPIOwner/ID,KPIOwner/Title&$expand=KPIOwner/Id&$filter=ID eq `+deptIdVal;
  this.context.spHttpClient
    .get(URL, SPHttpClient.configurations.v1)
    .then((response) => {
      return response.json().then((items: any): void => {
          //let listItems: KPIReportRequestItem[] = items.value;
          items.value.forEach((item: TECDepartments) => {
          
          this.CreateNewKPIRequest(item.KPIOwnerId);
          

        });
        
        //this.getRelatedDocuments(RequestTitle);
      }).catch(function(err) {  
        console.log(err);  
      });
    });
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

private ValidateDate(e:string):void{
  alert("Inside validate date");
    var dateRequest=$('#'+e).datepicker('getDate');
    var inputspan=$('#'+e).next("span");
  if(dateRequest == null || dateRequest == undefined)
  {
    inputspan.css("display","block");
    
  }
  else
  {
    inputspan.css("display","none");
  }
  }
  

  
}
