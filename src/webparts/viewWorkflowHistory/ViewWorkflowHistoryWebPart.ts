import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { ICCHistoryLogList } from './../Interfaces/ICCTVInternal';
//import  * as footable from 'footable';
// @ts-ignore
import * as footable from "footable";

//import {footable} from 'footable';
import { SPComponentLoader } from '@microsoft/sp-loader';

import * as $ from 'jquery';
import * as moment from 'moment';


// import styles from './ViewWorkflowHistoryWebPart.module.scss';
import * as strings from 'ViewWorkflowHistoryWebPartStrings';

export interface IViewWorkflowHistoryWebPartProps {
  description: string;

  
}


SPComponentLoader.loadScript("https://tecq8.sharepoint.com/sites/IntranetDev/Style%20Library/TEC/JS/footable.js");
export default class ViewWorkflowHistoryWebPart extends BaseClientSideWebPart<IViewWorkflowHistoryWebPartProps> {
 
  private reqItemID:number;
  private HistoryLogList: string= 'WorkflowLogs';
  //private masterListName:string=this.properties.description;

  public render(): void {
    this.domElement.innerHTML = `
    <div id='dvHistoryMain' style="margin-top:30px"><h2 style="margin-left:22px;color: #3999a7;">Request History</h2>

    <div class="container-fluid">
        <div class="row">
            <div class="col-12" id="dvDataTable">
    
            </div>

            

        </div>
    </div>
  </div>  
  `;

    this.PageLoad();
  }

  private PageLoad():void{
    
    
    const url : any = new URL(window.location.href);
    this.reqItemID= url.searchParams.get("ItemID");
    this.BindDetails();
  }

  private BindDetails()
  {
    var logComments='';

    var masterlistwfname=this.properties.description;
    this.context.spHttpClient.get(`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${this.HistoryLogList}')/items?$select=*,InitiatedBy/ID,InitiatedBy/Title,Author/ID,Author/Title&$expand=Author,InitiatedBy&$filter=((ItemID%20eq%20${this.reqItemID})and(Title%20eq%20%27${masterlistwfname}%27))&$orderby=ID%20desc`, SPHttpClient.configurations.v1)      
                                                                                ///items?$select=*,Author/ID,Author/Title&$expand=Author&$filter=((ItemID%20eq%20163)and(Title%20eq%20%27SuggestionsBox%27))
    .then(response=>{
      //console.log(response.json());

      return response.json()
      .then((items: any): void => {
        //console.log(items);
        //console.log(items["value"]);//table table-bordered table table-bordered table-hover footable
        
        var displaytable ='<table class="table table-bordered table-hover footable requesthistory"><thead><tr><th>SN</th><th>Requested On</th><th  data-breakpoints="xs">Requested By</th><th data-breakpoints="xs">Status</th><th  data-breakpoints="xs">Assigned Role</th><th  data-breakpoints="xs">Assignee Comments</th><th  data-breakpoints="xs">Assignee Actioned On</th><th  data-breakpoints="xs">Task Completed By</th></tr></thead><tbody>';

        //displaytable+='<tbody>';
        let listItems: ICCHistoryLogList[] = items["value"];
        for(var i:number=0;i<listItems.length;i++){
       // listItems.forEach(logItem => {
          //console.log(logItem);
          var logItem=listItems[i];
          
          logComments=logItem.Comments?logItem.Comments:"-";
          var formatApprovedDate;
          if(logItem.ApprovedDate!=null){
          var momentApprovedObj=moment(logItem.ApprovedDate)
          formatApprovedDate=momentApprovedObj.format('DD-MM-YYYY  HH:mm');
          }else{
            formatApprovedDate="";
          }
          var momentCreatedObj = moment(logItem.Created);           
          var formatCreatedDate=momentCreatedObj.format('DD-MM-YYYY  HH:mm');
          var intiatedperson=logItem.InitiatedBy!=null?logItem.InitiatedBy.Title:"";
          var assignedperson=logItem.AssignedTo!=null?logItem.AssignedTo:"";
          var TaskCompletedBy=logItem.TaskCompletedBy!=null?logItem.TaskCompletedBy:"";
          var sno=(listItems.length)-i;
          displaytable += "<tr><td>"
                    + sno + "</td><td>"   
                    + formatCreatedDate + "</td><td>"
                    + intiatedperson + "</td><td>"
                    + logItem.Status + "</td><td>"
                    + assignedperson + "</td><td>"
                    + logComments + "</td><td>" 
                    + formatApprovedDate + "</td><td>"
                    + TaskCompletedBy + "</td></tr>";
                    //+ logItem.Comments
                    // + "</td><tr>";
       // });
      }
        displaytable += '</tbody></table>';
        //displaytable.footable();
       //$('.table').footable();
        $('#dvDataTable').html(displaytable);
      });
    });


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
