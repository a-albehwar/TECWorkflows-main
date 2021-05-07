import * as React from 'react';
import styles from './CctvInternalMyTasks.module.scss';
import { ICctvInternalMyTasksProps } from './ICctvInternalMyTasksProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as moment from 'moment';
import BootstrapTable from 'react-bootstrap-table-next';  
import { sp } from "@pnp/sp/presets/all";
import * as $ from 'jquery';
import { SPComponentLoader } from '@microsoft/sp-loader';
import paginationFactory from 'react-bootstrap-table2-paginator';

import ToolkitProvider, { Search } from 'react-bootstrap-table2-toolkit';
  const { SearchBar } = Search;
 

  let groups: any[] = [];
  let LoginUserGroupName:string;
  const empTablecolumns = [     
  {    
      dataField: "Title",    
      text: "Request Title",    
      //headerStyle : {backgroundColor: '#3999a7',color: '#ffffff'}    
  
  },  
  {    
    dataField: "EmployeeName",    
    text: "Employee Name", 
  },  
  {    
    dataField: "Departments.Title",    
    text: "Department", 
  }, 
  {    
      dataField: "Status.Title",    
      text: "Status"    
  },    
  {    
      dataField: "Created",    
      text: "Created Date",
      formatter:(cell, row, rowIndex, formatExtraData) => {
         var momentObj = moment(row.Created);
         var formatCreatedDate=momentObj.format('DD-MM-YYYY'); 
         return formatCreatedDate;
         }
       
  },   
  {    
    
    text: "View",
    formatter: (cell, row, rowIndex, formatExtraData) => {
    //  console.log(row.id+"--"+row.Created);
    var viewurl=row.TaskUrl!=null?row.TaskUrl.Url:"#";
    //console.log(viewimgurl);
    var viewimgurl="/sites/IntranetDev/Style%20Library/TEC/images/view.svg";// remmove sites/tec while move prod
    //var viewimgurl="/sites/IntranetDev/Style%20Library/TEC/images/view.svg";// uncomment while prod

    return <a href={viewurl}><img src={viewimgurl} className={"img-fluid"}/></a>
    }
  }, 
 ];  

 export interface CctvInternalMyTasksStates{    
  employeeList :any[]    
}


const paginationOptions = {    
        sizePerPage: 10,    // change as per client request
        hideSizePerPage: true,    
        hidePageListOnlyOnePage: true    
}; 

export default class CctvInternalMyTasks extends React.Component<ICctvInternalMyTasksProps,CctvInternalMyTasksStates> {
  constructor(props: ICctvInternalMyTasksProps){    
    super(props);    
    //viewimgurl=this.props.siteurl;
    this.state ={    
      employeeList : []    
    }  
       
  } 
  
  public getEmployeeDetails = () =>{    
  
    //alert(LoginUserGroupName);
    sp.site.rootWeb.lists.getByTitle("CCTV_Internal_Incident").items.filter("AssignedTo/Title eq '"+LoginUserGroupName+"' and Status/ID ne 9 and Status/ID ne 6 and Status/ID ne 4").select("Title","TaskUrl","EmployeeName","Created","ID","Status/ID","Status/Title","AssignedTo/ID","AssignedTo/Title","Departments/ID","Departments/Title")
    .expand("AssignedTo,Departments,Status").orderBy("Created",false).get().
    then((results : any)=>{    
      
        this.setState({    
          employeeList:results    
        });    
      
    });    
  }  

  public componentDidMount(){  
    this._checkUserInGroup();   
      
    
  }

  private async _checkUserInGroup()
  {
    let groups1 = await  sp.site.rootWeb.currentUser.groups();
    var IsLegalTeamMember:number;
    var ISITMangerMember:number;
    var SSSMember:number
    if(groups1.length>0){
      for(var i=0;i<groups1.length;i++){
        groups.push(groups1[i].Title);
      }
    console.log(groups);
    }
    if(groups.length>0)
    {
      IsLegalTeamMember=$.inArray( "LegalManager", groups ) ;
      ISITMangerMember=$.inArray( "ITManager", groups ) ;
      SSSMember=$.inArray( "System Security Specialist", groups ) ;
    }
    if(IsLegalTeamMember<0 && ISITMangerMember<0 && SSSMember<0){
      alert("You are Not Authorized user");
     // window.location.href=this.props.weburl;
    }
    if(IsLegalTeamMember>=0)
    {
      LoginUserGroupName="LegalManager";
      this.getEmployeeDetails();
    }
    else if(ISITMangerMember>=0){
      LoginUserGroupName="ITManager";
      this.getEmployeeDetails();
    }
    else if(SSSMember>=0){
      LoginUserGroupName="System Security Specialist";
      this.getEmployeeDetails();
    }   
  } 
 
  public render(): React.ReactElement<ICctvInternalMyTasksProps> {
  
   // let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";    
   // SPComponentLoader.loadCss(cssURL);

    return (
      <div>    
      <div className={ styles.container }>    
        <div className={ styles.row }>    
          <div className={ styles.column }>        
          </div>    
        </div>   
        <ToolkitProvider keyField="id" data={this.state.employeeList} columns={ empTablecolumns } >
          {
            props => (
              <div>
                
                <BootstrapTable keyField='id' noDataIndication="No records found." data={this.state.employeeList} columns={ empTablecolumns }  pagination={paginationFactory(paginationOptions)} classes="table table-bordered table-hover footable" wrapperClasses="table-responsive"/>
              </div>
            )
          }
          </ToolkitProvider> 
          
      </div>    
    </div>
    );
    
  }
  //$(".table-bordered ").addClass(" table-hover");
  
}
