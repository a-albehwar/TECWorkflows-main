import * as React from 'react';
import styles from './CctvInternalAllRequests.module.scss';
import { ICctvInternalAllRequestsProps } from './ICctvInternalAllRequestsProps';
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
  const AllCCTVInternalReqcolumns = [     
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

 export interface CctvInternalAllRequestsStates{    
  AllCCTVInternalRequests :any[]    
}

const paginationOptions = {    
  sizePerPage: 10,    // change as per client request
  hideSizePerPage: true,    
  hidePageListOnlyOnePage: true    
}; 

export default class CctvInternalAllRequests extends React.Component<ICctvInternalAllRequestsProps,CctvInternalAllRequestsStates> {

  constructor(props: ICctvInternalAllRequestsProps){    
    super(props);    
    //viewimgurl=this.props.siteurl;
    this.state ={    
      AllCCTVInternalRequests : []    
    }  
       
  } 

  public getAllCCTVInternalRequests = () =>{    
  
    //alert(LoginUserGroupName);
    sp.site.rootWeb.lists.getByTitle("CCTV_Internal_Incident").items.select("Title","TaskUrl","EmployeeName","Created","ID","Status/ID","Status/Title","AssignedTo/ID","AssignedTo/Title","Departments/ID","Departments/Title")
    .expand("AssignedTo,Departments,Status").orderBy("Created",false).get().
    then((results : any)=>{    
      
        this.setState({    
          AllCCTVInternalRequests:results    
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
    
    if(groups1.length>0){
      for(var i=0;i<groups1.length;i++){
        groups.push(groups1[i].Title);
      }
    console.log(groups);
    }
    if(groups.length>0)
    {
      IsLegalTeamMember=$.inArray( "LegalManager", groups ) ;
    }
    
    if(IsLegalTeamMember>=0)
    {
      LoginUserGroupName="LegalManager";
      this.getAllCCTVInternalRequests();
    }
    else{
      alert("You are Not Authorized user");
      $("#btnNewRequest").hide();
    }
    
  } 
  public render(): React.ReactElement<ICctvInternalAllRequestsProps> {
    return (
      <div>    
      <div className={ styles.container }>    
        <div className={ styles.row }>    
          <div className={ styles.column }>        
          </div>    
        </div>   
          <div className={"row"}>
                <div className={"col-lg-12"}>
                <button className={"red-btn mt-4"} id="btnNewRequest" onClick={(e) => {
                                              e.preventDefault();
                                              window.location.href=this.props.weburl+"/Pages/TecPages/cctv/NewCCTVInternalRequest.aspx";
                                              }}>New CCTV Internal Request</button> 
                 </div>
          </div>
        <ToolkitProvider keyField="id" data={this.state.AllCCTVInternalRequests} columns={ AllCCTVInternalReqcolumns } >
          {
            props => (
              <div>
                
                <BootstrapTable keyField='id' noDataIndication="No records found." data={this.state.AllCCTVInternalRequests} columns={ AllCCTVInternalReqcolumns }  pagination={paginationFactory(paginationOptions)} classes="table table-bordered table-hover footable" wrapperClasses="table-responsive"/>
              </div>
            )
          }
          </ToolkitProvider> 
          
      </div>    
    </div>
    );
  }
}
