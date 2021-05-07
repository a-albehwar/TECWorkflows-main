export interface CCTVInternalList {
  Title: string;
  Departments:{
    Title:string;
    ID:number;
  }
  ID:number;
  Created:Date;
  DateOfIncident:Date;
  TimeOfIncident:string;
  DateOfIncident_To:Date,
  TimeOfIncident_To:string,
  ReasonForRequest:string;
  MoreInfofromLegalMgr:string;
  MoreInfoformSSS:string;
  Status:{
    Title:string;
    ID:number;
  }
  AssignedTo:{
    Title:string;
  }
  TaskUrl:{
    Description: string,
    Url: string
  }
  Author:{
    Title:string;
    ID:string;
  }
  MobileNumber:string;
  EmployeeName:string;
  EmployeeID:string;
  ITManagerComments: string;
  SSSComments :string;
  FootageAvailable:string;
  EmailAdress:string;
  LegalManagerComments:string;
 
 
}

export class CommonLinks
{
  public SSSActionUrl:string="/Pages/cctv/SSSIntCCTVReq.aspx";
  public LegalMgrActionUrl:string="/Pages/cctv/CCTVLegalMgr.aspx";
  public CCTVInternalTaskUrl:string="/Lists/CCTV_Internal_Incident";
}
  // field refers workflow logs list data
export interface ICCHistoryLogList {
  Title: string;
  ID:number;
  Created:Date;
  Comments:string;
  AssignedTo:string;
  Author:{
    Title:string;
    AuthorId:number;
  }
  Status:string;
  Modified:Date;
  InitiatedBy:{
    Title:string;
    ID:string;
  }
  ApprovedDate:Date;
  TaskCompletedBy:string;
}

