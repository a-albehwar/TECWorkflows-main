import { DocumentCardTitleBase } from "office-ui-fabric-react/lib/components/DocumentCard/DocumentCardTitle.base";


export interface KPIReportRequestItem
{
  Title: string;
  Department:{
    Title:string;
    ID:number;
  }
  ID:number;
  Created:Date;
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
  AuthorId:number,
  KPIOwnerId:number,
  KPIOwner:{
      Title:string;
      ID:number;
  }
  PreSetDate:Date,
  TimePeriod:string,
  Period:string,
  KPIOwnerComments:string,
  AnalystComments:string,
  RequesterComments:string,

}



export interface TECDepartments
{
  Title: string,
  ID:number,
  Created:Date,
  KPIOwnerId:number,

  Author:{
    Title:string;
    ID:string;
  }
  AuthorId:number;
  KPIOwner:{
      Title:string;
      ID:number;
  }
}

  export interface KPIValueItem
  {
      Title:string,
      ID:number,
      KPIReport:{
          Title:string;
          ID:number;
      }
      System:string,
      Value:string,
      Comments:string,
      Year:string,
      TimePeriod:string,
      Department:{
          Title:string;
          ID:number;
      }
  }

  export interface KPIDocItem{
      Title:string,
      ID:number,
      KPIReport:
      {
          Title:string,
          ID:number,
          Period:number,

      }
      File:
      {
        ServerRelativeUrl:string;
        Name:string;
      }

  }

