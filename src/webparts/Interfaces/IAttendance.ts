export interface AttendanceItem {

  Title: string;
  ID:number;
  Created:Date;
  Department:{
    Title:string;
    ID:number;
  }
  Status:{
    Title:string;
    ID:number;
  } 
  AssignedTo:{
    Title:string;
  }
  
  AuthorId:number,
  Author:{
    Title:string;
    ID:string;
  }
  EmployeeID:string,
  ContactNumber:string,
  Email:string,
  
  TimeofAbsence:string,
  ReasonForRequest:string,
  DateofRequest:Date,
  DateOfIncident_To:Date,
  TimeOfIncident_To:string,
  AssignedToId:number,
  TimeofRequest:string,
  //SSSComments:string,
  //SnapshotAvailable:string,
  
}

export interface AttendanceSnapshotDoc
{
  Title:string;
  Description:string;
  RequestId:number;
  Request:{
    Title:string;
    ID:string;
  }
  IsActive:boolean;
}