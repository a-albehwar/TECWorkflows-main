export interface CCTVExternalList {
    Title: string;
    LocationFacilityOfIncident:string;
    EmailAddress:string;
    ID:number;
    Created:Date;
    DateOfIncident:Date;
    TimeOfIncident:string;
    DateOfIncident_To:Date,
    TimeOfIncident_To:string,
    ReasonForRequest:string;
   
    Status:{
      Title:string;
      ID:number;
    }
    Author:{
      Title:string;
      ID:string;
    }
    Mobile_x002d_Tel_x0020_No:string;
    RequesterName:string;
  }
  