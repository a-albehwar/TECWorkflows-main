

import { Guid } from "@microsoft/sp-core-library";
export interface CCTVInternalList {
    Title: string;
    
    ID:number;
    Created:Date;
    Status:string;
    Modified:Date;
    AssignedTo:string;
    InitiatedBy:{
        Title:string;
        ID:string;
      }
    Author:{
      Title:string;
      ID:string;
    }
   
  Comments:string;
  }