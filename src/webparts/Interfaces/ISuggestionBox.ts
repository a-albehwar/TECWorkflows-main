import { Guid } from "@microsoft/sp-core-library";

export interface SuggestionBoxListCols {
    Title: string;
    Suggestion_Type:string;
    Assigned_Dept_Comments:string;
    Dept_Head_Comments:string;
    Description:string;
    Innovation_Team_Review:string;
    User_Department:string;
    User_JobTitle:string;
    User_Name:string;
    AssignedDepartment:{
      Title:string;
      ID:number;
    }
    ID:number;
    Created:Date;
    Suggestion_Status:{
      Title:string;
      ID:string;
    }
      AssignedTo:{
        Title:string;
      }
      TaskUrl:{
        Description: string,
        Url: string
      }
    ContentType:{
        Name:string;
        Id:Guid;
    }
    Author:{
        Name:string;
        Title:string;
    }
    AuthorId:number;
}