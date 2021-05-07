import { Guid } from "@microsoft/sp-core-library";

export interface RequestFormListCols {
    Title: string;
    Departments:{
      Title:string;
      ID:number;
    }
    ID:number;
    Created:Date;
    Status:string;
       
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

    TECDepartment:{
      Title:string;
      ID:number;
    }
    WorkflowStatus:string;
    // checkbox marcomm
    Doesrequestincludecommunicationc:string;
    Does_x0020_the_x0020_request_x00:string;
    Doestherequestinvolveexternalorc:string;
    IsRequestValid:string;
    DoesRequestMeetCriteria:string;
    IsthebudgetoriginatingfromMarCom:string;
    //Survey Fields...
    TypeOfSurvey:string;
    SurveyStartDate:Date;
    SurveyEndDate:Date;
    PurposeOfSurvey:string;
    WhoIsSurveyFor:string;  
    DoYouRequireSurveyReport:string,
    
    


    //Videography/Photography
    TypeOfShoot:string,
    DoesTheDepartmentHaveaBudgetForT:string,
    WhichDepartmentBudgetWillThisCom:string,
    BudgetAmount:string,
    PurposeOfShoot:string,
    DateOfShoot:Date,
    DateofShootTo:Date,
    Location:string,
    StyleOfShoot:string,
    WhereWillThisBePublished:string,
    IsAcastRequired:string,

    //Social Media type
    SocialMediaType:string, //multi select..
    DateOfPost:Date,
    DurationOfSponsoredAd:string,
    DateOfEvent:Date,
    DateOfInfluencerEngagement:Date,
    LocationOfEvent:string,
    Platforms:string[], //multi
    TypeOfEvent:string,

    //Event Fields
    EventDateTime:Date,
    EventDuration:string,
    Requirements:string[],
    IfDecorativePleaseSpecify:string,
    If_x0020_Other_x0020_Please_x002:string,
    TimeOfEvent:string,

    //Design And Production
    TypeOfDesign:string[],
    SpecifyDecorativeElements:string,
    SpecifyCollateral:string,
    Size:string,
    SupportingTextContentLanguage:string,
    IllustrationReference:string,
    DateOfDelivery:Date,
    WillYouRequireProduction:string,
    Quantity:string,
    PrefferedMaterial:string,
    //InstallationDeadline:string,
    InstallationDeadline:Date,

    //Media 
    TextContent:string,
    NewspaperMediaPlatformPrefrences:string,
    PublishDate:Date,
    ol_Department:string,

    //Content Type Form
    ContentTypeForm:string,
    PleaseProvideAllDetailsForConten:string,
    LengthOfContent:string,
    DoYouRequireBilingualContent:string,
    Language:string[],
    DeadLine_x0020_Content_x0020_Cre:Date,


    AnyAdditionalDetails:string;

}
//   public SSSActionUrl:string="/Pages/cctv/SSSIntCCTVReq.aspx";
//   public LegalMgrActionUrl:string="/Pages/cctv/CCTVLegalMgr.aspx";
//   public CCTVInternalTaskUrl:string="/Lists/CCTV_Internal_Incident";
// }