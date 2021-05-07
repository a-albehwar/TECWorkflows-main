
    import { WebPartContext } from "@microsoft/sp-webpart-base";  
    import { SPHttpClient } from '@microsoft/sp-http';
      
    export interface ICctvInternalLegalMgrProps {  
      description: string;  
      context: WebPartContext;  
      spHttpClient: SPHttpClient;
      siteurl: string;
      weburl:string;
      pagecultureId:string;
    }  


