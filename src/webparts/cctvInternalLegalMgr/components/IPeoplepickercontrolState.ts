
    import { MessageBarType } from 'office-ui-fabric-react';  
      
    export interface IPeoplepickercontrolState {  
        title: string;  
        users: number[];  
        showMessageBar: boolean;  
        messageType?: MessageBarType;  
        message?: string;  
    }  
