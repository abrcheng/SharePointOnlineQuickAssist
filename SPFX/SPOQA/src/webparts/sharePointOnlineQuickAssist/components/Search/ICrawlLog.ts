export interface ICrawlLog {  
    FullUrl:string;
    IsDeleted:string;
    ExclusionReason:string;
    DeleteReason:string;
    ErrorCode:string;  
    TimeStamp:string; 
    ErrorDesc:string;   
    DeletePending:string;      
  }

  export interface ICrawlLogs
  {
      items:ICrawlLog[];
  }
