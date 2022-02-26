import { Guid } from "@microsoft/sp-core-library";
export interface IRestoreItem {  
    DeletedByEmail:string;
    DeletedByName:string;
    DeletedDate:Date;
    Path:string;
    Id:string;            
  }

  export interface IRestoreItems
  {
      items:IRestoreItem[];
  }

  