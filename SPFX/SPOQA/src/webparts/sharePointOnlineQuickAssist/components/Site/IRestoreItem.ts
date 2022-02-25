import { Guid } from "@microsoft/sp-core-library";
export interface IRestoreItem {  
    DeletedByEmail:string;
    DeletedByName:string;
    DeletedDate:Date;
    LeafName:string;
    DirName:string;
    Id:string;            
  }

  export interface IRestoreItems
  {
      items:IRestoreItem[];
  }

  