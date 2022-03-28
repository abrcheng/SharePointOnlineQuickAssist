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

  