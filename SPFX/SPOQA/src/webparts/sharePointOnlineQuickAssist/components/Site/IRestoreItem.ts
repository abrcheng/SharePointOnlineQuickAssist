export interface IRestoreItem {  
    DeletedByEmail:string;
    DeletedByName:string;
    DeletedDate:Date;
    Path:string;
    Id:string;   
    Existing:boolean; 
  }

  export interface IRestoreItems
  {
      items:IRestoreItem[];
  }

  