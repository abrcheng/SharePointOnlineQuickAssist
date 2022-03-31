export interface IFile {
    ModifiedByEmail:string;
    ModifiedByName:string;
    ModifiedDate:string;
    Path:string;
    //Id:string;
    FileName:string;
    Library:string;
  }

  export interface IFiles
  {
      items:IFile[];
  }

  