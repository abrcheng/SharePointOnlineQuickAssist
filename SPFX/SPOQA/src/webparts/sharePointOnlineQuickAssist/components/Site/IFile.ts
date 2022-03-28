export interface IFile {
    ModifiedByEmail:string;
    ModifiedByName:string;
    ModifiedDate:string;
    Path:string;
    Id:string;
    FileName:string;
  }

  export interface IFiles
  {
      items:IFile[];
  }

  