export default class SPOQAHelper
{
    public static Show(id:string):void
    {        
         let sPOQASpinner:any = document.querySelector(`#${id}`);
         sPOQASpinner.style.display = "";
    }
 
    public static Hide(id:string):void
    {
     let sPOQASpinner:any = document.querySelector(`#${id}`);
     sPOQASpinner.style.display = "none";
    } 
    
    /*
      SPOQAErrorMessageBarContainer
      SPOQASuccessMessageBarContainer
      SPOQAWarningMessageBarContainer
      SPOQAInfoMessageBarContainer
    */
    public static ShowMessageBar(barType:string, message:string):void
    {
        // SPOQASuccessMessageBar
        const barTypes:String[] = ["Error", "Success", "Warning", "Info"]; 
        if(barTypes.indexOf(barType) >=0)
        {
            document.querySelector(`#SPOQA${barType}MessageBar`).innerHTML = `<span style="white-space:normal">${message}</span>`;
            SPOQAHelper.Show(`SPOQA${barType}MessageBarContainer`);
        }
        else
        {
            console.log(`Only accept bar type "Error", "Success", "Warning", "Info"`);
        }
    }

    public static ResetFormStaus():void
    {
        SPOQAHelper.Hide("SPOQAErrorMessageBarContainer");
        SPOQAHelper.Hide("SPOQASuccessMessageBarContainer");
        SPOQAHelper.Hide("SPOQAWarningMessageBarContainer");
        SPOQAHelper.Hide("SPOQAInfoMessageBarContainer");
    }

    public static ValidateEmail(email:string):boolean
    {
        const re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
        return re.test(email);
    }

    public static ValidateUrl(url:string):boolean
    {
        var pattern = new RegExp('^(https:\\/\\/)?'+ // protocol
        '((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|'+ // domain name
        '((\\d{1,3}\\.){3}\\d{1,3}))'+ // OR ip (v4) address
        '(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*'+ // port and path
        '(\\?[;&a-z\\d%_.~+=-]*)?'+ // query string
        '(\\#[-a-z\\d_]*)?$','i'); // fragment locator

        return pattern.test(url);
    }
}