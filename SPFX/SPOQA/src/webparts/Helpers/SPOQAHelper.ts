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

    public static ParseQueryString(queryString?: string): any {
        // if the query string is NULL or undefined
        if (!queryString) {
            queryString = window.location.search.substring(1);
        }
        const params = {};
        const queries = queryString.split("&");
        queries.forEach((indexQuery: string) => {
            const indexPair = indexQuery.split("=");
            const queryKey = decodeURIComponent(indexPair[0]);
            const queryValue = decodeURIComponent(indexPair.length > 1 ? indexPair[1] : "");
            params[queryKey] = queryValue;
        });
        return params;
    }

    public static GenerateUUID():string { 
        var d = new Date().getTime();//Timestamp
        var d2 = ((typeof performance !== 'undefined') && performance.now && (performance.now()*1000)) || 0;//Time in microseconds since page-load or 0 if unsupported
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
            var r = Math.random() * 16;//random number between 0 and 16
            if(d > 0){//Use timestamp until depleted
                r = (d + r)%16 | 0;
                d = Math.floor(d/16);
            } else {//Use microseconds since page-load if supported
                r = (d2 + r)%16 | 0;
                d2 = Math.floor(d2/16);
            }
            return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
        });
    }
}