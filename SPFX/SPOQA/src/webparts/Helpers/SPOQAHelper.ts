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
        // tslint:disable-next-line:no-function-expression
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

    public static JSONToCSVConvertor(JSONData:any, showLabel:boolean, fileName:string):void {
        //If JSONData is not an object then JSON.parse will parse the JSON string in an Object
        var arrData = typeof JSONData != 'object' ? JSON.parse(JSONData) : JSONData;
        
        var CSV = "";
    
        //This condition will generate the Label/Header
        if (showLabel) {
            var row = "";
            
            //This loop will extract the label from 1st index of on array
            for (var index in arrData[0]) {
                
                //Now convert each value to string and comma-seprated
                row += index + ',';
            }
    
            row = row.slice(0, -1);
            
            //append Label row with line break
            CSV += row + '\r\n';
        }
        
        //1st loop is to extract each row
        for (var i = 0; i < arrData.length; i++) {
            row = "";
            
            //2nd loop will extract each column and convert it in string comma-seprated
            for (index in arrData[i]) {
                row += '"' + arrData[i][index] + '",';
            }
    
            row.slice(0, row.length - 1);
            
            //add a line break after each row
            CSV += row + '\r\n';
        }
    
        if (CSV == '') {        
            alert("Invalid data");
            return;
        }          
        
        
        //Initialize file format you want csv or xls
        var uri = 'data:text/csv;charset=utf-8,%EF%BB%BF' + encodeURIComponent(CSV);
        
        // Now the little tricky part.
        // you can use either>> window.open(uri);
        // but this will not work in some browsers
        // or you will not get the correct file extension    
        
        //this trick will generate a temp <a /> tag
        var link = document.createElement("a");    
        link.href = uri;
        
        //set the visibility hidden so it will not effect on your web-layout
        link.style.display = "none";
        link.download = fileName + ".csv";
        
        //this part will append the anchor tag and remove it after automatic click
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }  

    // 2022-02-12T17:28:49.3538784+08:00
    
}