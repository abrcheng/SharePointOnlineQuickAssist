// Helper class for Remedy Steps
export class RemedyHelper
{
    private static remedyStyle = "color:black";

    // Build the html string according to the remedySteps 
    public static ShowRemedySteps(remedySteps:any[])
    {    
        var remedyHtml=`<br/><label class="ms-Label" style='${RemedyHelper.remedyStyle};font-size:14px;font-weight:bold'>Remedy Steps:</label><br/>`;
        // Dispaly remedy steps
        remedySteps.forEach(step=>{
            var message =step.message;
            if(step.message[step.message.length-1] ==".")
            {
                message = message.substr(0, step.message.length-1);                
            }
            var fixpage = "";
            if(step.url)
            {
                fixpage = ` can be fixed in <a href='${step.url}' target='_blank'>this page</a>`;
            }
            remedyHtml+=`<div style='${this.remedyStyle};margin-left:20px'>${message}${fixpage}.</div>`;
        }); 

        return remedyHtml; 
    }
} 