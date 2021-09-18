export default class SPOQASpinner
{
   public static Show(label:string)
   {
        document.querySelector("#SPOQASpinner >.ms-Spinner-label").innerHTML = label;
        let sPOQASpinner:any = document.querySelector("#SPOQASpinner");
        sPOQASpinner.style.display = "";
   }

   public static Hide()
   {
    let sPOQASpinner:any = document.querySelector("#SPOQASpinner");
    sPOQASpinner.style.display = "none";
   }
}