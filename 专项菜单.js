const saas=SpreadsheetApp.getActiveSpreadsheet();
const sbnSheetSet=saas.getSheetByName("Sheet设置");
const sbnMenuSet=saas.getSheetByName("菜单设置");
const sbnTimingSet=saas.getSheetByName("定时设置");
const sbnHistoryData=saas.getSheetByName("HistoryData");
const sbnCurrentData=saas.getSheetByName("CurrentData");
const currentTimeZone=saas.getSpreadsheetTimeZone();
const scriptProperties = PropertiesService.getScriptProperties();
const scriptRunOwner=Session.getActiveUser().getEmail();
let GOOGLE_CHAT_WEBHOOK_LINK = "https://chat.googleapis.com/v1/spaces/AAAAyKrTKQ4/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=1q38g4i2uw8Sp6-xbb-Rv2_95nnVrhx16-feSjrm-os%3D";    /**张俊【GoogleChat网络钩子】**/

function onOpen(){
  let sbnMenuSetLr=sbnMenuSet.getLastRow();
  let arrMenus=[];
  let arrMenusData=[];
  let arrRunScript=[];
  if(sbnMenuSetLr>2){
    arrMenusData=sbnMenuSet.getRange(3,1,sbnMenuSetLr-2,4).getValues();
    arrMenus=arrMenusData.filter(v=>{return v[0]!=""&&v[2]=="是"});
    arrOpenRun=arrMenusData.filter(v=>{return v[0]!=""&&v[3]=="是"});
  }
  if(arrMenus.length>0){
    arrRunScript.push(null);
    for(let i=0;i<arrMenus.length;i++){
      arrRunScript.push({name:arrMenus[i][0],functionName:arrMenus[i][1]});
      if(i<arrMenus.length-1){
        arrRunScript.push(null)
      }
    }
  }
  saas.addMenu("专项菜单",arrRunScript);
}
