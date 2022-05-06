
function execall(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sheet9");
  
  var Scenario = getcoldata("Scenario")
  var ID = getcoldata("ID")
  var callDate =	getcoldata("calldate")
  var callEnd =	getcoldata("callend")
  var callDuration = 	getcoldata("duration")
  var connectDuration = 	getcoldata("connect_duration")
  var caller = 	getcoldata("caller")
  var called  = 	getcoldata("called")
  var a_saddr = 	getcoldata("a_saddr")
  var b_saddr = 	getcoldata("b_saddr")
  var hold = 	getcoldata("hold")
  var legFlag = 	getcoldata("leg_flag")
  var sipResponseDesc = 	getcoldata("lastSIPresponseDesc")
  var audioDuration  = 	getcoldata("Audio Duration (Total length of call in seconds)")
  var ringing = 	getcoldata("Ringing Yes/No")
  var voiceStart = 	getcoldata("Voice Start Time (second)")
  var parties = 	getcoldata("Number of parties in call (#)")
  var clarity = 	getcoldata("Call Clarity (Echo, Noise, Mix Up, Silence)") 
  var comments = getcoldata("Comments");
  var sipcallerip = getcoldata("sipcallerip")
  var sipcalledip = getcoldata("sipcalledip")

  

  var data = sheet.getRange(2,1,1,199).getValues();
  var newdata = data[0];
  var mrgrng = sheet.getRange(1,1,sheet.getMaxRows, sheet.getMaxColumns()).getMergedRanges();
  console.log(mrgrng)

  
  
  //getting all column indexes
  var scenario_in = newdata.indexOf("Scenario");//this will print index of this number
  var id_in = newdata.indexOf("ID");
  var calldate_in = newdata.indexOf("calldate");
  var callEnd_in = newdata.indexOf("callend");
  var callDuration_in = newdata.indexOf("duration");
  var connectDuration_in = newdata.indexOf("connect_duration");
  var caller_in = newdata.indexOf("caller");
  var called_in = newdata.indexOf("called")
  var a_saddr_in	=	newdata.indexOf("a_saddr");
  var b_saddr_in	=	newdata.indexOf("b_saddr");
  var hold_in	=	newdata.indexOf("hold");
  var legFlag_in	=	newdata.indexOf("leg_flag");
  var sipResponseDesc_in	=	newdata.indexOf("lastSIPresponseDesc");
  var audioDuration_in	=	newdata.indexOf("Audio Duration (Total length of call in seconds)");
  var ringing_in	=	newdata.indexOf("Ringing Yes/No");
  var voiceStart_in	=	newdata.indexOf("Voice Start Time (second)");
  var parties_in	=	newdata.indexOf("Number of parties in call (#)");
  var clarity_in = newdata.indexOf("Call Clarity (Echo, Noise, Mix Up, Silence)")
  var comments_in = newdata.indexOf("Comments")
  var sipcallerip_in = newdata.indexOf("sipcallerip")
  var sipcalledip_in = newdata.indexOf("sipcalledip")

  console.log("index of scenario is = "+scenario_in)
  //var id_inb = id_in+1;
  
  

  var lastr = sheet.getLastRow();
  Logger.log(lastr)

  for(var i=3;i<lastr;i++){

    //getting values from all indexes
    var scenario2 = sheet.getRange(i,scenario_in+1).getValue();
    var id2 = sheet.getRange(i,id_in+1).getValue();
    var callDate2 = sheet.getRange(i,calldate_in+1).getValue();
    var callEnd2 = sheet.getRange(i,callEnd_in+1).getValue();
    var callDuration2 = sheet.getRange(i,callDuration_in+1).getValue();
    var connectDuration2 = sheet.getRange(i,connectDuration_in+1).getValue();
    var caller2 = sheet.getRange(i,caller_in+1).getValue();
    var called2 = sheet.getRange(i,called_in+1).getValue();
    var	a_saddr2	=	sheet.getRange(i,	a_saddr_in+1).getValue()
    var	b_saddr2	=	sheet.getRange(i,	b_saddr_in+1).getValue()
    var	hold_in2	=	sheet.getRange(i,	hold_in	+1).getValue()
    var	legFlag2	=	sheet.getRange(i,	legFlag_in+1).getValue()
    var	sipResponseDesc2	=	sheet.getRange(i,	sipResponseDesc_in+1).getValue()
    var	audioDuration2	=	sheet.getRange(i,	audioDuration_in	+1).getValue()
    var	ringing2	=	sheet.getRange(i,	ringing_in+1).getValue()
    var	voiceStart2	=	sheet.getRange(i,	voiceStart_in+1).getValue()
    var	parties2	=	sheet.getRange(i,	parties_in+1).getValue()
    var	clarity2	=	sheet.getRange(i,	clarity_in+1).getValue()
    var comment2 = sheet.getRange(i,comments_in+1).getValue();

    var	sipcallerip2	=	sheet.getRange(i,	sipcallerip_in+1).getValue()
    var sipcalledip2 = sheet.getRange(i,sipcalledip_in+1).getValue();
    //console.log("scenario = "+scenario2)

    
    

    if (scenario2 == null){
      sheet.getRange(i,scenario_in+1).setBackground("#cdcdcd")
      //if scenario is hold unhold, check for hold column 
    }
    if (id2 <=1){
      sheet.getRange(i,id_in+1).setBackground("#f2e06b")  
    }
    if (callDate2 <=1){
      sheet.getRange(i,calldate_in+1).setBackground("#f2e06b")      
    }
    if (callEnd2 <=1){
      sheet.getRange(i,callEnd_in+1).setBackground("#f2e06b")
    }
    if (callDuration2 <=1){
      sheet.getRange(i,callDuration_in+1).setBackground("#f2e06b")
    }
    if (connectDuration2 <=1){
      sheet.getRange(i,connectDuration_in+1).setBackground("#f2e06b")
      // sheet.getRange(i,a_saddr_in+1).setBackground("#ff4046")
      // sheet.getRange(i,	b_saddr_in+1).setBackground("#ff4046")
      sheet.getRange(i,comments_in+1).setNote("Audio Data May not Available ") && sheet.getRange(i,comments_in+1).setBackground("#ff4046");
      
    }
    if (caller2 <=1){
      sheet.getRange(i,caller_in+1).setBackground("#f2e06b") 
    }
    if (called2 <=1){
      sheet.getRange(i,called_in+1).setBackground("#f2e06b") 
    }
    if 	(	a_saddr2	<=1){
      	sheet.getRange(i,	a_saddr_in+1).setBackground("#f2e06b")	
    }
    if 	(	b_saddr2	<=1)	{
      	sheet.getRange(i,	b_saddr_in+1).setBackground("#f2e06b")
    }
    // if 	(	hold2	<=1)	{	
    //   sheet.getRange(i,	hold_in+1).setBackground("#f2e06b")
  	// }
    if 	(	legFlag2	<=1)	{
      	sheet.getRange(i,	legFlag_in+1).setBackground("#f2e06b")	
    }
    if 	(	sipResponseDesc2	!== '200 OK' || sipResponseDesc2	!==  "200 OK")	{
      	sheet.getRange(i,	sipResponseDesc_in+1).setBackground("#f2e06b")	
    }
    if 	(	audioDuration2	<=1)	{
      	sheet.getRange(i,	audioDuration_in+1).setBackground("#f2e06b")	
    }
    if 	(	ringing2	<=1)	{
      	sheet.getRange(i,	ringing_in+1).setBackground("#f2e06b")	
    }
    if 	(	voiceStart2	<=1)	{
      	sheet.getRange(i,	voiceStart_in+1).setBackground("#f2e06b")	
    }
    if 	(	parties2	<=1)	{
      	sheet.getRange(i,	parties_in+1).setBackground("#f2e06b")	
    }
    if 	(	clarity2	!= "Clear")	{
      	sheet.getRange(i,	clarity_in+1).setBackground("#f2e06b")	
    }
    if (callDuration2 >0 && connectDuration2 <=0  && sipResponseDesc2	==  "200 OK"){
      sheet.getRange(i,connectDuration_in+1).setBackground("#ff4046")
    }
     if (callDuration2 >0 && connectDuration2 <=0  && sipResponseDesc2	!==  "200 OK"){
      sheet.getRange(i,connectDuration_in+1).setBackground("#27cc53")
    }
    if(sipcallerip2 >0 && sipcalledip2>0 && sipcallerip2 === sipcalledip2 ){
      console.log("sip caller ip and called ip must not same")
      sheet.getRange(i,sipcallerip_in+1).setBackground("#d91646")
      sheet.getRange(i,sipcalledip_in+1).setBackground("#d91646")
      sheet.getRange(i,sipcallerip_in+1).setNote("Sip called and Caller IP must not same")
    }
    else{
      //Logger.log("executed")
    }
  }
  
}


function getcoldata(header) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sheet9"); // Change this according to your preferences
  var header;
  //const header = "ID"; // Change this according to your pareferences
  const values = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  const headers = values.shift();
  const columnIndex = headers.indexOf(header);
  const columnValues = values.map(row => row[columnIndex]);
  console.log(columnValues)
  //var field = SpreadsheetApp.getActiveSpreadsheet().getId()
  //console.log(field)
  return columnValues;  
}




const sS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("analysis-cpy"); 

function conditionalColors() {

  /* Filter the values to obtain true when the time changes */
  var range = getcoldata("calldate")
  const values = getcoldata("calldate")
    /* Evaluate end-date `B2:B` */
    //.getRange(range)
    // .filter(n => n[0] !== '')
    // .flat()
    .map(n => new Date(n).getTime())
     //
    .map((n, idx, crrArr) => idx === 0 ? true : crrArr[idx-1] !== crrArr[idx])  
    // 
    // Math.abs(crrArr[idx-1] - crrArr[idx]) > 5000 )

  /* defining the colors*/

  var data = sS.getRange(2,1,1,199).getValues();
  var newdata = data[0];
  //getting  column indexes
  var calldate_in = newdata.indexOf("calldate");
  var callEnd_in = newdata.indexOf("callend");
  console.log(calldate_in)
  
  console.log(getcoldata())
  const color1 = [204, 222, 188]
  const color2 = [238, 232, 170]
  let color = color1


  /* Loop to change the color every time we have a true */
  values.forEach((c, idx) => {
    if (c) {
      color = color === color1 ? color2 : color1
    }
    // if call date index is 4 it will add 1 =5, which means calldate  starts from 5th column   
    //provide a column range for calldate and call end
    sS.getRange(idx + 3, calldate_in+1).setBackgroundRGB(...color)
    sS.getRange(idx + 3, callEnd_in+1).setBackgroundRGB(...color)
  })
}
