/** 
 * Sheet 1 has customer full name in column1 and account status in column 2 [valid, invalid]
 * 
 * Sheet 2 has customer full name in column 1 and a payment status in column 2 [current, past due]
 * 
 * Goal: Rconcile the two sheets
 *    - highlighting the customer full name in SHEET1 if the account is [0] valid and the payment in sheet 2 is [1] past due
*/
var _scriptProps = PropertiesService.getScriptProperties();

function displayPrompt(promptMessage){
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(promptMessage);
  console.log(result.getResponseText());
  if (result.getSelectedButton() == ui.Button.CLOSE){
    return "Closed"
  }else{
    return result.getResponseText()
    }
}

function startUp(){
  const questions = ["Enter the Account Sheet's Name", "Enter the Payment's Sheet Name", "Data range of account sheet in A1 format? (ex. A1:B30)", "Data range of payment sheet in A1 format? (ex. A1:B30)"];
  const theKeys = ["accountSheetName", "paymentSheetName", "accountRange", "paymentRange"];

  let count = 0
  questions.forEach(ques =>{
    let result = displayPrompt(ques);
    if (result =="Closed"){console.log("skip")}else{
      _scriptProps.setProperty(theKeys[count], result);
    }
    count++
  })
  // ask if you want to have range prompted again gets it own because it's a yes or no
  const ui = SpreadsheetApp.getUi();
  let res = ui.prompt("Ask for ranges again when you run program? \n Click the Yes or No Button", ui.ButtonSet.YES_NO);
  if (res.getSelectedButton() == ui.Button.YES) {
    _scriptProps.setProperty("askForRanges", true);} else if (res.getSelectedButton() == ui.Button.NO){
      _scriptProps.setProperty("askForRanges", false);} else{return}
  // ask about highlight color
  let highlight = displayPrompt("Enter chosen highlight color \n #ffffff format or a css color (blue, green, white, etc");
  if (highlight == "Close"){let skip = "";}else{
    _scriptProps.setProperty("highlightColor", highlight);
  }
}
  




function setAccntSheetName(){
  let name = displayPrompt("Enter Account Sheet Name");
  if (name == "Closed"){return}
  else{_scriptProps.setProperty("accountSheetName", name);
  SpreadsheetApp.getUi().alert("Program will now search for account validity status in: " + name);
  }
}
function setPymntSheetName(){
  let name = displayPrompt("Enter Payment Sheet Name");
  if (name == "Closed"){return}
  else{_scriptProps.setProperty("paymentSheetName", name);
  SpreadsheetApp.getUi().alert("Program will now search for payment status in: " + name);
  }
}
function changeAskForR(){
  if(_scriptProps.getProperty("askForRanges") == true){
    _scriptProps.setProperty("askForRanges", false);
    
  }else{
    _scriptProps.setProperty("askForRanges", true);
  }
}
function setPayRange(){
  let r = displayPrompt("Data range of payment sheet in A1 format? (ex. A1:B30)");
  _scriptProps.setProperty("paymentRange", r);
}
function setAccountRange(){
  let r = displayPrompt("Data range of account sheet in A1 format? (ex. A1:B30)");
  _scriptProps.setProperty("accountRange", r);
}
function changeHighlightColor(){
  let r = displayPrompt("New highlight color? \n #ffffff or css color");
  _scriptProps.setProperty("highlightColor", r)

}


function doTheThing(){
  // get spreadsheetObj
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // getfirst sheet
  const accountSheet = ss.getSheetByName(_scriptProps.getProperty("accountSheetName"));
  const paymentSheet = ss.getSheetByName(_scriptProps.getProperty("paymentSheetName"));
  let accntRange = ""
  let pymntRange = ""
  if (_scriptProps.getProperty("askForRanges") == true){
    accntRange = displayPrompt("Data range of account sheet in A1 format? (ex. A1:B30)");
    pymntRange = displayPrompt("Data range of payment sheet in A1 format? (ex. A1:B30)");
  }else{
    accntRange = _scriptProps.getProperty("accountRange")
    pymntRange = _scriptProps.getProperty("paymentRange")
  }


  const sh1Range = accountSheet.getRange(accntRange).getValues();
  const sh2Range = paymentSheet.getRange(pymntRange).getValues();
  let accountsArray = [];


  const nameCol = accntRange.substring(0,1);

  const nameRow = accntRange.substring(1);

  let count = parseInt(nameRow, 10);
  sh1Range.forEach(sh1info=> {
    sh2Range.forEach(sh2info=> {
      if (sh1info[0] == sh2info[0]){
        let customer = {
          name: sh1info[0],
          account_status: sh1info[1].toLowerCase(),
          payment_status: sh2info[1].toLowerCase(),
          sh1CustomerLocation: nameCol + count
        };
        accountsArray.push(customer);
      }
      
    })
    count++
  })

  accountsArray.forEach(cust => {
    if(cust.account_status == "valid" && cust.payment_status == "past due"){
      accountSheet.getRange(cust.sh1CustomerLocation).setBackground(_scriptProps.getProperty("highlightColor"));
    } else{
      accountSheet.getRange(cust.sh1CustomerLocation).setBackground(null);
    }
  })

}

function helpme(){
  const message = "Click This One First, Then Don't Worry About It \n --> Sets up the program to read your data. Answer each question at least once. \n Highlight Valid Past Due Customers \n --> The Program \n Buttons past that let you reset the parameters to your liking";
  ui = SpreadsheetApp.getUi();
  ui.alert(message);

}

function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Special")
    .addItem("Click This One First, Then Don't Worry About It", "startUp")
    .addItem("Highlight Valid Past Due Customers", "doTheThing")
    .addItem("Set Account Sheet Name", "setAccntSheetName")
    .addItem("Set Payment Sheet Name", "setPymntSheetName")
    .addItem("Set Date Range [Account]", "setAccountRange")    
    .addItem("Set Data Range [Payment]", "setPayRange")
    .addItem("Change Highlight Color", "changeHighlightColor")
    .addItem("Toggle Ask For Range On/Off", "changeAskForR")
    .addItem("Help","helpme")
    .addToUi();
}

