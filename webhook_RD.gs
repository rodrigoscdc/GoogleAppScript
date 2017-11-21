var SHEET_ID = '1QM5XPyhgeAiEPBkbDGb7a0GXwsos9Hm9GtBE4OrCWsI';
var TOKEN = 'asdfg';

function doPost(req){

  
  var sheets = SpreadsheetApp.openById('1QM5XPyhgeAiEPBkbDGb7a0GXwsos9Hm9GtBE4OrCWsI');
  var params = req.postData.contents;
  
  
  
  //sheets.getRangeByName('nome').getCell(numRow, 1).setValue(params);
  if (req.parameters.token == 'asdfg'){
    
    var json = JSON.parse(params);
    var leads = json['leads'];
    
    var idContact = leads[0]['id'];
    var nameContact = leads[0]['name']; 
    var email = leads[0]['email'];
    var mobilePhone = leads[0]['mobile_phone'];
    var personalPhone = leads[0]['personal_phone'];
    var city = leads[0]['city'];
    var state = leads[0]['state'];
    var interest = leads[0]['custom_fields']['Interesse'];
    var lastContact = leads[0]['last_conversion']['created_at'];
    var origin = leads[0]['last_conversion']['source'];
    var esteira = 0
    if (req.parameters.esteira != undefined){
      esteira = Number(req.parameters.esteira);
    }
    
    var numRow = getNextRow(sheets, email);
    if(numRow[0] != false){
    
      sheets.getRangeByName('nome').getCell(numRow[1], 1).setValue(nameContact);
      sheets.getRangeByName('id').getCell(numRow[1], 1).setValue(idContact);
      sheets.getRangeByName('email').getCell(numRow[1], 1).setValue(email);
      sheets.getRangeByName('celular').getCell(numRow[1], 1).setValue(mobilePhone);
      sheets.getRangeByName('tel_pessoas').getCell(numRow[1], 1).setValue(personalPhone);
      sheets.getRangeByName('cidade').getCell(numRow[1], 1).setValue(city);
      sheets.getRangeByName('estado').getCell(numRow[1], 1).setValue(state);
      sheets.getRangeByName('interesse').getCell(numRow[1], 1).setValue(interest);
      sheets.getRangeByName('last_contato').getCell(numRow[1], 1).setValue(lastContact);
      sheets.getRangeByName('origem_lista').getCell(numRow[1], 1).setValue(origin);
      if(origin == "mood-preco-lancamento-bortolini"){
        sheets.getSheetByName('DefinitivoMOOD').getRange(numRow[1], 1, 1, 12).setBackground('#ffac24');
      }
    } else if(origin == 'esteiramood_investidorcontato'){
      var horasContact = leads[0]['last_conversion']['content']['contatar prospect'];
      sheets.getRangeByName('last_conversao').getCell(numRow[1], 1).setValue(origin);
      sheets.getRangeByName('hora_contato').getCell(numRow[1], 1).setValue(horasContact);
      
      
      sheets.getSheetByName('DefinitivoMOOD').getRange(numRow[1], 1, 1, 12).setBackground('#00ff66');
      
    } else if(esteira != 0){
      var fluxo = ['Recebeu email 1', 'Recebeu email 2', 'Recebeu email 3'];
      sheets.getRangeByName('esteira').getCell(numRow[1], 1).setValue(fluxo[--esteira]);
    } else {
      
      sheets.getRangeByName('last_conversao').getCell(numRow[1], 1).setValue(origin);
      sheets.getSheetByName('DefinitivoMOOD').getRange(numRow[1], 1, 1, 12).setBackground('#ffac24');
    }
    return ContentService.createTextOutput("Success");
  }else{
    return ContentService.createTextOutput("Error");
  }
  
}


function getNextRow(sheets, value) {
  var confId = sheets.getRangeByName("email").getValues();
  var response = [];
  for (i in confId) {
    if(confId[i][0] == value){
      i++;    
      response.push(false);
      response.push(Number(i));
      return response;
      break;
    } else if(confId[i][0] == ""){
      i++;
      response.push(true);
      response.push(Number(i));
      return response;
      break;
    }
  }
}