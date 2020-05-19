function levelCalculator() {
  
  var level = "";
  
  var ss = SpreadsheetApp.openById("1jYv5VzmQIHE572kRBmLouLXMCBddzGyy4ydWGehzLWQ"); //Handicap Group A
  var sheets = ss.getSheets();
  
  for(i = 2; i <= 45; i++) {
    
    var str = "D" + String(i);
    
    var pin = sheets[2].getRange(str).getValue();
  
    var url = "https://www.europeangodatabase.eu/EGD/GetPlayerDataByPIN.php?pin=" + pin;
  
    var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
  
    var json = response.getContentText();
    var list = JSON.parse(json);
    
    str2 = "P" + String(i);
    
    if(list.Gor < 150) level = "20k";
    else if(list.Gor < 250) level = "19k";
    else if(list.Gor < 350) level = "18k";
    else if(list.Gor < 450) level = "17k";
    else if(list.Gor < 550) level = "16k";
    else if(list.Gor < 650) level = "15k";
    else if(list.Gor < 750) level = "14k";
    else if(list.Gor < 850) level = "13k";
    else if(list.Gor < 950) level = "12k";
    else if(list.Gor < 1050) level = "11k";
    else if(list.Gor < 1150) level = "10k";
    else if(list.Gor < 1250) level = "9k";
    else if(list.Gor < 1350) level = "8k";
    else if(list.Gor < 1450) level = "7k";
    else if(list.Gor < 1550) level = "6k";
    else if(list.Gor < 1650) level = "5k";
    else if(list.Gor < 1750) level = "4k";
    else if(list.Gor < 1850) level = "3k";
    else if(list.Gor < 1950) level = "2k";
    else if(list.Gor < 2050) level = "1k";
    else if(list.Gor < 2150) level = "1d";
    else if(list.Gor < 2250) level = "2d";
    else if(list.Gor < 2350) level = "3d";
    else if(list.Gor < 2450) level = "4d";
    else if(list.Gor < 2550) level = "5d";
  
    sheets[2].getRange(str2).setValue(level);
    
  }
  
  var ss = SpreadsheetApp.openById("1neFq-utdwVBPwfE9BKVxiAiORw8ablCKBQb3nuVfOtM"); //Handicap Group B
  var sheets = ss.getSheets();
  
  for(i = 2; i <= 45; i++) {
    
    var str = "D" + String(i);
    
    var pin = sheets[2].getRange(str).getValue();
  
    var url = "https://www.europeangodatabase.eu/EGD/GetPlayerDataByPIN.php?pin=" + pin;
  
    var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
  
    var json = response.getContentText();
    var list = JSON.parse(json);
    
    str2 = "P" + String(i);
  
    if(list.Gor < 150) level = "20k";
    else if(list.Gor < 250) level = "19k";
    else if(list.Gor < 350) level = "18k";
    else if(list.Gor < 450) level = "17k";
    else if(list.Gor < 550) level = "16k";
    else if(list.Gor < 650) level = "15k";
    else if(list.Gor < 750) level = "14k";
    else if(list.Gor < 850) level = "13k";
    else if(list.Gor < 950) level = "12k";
    else if(list.Gor < 1050) level = "11k";
    else if(list.Gor < 1150) level = "10k";
    else if(list.Gor < 1250) level = "9k";
    else if(list.Gor < 1350) level = "8k";
    else if(list.Gor < 1450) level = "7k";
    else if(list.Gor < 1550) level = "6k";
    else if(list.Gor < 1650) level = "5k";
    else if(list.Gor < 1750) level = "4k";
    else if(list.Gor < 1850) level = "3k";
    else if(list.Gor < 1950) level = "2k";
    else if(list.Gor < 2050) level = "1k";
    else if(list.Gor < 2150) level = "1d";
    else if(list.Gor < 2250) level = "2d";
    else if(list.Gor < 2350) level = "3d";
    else if(list.Gor < 2450) level = "4d";
    else if(list.Gor < 2550) level = "5d";
  
    sheets[2].getRange(str2).setValue(level);
    
  }
  
  var ss = SpreadsheetApp.openById("1J9s0tq8MKPOckoaJCU9oUgu6TyHuI_a-ydgtQEEt_QQ"); //Group A
  var sheets = ss.getSheets();
  
  for(i = 2; i <= 53; i++) {
    
    var str = "D" + String(i);
    
    var pin = sheets[2].getRange(str).getValue();
  
    var url = "https://www.europeangodatabase.eu/EGD/GetPlayerDataByPIN.php?pin=" + pin;
  
    var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
  
    var json = response.getContentText();
    var list = JSON.parse(json);
    
    str2 = "P" + String(i);
  
    if(list.Gor < 150) level = "20k";
    else if(list.Gor < 250) level = "19k";
    else if(list.Gor < 350) level = "18k";
    else if(list.Gor < 450) level = "17k";
    else if(list.Gor < 550) level = "16k";
    else if(list.Gor < 650) level = "15k";
    else if(list.Gor < 750) level = "14k";
    else if(list.Gor < 850) level = "13k";
    else if(list.Gor < 950) level = "12k";
    else if(list.Gor < 1050) level = "11k";
    else if(list.Gor < 1150) level = "10k";
    else if(list.Gor < 1250) level = "9k";
    else if(list.Gor < 1350) level = "8k";
    else if(list.Gor < 1450) level = "7k";
    else if(list.Gor < 1550) level = "6k";
    else if(list.Gor < 1650) level = "5k";
    else if(list.Gor < 1750) level = "4k";
    else if(list.Gor < 1850) level = "3k";
    else if(list.Gor < 1950) level = "2k";
    else if(list.Gor < 2050) level = "1k";
    else if(list.Gor < 2150) level = "1d";
    else if(list.Gor < 2250) level = "2d";
    else if(list.Gor < 2350) level = "3d";
    else if(list.Gor < 2450) level = "4d";
    else if(list.Gor < 2550) level = "5d";
  
    sheets[2].getRange(str2).setValue(level);
    
  }
  
  var ss = SpreadsheetApp.openById("1zLHxxP20goVDXN39FHGihQDdXcsbZbToVJLDJwwqMvU"); //Group B
  var sheets = ss.getSheets();
  
  for(i = 2; i <= 53; i++) {
    
    var str = "D" + String(i);
    
    var pin = sheets[2].getRange(str).getValue();
  
    var url = "https://www.europeangodatabase.eu/EGD/GetPlayerDataByPIN.php?pin=" + pin;
  
    var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
  
    var json = response.getContentText();
    var list = JSON.parse(json);
    
    str2 = "P" + String(i);
  
    if(list.Gor < 150) level = "20k";
    else if(list.Gor < 250) level = "19k";
    else if(list.Gor < 350) level = "18k";
    else if(list.Gor < 450) level = "17k";
    else if(list.Gor < 550) level = "16k";
    else if(list.Gor < 650) level = "15k";
    else if(list.Gor < 750) level = "14k";
    else if(list.Gor < 850) level = "13k";
    else if(list.Gor < 950) level = "12k";
    else if(list.Gor < 1050) level = "11k";
    else if(list.Gor < 1150) level = "10k";
    else if(list.Gor < 1250) level = "9k";
    else if(list.Gor < 1350) level = "8k";
    else if(list.Gor < 1450) level = "7k";
    else if(list.Gor < 1550) level = "6k";
    else if(list.Gor < 1650) level = "5k";
    else if(list.Gor < 1750) level = "4k";
    else if(list.Gor < 1850) level = "3k";
    else if(list.Gor < 1950) level = "2k";
    else if(list.Gor < 2050) level = "1k";
    else if(list.Gor < 2150) level = "1d";
    else if(list.Gor < 2250) level = "2d";
    else if(list.Gor < 2350) level = "3d";
    else if(list.Gor < 2450) level = "4d";
    else if(list.Gor < 2550) level = "5d";
  
    sheets[2].getRange(str2).setValue(level);
    
  }
  
}
