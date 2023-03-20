function Test_ExtractStock_Button () {
  var ExtractStock_URL = "https://www.set.or.th/api/set/index/agri/composition?lang=th";
  var JSON_URLX = ExtractStock_URL;
  NotifyMessage("agri");
  try {
    var JSON_URL_ResponseX = UrlFetchApp.fetch(JSON_URLX);
  }
   catch (error) {
    //console.log(error.name + "：" + error.message);
    NotifyMessage(error.name + "：" + error.message);
    return;
  }

if (JSON_URL_ResponseX.getResponseCode() === 200) {

      var JSON_ContentX = JSON_URL_ResponseX.getContentText();
      console.log(JSON_ContentX);
  
      var requestObjX = JSON.parse(JSON_ContentX);
      console.log(requestObjX.composition.subIndices.length);

      var Industry_mai_Check = requestObjX.composition.subIndices;
      if (Industry_mai_Check == null) {
        
        //-----Begining mai------------

          var stockInfos_Count_M = requestObjX.composition.stockInfos.length;
          console.log("stockInfos_Count_M: " + stockInfos_Count_M);
          //NotifyMessage("stockInfos_Count_M: " + stockInfos_Count_M);

          var StockArrayM = new Array (stockInfos_Count_M);
          for (var mLoop = 0; mLoop < stockInfos_Count_M; mLoop++) {
            StockArrayM[mLoop] = new Array (9);

            StockArrayM[mLoop][0] = requestObjX.composition.stockInfos[mLoop].marketName;
            StockArrayM[mLoop][1] = requestObjX.composition.stockInfos[mLoop].sectorName;
            StockArrayM[mLoop][2] = requestObjX.composition.stockInfos[mLoop].symbol;
            
            var Temp_marketDateTime_mai = requestObjX.composition.stockInfos[mLoop].marketDateTime.substring(8,10) + "/" + requestObjX.composition.stockInfos[mLoop].marketDateTime.substring(5,7) + "/" + requestObjX.composition.stockInfos[mLoop].marketDateTime.substring(0,4);

            StockArrayM[mLoop][3] = Temp_marketDateTime_mai;
            StockArrayM[mLoop][4] = requestObjX.composition.stockInfos[mLoop].last;
            if (StockArrayM[mLoop][4] == null) StockArrayM[mLoop][4] = "-";
            StockArrayM[mLoop][5] = requestObjX.composition.stockInfos[mLoop].change;
            if (StockArrayM[mLoop][5] == null) {StockArrayM[mLoop][5] = "-";} else StockArrayM[mLoop][5] = requestObjX.composition.stockInfos[mLoop].change.toFixed(2);
            StockArrayM[mLoop][6] = requestObjX.composition.stockInfos[mLoop].percentChange;
            if (StockArrayM[mLoop][6] == null) {StockArrayM[mLoop][6] = "-";} else StockArrayM[mLoop][6] = requestObjX.composition.stockInfos[mLoop].percentChange.toFixed(2);
            StockArrayM[mLoop][7] = (requestObjX.composition.stockInfos[mLoop].totalValue/1000);
            StockArrayM[mLoop][8] = StockArrayM[mLoop][2] + StockArrayM[mLoop][3];
            StockArrayM[mLoop][3] = "'" + Temp_marketDateTime_mai;
        }
        sheet_Draft_JSON.getRange(sheet_Draft_JSON.getLastRow()+1,1,stockInfos_Count_M,9).setValues(StockArrayM);
        SpreadsheetApp.flush();

        //-----Ending mai------------

      } else {

        //-----Begining SET------------
        
        var subIndices_Count = requestObjX.composition.subIndices.length;
        var TotalStock_subIndices_Count = 0;
        var StockArrayX = new Array (subIndices_Count);
        for (var kLoop = 0; kLoop < subIndices_Count; kLoop++) {

          var stockInfos_Count = requestObjX.composition.subIndices[kLoop].stockInfos.length;
          TotalStock_subIndices_Count = TotalStock_subIndices_Count + stockInfos_Count;
          StockArrayX[kLoop] = new Array (stockInfos_Count);

          for (var LLoop = 0; LLoop < stockInfos_Count; LLoop++) {
            StockArrayX[kLoop][LLoop] = new Array (9);

            StockArrayX[kLoop][LLoop][0] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].marketName;
            StockArrayX[kLoop][LLoop][1] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].sectorName;
            StockArrayX[kLoop][LLoop][2] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].symbol;

            var Temp_marketDateTime = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].marketDateTime.substring(8,10) + "/" + requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].marketDateTime.substring(5,7) + "/" + requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].marketDateTime.substring(0,4);

            StockArrayX[kLoop][LLoop][3] = Temp_marketDateTime;
            StockArrayX[kLoop][LLoop][4] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].last;
            if (StockArrayX[kLoop][LLoop][4] == null) StockArrayX[kLoop][LLoop][4] = "-";
            StockArrayX[kLoop][LLoop][5] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].change;
            if (StockArrayX[kLoop][LLoop][5] == null) {StockArrayX[kLoop][LLoop][5] = "-";} else {
              StockArrayX[kLoop][LLoop][5] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].change.toFixed(2);
            }
            StockArrayX[kLoop][LLoop][6] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].percentChange;
            if (StockArrayX[kLoop][LLoop][6] == null) {StockArrayX[kLoop][LLoop][6] = "-";} else {
              StockArrayX[kLoop][LLoop][6] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].percentChange.toFixed(2);
            }
            StockArrayX[kLoop][LLoop][7] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].totalValue/1000;
            StockArrayX[kLoop][LLoop][8] = StockArrayX[kLoop][LLoop][2] + StockArrayX[kLoop][LLoop][3];
            StockArrayX[kLoop][LLoop][3] = "'" + Temp_marketDateTime;
          }
        sheet_Draft_JSON.getRange(sheet_Draft_JSON.getLastRow()+1,1,stockInfos_Count,9).setValues(StockArrayX[kLoop]);
        SpreadsheetApp.flush();
        }
          //console.log(sheet_Draft_JSON.getLastRow());
          //console.log(TotalStock_subIndices_Count);

          //-----Ending SET------------

      }

      /*
      var Industry_Count = 0;
      var StockArray = new Array (Industry_Count);
      for (var jLoop = 0; jLoop < requestObj.length; jLoop++) {
        if (requestObj[jLoop].level == "INDUSTRY") {
          StockArray[Industry_Count] = requestObj[jLoop].symbol;
          if (Market_Param == "mai") { StockArray[Industry_Count] = StockArray[Industry_Count] + "-m";}
          Industry_Count++;
        }
      }
      */

      //console.log(StockArrayX);

      //Logger.log("ExtractStock_Button has finished.");
      //SpreadsheetApp.flush();
 
  } else {
      console.log("Error : " + JSON_URL_responseX.getResponseCode());
      NotifyMessage("Error : " + JSON_URL_responseX.getResponseCode());
  }


}
//------------------------------------------------------------------------------------------------------------------
function ExtractStock_Button (ExtStock_Param) {

  if (ExtStock_Param == null) { var ExtStock_Param = "agro";
  }
  var ExtractStock_URL = "https://www.set.or.th/api/set/index/" + ExtStock_Param + "/composition";
  var JSON_URLX = ExtractStock_URL;
  //console.log(ExtStock_Param);
  NotifyMessage(ExtStock_Param);

  try {
    var JSON_URL_ResponseX = UrlFetchApp.fetch(JSON_URLX);
  }
  
  catch (error) {
    //console.log(error.name + "：" + error.message);
    NotifyMessage(error.name + "：" + error.message);
    return;
  }
  //console.log("getResponseCode ：" + JSON_URL_response.getResponseCode());

  if (JSON_URL_ResponseX.getResponseCode() === 200) {

      var JSON_ContentX = JSON_URL_ResponseX.getContentText();
      //console.log(JSON_ContentX);
  
      var requestObjX = JSON.parse(JSON_ContentX);
      //console.log(requestObjX.composition.subIndices.length);

      var Industry_mai_Check = requestObjX.composition.subIndices;
      if (Industry_mai_Check == null) {
        
        //-----Begining mai------------

          var stockInfos_Count_M = requestObjX.composition.stockInfos.length;
          //console.log("stockInfos_Count_M: " + stockInfos_Count_M);
          //NotifyMessage("stockInfos_Count_M: " + stockInfos_Count_M);

          var StockArrayM = new Array (stockInfos_Count_M);
          for (var mLoop = 0; mLoop < stockInfos_Count_M; mLoop++) {
            StockArrayM[mLoop] = new Array (9);

            StockArrayM[mLoop][0] = requestObjX.composition.stockInfos[mLoop].marketName;
            StockArrayM[mLoop][1] = requestObjX.composition.stockInfos[mLoop].sectorName;
            StockArrayM[mLoop][2] = requestObjX.composition.stockInfos[mLoop].symbol;
            
            var Temp_marketDateTime_mai = requestObjX.composition.stockInfos[mLoop].marketDateTime.substring(8,10) + "/" + requestObjX.composition.stockInfos[mLoop].marketDateTime.substring(5,7) + "/" + requestObjX.composition.stockInfos[mLoop].marketDateTime.substring(0,4);

            StockArrayM[mLoop][3] = Temp_marketDateTime_mai;
            StockArrayM[mLoop][4] = requestObjX.composition.stockInfos[mLoop].last;
            if (StockArrayM[mLoop][4] == null) StockArrayM[mLoop][4] = "-";
            StockArrayM[mLoop][5] = requestObjX.composition.stockInfos[mLoop].change;
            if (StockArrayM[mLoop][5] == null) {StockArrayM[mLoop][5] = "-";} else StockArrayM[mLoop][5] = requestObjX.composition.stockInfos[mLoop].change.toFixed(2);
            StockArrayM[mLoop][6] = requestObjX.composition.stockInfos[mLoop].percentChange;
            if (StockArrayM[mLoop][6] == null) {StockArrayM[mLoop][6] = "-";} else StockArrayM[mLoop][6] = requestObjX.composition.stockInfos[mLoop].percentChange.toFixed(2);
            StockArrayM[mLoop][7] = (requestObjX.composition.stockInfos[mLoop].totalValue/1000);
            StockArrayM[mLoop][8] = StockArrayM[mLoop][2] + StockArrayM[mLoop][3];
            StockArrayM[mLoop][3] = "'" + Temp_marketDateTime_mai;
        }
        sheet_Draft_JSON.getRange(sheet_Draft_JSON.getLastRow()+1,1,stockInfos_Count_M,9).setValues(StockArrayM);
        SpreadsheetApp.flush();

        //-----Ending mai------------

      } else {

        //-----Begining SET------------
        
        var subIndices_Count = requestObjX.composition.subIndices.length;
        var TotalStock_subIndices_Count = 0;
        var StockArrayX = new Array (subIndices_Count);
        for (var kLoop = 0; kLoop < subIndices_Count; kLoop++) {

          var stockInfos_Count = requestObjX.composition.subIndices[kLoop].stockInfos.length;
          TotalStock_subIndices_Count = TotalStock_subIndices_Count + stockInfos_Count;
          StockArrayX[kLoop] = new Array (stockInfos_Count);

          for (var LLoop = 0; LLoop < stockInfos_Count; LLoop++) {
            StockArrayX[kLoop][LLoop] = new Array (9);

            StockArrayX[kLoop][LLoop][0] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].marketName;
            StockArrayX[kLoop][LLoop][1] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].sectorName;
            StockArrayX[kLoop][LLoop][2] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].symbol;

            var Temp_marketDateTime = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].marketDateTime.substring(8,10) + "/" + requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].marketDateTime.substring(5,7) + "/" + requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].marketDateTime.substring(0,4);

            StockArrayX[kLoop][LLoop][3] = Temp_marketDateTime;
            StockArrayX[kLoop][LLoop][4] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].last;
            if (StockArrayX[kLoop][LLoop][4] == null) StockArrayX[kLoop][LLoop][4] = "-";
            StockArrayX[kLoop][LLoop][5] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].change;
            if (StockArrayX[kLoop][LLoop][5] == null) {StockArrayX[kLoop][LLoop][5] = "-";} else {
              StockArrayX[kLoop][LLoop][5] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].change.toFixed(2);
            }
            StockArrayX[kLoop][LLoop][6] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].percentChange;
            if (StockArrayX[kLoop][LLoop][6] == null) {StockArrayX[kLoop][LLoop][6] = "-";} else {
              StockArrayX[kLoop][LLoop][6] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].percentChange.toFixed(2);
            }
            StockArrayX[kLoop][LLoop][7] = requestObjX.composition.subIndices[kLoop].stockInfos[LLoop].totalValue/1000;
            StockArrayX[kLoop][LLoop][8] = StockArrayX[kLoop][LLoop][2] + StockArrayX[kLoop][LLoop][3];
            StockArrayX[kLoop][LLoop][3] = "'" + Temp_marketDateTime;
          }
        sheet_Draft_JSON.getRange(sheet_Draft_JSON.getLastRow()+1,1,stockInfos_Count,9).setValues(StockArrayX[kLoop]);
        SpreadsheetApp.flush();
        }
          //console.log(sheet_Draft_JSON.getLastRow());
          //console.log(TotalStock_subIndices_Count);

          //-----Ending SET------------

      }

      /*
      var Industry_Count = 0;
      var StockArray = new Array (Industry_Count);
      for (var jLoop = 0; jLoop < requestObj.length; jLoop++) {
        if (requestObj[jLoop].level == "INDUSTRY") {
          StockArray[Industry_Count] = requestObj[jLoop].symbol;
          if (Market_Param == "mai") { StockArray[Industry_Count] = StockArray[Industry_Count] + "-m";}
          Industry_Count++;
        }
      }
      */

      //console.log(StockArrayX);

      //Logger.log("ExtractStock_Button has finished.");
      //SpreadsheetApp.flush();
 
  } else {
      console.log("Error : " + JSON_URL_responseX.getResponseCode());
      NotifyMessage("Error : " + JSON_URL_responseX.getResponseCode());
  }
}

//-----------------------------------------------------------------------------------------------------------------
