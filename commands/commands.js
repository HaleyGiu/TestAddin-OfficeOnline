// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function btnConnectService(event) {
  console.log("Connect service button pressed");
  // Your code goes here
  g.state.setConnected(true);
  g.state.isConnectInProgress = true;
  updateRibbon();
  connectService();
  monitorSheetChanges();
  event.completed();
}


function btnDisconnectService(event) {
  console.log("Disconnect service button pressed");
  // Your code goes here
  g.state.setConnected(false);
  updateRibbon();
  updateTaskPaneUI();
  event.completed();
}

 function btnOpenTaskpane(event) {
  console.log("Open task pane button pressed");
  // Your code goes here
  SetRuntimeVisibleHelper(true);
  g.state.isTaskpaneOpen = true;
  updateRibbon();
  event.completed();
}

 function btnCloseTaskpane(event) {
  console.log("Open task pane button pressed");
  // Your code goes here
  SetRuntimeVisibleHelper(false);
  g.state.isTaskpaneOpen = false;
  updateRibbon();
  event.completed();
}

 function btnEnableAddinStart(event) {
  console.log("Enable add-in start button pressed");
  // Your code goes here
  SetStartupBehaviorHelper(true);
  g.state.isStartOnDocOpen = true;
  updateRibbon();
  event.completed();
}

 function btnDisableAddinStart(event) {
  console.log("Disable add-in start button pressed");
  // Your code goes here
  SetStartupBehaviorHelper(false);
  g.state.isStartOnDocOpen = false;
  updateRibbon();

  event.completed();
}

 function btnInsertData(event) {
  console.log("Insert data button pressed");
  // Mock code that pretends to insert data from a data source
  insertData();
  event.completed();
}

  function btnInsertApiResponseData(event) {
    console.log("Insert data button pressed");
    // Mock code that pretends to insert data from a data source
    insertDataResponse();
    event.completed();
  }

 async function btnSumData(event) {
  console.log("Insert data button pressed");
  // Mock code that pretends to insert data from a data source
  let address = g.state.selectionAddress;
  await Excel.run((context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange(address);
    range.load("values");

    let sum = 0;
    return context.sync().then(() => {
      range.values.forEach((v) => {
        let vnumber = +v.toString();
        sum += vnumber;
      });

      return context.sync().then(() => {
        let sheet = context.workbook.worksheets.getActiveWorksheet();

        let range = sheet.getRange("F1");
        range.values = [[sum]];
        range.format.autofitColumns();
        event.completed();
        console.log(sum);
        return context.sync();
      });
    });
  });
  event.completed();
}

async function insertData() {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
      expensesTable.name = "ExpensesTable";

      expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

      expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ["1/1/2017", "The Phone Company", "Communications", "$120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
        ["1/11/2017", "Bellows College", "Education", "$350"],
        ["1/15/2017", "Trey Research", "Other", "$135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"],
      ]);

      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }
      context.sync().then(() => {
        monitorSheetChanges();
        return context.sync();
      });
    });
  } catch (error) {
    console.log(error);
  }
}

async function insertDataResponse() {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      const apiUrl = "https://localhost/OOS.WebAPIExcel/api/DataQuery/Execute?QueryCode=rBDListPhoneNumbersGet&DatasourceCode&ChameleonStaffCode&Parameters&MenuID=4101";
      // Ya disponemos de la tabla, no podemos intentar crearla porque se solapa
      //let responseTable = sheet.tables.add("B6:R6", true);
      //responseTable.name = "Table1";
      //responseTable.getHeaderRowRange().values = [["Staff	Nombre", "Oficina", "Categoría", "División", "Entidad Legal", "Departamento", "Dirección", "Despacho", "Ext 1", "Ext 2", "Teléfono", "Móvil", "Email", "GUID", "StaffCode", "Secretaria", "Nombre Secretaria"]];

      let responseTable = sheet.tables.getItem("Table1");


      const response = await fetch(apiUrl, {
        method: "GET", 
        headers: {
          "Content-Type": "application/json"
        },
      });

      if (!response.ok) {
        throw new Error(`Error en la solicitud: ${response.statusText}`);
      }
  
      const data = await response.json();
      const result = data["ResultSets"][0]["Data"].map(obj => Object.values(obj));
      responseTable.rows.add(null, result);

      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }
      context.sync().then(() => {
        monitorSheetChanges();
        return context.sync();
      });

      
    });
  } catch (error) {
    console.log(error);
  }
}

const g = getGlobal();
  
Office.actions.associate("btnConnectService", btnConnectService);
Office.actions.associate("btnDisconnectService", btnDisconnectService);
Office.actions.associate("btnOpenTaskpane", btnOpenTaskpane);
Office.actions.associate("btnCloseTaskpane", btnCloseTaskpane);
Office.actions.associate("btnEnableAddinStart", btnEnableAddinStart);
Office.actions.associate("btnDisableAddinStart", btnDisableAddinStart);
Office.actions.associate("btnInsertData", btnInsertData);
Office.actions.associate("btnInsertApiResponseData", btnInsertApiResponseData);
Office.actions.associate("btnSumData", btnSumData);
