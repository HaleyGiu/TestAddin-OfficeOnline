// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    ensureStateInitialized(true);
    console.log("ensure state initialized from the office.initialize");
    isOfficeInitialized = true;
    monitorSheetChanges();

    document.getElementById("connectService").onclick = connectService; // in office-apis-helpers.js
    document.getElementById("selectFilter").onclick = insertFilteredData;
    document.getElementById("btnSandbox").onclick = btnSandbox; // in office-apis-helpers.js
    
    updateRibbon();
    updateTaskPaneUI();
    callApi();
  }
});

async function insertFilteredData() {
  try {
    //Determine which data source the user selected from the radio buttons.
    const radioExcel = document.getElementById("communicationFilter");
    if (radioExcel.checked) {
      generateCustomFunction("Communications");
    } else {
      generateCustomFunction("Groceries");
    }
  } catch (error) {
    console.error(error);
  }
}

async function callApi() {
  const apiUrl = "https://localhost/OOS.WebAPIExcel/api/Engine/Test";

  try {
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
    const apiSmokeTestResponseHTML = document.getElementById("apiSmokeTestResponse");
    apiSmokeTestResponseHTML.innerHTML = JSON.stringify(data);
  } catch (error) {
    apiSmokeTestResponseHTML.innerHTML = "No se ha podido conectar con el servicio.";
  }
}

function btnSandbox(event) {
  console.log("Sandbox iniciado");
  // ***************

  const authContext = Office.context.auth;
  authContext.getAccessTokenAsync(function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
          const token = result.value;
          console.log(token);
          console.log(result);
      } else {
          console.log("Error obtaining token", result.error);
      }
  }); 

  //***************
  g.state.setConnected(true);
  g.state.isConnectInProgress = true;
  updateRibbon();
  connectService();
  monitorSheetChanges();
  event.completed();
}