var loadData;
var checkForAuth;
var test;

(function () {
  $(document).ready(function () {
    tableau.extensions.initializeDialogAsync().then(
      function (s) {
        console.log("extension initialized");
        //alert('inizializzato');
        payload = s;
        loadGoogle();
        loadData();
      },
      function (err) {
        // Something went wrong in initialization.
        console.log("Error while Initializing: " + err.toString());
      }
    );
  });
  /*  extension can define other functions here as needed */

  var payload;
  var _worksheetname = "Creaz Budget";
  var worksheet;
  var worksheetData;

  var colRestaurantId_Name = "Restaurant Id";
  var colRestaurantId_NameIta = "Restaurant Id";
  var colRestaurantName_Name = "Name (Sales Restaurants)";
  var colRestaurantName_NameIta = "Name (Sales Restaurants)";
  var colNomeMisura_Name = "Measure Names";
  var colNomeMisura_NameIta = "Nomi misure";
  var colValoreMisura_Name = "Measure Values";
  var colValoreMisura_NameIta = "Valori misure";
  var colGmv1GenAp_Name = "AGG(GMV 1°gen-oggi AP)";
  var colGmv1GenAp_NameIta = "AGG(GMV 1°gen-oggi AP)";
  var colGmv1Gen_Name = "AGG(GMV 1°gen-oggi)";
  var colGmv1Gen_NameIta = "AGG(GMV 1°gen-oggi)";
  var colGmvMaxMensileAp_Name = "AGG(GMV Max Mensile AP)";
  var colGmvMaxMensileAp_NameIta = "AGG(GMV Max Mensile AP)";
  var colGmvMaxMensile_Name = "AGG(GMV Max Mensile)";
  var colGmvMaxMensile_NameIta = "AGG(GMV Max Mensile)";
  var colGmvMedioMensileAp_Name = "AGG(GMV Medio Mensile AP)";
  var colGmvMedioMensileAp_NameIta = "AGG(GMV Medio Mensile AP)";
  var colGmvMedioMensile_Name = "AGG(GMV Medio Mensile)";
  var colGmvMedioMensile_NameIta = "AGG(GMV Medio Mensile)";
  var colGMVTotAp_Name = "AGG(GMV TOT AP)";
  var colGMVTotAp_NameIta = "AGG(GMV TOT AP)";

  var PrevisioneMeasureName = "Previsione";
  var PotenzialeMeasureName = "Potenziale";

  var colRestaurantId = -1;
  var colRestaurantName = -1;
  var colGmv1GenAp = -1;
  var colGmv1Gen = -1;
  var colGmvMaxMensileAp = -1;
  var colGmvMaxMensile = -1;
  var colGmvMedioMensileAp = -1;
  var colGmvMedioMensile = -1;
  var colGmvTotaleAp = -1;
  var colNomeMisura = -1;
  var colValoreMisura = -1;

  var spreadsheetId = "1Ibv6YCUn5HqZg66FNwiv60dbMXL0361IrjeazteGJl8";

  const socket = io();
  socket.on("connect", () => {
    ///do nothoing
  });

  socket.on("signedin", function (token) {
    if (token) {
      localStorage.setItem("googleAuth", token);
      loadGoogle(function () {});
    } else {
      console.error("Token response was empty!");
    }
  });

  var authGapi = function () {
    gapi.auth.authorize(
      {
        immediate: true,
        client_id:
          "334355323973-nb13gnl993d8gb71n88ek33vla66jhj9.apps.googleusercontent.com",
        scope: "https://www.googleapis.com/auth/spreadsheets",
      },
      function (response) {
        return gapi.client
          .init({
            discoveryDocs: [
              "https://sheets.googleapis.com/$discovery/rest?version=v4",
            ],
          })
          .then(function () {
            var oauthToken = localStorage.getItem("googleAuth");
            console.log("token set: " + oauthToken);
            gapi.auth.setToken({ access_token: oauthToken });
            testGoogle();
          });
      }
    );
  };

  function loadGoogle() {
    var oauthToken = localStorage.getItem("googleAuth");
    if (oauthToken)
      gapi.load("client", async () => {
        authGapi();
      });
    else openSignInWindow();
  }

  function openSignInWindow() {
    var url =
      "" +
      "https://accounts.google.com/o/oauth2/v2/auth?" +
      "scope=https%3A//www.googleapis.com/auth/spreadsheets&" +
      "include_granted_scopes=true&" +
      "response_type=token&" +
      "state=" +
      socket.id +
      "&" +
      "redirect_uri=https%3A//soplaya.glitch.me/oathcomplete&" +
      "client_id=334355323973-nb13gnl993d8gb71n88ek33vla66jhj9.apps.googleusercontent.com";
    window.open(url, "_blank");
  }

  loadData = function () {
    //alert('carico i dati');
    const worksheets = tableau.extensions.dashboardContent.dashboard.worksheets;
    worksheet = worksheets.find(function (sheet) {
      return sheet.name === _worksheetname;
    });

    worksheet
      .getSummaryDataAsync({ ignoreSelection: true })
      .then(function (sumdata) {
        worksheetData = sumdata;
        buildTable(worksheetData);
      });
  };

  var buildTable = function (sumData) {
    $("#table").empty();
    //
    sumData.columns.forEach(function (c) {
      if (
        c.fieldName === colRestaurantId_Name ||
        c.fieldName === colRestaurantId_NameIta
      )
        colRestaurantId = c.index;
      if (
        c.fieldName === colRestaurantName_Name ||
        c.fieldName === colRestaurantName_NameIta
      )
        colRestaurantName = c.index;
      if (
        c.fieldName === colValoreMisura_Name ||
        c.fieldName === colValoreMisura_NameIta
      )
        colValoreMisura = c.index;
      if (
        c.fieldName === colNomeMisura_Name ||
        c.fieldName === colNomeMisura_NameIta
      )
        colNomeMisura = c.index;
      if (
        c.fieldName === colGmv1GenAp_Name ||
        c.fieldName === colGmv1GenAp_NameIta
      )
        colGmv1GenAp = c.index;
      if (c.fieldName === colGmv1Gen_Name || c.fieldName === colGmv1Gen_NameIta)
        colGmv1Gen = c.index;
      if (
        c.fieldName === colGmvMaxMensileAp_Name ||
        c.fieldName === colGmvMaxMensileAp_NameIta
      )
        colGmvMaxMensileAp = c.index;
      if (
        c.fieldName === colGmvMaxMensile_Name ||
        c.fieldName === colGmvMaxMensile_NameIta
      )
        colGmvMaxMensile = c.index;
      if (
        c.fieldName === colGmvMedioMensileAp_Name ||
        c.fieldName === colGmvMedioMensileAp_NameIta
      )
        colGmvMedioMensileAp = c.index;
      if (
        c.fieldName === colGmvMedioMensile_Name ||
        c.fieldName === colGmvMedioMensile_NameIta
      )
        colGmvMedioMensile = c.index;
      if (
        c.fieldName === colGMVTotAp_Name ||
        c.fieldName === colGMVTotAp_NameIta
      )
        colGmvTotaleAp = c.index;
    });

    //metto i dati in una struttura ad-hoc che mi va comoda
    var markup = "";
    var lastRowId = "";
    var rows = {};
    var currRow = {};
    var anni = {};
    var orderedRowIds=[];
    var anno = 2022;///to grab from the data

    sumData.data.forEach(function (d) {
      if (
        d[colNomeMisura].formattedValue === PrevisioneMeasureName ||
        d[colNomeMisura].formattedValue === PotenzialeMeasureName
      ) {
        var rowId = d[colRestaurantId].formattedValue+"#"+anno;

        if (rows[rowId]) currRow = rows[rowId];
        else {
          currRow = {};
          rows[rowId] = currRow;
        }
        
        if (!orderedRowIds.includes(rowId))
          orderedRowIds.push(rowId);

        currRow[colRestaurantId] = d[colRestaurantId].nativeValue;
        currRow[colRestaurantName] = d[colRestaurantName].formattedValue;
        currRow[colGmv1GenAp] = d[colGmv1GenAp].formattedValue;
        currRow[colGmv1Gen] = d[colGmv1Gen].formattedValue;
        currRow[colGmvMaxMensileAp] = d[colGmvMaxMensileAp].formattedValue;
        currRow[colGmvMaxMensile] = d[colGmvMaxMensile].formattedValue;
        currRow[colGmvMedioMensileAp] = d[colGmvMedioMensileAp].formattedValue;
        currRow[colGmvMedioMensile] = d[colGmvMedioMensile].formattedValue;
        currRow[colGmvTotaleAp] = d[colGmvTotaleAp].formattedValue;
        currRow[d[colNomeMisura].formattedValue] =
          d[colValoreMisura].nativeValue;
      }
    });

    //add headers
    $("#table").append(`  
                                <tr>
                                <th>Cliente</th>
                                <th>GMV Tot AP</th>
                                <th> GMV dal 1 Gen</th>
                                <th> GMV dal 1 Gen AP</th>
                                <th> GMV Mensile Medio</th>
                                <th> GMV Mensile Medio AP</th>
                                <th> GMV Mensile Max</th>
                                <th> GMV Mensile Max AP</th>
                                <th> Potenziale</th>
                                <th> Previsione</th>
                                </tr >
                                `);
    //costruimano la tabella
    var markup = "";
    for (let i = 0; i < orderedRowIds.length; i++) {
      var rowid = orderedRowIds[i];
      var rowToRender = rows[rowid];

      markup = markup + "<tr>";
      
      markup=markup+"<td>"+rowToRender[colRestaurantName]+"</td>";
      markup=markup+"<td>"+rowToRender[colGmvTotaleAp]+"</td>";
      markup=markup+"<td>"+rowToRender[colGmv1Gen]+"</td>";
      markup=markup+"<td>"+rowToRender[colGmv1GenAp]+"</td>";
      markup=markup+"<td>"+rowToRender[colGmvMedioMensile]+"</td>";
      markup=markup+"<td>"+rowToRender[colGmvMedioMensileAp]+"</td>";
      markup=markup+"<td>"+rowToRender[colGmvMaxMensile]+"</td>";
      markup=markup+"<td>"+rowToRender[colGmvMaxMensileAp]+"</td>";
      markup=markup+'<td>'+
         '<input type="number"  value="' +
          RenderStr(rowToRender[PotenzialeMeasureName]) +
          '" name="' +
          rowid+'#'+PotenzialeMeasureName+'#'+rowToRender[colRestaurantName]+'"'+
          '/>';
      markup=markup+'<td>'+
         '<input type="number"  value="' +
          RenderStr(rowToRender[PrevisioneMeasureName]) +
          '" name="' +
          rowid+'#'+PrevisioneMeasureName+'#'+rowToRender[colRestaurantName]+'"'+
          '/>';
      

      markup = markup + "</tr>";
    }
    $("#table").append(markup);

    $('input[type="number"]').on("change", focusoutHandler);

  };





  function focusoutHandler(ev) {
    var elementToUpdate = ev.target.name;
    var valueToUpdate = ev.target.value;
    if (gapi.client && gapi.client.sheets)
      saveData(elementToUpdate, valueToUpdate);
    else loadGoogle(); //.then (saveData(elementToUpdate, valueToUpdate));
  }



  function RenderStr(s) {
    var strOutPut = "";

    if (s || typeof s !== "undefined") {
      var ss = s + "";
      if (ss !== "Null" && ss !== "null" && ss != "NULL") strOutPut = ss;
    }

    return strOutPut;
  }

  function ManageUpdate(response) {
    $("#snackbar").html("dati salvati");
    $("#snackbar").addClass("show");
    setTimeout(function () {
      $("#snackbar").removeClass("show");
    }, 1000);
  }

  

  function testGoogle() {
    gapi.client.sheets.spreadsheets.values
      .get({
        spreadsheetId: spreadsheetId,
        range: "Budget!A:D",
      })
      .then(function (response) {})
      .catch((err) => manageErr(err));
  }

 

  function saveData(target, val) {

    var data = target.split("#");
    var restaurantId = data[0];
    var anno = data[1];
    var nomeMisura = data[2];
    var restauranteName = data[3];
    var notes = "";
    var valPotenziale;
    var valPrevisione;
    if (nomeMisura==PotenzialeMeasureName)
      valPotenziale=val;
    else
      valPrevisione=val;


    gapi.client.sheets.spreadsheets.values
      .get({
        spreadsheetId: spreadsheetId,
        range: "Budget!A:D",
      })
      .then(function (response) {
        var found = false;
        var i = 1;
        var indexRow = -1;
        response.result.values.forEach(function (r) {
          if (
            r[0] == restaurantId &&
            r[1] == anno 
          ) {
            found = true;
           if (nomeMisura==PotenzialeMeasureName)//grab previous value
              valPrevisione =r[3];
           else
            valPotenziale=r[2];
            
            indexRow = i;
          }
          i++;
        });

        if (found === false) {
          //append
          var submissionValues = [];
          submissionValues.push(restaurantId);
          submissionValues.push(anno);
          submissionValues.push(valPotenziale);
          submissionValues.push(valPrevisione);
          submissionValues.push(restauranteName);
          submissionValues.push(notes);
          
          const params = {
            spreadsheetId: spreadsheetId,
            range: `Budget!A${i}:F${i}`,
            valueInputOption: "RAW", //RAW = if no conversion or formatting of submitted data is needed. Otherwise USER_ENTERED
            insertDataOption: "INSERT_ROWS", //Choose OVERWRITE OR INSERT_ROWS
          };
          const valueRangeBody = {
            majorDimension: "ROWS", //log each entry as a new row (vs column)
            values: [submissionValues], //convert the object's values to an array
          };
          gapi.client.sheets.spreadsheets.values
            .append(params, valueRangeBody)
            .then(ManageUpdate(response))
            .catch((err) => manageErr(err));
        } else {
          //update
          var submissionValues = [];
          submissionValues.push(restaurantId);
          submissionValues.push(anno);
          submissionValues.push(valPotenziale);
          submissionValues.push(valPrevisione);
          submissionValues.push(restauranteName);
          submissionValues.push(notes);
          const params = {
            spreadsheetId: spreadsheetId,
            range: `Budget!A${indexRow}:F${indexRow}`,
            valueInputOption: "RAW", //RAW = if no conversion or formatting of submitted data is needed. Otherwise USER_ENTERED
          };
          const valueRangeBody = {
            majorDimension: "ROWS", //log each entry as a new row (vs column)
            values: [submissionValues], //convert the object's values to an array
          };
          gapi.client.sheets.spreadsheets.values
            .update(params, valueRangeBody)
            .then(ManageUpdate(response))
            .catch((err) => manageErr(err));
        }
      })
      .catch((err) => manageErr(err));
  }

  function manageErr(err) {
    if (err.result && err.result.error) {
      alert(
        err.result.error.code +
          "-" +
          err.result.error.message +
          "-" +
          err.result.error.status
      );
    } else {
      alert(err);
    }
    if (err.status === 401 || err.status === 403) {
      openSignInWindow();
    }
  }

  function pausecomp(millis) {
    var date = new Date();
    var curDate = null;
    do {
      curDate = new Date();
    } while (curDate - date < millis);
  }

  test = function () {
    loadData();
  };
})();
