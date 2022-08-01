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
  var _worksheetname = "Piani Attività drill cliente";
  var worksheet;
  var worksheetData;
  const _GMVMeasureName = "GMV";
  const _ForecastMeasureName = "Previsione";

  var colRestaurantId_Name = "Restaurant Id";
  var colRestaurantId_NameIta = "Restaurant Id";
  var colRestaurantName_Name = "Name";
  var colRestaurantName_NameIta = "Name";
  var colMese_Name = "MESE";
  var colMese_NameIta = "MESE";
  var colNomeMisura_Name = "Measure Names";
  var colNomeMisura_NameIta = "Nomi misure";
  var colAnno_Name = "YEAR(gg)";
  var colAnno_NameIta = "ANNO(gg)";
  var colSettimana_Name = "settimana";
  var colSettimana_NameIta = "settimana";
  var colValoreMisura_Name = "Measure Values";
  var colValoreMisura_NameIta = "Valori misure";
  var colForecastManuale_Name = "SUM(Forecast manuale)";
  var colForecastManuale_NameIta = "SOMMA(Forecast manuale)";
  var colNotes_Name = "AGG(Notes)";
  var colNotes_NameIta = "AGG(Notes)";
  var colColoreCella_Name = "AGG(colore cella)";
  var colColoreCella_NameIta = "AGG(colore cella)";

  var colRestaurantId = -1;
  var colRestaurantName = -1;
  var colMese = -1;
  var colNomeMisura = -1;
  var colAnno = -1;
  var colSettimana = -1;
  var colValoreMisura = -1;
  var googleLoaded = false;
  var colForecastManuale = -1;
  var colNotes = -1;
  var colColoreCella = -1;

  var colorYellowClass = "yellowClass";
  var colorRedClass = "redClass";

  var anno1;
  var anno2;

  var spreadsheetId = "1oTbG7vNTWjXjQkEjXBBT0qiGqJiMWMjrcaA28Upzv2s";

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
    SignalLoginNeeded();
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
      if (c.fieldName === colMese_Name || c.fieldName === colMese_NameIta)
        colMese = c.index;
      if (
        c.fieldName === colNomeMisura_Name ||
        c.fieldName === colNomeMisura_NameIta
      )
        colNomeMisura = c.index;
      if (c.fieldName === colAnno_Name || c.fieldName === colAnno_NameIta)
        colAnno = c.index;
      if (
        c.fieldName === colSettimana_Name ||
        c.fieldName === colSettimana_NameIta
      )
        colSettimana = c.index;
      if (
        c.fieldName === colValoreMisura_Name ||
        c.fieldName === colValoreMisura_NameIta
      )
        colValoreMisura = c.index;
      if (
        c.fieldName === colForecastManuale_Name ||
        c.fieldName === colForecastManuale_NameIta
      )
        colForecastManuale = c.index;
      if (c.fieldName === colNotes_Name || c.fieldName === colNotes_NameIta)
        colNotes = c.index;
      if (
        c.fieldName === colColoreCella_Name ||
        c.fieldName === colColoreCella_NameIta
      )
        colColoreCella = c.index;
    });

    //metto i dati in una struttura ad-hoc che mi va comoda
    var markup = "";
    var lastRowId = "";
    var rows = {};
    var currRow = {};
    var anni = {};

    sumData.data.forEach(function (d) {
      if (
        d[colNomeMisura].formattedValue === _GMVMeasureName ||
        d[colNomeMisura].formattedValue === _ForecastMeasureName
      ) {
        var rowId =
          d[colRestaurantId].formattedValue +
          "#" +
          d[colRestaurantName].formattedValue +
          "#" +
          d[colNomeMisura].formattedValue +
          "#" +
          d[colMese].formattedValue;

        if (rows[rowId]) currRow = rows[rowId];
        else {
          currRow = {};
          rows[rowId] = currRow;
        }

        //memorizziamo gli anni
        if (!anni[d[colAnno].formattedValue])
          anni[d[colAnno].formattedValue] = d[colAnno].formattedValue;

        currRow[colRestaurantName] = d[colRestaurantName].formattedValue;
        currRow[colMese] = d[colMese].formattedValue;
        currRow[colNomeMisura] = d[colNomeMisura].formattedValue;
        currRow[colRestaurantId] = d[colRestaurantId].formattedValue;

        if (
          d[colNotes].formattedValue &&
          d[colNotes].formattedValue.length > 0 &&
          d[colNotes].formattedValue.toUpperCase() !== "NULL"
        )
          currRow[colNotes] = d[colNotes].formattedValue;

        if (currRow[colNomeMisura] === _ForecastMeasureName) {
          currRow[
            d[colAnno].formattedValue + "#" + d[colSettimana].formattedValue
          ] = d[colValoreMisura].nativeValue;

          currRow[
            d[colAnno].formattedValue +
              "#" +
              d[colSettimana].formattedValue +
              "#colore"
          ] = d[colColoreCella].nativeValue;
        } else {
          currRow[
            d[colAnno].formattedValue + "#" + d[colSettimana].formattedValue
          ] = d[colValoreMisura].formattedValue;
        }
      }
    });

    var anniArr = [];
    for (const anno of Object.keys(anni)) {
      anniArr.push(anno);
    }
    anniArr.sort();

    anno1 = anniArr[1];
    anno2 = anniArr[0];
    //add headers
    $("#table").append(`  
                                <tr>
                                <th colspan="3"> </th>
                                <th colspan="4">${anno1} </th>
                                 <th colspan="4">${anno2}</th>
                                </tr >
                                `);
    $("#table").append(`  
                                <tr>
                                <th>Cliente</th>
                                <th>Mese</th>
                                <th> </th>
                                <th> 1</th>
                                 <th> 2</th>
                                <th> 3</th>
                                 <th> 4</th>
                                <th> 1</th>
                                 <th> 2</th>
                                <th> 3</th>
                                 <th> 4</th>
                                <th></th>
                                </tr >
                                `);
    //costruimano la tabella
    var r = 0;
    for (const rowid of Object.keys(rows)) {
      r++;
      var rowToRender = rows[rowid];
      markup = markup + "<tr>";
      if (r % 24 == 1)
        markup =
          markup + "<td rowspan=24>" + rowToRender[colRestaurantName] + "</td>";
      if (r % 2 != 0)
        markup = markup + "<td rowspan=2>" + rowToRender[colMese] + "</td>";
      markup = markup + "<td>" + rowToRender[colNomeMisura] + "</td>";

      if (rowToRender[colNomeMisura] === _ForecastMeasureName) {
        var retClassName = function (val) {
          if (val == "1") return colorYellowClass;
          else {
            if (val == "0") return colorRedClass;
            else return "";
          }
        };

        markup =
          markup +
          '<td><input type="number"  value="' +
          RenderStr(rowToRender[anno1 + "#1"]) +
          '" name="' +
          rowid +
          "#" +
          anno1 +
          "-1" +
          '" ' +
          ' class="' +
          retClassName(rowToRender[anno1 + "#1#colore"]) +
          '"/></td>';
        markup =
          markup +
          '<td><input type="number"  value="' +
          RenderStr(rowToRender[anno1 + "#2"]) +
          '" name="' +
          rowid +
          "#" +
          anno1 +
          "-2" +
          '" ' +
          ' class="' +
          retClassName(rowToRender[anno1 + "#2#colore"]) +
          '"/></td>';
        markup =
          markup +
          '<td><input type="number"  value="' +
          RenderStr(rowToRender[anno1 + "#3"]) +
          '" name="' +
          rowid +
          "#" +
          anno1 +
          "-3" +
          '" ' +
          ' class="' +
          retClassName(rowToRender[anno1 + "#3#colore"]) +
          '"/></td>';
        markup =
          markup +
          '<td><input type="number"  value="' +
          RenderStr(rowToRender[anno1 + "#4"]) +
          '" name="' +
          rowid +
          "#" +
          anno1 +
          "-4" +
          '" ' +
          ' class="' +
          retClassName(rowToRender[anno1 + "#4#colore"]) +
          '"/></td>';
        markup = markup + "<td>" + "Note:" + "</td>";

        var textNotes = "";
        var classNotes = colorRedClass;
        textNotes = RenderStr(rowToRender[colNotes]);
        if (textNotes && textNotes.length > 0) {
          classNotes = colorYellowClass;
        }

        markup =
          markup +
          '<td  colspan="3">' +
          '<input type="text" value="' +
          textNotes +
          '" name="' +
          rowid +
          "#" +
          anno1 +
          "#Notes" +
          '"' +
          ' class="' +
          classNotes +
          '" ' +
          "/>" +
          "</td>";

        markup =
          markup +
          '<td><button id="btn#' +
          rowid +
          '">' +
          ' <p ><i class="arrow down"></i></p>' +
          "</button></td>";

        /*
        markup =
          markup + "<td>" + RenderStr(rowToRender[anno2 + "#1"]) + "</td>";
        markup =
          markup + "<td>" + RenderStr(rowToRender[anno2 + "#2"]) + "</td>";
        markup =
          markup + "<td>" + RenderStr(rowToRender[anno2 + "#3"]) + "</td>";
        markup =
          markup + "<td>" + RenderStr(rowToRender[anno2 + "#4"]) + "</td>";
       */
      } else {
        markup =
          markup + "<td>" + RenderStr(rowToRender[anno1 + "#1"]) + "</td>";
        markup =
          markup + "<td>" + RenderStr(rowToRender[anno1 + "#2"]) + "</td>";
        markup =
          markup + "<td>" + RenderStr(rowToRender[anno1 + "#3"]) + "</td>";
        markup =
          markup + "<td>" + RenderStr(rowToRender[anno1 + "#4"]) + "</td>";
        markup =
          markup + "<td>" + RenderStr(rowToRender[anno2 + "#1"]) + "</td>";
        markup =
          markup + "<td>" + RenderStr(rowToRender[anno2 + "#2"]) + "</td>";
        markup =
          markup + "<td>" + RenderStr(rowToRender[anno2 + "#3"]) + "</td>";
        markup =
          markup + "<td>" + RenderStr(rowToRender[anno2 + "#4"]) + "</td>";
      }
      markup = markup + "</tr>";
    }
    $("#table").append(markup);

    $('input[type="number"]').on("change", focusoutHandler);
    $('input[type="text"]').on("change", focusoutHandlerText);
    $("button").on("click", btnPressedHandler);
  };

  function btnPressedHandler(ev) {
    var btnId = ev.currentTarget.id;
    if (confirm("vuoi copiare la previsione ai mesi successivi?")) {
      copyToFollowingMonths(btnId);
    }
  }

  function copyToFollowingMonths(btnId) {
    
    $('input[type="number"]').unbind();
    $('input[type="text"]').unbind();
    $("button").unbind();
    
    var rowid = btnId.substring(4, btnId.length);
    
    //
    var data = rowid.split("#");
    var restaurantId = data[0];
    var restauranteName = data[1];
    var nomeMisura = data[2];
    var mese = data[3];
    mese = parseInt(mese); //to convert to int

    //retrieve values to copy
    var set1;
    var set1Name = rowid + "#" + anno1 + "-1";
    var set2;
    var set2Name = rowid + "#" + anno1 + "-2";
    var set3;
    var set3Name = rowid + "#" + anno1 + "-3";
    var set4;
    var set4Name = rowid + "#" + anno1 + "-4";
    var Notes;
    var NotesName = rowid + "#" + anno1 + "#Notes";
    $("input").each(function () {
      var iname = $(this).attr("name");
      if (!iname) return;
      if (iname == set1Name) set1 = $(this).val();
      if (iname == set2Name) set2 = $(this).val();
      if (iname == set3Name) set3 = $(this).val();
      if (iname == set4Name) set4 = $(this).val();
      if (iname == NotesName) Notes = $(this).val();
    });
    var targetRows = [];
    var NoteToSetOnTargetRows=Notes;
    if (set1 && set2 && set3 && set4) {
      $("input").each(function () {
        var iname = $(this).attr("name");
        if (!iname) return;
        for (let i = mese + 1; i <= 12; i++) {
          var rowIdTarget =
            restaurantId + "#" + restauranteName + "#" + nomeMisura + "#" + i;
          targetRows.push(rowIdTarget);
          var set1NameTarg = rowIdTarget + "#" + anno1 + "-1";
          var set2NameTarg = rowIdTarget + "#" + anno1 + "-2";
          var set3NameTarg = rowIdTarget + "#" + anno1 + "-3";
          var set4NameTarg = rowIdTarget + "#" + anno1 + "-4";
          var NotesNameTarg = rowIdTarget + "#" + anno1 + "#Notes";

          if (iname == set1NameTarg) {
            $(this).val(set1);
            $(this).removeClass(colorRedClass);
            $(this).addClass(colorYellowClass);
          }
          if (iname == set2NameTarg) {
            $(this).val(set2);
            $(this).removeClass(colorRedClass);
            $(this).addClass(colorYellowClass);
          }
          if (iname == set3NameTarg) {
            $(this).val(set3);
            $(this).removeClass(colorRedClass);
            $(this).addClass(colorYellowClass);
          }
          if (iname == set4NameTarg) {
            $(this).val(set4);
            $(this).removeClass(colorRedClass);
            $(this).addClass(colorYellowClass);
          }

          if (iname == NotesNameTarg) {
            $(this).val(Notes);
            $(this).removeClass(colorRedClass);
            $(this).addClass(colorYellowClass);
          }
        }
      });
      saveAllData(targetRows,NoteToSetOnTargetRows);
    }
    $('input[type="number"]').on("change", focusoutHandler);
    $('input[type="text"]').on("change", focusoutHandlerText);
    $("button").on("click", btnPressedHandler);
  }

  function focusoutHandler(ev) {
    var elementToUpdate = ev.target.name;
    var valueToUpdate = ev.target.value;
    if (gapi.client && gapi.client.sheets)
      saveData(elementToUpdate, valueToUpdate);
    else loadGoogle(); //.then (saveData(elementToUpdate, valueToUpdate));
  }

  function focusoutHandlerText(ev) {
    var elementToUpdate = ev.target.name;
    var valueToUpdate = ev.target.value;
    if (gapi.client && gapi.client.sheets)
      saveNotes(elementToUpdate, valueToUpdate);
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
  
  function SignalLoginNeeded(){
    $("#snackbar").html("autenticazione a google sheet necessaria");
    $("#snackbar").addClass("show");
    setTimeout(function () {
      $("#snackbar").removeClass("show");
    }, 10000);
  }
  

  function saveAllData(targetRows,NoteToSetOnTargetRows) {
    if (!targetRows || targetRows.length === 0) return;
    gapi.client.sheets.spreadsheets.values
      .get({
        spreadsheetId: spreadsheetId,
        range: "FORECAST!A:D",
      })
      .then(function (response) {
        var batchUpdateData = [];
        var batchAppendData = [];
        var lastRowIndex = -1;
        var append1stRow = -1;

        $('input[type="number"]').each(function () {
          var iname = $(this).attr("name");
          if (!iname) return;
          var data = iname.split("#");
          var restaurantId = data[0];
          var restauranteName = data[1];
          var nomeMisura = data[2];
          var mese = data[3];
          var annoSettimana = data[4];

          var curRow =
            restaurantId +
            "#" +
            restauranteName +
            "#" +
            nomeMisura +
            "#" +
            mese;
          if (!targetRows.includes(curRow)) return;

          if (!annoSettimana) return;
          var anno = annoSettimana.split("-")[0];
          var settimana = annoSettimana.split("-")[1];
          var val = $(this).val();
          var found = false;
          var i = 1;
          var indexRow = -1;

          response.result.values.forEach(function (r) {
            if (
              r[0] == restaurantId &&
              r[1] == mese &&
              r[2] == anno &&
              r[3] == settimana
            ) {
              found = true;
              indexRow = i;
            }
            i++;
          });
          var submissionValues = [];
          submissionValues.push(restaurantId);
          submissionValues.push(mese);
          submissionValues.push(anno);
          submissionValues.push(settimana);
          submissionValues.push(val);
          submissionValues.push(restauranteName);
          submissionValues.push(NoteToSetOnTargetRows);
          if (found) {
            var dataItem = {
              majorDimension: "ROWS",
              range: `FORECAST!A${indexRow}:G${indexRow}`,
              values: [submissionValues],
            };
            batchUpdateData.push(dataItem);
          } else {
            if (lastRowIndex === -1) {
              lastRowIndex = i;
              append1stRow = i;
            } else lastRowIndex++;

            batchAppendData.push(submissionValues);
          }
        });
        //save update/////////////////////////////////////////
        if (batchUpdateData.length > 0) {
          var params = {
            spreadsheetId: spreadsheetId,
            valueInputOption: "RAW", //RAW = if no conversion or formatting of submitted data is needed. Otherwise USER_ENTERED
          };
          var resource = {
            data: batchUpdateData,
          };
          //now we can save
          gapi.client.sheets.spreadsheets.values
            .batchUpdate(params, resource)
            .then(ManageUpdate(response))
            .catch((err) => manageErr(err));
        }
        //save append////////////////////////////////////////////
        if (batchAppendData.length > 0) {
          const params2 = {
            spreadsheetId: spreadsheetId,
            range: `FORECAST!A${append1stRow}:G${lastRowIndex}`,
            valueInputOption: "RAW", //RAW = if no conversion or formatting of submitted data is needed. Otherwise USER_ENTERED
            insertDataOption: "INSERT_ROWS", //Choose OVERWRITE OR INSERT_ROWS
          };
          const valueRangeBody = {
            majorDimension: "ROWS", //log each entry as a new row (vs column)
            values: batchAppendData, //convert the object's values to an array
          };
          gapi.client.sheets.spreadsheets.values
            .append(params2, valueRangeBody)
            .then(ManageUpdate(response))
            .catch((err) => manageErr(err));
        }
      })
      .catch((err) => manageErr(err));
  }

  function testGoogle() {
    gapi.client.sheets.spreadsheets.values
      .get({
        spreadsheetId: spreadsheetId,
        range: "FORECAST!A:D",
      })
      .then(function (response) {})
      .catch((err) => manageErr(err));
  }

  function saveNotes(target, val) {
    $('input[name="' + target + '"]').removeClass(colorRedClass);
    $('input[name="' + target + '"]').addClass(colorYellowClass);
    var data = target.split("#");
    var restaurantId = data[0];
    var restauranteName = data[1];
    var mese = data[3];
    var anno = data[4];

    for (let week = 1; week <= 4; week++) {
      gapi.client.sheets.spreadsheets.values
        .get({
          spreadsheetId: spreadsheetId,
          range: "FORECAST!A:D",
        })
        .then(function (response) {
          var found = false;
          var i = 1;
          var indexRow = -1;
          response.result.values.forEach(function (r) {
            if (
              r[0] == restaurantId &&
              r[1] == mese &&
              r[2] == anno &&
              r[3] == week
            ) {
              found = true;
              indexRow = i;
            }
            i++;
          });

          if (found === false) {
            //append
            var submissionValues = [];
            submissionValues.push(restaurantId);
            submissionValues.push(mese);
            submissionValues.push(anno);
            submissionValues.push(week);
            submissionValues.push("");
            submissionValues.push(restauranteName);
            submissionValues.push(val);
            const params = {
              spreadsheetId: spreadsheetId,
              range: `FORECAST!A${i}:G${i}`,
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

            submissionValues.push(val);
            const params = {
              spreadsheetId: spreadsheetId,
              range: `FORECAST!G${indexRow}:G${indexRow}`,
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
  }

  function saveData(target, val) {
    $('input[name="' + target + '"]').removeClass(colorRedClass);
    $('input[name="' + target + '"]').addClass(colorYellowClass);
    var data = target.split("#");
    var restaurantId = data[0];
    var restauranteName = data[1];
    var nomeMisura = data[2];
    var mese = data[3];
    var annoSettimana = data[4];
    var anno = annoSettimana.split("-")[0];
    var settimana = annoSettimana.split("-")[1];

    gapi.client.sheets.spreadsheets.values
      .get({
        spreadsheetId: spreadsheetId,
        range: "FORECAST!A:D",
      })
      .then(function (response) {
        var found = false;
        var i = 1;
        var indexRow = -1;
        response.result.values.forEach(function (r) {
          if (
            r[0] == restaurantId &&
            r[1] == mese &&
            r[2] == anno &&
            r[3] == settimana
          ) {
            found = true;
            indexRow = i;
          }
          i++;
        });

        if (found === false) {
          //append
          var submissionValues = [];
          submissionValues.push(restaurantId);
          submissionValues.push(mese);
          submissionValues.push(anno);
          submissionValues.push(settimana);
          submissionValues.push(val);
          submissionValues.push(restauranteName);
          const params = {
            spreadsheetId: spreadsheetId,
            range: `FORECAST!A${i}:F${i}`,
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
          submissionValues.push(mese);
          submissionValues.push(anno);
          submissionValues.push(settimana);
          submissionValues.push(val);
          const params = {
            spreadsheetId: spreadsheetId,
            range: `FORECAST!A${indexRow}:F${indexRow}`,
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
