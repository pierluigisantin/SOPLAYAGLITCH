"use strict";

var _worksheetname = "Piani Attività drill cliente";
var worksheet;
var _popupurl = "./popup.html";
var worksheetDataSource;
var OpenPopUp;
var _ForecastDS='Forecast';

// Wrap everything in an anonymous function to avoid polluting the global namespace
(function () {
  $(document).ready(
    function () {
      tableau.extensions.initializeAsync().then(function () {
        ////SOPLAYA
        const worksheets =
          tableau.extensions.dashboardContent.dashboard.worksheets;
        worksheet = worksheets.find(function (sheet) {
          return sheet.name === _worksheetname;
        });
        worksheet.getDataSourcesAsync().then(function (fetchResults) {
          fetchResults.forEach(function (dataSourceForWorksheet) {
            if (
              dataSourceForWorksheet.name === _ForecastDS
            ) {
              worksheetDataSource = dataSourceForWorksheet;
            }
          });
        });
        
        $("#modifyBtn").click(OpenPopUp);
        
  /*
        const unregisterHandlerFunction = worksheet.addEventListener(
          tableau.TableauEventType.MarkSelectionChanged,
          MarkSelectionChangedHandler
        );
        */
      });
    },
    function (err) {
      // Something went wrong in initialization.
      console.log("Error while Initializing: " + err.toString());
    }
  );

  
  
  ///SOPLAYA
  OpenPopUp=function()
  {
        
      var payload = "";

      tableau.extensions.ui
        .displayDialogAsync(_popupurl, payload, { width: 1024, height: 768 })
        .then((closePayload) => {
          //
          alert(closePayload);

          // The promise is resolved when the dialog has been closed as expected, meaning that
          // the popup extension has called tableau.extensions.ui.closeDialog() method.
          // The close payload (closePayload) is returned from the popup extension
          // via the closeDialog() method.
          //
        })
        .catch((error) => {
          // One expected error condition is when the popup is closed by the user (meaning the user
          // clicks the 'X' in the top right of the dialog). This can be checked for like so:
          switch (error.errorCode) {
            case tableau.ErrorCodes.DialogClosedByUser:
              console.log("Dialog was closed by user");
              if (worksheetDataSource) worksheetDataSource.refreshAsync();
              break;
            default:
              console.error(error.message);
          }
        });
    
  }
  
  
  
  function MarkSelectionChangedHandler(marksEvent) {

  }
})();
