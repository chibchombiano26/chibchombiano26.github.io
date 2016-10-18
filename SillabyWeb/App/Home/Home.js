/// <reference path="../App.js" />
// global app

(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            
            window.countControls = 0;

            //List of all the controls
            //sillaby.controls

            app.initialize();
            //Load the dock file
            sillaby.loadDocx().done(function (documen64) {
                //Put the text of loaded file in the document
                sillaby.displayContentDocx(documen64).done(function () {
                    //initialize events
                    sillaby.init("#wizardContent", setValueCustomControl);
                    //List controls on the document
                    sillaby.listControls().done(function (data) {
                        //debugger
                    })
                })
            })           
            
            $('#get-data-from-selection').click(getParameters);            

        });
    };

    function setValueCustomControl(value, type){
            
        if (window.countControls <= sillaby.controls.length -1){

            var control = sillaby.controls[window.countControls];
            sillaby.executeFunctionOnControl(control.m_id, value);

            window.countControls = window.countControls + 1;
        }

    }

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
			function (result) {
			    if (result.status === Office.AsyncResultStatus.Succeeded) {
			        app.showNotification('The selected text is:', '"' + result.value + '"');
			    } else {
			        app.showNotification('Error:', result.error.message);
			    }
			}
		);
    }

    

})();