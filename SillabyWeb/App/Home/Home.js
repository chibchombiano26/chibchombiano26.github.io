/// <reference path="../App.js" />
// global app

(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#get-data-from-selection').click(getParameters);
        });
    };

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

    function getParameters() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Create a proxy object for the content controls collection that contains a specific tag.
            var contentControlsWithTag = context.document.contentControls.getByTag('input');

            // Queue a command to load the text property for all of content controls with a specific tag.
            context.load(contentControlsWithTag, 'text,title');

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                if (contentControlsWithTag.items.length === 0) {
                    console.log('error');
                } else {

                    var items = [];
                    var list = contentControlsWithTag.items;

                    for (var i = 0; i < list.length; i++) {
                        var item = list[i];
                        items.push({ id: item.m_title, value: item.m_text });
                    }

                    var itemsString = JSON.stringify(items);

                    app.showNotification('Info loaded:', '"' + itemsString + '"');

                    var url = "http://localhost:8540/syllabi/Syllabi/SaveInfo";
                    $.post(url, itemsString,
                    function (data, status) {
                        showNotification('Success', 'Information saved on db thanks');
                    }, function (err) {
                        debugger
                    });

                }

            });
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
    }


})();