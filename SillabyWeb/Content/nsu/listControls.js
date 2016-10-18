sillaby.listControls = function listControls() {
    var dfd = jQuery.Deferred();

    Word.run(function (context) {

        // Create a proxy object for the content controls collection that contains a specific tag.
        var contentControlsWithTag = context.document.contentControls.getByTag('input');

        // Queue a command to load the text property for all of content controls with a specific tag.
        context.load(contentControlsWithTag, 'text,title');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            sillaby.controls = contentControlsWithTag.items;            
            //contentControlsWithTag.items[0].insertText('Replaced text in the first content control.', 'Replace');
            dfd.resolve(sillaby.controls);
        });
    })
    .catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });

    return dfd.promise();
}

sillaby.listControlsRegex = function (wildcard) {
    var dfd = jQuery.Deferred();

    Word.run(function (context) {

        // Create a proxy object for the content controls collection that contains a specific tag.
        var searchResults = context.document.body.search(wildcard, { matchWildCards: true });

        // Queue a command to load the search results and get the font property values.
        context.load(searchResults, 'font', 'text');

        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Found count: ' + searchResults.items.length);
            sillaby.controls = searchResults.items;
            //contentControlsWithTag.items[0].insertText('Replaced text in the first content control.', 'Replace');
            dfd.resolve(sillaby.controls);
        });
    })
    .catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });

    return dfd.promise();
}


function getParameters() {

    $('#wizardContent').load(sillaby.getUrlRequest() + '/Content/nsu/wizard/template/input.html');

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

sillaby.executeFunctionOnControl = function (id, params) {
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the content controls collection.
        var contentControls = context.document.contentControls;

        // Queue a command to load the id property for all of the content controls. 
        context.load(contentControls, 'id, parent');

        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            if (contentControls.items.length === 0) {
                console.log('No content control found.');
            }
            else {

                var index = _.findIndex(contentControls.items, { "m_id": id });
                // Queue a command to replace text in the first content control. 

                if (index > -1) {
                    //contentControls.items[index].insertText(params, 'Replace');
                    contentControls.items[index].clear();
                }

                // Synchronize the document state by executing the queued commands, 
                // and return a promise to indicate task completion.
                return context.sync()
                    .then(function () {
                        console.log('Replaced text in the first content control.');
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

sillaby.search = function (wildcard, value, id) {
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Queue a command to search the document based on a prefix.
        var searchResults = context.document.body.search(wildcard, { matchWildCards: true });

        // Queue a command to load the search results and get the font property values.
        context.load(searchResults, 'font, text');

        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {

            var index = _.findIndex(searchResults.items, { "m__Id": id });

            // Queue a command to replace text in the first content control. 
            if (index > -1) {
                var currentText = searchResults.items[index].m_text;                
                var regex = /.*/i;
                value = currentText.replace(regex, value);
                console.log(value);

                searchResults.items[index].font.color = 'purple';
                searchResults.items[index].font.highlightColor = '#FFFF00'; //Yellow
                searchResults.items[index].font.bold = true;
                searchResults.items[index].insertText(value, Word.InsertLocation.replace);
            }

            return context.sync();
        });
    })
    .catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

sillaby.addWildCard = function (element) {
    return "--" + element + "--";
}