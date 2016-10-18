sillaby.loadDocx = function () {
    var dfd = jQuery.Deferred();
    $.get("/Content/Sample/base64.txt", function (data) {        
        dfd.resolve(data);
    })
   .done(function (e) {
       //alert("second success");
   })
    .fail(function (e) {
        //alert("error");
    })
    .always(function (e) {
        //alert("finished");
    });
    return dfd.promise();
}

sillaby.displayContentDocx = function (myBase64) {
        var dfd = jQuery.Deferred();
        Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;

        // Queue a command to clear the body contents. 
        thisDocument.body.clear();
        thisDocument.body.insertFileFromBase64(myBase64, "replace");
                    
        context.sync().then(function () {
            dfd.resolve();
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