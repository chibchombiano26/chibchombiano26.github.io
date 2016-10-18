
window.sillaby = (function () {    
    var lib = {};

    lib.init = function (container, callback) {
        
        $(container).on('click', '#saveInputButton', function () {
            var text = $('#saveInputText').val();
            callback(text, "input");
        });

    }

    return lib;
}());