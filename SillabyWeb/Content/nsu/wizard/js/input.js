
window.sillaby = (function () {    
    var lib = {};

    lib.init = function (container, callback) {

        $(container).on('click', '#saveInputButton', function () {
            var text = $('#saveInputText').val();
            callback(text, "input");
        });

    };

    lib.getBaseUrl = function () {
        var getUrl = window.location;
        var baseUrl = getUrl.protocol + "//" + getUrl.host + "/" + getUrl.pathname.split('/')[1];
        return baseUrl;
    }

    lib.getUrlRequest = function () {
        if (lib.getBaseUrl().indexOf("localhost") > -1) {
            return "";
        }
        else {
            return "/SillabyWeb";
        }
    }

    return lib;
}());