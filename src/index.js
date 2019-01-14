'use strict';

(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {
            if (Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {

                $('#create').click(show_signature_pad);
                $('#save').click(insert);
            } else {
                // Just letting you know that this code will not work with your version of Word.
                console.log('This add-in requires Excel 2016 or greater.');
            }
        });
    };

    function show_signature_pad() {
        document.getElementById('signature-pad').removeAttribute("style");
    }

    function insert() {
        var dataURL = signaturePad.toDataURL("image/png");
        dataURL = dataURL.replace('data:image/png;base64,', '');
        Office.context.document.setSelectedDataAsync(dataURL, {
            coercionType: Office.CoercionType.Image,
            imageLeft: 50,
            imageTop: 50,
            imageWidth: 100,
            imageHeight: 100
        },
           function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.log("Action failed with error: " + asyncResult.error.message);

                } else {
                    hide_signature_pad();
                }
            });
    }
    function hide_signature_pad() {
        var canvas = document.getElementById('canvas');
        var ctx = canvas.getContext('2d');
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        document.getElementById('signature-pad').setAttribute("style", "visibility:hidden;");
    }
       
   
})();