/// <reference path="/Scripts/FabricUI/MessageBanner.js" />


(function () {
    "use strict";

    var messageBanner;

    // Die Initialisierungsfunktion muss bei jedem Laden einer neuen Seite ausgeführt werden.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // Initialisiert den FabricUI-Benachrichtigungsmechanismus und blendet ihn aus.
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            
            // Initialize the Download File 
            $('#button-text1').text("Download aktuelle Folien Bibliothek");
            $('#button-desc1').text("Download aktuelle Folien Bibliothek");
            $('#btnDownload1').click(
               downloadFile);

            //Angeklickte Images senden
            $('.sendImageToClient').click(function () {
                //Get source of this image
                let src = this.src;
                
                //Get Base64 Data formatted image data
                toDataURL(src, function (dataUrl) {
                    let base64DataString = ';base64,';
                    let startBase64Data = dataUrl.toString().indexOf(base64DataString) + base64DataString.length;
                    console.log("dataUrl", dataUrl);
                    console.log("base64Data:", dataUrl.toString().substr(startBase64Data, dataUrl.toString().length));
                    insertPictureAtSelection(dataUrl.toString().substr(startBase64Data, dataUrl.toString().length));
                });
            });


        });
    };

    // Eine Hilfsfunktion zum Anzeigen von Benachrichtigungen.
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    //-------Image Functions---------
    //Insert image snipped from Office Apps Dev
    function insertPictureAtSelection(base64EncodedImageStr) {

        Office.context.document.setSelectedDataAsync(base64EncodedImageStr, {
            coercionType: Office.CoercionType.Image
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.log("Action failed with error: " + asyncResult.error.message);
                }
            });
    }
    //Sends Base64 String of an image to the Client
    function toDataURL(url, callback) {
        var xhr = new XMLHttpRequest();
        xhr.onload = function () {
            var reader = new FileReader();
            reader.onloadend = function () {
                console.log("reader.result", reader.result);
                callback(reader.result);
            }
            reader.readAsDataURL(xhr.response);
        };
        xhr.open('GET', url);
        xhr.responseType = 'blob';
        xhr.send();
    }
    //-------/Image Functions---------

    //-------Download File Functions---------
    var downloadUri = location.origin + '/Data/FoliensammlungJanusConsulting.pptx'; // PDF file to download from the remote server.
    

    function downloadFile() {
        // The following line is taken from http://stackoverflow.com/questions/7944460/detect-safari-browser
        if (/^((?!chrome|android).)*safari/i.test(navigator.userAgent)) {
            insertAndClickLink();
        }
        else {
            InsertURL();
            // First download the file, then prompt the user to save it. 
            // This function uses the file-saver.js library.
            getFile(downloadUri, saveDownloadAsFile);
        }
    }

    function getFile(url, callback) {
        // Use XHR to fetch the file from the remote server.
        var xhr = new XMLHttpRequest();
        xhr.onload = function () {
            try {
                if (xhr.status === 200) {
                    callback(xhr.response);
                }
                else if (xhr.status !== 200) {
                    console.log("getFile responseCode was not 200:", xhr);
                }
            } catch (e) {
                console.log("getFile produced an exception:", e);
            }
        };
        xhr.open('GET', url, true);
        xhr.responseType = 'blob';
        xhr.send();
    }

    function saveDownloadAsFile(file){
        saveAs(file, "FoliensammlungJanusConsulting.pptx");
    }


    function insertAndClickLink() {

        //Taken from http://stackoverflow.com/questions/3077242/force-download-a-pdf-link-using-javascript-ajax-jquery/29266135#29266135
        try {
            if (!window.ActiveXObject) {
                var save = document.createElement('a');
                save.href = downloadUri;
                save.target = '_blank';
                save.download = "myfile.pdf" || 'unknown';

                var evt = new MouseEvent('click', {
                    'view': window,
                    'bubbles': true,
                    'cancelable': false
                });

                save.dispatchEvent(evt);
                (window.URL || window.webkitURL).revokeObjectURL(save.href);
            }
        }
        catch (e) {
            console.error(e);
        }
    }

    /*
       This function inserts a URL into a DIV tag, which the user then clicks. This provides a good fallback experience. 
   */
    function InsertURL() {
        $("#InsertURL").append("<a href=\"" + downloadUri + "\" target=\"_blank\">PDF</a>");
        console.log("setting link; downloadUri:", downloadUri);
    }
    //-------/Download File Functions---------
        
})();
