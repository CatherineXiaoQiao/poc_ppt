
(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            $('#get-data-from-selection').click(getDataFromSelection);

            $("#get-shapes-only").click(getShapes);
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        var url = Office.context.document.url;
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                $("#r-div").text(asyncResult.error.message);
            } else {
                var selectedSlides = asyncResult.value.slides;
                if (!selectedSlides || selectedSlides.length == 0) return;
                var slideIndex = selectedSlides[0].index;
                convertToHtml(url, slideIndex);
            }
        });
    }

    function convertToHtml(pptUrl, slideIndex) {
        slideIndex = slideIndex - 1;
        $.ajax({
            type: 'GET',
            url: "/api/Presentation",
            success: function (response) {
                $("#h-div").text(response);
            },
            error: function () {
                $("#r-div").text("error");
            }
        });
    }

    function getShapes() {
        $.ajax({
            type: 'GET',
            url: "/api/Presentation/5",
            success: function (response) {
                var s = "<ul class='pf'>";
                for (var i = 0; i < response.length; i++) {
                    s += "<li><img src='/pf/" + response[i] + "'/><span>" + response[i]+"<span></li>";
                }
                s += "</ul>";
                $("#h-div").html(s);
            },
            error: function () {
                $("#r-div").text("error");
            }
        });
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
