// $('#verifica').on('click', function () {
//     console.log("gino");
// });

(function () {

    'use strict';

    Office.onReady();
    Office.initialize = function () {

        jQuery(document).ready(function () {

            $('#verifica').on('click', function () {
                console.log("gino");
                var item = Office.context.mailbox.item;
                item.subject.getAsync(
                    function (asyncResult) {
                        $('#txOrario').val(asyncResult.value);
                    });

            });
        });
    };

})();