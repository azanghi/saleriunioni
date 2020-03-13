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
                var start;
                // Orario
                item.start.getAsync(
                    function (asyncResult) {
                        start = asyncResult.value.getHours() + ':' + asyncResult.value.getMinutes();
                    }
                );
                var item = Office.context.mailbox.item;
                item.end.getAsync(
                    function (asyncResult) {
                        $('#txtOrario').val(start + '-' + asyncResult.value.getHours() + ':' + asyncResult.value.getMinutes());
                    }
                );
                // Luogo (Sala)
                var item = Office.context.mailbox.item;
                item.location.getAsync(
                    function (asyncResult) {
                        $('#txtSala').val(asyncResult.value);
                    }
                );
                // Tabella di partecipanti (Obbligatorio e Facoltativo)
                var attendeeNumber = 0;
                Office.context.mailbox.item.requiredAttendees.getAsync(function (result) {
                    if (result.error) {
                        console.log(result.error);
                    } else {
                        var msg = "";
                        $('#tableAttendees').removeAttr('hidden');
                        result.value.forEach(function (recip, index) {
                            msg = msg + recip.displayName + " (" + recip.emailAddress + ");";
                            attendeeNumber++;
                            $('#rowsAttendees').append('<tr><th scope="row">'+attendeeNumber+'</th><td>'+recip.displayName+'</td><td><img src="/assets/check.png" alt="Example" width="30" height="30"></td><td>Obbligatorio</td></tr>');
                        });
                    }
                });
                Office.context.mailbox.item.optionalAttendees.getAsync(function (result) {
                    if (result.error) {
                        console.log(result.error);
                    } else {
                        var msg = "";
                        result.value.forEach(function (recip, index) {
                            msg = msg + recip.displayName + " (" + recip.emailAddress + ");";
                            attendeeNumber++;
                            $('#rowsAttendees').append('<tr><th scope="row">'+attendeeNumber+'</th><td>'+recip.displayName+'</td><td><img src="/assets/check.png" alt="Example" width="30" height="30"></td><td>Facoltativo</td></tr>');
                        });
                    }
                });
            });
        });
    };

})();