'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            loadItemProps(Office.context.mailbox.item);
        });
    });

    function loadItemProps(item) {
        // Write message property values to the task pane
        console.log('item:');
        console.log(item);
        $('#item-title').text('2025-02-14 11:30');
        $('#item-id').text(item.itemId);
        $('#item-subject').text(item.subject);
        $('#item-internetMessageId').text(item.internetMessageId);
        $('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");
        
        item.body.getAsync(Office.CoercionType.Html, (result) => {
            console.log('item.body.getAsync');
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Contenido del correo:", result.value);
                $('#item-html').html(result.value);
            } else {
                console.error("oError al obtener el cuerpo:", result.error);
            }
        });
    }
})();
