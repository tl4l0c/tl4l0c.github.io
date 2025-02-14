// Loads the Office.js library.
Office.onReady().then(() => {
    console.log("Office está listo.");
    const item = Office.context.mailbox.item;

    if (item) {
        console.log('onReady -- setTimeout :: item');
        item.body.getAsync(Office.CoercionType.Html, (result) => {
            console.log('onReady -- setTimeout :: item.body.getAsync');
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("onReady -- Contenido del correo:", result.value);
            } else {
                console.error("onReady -- Error al obtener el cuerpo:", result.error);
            }
        });
    } else {
        console.error("onReady -- El correo aún no está disponible.");
    }
    console.log("Office end.");
});

function defaultStatus(event) {
    statusUpdate("icon16", "Hi 20250214 09:25!!!", event);
}


// Helper function to add a status message to the notification bar.
function statusUpdate(icon, text, event) {
    if (!window.jspdf) {
        console.log('!window.jspdf');
        const script = document.createElement("script");
        script.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
        script.onload = () => console.log("jsPDF cargado");
        document.head.appendChild(script);
    }


    setTimeout(() => {
        console.log('setTimeout Init');
        const item = Office.context.mailbox.item;
       
        if (item) {
            console.log('setTimeout :: item');
            item.body.getAsync(Office.CoercionType.Html, (result) => {
                console.log('setTimeout :: item.body.getAsync');
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("Contenido del correo:", result.value);
                } else {
                    console.error("Error al obtener el cuerpo:", result.error);
                }
            });
        } else {
            console.error("El correo aún no está disponible.");
        }
    }, 2000);



    const item = Office.context.mailbox.item;


    if (!item) {
        console.error("No se puede acceder al correo. El complemento podría estar en modo de redacción (Compose Mode).");
    }

    console.log("Modo del complemento:", Office.context.mailbox.diagnostics.hostName);
    console.log("Tipo de item:", item.itemType);




    let resultString = '';
    console.log('item: ', item);
    console.log("defaultStatus Init...");
    if (item) {


        item.body.getAsync(Office.CoercionType.Text, (result) => {
            console.log('item.body.getAsync.Text');
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Texto del correo:", result.value);
            } else {
                console.error("Error obteniendo el cuerpo:", result.error);
            }
        });

        console.log('if (item)');
        item.body.getAsync(Office.CoercionType.Html, (result) => {
            console.log('item.body.getAsync.Html');
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Contenido del correo:", result.value);
                resultString = "Contenido del correo: " + result.value;
                //generatePDF(result.value);
            } else {
                console.error("Error al obtener el contenido:", result.error);
                resultString = "Error al obtener el contenido: " + result.error;
            }
        });
    }
    else {
        console.log('Empty item');
    }  
  const details = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    icon: icon,
    message: text,
    persistent: false
  };
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", details, { asyncContext: event }, asyncResult => {
    const event = asyncResult.asyncContext;
    event.completed();
  });
}
// Displays a notification bar.


function generatePDF(htmlContent) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    doc.html(htmlContent, {
        callback: function (pdf) {
            console.log('Generar el PDF como un Blob');
            const pdfBlob = pdf.output("blob");

            console.log('Crear una URL del Blob');
            const url = URL.createObjectURL(pdfBlob);

            console.log('Crear un enlace invisible para forzar la descarga');
            const a = document.createElement("a");
            a.href = url;
            a.download = "correo.pdf";
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        },
        x: 10,
        y: 10
    });
}

// Maps the function name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("defaultStatus", defaultStatus);

