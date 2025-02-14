// Loads the Office.js library.
Office.onReady();

// Helper function to add a status message to the notification bar.
function statusUpdate(icon, text, event) {
    
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
function defaultStatus(event) {
    if (!window.jspdf) {
        console.log('!window.jspdf');
        const script = document.createElement("script");
        script.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
        script.onload = () => console.log("jsPDF cargado");
        document.head.appendChild(script);
    }

    const item = Office.context.mailbox.item;
    let resultString = '';
    console.log("defaultStatus Init...");
    if (item) {
        console.log('if (item)');
        item.body.getAsync(Office.CoercionType.Html, (result) => {
            console.log('item.body.getAsync');
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Contenido del correo:", result.value);
                resultString = "Contenido del correo: " + result.value;
                generatePDF(result.value);
            } else {
                console.error("Error al obtener el contenido:", result.error);
                resultString = "Error al obtener el contenido: " + result.error;
            }
        });
    }
    else {
        console.log('Empty item');
    }

  statusUpdate("icon16" , "Hi 20250213 8:52!!!", event);
}

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

