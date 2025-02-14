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
    const item = Office.context.mailbox.item;
    let resultString = '';
    console.log("defaultStatus Init...");
    if (item) {
        item.body.getAsync(Office.CoercionType.Html, (result) => {
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

  statusUpdate("icon16" , "Hi 20250213 8:13!!!", event);
}

function generatePDF2(htmlContent) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    console.log('holi');
    doc.html(htmlContent, {
        callback: function (pdf) {
            pdf.save("email.pdf");
            console.log('email.pdf');
        }
    });
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

