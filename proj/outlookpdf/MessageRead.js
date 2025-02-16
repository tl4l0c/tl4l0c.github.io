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
        $('#item-title').text('2025-02-15 20:10');
        //$('#item-id').text(item.itemId);
        $('#item-subject').text(item.subject);
        //$('#item-internetMessageId').text(item.internetMessageId);
        //$('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");
        
        item.body.getAsync(Office.CoercionType.Html, (result) => {
            console.log('item.body.getAsync');
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log('status ok');
                console.log('body:', result.value);
                // $('#item-html').html(result.value);
                generatePDF(result.value, item.subject);
            } else {
                console.error("Error al obtener el cuerpo:", result.error);
            }
        });
    }

    function generatePDF(htmlContent, subject) {
        console.log('generatePDF init.');
        const { jsPDF } = window.jspdf;

        if (!window.jspdf) {
            console.error("jsPDF no está cargado.");
        }
        if (!window.DOMPurify) {
            console.error("DOMPurify no está cargado.");
        }

        const doc = new jsPDF({
            orientation: "portrait",
            unit: "px",
            format: "a4"
        });

        console.log('generatePDF 1');

        let outlookHtml = ``;

        loadImageToBase64("https://cdn.graph.office.net/prod/media/shared/Microsoft_Logo_White.png", function (base64Image) {
            console.log(base64Image); // Reemplaza la imagen con su versión Base64
            outlookHtml = `
                <div style="width: 800px; margin: 10px auto;">
                <img src="${base64Image}" alt="Red dot" />
                     ${htmlContent}
                </div>
            `;
        });

        doc.html(outlookHtml, {
            callback: function (pdf) {
                console.log('generatePDF 2');
                pdf.save("Email." + formatFileName(subject) + ".pdf"); // Descarga el PDF automáticamente
            },
            x: 10,
            y: 10,
            html2canvas: {
                scale: 0.5,
                width: 800,
                useCORS: true
            } 
        });
        console.log('generatePDF end.');
    }

    function loadImageToBase64(url, callback) {
        const img = new Image();
        img.crossOrigin = "Anonymous";
        img.onload = function () {
            const canvas = document.createElement("canvas");
            canvas.width = img.width;
            canvas.height = img.height;
            const ctx = canvas.getContext("2d");
            ctx.drawImage(img, 0, 0);
            callback(canvas.toDataURL("image/png"));
        };
        img.onerror = function () {
            console.error("No se pudo cargar la imagen:", url);
        };
        img.src = url;
    }

    function formatFileName(subject) {
        subject = subject.replace(/[<>:"\/\\|?*\n\r]+/g, "");
        subject = subject.replace(/\s+/g, "");
        return subject.substring(0, 30);
    }

})();
