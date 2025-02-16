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
        $('#item-version').text('2025.02.16.11.18');
        //$('#item-id').text(item.itemId);
        $('#item-subject').text(item.subject);
        //$('#item-internetMessageId').text(item.internetMessageId);
        //$('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");

        let from = item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;";
        let to = item.to.displayName + " &lt;" + item.to.emailAddress + "&gt;";
        let cc = '';
        let bcc = '';
        if (item.cc.length != 0) {
            cc = item.cc.displayName + " &lt;" + item.cc.emailAddress + "&gt;";
        }
        if (item.bcc.length != 0) {
            bcc = item.bcc.displayName + " &lt;" + item.bcc.emailAddress + "&gt;";
        }
        
        item.body.getAsync(Office.CoercionType.Html, (result) => {
            console.log('item.body.getAsync');
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log('status ok');
                console.log('body:', result.value);
                // $('#item-html').html(result.value);
                generatePDF(result.value, item.subject, from, to, cc, bcc);
            } else {
                console.error("Error al obtener el cuerpo:", result.error);
            }
        });
    }

    function generatePDF(htmlContent, subject, from, to, cc, bcc) {
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

        outlookHtml = `
                <div style="width: 800px; margin: 10px auto;">
                    <table>
                        <tr>
                            <td>From:</td>
                            <td>${from}</td>
                        </tr>
                        <tr>
                            <td>To:</td>
                            <td>${to}</td>
                        </tr>
                        <tr>
                            <td>cc:</td>
                            <td>${cc}</td>
                        </tr>
                        <tr>
                            <td>bcc:</td>
                            <td>${bcc}</td>
                        </tr>
                         <tr>
                            <td>Subject:</td>
                            <td>${subject}</td>
                        </tr>
                    </table>

                     ${htmlContent}
                </div>
            `;

        /*
        getBase64Image("https://cdn.graph.office.net/prod/media/shared/Microsoft_Logo_White.png", function (base64Image) {
            console.log(base64Image); // Reemplaza la imagen con su versión Base64
            outlookHtml = `
                <div style="width: 800px; margin: 10px auto;">
                <img src="${base64Image}" alt="Red dot" />
                     ${htmlContent}
                </div>
            `;
        });
        */

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

    function getBase64Image(url, callback) {
        fetch(`http://lets.mx/api/img64.php?url=${encodeURIComponent(url)}`)
            .then(response => response.json())
            .then(data => {
                if (data.base64) {
                    callback(data.base64);
                } else {
                    console.error("Error en la conversión:", data.error);
                }
            })
            .catch(error => console.error("Error en la solicitud:", error));
    }

    function formatFileName(subject) {
        subject = subject.replace(/[<>:"\/\\|?*\n\r]+/g, "");
        subject = subject.replace(/\s+/g, "");
        return subject.substring(0, 30);
    }

})();
