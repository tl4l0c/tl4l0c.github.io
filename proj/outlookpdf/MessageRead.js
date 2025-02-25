﻿'use strict';

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
        $('#item-version').text('2025.02.17.18.55');
        //$('#item-id').text(item.itemId);
        $('#item-subject').text(item.subject);
        //$('#item-internetMessageId').text(item.internetMessageId);
        //$('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");

        let from = item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;";
        let to = '';
        let cc = '';
        let bcc = '';

        let totalTo = item.to.length;
        let totalCc = item.cc.length;
        let totalBcc = item.bcc.length;

        for (let toIdx = 0; toIdx < totalTo; toIdx++) {
            to += item.to[toIdx].displayName + " &lt;" + item.to[toIdx].emailAddress + "&gt; <br />";
        }

        if (totalCc != 0) {
            for (let toIdx = 0; toIdx < totalCc; toIdx++) {
                cc += item.cc[toIdx].displayName + " &lt;" + item.cc[toIdx].emailAddress + "&gt; <br />";
            }
        }
        if (totalBcc != 0) {
            for (let toIdx = 0; toIdx < totalBcc; toIdx++) {
                bcc += item.bcc[toIdx].displayName + " &lt;" + item.bcc[toIdx].emailAddress + "&gt; <br />";
            }
        }
        
        item.body.getAsync(Office.CoercionType.Html, (result) => {
            console.log('item.body.getAsync');
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log('status ok');
                console.log('body:', result.value);
                // $('#item-html').html(result.value);

                getAccessToken();
                /*
                getAttachments((attachments) => {
                    generatePDF(result.value, item.subject, from, to, cc, bcc, attachments);
                });
                */
            } else {
                console.error("Error al obtener el cuerpo:", result.error);
            }
        });
    }

    async function getAccessToken() {
        console.log('getAccessToken Init');
        const clientId = "78283a7f-c3ed-4dc2-9b04-0f411555145a"; // Reemplázalo con "Application (client) ID"
        const tenantId = "faf9c572-8934-408f-bffb-41ffaee3edc4"; // Reemplázalo con "Directory (tenant) ID"

        const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`;
        console.log('authUrl:', authUrl);
        const params = new URLSearchParams({
            client_id: clientId,
            response_type: "token",
            redirect_uri: "https://login.microsoftonline.com/common/oauth2/nativeclient",
            scope: "https://graph.microsoft.com/.default offline_access"
        });

        window.open(`${authUrl}?${params.toString()}`, "_blank");
        console.log('getAccessToken End');
    }

    function getAuthToken(callback) {
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
            if (result.status === "succeeded") {
                callback(result.value); // Token obtenido
            } else {
                console.error("Error al obtener el token de autenticación:", result.error);
                callback(null);
            }
        });
    }

    function downloadAttachment(attachment, token, callback) {
        const baseUrl = Office.context.mailbox.restUrl;
        const attachmentUrl = `${baseUrl}/v2.0/me/messages/${Office.context.mailbox.item.itemId}/attachments/${attachment.id}/$value`;
        console.log('baseUrl:', baseUrl);
        console.log('attachmentUrl:', attachmentUrl);
        fetch(attachmentUrl, {
            method: "GET",
            headers: {
                "Authorization": `Bearer ${token}`,
                "Accept": "application/octet-stream"
            }
        })
            .then(response => {
                if (!response.ok) {
                    throw new Error(`Error al descargar el archivo: ${response.status}`);
                }
                return response.arrayBuffer();
            })
            .then(arrayBuffer => {
                callback(arrayBuffer);
            })
            .catch(error => {
                console.error("Error descargando el adjunto:", error);
                callback(null);
            });
    }


    function getAttachments(callback) {
        const item = Office.context.mailbox.item;

        if (!item || !item.attachments || item.attachments.length === 0) {
            console.log("No hay archivos adjuntos.");
            callback([]); // Llamar callback con un array vacío para evitar errores
            return;
        }

        // Filtrar solo archivos PDF
        const attachments = item.attachments.filter(att => att.name.toLowerCase().endsWith(".pdf"));

        if (attachments.length === 0) {
            console.log("No hay archivos PDF adjuntos.");
            callback([]); // No hay PDFs, continuar con la generación del PDF sin adjuntos
            return;
        }

        console.log("Archivos PDF adjuntos encontrados:", attachments);

        // Obtener el token de autenticación para descargar los archivos
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
            if (result.status !== "succeeded") {
                console.error("Error al obtener el token de autenticación:", result.error);
                callback([]); // No se puede descargar nada sin token
                return;
            }

            const token = result.value;
            const downloadedAttachments = [];

            // Descargar cada PDF adjunto
            let count = 0;
            attachments.forEach(attachment => {
                downloadAttachment(attachment, token, function (pdfBytes) {
                    if (pdfBytes) {
                        downloadedAttachments.push({
                            name: attachment.name,
                            data: pdfBytes
                        });
                    }

                    count++;

                    // Si ya descargamos todos los adjuntos, llamamos al callback
                    if (count === attachments.length) {
                        callback(downloadedAttachments);
                    }
                });
            });
        });
    }


    async function generatePDF(htmlContent, subject, from, to, cc, bcc, attachments) {
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
                            <td colspan="2"><hr /></td>
                        </tr>

                        <tr>
                            <td>To:</td>
                            <td>${to}</td>
                        </tr>

                        <tr>
                            <td colspan="2"><hr /></td>
                        </tr>

                        <tr>
                            <td>Cc:</td>
                            <td>${cc}</td>
                        </tr>

                        <tr>
                            <td colspan="2"><hr /></td>
                        </tr>

                        <tr>
                            <td>Bcc:</td>
                            <td>${bcc}</td>
                        </tr>

                        <tr>
                            <td colspan="2"><hr /></td>
                        </tr>

                        <tr>
                            <td>Subject:</td>
                            <td>${subject}</td>
                        </tr>
                    </table>

                     ${htmlContent}
                </div>
            `;

        console.log('generatePDF end.');

        return new Promise((resolve) => {
            doc.html(outlookHtml, {
                callback: async function (pdf) {
                    console.log('generatePDF 2');
                    const pdfBytes = pdf.output("arraybuffer");

                    getAuthToken(async (token) => {
                        if (!token) {
                            console.error("No se pudo obtener el token de autenticación.");
                            return;
                        }

                        const mergedPdfBytes = await mergePDFs(pdfBytes, attachments, token);

                        const blob = new Blob([mergedPdfBytes], { type: "application/pdf" });
                        const url = URL.createObjectURL(blob);
                        const a = document.createElement("a");
                        a.href = url;
                        a.download = `Email_${formatFileName(subject)}.pdf`;
                        document.body.appendChild(a);
                        a.click();
                        document.body.removeChild(a);

                        resolve();
                    });
                },
                x: 10,
                y: 10,
                html2canvas: { scale: 0.5, width: 800, useCORS: true }
            });
        });
    }


    async function mergePDFs(emailPdfBytes, attachments, token) {
        console.log("Iniciando combinación de PDFs...");

        if (typeof PDFLib === "undefined") {
            console.error("Error: PDFLib no está definido.");
            return emailPdfBytes;
        }

        const mergedPdf = await PDFLib.PDFDocument.create();
        const mainPdf = await PDFLib.PDFDocument.load(emailPdfBytes);
        const copiedPages = await mergedPdf.copyPages(mainPdf, mainPdf.getPageIndices());
        copiedPages.forEach((page) => mergedPdf.addPage(page));

        for (const attachment of attachments) {
            if (attachment.name.endsWith(".pdf")) {
                console.log(`Descargando: ${attachment.name}`);

                const pdfBytes = await new Promise((resolve) => {
                    downloadAttachment(attachment, token, resolve);
                });

                if (pdfBytes) {
                    try {
                        const attachmentPdf = await PDFLib.PDFDocument.load(pdfBytes);
                        const pages = await mergedPdf.copyPages(attachmentPdf, attachmentPdf.getPageIndices());
                        pages.forEach((page) => mergedPdf.addPage(page));
                    } catch (error) {
                        console.error(`Error procesando el PDF ${attachment.name}:`, error);
                    }
                } else {
                    console.warn(`No se pudo descargar el archivo: ${attachment.name}`);
                }
            }
        }

        return await mergedPdf.save();
    }


    async function downloadAttachmentAsBinary(attachment) {
        return new Promise((resolve, reject) => {
            Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
                if (result.status === "succeeded") {
                    const token = result.value;
                    const attachmentUrl = Office.context.mailbox.ewsUrl + "/ews/Exchange.asmx";

                    fetch(attachmentUrl, {
                        method: "POST",
                        headers: {
                            "Authorization": "Bearer " + token,
                            "Content-Type": "application/json"
                        },
                        body: JSON.stringify({
                            // Aquí se arma la petición para descargar el adjunto en binario
                        })
                    })
                        .then(response => response.arrayBuffer()) // Convertir a ArrayBuffer
                        .then(pdfBytes => resolve(pdfBytes))
                        .catch(error => {
                            console.error("Error descargando adjunto:", error);
                            resolve(null);
                        });
                } else {
                    console.error("No se pudo obtener el token:", result.error);
                    resolve(null);
                }
            });
        });
    }


    function generatePDF_onlypdf(htmlContent, subject, from, to, cc, bcc) {
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
                            <td colspan="2"><hr /></td>
                        </tr>

                        <tr>
                            <td>To:</td>
                            <td>${to}</td>
                        </tr>

                        <tr>
                            <td colspan="2"><hr /></td>
                        </tr>

                        <tr>
                            <td>Cc:</td>
                            <td>${cc}</td>
                        </tr>

                        <tr>
                            <td colspan="2"><hr /></td>
                        </tr>

                        <tr>
                            <td>Bcc:</td>
                            <td>${bcc}</td>
                        </tr>

                        <tr>
                            <td colspan="2"><hr /></td>
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
