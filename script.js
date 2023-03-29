<!DOCTYPE html>
<html>
  <head>
    <meta charset="UTF-8">
    <title>Converter PDF para Excel</title>
  </head>
  <body>
    <input type="file" id="pdf-file">
    <button id="convert-button">Converter</button>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.11.338/pdf.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.16.8/xlsx.full.min.js"></script>
    <script>
      let pdfData = []; // Variável global para armazenar dados do PDF

      const pdfFileInput = document.getElementById('pdf-file');
      const convertButton = document.getElementById('convert-button');

      convertButton.addEventListener('click', () => {
        const pdfFile = pdfFileInput.files[0];

        // Carrega o arquivo PDF usando a biblioteca pdf.js
        pdfjsLib.getDocument(URL.createObjectURL(pdfFile)).promise.then(pdf => {
          const numPages = pdf.numPages;

          // Itera sobre cada página do PDF
          let pageDataPromises = [];
          for (let pageNum = 1; pageNum <= numPages; pageNum++) {
            pageDataPromises.push(pdf.getPage(pageNum).then(page => {
              return page.getTextContent().then(content => {
                // Extrai o texto da página e formata em um array de arrays para o sheet.js
                let pageData = content.items.map(item => item.str.trim());
                pdfData.push(pageData);
              });
            }));
          }

          // Aguarda a extração de todas as páginas do PDF antes de converter em um arquivo Excel
          Promise.all(pageDataPromises).then(() => {
            const workbook = XLSX.utils.book_new();
            const worksheet = XLSX.utils.aoa_to_sheet(pdfData);
            XLSX.utils.book_append_sheet(workbook, worksheet, 'PDF para Excel');
            const excelDataArrayBuffer = XLSX.write(workbook, { type: 'array' });

            // Salva o arquivo Excel em um blob e cria um link para download
            const excelBlob = new Blob([excelDataArrayBuffer], { type: 'application/octet-stream' });
            const downloadLink = document.createElement('a');
            downloadLink.href = URL.createObjectURL(excelBlob);
            downloadLink.download = `${pdfFile.name}.xlsx`;
            downloadLink.click();
          });
        });
      });
    </script>
  </body>
</html>
