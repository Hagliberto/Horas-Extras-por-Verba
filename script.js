document.addEventListener("DOMContentLoaded", function () {
    const fileInput = document.getElementById("file-input");
    const processButton = document.getElementById("process-button");
    processButton.addEventListener("click", function () {
      const file = fileInput.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = function (e) {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const worksheet = workbook.Sheets[workbook.SheetNames[0]];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  
          // Agrupa os dados com base nos valores da segunda coluna
          const groupedData = {};
          jsonData.forEach(function (row) {
            const value = row[1];
            if (!groupedData[value]) {
              groupedData[value] = [];
            }
            // Realiza o processamento e formatação dos dados
            if (typeof row[2] === "number") {
              row[2] = row[2].toFixed(2).replace(".", ",");
            } else if (typeof row[2] === "string") {
              row[2] = row[2].replace(".", ",");
            }
            row[0] = ("000000" + row[0]).slice(-6);
            groupedData[value].push(row);
          });
  
          // Cria uma planilha separada para cada grupo de dados
          for (const value in groupedData) {
            if (groupedData.hasOwnProperty(value)) {
              const groupData = groupedData[value];
              const newWorkbook = XLSX.utils.book_new();
              const newWorksheet = XLSX.utils.aoa_to_sheet(groupData);
  
              // Preserva a formatação existente copiando as propriedades de estilo de cada célula do worksheet original
              for (const key in worksheet) {
                if (
                  key !== "!ref" &&
                  worksheet.hasOwnProperty(key) &&
                  newWorksheet.hasOwnProperty(key)
                ) {
                  newWorksheet[key].s = worksheet[key].s;
                }
              }
  
              XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");
  
              // Converte o novo workbook em um arquivo Excel e inicia o download
              const newFileData = XLSX.write(newWorkbook, {
                bookType: "xlsx",
                type: "array",
              });
              const blob = new Blob([new Uint8Array(newFileData)], {
                type: "application/octet-stream",
              });
              const url = URL.createObjectURL(blob);
              const a = document.createElement("a");
              a.href = url;
              a.download = `Verba_${value}.xlsx`;
              a.click();
            }
          }
  
          console.log("Arquivos formatados e salvos com sucesso!");
        };
        reader.readAsArrayBuffer(file);
      } else {
        console.log("Nenhum arquivo selecionado.");
      }
    });
  });
  