document.addEventListener("DOMContentLoaded", function () {
  const fileInput = document.getElementById("file-input");
  const processButton = document.getElementById("process-button");
  const errorContainer = document.getElementById("error-container");

  processButton.addEventListener("click", function () {
    const file = fileInput.files[0];

    if (!file) {
      showError("Nenhum arquivo selecionado.");
      return;
    }

    try {
      const reader = new FileReader();
      reader.onload = function (e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const worksheet = workbook.Sheets[workbook.SheetNames[0]];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

          const groupedData = {};

          jsonData.forEach(function (row) {
            const value = row[1];
            if (!groupedData[value]) {
              groupedData[value] = [];
            }
            if (typeof row[2] === "number") {
              row[2] = row[2].toFixed(2).replace(",", ".");
            } else if (typeof row[2] === "string") {
              row[2] = row[2].replace(",", ".");
            }
            row[0] = ("000000" + row[0]).slice(-6);
            groupedData[value].push(row);
          });

          const successMessage = "Arquivo formatado e salvo com sucesso!";
          const allData = [];

          for (const value in groupedData) {
            if (groupedData.hasOwnProperty(value)) {
              const groupData = groupedData[value];
              const newWorkbook = XLSX.utils.book_new();
              const newWorksheet = XLSX.utils.aoa_to_sheet(groupData);

              XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");

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

              allData.push(...groupData);
            }
          }

          // Generate Verba_geral.xlsx with all data
          const allWorkbook = XLSX.utils.book_new();
          const allWorksheet = XLSX.utils.aoa_to_sheet(allData);
          XLSX.utils.book_append_sheet(allWorkbook, allWorksheet, "Verba_Geral");

          const allFileData = XLSX.write(allWorkbook, {
            bookType: "xlsx",
            type: "array",
          });
          const allBlob = new Blob([new Uint8Array(allFileData)], {
            type: "application/octet-stream",
          });
          const allUrl = URL.createObjectURL(allBlob);
          const allA = document.createElement("a");
          allA.href = allUrl;
          allA.download = "Verba_geral.xlsx";
          allA.click();

          showSuccess(successMessage);
        } catch (error) {
          showError("Erro ao processar o arquivo.");
          console.error(error);
        }
      };
      reader.readAsArrayBuffer(file);
    } catch (error) {
      showError("Erro ao ler o arquivo.");
      console.error(error);
    }
  });

  function showError(errorMessage) {
    const errorContainer = document.getElementById("error-container");

    errorContainer.textContent = errorMessage;
    errorContainer.style.display = "block";

    // Recarrega a página após 5 segundos
    setTimeout(function () {
      location.reload();
    }, 10000);
  }

  function showSuccess(message) {
    const successContainer = document.getElementById("success-container");
    successContainer.textContent = message;
    successContainer.style.display = "block";

    // Oculta a mensagem de sucesso após 5 segundos
    setTimeout(function () {
      location.reload();
    }, 10000);
  }
});