document.addEventListener("DOMContentLoaded", function () {
  const fileInput = document.getElementById("file-input");
  const processButton = document.getElementById("process-button");
  const errorContainer = document.getElementById("error-container");

  processButton.addEventListener("click", function () {
    const file = fileInput.files[0];
    if (file) {
      const fileExtension = file.name.split(".").pop();
      if (fileExtension !== "xlsx") {
        showError(
          "Erro: O arquivo selecionado deve ser um arquivo Excel (.xlsx)."
        );
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

            for (const value in groupedData) {
              if (groupedData.hasOwnProperty(value)) {
                const groupData = groupedData[value];
                const newWorkbook = XLSX.utils.book_new();
                const newWorksheet = XLSX.utils.aoa_to_sheet(groupData);

                for (const key in worksheet) {
                  if (
                    key !== "!ref" &&
                    worksheet.hasOwnProperty(key) &&
                    newWorksheet.hasOwnProperty(key)
                  ) {
                    newWorksheet[key].s = worksheet[key].s;
                  }
                }

                XLSX.utils.book_append_sheet(
                  newWorkbook,
                  newWorksheet,
                  "Sheet1"
                );

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

            showSuccess("Arquivos formatados e salvos com sucesso!");
          } catch (error) {
            showError("Erro ao processar o arquivo: " + error.message);
          }
        };
        reader.readAsArrayBuffer(file);
      } catch (error) {
        showError("Erro ao ler o arquivo: " + error.message);
      }
    } else {
      showError("Nenhum arquivo selecionado.");
    }

    if (!file) {
      const fileUploadMessage = document.getElementById("file-upload-message");
      fileUploadMessage.style.display = "block";
      return;
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
    }, 15000);
  }
});
