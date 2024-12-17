document.addEventListener("DOMContentLoaded", () => {
    const inputFolder = document.getElementById("input-folder");
    const outputFileName = "output.xlsx";
    const fileInput = document.getElementById("file-input");
    const parseButton = document.getElementById("parse-button");
    const downloadLink = document.getElementById("download-link");
    const status = document.getElementById("status");

    let files = [];
    let currentIndex = 0;
    let combinedData = [];
    let stepCounter = 1;

    // При загрузке файлов
    fileInput.addEventListener("change", (event) => {
        files = Array.from(event.target.files);
        status.textContent = `Загружено файлов: ${files.length}`;
    });

    // Парсинг файлов
    parseButton.addEventListener("click", async () => {
        if (files.length === 0) {
            status.textContent = "Нет загруженных файлов.";
            return;
        }

        status.textContent = "Начинаем обработку файлов...";
        await processNextFile();
    });

    async function processNextFile() {
        if (currentIndex >= files.length) {
            status.textContent = "Обработка всех файлов завершена.";
            generateOutputFile();
            return;
        }

        const file = files[currentIndex];
        status.textContent = `Обрабатывается файл: ${file.name}`;

        try {
            const data = await readExcelFile(file);
            const suiteName = file.name.split(".")[0];

            // Удаление дубликатов и пустых строк
            const uniqueData = removeDuplicatesAndEmpty(data);

            // Создание блока данных
            const block = uniqueData.map((row, index) => ({
                Сьют: index === 0 ? suiteName : "",
                Наименование: row["Наименование"],
                Шаги: `Шаг ${stepCounter + index}`
            }));

            combinedData = [...combinedData, ...block];
            stepCounter += uniqueData.length;

            // Генерация промежуточного файла
            generateOutputFile();

            currentIndex++;
            processNextFile();
        } catch (error) {
            status.textContent = `Ошибка обработки файла: ${file.name}. ${error.message}`;
        }
    }

    function readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (event) => {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: "array" });
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                // Преобразование в массив объектов
                const headers = json[0];
                const rows = json.slice(1);
                const result = rows.map((row) => {
                    const obj = {};
                    headers.forEach((header, i) => {
                        obj[header] = row[i];
                    });
                    return obj;
                });
                resolve(result);
            };
            reader.onerror = (error) => reject(error);
            reader.readAsArrayBuffer(file);
        });
    }

    function removeDuplicatesAndEmpty(data) {
        const seen = new Set();
        return data.filter((row) => {
            const name = row["Наименование"];
            if (!name || seen.has(name)) return false;
            seen.add(name);
            return true;
        });
    }

    function generateOutputFile() {
        const worksheet = XLSX.utils.json_to_sheet(combinedData, {
            header: ["Сьют", "Наименование", "Шаги"]
        });
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Output");

        const outputData = XLSX.write(workbook, {
            bookType: "xlsx",
            type: "array"
        });

        const blob = new Blob([outputData], { type: "application/octet-stream" });
        const url = URL.createObjectURL(blob);

        downloadLink.href = url;
        downloadLink.download = outputFileName;
        downloadLink.textContent = "Скачать результат";
        downloadLink.style.display = "block";
    }
});
