async function fetchExcelData() {
    try {
        const url = "https://raw.githubusercontent.com/klef-ece/student_information/main/data.xlsx";
        const response = await fetch(url);
        if (!response.ok) throw new Error("Failed to fetch Excel file");

        const data = await response.arrayBuffer();
        const workbook = XLSX.read(data, { type: "array" });

        let combinedData = [];

        workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            if (sheetData.length > 1) {
                combinedData.push({
                    headers: sheetData[0],
                    rows: sheetData.slice(1)
                });
            }
        });

        return combinedData;

    } catch (error) {
        console.error("❌ Error loading Excel file:", error);
        return [];
    }
}

async function searchData() {
    const globalSearch = document.getElementById("searchInputGlobal").value.toLowerCase().trim();

    const searchTerms = [
        document.getElementById("searchInput1").value.toLowerCase().trim(),
        document.getElementById("searchInput2").value.toLowerCase().trim(),
        document.getElementById("searchInput3").value.toLowerCase().trim(),
        document.getElementById("searchInput4").value.toLowerCase().trim(),
        document.getElementById("searchInput5").value.toLowerCase().trim(),
        document.getElementById("searchInput6").value.toLowerCase().trim()
    ];

    const allEmpty = searchTerms.every(term => term === "") && globalSearch === "";

    const table = document.getElementById("dataTable");
    const tableHead = document.getElementById("tableHead");
    const tableBody = document.getElementById("tableBody");
    const noMessage = document.getElementById("noSearchMessage");

    if (allEmpty) {
        table.style.display = "none";
        noMessage.style.display = "block";
        return;
    }

    const allSheetsData = await fetchExcelData();

    tableHead.innerHTML = "";
    tableBody.innerHTML = "";

    let foundResults = false;

    allSheetsData.forEach(({ headers, rows }) => {
        const matchingRows = rows.filter(row => {
            const columnMatch = searchTerms.every((term, index) => {
                if (!term) return true;
                const cell = row[index] ?? "";
                return cell.toString().toLowerCase().includes(term);
            });

            const globalMatch = !globalSearch || row.some(cell =>
                cell?.toString().toLowerCase().includes(globalSearch)
            );

            return columnMatch && globalMatch;
        });

        if (matchingRows.length > 0) {
            foundResults = true;

            // Add table headers (only once)
            if (tableHead.childElementCount === 0) {
                headers.forEach(header => {
                    const th = document.createElement("th");
                    th.textContent = header;
                    tableHead.appendChild(th);
                });
            }

            // Add matching rows
            matchingRows.forEach(row => {
                const tr = document.createElement("tr");

                const containsCGPA = row.some(cell =>
                    cell?.toString().toLowerCase().includes("cgpa")
                );
                if (containsCGPA) {
                    tr.classList.add("highlight-row");
                }

                headers.forEach((_, index) => {
                    const td = document.createElement("td");
                    td.textContent = row[index] ?? "";
                    tr.appendChild(td);
                });

                tableBody.appendChild(tr);
            });
        }
    });

    if (foundResults) {
        table.style.display = "table";
        noMessage.style.display = "none";
    } else {
        tableHead.innerHTML = "";
        tableBody.innerHTML = `<tr><td colspan="100%">No matching results found.</td></tr>`;
        table.style.display = "table";
        noMessage.style.display = "none";
    }
}
