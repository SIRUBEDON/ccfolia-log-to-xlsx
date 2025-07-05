const fileInput = document.getElementById('fileInput');
const convertButton = document.getElementById('convertButton');
const downloadLink = document.getElementById('downloadLink');
let parsedRows = []; // パース結果を保持する変数

fileInput.addEventListener('change', handleFileSelect);
convertButton.addEventListener('click', convertFile);

function toggleDescription() {
    const description = document.getElementById('description');
    description.style.display = description.style.display === 'none' ? 'block' : 'none';
}

async function handleFileSelect() {
    const file = fileInput.files[0];
    const outputOptions = document.getElementById('outputOptions');
    if (!file) {
        outputOptions.style.display = 'none';
        parsedRows = [];
        return;
    }

    const content = await file.text();
    const parser = new DOMParser();
    const doc = parser.parseFromString(content, 'text/html');

    parsedRows = Array.from(doc.querySelectorAll('body > p')).map(p => {
        const spans = p.querySelectorAll('span');
        const tabText = spans[0] ? spans[0].textContent.trim().replace(/^[\[]|[\]]$/g, '') : '';
        const speaker = spans[1] ? spans[1].textContent.trim().replace(/:$/, '') : '';
        const contentSpan = spans[2];
        const content = contentSpan ? contentSpan.innerHTML.replace(/<br\s*\/?>/g, '\n').trim() : '';
        let color = '';
        const styleAttr = p.getAttribute('style') || '';
        const colorMatch = styleAttr.match(/color\s*:\s*(#[0-9A-Fa-f]{3,6})/);
        if (colorMatch) {
            color = colorMatch[1];
            if (color.length === 4) {
                color = '#' + color[1].repeat(2) + color[2].repeat(2) + color[3].repeat(2);
            }
        }
        return [speaker, tabText, content, color];
    });

    displayTabSelection();
    outputOptions.style.display = 'block';
}

function displayTabSelection() {
    const tabs = Array.from(new Set(parsedRows.map(row => row[1])));
    const uniqueTabs = ['all', ...new Set(tabs.filter(t => t))];

    const tabSelectionDiv = document.getElementById('tabSelection');
    tabSelectionDiv.innerHTML = '<h4>変換するタブを選択してください:</h4>';
    uniqueTabs.forEach(tabName => {
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.id = tabName;
        checkbox.value = tabName;
        checkbox.checked = true;
        const label = document.createElement('label');
        label.htmlFor = tabName;
        label.appendChild(document.createTextNode(tabName));
        tabSelectionDiv.appendChild(checkbox);
        tabSelectionDiv.appendChild(label);
        tabSelectionDiv.appendChild(document.createElement('br'));
    });
}

async function convertFile() {
    if (parsedRows.length === 0) {
        alert('ファイルをまず選択してください。');
        return;
    }

    const selectedTabs = Array.from(document.querySelectorAll('#tabSelection input[type="checkbox"]:checked')).map(cb => cb.value);
    if (selectedTabs.length === 0) {
        alert('少なくとも1つのタブを選択してください。');
        return;
    }

    const outputFormat = document.querySelector('input[name="format"]:checked').value;

    if (outputFormat === 'xlsx') {
        const workbook = new ExcelJS.Workbook();
        const headers = ["発言者", "発言タブ", "発言内容"];

        selectedTabs.forEach(tabName => {
            const sheet = workbook.addWorksheet(tabName);
            sheet.addRow(headers);
            let sheetData;
            if (tabName === 'all') {
                sheetData = parsedRows;
            } else {
                sheetData = parsedRows.filter(row => row[1] === tabName);
            }
            sheetData.forEach(row => {
                const excelRow = sheet.addRow(row.slice(0, 3));
                if (row[3]) {
                    const hex = row[3].replace('#', '').toUpperCase();
                    excelRow.getCell(1).font = {
                        color: { argb: `FF${hex}` }
                    };
                }
            });
            sheet.columns.forEach((column, index) => {
                let maxLength = 0;
                column.eachCell({ includeEmpty: true }, (cell) => {
                    const columnLength = cell.value ? cell.value.toString().length : 10;
                    if (columnLength > maxLength) {
                        maxLength = columnLength;
                    }
                });
                column.width = Math.min(maxLength + 2, 100);
            });
            sheet.getRow(1).font = { bold: true };
            sheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
        });

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        const url = URL.createObjectURL(blob);
        downloadLink.innerHTML = `<a href="${url}" download="ccfolia_log_converted.xlsx">Excelファイルをダウンロード</a>`;
        downloadLink.style.display = 'block';
    } else if (outputFormat === 'csv') {
        let csvContent = "\uFEFF"; // BOM for UTF-8
        const headers = ["発言者", "発言タブ", "発言内容"];
        
        selectedTabs.forEach(tabName => {
            let sheetData;
            if (tabName === 'all') {
                sheetData = parsedRows;
            } else {
                sheetData = parsedRows.filter(row => row[1] === tabName);
            }

            if(selectedTabs.length > 1) {
                csvContent += `\n${tabName}\n`;
            }
            
            csvContent += headers.join(",") + "\r\n";

            sheetData.forEach(rowArray => {
                let row = rowArray.slice(0, 3).map(cell => `"${(cell || "").toString().replace(/"/g, '""')}"`).join(",");
                csvContent += row + "\r\n";
            });
        });

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        downloadLink.innerHTML = `<a href="${url}" download="ccfolia_log_converted.csv">CSVファイルをダウンロード</a>`;
        downloadLink.style.display = 'block';
    }
}