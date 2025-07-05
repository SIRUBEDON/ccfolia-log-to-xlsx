const fileInput = document.getElementById('fileInput');
const convertButton = document.getElementById('convertButton');
const downloadLink = document.getElementById('downloadLink');

convertButton.addEventListener('click', convertFile);

function toggleDescription() {
    const description = document.getElementById('description');
    description.style.display = description.style.display === 'none' ? 'block' : 'none';
}

async function convertFile() {
    const file = fileInput.files[0];
    if (!file) {
        alert('ファイルを選択してください。');
        return;
    }

    const content = await file.text();
    const parser = new DOMParser();
    const doc = parser.parseFromString(content, 'text/html');

    const rows = Array.from(doc.querySelectorAll('body > p')).map(p => {
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

    const tabs = Array.from(new Set(rows.map(row => row[1])));
    const uniqueTabs = Array.from(new Set(['all', 'main', 'info', 'other', ...tabs]));

    const workbook = new ExcelJS.Workbook();
    const headers = ["発言者", "発言タブ", "発言内容"];

    uniqueTabs.forEach(tabName => {
        const sheet = workbook.addWorksheet(tabName);
        sheet.addRow(headers);
        let sheetData;
        if (tabName === 'all') {
            sheetData = rows;
        } else {
            sheetData = rows.filter(row => row[1] === tabName);
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
}