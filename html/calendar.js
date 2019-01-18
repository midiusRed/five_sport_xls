const newLineRE = new RegExp('\r?\n', 'g');
function getString(sheet, rowIndex, col) {
    const cell = sheet[col + rowIndex];
    if (cell && cell.v) {
        return String(cell.v).trim().replace(newLineRE, '<br>');
    }
    return '';
}
function getNumber(sheet, rowIndex, col) {
    const cell = sheet[col + rowIndex];
    if (cell && cell.t === 'n' && typeof cell.v === 'number') {
        return cell.v;
    }
    return 0;
}
const f = function (event) {
    const input = event.target;
    const file = input.files[0];
    let len = file.name.length;
    if (file.name.substr(len - 4) !== '.xls' && file.name.substr(len - 5) !== '.xlsx') {
        console.error('поддерживаются тольк xls/xlsx-файлы');
        return;
    }
    const reader = new FileReader();
    reader.onload = function () {
        if (!(reader.result instanceof ArrayBuffer)) {
            return;
        }
        const XLSX = window['XLSX'];
        let workbook = XLSX.read(new Uint8Array(reader.result), { type: 'array' });
        if (workbook && workbook.SheetNames.length > 0) {
            let str = '';
            let sheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.decode_range(sheet['!ref']).e.r + 1;
            for (let rowIndex = 7; rowIndex <= rows; rowIndex++) {
                let date = getString(sheet, rowIndex, 'E');
                let i = date.indexOf('-');
                if (i > 0) {
                    let date2 = date.substr(i + 1);
                    i = date2.indexOf('.');
                    let vD = parseInt(date2.substr(0, i));
                    let vM = parseInt(date2.substr(i + 1));
                    if (!isNaN(vD) && !isNaN(vM)) {
                        str += vD + '.' + vM;
                    }
                }
                str += ';' + getString(sheet, rowIndex, 'B') + ';' + getString(sheet, rowIndex, 'C') + ';' + getString(sheet, rowIndex, 'D') + ';' + date + ';' +
                    getString(sheet, rowIndex, 'F') + ';' + getString(sheet, rowIndex, 'G') + ';' + getString(sheet, rowIndex, 'H') + ';' + getString(sheet, rowIndex, 'I') + '\n';
            }
            console.log(str);
        }
    };
    reader.readAsArrayBuffer(file);
};
document.getElementById('file').addEventListener('change', f, false);
