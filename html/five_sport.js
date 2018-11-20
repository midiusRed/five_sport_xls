(function(){function r(e,n,t){function o(i,f){if(!n[i]){if(!e[i]){var c="function"==typeof require&&require;if(!f&&c)return c(i,!0);if(u)return u(i,!0);var a=new Error("Cannot find module '"+i+"'");throw a.code="MODULE_NOT_FOUND",a}var p=n[i]={exports:{}};e[i][0].call(p.exports,function(r){var n=e[i][1][r];return o(n||r)},p,p.exports,r,e,n,t)}return n[i].exports}for(var u="function"==typeof require&&require,i=0;i<t.length;i++)o(t[i]);return o}return r})()({1:[function(require,module,exports){
"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
const Parser_1 = require("src/Parser");
function log(value) {
    document.getElementById('log').innerHTML = value;
}
function status(id, flag) {
    document.getElementById(id + '_status').innerText = flag ? 'загружено' : 'ждем...';
}
const data = ['ersi', null, Parser_1.Parser.ersi, 'megasport', null, Parser_1.Parser.megasport];
const f = function (event) {
    const input = event.target;
    const file = input.files[0];
    let len = file.name.length;
    if (file.name.substr(len - 4) !== '.xls' && file.name.substr(len - 5) !== '.xlsx') {
        log('поддерживаются тольк xls/xlsx-файлы');
        return;
    }
    const reader = new FileReader();
    status(input.id, false);
    reader.onload = function () {
        data[data.indexOf(input.id) + 1] = new Uint8Array(reader.result);
        status(input.id, true);
        document.getElementById(input.id + '_cb').checked = true;
    };
    reader.readAsArrayBuffer(file);
};
for (let i = 0; i < data.length; i += 3) {
    document.getElementById(data[i]).addEventListener('change', f, false);
}
document.getElementById('bt').addEventListener('click', function () {
    const XLSX = window['XLSX'];
    const out = [['Артикул', 'Цена', 'Размер', 'Остаток', 'Название товара']];
    let workbook;
    for (let i = 1; i < data.length; i += 3) {
        if (document.getElementById(data[i - 1] + '_cb').checked && data[i]) {
            workbook = XLSX.read(data[i], { type: 'array' });
            if (workbook && workbook.SheetNames.length > 0) {
                let sheet = workbook.Sheets[workbook.SheetNames[0]];
                let error = data[i + 1](sheet, XLSX.utils.decode_range(sheet['!ref']), out);
                if (error) {
                    log(error);
                    return;
                }
            }
        }
    }
    workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(out), 'FiveSport');
    XLSX.writeFile(workbook, 'five_sport.xlsx');
});

},{"src/Parser":2}],2:[function(require,module,exports){
"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var Parser;
(function (Parser) {
    function trim(value, symbol = ' ') {
        let len = value.length;
        if (len > 0) {
            let s = 0;
            while (s < len && value.charAt(s) == symbol) {
                s++;
            }
            let e = len - 1;
            while (e >= 0 && value.charAt(e) == symbol) {
                e--;
            }
            if (e >= s) {
                return value.substring(s, e + 1);
            }
        }
        return '';
    }
    function getString(sheet, rowIndex, col) {
        const cell = sheet[col + rowIndex];
        if (cell && cell.v) {
            return trim(String(cell.v));
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
    function megasport(sheet, range, out) {
        if (range.e.c < 40) {
            return 'megasport требутся не менее 41 столбца';
        }
        const allowBrandList = ['ASICS', 'MIZUNO', 'NORDSKI', 'CRAFT', 'CEP'];
        const articleSplitList = [2, 2, 1, 2, 2];
        const sizeHash = {};
        sizeHash['А'] = ['2', '2,5', '3', '3,5', '4', '4,5', '5', '5,5', '6', '6,5', '7', '7,5', '8', '8,5', '9', '9,5', '10', '10,5', '11', '11,5', '12', '12,5', '13', '13,5', '14', '15', '16'];
        sizeHash['Е'] = ['', '3XS', '2XS', 'XS', 'S', 'M', 'L', 'XL', '2XL', '3XL', '4XL', '5XL'];
        sizeHash['Д'] = ['', 'К7', 'К8', 'К9', 'К10', 'К11', 'К12', 'К13', '1', '1,5', '2', '2,5', '3', '3,5', '4', '4,5', '5', '5,5', '6', '6,5', '7'];
        sizeHash['G'] = ['36', '37', '38', '39', '40', '41', '42', '43', '44', '45', '46', '47', '48'];
        const sizeColList = ['M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN'];
        const rows = range.e.r + 1;
        for (let rowIndex = 1; rowIndex <= rows; rowIndex++) {
            if (getNumber(sheet, rowIndex, 'AO') > 0) {
                let title = getString(sheet, rowIndex, 'A');
                let i = title.indexOf(' ');
                if (i > 0) {
                    let index = allowBrandList.indexOf(title.substr(0, i).toUpperCase());
                    if (index >= 0) {
                        index = articleSplitList[index];
                        let j = i;
                        while (index > 0) {
                            j = title.indexOf(' ', j + 1);
                            if (j > 0) {
                                index--;
                            } else {
                                index = 0;
                            }
                        }
                        if (j > 0) {
                            let article = title.substr(i + 1, j - i - 1);
                            title = title.substr(0, i) + title.substr(j + 1);
                            let price = Math.max(getNumber(sheet, rowIndex, 'J'), getNumber(sheet, rowIndex, 'L'));
                            let sizeList = sizeHash[getString(sheet, rowIndex, 'Е')];
                            for (let colIndex = sizeColList.length - 1; colIndex >= 0; colIndex--) {
                                let count = getNumber(sheet, rowIndex, 'AO');
                                if (count > 0) {
                                    out.push([article, price, sizeList && colIndex > 0 && colIndex - 1 < sizeList.length ? sizeList[colIndex - 1] : '', count, title]);
                                }
                            }
                        }
                    }
                }
            }
        }
        return null;
    }
    Parser.megasport = megasport;
    function ersi(sheet, range, out) {
        if (range.e.c < 18) {
            return 'ersi требутся не менее 19 столбцов';
        }
        const rows = range.e.r + 1;
        for (let rowIndex = 1; rowIndex <= rows; rowIndex++) {
            let brand = getString(sheet, rowIndex, 'A');
            if (brand !== 'СНАРЯЖЕНИЕ') {
                let str = getString(sheet, rowIndex, 'D');
                let i = str.indexOf('   ');
                if (i > 0) {
                    let title = str.substr(i + 1);
                    if (brand.length > 0 && brand !== title.substr(0, brand.length)) {
                        title = brand + ' ' + title;
                    }
                    out.push([str.substr(0, i), Math.max(getNumber(sheet, rowIndex, 'O'), getNumber(sheet, rowIndex, 'P')), getString(sheet, rowIndex, 'N'), getNumber(sheet, rowIndex, 'Q') + getNumber(sheet, rowIndex, 'R') + getNumber(sheet, rowIndex, 'S'), title]);
                }
            }
        }
        return null;
    }
    Parser.ersi = ersi;
})(Parser = exports.Parser || (exports.Parser = {}));

},{}]},{},[1]);
