namespace Parser { 	
	const ersiColumns:Array<string> = [
		'A'  /* Марка (бренд) */,
		'F'  /* Номенклатура */,
		'N'  /* Характеристика */,
		'O'  /* 01. РРЦ */,
		'P'  /* 08.МЕГАДИЛЕР */,
		'Q'  /* Магазин №2 (КЗ) */,
		''   /* ОСНОВНОЙ СКЛАД 3 */,
		'R'  /* ОСНОВНОЙ СКЛАД 6 */
	];

	interface Cell {
		v:string | number;
		t:string;
		w:string;
	}

	function trim(value:string, symbol:string = ' '):string {
		let len:number = value.length;
		if (len > 0) {
			let s:number = 0;
			while (s < len && value.charAt(s) == symbol) {
				s++;
			}
			let e:number = len - 1;
			while (e >= 0 && value.charAt(e) == symbol) {
				e--;
			}
			if (e >= s) {
				return value.substring(s, e + 1);
			}
		}
		return '';
	}

	function getString(sheet, rowIndex:number, col:string):string {
		const cell:Cell = sheet[col + rowIndex];
		if (cell && cell.v) {
			return trim(String(cell.v));
		}
		return '';
	}

	function getNumber(sheet, rowIndex:number, col:string):number {
		const cell:Cell = sheet[col + rowIndex];
		if (cell && cell.t === 'n' && typeof cell.v === 'number') {
			return cell.v;
		}
		return 0;
	}

	export function megasport(sheet, range, out:Array<Array<any>>):string {
		//ASICS 1051A002 400
		//MIZUNO V1GA1612 71
		//ADIDAS B40807 
		//SALOMON L38318100 
		//JOMA DRIW.815.IN 
		//MIKASA MT350 0003
		//TORNADO T312 2650 
		//GIVOVA AINF05 0004
		//MEGASPORT MS609Z1 0050
		//ERREA A245000009
		//UMBRO 350118 09S
		//NORDSKI NSM435700 
		//ERREA D505000009 
		//ZASPORT OFA217-063/001-BLU
		//CRAFT 1903716 B999 
		//WILSON WTB1033XB 
		//SELECT 810015 052 
		//TORRES AL00221
		//KV.REZAC 15015801 
		//GALA XX41009
		//MUELLER 130104 
		//POWERUP 00412 
		//MACRON 49035 
		//BUFF 110992.522.10.00
		//EXENZA G01 EMPIRE
		//HEAD 285511 
		//SKINS ZB99320059001 A400 
		//CEP C188M 5 
		//нужны asics, mizuno, nordski, craft, cep
		if (range.e.c < 40) {
			return 'megasport требутся не менее 41 столбца';
		}

		const allowBrandList:Array<string> = ['ASICS', 'MIZUNO', 'NORDSKI', 'CRAFT', 'CEP'];
		const articleSplitList:Array<number> = [2, 2, 1, 2, 2];
		const sizeHash:{ [kind:string]:Array<string> } = {};
		sizeHash['А'] = ['1', '1.5', '2',   '2.5', '3',   '3.5', '4',   '4.5', '5',   '5.5', '6',   '6.5', '7',   '7.5', '8', '8.5', '9', '9.5', '10', '10.5', '11', '11.5', '12', '12.5', '13', '13.5', '14', '15', '16'];
		sizeHash['Е'] = ['',  '3XS', '2XS', 'XS',  'S',   'M',   'L',   'XL',  '2XL', '3XL', '4XL', '5XL'];
		sizeHash['Д'] = ['',  'К7',  'К8',  'К9',  'К10', 'К11', 'К12', 'К13', '1',   '1.5', '2',   '2.5', '3',   '3.5', '4', '4.5', '5', '5.5', '6', '6.5', '7'];
		sizeHash['G'] = ['',  '36',  '37',  '38',  '39',  '40',  '41',  '42',  '43',  '44',  '45',  '46',  '47',  '48'];
		const sizeColList:Array<string> = ['M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO'];

		const rows:number = range.e.r + 1;
		for (let rowIndex:number = 1; rowIndex <= rows; rowIndex++) {
			if (getNumber(sheet, rowIndex, 'AP') > 0) {
				let title:string = getString(sheet, rowIndex, 'A');
				let i:number = title.indexOf(' ');
				if (i > 0) {
					let index:number = allowBrandList.indexOf(title.substr(0, i).toUpperCase());
					if (index >= 0) {
						index = articleSplitList[index];
						let j:number = i;
						while (index > 0) {
							j = title.indexOf(' ', j + 1);
							if (j > 0) {
								index--;
							} else {
								index = 0;
							}
						}
						if (j > 0) {
							let article:string = title.substr(i + 1, j - i - 1);
							title = title.substr(0, i) + title.substr(j);
							let price:number = Math.max(
								getNumber(sheet, rowIndex, 'I'),
								getNumber(sheet, rowIndex, 'K')
							);
							let sizeList:Array<string> = sizeHash[getString(sheet, rowIndex, 'E')];

							for (let colIndex:number = sizeColList.length - 1; colIndex >= 0; colIndex--) {
								let count:number = getNumber(sheet, rowIndex, sizeColList[colIndex]);
								if (count > 0) {
									out.push([
										article,
										price,
										sizeList && colIndex > 0 && colIndex - 1 < sizeList.length ? sizeList[colIndex - 1] : '',
										count,
										title
									]);
								}
							}
						}
					}
				}
			}
		}
		return null;
	}

	export function ersi(sheet, range, out:Array<Array<any>>):string {
		// if (range.e.c < 18) {
		// 	return 'ersi требутся не менее 19 столбцов';
		// }
		const rows:number = range.e.r + 1;
		for (let rowIndex:number = 1; rowIndex <= rows; rowIndex++) {
			let brand:string = getString(sheet, rowIndex, ersiColumns[0]);
			if (brand !== 'СНАРЯЖЕНИЕ') {
				let str:string = getString(sheet, rowIndex, ersiColumns[1]);
				let i:number = str.indexOf('   ');
				if (i > 0) {
					let title:string = trim(str.substr(i + 1));
					if (brand.length > 0 && brand !== title.substr(0, brand.length)) {
						title = brand + ' ' + title;
					}

					let count:number = 0;
					for (let k:number = 5; k <= 7; k++) {
						if (ersiColumns[k]) {
							count += getNumber(sheet, rowIndex, ersiColumns[k]);
						}
					}

					out.push([
						str.substr(0, i), //артикл
						Math.max(getNumber(sheet, rowIndex, ersiColumns[3]), getNumber(sheet, rowIndex, ersiColumns[4])), //цена
						getString(sheet, rowIndex, ersiColumns[2]),
						count,
						title
					]);
				}
			}
		}
		return null;
	}

} //end namespace Parser

function showError(value:string):void {
	document.getElementById('log').innerHTML = value;
}

function showStatus(id:string, flag:boolean):void {
	document.getElementById(id + '_status').innerText = flag ? 'загружено' : 'ждем...';
}

const data:Array<any> = [
	'ersi', null, Parser.ersi,
	'megasport', null, Parser.megasport
];

const f = function(event):void {
	const input = event.target as HTMLInputElement;
	const file:File = input.files[0];
	let len:number = file.name.length;
	if (file.name.substr(len - 4) !== '.xls' && file.name.substr(len - 5) !== '.xlsx') {
		showError('поддерживаются тольк xls/xlsx-файлы');
		return;
	}
	const reader:FileReader = new FileReader();
	showStatus(input.id, false);
	reader.onload = function() {
		data[data.indexOf(input.id) + 1] = new Uint8Array(reader.result as ArrayBuffer);
		showStatus(input.id, true);
		(document.getElementById(input.id + '_cb') as HTMLInputElement).checked = true;
	};
	reader.readAsArrayBuffer(file);
};
for (let i:number = 0; i < data.length; i += 3) {
	(document.getElementById(data[i]) as HTMLInputElement).addEventListener('change', f, false);
}

(document.getElementById('bt') as HTMLInputElement).addEventListener('click', function():void {
	const XLSX = window['XLSX'];
	const out:Array<Array<any>> = [
		['Артикул', 'Цена', 'Размер', 'Остаток', 'Название товара']
	];
	let workbook;
	for (let i:number = 1; i < data.length; i += 3) {
		if ((document.getElementById(data[i - 1] + '_cb') as HTMLInputElement).checked && data[i]) {
			workbook = XLSX.read(data[i], { type:'array' });
			if (workbook && workbook.SheetNames.length > 0) {
				let sheet = workbook.Sheets[workbook.SheetNames[0]];
				let error:string = data[i + 1](sheet, XLSX.utils.decode_range(sheet['!ref']), out);
				if (error) {
					showError(error);
					return;
				}
			}
		}
	}
	//сохраняем xlsx
	workbook = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(out), 'FiveSport');
	XLSX.writeFile(workbook, 'five_sport.xlsx');
});