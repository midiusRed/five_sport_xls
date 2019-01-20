const eRSi = {};
eRSi['одежда'] = { H:'XS', I:'S', J:'M', K:'L', L:'XL', M:'XXL', N:'XXXL' };
eRSi['обувь'] = { G:'4', H:'4.5', I:'5', J:'5.5', K:'6', L:'6.5', M:'7', N:'7.5', O:'8', P:'8.5', Q:'9', R:'9.5', S:'10', T:'10.5', U:'11',	V:'11.5', W:'12', X:'12.5', Y:'13', Z:'13.5', AA:'14', AB:'15' };

const sizeTransform = {
	'4.5':'"4,5"',
	'5.5':'"5,5"',
	'6.5':'"6,5"',
	'7.5':'"7,5"',
	'8.5':'"8,5"',
	'9.5':'"9,5"',
	'10.5':'"10,5"',
	'11.5':'"11,5"',
	'12.5':'"12,5"',
	'13.5':'"13,5"'
};

function getSizeTitle(value:string):string {
	return value in sizeTransform ? sizeTransform[value] : value; 
}

function log(value:string):void {
	document.getElementById('log').innerHTML = value;
}

function showStatus(id:string, flag:boolean):void {
	document.getElementById(id + '_status').innerText = flag ? ' загружено' : ' ждем...';
}

let products:Array<Array<string>>;

(document.getElementById('products') as HTMLInputElement).addEventListener('change', event => {
	const input = event.target as HTMLInputElement;
	const file:File = input.files[0];
	if (file.name.substr(file.name.length - 4) !== '.csv') {
		log('для товаров используется csv с разделителем ;');
		return;
	}
	const reader:FileReader = new FileReader();
	showStatus(input.id, false);
	reader.onload = () => {
		products = CSVToArray(reader.result as string, ';');
		showStatus(input.id, true);
	};
	reader.readAsText(file);
}, false);

let ignoreList:Array<string>;

(document.getElementById('ignore') as HTMLInputElement).addEventListener('change', event => {
	const input = event.target as HTMLInputElement;
	const reader:FileReader = new FileReader();
	showStatus(input.id, false);
	reader.onload = () => {
		ignoreList = (reader.result as string).split(',').map(art => art.trim());
		showStatus(input.id, true);
	};
	reader.readAsText(input.files[0]);
}, false);

const data:Array<any> = [
	'ersi', null
];

const f = function(event):void {
	const input = event.target as HTMLInputElement;
	const file:File = input.files[0];
	let len:number = file.name.length;
	if (file.name.substr(len - 4) !== '.xls' && file.name.substr(len - 5) !== '.xlsx') {
		log('поддерживаются только xls/xlsx-файлы');
		return;
	}
	const reader:FileReader = new FileReader();
	showStatus(input.id, false);
	reader.onload = () => {
		data[data.indexOf(input.id) + 1] = new Uint8Array(reader.result as ArrayBuffer);
		showStatus(input.id, true);
		(document.getElementById(input.id + '_cb') as HTMLInputElement).checked = true;
	};
	reader.readAsArrayBuffer(file);
};
for (let i:number = 0; i < data.length; i += 2) {
	(document.getElementById(data[i]) as HTMLInputElement).addEventListener('change', f, false);
}

(document.getElementById('bt') as HTMLInputElement).addEventListener('click', function():void {
	if (!products) {
		log('нужно загрузить товары');
		return;
	}
	const urlRow:number = products.length > 0 ? products[0].indexOf('Ссылка на витрину') : -1;
	if (urlRow < 0) {
		log('в товарах не удается найти столбец "Ссылка на витрину"');
		return;
	}
	const artRow:number = products.length > 0 ? products[0].indexOf('Код артикула') : -1;
	if (artRow < 0) {
		log('в товарах не удается найти столбец "Код артикула"');
		return;
	}

	let limit = (document.getElementById('limit') as HTMLInputElement).valueAsNumber;
	if (isNaN(limit) || limit <= 0) {
		limit = Number.MAX_VALUE;
	} 

	let count:number = 0;
	let out:string = '"Ссылка на витрину";"Наименование артикула";"Размеры";"В наличии";"Закупочная цена";"Цена";"Код артикула"';
	const idFilter = new RegExp('-', 'g');
	
	const XLSX = window['XLSX'];
	let workbook;
	loop: for (let i:number = 0; i < data.length; i += 2) {
		if ((document.getElementById(data[i] + '_cb') as HTMLInputElement).checked && data[i]) {
			workbook = XLSX.read(data[i + 1], { type:'array' });
			if (workbook && workbook.SheetNames.length > 0) {
				for (let sheetIndex:number = 0; sheetIndex < workbook.SheetNames.length; sheetIndex++) {
					let sheetName:string = workbook.SheetNames[sheetIndex];
					let cols:{ [col:string]:string } = eRSi[sheetName];
					if (!cols) {
						continue;
					}
					let sizes:string = '';
					for (let col in cols) {
						if (sizes.length > 0) {
							sizes += ',';
						}
						sizes += getSizeTitle(cols[col]);
					}
					
					let sheet = workbook.Sheets[sheetName];
					let range = XLSX.utils.decode_range(sheet['!ref']);
					const rows:number = range.e.r + 1;
					for (let rowIndex:number = rows; rowIndex > 0; rowIndex--) {
						let id:string = getString(sheet, rowIndex, 'A');
						if (id && id.length > 0 && (!ignoreList || ignoreList.indexOf(id) < 0)) {
							for (let product of products) {
								if (!product[artRow]) {
									continue;
								}
								let curId = product[artRow].trim();
								if (curId === id || curId.replace(idFilter, ' ') === id) {
									count++;
									let sellPrice:number = getNumber(sheet, rowIndex, 'C');
									let buyPrice:number = getNumber(sheet, rowIndex, 'E');
									out += '\n"' + product[urlRow] + '";;<{' + sizes + '}>;;' + buyPrice + ';' + sellPrice + ';"' + curId + '"';
									for (let col in cols) {
										out += '\n"' + product[urlRow] + '";' + getSizeTitle(cols[col]) + ';' + getSizeTitle(cols[col]) + ';' + getNumber(sheet, rowIndex, col) +
											';' + buyPrice + ';' + sellPrice + ';"' + curId + '"';
									}
									if (count >= limit) {
										break loop;
									} else {
										break;
									}
								}
							}
						}
					} 
				}
			}
		}
	}
	log('товаров затронуто: ' + count);
	window['saveAs'](new Blob([out], { type:'text/plain;charset=utf-8' }), 'wt_products.csv');
});

interface Cell {
	v:string | number;
	t:string;
	w:string;
}

function getString(sheet, rowIndex:number, col:string):string {
	const cell:Cell = sheet[col + rowIndex];
	if (cell && cell.v) {
		return String(cell.v).trim();
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

function CSVToArray(strData:string, strDelimiter:string):Array<Array<string>> {
	// Check to see if the delimiter is defined. If not,
	// then default to comma.
	strDelimiter = (strDelimiter || ",");

	// Create a regular expression to parse the CSV values.
	let objPattern = new RegExp(
		(
			// Delimiters.
			"(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +

			// Quoted fields.
			"(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +

			// Standard fields.
			"([^\"\\" + strDelimiter + "\\r\\n]*))"
		),
		"gi"
	);


	// Create an array to hold our data. Give the array
	// a default empty first row.
	let arrData = [[]];

	// Create an array to hold our individual pattern
	// matching groups.
	let arrMatches;


	// Keep looping over the regular expression matches
	// until we can no longer find a match.
	while (Boolean(arrMatches = objPattern.exec( strData ))){

		// Get the delimiter that was found.
		let strMatchedDelimiter = arrMatches[ 1 ];

		// Check to see if the given delimiter has a length
		// (is not the start of string) and if it matches
		// field delimiter. If id does not, then we know
		// that this delimiter is a row delimiter.
		if (
			strMatchedDelimiter.length &&
			strMatchedDelimiter !== strDelimiter
		){

			// Since we have reached a new row of data,
			// add an empty row to our data array.
			arrData.push( [] );

		}

		let strMatchedValue;

		// Now that we have our delimiter out of the way,
		// let's check to see which kind of value we
		// captured (quoted or unquoted).
		if (arrMatches[ 2 ]){

			// We found a quoted value. When we capture
			// this value, unescape any double quotes.
			strMatchedValue = arrMatches[ 2 ].replace(
				new RegExp( "\"\"", "g" ),
				"\""
			);

		} else {

			// We found a non-quoted value.
			strMatchedValue = arrMatches[ 3 ];

		}


		// Now that we have our value string, let's add
		// it to the data array.
		arrData[ arrData.length - 1 ].push( strMatchedValue );
	}

	// Return the parsed data.
	return arrData;
}