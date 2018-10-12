import {Parser} from 'src/Parser';

function log(value:string):void {
	document.getElementById('log').innerHTML = value;
}

function status(id:string, flag:boolean):void {
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
		log('поддерживаются тольк xls/xlsx-файлы');
		return;
	}
	const reader:FileReader = new FileReader();
	status(input.id, false);
	reader.onload = function() {
		data[data.indexOf(input.id) + 1] = new Uint8Array(reader.result as ArrayBuffer);
		status(input.id, true);
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
					log(error);
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
