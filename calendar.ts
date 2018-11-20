interface Cell {
	v:string | number;
	t:string;
	w:string;
}

const newLineRE:RegExp = new RegExp('\r?\n', 'g');
function getString(sheet, rowIndex:number, col:string):string {
	const cell:Cell = sheet[col + rowIndex];
	if (cell && cell.v) {
		return String(cell.v).trim().replace(newLineRE, '<br>');
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

const f = function(event):void {
	const input = event.target as HTMLInputElement;
	const file:File = input.files[0];
	let len:number = file.name.length;
	if (file.name.substr(len - 4) !== '.xls' && file.name.substr(len - 5) !== '.xlsx') {
		console.error('поддерживаются тольк xls/xlsx-файлы');
		return;
	}
	const reader:FileReader = new FileReader();

	reader.onload = function() {
		if (!(reader.result instanceof ArrayBuffer)) {
			return;
		}
		
		const XLSX = window['XLSX'];
		let workbook = XLSX.read(new Uint8Array(reader.result as ArrayBuffer), { type:'array' });
		if (workbook && workbook.SheetNames.length > 0) {
			let str:string = '';
			let sheet = workbook.Sheets[workbook.SheetNames[0]];
			const rows:number = XLSX.utils.decode_range(sheet['!ref']).e.r + 1;
			for (let rowIndex:number = 7; rowIndex <= rows; rowIndex++) {
				let date:string = getString(sheet, rowIndex, 'E');
				let i:number = date.indexOf('-');
				if (i > 0) {
					let date2:string = date.substr(i + 1);
					i = date2.indexOf('.');
					let vD:number = parseInt(date2.substr(0, i));
					let vM:number = parseInt(date2.substr(i + 1));
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
(document.getElementById('file') as HTMLInputElement).addEventListener('change', f, false);
