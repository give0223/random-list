import {
  saveAs
} from 'file-saver'
import XLSX from 'xlsx'


function Workbook() {
  if (!(this instanceof Workbook)) return new Workbook();
  this.SheetNames = [];
  this.Sheets = {};
}

export function appendSheet(sheet, name = `sheet${ this.SheetNames.length + 1 }`) {
  this.SheetNames = [...this.SheetNames, name];
  this.Sheets[name] = sheet;
}