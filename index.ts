import XLSX, { utils } from 'xlsx';
import { existsSync, writeFile } from 'fs';
import json2csv, { Parser, Transform } from 'json2csv';
import moment from 'moment';
moment.locale("id");

interface Line {
  Nama: string
  Terduga: string
  WN: string
  Alamat: string
  'Tpt Lahir': string
  'Tgl Lahir': string
  'Kode Densus': string
}

interface LineOutput {
  nama: string
  terduga: string | null
  warga_negara: string | null
  alamat: string | null
  tempat_lahir: string | null
  tanggal_lahir: string | null
  kode_densus: string | null
  ref_line: number
}

function cleanRow(line: Line, line_number: number): LineOutput[] {
  const list_nama = (line.Nama ?? '').replace(/\;+/g, '')
    .replace(/Alias/g, 'alias')
    .replace(/ALIAS/g, 'alias')
    .split(/\s+alias\s+/);
  const terduga = (line.Terduga ?? '').replace(/\;+/g, '') ?? null;
  const warga_negara = (line.WN ?? '').replace(/\;+/g, '') ?? null;
  const alamat = (line.Alamat ?? '').replace(/\;+/g, '') ?? null;
  const tempat_lahir = (line['Tpt Lahir'] ?? '').replace(/\;+/g, '') ?? null;
  const tanggal_lahir = String((line['Tgl Lahir'] ?? '')).replace(/\;+/g, '') ?? null;
  const kode_densus = (line['Kode Densus'] ?? '').replace(/\;+/g, '') ?? null;
  const ref_line = line_number + 1;

  return list_nama.map((nama: string) => {
    let list_tanggal_lahir = [tanggal_lahir];
    if (tanggal_lahir && tanggal_lahir.includes('atau')) {
      list_tanggal_lahir = tanggal_lahir.split(/\s+atau\s+/);
    }

    for (let i = 0; i < list_tanggal_lahir.length; i++) {
      if (moment(list_tanggal_lahir[i], 'D MMMM YYYY', true).isValid()) {
        list_tanggal_lahir[i] = moment(list_tanggal_lahir[i], 'D MMMM YYYY').toISOString();
        continue;
      }

      if (moment(list_tanggal_lahir[i], 'DD-MM-YYYY', true).isValid()) {
        list_tanggal_lahir[i] = moment(list_tanggal_lahir[i], 'DD-MM-YYYY').toISOString();
        continue;
      }

      list_tanggal_lahir[i] = null;
    }

    return list_tanggal_lahir.map<LineOutput>((tgl_lahir_iso: string) => ({
      nama,
      terduga,
      warga_negara,
      alamat,
      tempat_lahir,
      tanggal_lahir: tgl_lahir_iso,
      kode_densus,
      ref_line
    }));
  }).reduce((a: LineOutput[], c: LineOutput[]) => [...a, ...c], []);
}

function exportToCSV(filename: string, data: LineOutput[]) {
  const json2csvParser = new Parser();
  const csv = json2csvParser.parse(data); 
  writeFile(`${filename}.csv`, csv, (err) => {
    if (err)
      console.log(err);
    else {
      console.log(`Export to CSV ${filename}.csv sucessfully\n`);
    }
  });
}

function exportToExcel(filename: string, data: LineOutput[]) {
  const ws = utils.json_to_sheet(data);
  const wb = utils.book_new();
  utils.book_append_sheet(wb, ws, "Data");
  XLSX.writeFile(wb, `${process.argv[3]}.xlsx`);
  console.log(`Export to Excel ${filename}.xlsx sucessfully\n`);
}

function main() {
  if (process.argv.length < 3 || !process.argv[2]) {
    throw new Error(`file input tidak boleh kosong`);
  }

  if (!existsSync(process.argv[2])) {
    throw new Error(`file input "${process.argv[2]}" tidak ditemukan`);
  }

  if (process.argv.length < 4 || !process.argv[3]) {
    throw new Error(`nama file output tidak boleh kosong`);
  }

  const workbook = XLSX.readFile(process.argv[2]);
  const list_sheet = workbook.SheetNames;
  const excel_data: Line[] = XLSX.utils.sheet_to_json(workbook.Sheets[list_sheet[0]]);

  const list_output: LineOutput[] = excel_data.map(cleanRow).reduce((a: LineOutput[], c: LineOutput[]) => [...a, ...c], []);

  exportToExcel(process.argv[3], list_output);
  exportToCSV(process.argv[3], list_output);
}

try {
  main();
} catch (err: any) {
  console.error(err.toString());
}
