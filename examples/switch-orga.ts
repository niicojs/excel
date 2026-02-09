import { toAddress, Workbook } from '../src';

const files = ['./examples/output/CA par Compte et BU 2025.xlsx'];

async function convert_orga(path: string) {
  console.info('Converting organisation in file:', path);

  const wb = await Workbook.fromFile(path);
  const data = wb.sheet('data').toJson();

  for (const line of data) {
    if (line.bu === 'Data & Technologies (FR)') {
      if (line.pole === 'DTF_Tech for Data') line.bu = 'Data & AI';
      else line.bu = 'BDT';
    } else if (line.bu === 'Micropole France') {
      if (line.pole === 'M--Régions') {
        line.bu = 'Régions';
        if (line.domain === 'M-Centre-Est') line.pole = 'REG_Sud-Est';
        if (line.domain === 'M-Digital Centre-Est - IDEA') line.pole = 'REG_Sud-Est';
        if (line.domain === 'M-Grand-Ouest') line.pole = 'REG_Ouest';
      } else {
        if (line.pole === 'M--DATA&AI') line.bu = 'Data & AI';
        if (line.pole === 'M--Open Innovation') line.bu = 'Data & AI';
        if (line.pole === 'M--DIGITAL') line.bu = 'BDT';
        if (line.pole === 'M---') line.bu = 'Data & AI';
      }
    }
  }

  console.info('Saving...');

  const new_wb = Workbook.create();
  new_wb.addSheetFromData({ name: 'data', data });

  // const fields = Math.max(...data.map((row) => Object.keys(row).length));

  const [w, h] = [Object.keys(data[0]).length - 1, data.length];
  new_wb.addSheet('by-top');
  new_wb
    .createPivotTable({
      name: 'by top',
      source: `data!A1:${toAddress(h, w)}`,
      target: 'by-top!A1',
    })
    .addRowField('top')
    .addValueField({
      field: 'ca',
      aggregation: 'sum',
      numberFormat: '#,##0 €',
      name: 'Total CA',
    })
    .sortField('top', 'desc');

  await new_wb.toFile(path.replace('.xlsx', '_converted.xlsx'));
}

for (const file of files) {
  await convert_orga(file);
}
