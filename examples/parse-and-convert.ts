import { Workbook } from '../src';

console.log('Load data...');
const input = await Workbook.fromFile('examples/output/CA par Compte et BU 2025.xlsx');

const data = input.sheet('data').toJson();

const x_sell = new Map<
  string,
  {
    top: string;
    total: number;
    bdt: number;
    data: number;
    bus: number;
    regions: number;
    coe: number;
    conseil: number;
    coexya: number;
    emea: number;
    nam: number;
  }
>();

for (const row of data) {
  const top = row.top as string;
  let line = x_sell.get(top);
  if (!line) {
    line = { top, total: 0, bdt: 0, data: 0, bus: 0, conseil: 0, regions: 0, coexya: 0, coe: 0, emea: 0, nam: 0 };
    x_sell.set(top, line);
  }

  const [bu, pole] = [row.bu as string, row.pole as string];
  const ca = row.ca as number;
  line.total += ca;

  if (bu === 'Cloud and Applications services') line.bus += ca;
  else if (bu === 'Data & Technologies (FR)' && pole === 'DTF_Tech for Data') line.data += ca;
  else if (bu === 'Data & Technologies (FR)' && pole === 'DTF_Tech for Data') line.data += ca;
  else if (bu === 'Data & Technologies (FR)' && pole !== 'DTF_Tech for Data') line.bdt += ca;
  else if (bu === 'Supply Chain') line.conseil += ca;
  else if (bu === 'Consulting') line.conseil += ca;
  else if (bu === 'TenSquare') line.conseil += ca;
  else if (pole === 'M--DATA&AI') line.data += ca;
  else if (pole === 'M---') line.data += ca;
  else if (pole === 'M--DIGITAL') line.bdt += ca;
  else if (pole === 'M--Open Innovation') line.data += ca;
  else if (pole === 'M--Régions') line.regions += ca;
  else if (bu === 'Régions') line.regions += ca;
  else if (bu === 'Coexya') line.coexya += ca;
  else if (bu === 'SAP (NAM)') line.nam += ca;
  else if (bu === 'Microsoft (NAM)') line.nam += ca;
  else if (bu === 'Oracle Technologies') line.nam += ca;
  else if (bu === 'Corporate North America') line.nam += ca;
  else if (bu === 'Data & Technologies (NAM)') line.nam += ca;
  else if (bu === 'EAM') line.nam += ca;
  else if (bu === 'Consulting (NAM)') line.nam += ca;
  else if (bu === 'Talan Tech (NAM)') line.nam += ca;
  else if (bu === 'Talan USA') line.nam += ca;
  else if (bu === 'Switzerland') line.emea += ca;
  else if (bu === 'DATAROOTS') line.emea += ca;
  else if (bu === 'Micropole Switzerland') line.emea += ca;
  else if (bu === 'PaP Switzerland') line.emea += ca;
  else if (bu === 'Talan UK') line.emea += ca;
  else if (bu === '9FT') line.emea += ca;
  else if (bu === 'GEMSERV') line.emea += ca;
  else if (bu === 'BDP') line.emea += ca;
  else if (bu === 'Luxembourg') line.emea += ca;
  else if (bu === 'Spain') line.emea += ca;
  else if (bu === 'Belgium') line.emea += ca;
  else if (bu === 'Poland') line.emea += ca;
  else if (bu === 'Singapore') line.emea += ca;
  else if (bu.startsWith('COE')) line.coe += ca;
  else console.log(`  ${bu}/${pole}: ${ca}`);
}

const output = Workbook.create();

output.addSheetFromData({
  name: 'x-sell',
  data: Array.from(x_sell.values()),
  columns: [
    { key: 'top', header: 'Top' },
    { key: 'total', header: 'Total', style: { numberFormat: '#,##0' } },
    { key: 'bus', header: 'BUS', style: { numberFormat: '#,##0' } },
    { key: 'data', header: 'Data', style: { numberFormat: '#,##0' } },
    { key: 'bdt', header: 'BDT', style: { numberFormat: '#,##0' } },
    { key: 'conseil', header: 'Conseil', style: { numberFormat: '#,##0' } },
    { key: 'regions', header: 'Régions', style: { numberFormat: '#,##0' } },
    { key: 'coexya', header: 'Coexya', style: { numberFormat: '#,##0' } },
    { key: 'coe', header: 'COE', style: { numberFormat: '#,##0' } },
    { key: 'emea', header: 'EMEA', style: { numberFormat: '#,##0' } },
    { key: 'nam', header: 'NAM', style: { numberFormat: '#,##0' } },
  ],
});

output.addSheetFromData({
  name: 'x-sell percent',
  data: Array.from(x_sell.values()).map((line: any) => {
    return {
      top: line.top,
      total: line.total,
      bus: line.bus / line.total,
      data: line.data / line.total,
      bdt: line.bdt / line.total,
      conseil: line.conseil / line.total,
      regions: line.regions / line.total,
      coexya: line.coexya / line.total,
      coe: line.coe / line.total,
      emea: line.emea / line.total,
      nam: line.nam / line.total,
    };
  }),
  columns: [
    { key: 'top', header: 'Top' },
    { key: 'total', header: 'Total', style: { numberFormat: '#,##0' } },
    { key: 'bus', header: 'BUS', style: { numberFormat: '0%' } },
    { key: 'data', header: 'Data', style: { numberFormat: '0%' } },
    { key: 'bdt', header: 'BDT', style: { numberFormat: '0%' } },
    { key: 'conseil', header: 'Conseil', style: { numberFormat: '0%' } },
    { key: 'regions', header: 'Régions', style: { numberFormat: '0%' } },
    { key: 'coexya', header: 'Coexya', style: { numberFormat: '0%' } },
    { key: 'coe', header: 'COE', style: { numberFormat: '0%' } },
    { key: 'emea', header: 'EMEA', style: { numberFormat: '0%' } },
    { key: 'nam', header: 'NAM', style: { numberFormat: '0%' } },
  ],
});

output.addSheetFromData({ name: 'details', data });

await output.toFile('examples/output/x-sell.xlsx');
console.log('Done.');
