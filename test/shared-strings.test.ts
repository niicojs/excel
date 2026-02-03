import { describe, it, expect } from 'vitest';
import { SharedStrings } from '../src/shared-strings';

describe('SharedStrings', () => {
  it('preserves rich text nodes and counts', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="2">
  <si><t>Plain</t></si>
  <si><r><t xml:space="preserve"> Rich</t></r><r><t>Text</t></r></si>
</sst>`;

    const ss = SharedStrings.parse(xml);
    const out = ss.toXml();

    expect(out).toContain('count="3"');
    expect(out).toContain('uniqueCount="2"');
    expect(out).toContain('<r>');
    expect(out).toContain('Rich');
    expect(out).toContain('Text');
  });

  it('updates total count for duplicate strings', () => {
    const ss = new SharedStrings();
    ss.addString('Alpha');
    ss.addString('Alpha');

    const out = ss.toXml();
    expect(out).toContain('count="2"');
    expect(out).toContain('uniqueCount="1"');
  });
});
