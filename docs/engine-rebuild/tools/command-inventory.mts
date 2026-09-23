// Enumerates real built-in commands. Does not read/write the user's symbol library.
import { createHash } from 'node:crypto';
import { commandRegistry } from '../../../apps/windows/src/autocomplete/commandRegistry.ts';
const rows = commandRegistry.map(c => ({ id: c.id, command: c.command, template: c.insertTemplate,
  preview: c.previewLatex, category: c.category, supportedByLegacy: c.supportedInMathMode,
  newCore: { parse: 'NOT_IMPLEMENTED', edit: 'NOT_IMPLEMENTED', layout: 'NOT_IMPLEMENTED', latex: 'NOT_IMPLEMENTED', svg: 'NOT_IMPLEMENTED', mathMl: 'NOT_IMPLEMENTED', office: 'NOT_RUN' } }));
const summary = { schema: 'visualtex-command-inventory-v1', count: rows.length,
  uniqueIds: new Set(rows.map(r => r.id)).size,
  dataSha256: createHash('sha256').update(JSON.stringify(rows)).digest('hex'),
  categories: Object.fromEntries([...new Set(rows.map(r => r.category))].sort().map(c => [c, rows.filter(r => r.category === c).length])),
  userCustomSymbols: 'NOT_INVENTORIED; preserve existing library, P09 gate',
};
console.log(JSON.stringify(process.argv.includes('--all') ? { ...summary, rows } : summary, null, 2));
