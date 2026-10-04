import {latexToMathMl} from '../src/export/runtime.ts';
const cases=[String.raw`\mathrm{\hslash}`,String.raw`\mathrm{ℏ}`,String.raw`\mathbfit{ℏ}`,String.raw`\mathbfit{\hslash}`,String.raw`\mathrm{\Bbbk}`,String.raw`\mathit{\sin x}`,String.raw`\mathit{\operatorname{x}}`,String.raw`\mathrm{\alpha}`,String.raw`\mathrm{α}`];
for(const latex of cases) console.log(JSON.stringify({latex,mathMl:await latexToMathMl(latex,false)}));
