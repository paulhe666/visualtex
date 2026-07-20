import { readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";

const root = resolve(
  process.argv[2] ?? "src-windows/artifacts/real-formula-fixtures",
);
const manifest = JSON.parse(
  await readFile(resolve(root, "manifest.json"), "utf8"),
) as Array<{
  id: string;
  fileName: string;
  width: number;
  height: number;
}>;

for (const item of manifest) {
  const svg = await readFile(resolve(root, item.fileName), "utf8");
  const encoded = Buffer.from(svg, "utf8").toString("base64");
  const canvasWidth = Math.max(1, Math.ceil(item.width * 2));
  const canvasHeight = Math.max(1, Math.ceil(item.height * 2));
  const html = `<!doctype html>
<html><head><meta charset="utf-8"><style>
html,body{margin:0;padding:0;width:${canvasWidth}px;height:${canvasHeight}px;overflow:hidden;background:transparent}
canvas{display:block;width:${canvasWidth}px;height:${canvasHeight}px}
</style></head><body><canvas id="c" width="${canvasWidth}" height="${canvasHeight}"></canvas><script>
const canvas=document.getElementById('c');
const context=canvas.getContext('2d');
const image=new Image();
image.onload=()=>{context.clearRect(0,0,canvas.width,canvas.height);context.drawImage(image,0,0,canvas.width,canvas.height);document.title='ready';};
image.src='data:image/svg+xml;base64,${encoded}';
</script></body></html>`;
  await writeFile(resolve(root, `${item.id}-canvas.html`), html, "utf8");
  console.log(`${item.id}\t${canvasWidth}x${canvasHeight}`);
}
