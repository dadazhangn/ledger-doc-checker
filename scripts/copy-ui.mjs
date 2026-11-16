import { cp, mkdir } from 'node:fs/promises';
await mkdir(new URL('../dist/public/', import.meta.url), { recursive: true });
for (const name of ['index.html', 'style.css', 'app.js', 'word-mode.js']) {
  await cp(new URL(`../public/${name}`, import.meta.url), new URL(`../dist/public/${name}`, import.meta.url));
}
