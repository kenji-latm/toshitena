// アドイン用アイコンPNG（点線円）を依存ライブラリなしで生成する。
// 使い方: node app/oshitena/tools/make-icons.mjs
import { deflateSync } from "node:zlib";
import { writeFileSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const OUT_DIR = join(dirname(fileURLToPath(import.meta.url)), "..", "assets");
const COLOR = [217, 58, 34]; // 朱 #D93A22

function crc32(buf) {
  let table = crc32.table;
  if (!table) {
    table = crc32.table = new Int32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c;
    }
  }
  let crc = -1;
  for (let i = 0; i < buf.length; i++) crc = (crc >>> 8) ^ table[(crc ^ buf[i]) & 0xff];
  return (crc ^ -1) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([len, body, crc]);
}

function encodePng(size, rgba) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // RGBA
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0; // filter: none
    rgba.copy(raw, y * (size * 4 + 1) + 1, y * size * 4, (y + 1) * size * 4);
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(raw, { level: 9 })),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

// スーパーサンプリングで破線円を描く
function drawDashedCircle(size) {
  const rgba = Buffer.alloc(size * size * 4);
  const ss = 4;
  const center = size / 2;
  const radius = size * 0.4;
  const stroke = Math.max(1.1, size * 0.085);
  const dashes = 8; // 破線の本数
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      let hits = 0;
      for (let sy = 0; sy < ss; sy++) {
        for (let sx = 0; sx < ss; sx++) {
          const px = x + (sx + 0.5) / ss - center;
          const py = y + (sy + 0.5) / ss - center;
          const dist = Math.hypot(px, py);
          if (Math.abs(dist - radius) > stroke / 2) continue;
          const angle = Math.atan2(py, px) + Math.PI; // 0..2π
          const seg = Math.floor((angle / (2 * Math.PI)) * dashes * 2) % 2;
          if (size <= 16 || seg === 0) hits++; // 16pxは実線の方が視認しやすい
        }
      }
      const alpha = Math.round((hits / (ss * ss)) * 255);
      if (alpha > 0) {
        const i = (y * size + x) * 4;
        rgba[i] = COLOR[0];
        rgba[i + 1] = COLOR[1];
        rgba[i + 2] = COLOR[2];
        rgba[i + 3] = alpha;
      }
    }
  }
  return rgba;
}

mkdirSync(OUT_DIR, { recursive: true });
for (const size of [16, 32, 64, 80]) {
  const png = encodePng(size, drawDashedCircle(size));
  const file = join(OUT_DIR, `icon-${size}.png`);
  writeFileSync(file, png);
  console.log(`wrote ${file} (${png.length} bytes)`);
}
