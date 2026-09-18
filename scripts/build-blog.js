// articles/*.md から app/blog/ 配下の静的HTMLを生成する（依存パッケージなし）
// 使い方: npm run build:blog
"use strict";

const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const SRC_DIR = path.join(ROOT, "articles");
const OUT_DIR = path.join(ROOT, "app", "blog");
const SITE_NAME = "トウシテナ";
const AUTHOR = "司法書士 石本憲史";

function escapeHtml(s) {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// front matter（--- で囲んだ key: value）を読む
function parseFrontMatter(text) {
  const m = text.match(/^---\r?\n([\s\S]*?)\r?\n---\r?\n?([\s\S]*)$/);
  if (!m) return { meta: {}, body: text };
  const meta = {};
  for (const line of m[1].split(/\r?\n/)) {
    const idx = line.indexOf(":");
    if (idx < 0) continue;
    const key = line.slice(0, idx).trim();
    let value = line.slice(idx + 1).trim();
    if (key === "tags") {
      value = value
        .replace(/^\[|\]$/g, "")
        .split(",")
        .map((t) => t.trim())
        .filter(Boolean);
    }
    meta[key] = value;
  }
  return { meta, body: m[2] };
}

// インライン記法: **強調**、`コード`、[文字](URL)
function inline(s) {
  let out = escapeHtml(s);
  out = out.replace(/`([^`]+)`/g, "<code>$1</code>");
  out = out.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
  out = out.replace(/\[([^\]]+)\]\((https?:\/\/[^)\s]+|[^)\s]+\.html)\)/g, (_, t, u) => {
    const external = /^https?:/.test(u);
    return `<a href="${u}"${external ? ' target="_blank" rel="noreferrer"' : ""}>${t}</a>`;
  });
  return out;
}

// ごく小さなMarkdown変換（見出し・段落・箇条書き・番号付き・表・引用・注意ボックス）
function markdownToHtml(md) {
  const lines = md.replace(/\r\n/g, "\n").split("\n");
  const html = [];
  const toc = [];
  let i = 0;

  const flushParagraph = (buf) => {
    if (buf.length) html.push(`<p>${inline(buf.join(" "))}</p>`);
    buf.length = 0;
  };

  let para = [];
  while (i < lines.length) {
    const line = lines[i];

    if (/^\s*$/.test(line)) {
      flushParagraph(para);
      i++;
      continue;
    }

    const h = line.match(/^(#{2,4})\s+(.+)$/);
    if (h) {
      flushParagraph(para);
      const level = h[1].length;
      const text = h[2].trim();
      const id = `s${toc.length + 1}`;
      if (level === 2) toc.push({ id, text });
      html.push(`<h${level}${level === 2 ? ` id="${id}"` : ""}>${inline(text)}</h${level}>`);
      i++;
      continue;
    }

    if (/^:::\s*(note|warn)\s*$/.test(line)) {
      flushParagraph(para);
      const kind = line.match(/^:::\s*(note|warn)/)[1];
      const buf = [];
      i++;
      while (i < lines.length && !/^:::\s*$/.test(lines[i])) {
        buf.push(lines[i]);
        i++;
      }
      i++;
      html.push(`<div class="box box-${kind}">${markdownToHtml(buf.join("\n")).html}</div>`);
      continue;
    }

    if (/^>\s?/.test(line)) {
      flushParagraph(para);
      const buf = [];
      while (i < lines.length && /^>\s?/.test(lines[i])) {
        buf.push(lines[i].replace(/^>\s?/, ""));
        i++;
      }
      html.push(`<blockquote>${markdownToHtml(buf.join("\n")).html}</blockquote>`);
      continue;
    }

    if (/^\|/.test(line)) {
      flushParagraph(para);
      const rows = [];
      while (i < lines.length && /^\|/.test(lines[i])) {
        rows.push(lines[i]);
        i++;
      }
      const cells = (r) =>
        r
          .replace(/^\|/, "")
          .replace(/\|\s*$/, "")
          .split("|")
          .map((c) => c.trim());
      const header = cells(rows[0]);
      const body = rows.slice(1).filter((r) => !/^\|\s*:?-+/.test(r));
      let t = '<div class="table-wrap"><table><thead><tr>';
      t += header.map((c) => `<th>${inline(c)}</th>`).join("");
      t += "</tr></thead><tbody>";
      for (const r of body) {
        t += "<tr>" + cells(r).map((c) => `<td>${inline(c)}</td>`).join("") + "</tr>";
      }
      t += "</tbody></table></div>";
      html.push(t);
      continue;
    }

    if (/^(-|\*)\s+/.test(line) || /^\d+\.\s+/.test(line)) {
      flushParagraph(para);
      const ordered = /^\d+\.\s+/.test(line);
      const re = ordered ? /^\d+\.\s+/ : /^(-|\*)\s+/;
      const items = [];
      while (i < lines.length && re.test(lines[i])) {
        let item = lines[i].replace(re, "");
        i++;
        // 継続行（先頭が空白）は同じ項目に結合
        while (i < lines.length && /^\s+\S/.test(lines[i]) && !re.test(lines[i])) {
          item += " " + lines[i].trim();
          i++;
        }
        items.push(`<li>${inline(item)}</li>`);
      }
      html.push(`<${ordered ? "ol" : "ul"}>${items.join("")}</${ordered ? "ol" : "ul"}>`);
      continue;
    }

    para.push(line.trim());
    i++;
  }
  flushParagraph(para);
  return { html: html.join("\n"), toc };
}

function layout({ title, description, bodyHtml, isIndex }) {
  const rel = isIndex ? "../" : "../";
  return `<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(title)}｜${SITE_NAME} 実務コラム</title>
  <meta name="description" content="${escapeHtml(description || "")}">
  <link rel="stylesheet" href="blog.css">
</head>
<body>
  <header class="topbar">
    <a class="brand" href="${rel}">
      <span class="brand-mark" aria-hidden="true">IL</span>
      <span>ISHIMOTO Legal</span>
    </a>
    <nav class="topnav">
      <a href="${rel}">トウシテナ（電子署名チェッカー）</a>
      <a href="index.html">実務コラム</a>
    </nav>
  </header>
  <main>
${bodyHtml}
  </main>
  <footer class="site-footer">
    <p>本コラムは商業登記・会社法実務に関する一般的な情報提供を目的としたもので、個別の事案に対する法的助言ではありません。具体的な手続は、最新の法令・先例および管轄法務局の取扱いをご確認ください。</p>
    <p>〒541-0042 大阪市中央区今橋一丁目７番１４号 堺筋北浜宗田ビル８Ｆ　${AUTHOR}</p>
  </footer>
</body>
</html>
`;
}

function renderArticle(slug, meta, body) {
  const { html, toc } = markdownToHtml(body);
  const tags = (meta.tags || []).map((t) => `<span class="tag">${escapeHtml(t)}</span>`).join("");
  const tocHtml =
    toc.length >= 3
      ? `<nav class="toc" aria-label="目次"><div class="toc-title">目次</div><ol>${toc
          .map((t) => `<li><a href="#${t.id}">${inline(t.text)}</a></li>`)
          .join("")}</ol></nav>`
      : "";
  const bodyHtml = `
    <article class="post">
      <p class="crumbs"><a href="index.html">実務コラム</a> / ${escapeHtml(meta.category || "")}</p>
      <h1>${escapeHtml(meta.title)}</h1>
      <p class="post-meta"><time datetime="${escapeHtml(meta.date)}">${escapeHtml(meta.date)}</time>　${AUTHOR}</p>
      <div class="tags">${tags}</div>
      ${meta.description ? `<p class="lead">${inline(meta.description)}</p>` : ""}
      ${tocHtml}
      <div class="post-body">
${html}
      </div>
      <p class="back"><a href="index.html">← 記事一覧へ戻る</a></p>
    </article>`;
  return layout({ title: meta.title, description: meta.description, bodyHtml });
}

function renderIndex(posts) {
  const byCategory = new Map();
  for (const p of posts) {
    const key = p.meta.category || "その他";
    if (!byCategory.has(key)) byCategory.set(key, []);
    byCategory.get(key).push(p);
  }
  let list = "";
  for (const [cat, items] of byCategory) {
    list += `<section class="cat"><h2>${escapeHtml(cat)}</h2><ul class="post-list">`;
    for (const p of items) {
      list += `<li>
        <a class="post-link" href="${p.slug}.html">${escapeHtml(p.meta.title)}</a>
        <p class="post-summary">${inline(p.meta.description || "")}</p>
        <p class="post-meta"><time datetime="${escapeHtml(p.meta.date)}">${escapeHtml(p.meta.date)}</time>
        ${(p.meta.tags || []).map((t) => `<span class="tag">${escapeHtml(t)}</span>`).join("")}</p>
      </li>`;
    }
    list += "</ul></section>";
  }
  const bodyHtml = `
    <div class="title-row">
      <div>
        <h1>実務コラム</h1>
        <p class="lead">商業登記・組織再編・種類株式・役員変更など、会社法人登記の現場で実際に判断に迷いやすい論点を、司法書士の視点で整理しています。</p>
      </div>
      <div class="meta-pill">${posts.length} 記事</div>
    </div>
${list}`;
  return layout({
    title: "記事一覧",
    description: "商業登記・組織再編の実務論点を司法書士が解説するコラム一覧",
    bodyHtml,
    isIndex: true,
  });
}

function main() {
  if (!fs.existsSync(SRC_DIR)) {
    console.error(`原稿フォルダがありません: ${SRC_DIR}`);
    process.exit(1);
  }
  fs.mkdirSync(OUT_DIR, { recursive: true });

  const files = fs.readdirSync(SRC_DIR).filter((f) => f.endsWith(".md")).sort();
  const posts = [];
  for (const f of files) {
    const slug = f.replace(/\.md$/, "");
    const text = fs.readFileSync(path.join(SRC_DIR, f), "utf8");
    const { meta, body } = parseFrontMatter(text);
    if (!meta.title || !meta.date) {
      console.error(`title / date が未設定です: ${f}`);
      process.exit(1);
    }
    if (meta.draft === "true") continue;
    fs.writeFileSync(path.join(OUT_DIR, `${slug}.html`), renderArticle(slug, meta, body), "utf8");
    posts.push({ slug, meta });
  }
  posts.sort((a, b) => (a.meta.order || "999").localeCompare(b.meta.order || "999") || b.meta.date.localeCompare(a.meta.date));
  fs.writeFileSync(path.join(OUT_DIR, "index.html"), renderIndex(posts), "utf8");
  console.log(`生成完了: ${posts.length} 記事 → ${path.relative(ROOT, OUT_DIR)}/`);
}

main();
