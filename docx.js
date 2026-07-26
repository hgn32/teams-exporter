// ============================================================
// Teams メッセージ抽出ツール - Word (.docx) 生成モジュール
//
// background.js (service worker) から importScripts で読み込む。
// .docx の実体は OOXML (XML群) を ZIP に固めたものなので、外部
// ライブラリを使わずに「無圧縮ZIP + 最小限のWordprocessingML」を
// ここで自前生成する（外部サーバーへの送信を一切しない方針のため、
// CDN等からライブラリを読み込む選択肢はとらない）。
//
// 生成する内容:
// - メッセージごとに「送信者・時刻」行 + 本文 + 画像 + 添付リンク
// - 本文中のリンク・添付ファイルはWord上でもクリックできるハイパーリンク
// - base64取得済みの画像はWord文書内に埋め込み（word/media/）
// - 返信メッセージは左インデントで区別
// ============================================================

(() => {
  // ---------- ZIP (無圧縮) ----------

  // ZIPの各エントリに必須のCRC-32。標準的なテーブル方式
  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    let c = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) {
      c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    }
    return (c ^ 0xffffffff) >>> 0;
  }

  // entries: [{ name: string, bytes: Uint8Array }] を無圧縮(stored)ZIPにする。
  // 画像は元々圧縮済み(PNG/JPEG)でXML部分のサイズは相対的に小さいため、
  // Deflate実装を持ち込まず無圧縮で十分と判断
  function buildZip(entries) {
    const encoder = new TextEncoder();
    const chunks = [];
    const central = [];
    let offset = 0;

    const now = new Date();
    const dosTime =
      ((now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1)) & 0xffff;
    const dosDate =
      (((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate()) & 0xffff;

    for (const e of entries) {
      const nameBytes = encoder.encode(e.name);
      const crc = crc32(e.bytes);
      const local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true); // local file header signature
      local.setUint16(4, 20, true); // version needed to extract
      local.setUint16(6, 0x0800, true); // general purpose flag: UTF-8ファイル名
      local.setUint16(8, 0, true); // method: stored
      local.setUint16(10, dosTime, true);
      local.setUint16(12, dosDate, true);
      local.setUint32(14, crc, true);
      local.setUint32(18, e.bytes.length, true); // compressed size (= raw)
      local.setUint32(22, e.bytes.length, true); // uncompressed size
      local.setUint16(26, nameBytes.length, true);
      local.setUint16(28, 0, true); // extra length
      chunks.push(new Uint8Array(local.buffer), nameBytes, e.bytes);
      central.push({ nameBytes, crc, size: e.bytes.length, offset });
      offset += 30 + nameBytes.length + e.bytes.length;
    }

    const centralStart = offset;
    for (const c of central) {
      const h = new DataView(new ArrayBuffer(46));
      h.setUint32(0, 0x02014b50, true); // central directory header signature
      h.setUint16(4, 20, true); // version made by
      h.setUint16(6, 20, true); // version needed
      h.setUint16(8, 0x0800, true);
      h.setUint16(10, 0, true);
      h.setUint16(12, dosTime, true);
      h.setUint16(14, dosDate, true);
      h.setUint32(16, c.crc, true);
      h.setUint32(20, c.size, true);
      h.setUint32(24, c.size, true);
      h.setUint16(28, c.nameBytes.length, true);
      h.setUint32(42, c.offset, true);
      chunks.push(new Uint8Array(h.buffer), c.nameBytes);
      offset += 46 + c.nameBytes.length;
    }

    const eocd = new DataView(new ArrayBuffer(22));
    eocd.setUint32(0, 0x06054b50, true); // end of central directory signature
    eocd.setUint16(8, central.length, true);
    eocd.setUint16(10, central.length, true);
    eocd.setUint32(12, offset - centralStart, true);
    eocd.setUint32(16, centralStart, true);
    chunks.push(new Uint8Array(eocd.buffer));
    offset += 22;

    const out = new Uint8Array(offset);
    let pos = 0;
    for (const ch of chunks) {
      out.set(ch, pos);
      pos += ch.length;
    }
    return out;
  }

  // ---------- ユーティリティ ----------

  // XMLに書けない制御文字はWord側で「ファイルが壊れている」扱いに
  // なるため、エスケープと同時に除去する
  function xmlEscape(s) {
    return String(s == null ? '' : s)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  // bodyToSafeHtml(content.js)が出力するのは「エスケープ済みテキスト +
  // <br> + <a href>」だけなので、その3種だけ元に戻せればよい
  function htmlUnescape(s) {
    return String(s || '')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&amp;/g, '&');
  }

  // bodyHtml をテキスト/リンク/改行のセグメント列に分解する
  function parseBodyHtml(bodyHtml) {
    const segments = [];
    const re = /<a href="([^"]*)"[^>]*>([\s\S]*?)<\/a>|<br>/g;
    let last = 0;
    let m;
    while ((m = re.exec(bodyHtml)) !== null) {
      if (m.index > last) {
        segments.push({ type: 'text', text: htmlUnescape(bodyHtml.slice(last, m.index)) });
      }
      if (m[0] === '<br>') {
        segments.push({ type: 'br' });
      } else {
        // リンク内テキストに<br>が入っている場合は改行として残す
        const inner = htmlUnescape(m[2].replace(/<br>/g, '\n').replace(/<[^>]+>/g, ''));
        const href = htmlUnescape(m[1]);
        segments.push({ type: 'link', href, text: inner || href });
      }
      last = re.lastIndex;
    }
    if (last < bodyHtml.length) {
      segments.push({ type: 'text', text: htmlUnescape(bodyHtml.slice(last)) });
    }
    return segments;
  }

  function dataUriToBytes(dataUri) {
    const m = /^data:([^;,]*);base64,([\s\S]*)$/.exec(dataUri);
    if (!m) return null;
    try {
      const bin = atob(m[2]);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return { mime: (m[1] || '').toLowerCase(), bytes };
    } catch (e) {
      return null;
    }
  }

  // 画像のピクセル寸法をバイナリヘッダから読み取る。service workerには
  // <img>やImageオブジェクトが無いため、フォーマット別に自前でパースする。
  // 取れなければnull（呼び出し側で既定サイズにフォールバック）
  function getImageSizePx(bytes) {
    try {
      const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      // PNG: IHDRチャンク先頭（オフセット16）に幅・高さ（ビッグエンディアン）
      if (bytes.length > 24 && bytes[0] === 0x89 && bytes[1] === 0x50) {
        return { w: dv.getUint32(16), h: dv.getUint32(20) };
      }
      // GIF: オフセット6に幅・高さ（リトルエンディアン）
      if (bytes.length > 10 && bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) {
        return { w: dv.getUint16(6, true), h: dv.getUint16(8, true) };
      }
      // BMP: BITMAPINFOHEADERのオフセット18/22（高さは負値=トップダウンあり）
      if (bytes.length > 26 && bytes[0] === 0x42 && bytes[1] === 0x4d) {
        return { w: Math.abs(dv.getInt32(18, true)), h: Math.abs(dv.getInt32(22, true)) };
      }
      // JPEG: SOFマーカー(C0-CF、ただしC4/C8/CCはSOFではない)を走査
      if (bytes.length > 4 && bytes[0] === 0xff && bytes[1] === 0xd8) {
        let i = 2;
        while (i + 9 < bytes.length) {
          if (bytes[i] !== 0xff) {
            i++;
            continue;
          }
          const marker = bytes[i + 1];
          if (marker === 0xff) {
            i++;
            continue;
          }
          if (marker >= 0xd0 && marker <= 0xd9) {
            i += 2; // RST/SOI/EOIは長さフィールドを持たない
            continue;
          }
          if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
            return { h: dv.getUint16(i + 5), w: dv.getUint16(i + 7) };
          }
          i += 2 + dv.getUint16(i + 2);
        }
      }
    } catch (e) {
      /* フォールバックに任せる */
    }
    return null;
  }

  // ---------- WordprocessingML 生成 ----------

  const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const REL_HYPERLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
  const REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
  const REL_OFFICE_DOC = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

  // Wordが確実に扱える画像形式のみ埋め込む（それ以外はリンク表示に
  // フォールバック）。拡張子は[Content_Types].xmlのDefault宣言と一致させる
  const MIME_EXT = {
    'image/png': 'png',
    'image/jpeg': 'jpg',
    'image/jpg': 'jpg',
    'image/gif': 'gif',
    'image/bmp': 'bmp',
  };

  // レイアウト定数
  const REPLY_INDENT_TWIP = 567; // 返信の左インデント（1cm）
  const MAX_IMAGE_WIDTH_EMU = 5400000; // 画像の最大幅15cm（A4本文幅に収める）
  const EMU_PER_PX = 9525; // 96dpi換算

  // ハイパーリンクにしてよいURLか。Relationshipsに不正なターゲットを
  // 書くとWordがファイル全体を破損扱いすることがあるため、通常の
  // http(s)/mailto以外（blob:や巨大なdata:等）はテキスト表示に落とす
  function isLinkableHref(href) {
    return /^(https?:|mailto:)/i.test(String(href || ''));
  }

  function buildDocx(opts) {
    const title = opts.title || 'Teams チャット抽出結果';
    const messages = opts.messages || [];

    // word/_rels/document.xml.rels に書く関係(画像・ハイパーリンク)
    const rels = [];
    const media = []; // word/media/ に格納する画像 [{name, bytes}]
    const linkRelIds = new Map(); // 同一URLのrIdを使い回す

    function addRel(type, target, external) {
      const id = 'rId' + (rels.length + 1);
      rels.push({ id, type, target, external: !!external });
      return id;
    }

    function linkRelId(url) {
      if (!linkRelIds.has(url)) {
        linkRelIds.set(url, addRel(REL_HYPERLINK, url, true));
      }
      return linkRelIds.get(url);
    }

    function runProps(style) {
      const s = style || {};
      const props = [];
      if (s.bold) props.push('<w:b/>');
      if (s.color) props.push(`<w:color w:val="${s.color}"/>`);
      if (s.sizeHalfPt) props.push(`<w:sz w:val="${s.sizeHalfPt}"/><w:szCs w:val="${s.sizeHalfPt}"/>`);
      if (s.underline) props.push('<w:u w:val="single"/>');
      return props.length ? `<w:rPr>${props.join('')}</w:rPr>` : '';
    }

    // テキスト中の改行(\n)は<w:br/>に変換する
    function textRun(text, style) {
      const rpr = runProps(style);
      return String(text)
        .split('\n')
        .map(
          (part, idx) =>
            `<w:r>${rpr}${idx > 0 ? '<w:br/>' : ''}${
              part ? `<w:t xml:space="preserve">${xmlEscape(part)}</w:t>` : ''
            }</w:r>`
        )
        .join('');
    }

    function linkRun(href, text, style) {
      if (!isLinkableHref(href)) {
        return textRun(text, style);
      }
      const merged = Object.assign({ color: '0563C1', underline: true }, style);
      return `<w:hyperlink r:id="${linkRelId(href)}">${textRun(text, merged)}</w:hyperlink>`;
    }

    function para(childrenXml, opt) {
      const o = opt || {};
      const pPr = [];
      const spacing = [];
      if (o.spacingBefore != null) spacing.push(`w:before="${o.spacingBefore}"`);
      if (o.spacingAfter != null) spacing.push(`w:after="${o.spacingAfter}"`);
      if (spacing.length) pPr.push(`<w:spacing ${spacing.join(' ')}/>`);
      if (o.indent) pPr.push(`<w:ind w:left="${o.indent}"/>`);
      return `<w:p>${pPr.length ? `<w:pPr>${pPr.join('')}</w:pPr>` : ''}${childrenXml}</w:p>`;
    }

    // data URIの画像をword/mediaに追加し、埋め込み表示する段落を返す。
    // 埋め込めない形式・壊れたdata URIの場合はnull（呼び出し側でリンクに
    // フォールバック）
    function imageParagraph(dataUri, alt, indent) {
      const parsed = dataUriToBytes(dataUri);
      if (!parsed || !MIME_EXT[parsed.mime]) return null;
      const name = `image${media.length + 1}.${MIME_EXT[parsed.mime]}`;
      const relId = addRel(REL_IMAGE, `media/${name}`, false);
      media.push({ name, bytes: parsed.bytes });

      const px = getImageSizePx(parsed.bytes) || { w: 480, h: 320 };
      let cx = Math.max(1, px.w * EMU_PER_PX);
      let cy = Math.max(1, px.h * EMU_PER_PX);
      if (cx > MAX_IMAGE_WIDTH_EMU) {
        cy = Math.max(1, Math.round((cy * MAX_IMAGE_WIDTH_EMU) / cx));
        cx = MAX_IMAGE_WIDTH_EMU;
      }

      const docPrId = media.length; // 文書内で一意なら良いので連番
      const label = xmlEscape(alt || name);
      const drawing =
        `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
        `<wp:extent cx="${cx}" cy="${cy}"/>` +
        `<wp:docPr id="${docPrId}" name="${label}"/>` +
        `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
        `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
        `<pic:nvPicPr><pic:cNvPr id="${docPrId}" name="${label}"/><pic:cNvPicPr/></pic:nvPicPr>` +
        `<pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
        `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
        `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
      return para(`<w:r>${drawing}</w:r>`, { indent, spacingAfter: 80 });
    }

    // ---- 本文組み立て ----

    const body = [];

    body.push(para(textRun(title, { bold: true, sizeHalfPt: 32 }), { spacingAfter: 80 }));
    body.push(
      para(
        textRun(
          `抽出元: ${opts.sourceUrl || ''} / 抽出日時: ${opts.exportedAt || ''} / 件数: ${messages.length}`,
          { color: '888888', sizeHalfPt: 16 }
        ),
        { spacingAfter: 240 }
      )
    );

    for (const m of messages) {
      const indent = m.kind === '返信' ? REPLY_INDENT_TWIP : 0;

      // メタ行: [種別] 送信者  時刻
      const metaRuns = [];
      if (m.kind) metaRuns.push(textRun(`[${m.kind}] `, { color: '888888', sizeHalfPt: 18 }));
      metaRuns.push(textRun(m.author || '(不明な送信者)', { bold: true, sizeHalfPt: 20 }));
      const time = m.displayTime || m.isoTime || '';
      if (time) metaRuns.push(textRun('  ' + time, { color: '888888', sizeHalfPt: 18 }));
      body.push(para(metaRuns.join(''), { indent, spacingBefore: 200, spacingAfter: 40 }));

      // 本文（bodyHtmlがあればリンク・改行を復元、なければプレーンテキスト）
      const segments = m.bodyHtml
        ? parseBodyHtml(m.bodyHtml)
        : [{ type: 'text', text: m.text || '' }];
      const bodyRuns = segments
        .map((seg) => {
          if (seg.type === 'br') return '<w:r><w:br/></w:r>';
          if (seg.type === 'link') return linkRun(seg.href, seg.text);
          return textRun(seg.text);
        })
        .join('');
      if (bodyRuns) body.push(para(bodyRuns, { indent, spacingAfter: 40 }));

      // 画像: base64取得済みは埋め込み、失敗分は元URLへのリンク
      //（URLがhttp(s)でない場合はリンクにできないためグレーのテキスト表示）
      for (const img of m.images || []) {
        let p = null;
        if (img.dataUri) p = imageParagraph(img.dataUri, img.alt, indent);
        if (!p) {
          const label = `🖼️ ${img.alt || '(画像)'}`;
          const runs = isLinkableHref(img.src)
            ? linkRun(img.src, label)
            : textRun(`${label} (取得失敗)`, { color: '888888' });
          p = para(runs, { indent, spacingAfter: 40 });
        }
        body.push(p);
      }

      // 本文外のファイル添付カード
      for (const f of m.files || []) {
        body.push(para(linkRun(f.href, `📎 ${f.name || f.href}`), { indent, spacingAfter: 40 }));
      }
    }

    const documentXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="${NS_WP}">` +
      `<w:body>${body.join('')}` +
      `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>` + // A4縦
      `<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/>` +
      `</w:sectPr></w:body></w:document>`;

    const contentTypesXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Default Extension="png" ContentType="image/png"/>` +
      `<Default Extension="jpg" ContentType="image/jpeg"/>` +
      `<Default Extension="gif" ContentType="image/gif"/>` +
      `<Default Extension="bmp" ContentType="image/bmp"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `</Types>`;

    const rootRelsXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL_OFFICE_DOC}" Target="word/document.xml"/>` +
      `</Relationships>`;

    const docRelsXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      rels
        .map(
          (r) =>
            `<Relationship Id="${r.id}" Type="${r.type}" Target="${xmlEscape(r.target)}"${
              r.external ? ' TargetMode="External"' : ''
            }/>`
        )
        .join('') +
      `</Relationships>`;

    const encoder = new TextEncoder();
    const entries = [
      { name: '[Content_Types].xml', bytes: encoder.encode(contentTypesXml) },
      { name: '_rels/.rels', bytes: encoder.encode(rootRelsXml) },
      { name: 'word/document.xml', bytes: encoder.encode(documentXml) },
      { name: 'word/_rels/document.xml.rels', bytes: encoder.encode(docRelsXml) },
    ];
    for (const img of media) {
      entries.push({ name: `word/media/${img.name}`, bytes: img.bytes });
    }

    return buildZip(entries);
  }

  self.TeamsDocx = { buildDocx };

  // pdf.js と共有するユーティリティ（backgroundは docx.js → pdf.js の順に
  // importScripts するため、pdf.js からはこの名前空間が見える）
  self.TeamsExportShared = {
    parseBodyHtml,
    htmlUnescape,
    dataUriToBytes,
    getImageSizePx,
    isLinkableHref,
  };
})();
