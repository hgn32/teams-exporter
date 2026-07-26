// ============================================================
// Teams メッセージ抽出ツール - PDF 生成モジュール
//
// background.js (service worker) から importScripts で読み込む
// （docx.js が先。共通ユーティリティ self.TeamsExportShared を使うため）。
//
// docx.jsと同じく外部ライブラリを使わず、PDFのオブジェクト構造
// （カタログ/ページ/フォント/画像XObject/リンク注釈 + xref）を
// 直接組み立てる。
//
// 日本語テキストは「非埋め込みCIDフォント」方式で出力する:
//   Type0フォント + BaseFont MS-Gothic + Encoding UniJIS-UCS2-H。
//   フォントファイルを同梱せず、閲覧側（Acrobat/Chrome/EdgeのPDF
//   ビューア）のフォント置換に表示を任せる、日本語PDFの定石。
//   MS-Gothicは等幅（ASCII=半角0.5em、その他=全角1em）なので、
//   行折り返しの幅計算が単純な文字数ベースで正確に成立する。
//   ※UCS2で表せない絵文字等（サロゲートペア）はAdobe-Japan1に
//     グリフが無いため「□」に置き換える。
//
// 画像はJPEGとしてDCTDecodeで直接埋め込む。JPEG以外（PNG等）は
// service workerで使えるOffscreenCanvasでJPEGへ変換してから埋め込む。
// ============================================================

(() => {
  const shared = self.TeamsExportShared;

  // A4縦・単位はpt
  const PAGE_W = 595.28;
  const PAGE_H = 841.89;
  const MARGIN = 48;
  const CONTENT_W = PAGE_W - MARGIN * 2;
  const REPLY_INDENT = 20; // 返信の左インデント(pt)
  const PX_TO_PT = 0.75; // 画像表示サイズ: 96dpi px → 72dpi pt

  const BLACK = [0, 0, 0];
  const GRAY = [0.45, 0.45, 0.45];
  const LINK_BLUE = [0.02, 0.39, 0.76];

  // XMLと同様、PDFの文字列にも制御文字は入れない。改行(\n)だけは
  // レイアウト側で行送りとして解釈するため残す。サロゲートペア
  // （絵文字等）はUCS2エンコーディングで表せないため□に置き換える
  function sanitizeText(s) {
    return String(s == null ? '' : s)
      .replace(/[\u0000-\u0009\u000B-\u001F\u007F]/g, '')
      .replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, '□')
      .replace(/[\uD800-\uDFFF]/g, '');
  }

  // MS-Gothic(等幅)の幅: ASCIIは半角(0.5em)、それ以外は全角(1em)。
  // CIDフォント側のW配列宣言と一致させること
  function charEm(ch) {
    const code = ch.charCodeAt(0);
    return code >= 0x20 && code <= 0x7e ? 0.5 : 1.0;
  }

  // UniJIS-UCS2-H のテキストはUTF-16BEのhex文字列として書く
  function utf16beHex(s) {
    let hex = '';
    for (let i = 0; i < s.length; i++) {
      hex += s.charCodeAt(i).toString(16).padStart(4, '0');
    }
    return hex;
  }

  // PDFリテラル文字列 ( ) 用のエスケープ。非ASCIIはUTF-8バイトの
  // 8進数表記にする（URI等に使う）
  function pdfString(s) {
    const bytes = new TextEncoder().encode(String(s == null ? '' : s));
    let out = '';
    for (const b of bytes) {
      if (b === 0x28 || b === 0x29 || b === 0x5c) out += '\\' + String.fromCharCode(b);
      else if (b < 0x20 || b > 0x7e) out += '\\' + b.toString(8).padStart(3, '0');
      else out += String.fromCharCode(b);
    }
    return out;
  }

  // 座標・寸法の数値表記（指数表記を避け小数2桁に丸める）
  function fmt(n) {
    return String(Math.round(n * 100) / 100);
  }

  // 画像をPDFに直接埋め込めるJPEGへ正規化する。
  // - 元からJPEGなら無変換（寸法はバイナリヘッダから取得）
  // - それ以外はOffscreenCanvasでJPEG化（JPEGは透過を持てないため白地に合成）
  // - 変換できない場合はnull（呼び出し側でリンク表示にフォールバック）
  async function toJpeg(dataUri) {
    const parsed = shared.dataUriToBytes(dataUri);
    if (!parsed) return null;

    if (parsed.mime === 'image/jpeg' || parsed.mime === 'image/jpg') {
      const size = shared.getImageSizePx(parsed.bytes);
      if (size && size.w > 0 && size.h > 0) {
        return { bytes: parsed.bytes, w: size.w, h: size.h };
      }
    }

    if (typeof createImageBitmap !== 'function' || typeof OffscreenCanvas !== 'function') {
      return null;
    }
    const bitmap = await createImageBitmap(new Blob([parsed.bytes], { type: parsed.mime }));
    const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, bitmap.width, bitmap.height);
    ctx.drawImage(bitmap, 0, 0);
    const blob = await canvas.convertToBlob({ type: 'image/jpeg', quality: 0.9 });
    return { bytes: new Uint8Array(await blob.arrayBuffer()), w: bitmap.width, h: bitmap.height };
  }

  async function buildPdf(opts) {
    const title = opts.title || 'Teams チャット抽出結果';
    const messages = opts.messages || [];

    // 画像のJPEG化は非同期なので、レイアウト前にまとめて済ませておく
    //（レイアウト処理自体は同期で書けるようにするため）
    const jpegByImage = new Map();
    const jobs = [];
    for (const m of messages) {
      for (const img of m.images || []) {
        if (!img.dataUri) continue;
        jobs.push(
          toJpeg(img.dataUri)
            .then((r) => jpegByImage.set(img, r))
            .catch(() => jpegByImage.set(img, null))
        );
      }
    }
    await Promise.all(jobs);

    // ---- オブジェクトテーブル（index = オブジェクト番号） ----
    const objects = [null];
    function reserve() {
      objects.push(null);
      return objects.length - 1;
    }

    const catalogNum = reserve();
    const pagesNum = reserve();
    const infoNum = reserve();
    const fontNum = reserve();
    const cidFontNum = reserve();
    const fdNum = reserve();

    objects[fontNum] =
      `<< /Type /Font /Subtype /Type0 /BaseFont /MS-Gothic /Encoding /UniJIS-UCS2-H ` +
      `/DescendantFonts [${cidFontNum} 0 R] >>`;
    // W配列: ASCII相当のCID（プロポーショナル欧文1-95と半角欧文231-325の
    // 両レンジ）を半角500に宣言し、charEm()の幅計算と一致させる。
    // それ以外はDW=1000（全角）
    objects[cidFontNum] =
      `<< /Type /Font /Subtype /CIDFontType0 /BaseFont /MS-Gothic ` +
      `/CIDSystemInfo << /Registry (Adobe) /Ordering (Japan1) /Supplement 6 >> ` +
      `/FontDescriptor ${fdNum} 0 R /DW 1000 /W [1 95 500 231 325 500] >>`;
    // 非埋め込み（FontFile無し）。数値はMS-Gothic相当の概算メトリクス
    objects[fdNum] =
      `<< /Type /FontDescriptor /FontName /MS-Gothic /Flags 4 ` +
      `/FontBBox [-121 -137 1000 859] /ItalicAngle 0 /Ascent 859 /Descent -141 ` +
      `/CapHeight 769 /StemV 78 >>`;
    objects[infoNum] =
      `<< /Title <FEFF${utf16beHex(sanitizeText(title).replace(/\n/g, ' '))}> ` +
      `/Producer (Teams Message Exporter) >>`;

    // ---- ページレイアウト ----
    const pages = []; // {ops:[], annots:[], imgs:[], pageNum}
    let cur = null;
    let cursorY = 0; // 次に置く行の「上端」のy座標（PDF座標系: 下が0）

    function newPage() {
      cur = { ops: [], annots: [], imgs: [] };
      pages.push(cur);
      cursorY = PAGE_H - MARGIN;
    }
    newPage();

    // 1行分の確定済みチャンク（x位置・幅計算済み）を描画コマンドにする
    function flushLine(lineRuns, lineH) {
      if (!lineRuns.length) {
        cursorY -= lineH; // 空行（連続改行）も行送りだけは進める
        return;
      }
      if (cursorY - lineH < MARGIN) newPage();
      const baseline = cursorY - lineH * 0.8;
      for (const r of lineRuns) {
        const rgb = (r.color || BLACK).map(fmt).join(' ');
        let op = `BT /F1 ${fmt(r.size)} Tf ${rgb} rg`;
        // 太字フォントを持たないため、塗り+輪郭の描画モードで疑似ボールドにする。
        // Trや線色はBT/ETを跨いで持続するグラフィックス状態のため、
        // 非ボールド側でも毎回明示的に通常モード(0 Tr)へ戻す
        if (r.bold) op += ` ${rgb} RG 2 Tr ${fmt(r.size * 0.03)} w`;
        else op += ` 0 Tr`;
        op += ` 1 0 0 1 ${fmt(r.x)} ${fmt(baseline)} Tm <${utf16beHex(r.text)}> Tj ET`;
        cur.ops.push(op);
        if (r.href) {
          cur.ops.push(
            `${rgb} RG 0.5 w ${fmt(r.x)} ${fmt(baseline - 1.5)} m ${fmt(r.x + r.w)} ${fmt(baseline - 1.5)} l S`
          );
          cur.annots.push({
            rect: [r.x, baseline - 2, r.x + r.w, baseline + r.size],
            uri: r.href,
          });
        }
      }
      cursorY -= lineH;
    }

    // 1段落分のラン（テキスト/リンク）を折り返しながら配置する。
    // MS-Gothicは等幅なので、幅は文字種（半角/全角）だけで決まる
    function layoutRuns(runs, opt) {
      const o = opt || {};
      const left = MARGIN + (o.indent || 0);
      const right = PAGE_W - MARGIN;
      const lineH = Math.max.apply(null, runs.map((r) => r.size || 10)) * 1.45;
      let line = [];
      let x = left;

      function breakLine() {
        flushLine(line, lineH);
        line = [];
        x = left;
      }

      for (const run of runs) {
        const size = run.size || 10;
        const text = sanitizeText(run.text);
        let chunk = '';
        let chunkW = 0;

        function emitChunk() {
          if (!chunk) return;
          line.push({ text: chunk, size, bold: run.bold, color: run.color, href: run.href, x, w: chunkW });
          x += chunkW;
          chunk = '';
          chunkW = 0;
        }

        for (const ch of text) {
          if (ch === '\n') {
            emitChunk();
            breakLine();
            continue;
          }
          const w = charEm(ch) * size;
          if (x + chunkW + w > right && (chunk || line.length)) {
            emitChunk();
            breakLine();
          }
          chunk += ch;
          chunkW += w;
        }
        emitChunk();
      }
      if (line.length) flushLine(line, lineH);
    }

    // JPEG化済み画像をXObjectとして登録し、表示コマンドを置く
    function layoutImage(jpeg, indent) {
      const maxW = CONTENT_W - indent;
      const maxH = PAGE_H - MARGIN * 2 - 4;
      let w = jpeg.w * PX_TO_PT;
      let h = jpeg.h * PX_TO_PT;
      if (w > maxW) {
        h *= maxW / w;
        w = maxW;
      }
      if (h > maxH) {
        w *= maxH / h;
        h = maxH;
      }
      if (cursorY - h - 4 < MARGIN) newPage();

      const objNum = reserve();
      objects[objNum] = {
        dict:
          `<< /Type /XObject /Subtype /Image /Width ${jpeg.w} /Height ${jpeg.h} ` +
          `/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${jpeg.bytes.length} >>`,
        stream: jpeg.bytes,
      };
      const name = 'Im' + objNum;
      cur.imgs.push({ name, objNum });
      cursorY -= h;
      cur.ops.push(`q ${fmt(w)} 0 0 ${fmt(h)} ${fmt(MARGIN + indent)} ${fmt(cursorY)} cm /${name} Do Q`);
      cursorY -= 4;
    }

    // ---- 本文 ----

    layoutRuns([{ text: title, size: 14, bold: true, color: BLACK }]);
    cursorY -= 2;
    layoutRuns([
      {
        text: `抽出元: ${opts.sourceUrl || ''} / 抽出日時: ${opts.exportedAt || ''} / 件数: ${messages.length}`,
        size: 7.5,
        color: GRAY,
      },
    ]);
    cursorY -= 8;

    for (const m of messages) {
      const indent = m.kind === '返信' ? REPLY_INDENT : 0;
      cursorY -= 6; // メッセージ間の余白

      // メタ行: [種別] 送信者  時刻
      const metaRuns = [];
      if (m.kind) metaRuns.push({ text: `[${m.kind}] `, size: 9, color: GRAY });
      metaRuns.push({ text: m.author || '(不明な送信者)', size: 10, bold: true, color: BLACK });
      const time = m.displayTime || m.isoTime || '';
      if (time) metaRuns.push({ text: '  ' + time, size: 9, color: GRAY });
      layoutRuns(metaRuns, { indent });

      // 本文（bodyHtmlがあればリンク・改行を復元、なければプレーンテキスト）
      const segments = m.bodyHtml
        ? shared.parseBodyHtml(m.bodyHtml)
        : [{ type: 'text', text: m.text || '' }];
      const bodyRuns = [];
      for (const seg of segments) {
        if (seg.type === 'br') {
          bodyRuns.push({ text: '\n', size: 10, color: BLACK });
        } else if (seg.type === 'link' && shared.isLinkableHref(seg.href)) {
          bodyRuns.push({ text: seg.text, size: 10, color: LINK_BLUE, href: seg.href });
        } else {
          bodyRuns.push({ text: seg.text, size: 10, color: BLACK });
        }
      }
      if (bodyRuns.some((r) => r.text)) layoutRuns(bodyRuns, { indent });

      // 画像: JPEG化できたものは埋め込み、できなかったものは元URLへのリンク
      //（URLがhttp(s)でない場合はリンクにできないためグレーのテキスト表示）
      for (const img of m.images || []) {
        const jpeg = jpegByImage.get(img);
        if (jpeg) {
          layoutImage(jpeg, indent);
          continue;
        }
        const label = `[画像] ${img.alt || '(画像)'}`;
        if (shared.isLinkableHref(img.src)) {
          layoutRuns([{ text: label, size: 10, color: LINK_BLUE, href: img.src }], { indent });
        } else {
          layoutRuns([{ text: `${label} (取得失敗)`, size: 10, color: GRAY }], { indent });
        }
      }

      // 本文外のファイル添付カード
      for (const f of m.files || []) {
        const label = `[添付] ${f.name || f.href}`;
        if (shared.isLinkableHref(f.href)) {
          layoutRuns([{ text: label, size: 10, color: LINK_BLUE, href: f.href }], { indent });
        } else {
          layoutRuns([{ text: label, size: 10, color: GRAY }], { indent });
        }
      }
    }

    // ---- ページオブジェクト化 ----
    const encoder = new TextEncoder();
    for (const p of pages) {
      const content = encoder.encode(p.ops.join('\n'));
      const contentNum = reserve();
      objects[contentNum] = { dict: `<< /Length ${content.length} >>`, stream: content };

      const pageNum = reserve();
      p.pageNum = pageNum;
      const xobj = p.imgs.length
        ? ` /XObject << ${p.imgs.map((i) => `/${i.name} ${i.objNum} 0 R`).join(' ')} >>`
        : '';
      const annots = p.annots.length
        ? ` /Annots [ ${p.annots
            .map(
              (a) =>
                `<< /Type /Annot /Subtype /Link /Rect [${a.rect.map(fmt).join(' ')}] ` +
                `/Border [0 0 0] /A << /S /URI /URI (${pdfString(a.uri)}) >> >>`
            )
            .join(' ')} ]`
        : '';
      objects[pageNum] =
        `<< /Type /Page /Parent ${pagesNum} 0 R /MediaBox [0 0 ${PAGE_W} ${PAGE_H}] ` +
        `/Resources << /Font << /F1 ${fontNum} 0 R >>${xobj} >> ` +
        `/Contents ${contentNum} 0 R${annots} >>`;
    }
    objects[pagesNum] = `<< /Type /Pages /Kids [${pages.map((p) => `${p.pageNum} 0 R`).join(' ')}] /Count ${pages.length} >>`;
    objects[catalogNum] = `<< /Type /Catalog /Pages ${pagesNum} 0 R >>`;

    // ---- シリアライズ（ヘッダ + 各オブジェクト + xref + trailer） ----
    const chunks = [];
    let offset = 0;
    function push(bytes) {
      chunks.push(bytes);
      offset += bytes.length;
    }
    // ヘッダ。2行目はバイナリファイルであることを示す慣習のコメント
    push(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31, 0x2e, 0x37, 0x0a, 0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a]));

    const offsets = [0];
    for (let i = 1; i < objects.length; i++) {
      offsets[i] = offset;
      const body = objects[i];
      push(encoder.encode(`${i} 0 obj\n`));
      if (body && typeof body === 'object' && body.stream) {
        push(encoder.encode(body.dict + '\nstream\n'));
        push(body.stream);
        push(encoder.encode('\nendstream'));
      } else {
        push(encoder.encode(String(body)));
      }
      push(encoder.encode('\nendobj\n'));
    }

    const xrefStart = offset;
    let xref = `xref\n0 ${objects.length}\n0000000000 65535 f \n`;
    for (let i = 1; i < objects.length; i++) {
      xref += `${String(offsets[i]).padStart(10, '0')} 00000 n \n`;
    }
    xref +=
      `trailer\n<< /Size ${objects.length} /Root ${catalogNum} 0 R /Info ${infoNum} 0 R >>\n` +
      `startxref\n${xrefStart}\n%%EOF\n`;
    push(encoder.encode(xref));

    const out = new Uint8Array(offset);
    let pos = 0;
    for (const ch of chunks) {
      out.set(ch, pos);
      pos += ch.length;
    }
    return out;
  }

  self.TeamsPdf = { buildPdf };
})();
