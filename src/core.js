/* ===================================================================
 * Paper2Docx core — 组卷网试卷网页 -> Word(.docx) 结构化转换引擎
 * 纯前端、零依赖：解析 DOM -> 生成 OOXML -> 自己打包 ZIP
 * =================================================================== */
(function (global) {
  'use strict';

  var EMU = 360000;            // 1cm = 360000 EMU
  var PT2CM = 2.54 / 72;
  var TWIP = 567;              // 1cm = 567 twips

  var DEFAULT_OPTIONS = {
    fontSize: 12,             // 小四=12, 四号=14, 五号=10.5, 三号=16
    lineSpacing: 1.05,        // 行距倍数
    imgScale: 1.0,            // 图片额外缩放
    margin: 1.9,              // 页边距 cm
    includeAnswers: false,    // 是否导出【答案】【解析】
    questionGap: 6,           // 题间距 pt
    titleFont: '黑体',
    bodyFont: '宋体',
    latinFont: 'Times New Roman',
    showPageNumber: true,
    keepImageWithText: true,  // 图与前文不跨页拆开
    autoOptionColumns: true   // 按网页 cols 排选项
  };

  /* ---------------- 小工具 ---------------- */
  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function tw(cm) { return Math.round(cm * TWIP); }
  function emu(cm) { return Math.round(cm * EMU); }
  function log() { }

  /* ---------------- CRC32 + ZIP ---------------- */
  var CRC_TABLE = (function () {
    var t = new Int32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c;
    }
    return t;
  })();

  function crc32(u8) {
    var c = -1;
    for (var i = 0; i < u8.length; i++) c = CRC_TABLE[(c ^ u8[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ -1) >>> 0;
  }

  function u8(str) {
    return new TextEncoder().encode(str);
  }

  async function deflateRaw(bytes) {
    if (typeof CompressionStream === 'undefined') return null;
    try {
      var cs = new CompressionStream('deflate-raw');
      var stream = new Blob([bytes]).stream().pipeThrough(cs);
      var buf = await new Response(stream).arrayBuffer();
      return new Uint8Array(buf);
    } catch (e) { return null; }
  }

  /** entries: [{name, data:Uint8Array, store:bool}] -> Blob(zip) */
  async function makeZip(entries) {
    var chunks = [];
    var central = [];
    var offset = 0;
    var now = new Date();
    var dosTime = ((now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1)) & 0xFFFF;
    var dosDate = (((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate()) & 0xFFFF;

    for (var i = 0; i < entries.length; i++) {
      var e = entries[i];
      var data = e.data;
      var method = 0;
      var comp = data;
      if (!e.store) {
        var d = await deflateRaw(data);
        if (d && d.length < data.length) { comp = d; method = 8; }
      }
      var nameBytes = u8(e.name);
      var crc = crc32(data);

      var lh = new Uint8Array(30 + nameBytes.length);
      var dv = new DataView(lh.buffer);
      dv.setUint32(0, 0x04034b50, true);
      dv.setUint16(4, 20, true);
      dv.setUint16(6, 0x0800, true);        // UTF-8 filename
      dv.setUint16(8, method, true);
      dv.setUint16(10, dosTime, true);
      dv.setUint16(12, dosDate, true);
      dv.setUint32(14, crc, true);
      dv.setUint32(18, comp.length, true);
      dv.setUint32(22, data.length, true);
      dv.setUint16(26, nameBytes.length, true);
      dv.setUint16(28, 0, true);
      lh.set(nameBytes, 30);

      chunks.push(lh, comp);

      var cd = new Uint8Array(46 + nameBytes.length);
      var dv2 = new DataView(cd.buffer);
      dv2.setUint32(0, 0x02014b50, true);
      dv2.setUint16(4, 20, true);
      dv2.setUint16(6, 20, true);
      dv2.setUint16(8, 0x0800, true);
      dv2.setUint16(10, method, true);
      dv2.setUint16(12, dosTime, true);
      dv2.setUint16(14, dosDate, true);
      dv2.setUint32(16, crc, true);
      dv2.setUint32(20, comp.length, true);
      dv2.setUint32(24, data.length, true);
      dv2.setUint16(28, nameBytes.length, true);
      dv2.setUint32(42, offset, true);
      cd.set(nameBytes, 46);
      central.push(cd);

      offset += lh.length + comp.length;
    }

    var cdSize = 0;
    for (var j = 0; j < central.length; j++) cdSize += central[j].length;
    var eocd = new Uint8Array(22);
    var dv3 = new DataView(eocd.buffer);
    dv3.setUint32(0, 0x06054b50, true);
    dv3.setUint16(8, central.length, true);
    dv3.setUint16(10, central.length, true);
    dv3.setUint32(12, cdSize, true);
    dv3.setUint32(16, offset, true);

    return new Blob(chunks.concat(central, [eocd]), { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  }

  /* ---------------- 图片：加载 / 栅格化 ---------------- */
  function isSvg(name, blob) {
    return /\.svg$/i.test(name) || (blob && blob.type === 'image/svg+xml');
  }

  function readImageSize(blob) {
    return new Promise(function (res) {
      var img = new Image();
      var url = URL.createObjectURL(blob);
      img.onload = function () { res({ w: img.naturalWidth || img.width, h: img.naturalHeight || img.height, url: url }); };
      img.onerror = function () { res({ w: 0, h: 0, url: url }); };
      img.src = url;
    });
  }

  function svgIntrinsic(blob) {
    return blob.text().then(function (text) {
      var m = text.match(/<svg[^>]*>/i);
      if (!m) return { w: 0, h: 0 };
      var tag = m[0];
      var w = /width="([\d.]+)/.exec(tag);
      var h = /height="([\d.]+)/.exec(tag);
      if (w && h) return { w: parseFloat(w[1]), h: parseFloat(h[1]) };
      var vb = /viewBox="([-\d.\s]+)"/.exec(tag);
      if (vb) {
        var p = vb[1].trim().split(/\s+/).map(Number);
        if (p.length === 4) return { w: p[2], h: p[3] };
      }
      return { w: 0, h: 0 };
    });
  }

  /** SVG -> 高清 PNG（canvas 矢量渲染，6 倍） */
  async function svgToPng(blob, cssW, cssH, scale) {
    var url = URL.createObjectURL(blob);
    try {
      var img = await new Promise(function (res, rej) {
        var i = new Image();
        i.onload = function () { res(i); };
        i.onerror = function () { rej(new Error('svg load fail')); };
        i.src = url;
      });
      var w = cssW || img.naturalWidth || 40;
      var h = cssH || img.naturalHeight || 20;
      scale = scale || 6;
      var cw = Math.max(8, Math.round(w * scale));
      var ch = Math.max(8, Math.round(h * scale));
      if (cw * ch > 4000000) { var k = Math.sqrt(4000000 / (cw * ch)); cw = Math.round(cw * k); ch = Math.round(ch * k); }
      var c = document.createElement('canvas');
      c.width = cw; c.height = ch;
      var ctx = c.getContext('2d');
      ctx.clearRect(0, 0, cw, ch);
      ctx.drawImage(img, 0, 0, cw, ch);
      var out = await new Promise(function (res) { c.toBlob(function (b) { res(b); }, 'image/png'); });
      if (!out) throw new Error('svg raster fail');
      return { blob: out, ext: 'png' };
    } finally {
      URL.revokeObjectURL(url);
    }
  }

  async function shrinkRaster(blob, maxPx) {
    var info = await readImageSize(blob);
    if (!info.w || info.w <= maxPx) {
      return { blob: blob, ext: mimeExt(blob.type), natW: info.w, natH: info.h };
    }
    var scale = maxPx / info.w;
    var w = Math.round(info.w * scale), h = Math.round(info.h * scale);
    var img = await new Promise(function (res) {
      var i = new Image();
      i.onload = function () { res(i); };
      i.src = info.url;
    });
    var c = document.createElement('canvas');
    c.width = w; c.height = h;
    var ctx = c.getContext('2d');
    ctx.fillStyle = '#fff';
    // 检测是否有透明像素（缩略采样）
    var small = document.createElement('canvas');
    small.width = 48; small.height = 48;
    var sctx = small.getContext('2d');
    sctx.drawImage(img, 0, 0, 48, 48);
    var hasAlpha = false;
    try {
      var d = sctx.getImageData(0, 0, 48, 48).data;
      for (var i = 3; i < d.length; i += 4) { if (d[i] < 250) { hasAlpha = true; break; } }
    } catch (e) { }
    if (!hasAlpha) ctx.fillRect(0, 0, w, h);
    ctx.drawImage(img, 0, 0, w, h);
    var type = hasAlpha ? 'image/png' : 'image/jpeg';
    var out = await new Promise(function (res) { c.toBlob(function (b) { res(b); }, type, 0.92); });
    URL.revokeObjectURL(info.url);
    return { blob: out || blob, ext: mimeExt(type), natW: w, natH: h };
  }

  function mimeExt(t) {
    if (!t) return 'png';
    if (/jpe?g/i.test(t)) return 'jpeg';
    if (/gif/i.test(t)) return 'gif';
    if (/svg/i.test(t)) return 'svg';
    if (/webp/i.test(t)) return 'webp';
    return 'png';
  }

  /* ---------------- HTML 解析 ---------------- */
  function cleanText(s) {
    return s.replace(/[\t\r\n]+/g, ' ').replace(/\u00a0/g, ' ');
  }

  function detectMathImg(el) {
    return /\.svg($|\?)/i.test(el.getAttribute('src') || '');
  }

  function pxAttr(el, name) {
    var v = el.getAttribute(name);
    if (!v) return null;
    var m = /([\d.]+)/.exec(v);
    return m ? parseFloat(m[1]) : null;
  }

  /** 遍历题目内容，产出 block 列表 */
  function Walker(ctx) {
    this.ctx = ctx;
    this.blocks = [];
    this.cur = [];
    this.keepWs = false;
  }

  Walker.prototype.flush = function (align) {
    if (!this.cur.length) return;
    var items = this.cur;
    this.cur = [];
    while (items.length && items[0].t === 'text' && !items[0].text.trim()) items.shift();
    while (items.length && items[items.length - 1].t === 'text' && !items[items.length - 1].text.trim()) items.pop();
    while (items.length && items[items.length - 1].t === 'br') items.pop();
    if (!items.length) return;
    this.blocks.push({ kind: 'para', items: items, align: align || null });
  };

  Walker.prototype.pushText = function (el, text, style) {
    if (!text) return;
    var last = this.cur[this.cur.length - 1];
    if (last && last.t === 'text' && sameStyle(last, style)) {
      last.text += text;
      return;
    }
    this.cur.push({ t: 'text', text: text, b: style.b, i: style.i, u: style.u, va: style.va });
  };

  function sameStyle(a, b) {
    return !!a.b === !!b.b && !!a.i === !!b.i && !!a.u === !!b.u && (a.va || '') === (b.va || '');
  }

  Walker.prototype.walk = function (node, style) {
    style = style || {};
    for (var i = 0; i < node.childNodes.length; i++) {
      var ch = node.childNodes[i];
      if (ch.nodeType === 3) {
        var t = cleanText(ch.nodeValue);
        if (!t) continue;
        if (/^\s+$/.test(t)) {
          // 折叠多余空白，但保留一个空格
          var last = this.cur[this.cur.length - 1];
          if (last && last.t === 'text' && / $/.test(last.text)) continue;
          t = ' ';
        }
        this.pushText(ch, t, style);
        continue;
      }
      if (ch.nodeType !== 1) continue;
      var name = ch.tagName.toLowerCase();
      if (name === 'script' || name === 'style' || name === 'noscript') continue;
      if (ch.hasAttribute('hidden')) {
        var cs = ch.getAttribute('style') || '';
        if (/display\s*:\s*none/i.test(cs)) continue;
      }
      if (/display\s*:\s*none/i.test(ch.getAttribute('style') || '')) continue;

      if (name === 'br') { this.cur.push({ t: 'br' }); continue; }

      if (name === 'img') {
        var src = ch.getAttribute('src') || '';
        var key = this.ctx.resolveFile(src);
        var w = pxAttr(ch, 'width'), h = pxAttr(ch, 'height');
        if (key) {
          this.cur.push({ t: 'img', key: key, w: w, h: h, math: detectMathImg(ch) });
        }
        continue;
      }

      if (name === 'p') {
        this.flush();
        var align = /center/i.test(ch.getAttribute('style') || '') ? 'center' : null;
        this.walk(ch, style);
        this.flush(align);
        continue;
      }

      if (name === 'table') {
        this.flush();
        this.blocks.push(this.table(ch));
        continue;
      }

      if (name === 'sub' || name === 'sup') {
        this.walk(ch, mix(style, { va: name === 'sub' ? 'sub' : 'sup' }));
        continue;
      }
      if (name === 'i' || name === 'em' || name === 'cite') {
        this.walk(ch, mix(style, { i: true }));
        continue;
      }
      if (name === 'b' || name === 'strong' || name === 'dot') {
        this.walk(ch, mix(style, { b: true }));
        continue;
      }
      if (name === 'u' || name === 'ins') {
        this.walk(ch, mix(style, { u: true }));
        continue;
      }
      if (name === 'bk') {
        var size = parseInt(ch.getAttribute('size') || '0', 10);
        var raw = (ch.textContent || '').trim();
        var n = size > 0 ? size : Math.max(4, Math.min(50, raw.length || 8));
        this.pushText(ch, new Array(n + 1).join('_'), style);
        continue;
      }
      if (name === 'span' && ch.getAttribute('qml-space-size')) {
        this.pushText(ch, '        ', style);
        continue;
      }
      if (name === 'div') {
        this.flush();
        this.walk(ch, style);
        this.flush();
        continue;
      }
      this.walk(ch, style);
    }
  };

  function mix(a, b) {
    var o = {};
    for (var k in a) o[k] = a[k];
    for (var k2 in b) o[k2] = b[k2];
    return o;
  }

  Walker.prototype.cellItems = function (cell) {
    var sub = new Walker(this.ctx);
    sub.walk(cell, {});
    sub.flush();
    if (!sub.blocks.length) return [];
    if (sub.blocks.length === 1) return sub.blocks[0].items;
    var items = [];
    for (var i = 0; i < sub.blocks.length; i++) {
      if (sub.blocks[i].kind === 'para') items = items.concat(sub.blocks[i].items);
      if (i < sub.blocks.length - 1) items.push({ t: 'br' });
    }
    return items;
  };

  Walker.prototype.table = function (tbl) {
    var rowsEl = tbl.querySelectorAll('tr');
    var rows = [];
    var maxCols = 1;
    for (var i = 0; i < rowsEl.length; i++) {
      var cellsEl = rowsEl[i].querySelectorAll('td,th');
      var row = [];
      for (var j = 0; j < cellsEl.length; j++) row.push(this.cellItems(cellsEl[j]));
      maxCols = Math.max(maxCols, row.length);
      rows.push(row);
    }
    var tname = tbl.getAttribute('name');
    if (tname === 'optionsTable') {
      var cols = parseInt(tbl.getAttribute('cols') || '', 10);
      return { kind: 'options', cols: cols > 0 ? cols : maxCols, rows: rows };
    }
    // 数据表格：取首行单元格宽度
    var widths = [];
    if (rowsEl.length) {
      var first = rowsEl[0].querySelectorAll('td,th');
      for (var k = 0; k < first.length; k++) {
        var m = /width\s*:\s*([\d.]+)/.exec(first[k].getAttribute('style') || '');
        widths.push(m ? parseFloat(m[1]) : null);
      }
    }
    var tw = null;
    var mt = /width\s*:\s*([\d.]+)px/.exec(tbl.getAttribute('style') || '');
    if (mt) tw = parseFloat(mt[1]);
    return { kind: 'data', rows: rows, widths: widths, tw: tw, ncols: maxCols };
  };

  /* 解析整份试卷 */
  function parsePaper(htmlText, options) {
    var doc = new DOMParser().parseFromString(htmlText, 'text/html');
    var body = doc.querySelector('.paper-body') || doc.body;
    var fileNameMap = null;

    var ctx = {
      resolveFile: function (src) {
        if (!src) return null;
        var s = src.split('#')[0].split('?')[0].replace(/\\/g, '/');
        var base = s.substring(s.lastIndexOf('/') + 1);
        try { base = decodeURIComponent(base); } catch (e) { }
        return base || null;
      }
    };

    var titleEl = doc.querySelector('.paper-title .main-title') || doc.querySelector('.main-title') || doc.querySelector('.paper-title');
    var title = titleEl ? (titleEl.textContent || '').trim() : '试卷';
    var stdEl = doc.querySelector('.std-input');
    var std = null;
    if (stdEl && !stdEl.hasAttribute('hidden')) std = (stdEl.textContent || '').replace(/\s+/g, '').trim();

    var model = { title: title, std: std, sections: [], stats: { questions: 0, figures: 0, mathImages: 0, types: [] } };
    var gIndex = 0;

    function pushQType(qt, partTitle) {
      var head = qt.querySelector('.questypetitle');
      var label = '';
      if (head) {
        var idx = head.querySelector('.questypeindex');
        var nm = head.querySelector('.questypename');
        label = (idx ? idx.textContent : '') + (nm ? nm.textContent : '');
      }
      var qs = qt.querySelectorAll('.tk-quest-item');
      if (!qs.length) return;
      var sec = { type: 'qtype', label: label.trim(), partTitle: partTitle || '', questions: [] };
      for (var i = 0; i < qs.length; i++) {
        var cnt = qs[i].querySelector('.exam-item__cnt');
        if (!cnt) continue;
        var w = new Walker(ctx);
        w.walk(cnt, {});
        w.flush();
        // 去掉末尾换行
        for (var b = 0; b < w.blocks.length; b++) {
          if (w.blocks[b].kind !== 'para') continue;
          var it = w.blocks[b].items;
          while (it.length && it[it.length - 1].t === 'br') it.pop();
        }
        // 答案 / 解析
        var ans = { answer: [], parse: [] };
        if (options.includeAnswers) {
          var an = qs[i].querySelector('.item.answer');
          var pa = qs[i].querySelector('.item.parse');
          [['answer', an], ['parse', pa]].forEach(function (pair) {
            var el = pair[1];
            if (!el) return;
            var txt = (el.textContent || '').replace(/\s+/g, ' ').trim();
            if (!txt) return;
            var ww = new Walker(ctx);
            ww.walk(el, {});
            ww.flush();
            ans[pair[0]] = ww.blocks;
          });
        }
        sec.questions.push({ index: (++gIndex), blocks: w.blocks, answers: ans });
        model.stats.questions++;
        countMedia(w.blocks, model.stats);
        countMedia(ans.answer.concat(ans.parse), model.stats);
      }
      model.stats.types.push({ label: sec.label, count: sec.questions.length });
      model.sections.push(sec);
    }

    var qtypes = body.querySelectorAll('.ques-type');
    if (qtypes.length) {
      var parts = body.querySelectorAll('.paper-part');
      if (parts.length) {
        for (var p = 0; p < parts.length; p++) {
          var pn = parts[p].querySelector('.part-name');
          var partTitle = '';
          var ph = parts[p].querySelector('.part-head');
          if (pn && !(ph && ph.hasAttribute('hidden'))) partTitle = (pn.textContent || '').trim();
          var inner = parts[p].querySelectorAll('.ques-type');
          for (var q = 0; q < inner.length; q++) pushQType(inner[q], partTitle);
        }
      } else {
        for (var t = 0; t < qtypes.length; t++) pushQType(qtypes[t], '');
      }
    } else {
      // 退化：直接抓题块
      var all = body.querySelectorAll('.tk-quest-item');
      if (all.length) {
        var sec2 = { type: 'qtype', label: '', partTitle: '', questions: [] };
        for (var a = 0; a < all.length; a++) {
          var c2 = all[a].querySelector('.exam-item__cnt');
          if (!c2) continue;
          var w2 = new Walker(ctx);
          w2.walk(c2, {});
          w2.flush();
          sec2.questions.push({ index: (++gIndex), blocks: w2.blocks, answers: { answer: [], parse: [] } });
          model.stats.questions++;
          countMedia(w2.blocks, model.stats);
        }
        model.stats.types.push({ label: '', count: sec2.questions.length });
        model.sections.push(sec2);
      }
    }
    return model;
  }

  function countMedia(blocks, stats) {
    if (!blocks) return;
    for (var i = 0; i < blocks.length; i++) {
      var b = blocks[i];
      if (b.kind === 'para') {
        for (var j = 0; j < b.items.length; j++) {
          if (b.items[j].t === 'img') {
            stats.figures++;
            if (b.items[j].math) stats.mathImages++;
          }
        }
      } else if (b.kind === 'options') {
        for (var r = 0; r < b.rows.length; r++)
          for (var c = 0; c < b.rows[r].length; c++)
            countMedia([{ kind: 'para', items: b.rows[r][c] }], stats);
      } else if (b.kind === 'data') {
        for (var r2 = 0; r2 < b.rows.length; r2++)
          for (var c2 = 0; c2 < b.rows[r2].length; c2++)
            countMedia([{ kind: 'para', items: b.rows[r2][c2] }], stats);
      }
    }
  }

  global.P2D = {
    DEFAULT_OPTIONS: DEFAULT_OPTIONS,
    parsePaper: parsePaper,
    esc: esc,
    u8: u8,
    makeZip: makeZip,
    crc32: crc32,
    deflateRaw: deflateRaw,
    util: {
      readImageSize: readImageSize,
      svgToPng: svgToPng,
      svgIntrinsic: svgIntrinsic,
      shrinkRaster: shrinkRaster,
      mimeExt: mimeExt,
      isSvg: isSvg
    }
  };
})(typeof window !== 'undefined' ? window : globalThis);
