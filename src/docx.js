/* ===================================================================
 * Paper2Docx — DOCX 生成（OOXML）与打包
 * =================================================================== */
(function (global) {
  'use strict';
  var P2D = global.P2D;
  var esc = P2D.esc;
  var readImageSize = P2D.util.readImageSize;
  var svgToPng = P2D.util.svgToPng;
  var shrinkRaster = P2D.util.shrinkRaster;

  var W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

  function ptTwips(pt) { return Math.round(pt * 20); }
  function tw(cm) { return Math.round(cm * 567); }
  function emu(cm) { return Math.round(cm * 360000); }

  /* ---------- 图片准备 ---------- */
  async function prepareImages(model, files, opt, warnings) {
    var k = (opt.fontSize / 10.5) * opt.imgScale;     // 1 CSS px -> k pt
    var descs = new Map();
    function scan(blocks) {
      if (!blocks) return;
      blocks.forEach(function (b) {
        if (!b) return;
        if (b.kind === 'para') {
          b.items.forEach(function (it) { if (it.t === 'img' && !descs.has(it.key)) descs.set(it.key, it); });
        } else if (b.kind === 'options') {
          b.rows.forEach(function (r) { r.forEach(function (items) { scan([{ kind: 'para', items: items }]); }); });
        } else if (b.kind === 'data') {
          b.rows.forEach(function (r) { r.forEach(function (items) { scan([{ kind: 'para', items: items }]); }); });
        }
      });
    }
    model.sections.forEach(function (sec) {
      sec.questions.forEach(function (q) {
        scan(q.blocks);
        scan(q.answers && q.answers.answer);
        scan(q.answers && q.answers.parse);
      });
    });

    var media = [];
    var map = new Map();
    var missing = [];
    var used = new Set();

    for (var ent of descs) {
      var key = ent[0], d = ent[1];
      var file = files.get(key);
      if (!file) {
        var lower = key.toLowerCase();
        for (var fk of files.keys()) {
          if (fk.toLowerCase() === lower) { file = files.get(fk); break; }
        }
      }
      if (!file) { missing.push(key); continue; }

      var ext = 'png', bytesBlob = file, natW = 0, natH = 0;
      try {
        if (/\.svg$/i.test(key) || file.type === 'image/svg+xml') {
          var intr = await svgIntrinsicSafe(file);
          var cw = d.w || intr.w, chh = d.h || intr.h;
          if (!cw || !chh) { cw = 40; chh = 20; }
          var r = await svgToPng(file, cw, chh, 6);
          bytesBlob = r.blob; ext = 'png'; natW = cw; natH = chh;
        } else {
          var info = await readImageSize(file);
          natW = info.w; natH = info.h;
          var shr = await shrinkRaster(file, 3000);
          bytesBlob = shr.blob; ext = shr.ext;
          if (shr.natW) { natW = shr.natW; natH = shr.natH; }
          URL.revokeObjectURL(info.url);
        }
      } catch (e) {
        warnings.push('图片处理失败：' + key + '（' + e.message + '）');
        continue;
      }

      // 显示尺寸（CSS px）
      var cssW = d.w, cssH = d.h;
      if (!cssW || !cssH) {
        if (/\.svg$/i.test(key)) { cssW = cssW || natW; cssH = cssH || natH; }
        else { cssW = cssW || natW; cssH = cssH || natH; }
      }
      if (!cssH && cssW && natW) cssH = cssW * natH / natW;
      if (!cssW && cssH && natH) cssW = cssH * natW / natH;

      var wCm = cssW * 0.75 * k * 2.54 / 72;
      var hCm = cssH * 0.75 * k * 2.54 / 72;
      var buf = new Uint8Array(await bytesBlob.arrayBuffer());
      var name = 'image' + (media.length + 1) + '.' + ext;
      var rec = { key: key, name: name, data: buf, ext: ext, wCm: wCm, hCm: hCm, rId: 'rIdImg' + (media.length + 1) };
      media.push(rec);
      map.set(key, rec);
    }
    return { map: map, media: media, missing: missing };
  }

  function svgIntrinsicSafe(blob) {
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
    }).catch(function () { return { w: 0, h: 0 }; });
  }

  /* ---------- OOXML 片段 ---------- */
  function rPr(o) {
    var s = '<w:rPr>';
    var en = o.latinFont || 'Times New Roman', cn = o.bodyFont || '宋体';
    s += '<w:rFonts w:ascii="' + esc(en) + '" w:hAnsi="' + esc(en) + '" w:eastAsia="' + esc(cn) + '" w:cs="' + esc(en) + '"/>';
    if (o.b) s += '<w:b/><w:bCs/>';
    if (o.i) s += '<w:i/><w:iCs/>';
    var half = Math.round((o.size || 12) * 2);
    s += '<w:sz w:val="' + half + '"/><w:szCs w:val="' + half + '"/>';
    if (o.u) s += '<w:u w:val="single"/>';
    if (o.va === 'sub') s += '<w:vertAlign w:val="subscript"/>';
    else if (o.va === 'sup') s += '<w:vertAlign w:val="superscript"/>';
    return s + '</w:rPr>';
  }

  function pPr(o) {
    var s = '<w:pPr>';
    if (o.keepNext) s += '<w:keepNext/>';
    if (o.keepLines) s += '<w:keepLines/>';
    s += '<w:widowControl/>';
    var before = o.before || 0, after = o.after == null ? 0 : o.after;
    s += '<w:spacing w:before="' + ptTwips(before) + '" w:after="' + ptTwips(after) + '"' +
      ' w:line="' + Math.round((o.line || 1) * 240) + '" w:lineRule="auto"/>';
    if (o.left) s += '<w:ind w:left="' + tw(o.left) + '"/>';
    var jc = o.align === 'center' ? 'center' : (o.align === 'both' ? 'both' : 'left');
    s += '<w:jc w:val="' + jc + '"/>';
    return s + '</w:pPr>';
  }

  var imgSeq = 1;
  function imgXml(rec, wCm, hCm) {
    var cx = emu(wCm), cy = emu(hCm);
    var id = imgSeq++;
    return '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
      '<wp:extent cx="' + cx + '" cy="' + cy + '"/>' +
      '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
      '<wp:docPr id="' + id + '" name="Picture' + id + '"/>' +
      '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:nvPicPr><pic:cNvPr id="' + id + '" name="Picture' + id + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="' + rec.rId + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  }

  /* ---------- 文档构建 ---------- */
  async function buildDocx(htmlText, files, optIn, onProgress) {
    var opt = Object.assign({}, P2D.DEFAULT_OPTIONS, optIn || {});
    var warnings = [];
    var progress = onProgress || function () { };

    progress('解析网页结构…');
    var model = P2D.parsePaper(htmlText, opt);
    if (!model.stats.questions) {
      throw new Error('没有识别到题目。请确认上传的是组卷网「试卷预览/组卷中心」页面保存的 HTML。');
    }

    progress('处理图片（' + model.stats.figures + ' 处）…');
    var prep = await prepareImages(model, files, opt, warnings);
    if (prep.missing.length) {
      warnings.push('缺少 ' + prep.missing.length + ' 个图片文件：' + prep.missing.slice(0, 8).join('、') + (prep.missing.length > 8 ? ' …' : ''));
    }

    var textW = 21 - opt.margin * 2;
    var maxImgW = textW - 0.3;
    var body = [];
    var lastIsTable = false;
    function pushP(xml) { body.push(xml); lastIsTable = false; }
    function pushT(xml) {
      if (lastIsTable) body.push('<w:p>' + pPr({ before: 0, after: 0, line: 1.0 }) + '</w:p>');
      body.push(xml); lastIsTable = true;
    }

    function resolveItems(items, maxW) {
      var out = [];
      for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (it.t === 'img') {
          var rec = prep.map.get(it.key);
          if (!rec) continue;
          var w = rec.wCm, h = rec.hCm;
          var lim = maxW || maxImgW;
          if (w > lim) { h = h * lim / w; w = lim; }
          out.push({ t: 'img', rec: rec, w: w, h: h });
        } else out.push(it);
      }
      return out;
    }

    function itemsXml(items, o) {
      var s = '';
      var buf = '';
      function flush() {
        if (!buf) return;
        s += '<w:r>' + rPr(o) + '<w:t xml:space="preserve">' + esc(buf) + '</w:t></w:r>';
        buf = '';
      }
      for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (it.t === 'text') {
          if (it.b || it.i || it.u || it.va) {
            flush();
            s += '<w:r>' + rPr(Object.assign({}, o, { b: it.b, i: it.i, u: it.u, va: it.va })) +
              '<w:t xml:space="preserve">' + esc(it.text) + '</w:t></w:r>';
          } else buf += it.text;
        } else if (it.t === 'br') {
          flush(); s += '<w:r>' + rPr(o) + '<w:br/></w:r>';
        } else if (it.t === 'img') {
          flush(); s += imgXml(it.rec, it.w, it.h);
        }
      }
      flush();
      return s;
    }

    function paraXml(opts, items, o) {
      return '<w:p>' + pPr(opts) + itemsXml(items, o) + '</w:p>';
    }

    function hasImg(items) { return items.some(function (x) { return x.t === 'img'; }); }
    function onlyImg(items) {
      var rest = items.filter(function (x) { return !(x.t === 'text' && !x.text.trim()) && x.t !== 'br'; });
      return rest.length > 0 && rest.every(function (x) { return x.t === 'img'; });
    }

    function tableXml(rows, cols, widthsCm, opts) {
      var total = widthsCm.reduce(function (a, b) { return a + b; }, 0);
      var s = '<w:tbl><w:tblPr>';
      s += '<w:tblW w:w="' + tw(total) + '" w:type="dxa"/>';
      s += '<w:jc w:val="' + (opts.center ? 'center' : 'left') + '"/>';
      if (opts.border) {
        s += '<w:tblBorders>';
        ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].forEach(function (sd) {
          s += '<w:' + sd + ' w:val="single" w:sz="6" w:space="0" w:color="000000"/>';
        });
        s += '</w:tblBorders>';
      } else {
        s += '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(function (sd) {
          return '<w:' + sd + ' w:val="none" w:sz="0" w:space="0" w:color="auto"/>';
        }).join('') + '</w:tblBorders>';
      }
      s += '<w:tblLayout w:type="fixed"/>';
      s += '<w:tblCellMar>' + ['top', 'left', 'bottom', 'right'].map(function (sd) {
        return '<w:' + sd + ' w:w="' + (opts.border ? 57 : 0) + '" w:type="dxa"/>';
      }).join('') + '</w:tblCellMar>';
      s += '</w:tblPr><w:tblGrid>';
      widthsCm.forEach(function (w) { s += '<w:gridCol w:w="' + tw(w) + '"/>'; });
      s += '</w:tblGrid>';

      for (var r = 0; r < rows.length; r++) {
        s += '<w:tr><w:trPr><w:cantSplit/></w:trPr>';
        for (var c = 0; c < cols; c++) {
          var cellW = widthsCm[c] || (total / cols);
          s += '<w:tc><w:tcPr><w:tcW w:w="' + tw(cellW) + '" w:type="dxa"/>' +
            '<w:vAlign w:val="' + (opts.vAlign || 'center') + '"/></w:tcPr>';
          s += opts.cellXml(r, c, cellW);
          s += '</w:tc>';
        }
        s += '</w:tr>';
      }
      return s + '</w:tbl>';
    }

    /* 题头 */
    var titleSize = Math.max(16, opt.fontSize + 6);
    pushP(paraXml({ align: 'center', before: 0, after: 10, line: 1.0 },
      [{ t: 'text', text: model.title || '试卷' }],
      { size: titleSize, b: true, bodyFont: opt.titleFont, latinFont: opt.titleFont }));
    if (model.std) {
      pushP(paraXml({ align: 'center', before: 0, after: 10, line: 1.0 },
        [{ t: 'text', text: model.std }], { size: Math.max(10, opt.fontSize - 1.5) }));
    }

    var lastPart = '';
    for (var si = 0; si < model.sections.length; si++) {
      var sec = model.sections[si];
      if (sec.partTitle && sec.partTitle !== lastPart) {
        lastPart = sec.partTitle;
        pushP(paraXml({ align: 'center', before: 12, after: 4, line: 1.0 },
          [{ t: 'text', text: sec.partTitle }],
          { size: opt.fontSize + 2, b: true, bodyFont: opt.titleFont, latinFont: opt.titleFont }));
      }
      if (sec.label) {
        pushP(paraXml({ align: null, before: 12, after: 4, line: 1.0, keepNext: true },
          [{ t: 'text', text: sec.label }],
          { size: opt.fontSize + 2, b: true, bodyFont: opt.titleFont, latinFont: opt.titleFont }));
      }

      for (var qi = 0; qi < sec.questions.length; qi++) {
        var q = sec.questions[qi];
        var blocks = q.blocks;
        // 若首块没有题号，补一个
        if (blocks.length && blocks[0].kind === 'para' && blocks[0].items.length) {
          var firstTxt = blocks[0].items[0];
          if (firstTxt.t === 'text' && !/^\s*\d/.test(firstTxt.text)) {
            blocks[0].items.unshift({ t: 'text', text: (q.index) + '．' });
          }
        }
        for (var bi = 0; bi < blocks.length; bi++) {
          var b = blocks[bi];
          var nxt = blocks[bi + 1];
          var nxtImg = nxt && nxt.kind === 'para' && onlyImg(resolveItems(nxt.items));
          var nxtTbl = nxt && (nxt.kind === 'options' || nxt.kind === 'data');
          var o = { size: opt.fontSize };
          if (b.kind === 'para') {
            var items = resolveItems(b.items);
            var isImg = onlyImg(items);
            var opts2 = {
              align: isImg ? 'center' : (b.align === 'center' ? 'center' : null),
              before: bi === 0 ? opt.questionGap : (isImg ? 1 : 0),
              after: nxtTbl ? 4 : 2,
              line: opt.lineSpacing,
              keepNext: (bi === 0) || (opt.keepImageWithText && (nxtImg || nxtTbl)),
              keepLines: isImg && opt.keepImageWithText
            };
            pushP(paraXml(opts2, items, o));
          } else if (b.kind === 'options') {
            pushT(optionsXml(b, resolveItems, tableXml, opt, textW, bi === 0, o));
          } else if (b.kind === 'data') {
            pushT(dataXml(b, resolveItems, tableXml, opt, textW, o));
          }
        }

        // 答案 / 解析
        if (opt.includeAnswers) {
          var ansBlocks = (q.answers && q.answers.answer) || [];
          var parseBlocks = (q.answers && q.answers.parse) || [];
          if (ansBlocks.length) pushP(ansXml('【答案】', ansBlocks, resolveItems, paraXml, opt, maxImgW));
          if (parseBlocks.length) pushP(ansXml('【解析】', parseBlocks, resolveItems, paraXml, opt, maxImgW));
        }
      }
    }

    function optionsXml(b, resolveItems, tableXml, opt, textW, first, o) {
      var cols = opt.autoOptionColumns ? b.cols : 1;
      var ncol = 1;
      b.rows.forEach(function (r) { ncol = Math.max(ncol, r.length); });
      var out = [];
      if (cols <= 1) {
        for (var r = 0; r < b.rows.length; r++) {
          for (var c = 0; c < b.rows[r].length; c++) {
            var items = resolveItems(b.rows[r][c], maxImgW);
            out.push(paraXml({
              before: (r === 0 && c === 0 && first) ? 2 : 0, after: 1, line: opt.lineSpacing, left: 0.35,
              keepNext: false
            }, items, o));
          }
        }
        return out.join('');
      }
      var n = Math.min(cols, ncol);
      var cw = textW / n;
      var widths = new Array(n).fill(cw);
      var cellXml = function (r, c, cellW) {
        var items = resolveItems(b.rows[r] && b.rows[r][c] ? b.rows[r][c] : [], cellW - 0.35);
        var hasI = hasImg(items);
        if (hasI) {
          // 题号与图分两行居中
          var label = '', rest = items;
          for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (it.t === 'text') {
              var t = it.text.trim();
              if (/^[A-D][．.、]?$/.test(t) || /^[A-D][．.、]/.test(t)) { label = t; rest = items.slice(i + 1); break; }
            } else break;
          }
          var s = '<w:p>' + pPr({ align: 'center', before: 0, after: 0, line: 1.0 }) +
            (label ? '<w:r>' + rPr(o) + '<w:t xml:space="preserve">' + esc(label) + '</w:t></w:r>' : '') + '</w:p>';
          s += '<w:p>' + pPr({ align: 'center', before: 0, after: 0, line: 1.0 }) + itemsXml(rest, o) + '</w:p>';
          return s;
        }
        return '<w:p>' + pPr({ align: 'left', before: 0, after: 0, line: opt.lineSpacing }) + itemsXml(items, o) + '</w:p>';
      };
      return tableXml(b.rows, n, widths, { border: false, center: false, vAlign: 'top', cellXml: cellXml });
    }

    function dataXml(b, resolveItems, tableXml, opt, textW, o) {
      var ncols = 1;
      b.rows.forEach(function (r) { ncols = Math.max(ncols, r.length); });
      var nat = (b.tw || 400) * 0.75 * (opt.fontSize / 10.5) * opt.imgScale * 2.54 / 72;
      var tblW = Math.min(textW, nat);
      var widths = [];
      var known = b.widths.filter(function (x) { return x; });
      var sum = known.reduce(function (a, x) { return a + x; }, 0) || 1;
      for (var i = 0; i < ncols; i++) {
        var v = b.widths[i];
        widths.push(((v || sum / ncols) / sum) * tblW);
      }
      var cellXml = function (r, c, cellW) {
        var items = resolveItems(b.rows[r] && b.rows[r][c] ? b.rows[r][c] : [], Math.max(0.5, cellW - 0.3));
        return '<w:p>' + pPr({ align: 'center', before: 1, after: 1, line: 1.0 }) + itemsXml(items, o) + '</w:p>';
      };
      return tableXml(b.rows, ncols, widths, { border: true, center: true, vAlign: 'center', cellXml: cellXml });
    }

    function ansXml(prefix, blocks, resolveItems, paraXml, opt, maxW) {
      var s = '';
      var o = { size: opt.fontSize - 1 };
      for (var i = 0; i < blocks.length; i++) {
        var bl = blocks[i];
        if (bl.kind !== 'para') continue;
        var items = resolveItems(bl.items, maxW);
        if (i === 0) items.unshift({ t: 'text', text: prefix });
        s += paraXml({ before: 0, after: 0, line: opt.lineSpacing, left: 0.35 }, items, o);
      }
      return s;
    }

    progress('生成文档…');
    var sectPr = '<w:sectPr><w:footerReference w:type="default" r:id="rIdFtr"/>' +
      '<w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="' + tw(opt.margin) + '" w:right="' + tw(opt.margin) + '" w:bottom="' + tw(opt.margin - 0.2) +
      '" w:left="' + tw(opt.margin) + '" w:header="' + tw(1.2) + '" w:footer="' + tw(1.0) + '" w:gutter="0"/>' +
      '<w:cols w:space="425"/></w:sectPr>';

    var docXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:document ' + W_NS + '><w:body>' + body.join('') + sectPr + '</w:body></w:document>';

    var stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr>' +
      '<w:rFonts w:ascii="' + esc(opt.latinFont) + '" w:hAnsi="' + esc(opt.latinFont) + '" w:eastAsia="' + esc(opt.bodyFont) + '" w:cs="' + esc(opt.latinFont) + '"/>' +
      '<w:sz w:val="' + Math.round(opt.fontSize * 2) + '"/><w:szCs w:val="' + Math.round(opt.fontSize * 2) + '"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:line="' + Math.round(opt.lineSpacing * 240) + '" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
      '</w:styles>';

    var settingsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:zoom w:percent="100"/><w:defaultTabStop w:val="420"/><w:characterSpacingControl w:val="compressPunctuation"/>' +
      '<w:compat><w:useFELayout/></w:compat></w:settings>';

    var footerXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
      '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">第 </w:t></w:r>' +
      '<w:fldSimple w:instr=" PAGE "><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>' +
      '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve"> 页 / 共 </w:t></w:r>' +
      '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>' +
      '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve"> 页</w:t></w:r></w:p></w:ftr>';

    var mediaList = prep.media;
    var ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>';
    var exts = {};
    mediaList.forEach(function (m) { exts[m.ext] = true; });
    Object.keys(exts).forEach(function (e) {
      var ctType = e === 'jpeg' || e === 'jpg' ? 'image/jpeg' : ('image/' + (e === 'svg' ? 'svg+xml' : e));
      ct += '<Default Extension="' + e + '" ContentType="' + ctType + '"/>';
    });
    ct += '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
      '</Types>';

    var rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
      '</Relationships>';

    var docRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>' +
      '<Relationship Id="rIdFtr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>';
    mediaList.forEach(function (m) {
      docRels += '<Relationship Id="' + m.rId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' + m.name + '"/>';
    });
    docRels += '</Relationships>';

    var now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    var coreXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
      'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:title>' + esc(model.title || '试卷') + '</dc:title><dc:creator>组卷转Word工具</dc:creator>' +
      '<cp:lastModifiedBy>组卷转Word工具</cp:lastModifiedBy>' +
      '<dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created>' +
      '<dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified>' +
      '</cp:coreProperties>';

    var appXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
      'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      '<Application>组卷转Word工具</Application></Properties>';

    progress('打包 .docx…');
    var entries = [
      { name: '[Content_Types].xml', data: P2D.u8(ct) },
      { name: '_rels/.rels', data: P2D.u8(rootRels) },
      { name: 'docProps/core.xml', data: P2D.u8(coreXml) },
      { name: 'docProps/app.xml', data: P2D.u8(appXml) },
      { name: 'word/document.xml', data: P2D.u8(docXml) },
      { name: 'word/styles.xml', data: P2D.u8(stylesXml) },
      { name: 'word/settings.xml', data: P2D.u8(settingsXml) },
      { name: 'word/footer1.xml', data: P2D.u8(footerXml) },
      { name: 'word/_rels/document.xml.rels', data: P2D.u8(docRels) }
    ];
    mediaList.forEach(function (m) {
      entries.push({ name: 'word/media/' + m.name, data: m.data, store: true });
    });

    var blob = await P2D.makeZip(entries);
    return {
      blob: blob,
      stats: model.stats,
      title: model.title,
      warnings: warnings,
      missing: prep.missing,
      images: mediaList.length
    };
  }

  P2D.buildDocx = buildDocx;

})(typeof window !== 'undefined' ? window : globalThis);

