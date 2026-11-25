/* ===================================================================
 * Paper2Docx UI
 * =================================================================== */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };
  var state = { htmlFile: null, htmlText: '', htmlName: '', files: new Map(), allFiles: [], model: null };

  var $htmlInput = $('htmlInput'), $dirInput = $('dirInput');
  var $htmlZone = $('htmlZone'), $dirZone = $('dirZone');
  var $htmlInfo = $('htmlInfo'), $dirInfo = $('dirInfo');
  var $btn = $('convertBtn');
  var $status = $('status');
  var $result = $('result');
  var $preview = $('preview');
  var $dl = $('downloadLink');

  var IMG_RE = /\.(png|jpe?g|gif|svg|webp|bmp)$/i;
  var HTML_RE = /\.html?$/i;

  /* ---------- 提示 ---------- */
  function setStatus(msg, type) {
    if (!msg) { $status.className = 'status'; $status.textContent = ''; return; }
    $status.className = 'status ' + (type || 'info');
    $status.textContent = msg;
  }

  /* ---------- 文件收集 ---------- */
  function addFiles(fileList) {
    var imgs = 0, htmls = 0;
    for (var i = 0; i < fileList.length; i++) {
      var f = fileList[i];
      var name = f.name;
      if (HTML_RE.test(name)) {
        htmls++;
        if (!state.htmlFile || f.size > state.htmlFile.size) state.htmlFile = f;
      } else if (IMG_RE.test(name)) {
        imgs++;
        state.files.set(name, f);
        state.allFiles.push(f);
      } else {
        state.allFiles.push(f);
      }
    }
    return { imgs: imgs, htmls: htmls };
  }

  function readEntries(dirReader) {
    return new Promise(function (resolve) {
      var out = [];
      function step() {
        dirReader.readEntries(function (entries) {
          if (!entries.length) return resolve(out);
          out = out.concat(Array.prototype.slice.call(entries));
          step();
        }, function () { resolve(out); });
      }
      step();
    });
  }

  function entryFile(entry) {
    return new Promise(function (res) {
      entry.file(function (f) { res(f); }, function () { res(null); });
    });
  }

  async function walkEntry(entry, acc) {
    if (entry.isFile) {
      var f = await entryFile(entry);
      if (f) acc.push(f);
    } else if (entry.isDirectory) {
      var reader = entry.createReader();
      var entries = await readEntries(reader);
      for (var i = 0; i < entries.length; i++) await walkEntry(entries[i], acc);
    }
  }

  async function handleDrop(e) {
    e.preventDefault();
    $htmlZone.classList.remove('active'); $dirZone.classList.remove('active');
    var dt = e.dataTransfer;
    var acc = [];
    if (dt.items && dt.items.length && dt.items[0].webkitGetAsEntry) {
      var entries = [];
      for (var i = 0; i < dt.items.length; i++) {
        var en = dt.items[i].webkitGetAsEntry();
        if (en) entries.push(en);
      }
      for (var j = 0; j < entries.length; j++) await walkEntry(entries[j], acc);
    } else {
      acc = Array.prototype.slice.call(dt.files);
    }
    setStatus('正在读取拖入的文件…');
    addFiles(acc);
    await afterFiles('已拖入 ' + acc.length + ' 个文件');
  }

  async function afterFiles(msg) {
    if (state.htmlFile) {
      state.htmlText = await state.htmlFile.text();
      state.htmlName = state.htmlFile.name;
      $htmlInfo.textContent = state.htmlFile.name + '（' + (state.htmlFile.size / 1024).toFixed(0) + ' KB）';
      $htmlInfo.className = 'file-info ok';
    }
    var imgCount = 0;
    state.files.forEach(function () { imgCount++; });
    $dirInfo.textContent = imgCount ? ('已载入 ' + imgCount + ' 个图片/资源文件') : '尚未选择素材文件夹';
    $dirInfo.className = 'file-info' + (imgCount ? ' ok' : '');

    refreshBtn();
    if (state.htmlText) await doPreview(msg);
    else if (msg) setStatus(msg + '。还需要选择试卷 HTML 文件。', 'info');
  }

  function refreshBtn() {
    $btn.disabled = !(state.htmlText && state.files.size);
  }

  /* ---------- 解析预览 ---------- */
  async function doPreview(msg) {
    try {
      var model = P2D.parsePaper(state.htmlText, readOptions());
      state.model = model;
      var st = model.stats;
      var miss = 0, missList = [];
      var seen = new Set();
      (function scan(blocks) {
        if (!blocks) return;
        blocks.forEach(function (b) {
          if (!b) return;
          if (b.kind === 'para') b.items.forEach(function (it) {
            if (it.t !== 'img' || seen.has(it.key)) return;
            seen.add(it.key);
            if (!state.files.get(it.key)) { miss++; missList.push(it.key); }
          });
          else if (b.kind === 'options' || b.kind === 'data') b.rows.forEach(function (r) {
            r.forEach(function (items) { scan([{ kind: 'para', items: items }]); });
          });
        });
      })(flattenBlocks(model));

      var types = st.types.filter(function (t) { return t.count; })
        .map(function (t) { return (t.label || '未命名题型') + ' ' + t.count + ' 题'; }).join('、');

      $preview.innerHTML = '';
      $preview.className = 'preview show';
      var lines = [
        ['试卷标题', model.title || '（未识别）'],
        ['题目数量', st.questions + ' 题（' + st.types.length + ' 个题型）'],
        ['题型分布', types || '—'],
        ['插图数量', st.figures + ' 处（其中公式小图 ' + st.mathImages + ' 处）'],
        ['图片资源', miss ? ('<b class="warn">缺 ' + miss + ' 个</b>：' + missList.slice(0, 6).join('、')) : '全部匹配 ✓']
      ];
      lines.forEach(function (kv) {
        var row = document.createElement('div');
        row.className = 'pv-row';
        row.innerHTML = '<span class="pv-k">' + kv[0] + '</span><span class="pv-v">' + kv[1] + '</span>';
        $preview.appendChild(row);
      });
      setStatus(msg || '识别完成，可以生成 Word 了。', 'success');
    } catch (e) {
      setStatus('解析失败：' + e.message, 'error');
    }
  }

  function flattenBlocks(model) {
    var out = [];
    model.sections.forEach(function (s) {
      s.questions.forEach(function (q) {
        out = out.concat(q.blocks);
        if (q.answers) {
          out = out.concat(q.answers.answer || []);
          out = out.concat(q.answers.parse || []);
        }
      });
    });
    return out;
  }

  /* ---------- 参数 ---------- */
  var LAYOUT_HINT = {
    'a4-1': 'A4 单栏：一张 A4 一页，普通打印机直接出。',
    'a3-2': '<b>A3 横向双栏</b>：每张 A3 左右两栏（约等于两张 A4 并排），正反印共 4 面，最接近正式考试卷；卷头自动横跨两栏。要装订成册请用 Word「书籍折页」打印或让文印店拼版。',
    '8k-2': '<b>8 开横向双栏</b>（420×285mm）：国内试卷常用开本，家用打印机一般不支持，适合送文印店。',
    'a4-2': '<b>A4 双栏</b>：左右分栏、题目自动流到下一栏；图会缩到栏宽（约 8cm），适合题量少的卷子。',
    '16k-1': '16 开单栏（195×270mm）：比 A4 略窄，用 16 开纸打印时选它。'
  };

  function syncLayoutHint() {
    var el = $('layoutHint');
    if (!el) return;
    el.innerHTML = LAYOUT_HINT[$('optLayout').value] || '';
  }

  function readOptions() {
    return {
      layout: $('optLayout').value,
      fontSize: parseFloat($('optFont').value),
      lineSpacing: parseFloat($('optLine').value),
      imgScale: parseFloat($('optImg').value),
      margin: parseFloat($('optMargin').value),
      includeAnswers: $('optAns').value === '1',
      questionGap: parseFloat($('optGap').value)
    };
  }

  /* ---------- 生成 ---------- */
  function safeName(s) {
    return (s || '试卷').replace(/[\\/:*?"<>|\r\n\t]/g, '').trim().slice(0, 60) || '试卷';
  }

  function download(blob, name) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 4000);
  }

  async function convert() {
    if ($btn.disabled) return;
    $btn.disabled = true;
    $result.className = 'result';
    setStatus('开始处理…');
    try {
      var t0 = performance.now();
      var res = await P2D.buildDocx(state.htmlText, state.files, readOptions(), function (m) {
        setStatus(m);
      });
      var name = safeName(res.title) + '.docx';
      download(res.blob, name);
      var secs = ((performance.now() - t0) / 1000).toFixed(1);
      $result.className = 'result show';
      $result.innerHTML =
        '<div class="ok-title">✓ 已生成：' + name + '</div>' +
        '<div class="ok-meta">' + res.stats.questions + ' 题 · ' + res.stats.figures + ' 处插图 · ' +
        (res.blob.size / 1048576).toFixed(1) + ' MB · 用时 ' + secs + ' 秒</div>' +
        (res.warnings.length ? '<div class="warn-box">' + res.warnings.map(function (w) { return '· ' + w; }).join('<br>') + '</div>' : '') +
        '<div class="hint">如果浏览器没有自动下载，请检查下载拦截提示。</div>';
      setStatus('完成。', 'success');
    } catch (e) {
      setStatus('生成失败：' + e.message, 'error');
      console.error(e);
    } finally {
      $btn.disabled = false;
    }
  }

  /* ---------- 事件绑定 ---------- */
  $htmlZone.addEventListener('click', function () { $htmlInput.click(); });
  $dirZone.addEventListener('click', function () { $dirInput.click(); });

  $htmlInput.addEventListener('change', function () {
    var fs = $htmlInput.files;
    if (!fs.length) return;
    state.htmlFile = fs[0];
    afterFiles();
  });

  $dirInput.addEventListener('change', function () {
    var fs = $dirInput.files;
    if (!fs.length) return;
    addFiles(fs);
    afterFiles();
  });

  ['dragenter', 'dragover'].forEach(function (ev) {
    document.addEventListener(ev, function (e) {
      e.preventDefault();
      document.body.classList.add('dragging');
    });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    document.addEventListener(ev, function (e) {
      if (ev === 'dragleave' && e.relatedTarget) return;
      document.body.classList.remove('dragging');
    });
  });
  document.addEventListener('drop', handleDrop);

  $btn.addEventListener('click', convert);

  ['optLayout', 'optFont', 'optLine', 'optImg', 'optMargin', 'optAns', 'optGap'].forEach(function (id) {
    $(id).addEventListener('change', function () {
      if (state.htmlText) doPreview();
    });
  });

  syncLayoutHint();
  $('optLayout').addEventListener('change', syncLayoutHint);

  // 默认展开说明
  var tips = $('tips');
  if (tips) tips.open = false;
})();
