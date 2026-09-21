/*
 * core.js — 账单/发货单 通用合并核心逻辑（浏览器 + Node 通用 UMD）
 * 自适应跨格式：自动识别表头行、按同义词映射列、跳过主表/合计行。
 */
(function (global) {
  'use strict';
  var root = (typeof globalThis !== 'undefined') ? globalThis : global;
  var XLSX = root.XLSX;
  var VERSION = '1.4.1';            // 工具版本号：每次更新必须递增（唯一来源，见 CHANGELOG.md）
  var BUILD_DATE = '2026-09-21';    // 本版本日期

  /* ---------------- 文本工具 ---------------- */

  // 统一为可比较的“干净”文本：去换行/空白/全角空格，转小写(ASCII)
  function clean(v) {
    if (v === null || v === undefined) return '';
    var s = String(v);
    s = s.replace(/[\r\n\t ]/g, '');
    s = s.replace(/\u3000/g, '');
    return s;
  }

  // 展示用文本（保留空格换行但折叠空白）
  function display(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return fmtDate(v);
    return String(v).replace(/\s+/g, ' ');
  }

  // 检测文本是否为“表头类”内容（含标点连续语段不算）
  function isHeaderish(h) {
    if (!h) return false;
    var len = h.length;
    if (len > 40) return false;              // 长句说明文字不像列名
    return true;
  }

  // 拼接一行前 nCols 列的非空文本（用空格分隔），用于识别出库总量等校核行
  function rowText(ws, r, nCols) {
    var parts = [];
    for (var c = 0; c < nCols; c++) {
      var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
      if (cell && cell.v !== null && cell.v !== undefined && String(cell.v).trim() !== '') {
        parts.push(String(cell.v).trim());
      }
    }
    return parts.join(' ');
  }

  // 从文本中按“标签：数值”格式提取数字，取不到返回 null
  function pickNum(s, re) {
    var m = String(s).match(re);
    return m ? parseFloat(m[1]) : null;
  }

  // 读取单元格数值（非数字返回 null）
  function cellNum(ws, r, c) {
    var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
    if (!cell || cell.v === null || cell.v === undefined) return null;
    var v = cell.v;
    if (typeof v === 'number') return v;
    var n = parseFloat(String(v).replace(/[^\-0-9.]/g, ''));
    return isNaN(n) ? null : n;
  }

  function round3(n) { return Math.round(n * 1000) / 1000; }

  /*
   * 「单块体积/面积」这类合并列的语义判定：按源表头文本判断这一列装的是体积还是面积。
   * 台账/汇总据此把数值放进正确的一格（体积列或面积列），另一格缺失时再由尺寸推算。
   * 返回 'volume' | 'area' | ''
   */
  function qtySemantics(text) {
    var t = clean(text);
    if (!t) return '';
    if (/面积|平方|m2|m²/i.test(t)) return 'area';
    if (/体积|方量|立方|m3|m³/i.test(t)) return 'volume';
    return '';
  }

  /*
   * 表头区（表头行之上）的“出库总量”标注。有些厂的出库单把总量印在页面顶部：
   *   A4='出库总量' | D4='体积' | E4='57.716m³'   → 标签与数值分列
   *   或 总块数：36块 / 总方量：6.806m³           → 同一格“标签：数值”
   * 返回 {blocks, volume, area, weight, text} 或 null（解析不到值时返回 null，避免误报）
   */
  function findHeaderTotal(ws, headerRow) {
    if (!headerRow || headerRow < 2) return null;
    var rmax = Math.min(headerRow - 1, 12);
    for (var r = 0; r < rmax; r++) {
      var line = rowText(ws, r, 32);
      if (!/出库总量|总块数|总数量|总方量|总体积|总面积|总重量/.test(line)) continue;
      var eB = pickNum(line, /总块数[：:]?\s*([0-9]+(?:\.[0-9]+)?)/);
      var eV = pickNum(line, /总方量[：:]?\s*([0-9]+(?:\.[0-9]+)?)/);
      if (eV === null) eV = pickNum(line, /总体积[：:]?\s*([0-9]+(?:\.[0-9]+)?)/);
      var eA = pickNum(line, /总面积[：:]?\s*([0-9]+(?:\.[0-9]+)?)/);
      var eW = pickNum(line, /总重量[：:]?\s*([0-9]+(?:\.[0-9]+)?)/);
      if (eB === null && eV === null && eA === null && eW === null) {
        var cols = [];
        for (var c = 0; c < 32; c++) {
          var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
          cols.push(cell && cell.v !== null && cell.v !== undefined ? String(cell.v).trim() : '');
        }
        for (var i = 0; i < cols.length; i++) {
          var lab = clean(cols[i]);
          if (!lab) continue;
          var kind = /块数|数量|件数/.test(lab) ? 'b'
                   : /方量|体积/.test(lab) ? 'v'
                   : /面积/.test(lab) ? 'a'
                   : /重量|质量/.test(lab) ? 'w' : '';
          if (!kind) continue;
          for (var j = i + 1; j < cols.length && j <= i + 3; j++) {
            var mm = String(cols[j]).match(/([0-9]+(?:\.[0-9]+)?)/);
            if (!mm) continue;
            var val = parseFloat(mm[1]);
            if (kind === 'b' && eB === null) eB = val;
            else if (kind === 'v' && eV === null) eV = val;
            else if (kind === 'a' && eA === null) eA = val;
            else if (kind === 'w' && eW === null) eW = val;
            break;
          }
        }
      }
      if (eB !== null || eV !== null || eA !== null || eW !== null) {
        return { blocks: eB, volume: eV, area: eA, weight: eW, text: line };
      }
    }
    return null;
  }

  // 比对“出库总量”标注值与明细合计，返回校核记录
  function buildVerify(source, exp, act) {
    function cmp(label, e, a, unit, tol) {
      if (e === null || e === undefined) return null;
      var ok = Math.abs(e - a) <= tol;
      return {
        label: label, exp: e, act: round3(a), unit: unit || '', ok: ok,
        text: label + ' 标注 ' + e + (unit || '') + ' ／ 明细合计 ' + round3(a) + (unit || '') + (ok ? ' ✔' : ' ✘ 不符')
      };
    }
    var items = [];
    var t;
    t = cmp('块数', exp.blocks, act.count, '块', 0.001); if (t) items.push(t);
    t = cmp('方量', exp.volume, act.volume, 'm³', Math.max(0.02, (exp.volume || 0) * 0.005)); if (t) items.push(t);
    t = cmp('面积', exp.area, act.area, 'm²', Math.max(0.05, (exp.area || 0) * 0.005)); if (t) items.push(t);
    t = cmp('重量', exp.weight, act.weight, 'T', Math.max(0.02, (exp.weight || 0) * 0.005)); if (t) items.push(t);
    return {
      source: source,
      ok: items.length > 0 && items.every(function (x) { return x.ok; }),
      items: items,
      text: items.map(function (x) { return x.text; }).join('；')
    };
  }

  /* ---------------- 日期工具 ---------------- */

  function pad(n) { return n < 10 ? '0' + n : String(n); }

  function fmtDate(d) {
    if (!(d instanceof Date) || isNaN(d.getTime())) return '';
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  }

  // 从任意值里尽量抠出一个 yyyy-mm-dd 文本
  function toDateText(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return fmtDate(v);
    var s = String(v).trim();
    if (!s) return '';
    // 纯数字 Excel 序列号
    if (/^\d+(\.\d+)?$/.test(s) && XLSX && XLSX.SSF) {
      try {
        var num = parseFloat(s);
        if (num > 20000 && num < 80000) {
          var d = XLSX.SSF.parse_date_code(num);
          if (d) return d.y + '-' + pad(d.m) + '-' + pad(d.d);
        }
      } catch (e) { /* ignore */ }
    }
    // 中文/横线/斜杠日期
    var m = s.match(/(\d{4})[年\/\-.](\d{1,2})[月\/\-.](\d{1,2})/);
    if (m) return m[1] + '-' + pad(+m[2]) + '-' + pad(+m[3]);
    m = s.match(/(\d{4})(\d{2})(\d{2})/);
    if (m && m[1] > '1900') return m[1] + '-' + m[2] + '-' + m[3];
    return '';
  }

  /* ---------------- 读取工作簿 ---------------- */

  // data: ArrayBuffer 或 Uint8Array
  function loadWorkbook(data) {
    var wb = XLSX.read(data, { type: 'array', cellDates: true });
    var out = { wb: wb, sheets: [] };
    wb.SheetNames.forEach(function (name) {
      var ws = wb.Sheets[name];
      out.sheets.push({ name: name, ws: ws });
    });
    return out;
  }

  // 取工作表前 maxRows 行的预览网格（文本）
  function previewGrid(ws, maxRows, maxCols) {
    maxRows = maxRows || 14; maxCols = maxCols || 30;
    var rows = [];
    var ref = ws['!ref'];
    if (!ref) return rows;
    var lastR = Math.min(XLSX.utils.decode_range(ref).e.r + 1, maxRows);
    for (var r = 0; r < lastR; r++) {
      var line = [];
      for (var c = 0; c < maxCols; c++) {
        var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
        line.push(cell ? display(cell.v) : '');
      }
      rows.push(line);
    }
    return rows;
  }

  /* ---------------- 表头行自动识别 ---------------- */

  // 每个单元格命中一次记一次分的强词表（表头列的典型词）
  var STRONG_WORDS = [
    '构件编号', '编号', '构件', '楼栋', '楼层', '类型', '型号', '规格',
    '体积', '面积', '质量', '重量', '砼', '标号', '备注', '序号',
    '数量', '日期', '长度', '宽度', '厚度', '名称', '编码', '强度',
    '尺寸', '单价', '金额', '位号', '部位', '层号', '长', '宽', '厚'
  ];

  // 只统计“干净短文本”单元格，避免把整段提示/签名/批注算进去
  function rowHeaderScore(ws, r1) {
    var score = 0, nonEmpty = 0;
    for (var c = 0; c < 60; c++) {
      var cell = ws[XLSX.utils.encode_cell({ r: r1, c: c })];
      if (!cell || cell.v === null || cell.v === undefined) continue;
      var t = clean(cell.v);
      if (!t) continue;
      nonEmpty++;
      if (t.length > 40) continue;                    // 长句不算
      var hit = 0;
      for (var i = 0; i < STRONG_WORDS.length; i++) {
        if (t.indexOf(STRONG_WORDS[i]) >= 0) hit++;
      }
      if (hit > 0) score += 1 + Math.min(hit, 3);     // 命中词越多分越高，单格上限4
    }
    // 行内有内容才算候选
    if (nonEmpty === 0) return 0;
    score += Math.min(nonEmpty, 12);                  // 分散多列的整行表头分更高
    return score;
  }

  // 找出得分最高的“表头行”（1-based），找不到返回 0
  function detectHeaderRow(ws) {
    var ref = ws['!ref'];
    if (!ref) return 0;
    var maxR = Math.min(XLSX.utils.decode_range(ref).e.r + 1, 40);
    var best = 0, bestScore = 0;
    for (var r = 0; r < maxR; r++) {
      var s = rowHeaderScore(ws, r);
      if (s > bestScore) { bestScore = s; best = r + 1; }
    }
    return bestScore >= 6 ? best : 0;
  }

  /* ---------------- 发货单 vs 主表/普通表 ---------------- */

  var DELIVERY_MARKERS = [
    '项目名称', '客户姓名', '客户', '运输车号', '运输人', '运输公司',
    '出库', '收货', '发货单', '送货', '司机', '承运', '打印时间'
  ];

  // 表头上方(前若干行)是否有单据头信息
  function hasDeliveryMarkers(ws, headerRow) {
    var upto = headerRow > 0 ? Math.min(headerRow - 1, 9) : 8;
    var ref = ws['!ref'];
    if (!ref) return false;
    var rmax = Math.min(XLSX.utils.decode_range(ref).e.r + 1, upto);
    for (var r = 0; r < rmax; r++) {
      for (var c = 0; c < 40; c++) {
        var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
        if (!cell) continue;
        var t = clean(cell.v);
        if (!t) continue;
        for (var i = 0; i < DELIVERY_MARKERS.length; i++) {
          if (t.indexOf(DELIVERY_MARKERS[i]) >= 0) return true;
        }
      }
    }
    return false;
  }

  /* ---------------- 列同义词映射 ---------------- */

  // 目标列名 -> 匹配规则(clean 后文本)。返回 true 即命中。
  // 注意顺序即优先级，一个源列只被一个目标列消费。
  var SYNONYMS = {
    '序号': function (h) { return h === '序号' || h === '序' || /^序号/.test(h); },
    '构件编号': function (h) {
      return h.indexOf('构件编号') >= 0 ||
             (h.indexOf('构件') >= 0 && h.indexOf('编号') >= 0) ||
             h.indexOf('编码') >= 0 ||
             (h.indexOf('编号') >= 0 && h.indexOf('序号') < 0) ||
             h === '编号';
    },
    '楼栋': function (h) {
      return h.indexOf('楼栋') >= 0 || h.indexOf('楼号') >= 0 ||
             h.indexOf('栋号') >= 0 || /^号楼/.test(h) || h === '楼栋号';
    },
    '楼层': function (h) {
      return h.indexOf('楼层') >= 0 || h === '层' || h === '层号' ||
             h.indexOf('层号') >= 0;
    },
    '构件类型': function (h) {
      return h.indexOf('构件类型') >= 0 || h.indexOf('构件名称') >= 0 ||
             h.indexOf('类型') >= 0 || h.indexOf('型号') >= 0 ||
             h.indexOf('类别') >= 0 || h.indexOf('名称') >= 0;
    },
    // '单块体积'与'单块面积'合并为一列：源表通常带/不带单位字眼（m³/m²/方量/砼量）
    // 直接统一识别为单模板列 → 源列写\"单块体积\"、\"单块体积(m³)\"、\"单块面积\"、\"单块面积(m²)\"、\"m³体积\"、\"m²面积\" 等都匹配同一列
    '单块体积/面积': function (h) {
      return h.indexOf('体积') >= 0 || h.indexOf('面积') >= 0 ||
             h.indexOf('方量') >= 0 || h.indexOf('砼量') >= 0 ||
             h.indexOf('m3') >= 0 || h.indexOf('m²') >= 0 || h.indexOf('m2') >= 0;
    },
    '单体质量': function (h) {
      return h.indexOf('质量') >= 0 || h.indexOf('重量') >= 0 || h.indexOf('吨') >= 0;
    },
    '板宽(mm)': function (h) { return h.indexOf('板宽') >= 0 || h.indexOf('宽度') >= 0 || h === '宽'; },
    '板长(mm)': function (h) { return h.indexOf('板长') >= 0 || h.indexOf('长度') >= 0 || h === '长'; },
    '板厚(mm)': function (h) { return h.indexOf('板厚') >= 0 || h.indexOf('厚度') >= 0 || h === '厚'; },
    '砼标号': function (h) {
      return h.indexOf('砼') >= 0 || h.indexOf('强度') >= 0 || h === '标号';
    },
    '备注': function (h) { return h === '备注' || h === '注' || h === '说明' || h === '备注/说明'; }
  };

  // 读表头行文本数组
  function headerTexts(ws, headerRow, maxCols) {
    maxCols = maxCols || 80;
    var arr = [];
    for (var c = 0; c < maxCols; c++) {
      var cell = ws[XLSX.utils.encode_cell({ r: headerRow - 1, c: c })];
      arr.push(cell ? clean(cell.v) : '');
    }
    return arr;
  }

  // 兜底匹配：目标列名与源表头做“包含/去单位”宽松比较（用于用户自定义列名）
  function genericFor(name) {
    return function (h) {
      if (!name) return false;
      var tn = clean(name).replace(/[（(].*?[)）]/g, '');   // 去掉 (mm)/(m³) 等单位后缀
      if (!tn) return false;
      if (h === tn) return true;
      if (tn.length >= 2 && h.length >= 2 && (h.indexOf(tn) >= 0 || tn.indexOf(h) >= 0)) return true;
      return false;
    };
  }

  // 对一张表做列映射：返回 {colName: {idx, text}}；consumed 防止一列多配
  function mapColumns(ws, headerRow, headers) {
    var texts = headerTexts(ws, headerRow);
    var consumed = {};
    var map = {};
    headers.forEach(function (name) {
      map[name] = null;
      if (name === '发货时间') return;             // 特殊列，不占源列
      var fn = SYNONYMS[name] || genericFor(name);
      if (!fn) return;
      for (var i = 0; i < texts.length; i++) {
        if (consumed[i] || !texts[i] || !isHeaderish(texts[i])) continue;
        if (fn(texts[i])) {
          consumed[i] = true;
          map[name] = { idx: i, text: texts[i] };
          break;
        }
      }
    });
    return map;
  }

  /* ---------------- 发货时间识别 ---------------- */

  // 从表头区解析发货日期文本；找不到返回 ''
  function parseSheetDate(ws) {
    var ref = ws['!ref'];
    var rmax = ref ? Math.min(XLSX.utils.decode_range(ref).e.r + 1, 9) : 9;
    for (var r = 0; r < rmax; r++) {
      for (var c = 0; c < 30; c++) {
        var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
        if (!cell || cell.v === null || cell.v === undefined) continue;
        var dt = toDateText(cell.v);
        if (dt) return dt;
      }
    }
    return '';
  }

  /* ---------------- 主合并逻辑 ---------------- */

  /*
   * config = {
   *   files: [ { name, wb } ],            // 已 load 的工作簿
   *   sheets: { "<文件名>": { "<表名>": {
   *        include: bool,
   *        headerRow: int (>0 有效),
   *        dateFixed: '' | 'yyyy-mm-dd'   // 本表强制日期（可选）
   *   } } },
   *   cols: [ { name, enabled,
   *             matchHeader?: string      // 手动指定源表头文本（数据列）
   *             date?: { mode:'auto'|'fixed'|'col', fixed?:string, col?:string }  // 发货时间专用
   *   } ]
   * }
   * 返回 { headers, rows, rowMeta, stats, warnings, verify }
   *   rowMeta: 与 rows 一一对应的来源信息 [{file, sheet, date}]（汇总/校核扩展用，不影响导出列）
   *   verify: 出库总量校核结果 [{source, ok, items, text, exp, act}]，
   *           源单底部“出库总量”行不作为明细导出，其标注总量与明细实际合计逐项比对
   */
  function merge(config) {
    var headers = [], dateEntry = null, dataCols = [];
    config.cols.forEach(function (c) {
      if (!c.enabled) return;
      headers.push(c.name);
      if (c.name === '发货时间') dateEntry = c;
      else dataCols.push(c);
    });
    var rows = [];
    var rowMeta = [];         // 与 rows 平行：每行来自哪个文件/哪张表/哪天
    var warnings = [];
    var stats = {};           // 文件名 -> 行数
    var verify = [];          // 出库总量校核结果
    var dateColIdx = headers.indexOf('发货时间');

    // 楼栋兜底：某些出库单表里没有“楼栋”列（楼栋只写在文件名里，如“…翔安一中12#出库单.xlsx”）。
    // 只在某张表的“楼栋”列完全没匹配到时才填，不会覆盖表里已有的楼栋值。
    var bldFallback = config.bldFallback || '';
    function bldFallbackFor(fname) {
      if (!bldFallback) return '';
      if (typeof bldFallback === 'object') return clean(bldFallback[fname] || '');
      return clean(bldFallback);
    }

    // 在指定表里按“手动指定表头文本”精确找列
    function findByHeader(ws, headerRow, text) {
      var arr = headerTexts(ws, headerRow);
      var t = clean(text);
      for (var i = 0; i < arr.length; i++) {
        if (arr[i] === t) return i;
      }
      return -1;
    }

    config.files.forEach(function (f) {
      var fn = f.name;
      var sheetCfg = (config.sheets[fn] || {});
      var wb = f.wb;
      wb.SheetNames.forEach(function (sn) {
        var sc = (sheetCfg[sn] || {});
        if (!sc.include || !sc.headerRow) return;
        var ws = wb.Sheets[sn];
        var headerRow = sc.headerRow;

        // ---- 列映射 ----
        var map = {};
        // 1) 自动匹配（synonym / generic）
        var autoNames = dataCols.filter(function (c) { return !c.matchHeader; }).map(function (c) { return c.name; });
        var autoMap = mapColumns(ws, headerRow, autoNames);
        dataCols.forEach(function (c) {
          if (c.matchHeader) {
            var idx = findByHeader(ws, headerRow, c.matchHeader);
            map[c.name] = idx >= 0 ? { idx: idx, text: c.matchHeader } : null;
          } else {
            map[c.name] = autoMap[c.name] || null;
          }
        });

        // ---- 主键列 ----
        var keyCol = map['构件编号'] ? map['构件编号'].idx
                   : (function () { for (var i = 0; i < dataCols.length; i++) { if (map[dataCols[i].name]) return map[dataCols[i].name].idx; } return -1; })();
        if (keyCol < 0) {
          warnings.push(fn + ' / ' + sn + ': 未匹配到任何数据列，已跳过');
          return;
        }

        // ---- 「单块体积/面积」列的语义（体积 or 面积）----
        // 源表可能只有体积列（如翔安一中出库单），而构件按面积计价；
        // 这里记下语义，台账/汇总才能把数值放进正确的一格，并在缺另一项时用尺寸推算。
        var qSem = '';
        (function () {
          for (var qi = 0; qi < dataCols.length; qi++) {
            var nm = dataCols[qi].name;
            if (/体积/.test(nm) && /面积/.test(nm) && map[nm]) {
              qSem = qtySemantics(map[nm].text);
              break;
            }
          }
        })();
        var bldFb = bldFallbackFor(fn);

        // ---- 日期解析方式 ----
        var dColIdx = -1;                 // 'col' 模式：本表里日期列位置
        if (dateEntry && dateEntry.date && dateEntry.date.mode === 'col') {
          dColIdx = findByHeader(ws, headerRow, dateEntry.date.col);
        }
        var baseDate = sc.dateFixed || (dateEntry && dateEntry.date && dateEntry.date.mode === 'fixed' ? (dateEntry.date.fixed || '') : '') || '';
        if (!baseDate) baseDate = parseSheetDate(ws);   // auto 兜底

        // ---- 数据行范围 ----
        var ref = ws['!ref'];
        var maxR = ref ? XLSX.utils.decode_range(ref).e.r + 1 : headerRow + 1;
        var lastData = headerRow + 1;
        for (var r = headerRow; r < maxR; r++) {
          var isEmpty = true;
          for (var i = 0; i < dataCols.length; i++) {
            var m = map[dataCols[i].name];
            if (!m) continue;
            var cv = ws[XLSX.utils.encode_cell({ r: r, c: m.idx })];
            if (cv && cv.v !== null && cv.v !== undefined && String(cv.v).trim() !== '') { isEmpty = false; break; }
          }
          if (!isEmpty) lastData = r + 1;
        }

        var fileRows = 0;
        // 明细实际合计（用于与“出库总量”标注值校核）——直接按源表列头识别体积/面积/重量列
        var actSum = { count: 0, volume: 0, area: 0, weight: 0 };
        var sVol = -1, sArea = -1, sWt = -1;
        (function () {
          var hdrs = headerTexts(ws, headerRow);
          for (var i = 0; i < hdrs.length; i++) {
            var h = hdrs[i] || '';
            if (sVol < 0 && /体积|方量/.test(h)) sVol = i;
            if (sArea < 0 && /面积/.test(h)) sArea = i;
            if (sWt < 0 && /重量|质量/.test(h)) sWt = i;
          }
        })();
        var expSum = null;        // 本表解析出的“出库总量”标注值
        for (var r2 = headerRow; r2 < lastData; r2++) {
          var kcell = ws[XLSX.utils.encode_cell({ r: r2, c: keyCol })];
          var kv = kcell ? kcell.v : null;
          if (kv === null || kv === undefined) continue;
          var kt = String(kv).trim();
          if (!kt) continue;
          if (/合计|总计|小计|SUM/i.test(kt)) continue;      // 汇总行
          if (/^(司机|驾驶员|送货人|收货人|签收单位|签收人|打印人|制单人|质检员?|验收人)/.test(kt)) continue; // 签收留尾行
          // 出库总量校核行（如“出库总量|总块数：36块|总方量：6.806m³|总面积…|总重量…”）：
          // 不作为明细导出，解析标注总量用于校核；明细扫描到此为止（其后为签收/周转物料等留尾内容）
          var sumLine = rowText(ws, r2, 32);
          if (/出库总量|总块数|总数量|总方量|总体积|总面积|总重量/.test(sumLine)) {
            var eB = pickNum(sumLine, /总块数[：:]?\s*([0-9]+(?:\.[0-9]+)?)/);
            var eV = pickNum(sumLine, /总方量[：:]?\s*([0-9]+(?:\.[0-9]+)?)/);
            if (eV === null) eV = pickNum(sumLine, /总体积[：:]?\s*([0-9]+(?:\.[0-9]+)?)/);
            var eA = pickNum(sumLine, /总面积[：:]?\s*([0-9]+(?:\.[0-9]+)?)/);
            var eW = pickNum(sumLine, /总重量[：:]?\s*([0-9]+(?:\.[0-9]+)?)/);
            if (eB !== null || eV !== null || eA !== null || eW !== null || /出库总量/.test(sumLine)) {
              expSum = { blocks: eB, volume: eV, area: eA, weight: eW };
              break;
            }
          }
          var out = [];
          for (var hi = 0; hi < headers.length; hi++) {
            var hname = headers[hi];
            if (hname === '发货时间') {
              if (dColIdx >= 0) {
                var dc = ws[XLSX.utils.encode_cell({ r: r2, c: dColIdx })];
                out.push(dc ? toDateText(dc.v) : '');
              } else {
                out.push(baseDate || '');
              }
              continue;
            }
            var mm = map[hname];
            if (!mm) { out.push((hname === '楼栋' && bldFb) ? bldFb : ''); continue; }
            var cell = ws[XLSX.utils.encode_cell({ r: r2, c: mm.idx })];
            var val = cell ? cell.v : null;
            if (val instanceof Date) val = fmtDate(val);
            else if (typeof val === 'string') val = val.trim();
            out.push(val === null || val === undefined ? '' : val);
          }
          rows.push(out);
          rowMeta.push({ file: fn, sheet: sn, qty: qSem,
            date: (dateColIdx >= 0 ? (out[dateColIdx] || '') : (baseDate || '')) });
          fileRows++;
          // 累计明细实际值，供出库总量校核
          actSum.count++;
          if (sVol >= 0)  { var nv = cellNum(ws, r2, sVol);  if (nv !== null) actSum.volume += nv; }
          if (sArea >= 0) { var na = cellNum(ws, r2, sArea); if (na !== null) actSum.area  += na; }
          if (sWt >= 0)   { var nw = cellNum(ws, r2, sWt);   if (nw !== null) actSum.weight += nw; }
        }
        // 表里没有底部“出库总量”行时，再看表头区是否印了总量（如翔安一中出库单 A4/D4/E4）
        if (!expSum) expSum = findHeaderTotal(ws, headerRow);
        if (expSum) verify.push(buildVerify(fn + ' / ' + sn, expSum, actSum));
        stats[fn] = (stats[fn] || 0) + fileRows;
        var missing = dataCols.filter(function (c) { return map[c.name] === null; }).map(function (c) { return c.name; });
        if (missing.length && fileRows > 0) {
          warnings.push(fn + ' / ' + sn + ': 未找到列 ' + missing.join('、') + '（留空）');
        }
        if (bldFb && fileRows > 0 && map['楼栋'] === null) {
          warnings.push(fn + ' / ' + sn + ': 表内没有“楼栋”列，已用兜底值「' + bldFb + '」补齐（可在步骤3改/清空）');
        }
      });
    });
    return { headers: headers, rows: rows, rowMeta: rowMeta, stats: stats, warnings: warnings, verify: verify };
  }

  /* ================= 台账生成（供货明细模板） =================
   * 输入 dataRows（调用方已从合并结果提取好字段）：
   *   { d:'yyyy-mm-dd' 发货日期, bld, fl, type, code:文本,
   *     w, l, h: 板宽/长/厚(拼规格型号), vol, area, c:忽略 }
   * opts = { title, priceRules, segment:{mode:'month'|'cutday'|'single', cutDay},
   *          singleLabel, bldOrder:[楼栋页顺序] }
   * priceRules: [{label, kw[], exclude[], unit:'体积'|'面积', price}] 按顺序首中即用；
   *   规则顺序同时决定同一天内小计块的排列顺序
   * 小计块口径 = (页楼栋 + 发货日期 + 产品名)合并一块，尾部跟“小计：”
   * 楼栋为空的明细不丢弃，归入 opts.blankBldLabel（默认「未填楼栋」）页
   * 返回 { pages:[{name, aoa, mark, amount}], warnings:[...] }
   */
  function buildLedger(dataRows, opts) {
    opts = opts || {};
    var title = opts.title || '预制构件供货明细';
    var blankBld = opts.blankBldLabel || '未填楼栋';   // 楼栋为空的明细归入该页，不再丢弃
    var rules = opts.priceRules && opts.priceRules.length ? opts.priceRules
      : [{ label: 'PC叠合板', kw: ['叠合板'], exclude: ['桁架'], unit: '体积', price: 2300 },
         { label: 'PC楼梯', kw: ['楼梯'], exclude: [], unit: '体积', price: 2500 },
         { label: '钢管桁架预应力叠合板', kw: ['桁架'], exclude: [], unit: '面积', price: 160 }];
    var seg = opts.segment || { mode: 'month', cutDay: 26 };
    var segMode = seg.mode || 'month';
    var cutDay = seg.cutDay || 26;
    var warnings = [];

    function num(v) { var n = parseFloat(v); return isNaN(n) ? null : n; }
    function segKeyOf(d) {
      if (!/^\d{4}-\d{2}-\d{2}/.test(d || '')) return { key: 0, label: '' };
      var y = +d.slice(0, 4), m = +d.slice(5, 7), day = +d.slice(8, 10);
      if (segMode === 'cutday') {
        if (day >= cutDay) return { key: y * 100 + m, label: m + '月' };
        var py = m === 1 ? y - 1 : y, pm = m === 1 ? 12 : m - 1;
        return { key: py * 100 + pm, label: pm + '月' };
      }
      return { key: y * 100 + m, label: m + '月' };
    }
    function pct(x) { return Math.round(x * 100) / 100; }

    var recs = [];
    dataRows.forEach(function (r) {
      var bldRaw = clean(r.bld);
      var bld = bldRaw || blankBld;
      if (!bldRaw) warnings.push('明细缺少楼栋（' + (r.code || '?') + '），已归入「' + blankBld + '」页');
      var type = clean(r.type);
      var rule = null, ruleIdx = 1e9, ri, k;
      for (ri = 0; ri < rules.length; ri++) {
        var rr = rules[ri], hit = false;
        for (k = 0; k < (rr.kw || []).length; k++) if (type.indexOf(rr.kw[k]) >= 0) { hit = true; break; }
        if (!hit) continue;
        var bad = false;
        for (k = 0; k < (rr.exclude || []).length; k++) if (type.indexOf(rr.exclude[k]) >= 0) { bad = true; break; }
        if (!bad) { rule = rr; ruleIdx = ri; break; }
      }
      var w = num(r.w), l = num(r.l), h = num(r.h);
      var spec = (w !== null && l !== null && h !== null) ? w + '*' + l + '*' + h : '';
      if (!rule) {
        warnings.push('构件类型“' + (type || '(空)') + '”未匹配计价规则（' + bld + ' ' + (r.code || '') + '），按单价0记入，请检查');
        rule = { label: type || '未分类', unit: '体积', price: 0 };
      }
      var qty = num(rule.unit === '面积' ? r.area : r.vol);
      if (qty === null || qty < 0) {
        qty = '';
        warnings.push('构件缺少' + (rule.unit === '面积' ? '面积' : '体积') + '：' + rule.label + ' ' + (r.code || '') + '（' + bld + '），金额留空');
      }
      recs.push({
        c: r.c || 0, d: r.d || '', bld: bld, bldRaw: bldRaw, fl: clean(r.fl), li: ruleIdx,
        label: rule.label, price: rule.price, unit: rule.unit,
        code: String(r.code == null ? '' : r.code).trim(), spec: spec,
        qty: qty === '' ? '' : qty,
        amt: qty === '' ? '' : pct(qty * rule.price)
      });
    });

    var maxLbl = '';
    recs.forEach(function (r) {
      var sk = segKeyOf(r.d);
      if (segMode === 'single') { r.sk = 1; r.sl = ''; return; }
      r.sk = sk.key; r.sl = sk.label;
      if (sk.label) maxLbl = sk.label;
    });
    var singleLabel = segMode === 'single' ? (opts.singleLabel || maxLbl || '') : '';

    var pageSeen = {}, seenDate = {}, pageList = [];
    recs.forEach(function (r) {
      if (!pageSeen[r.bld]) { pageSeen[r.bld] = true; seenDate[r.bld] = r.d; pageList.push(r.bld); }
    });
    var prefer = (opts.bldOrder || []).map(function (b) { return clean(b); });
    pageList.sort(function (a, b) {
      var ia = prefer.indexOf(a), ib = prefer.indexOf(b);
      if (ia >= 0 || ib >= 0) {
        if (ia >= 0 && ib >= 0) return ia - ib;
        return ia >= 0 ? -1 : 1;
      }
      return seenDate[a] < seenDate[b] ? -1 : seenDate[a] > seenDate[b] ? 1 : 0;
    });

    var HEAD = ['日期', '产品名称', '构件编号', '规格型号', '楼栋', '楼层', '单块体积/面积', '含税单价', '含税金额'];
    var pages = [];
    pageList.forEach(function (bld) {
      // 小计块口径 = (发货日期 + 产品名)：同一天同楼栋同类构件合并一块（对应样例台账的小计）
      var cars = {}, orderC = [];
      recs.forEach(function (r) {
        if (r.bld !== bld) return;
        var key = r.d + '\u0001' + r.label;
        if (!cars[key]) { cars[key] = { key: key, d: r.d, label: r.label, li: r.li, rows: [] }; orderC.push(key); }
        var car = cars[key];
        car.rows.push(r);
      });
      var carArr = orderC.map(function (key) { return cars[key]; });
      carArr.sort(function (a, b) {
        var ka = segKeyOf(a.d).key || 1, kb = segKeyOf(b.d).key || 1;
        if (ka !== kb) return ka - kb;
        if (a.d !== b.d) return a.d < b.d ? -1 : 1;
        if (a.li !== b.li) return a.li - b.li;
        return 0;
      });
      var segs = [], cur = null;
      carArr.forEach(function (car) {
        var key = segMode === 'single' ? 1 : (segKeyOf(car.d).key || 0);
        if (!cur || cur.key !== key) {
          cur = { key: key, cars: [], label: segMode === 'single' ? singleLabel : (segKeyOf(car.d).label || '') };
          segs.push(cur);
        }
        cur.cars.push(car);
      });
      var aoa = [[title, '', '', '', '', '', '', '', '']];
      var mark = [{ r: 0, label: title, kind: 'title' }];
      var pageTotal = 0, first = true;
      segs.forEach(function (sg) {
        aoa.push([sg.label, '', '', '', '', '', '', '', '']);
        mark.push({ r: aoa.length - 1, label: sg.label, kind: 'segname' });
        if (first) { aoa.push(HEAD.slice()); mark.push({ r: aoa.length - 1, label: '', kind: 'hdr' }); first = false; }
        var sgSum = 0;
        sg.cars.forEach(function (car) {
          var carSum = 0;
          car.rows.forEach(function (r) {
            aoa.push([r.d, r.label, r.code, r.spec, r.bldRaw || '', r.fl, r.qty, r.price, r.amt]);
            mark.push({ r: aoa.length - 1, label: r.code, kind: 'data' });
            carSum += (typeof r.amt === 'number' ? r.amt : 0);
          });
          sgSum += carSum;
          aoa.push(['小计：', '', '', '', '', '', '', '', pct(carSum)]);
          mark.push({ r: aoa.length - 1, label: '小计', kind: 'sub' });
        });
        pageTotal += sgSum;
        aoa.push(['本月合计：', '', '', '', '', '', '', '', pct(sgSum)]);
        mark.push({ r: aoa.length - 1, label: '本月合计', kind: 'segsum' });
      });
      if (segs.length > 1) {
        aoa.push(['总计：', '', '', '', '', '', '', '', pct(pageTotal)]);
        mark.push({ r: aoa.length - 1, label: '总计', kind: 'total' });
      }
      pages.push({ name: bld, aoa: aoa, mark: mark, amount: pct(pageTotal) });
    });

    function uniqWarn(list) {
      var cnt = {}, ord = [];
      list.forEach(function (m) { if (!(m in cnt)) { cnt[m] = 0; ord.push(m); } cnt[m]++; });
      return ord.map(function (m) { return cnt[m] > 1 ? m + '（共 ' + cnt[m] + ' 条）' : m; });
    }
    return { pages: pages, warnings: uniqWarn(warnings) };
  }

  /* ================= 当天单楼栋发货汇总（日期 × 楼栋 × 构件类型） =================
   * items（调用方从合并结果提取）：
   *   { d:'yyyy-mm-dd', bld, type, va(单块体积/面积值), wt(单体质量), src:'文件/表'(计车次) }
   * opts = { priceRules, blankBldLabel, splitType(默认true: 同日同楼栋不同构件类型拆成多行) }
   * 计价规则与台账一致（首中即用）；金额 = 数量×含税单价，数量按规则的计量方式取 va。
   * 返回 { groups:[{date,bld,label,labels,cars,count,vol,wt,area,amt}], totals:{...}, warnings }
   * totals.cars = 去重后的实际送货单张数（不因拆类型重复计）
   */
  function matchRule(type, rules) {
    for (var ri = 0; ri < rules.length; ri++) {
      var rr = rules[ri], hit = false, k;
      for (k = 0; k < (rr.kw || []).length; k++) if (type.indexOf(rr.kw[k]) >= 0) { hit = true; break; }
      if (!hit) continue;
      var bad = false;
      for (k = 0; k < (rr.exclude || []).length; k++) if (type.indexOf(rr.exclude[k]) >= 0) { bad = true; break; }
      if (!bad) return { rule: rr, idx: ri };
    }
    return null;
  }

  function buildDailySummary(items, opts) {
    opts = opts || {};
    var blankBld = opts.blankBldLabel || '未填楼栋';
    var rules = opts.priceRules && opts.priceRules.length ? opts.priceRules
      : [{ label: 'PC叠合板', kw: ['叠合板'], exclude: ['桁架'], unit: '体积', price: 2300 },
         { label: 'PC楼梯', kw: ['楼梯'], exclude: [], unit: '体积', price: 2500 },
         { label: '钢管桁架预应力叠合板', kw: ['桁架'], exclude: [], unit: '面积', price: 160 }];
    var warnings = [];
    function num(v) { var n = parseFloat(v); return isNaN(n) ? null : n; }
    function r2(x) { return Math.round(x * 100) / 100; }

    var splitType = opts.splitType !== false;   // 默认按 构件类型 拆分：同一天同楼栋不同构件各占一行
    var map = {};      // key: date+\u0001+bld(\u0001+label)
    var order = [];
    var carSeen = {};  // 全局车次去重（合计行的车次=实际送货单张数，不因拆类型而重复计）
    items.forEach(function (it) {
      var d = toDateText(it.d);
      if (!d) { warnings.push('明细缺少日期（' + (it.src || '?') + '），未计入汇总'); return; }
      var bld = clean(it.bld) || blankBld;
      var type = clean(it.type);
      var m = matchRule(type, rules);
      if (!m) {
        warnings.push('构件类型“' + (type || '(空)') + '”未匹配计价规则（' + bld + '），金额按 0 记');
        m = { rule: { label: type || '未分类', unit: '体积', price: 0 }, idx: 1e9 };
      }
      // 取值口径：优先用调用方按语义分开给的 vol / area（缺失的那项可由尺寸推算）；
      // 没给这两项时退回旧的单列 va，保证既有项目（如太和、尚谷大院）数值不变。
      var raw = (m.rule.unit === '面积') ? it.area : it.vol;
      var qty = (raw === undefined || raw === null || raw === '') ? num(it.va) : num(raw);
      if (qty === null || qty < 0) {
        qty = 0;
        warnings.push('构件缺少' + (m.rule.unit === '面积' ? '面积' : '体积') + '：' + (it.src || '') + '（' + bld + '），按 0 计');
      }
      var wt = num(it.wt); if (wt === null || wt < 0) wt = 0;
      // 体积/面积两列都尽量填：单价仍按规则的计量方式（rule.unit）算，但汇总表两栏都有数
      // （源表缺的那一栏由调用方用尺寸推算）。调用方没给 vol/area 时退回旧口径。
      var qVol = (it.vol === undefined || it.vol === null || it.vol === '') ? null : num(it.vol);
      var qArea = (it.area === undefined || it.area === null || it.area === '') ? null : num(it.area);
      if (qVol !== null && qVol < 0) qVol = null;
      if (qArea !== null && qArea < 0) qArea = null;
      var label = m.rule.label;
      var key = d + '\u0001' + bld + (splitType ? '\u0001' + label : '');
      var g = map[key];
      if (!g) {
        g = map[key] = { date: d, bld: bld, label: label, li: m.idx,
                         labels: [label], labelSet: {}, cars: {}, carN: 0,
                         count: 0, vol: 0, wt: 0, area: 0, amt: 0 };
        g.labelSet[label] = 1;
        order.push(key);
      }
      if (!splitType && !g.labelSet[label]) { g.labelSet[label] = 1; g.labels.push(label); }
      g.count++;
      if (qVol === null && qArea === null) {
        if (m.rule.unit === '面积') g.area += qty; else g.vol += qty;
      } else {
        if (qVol !== null) g.vol += qVol;
        if (qArea !== null) g.area += qArea;
      }
      g.wt += wt;
      g.amt += qty * m.rule.price;
      var src = clean(it.src);
      if (src) {
        if (!g.cars[src]) { g.cars[src] = 1; g.carN++; }
        if (!carSeen[src]) carSeen[src] = 1;
      }
    });

    order.sort(function (a, b) {
      var ga = map[a], gb = map[b];
      if (ga.date !== gb.date) return ga.date < gb.date ? -1 : 1;
      if (ga.bld !== gb.bld) return ga.bld < gb.bld ? -1 : 1;
      if (ga.li !== gb.li) return ga.li - gb.li;
      if (ga.label !== gb.label) return ga.label < gb.label ? -1 : 1;
      return 0;
    });
    var groups = order.map(function (k) {
      var g = map[k];
      return { date: g.date, bld: g.bld, label: g.label, labels: g.labels.join('、'),
               cars: g.carN, count: g.count, vol: round3(g.vol), wt: round3(g.wt),
               area: round3(g.area), amt: r2(g.amt) };
    });
    var totals = { cars: 0, count: 0, vol: 0, wt: 0, area: 0, amt: 0 };
    Object.keys(carSeen).forEach(function () { totals.cars++; });
    groups.forEach(function (g) {
      totals.count += g.count;
      totals.vol += g.vol; totals.wt += g.wt; totals.area += g.area; totals.amt += g.amt;
    });
    totals.vol = round3(totals.vol); totals.wt = round3(totals.wt);
    totals.area = round3(totals.area); totals.amt = r2(totals.amt);
    function uniqWarn(list) {
      var cnt = {}, ord = [];
      list.forEach(function (m) { if (!(m in cnt)) { cnt[m] = 0; ord.push(m); } cnt[m]++; });
      return ord.map(function (m) { return cnt[m] > 1 ? m + '（共 ' + cnt[m] + ' 条）' : m; });
    }
    return { groups: groups, totals: totals, warnings: uniqWarn(warnings) };
  }

  /* ---------- 合计校核：明细实际合计 vs 期望值（如对账单汇总页抄来的数） ---------- */
  /* actual / expected: {count, vol, wt, area, amt} 中任意子集，null 忽略该项 */
  function verifyTotals(actual, expected) {
    var items = [];
    function cmp(label, e, a, unit, tol) {
      if (e === null || e === undefined || isNaN(e)) return;
      a = a || 0;
      var ok = Math.abs(e - a) <= tol;
      items.push({
        label: label, exp: e, act: round3(a), unit: unit || '', ok: ok,
        text: label + ' 期望 ' + e + (unit || '') + ' ／ 实际合计 ' + round3(a) + (unit || '') +
              (ok ? ' ✔' : ' ✘ 差 ' + round3(a - e) + (unit || ''))
      });
    }
    var e = expected || {};
    cmp('件数', e.count, actual.count, '件', 0.001);
    cmp('体积', e.vol, actual.vol, 'm³', Math.max(0.02, (e.vol || 0) * 0.005));
    cmp('重量', e.wt, actual.wt, 'T', Math.max(0.02, (e.wt || 0) * 0.005));
    cmp('面积', e.area, actual.area, 'm²', Math.max(0.05, (e.area || 0) * 0.005));
    cmp('金额', e.amt, actual.amt, '元', Math.max(0.05, (e.amt || 0) * 0.005));
    return {
      ok: items.length > 0 && items.every(function (x) { return x.ok; }),
      items: items,
      text: items.map(function (x) { return x.text; }).join('；')
    };
  }

  /* ---------- 汇总导出（浏览器用） ---------- */
  var SUM_HEAD = ['发货日期', '楼栋', '产品名称', '车次数', '构件件数', '体积(m³)', '重量(t)', '面积(m²)', '含税金额(元)'];
  function exportSummaryWorkbook(sum) {
    var aoa = [SUM_HEAD.slice()];
    (sum.groups || []).forEach(function (g) {
      aoa.push([g.date, g.bld, g.labels, g.cars, g.count, g.vol, g.wt, g.area, g.amt]);
    });
    var t = sum.totals || {};
    aoa.push(['合计', '', '', t.cars || 0, t.count || 0, t.vol || 0, t.wt || 0, t.area || 0, t.amt || 0]);
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 12 }, { wch: 10 }, { wch: 26 }, { wch: 8 }, { wch: 10 }, { wch: 11 }, { wch: 10 }, { wch: 11 }, { wch: 14 }];
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '当天单楼栋发货汇总');
    return wb;
  }

  /* ---------------- 导出（浏览器用） ---------------- */
  function exportWorkbook(rows, headers, sheetName) {
    var aoa = [headers].concat(rows);
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName || '汇总');
    return wb;
  }

  // 台账导出：pages=[{name,aoa,mark}] → 每楼栋一页，标题/段名/小计/合计行做横向合并
  function exportLedgerWorkbook(pages) {
    var wb = XLSX.utils.book_new();
    var used = {};
    (pages || []).forEach(function (pg, pi) {
      if (!pg || !pg.aoa || !pg.aoa.length) return;
      var ws = XLSX.utils.aoa_to_sheet(pg.aoa);
      var merges = [];
      (pg.mark || []).forEach(function (m) {
        var r = m.r;
        if (!m.kind) return;
        if (m.kind === 'title' || m.kind === 'segname') merges.push({ s: { r: r, c: 0 }, e: { r: r, c: 8 } });
        else if (m.kind === 'sub' || m.kind === 'segsum' || m.kind === 'total') merges.push({ s: { r: r, c: 0 }, e: { r: r, c: 7 } });
      });
      if (merges.length) ws['!merges'] = merges;
      ws['!cols'] = [{ wch: 11 }, { wch: 18 }, { wch: 15 }, { wch: 15 }, { wch: 9 }, { wch: 9 }, { wch: 12 }, { wch: 9 }, { wch: 12 }];
      var nm = String(pg.name == null ? '页' + (pi + 1) : pg.name).replace(/[\\\/\?\*\[\]:]/g, '').slice(0, 31) || ('页' + (pi + 1));
      var base = nm, k = 2;
      while (used[nm]) nm = base.slice(0, 28) + '(' + (k++) + ')';
      used[nm] = 1;
      XLSX.utils.book_append_sheet(wb, ws, nm);
    });
    return wb;
  }

  /* ---------- 原始货单附带导出 ----------
   * 从 merge 的同一份 config 里，取出所有被勾选(include)的工作表原始内容，
   * 供导出时以“每张原始单 = 一个工作表”的方式附到导出工作簿后面。
   * 返回 [{file, sheet, name(建议表名，≤31字), aoa, cols(建议列宽)}]
   */
  function collectSourceSheets(config) {
    var out = [];
    (config.files || []).forEach(function (f) {
      var fn = f.name || '文件';
      var sheetCfg = (config.sheets && config.sheets[fn]) || {};
      var wb = f.wb;
      if (!wb || !wb.SheetNames) return;
      // 本文件里被勾选的表数（>1 时表名要带 sheet 名区分）
      var incNames = wb.SheetNames.filter(function (sn) {
        var sc = sheetCfg[sn] || {};
        return sc.include !== false && (sc.include === true || sc.include);
      });
      wb.SheetNames.forEach(function (sn) {
        var sc = sheetCfg[sn] || {};
        if (!sc.include) return;
        var ws = wb.Sheets[sn];
        if (!ws || !ws['!ref']) return;
        var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
        // 建议表名（≤31 字）：优先用文件名；一单多表时用“文件名-sheet名”，
        // 文件名过长放不下时退回 sheet 名（信息在日期/车次上，比截断的文件名有用）
        var base = String(fn).replace(/\.xlsx$/i, '').replace(/[\\\/\?\*\[\]:]/g, '');
        var nm;
        if (incNames.length > 1) {
          var full = base + '-' + sn;
          nm = full.length <= 31 ? full : sn.slice(0, 31);
        } else {
          nm = base.slice(0, 31);
        }
        out.push({
          file: fn, sheet: sn, name: nm, aoa: aoa,
          cols: (ws['!cols'] || []).slice(0, 40)
        });
      });
    });
    return out;
  }

  /* ---------------- 导出接口 ---------------- */
  var api = {
    VERSION: VERSION,
    BUILD_DATE: BUILD_DATE,
    loadWorkbook: loadWorkbook,
    previewGrid: previewGrid,
    detectHeaderRow: detectHeaderRow,
    hasDeliveryMarkers: hasDeliveryMarkers,
    mapColumns: mapColumns,
    headerTexts: headerTexts,
    parseSheetDate: parseSheetDate,
    toDateText: toDateText,
    clean: clean,
    merge: merge,
    buildLedger: buildLedger,
    buildDailySummary: buildDailySummary,
    verifyTotals: verifyTotals,
    exportWorkbook: exportWorkbook,
    exportLedgerWorkbook: exportLedgerWorkbook,
    exportSummaryWorkbook: exportSummaryWorkbook,
    collectSourceSheets: collectSourceSheets,
    SYNONYMS: SYNONYMS
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  (typeof globalThis !== 'undefined' ? globalThis : global).BillMerge = api;
})(this);
