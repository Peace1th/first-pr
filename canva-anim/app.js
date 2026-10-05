(function () {
  'use strict';

  var A = window.PptxAnim;
  var NS = {
    p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
    mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006'
  };
  var EMU_PX = 9525;
  var DIR_LABEL = { left: '左から', right: '右から', top: '上から', bottom: '下から' };
  var SHAPE_TAGS = ['sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp'];
  var KIND_LABEL = { sp: '図形/テキスト', pic: '画像', grpSp: 'グループ', graphicFrame: '表/グラフ', cxnSp: '線' };

  var $ = function (s) { return document.querySelector(s); };
  var state = { zip: null, fileName: '', size: { cx: 9144000, cy: 5143500 }, slides: [], cur: 0, selSpid: null, urls: [], playing: [] };

  /* ---------- 読み込み ---------- */

  function parseXml(text, label) {
    var doc = new DOMParser().parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error(label + ' を読み取れませんでした');
    return doc;
  }

  function dirname(path) { return path.slice(0, path.lastIndexOf('/') + 1); }

  function resolvePath(base, target) {
    if (target.charAt(0) === '/') return target.slice(1);
    var parts = (dirname(base) + target).split('/'), out = [];
    parts.forEach(function (p) {
      if (p === '..') out.pop(); else if (p && p !== '.') out.push(p);
    });
    return out.join('/');
  }

  function readRels(zip, relsPath, ownerPath) {
    var f = zip.file(relsPath);
    if (!f) return Promise.resolve({});
    return f.async('string').then(function (text) {
      var map = {};
      var doc = parseXml(text, relsPath);
      Array.prototype.forEach.call(doc.getElementsByTagNameNS(NS.rel, 'Relationship'), function (el) {
        if (el.getAttribute('TargetMode') === 'External') return;
        map[el.getAttribute('Id')] = resolvePath(ownerPath, el.getAttribute('Target'));
      });
      return map;
    });
  }

  function relsPathOf(path) { return dirname(path) + '_rels/' + path.slice(path.lastIndexOf('/') + 1) + '.rels'; }

  function childByName(el, ns, name) {
    for (var c = el.firstElementChild; c; c = c.nextElementSibling) {
      if (c.localName === name && c.namespaceURI === ns) return c;
    }
    return null;
  }

  function getXfrm(el) {
    var tag = el.localName, x = null;
    if (tag === 'graphicFrame') x = childByName(el, NS.p, 'xfrm');
    else {
      var pr = childByName(el, NS.p, tag === 'grpSp' ? 'grpSpPr' : 'spPr');
      x = pr && childByName(pr, NS.a, 'xfrm');
    }
    if (!x) return null;
    var off = childByName(x, NS.a, 'off'), ext = childByName(x, NS.a, 'ext');
    if (!off || !ext) return null;
    return {
      x: +off.getAttribute('x') || 0, y: +off.getAttribute('y') || 0,
      w: +ext.getAttribute('cx') || 0, h: +ext.getAttribute('cy') || 0,
      rot: (+x.getAttribute('rot') || 0) / 60000
    };
  }

  function solidColor(parent) {
    var fill = parent && childByName(parent, NS.a, 'solidFill');
    var c = fill && childByName(fill, NS.a, 'srgbClr');
    return c ? '#' + c.getAttribute('val') : null;
  }

  function extractShape(el, rels) {
    var tag = el.localName;
    var nv = el.getElementsByTagNameNS(NS.p, 'cNvPr')[0];
    if (!nv) return null;
    var texts = Array.prototype.map.call(el.getElementsByTagNameNS(NS.a, 't'), function (t) { return t.textContent; });
    var paras = Array.prototype.map.call(el.getElementsByTagNameNS(NS.a, 'p'), function (p) {
      return Array.prototype.map.call(p.getElementsByTagNameNS(NS.a, 't'), function (t) { return t.textContent; }).join('');
    });
    var shape = {
      spid: +nv.getAttribute('id'), name: nv.getAttribute('name') || '', tag: tag,
      box: getXfrm(el), text: paras.join('\n').trim(), snippet: texts.join('').trim().slice(0, 24),
      img: null, crop: null, fill: null, fontPx: 0, color: null
    };
    var spPr = childByName(el, NS.p, 'spPr');
    shape.fill = solidColor(spPr);
    if (tag === 'pic') {
      var blip = el.getElementsByTagNameNS(NS.a, 'blip')[0];
      var rid = blip && blip.getAttributeNS(NS.r, 'embed');
      shape.img = rid && rels[rid] || null;
      var src = el.getElementsByTagNameNS(NS.a, 'srcRect')[0];
      if (src) {
        shape.crop = ['l', 't', 'r', 'b'].map(function (k) { return (+src.getAttribute(k) || 0) / 100000; });
      }
    }
    var rPr = el.getElementsByTagNameNS(NS.a, 'rPr')[0];
    if (rPr) {
      if (rPr.getAttribute('sz')) shape.fontPx = (+rPr.getAttribute('sz') / 100) * 96 / 72;
      shape.color = solidColor(rPr);
    }
    return shape;
  }

  function loadFile(file) {
    showMsg('読み込み中…');
    revokeUrls();
    var zip;
    return JSZip.loadAsync(file).then(function (z) {
      zip = z;
      if (!zip.file('ppt/presentation.xml')) throw new Error('PowerPointファイル(.pptx)ではないようです');
      return Promise.all([
        zip.file('ppt/presentation.xml').async('string'),
        readRels(zip, 'ppt/_rels/presentation.xml.rels', 'ppt/presentation.xml')
      ]);
    }).then(function (res) {
      var pres = parseXml(res[0], 'presentation.xml'), rels = res[1];
      var sz = pres.getElementsByTagNameNS(NS.p, 'sldSz')[0];
      var size = sz ? { cx: +sz.getAttribute('cx'), cy: +sz.getAttribute('cy') } : { cx: 9144000, cy: 5143500 };
      var ids = Array.prototype.map.call(pres.getElementsByTagNameNS(NS.p, 'sldId'), function (el) {
        return rels[el.getAttributeNS(NS.r, 'id')];
      }).filter(Boolean);
      if (!ids.length) throw new Error('スライドが見つかりませんでした');
      return Promise.all(ids.map(function (path) { return loadSlide(zip, path); })).then(function (slides) {
        state.zip = zip; state.fileName = file.name; state.size = size; state.slides = slides;
        state.cur = 0; state.selSpid = null;
        $('#app').hidden = false;
        var hadAnim = slides.some(function (s) { return s.hadTiming; });
        showMsg(slides.length + ' 枚のスライドを読み込みました。' +
          (hadAnim ? ' 既存のアニメーションがあるスライドは、設定した場合のみ置き換えます。' : ''));
        renderAll();
      });
    }).catch(function (err) {
      showMsg(err.message || String(err), true);
    });
  }

  function loadSlide(zip, path) {
    return Promise.all([zip.file(path).async('string'), readRels(zip, relsPathOf(path), path)]).then(function (r) {
      var doc = parseXml(r[0], path), rels = r[1];
      var tree = doc.getElementsByTagNameNS(NS.p, 'spTree')[0];
      var shapes = [];
      for (var c = tree ? tree.firstElementChild : null; c; c = c.nextElementSibling) {
        if (c.namespaceURI === NS.p && SHAPE_TAGS.indexOf(c.localName) >= 0) {
          var s = extractShape(c, rels);
          if (s) shapes.push(s);
        }
      }
      var urlJobs = shapes.filter(function (s) { return s.img && zip.file(s.img); }).map(function (s) {
        return zip.file(s.img).async('blob').then(function (b) {
          var ext = s.img.split('.').pop().toLowerCase();
          var types = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml', webp: 'image/webp' };
          var url = URL.createObjectURL(types[ext] ? new Blob([b], { type: types[ext] }) : b);
          state.urls.push(url); s.url = url;
        });
      });
      return Promise.all(urlJobs).then(function () {
        return {
          path: path, doc: doc, shapes: shapes, anims: [], transition: 'none', touched: false,
          hadTiming: !!doc.getElementsByTagNameNS(NS.p, 'timing').length
        };
      });
    });
  }

  function revokeUrls() { state.urls.forEach(function (u) { URL.revokeObjectURL(u); }); state.urls = []; }

  /* ---------- 状態の操作 ---------- */

  function curSlide() { return state.slides[state.cur]; }
  function shapeOf(slide, spid) { return slide.shapes.filter(function (s) { return s.spid === spid; })[0]; }
  function shapeLabel(s) {
    return (KIND_LABEL[s.tag] || s.tag) + (s.snippet ? ' 「' + s.snippet + '」' : ' ' + s.name);
  }
  function touch() { curSlide().touched = true; }

  function addAnim(spid, props) {
    var s = curSlide();
    s.anims.push(Object.assign({ spid: spid, effect: 'fade', dir: null, trigger: 'click', dur: 500, delay: 0 }, props || {}));
    touch();
  }

  function isBackground(shape) {
    var b = shape.box;
    return !!b && b.w * b.h >= state.size.cx * state.size.cy * 0.9;
  }

  function readingOrder(slide) {
    var row = state.size.cy * 0.04;
    return slide.shapes.filter(function (s) { return s.box && !isBackground(s); }).sort(function (a, b) {
      var ra = Math.round(a.box.y / row), rb = Math.round(b.box.y / row);
      return ra - rb || a.box.x - b.box.x;
    });
  }

  function bulkApply(slides, effect, dir, mode) {
    slides.forEach(function (slide) {
      var list = readingOrder(slide);
      slide.anims = list.map(function (shape, i) {
        var trigger;
        if (mode === 'click') trigger = 'click';
        else if (mode === 'auto') trigger = 'after';
        else trigger = i === 0 ? 'click' : (mode === 'click-with' ? 'with' : 'after');
        return { spid: shape.spid, effect: effect, dir: dir, trigger: trigger, dur: 500, delay: 0 };
      });
      slide.touched = true;
    });
  }

  /* ---------- 画面描画 ---------- */

  function renderAll() {
    stopPlay();
    renderSlides(); renderStage(); renderPanel(); renderOrder(); renderElements(); renderSummary();
    $('#transition').value = curSlide().transition;
  }

  function renderSlides() {
    var ol = $('#slideList');
    ol.textContent = '';
    state.slides.forEach(function (s, i) {
      var li = document.createElement('li'), b = document.createElement('button');
      b.type = 'button';
      b.textContent = (i + 1) + ' 枚目';
      if (i === state.cur) b.setAttribute('aria-current', 'true');
      var badge = document.createElement('span');
      badge.className = 'badge';
      badge.textContent = s.anims.length ? '✨' + s.anims.length : '';
      b.appendChild(badge);
      b.onclick = function () { state.cur = i; state.selSpid = null; renderAll(); };
      li.appendChild(b); ol.appendChild(li);
    });
  }

  function fitStage() {
    var wrap = $('#stageWrap'), st = $('#stage');
    var w = state.size.cx / EMU_PX, h = state.size.cy / EMU_PX;
    var k = wrap.clientWidth / w;
    st.style.width = w + 'px'; st.style.height = h + 'px';
    st.style.transform = 'scale(' + k + ')';
    wrap.style.height = h * k + 'px';
  }

  function renderStage() {
    var st = $('#stage'), slide = curSlide();
    st.textContent = '';
    st.classList.remove('playing');
    slide.shapes.forEach(function (s) {
      var d = document.createElement('div');
      d.className = 'shp';
      d.dataset.spid = s.spid;
      if (s.box) {
        d.classList.add('has-box');
        d.style.left = s.box.x / EMU_PX + 'px'; d.style.top = s.box.y / EMU_PX + 'px';
        d.style.width = s.box.w / EMU_PX + 'px'; d.style.height = s.box.h / EMU_PX + 'px';
        if (s.box.rot) d.style.transform = 'rotate(' + s.box.rot + 'deg)';
      } else d.style.display = 'none';
      if (s.fill) d.style.backgroundColor = s.fill;
      if (s.url) {
        d.style.backgroundImage = 'url("' + s.url + '")';
        if (s.crop && (s.crop[0] + s.crop[2] < 1) && (s.crop[1] + s.crop[3] < 1)) {
          var vw = 1 - s.crop[0] - s.crop[2], vh = 1 - s.crop[1] - s.crop[3];
          d.style.backgroundSize = 100 / vw + '% ' + 100 / vh + '%';
          var px = s.crop[0] + s.crop[2] > 0 ? s.crop[0] / (s.crop[0] + s.crop[2]) * 100 : 0;
          var py = s.crop[1] + s.crop[3] > 0 ? s.crop[1] / (s.crop[1] + s.crop[3]) * 100 : 0;
          d.style.backgroundPosition = px + '% ' + py + '%';
        } else d.style.backgroundSize = '100% 100%';
      }
      if (s.text) {
        d.textContent = s.text;
        d.style.fontSize = (s.fontPx || 24) + 'px';
        if (s.color) d.style.color = s.color;
      }
      if (s.spid === state.selSpid) d.classList.add('sel');
      var idx = animIndexes(slide, s.spid);
      if (idx.length) {
        var tag = document.createElement('span');
        tag.className = 'tag'; tag.textContent = idx.join(',');
        d.appendChild(tag);
      }
      d.onclick = function () { select(s.spid); };
      st.appendChild(d);
    });
    fitStage();
  }

  function animIndexes(slide, spid) {
    var out = [];
    slide.anims.forEach(function (a, i) { if (a.spid === spid) out.push(i + 1); });
    return out;
  }

  function select(spid) {
    state.selSpid = spid;
    stopPlay();
    renderStage(); renderPanel(); renderOrder(); renderElements();
  }

  function option(value, label) {
    var o = document.createElement('option');
    o.value = value; o.textContent = label;
    return o;
  }

  function field(label, control, full) {
    var l = document.createElement('label');
    if (full) l.className = 'full';
    l.appendChild(document.createTextNode(label));
    l.appendChild(control);
    return l;
  }

  function renderPanel() {
    var box = $('#panel'), slide = curSlide();
    box.textContent = '';
    var shape = state.selSpid != null && shapeOf(slide, state.selSpid);
    if (!shape) {
      var p = document.createElement('p');
      p.className = 'empty'; p.textContent = '左のプレビューか下の要素一覧から、要素を選んでください。';
      box.appendChild(p);
      return;
    }
    var n = document.createElement('div');
    n.className = 'name'; n.textContent = shapeLabel(shape);
    box.appendChild(n);
    var sub = document.createElement('div');
    sub.className = 'sub'; sub.textContent = 'ID ' + shape.spid + (shape.text ? ' ・ ' + shape.text.replace(/\s+/g, ' ').slice(0, 60) : '');
    box.appendChild(sub);

    slide.anims.forEach(function (a, i) {
      if (a.spid !== shape.spid) return;
      box.appendChild(animCard(a, i));
    });

    var add = document.createElement('button');
    add.type = 'button'; add.textContent = '＋ アニメーションを追加';
    add.onclick = function () {
      var hasEntr = slide.anims.some(function (x) { return x.spid === shape.spid && A.EFFECTS[x.effect].cls === 'entr'; });
      addAnim(shape.spid, hasEntr ? { effect: 'fadeOut' } : {});
      renderAll();
    };
    box.appendChild(add);
  }

  function animCard(a, i) {
    var card = document.createElement('div');
    card.className = 'anim-card';

    var eff = document.createElement('select');
    Object.keys(A.EFFECTS).forEach(function (k) {
      eff.appendChild(option(k, (A.EFFECTS[k].cls === 'exit' ? '[終了] ' : '[開始] ') + A.EFFECTS[k].label));
    });
    eff.value = a.effect;
    eff.onchange = function () {
      a.effect = eff.value; a.dir = A.EFFECTS[a.effect].dirs ? (A.EFFECTS[a.effect].defaultDir) : null;
      touch(); renderAll();
    };
    card.appendChild(field('効果 ' + (i + 1) + ' 番目', eff, true));

    if (A.EFFECTS[a.effect].dirs) {
      var dir = document.createElement('select');
      Object.keys(DIR_LABEL).forEach(function (k) { dir.appendChild(option(k, DIR_LABEL[k])); });
      dir.value = a.dir || A.EFFECTS[a.effect].defaultDir;
      dir.onchange = function () { a.dir = dir.value; touch(); renderAll(); };
      card.appendChild(field('向き', dir, true));
    }

    var trg = document.createElement('select');
    Object.keys(A.TRIGGERS).forEach(function (k) { trg.appendChild(option(k, A.TRIGGERS[k])); });
    trg.value = a.trigger;
    trg.onchange = function () { a.trigger = trg.value; touch(); renderAll(); };
    card.appendChild(field('開始のタイミング', trg, true));

    [['dur', '時間(ミリ秒)', 50], ['delay', '遅延(ミリ秒)', 0]].forEach(function (f) {
      var inp = document.createElement('input');
      inp.type = 'number'; inp.min = f[2]; inp.step = 100; inp.value = a[f[0]];
      inp.onchange = function () {
        var v = Number(inp.value);
        a[f[0]] = isFinite(v) ? Math.max(f[2], v) : a[f[0]];
        touch(); renderAll();
      };
      card.appendChild(field(f[1], inp));
    });

    var del = document.createElement('button');
    del.type = 'button'; del.className = 'ghost full'; del.textContent = 'このアニメーションを削除';
    del.onclick = function () { curSlide().anims.splice(i, 1); touch(); renderAll(); };
    card.appendChild(del);
    return card;
  }

  function summaryText(a) {
    var e = A.EFFECTS[a.effect];
    return e.label + (a.dir ? '(' + DIR_LABEL[a.dir] + ')' : '') + ' / ' + A.TRIGGERS[a.trigger];
  }

  function renderOrder() {
    var ol = $('#orderList'), slide = curSlide();
    ol.textContent = '';
    if (!slide.anims.length) {
      var li = document.createElement('li');
      li.className = 'empty'; li.textContent = 'まだありません';
      ol.appendChild(li);
      return;
    }
    slide.anims.forEach(function (a, i) {
      var shape = shapeOf(slide, a.spid);
      var li = document.createElement('li');
      var pick = document.createElement('button');
      pick.type = 'button'; pick.className = 'pick';
      pick.textContent = (i + 1) + '. ' + (shape ? shapeLabel(shape) : 'ID ' + a.spid) + ' — ' + summaryText(a);
      if (a.spid === state.selSpid) pick.setAttribute('aria-current', 'true');
      pick.onclick = function () { select(a.spid); };
      li.appendChild(pick);
      [['▲', -1], ['▼', 1]].forEach(function (m) {
        var b = document.createElement('button');
        b.type = 'button'; b.className = 'mini'; b.textContent = m[0];
        b.setAttribute('aria-label', m[1] < 0 ? '上へ' : '下へ');
        var j = i + m[1];
        b.disabled = j < 0 || j >= slide.anims.length;
        b.onclick = function () {
          var t = slide.anims[i]; slide.anims[i] = slide.anims[j]; slide.anims[j] = t;
          touch(); renderAll();
        };
        li.appendChild(b);
      });
      ol.appendChild(li);
    });
  }

  function renderElements() {
    var ul = $('#elList'), slide = curSlide();
    ul.textContent = '';
    slide.shapes.forEach(function (s) {
      var li = document.createElement('li'), b = document.createElement('button');
      b.type = 'button'; b.className = 'pick';
      var n = animIndexes(slide, s.spid).length;
      b.textContent = shapeLabel(s) + (n ? ' ✨' + n : '');
      if (s.spid === state.selSpid) b.setAttribute('aria-current', 'true');
      b.onclick = function () { select(s.spid); };
      li.appendChild(b); ul.appendChild(li);
    });
  }

  function renderSummary() {
    var n = state.slides.filter(function (s) { return s.anims.length || s.transition !== 'none'; }).length;
    $('#summary').textContent = n ? n + ' 枚のスライドに効果を設定済みです。' : 'まだ効果を設定したスライドはありません。';
  }

  /* ---------- プレビュー再生 ---------- */

  function stopPlay() {
    state.playing.forEach(function (an) { try { an.cancel(); } catch (e) { /* ignore */ } });
    state.playing = [];
    var st = $('#stage');
    st.classList.remove('playing');
    Array.prototype.forEach.call(st.querySelectorAll('.shp'), function (el) { el.style.opacity = ''; });
  }

  function keyframes(entry, shape) {
    var b = shape.box || { x: 0, y: 0, w: 0, h: 0 };
    var W = state.size.cx / EMU_PX, H = state.size.cy / EMU_PX;
    var x = b.x / EMU_PX, y = b.y / EMU_PX, w = b.w / EMU_PX, h = b.h / EMU_PX;
    switch (entry.effect) {
      case 'appear': return [{ opacity: 1 }, { opacity: 1 }];
      case 'fade': return [{ opacity: 0 }, { opacity: 1 }];
      case 'zoom': return [{ opacity: 0, scale: 0 }, { opacity: 1, scale: 1 }];
      case 'fly': {
        var from = { left: [-(x + w), 0], right: [W - x, 0], top: [0, -(y + h)], bottom: [0, H - y] }[entry.dir] || [0, 0];
        return [{ opacity: 1, translate: from[0] + 'px ' + from[1] + 'px' }, { opacity: 1, translate: '0px 0px' }];
      }
      case 'wipe': {
        var hid = { left: 'inset(0 100% 0 0)', right: 'inset(0 0 0 100%)', top: 'inset(0 0 100% 0)', bottom: 'inset(100% 0 0 0)' }[entry.dir];
        return [{ opacity: 1, clipPath: hid }, { opacity: 1, clipPath: 'inset(0 0 0 0)' }];
      }
      case 'disappear': return [{ opacity: 0 }, { opacity: 0 }];
      case 'fadeOut': return [{ opacity: 1 }, { opacity: 0 }];
    }
    return [{ opacity: 1 }, { opacity: 1 }];
  }

  function play() {
    stopPlay();
    var slide = curSlide(), st = $('#stage');
    if (!slide.anims.length) { showMsg('このスライドにはアニメーションが設定されていません。'); return; }
    st.classList.add('playing');
    var seen = {};
    slide.anims.forEach(function (a) {
      if (seen[a.spid]) return;
      seen[a.spid] = true;
      if (A.EFFECTS[a.effect].cls === 'entr') {
        var el = st.querySelector('[data-spid="' + a.spid + '"]');
        if (el) el.style.opacity = '0';
      }
    });
    A.schedule(slide.anims, 500).forEach(function (entry) {
      var el = st.querySelector('[data-spid="' + entry.spid + '"]');
      var shape = shapeOf(slide, entry.spid);
      if (!el || !shape) return;
      var dur = entry.effect === 'appear' || entry.effect === 'disappear' ? 1 : entry.dur;
      state.playing.push(el.animate(keyframes(entry, shape), { duration: dur, delay: entry.start, fill: 'forwards', easing: 'ease-out' }));
    });
  }

  /* ---------- 書き出し ---------- */

  function removeChildrenNamed(root, names) {
    Array.prototype.slice.call(root.children).forEach(function (c) {
      var hit = c.namespaceURI === NS.p && names.indexOf(c.localName) >= 0;
      if (!hit && c.namespaceURI === NS.mc && c.localName === 'AlternateContent' &&
          c.getElementsByTagNameNS(NS.p, 'transition').length && names.indexOf('transition') >= 0) hit = true;
      if (hit) root.removeChild(c);
    });
  }

  function applyToSlideXml(slide) {
    var doc = slide.doc, root = doc.documentElement;
    removeChildrenNamed(root, ['transition', 'timing']);
    var tags = {};
    slide.shapes.forEach(function (s) { tags[s.spid] = s.tag; });
    var parts = [A.buildTransition(slide.transition), A.buildTiming(slide.anims, tags)].filter(Boolean);
    var ext = childByName(root, NS.p, 'extLst');
    parts.forEach(function (xml) {
      var node = doc.importNode(parseXml(xml, 'animation').documentElement, true);
      root.insertBefore(node, ext);
    });
    var out = new XMLSerializer().serializeToString(doc);
    return /^<\?xml/.test(out) ? out : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + out;
  }

  function download() {
    if (!state.zip) return;
    showMsg('作成中…');
    state.slides.forEach(function (s) {
      if (s.touched) state.zip.file(s.path, applyToSlideXml(s), { createFolders: false });
    });
    state.zip.generateAsync({
      type: 'blob', compression: 'DEFLATE',
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    }).then(function (blob) {
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = state.fileName.replace(/\.pptx$/i, '') + '_animated.pptx';
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 10000);
      showMsg('ダウンロードしました。PowerPointで開いて「スライドショー」で確認してください。');
    }).catch(function (err) { showMsg('書き出しに失敗しました: ' + err.message, true); });
  }

  /* ---------- 起動 ---------- */

  function showMsg(text, isError) {
    var m = $('#msg');
    m.hidden = false; m.textContent = text;
    m.classList.toggle('error', !!isError);
  }

  function init() {
    var tsel = $('#transition');
    Object.keys(A.TRANSITIONS).forEach(function (k) { tsel.appendChild(option(k, A.TRANSITIONS[k].label)); });
    var be = $('#bulkEffect'), bd = $('#bulkDir');
    ['fade', 'fly', 'wipe', 'zoom', 'appear'].forEach(function (k) { be.appendChild(option(k, A.EFFECTS[k].label)); });
    Object.keys(DIR_LABEL).forEach(function (k) { bd.appendChild(option(k, DIR_LABEL[k])); });
    var syncDir = function () { bd.disabled = !A.EFFECTS[be.value].dirs; };
    be.onchange = syncDir; syncDir();

    var drop = $('#drop'), input = $('#file');
    drop.onclick = function () { input.click(); };
    drop.onkeydown = function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } };
    input.onchange = function () { if (input.files[0]) loadFile(input.files[0]); input.value = ''; };
    ['dragenter', 'dragover'].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('over'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('over'); });
    });
    drop.addEventListener('drop', function (e) { if (e.dataTransfer.files[0]) loadFile(e.dataTransfer.files[0]); });

    tsel.onchange = function () { curSlide().transition = tsel.value; touch(); renderSlides(); renderSummary(); };
    $('#transitionAll').onclick = function () {
      state.slides.forEach(function (s) { s.transition = tsel.value; s.touched = true; });
      renderSummary();
    };
    var bulkArgs = function () {
      var eff = be.value;
      return [eff, A.EFFECTS[eff].dirs ? bd.value : null, $('#bulkTrigger').value];
    };
    $('#bulkSlide').onclick = function () {
      var a = bulkArgs(); bulkApply([curSlide()], a[0], a[1], a[2]); renderAll();
    };
    $('#bulkAll').onclick = function () {
      var a = bulkArgs(); bulkApply(state.slides, a[0], a[1], a[2]); renderAll();
    };
    $('#clearSlide').onclick = function () { curSlide().anims = []; touch(); renderAll(); };
    $('#play').onclick = play;
    $('#stop').onclick = stopPlay;
    $('#download').onclick = download;
    window.addEventListener('resize', function () { if (state.slides.length) fitStage(); });
  }

  init();
})();
