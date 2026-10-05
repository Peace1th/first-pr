/* PowerPoint のアニメーション(p:timing)と画面切り替え(p:transition)の XML を作る純粋ロジック。
 * ブラウザでは window.PptxAnim、Node では module.exports として使える。 */
(function (root) {
  'use strict';

  var NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';

  // 方向 → presetSubtype(PowerPoint の定義: 1=上 2=右 4=下 8=左)
  var SUBTYPE = { top: 1, right: 2, bottom: 4, left: 8 };
  // ワイプのフィルタ名(「下から」= wipe(down) のように PowerPoint は逆向きの名前を使う)
  var WIPE_FILTER = { top: 'wipe(up)', right: 'wipe(right)', bottom: 'wipe(down)', left: 'wipe(left)' };

  // effect: { presetID, cls, dirs } 。dirs が true なら方向を選べる
  var EFFECTS = {
    appear:    { label: '表示',           id: 1,  cls: 'entr', dirs: false },
    fade:      { label: 'フェードイン',   id: 10, cls: 'entr', dirs: false },
    fly:       { label: 'スライドイン',   id: 2,  cls: 'entr', dirs: true, defaultDir: 'left' },
    wipe:      { label: 'ワイプ',         id: 22, cls: 'entr', dirs: true, defaultDir: 'left' },
    zoom:      { label: 'ズームイン',     id: 53, cls: 'entr', dirs: false },
    disappear: { label: '消える',         id: 1,  cls: 'exit', dirs: false },
    fadeOut:   { label: 'フェードアウト', id: 10, cls: 'exit', dirs: false }
  };

  var TRIGGERS = {
    click: 'クリックで',
    with: '前と同時',
    after: '前の後に'
  };

  var TRANSITIONS = {
    none: { label: 'なし' },
    fade: { label: 'フェード', xml: '<p:fade/>' },
    push: { label: 'プッシュ', xml: '<p:push dir="u"/>' },
    wipe: { label: 'ワイプ', xml: '<p:wipe dir="r"/>' },
    zoom: { label: 'ズーム', xml: '<p:zoom/>' }
  };

  function num(v, fallback, min, max) {
    v = Number(v);
    if (!isFinite(v)) v = fallback;
    return Math.min(max, Math.max(min, Math.round(v)));
  }

  /** 1 つの shape ターゲット */
  function tgt(spid) {
    return '<p:tgtEl><p:spTgt spid="' + spid + '"/></p:tgtEl>';
  }

  function setVisibility(ctx, spid, value, delayMs) {
    return '<p:set><p:cBhvr><p:cTn id="' + ctx.nextId() + '" dur="1" fill="hold">' +
      '<p:stCondLst><p:cond delay="' + delayMs + '"/></p:stCondLst></p:cTn>' + tgt(spid) +
      '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>' +
      '<p:to><p:strVal val="' + value + '"/></p:to></p:set>';
  }

  function animEffect(ctx, spid, transition, filter, dur) {
    return '<p:animEffect transition="' + transition + '" filter="' + filter + '">' +
      '<p:cBhvr><p:cTn id="' + ctx.nextId() + '" dur="' + dur + '"/>' + tgt(spid) + '</p:cBhvr></p:animEffect>';
  }

  function animNum(ctx, spid, attr, from, to, dur, additive) {
    var val = function (v) {
      return typeof v === 'number'
        ? '<p:fltVal val="' + v + '"/>'
        : '<p:strVal val="' + v + '"/>';
    };
    return '<p:anim calcmode="lin" valueType="num"><p:cBhvr' + (additive ? ' additive="base"' : '') + '>' +
      '<p:cTn id="' + ctx.nextId() + '" dur="' + dur + '" fill="hold"/>' + tgt(spid) +
      '<p:attrNameLst><p:attrName>' + attr + '</p:attrName></p:attrNameLst></p:cBhvr>' +
      '<p:tavLst><p:tav tm="0"><p:val>' + val(from) + '</p:val></p:tav>' +
      '<p:tav tm="100000"><p:val>' + val(to) + '</p:val></p:tav></p:tavLst></p:anim>';
  }

  /** エフェクト本体(cTn の子要素)を作る */
  function effectBody(ctx, a, dur) {
    var spid = a.spid;
    switch (a.effect) {
      case 'appear':
        return setVisibility(ctx, spid, 'visible', 0);
      case 'fade':
        return setVisibility(ctx, spid, 'visible', 0) + animEffect(ctx, spid, 'in', 'fade', dur);
      case 'wipe':
        return setVisibility(ctx, spid, 'visible', 0) + animEffect(ctx, spid, 'in', WIPE_FILTER[a.dir], dur);
      case 'zoom':
        return setVisibility(ctx, spid, 'visible', 0) +
          animNum(ctx, spid, 'ppt_w', 0, '#ppt_w', dur, false) +
          animNum(ctx, spid, 'ppt_h', 0, '#ppt_h', dur, false) +
          animEffect(ctx, spid, 'in', 'fade', dur);
      case 'fly': {
        var fromX = { left: '0-#ppt_w/2', right: '1+#ppt_w/2' }[a.dir] || '#ppt_x';
        var fromY = { top: '0-#ppt_h/2', bottom: '1+#ppt_h/2' }[a.dir] || '#ppt_y';
        return setVisibility(ctx, spid, 'visible', 0) +
          animNum(ctx, spid, 'ppt_x', fromX, '#ppt_x', dur, true) +
          animNum(ctx, spid, 'ppt_y', fromY, '#ppt_y', dur, true);
      }
      case 'disappear':
        return setVisibility(ctx, spid, 'hidden', 0);
      case 'fadeOut':
        return animEffect(ctx, spid, 'out', 'fade', dur) + setVisibility(ctx, spid, 'hidden', dur);
      default:
        throw new Error('unknown effect: ' + a.effect);
    }
  }

  /** 設定値を正規化。不正な値はデフォルトへ */
  function normalize(a) {
    var def = EFFECTS[a.effect];
    if (!def) throw new Error('unknown effect: ' + a.effect);
    var dir = def.dirs ? (SUBTYPE[a.dir] ? a.dir : def.defaultDir) : null;
    return {
      spid: num(a.spid, 0, 0, 2147483647),
      effect: a.effect,
      dir: dir,
      trigger: TRIGGERS[a.trigger] ? a.trigger : 'click',
      dur: num(a.dur, 500, 50, 60000),
      delay: num(a.delay, 0, 0, 60000)
    };
  }

  /**
   * アニメーション一覧から <p:timing> の XML 文字列を作る。
   * anims: [{spid, effect, dir, trigger, dur, delay}] を再生順に並べたもの
   * shapeTags: { [spid]: 'sp' | 'pic' | ... } 。sp のみ bldLst に登録する
   */
  function buildTiming(anims, shapeTags) {
    var list = anims.map(normalize);
    if (!list.length) return '';
    shapeTags = shapeTags || {};

    var id = 2; // 1 = tmRoot, 2 = mainSeq
    var ctx = { nextId: function () { return ++id; } };

    // クリックごとのグループ → グループ内の「前の後」ごとの束 に分ける
    var clickGroups = [];
    var group = null, bundle = null, groupEnd = 0;
    list.forEach(function (a, i) {
      if (!group || a.trigger === 'click') {
        group = { auto: i === 0 && a.trigger !== 'click', bundles: [] };
        clickGroups.push(group);
        bundle = null; groupEnd = 0;
      }
      if (!bundle || a.trigger === 'after') {
        bundle = { start: bundle ? groupEnd : 0, items: [] };
        group.bundles.push(bundle);
      }
      bundle.items.push(a);
      groupEnd = Math.max(groupEnd, bundle.start + a.delay + a.dur);
    });

    var seen = {};
    var bld = '';
    var body = clickGroups.map(function (g) {
      var outerId = ctx.nextId();
      var inner = g.bundles.map(function (b) {
        var innerId = ctx.nextId();
        var effects = b.items.map(function (a) {
          var def = EFFECTS[a.effect];
          var isSp = shapeTags[a.spid] === 'sp';
          if (isSp && !seen[a.spid]) {
            seen[a.spid] = true;
            bld += '<p:bldP spid="' + a.spid + '" grpId="0" animBg="1"/>';
          }
          // click は必ずグループの先頭、after は必ず束の先頭になるので、トリガーだけで決まる
          var nodeType = { click: 'clickEffect', after: 'afterEffect', with: 'withEffect' }[a.trigger];
          var effId = ctx.nextId();
          return '<p:par><p:cTn id="' + effId + '" presetID="' + def.id + '" presetClass="' + def.cls +
            '" presetSubtype="' + (a.dir ? SUBTYPE[a.dir] : (a.effect === 'zoom' ? 16 : 0)) + '" fill="hold"' +
            (isSp ? ' grpId="0"' : '') + ' nodeType="' + nodeType + '">' +
            '<p:stCondLst><p:cond delay="' + a.delay + '"/></p:stCondLst><p:childTnLst>' +
            effectBody(ctx, a, a.dur) + '</p:childTnLst></p:cTn></p:par>';
        }).join('');
        return '<p:par><p:cTn id="' + innerId + '" fill="hold"><p:stCondLst><p:cond delay="' + b.start +
          '"/></p:stCondLst><p:childTnLst>' + effects + '</p:childTnLst></p:cTn></p:par>';
      }).join('');
      var start = '<p:cond delay="indefinite"/>' +
        (g.auto ? '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>' : '');
      return '<p:par><p:cTn id="' + outerId + '" fill="hold"><p:stCondLst>' + start +
        '</p:stCondLst><p:childTnLst>' + inner + '</p:childTnLst></p:cTn></p:par>';
    }).join('');

    return '<p:timing xmlns:p="' + NS_P + '"><p:tnLst><p:par>' +
      '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>' +
      '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>' +
      body +
      '</p:childTnLst></p:cTn>' +
      '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
      '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>' +
      '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>' +
      (bld ? '<p:bldLst>' + bld + '</p:bldLst>' : '') + '</p:timing>';
  }

  function buildTransition(kind) {
    var t = TRANSITIONS[kind];
    if (!t || !t.xml) return '';
    return '<p:transition xmlns:p="' + NS_P + '" spd="med">' + t.xml + '</p:transition>';
  }

  /**
   * プレビュー用の再生スケジュールを作る(PowerPoint とは別実装なので近似)。
   * 戻り値: [{ spid, effect, dir, start, dur }] start/dur はミリ秒。
   * クリック待ちは clickGapMs の間隔に置き換える。
   */
  function schedule(anims, clickGapMs) {
    var list = anims.map(normalize);
    var out = [];
    var t0 = 0, cursor = 0, groupEnd = 0, bundleStart = 0, first = true;
    list.forEach(function (a) {
      if (a.trigger === 'click' && !first) {
        t0 = t0 + groupEnd + clickGapMs; groupEnd = 0; bundleStart = 0;
      } else if (a.trigger === 'after') {
        bundleStart = groupEnd;
      }
      first = false;
      var start = t0 + bundleStart + a.delay;
      out.push({ spid: a.spid, effect: a.effect, dir: a.dir, start: start, dur: a.dur });
      groupEnd = Math.max(groupEnd, bundleStart + a.delay + a.dur);
    });
    return out;
  }

  var api = {
    EFFECTS: EFFECTS, TRIGGERS: TRIGGERS, TRANSITIONS: TRANSITIONS, SUBTYPE: SUBTYPE,
    buildTiming: buildTiming, buildTransition: buildTransition, schedule: schedule, normalize: normalize
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PptxAnim = api;
})(typeof window !== 'undefined' ? window : globalThis);
