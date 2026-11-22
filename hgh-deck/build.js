// Builds the deck "杭州萧山国际机场" (light, cute theme).
//   node build.js            -> dist/hgh-deck.pptx
// PptxGenJS lays out the slides; lib/fx.js adds transitions, animations, gradients and
// rounded photos to the slide XML afterwards.
const fs = require('fs');
const path = require('path');
const sharp = require('sharp');
const JSZip = require('jszip');
const PptxGenJS = require('pptxgenjs');
const { gradient, decorateSlide } = require('./lib/fx');
const A = require('./lib/art');

const W = 13.333;
const H = 7.5;
const PX = 150;
const TOTAL = 18;
const IMG = path.join(__dirname, 'img');
const OUT_DIR = path.join(__dirname, 'dist');
const OUT_FILE = path.join(OUT_DIR, 'hgh-deck.pptx');
const PHOTOS = Object.fromEntries(require('./photos.json').map((p) => [p.key, p]));

const C = {
  navy: '274060',
  text: '33475F',
  sub: '6E8098',
  mute: '93A3B8',
  line: 'DCE8F3',
  white: 'FFFFFF',
  blue: '4BA3E3',
  blueDeep: '2F6FB5',
  blueSoft: '9AD0F5',
  bluePale: 'E3F2FD',
  pink: 'FF8FAB',
  pinkDeep: 'F0668D',
  pinkPale: 'FFE9EF',
  green: '63C7A0',
  greenDeep: '3FA57F',
  greenPale: 'E2F6EC',
  yellow: 'FFC94D',
  yellowDeep: 'F2A12E',
  yellowPale: 'FFF3D1',
};
const FONT = { cn: 'Microsoft YaHei', num: 'Bahnschrift' };
const SEC = {
  1: { name: '初识机场', color: C.blue, deep: C.blueDeep, pale: C.bluePale },
  2: { name: '成长之路', color: C.pink, deep: C.pinkDeep, pale: C.pinkPale },
  3: { name: '莲叶下的 T4', color: C.green, deep: C.greenDeep, pale: C.greenPale },
  4: { name: '枢纽实力', color: C.blue, deep: C.blueDeep, pale: C.bluePale },
  5: { name: '四通八达', color: C.yellowDeep, deep: C.yellowDeep, pale: C.yellowPale },
  6: { name: '亚运与未来', color: C.pink, deep: C.pinkDeep, pale: C.pinkPale },
};

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: W, height: H });
pptx.layout = 'WIDE';
pptx.title = '杭州萧山国际机场';
pptx.subject = '机场介绍 · 期末汇报';
pptx.author = ' ';
pptx.company = ' ';
pptx.theme = { headFontFace: FONT.cn, bodyFontFace: FONT.cn };

const specs = [];

// Mainland typography: no space between Chinese characters and Latin letters or digits.
const CJK = '\\u2E80-\\u9FFF\\u3000-\\u303F\\uFF00-\\uFFEF“”‘’';
const RE_A = new RegExp(`([${CJK}]) +(?=[A-Za-z0-9+%−-])`, 'g');
const RE_B = new RegExp(`([A-Za-z0-9+%]) +(?=[${CJK}])`, 'g');
const zh = (t) => (typeof t === 'string' ? t.replace(RE_A, '$1').replace(RE_B, '$1') : t.map((r) => ({ ...r, text: zh(r.text) })));
const shadow = () => ({ type: 'outer', color: '4F7FA8', opacity: 0.2, blur: 12, offset: 4, angle: 90 });

// ---------- images ----------------------------------------------------------------------

async function crop(key, w, h, o = {}) {
  const pw = Math.round(w * PX);
  const ph = Math.round(h * PX);
  const file = path.join(IMG, `${key}_${pw}x${ph}.jpg`);
  const src = path.join(__dirname, 'photos', key + '.jpg');
  const meta = await sharp(src).metadata();
  const scale = Math.max(pw / meta.width, ph / meta.height);
  const rw = Math.ceil(meta.width * scale);
  const rh = Math.ceil(meta.height * scale);
  const left = Math.min(rw - pw, Math.round((rw - pw) * (o.fx === undefined ? 0.5 : o.fx)));
  const top = Math.min(rh - ph, Math.round((rh - ph) * (o.fy === undefined ? 0.5 : o.fy)));
  let img = sharp(src).resize(rw, rh).extract({ left, top, width: pw, height: ph });
  if (o.bright || o.sat) img = img.modulate({ brightness: o.bright || 1, saturation: o.sat || 1 });
  await img.jpeg({ quality: 86, mozjpeg: true }).toFile(file);
  return file;
}

const ART = {};
const RATIO = {};
async function makeArt(name, svg, widthPx) {
  const m = /viewBox="0 0 (\d+) (\d+)"/.exec(svg);
  const file = path.join(IMG, `art-${name}.png`);
  await sharp(Buffer.from(svg), { density: Math.min(600, (72 * widthPx) / Number(m[1])) })
    .resize({ width: widthPx })
    .png({ compressionLevel: 9, palette: name.startsWith('bg') ? false : true, quality: 95 })
    .toFile(file);
  ART[name] = file;
  RATIO[name] = Number(m[1]) / Number(m[2]);
}

// ---------- slide API -------------------------------------------------------------------

// PptxGenJS turns a line width of 0 into 1 pt; `width: 0` here means "no outline".
const noLine = (o) => {
  if (!o.line || o.line.width !== 0) return o;
  const { line, ...rest } = o;
  return rest;
};

function newSlide(transition, bg) {
  const slide = pptx.addSlide();
  slide.background = { path: ART[bg || 'bgMain'] };
  const spec = { transition, effects: [], pics: {}, callouts: {}, bulletColor: C.pink };
  specs.push(spec);
  let n = 0;
  const nm = (p) => `${p}${++n}`;
  const api = {
    spec,
    shape(type, o) {
      const objectName = nm('s');
      slide.addShape(type, { ...noLine(o), objectName });
      return objectName;
    },
    text(t, o) {
      const objectName = nm('t');
      slide.addText(zh(t), { fontFace: FONT.cn, color: C.text, margin: 0, ...noLine(o), objectName });
      return objectName;
    },
    image(file, x, y, w, h, o = {}) {
      const objectName = nm(o.round ? 'rp' : 'p');
      const opts = { path: file, x, y, w, h, altText: o.alt || '插图', objectName };
      if (o.rotate) opts.rotate = o.rotate;
      if (o.hyperlink) opts.hyperlink = o.hyperlink;
      slide.addImage(opts);
      if (o.round) spec.pics[objectName] = { adj: o.adj || 6000, border: o.border, borderPt: o.borderPt, shadow: o.shadow !== false };
      return objectName;
    },
    /** A self-drawn illustration; the height follows the drawing's proportions. */
    art(name, x, y, w, o = {}) {
      return api.image(ART[name], x, y, w, w / RATIO[name], { alt: o.alt || '手绘插画', rotate: o.rotate, hyperlink: o.hyperlink });
    },
    chart(type, data, o) {
      const objectName = nm('c');
      slide.addChart(type, data, { ...o, objectName });
      return objectName;
    },
    anim(target, type, delay = 0, dur = 500, extra = {}) {
      spec.effects.push({ target, type, delay, dur, ...extra });
      return target;
    },
    notes(t) {
      slide.addNotes(zh(t));
    },
  };
  return api;
}

// ---------- shared pieces ---------------------------------------------------------------

const cjkWidth = (str, pt) => [...str].reduce((sum, ch) => sum + (/[\u2E80-\uFFFF]/.test(ch) ? pt / 72 : (pt / 72) * 0.56), 0);

function header(s, sec, title, page) {
  const S = SEC[sec];
  const label = `${String(sec).padStart(2, '0')}  ${S.name}`;
  const pill = s.text(label, {
    x: 0.6, y: 0.42, w: cjkWidth(label, 12) + 0.5, h: 0.38, shape: 'roundRect', rectRadius: 0.19,
    fill: { color: S.color }, color: C.white, bold: true, fontSize: 12, align: 'center', valign: 'middle',
  });
  const t = s.text(title, { x: 0.6, y: 0.9, w: 10.8, h: 0.72, fontSize: 30, bold: true, color: C.navy, valign: 'middle' });
  const bar = s.shape('roundRect', { x: 0.62, y: 1.68, w: 0.72, h: 0.075, rectRadius: 0.037, fill: { color: S.color }, line: { color: S.color, width: 0 } });
  const d1 = s.shape('ellipse', { x: 1.42, y: 1.68, w: 0.075, h: 0.075, fill: { color: C.yellow }, line: { color: C.yellow, width: 0 } });
  const d2 = s.shape('ellipse', { x: 1.57, y: 1.68, w: 0.075, h: 0.075, fill: { color: C.blueSoft }, line: { color: C.blueSoft, width: 0 } });
  const cloud = s.art('cloud', 11.75, 0.34, 1.0);
  const num = s.text(
    [
      { text: String(page).padStart(2, '0'), options: { fontFace: FONT.num, fontSize: 14, bold: true, color: C.blueDeep } },
      { text: ` / ${TOTAL}`, options: { fontFace: FONT.num, fontSize: 9, color: C.mute } },
    ],
    { x: 11.75, y: 0.5, w: 1.0, h: 0.34, align: 'center', valign: 'middle' }
  );
  s.anim(pill, 'wipe', 0, 350, { dir: 'left' });
  s.anim(t, 'rise', 120, 500);
  s.anim(bar, 'wipe', 380, 350, { dir: 'left' });
  s.anim(d1, 'zoom', 650, 250);
  s.anim(d2, 'zoom', 760, 250);
  s.anim(cloud, 'path', 0, 6000, { path: 'M 0 0 L -0.012 0 E', autoRev: true, repeat: 'forever', ease: false });
  return { pill, t, num };
}

function footer(s, source) {
  if (source) s.text('资料来源：' + source, { x: 0.6, y: 7.04, w: 9.6, h: 0.28, fontSize: 9, color: C.mute, valign: 'middle' });
  s.text('杭州萧山国际机场 · HGH', { x: 10.2, y: 7.04, w: 2.55, h: 0.28, fontSize: 9, color: C.mute, align: 'right', valign: 'middle' });
}

function card(s, x, y, w, h, o = {}) {
  return s.shape('roundRect', {
    x, y, w, h, rectRadius: o.radius === undefined ? 0.22 : o.radius,
    fill: { color: o.fill || C.white }, line: { color: o.line || C.line, width: o.lineWidth === undefined ? 0.75 : o.lineWidth },
    shadow: o.flat ? undefined : shadow(),
  });
}

// Photos actually placed on slides, for the credits page (filled by `photo`).
const USED_PHOTOS = [];
const PHOTO_NAMES = {
  lotus: 'T4“荷花谷”', hall: 'T4 出发大厅', checkin: 'T4 值机岛', concourse: 'T4 国际出发层指廊', columns: 'T4 出发层大厅', curb: 'T4 出发层车道边',
  aerial: '空中俯瞰机场', aerial2: '跑道与航站区', loong: '长龙航空客机', c919: '国航 C919', b787: '南航波音 787', sfhub: '顺丰物流枢纽',
  sf: '顺丰航空货机', metro19: '19 号线机场站', sign: '地铁站名墙',
};

/** A rounded photo with a small credit label in its bottom-left corner. */
async function photo(s, key, x, y, w, h, o = {}) {
  const meta = PHOTOS[key];
  if (!meta) throw new Error('no credit data for photo ' + key);
  if (!PHOTO_NAMES[key]) throw new Error('no short name for photo ' + key);
  USED_PHOTOS.push(key);
  const file = await crop(key, w, h, o);
  const pic = s.image(file, x, y, w, h, { round: true, adj: o.adj || Math.round(Math.min(18000, (0.2 / Math.min(w, h)) * 100000)), border: C.white, borderPt: o.borderPt || 3, alt: o.caption || meta.title });
  let label;
  if (o.credit !== false) {
    const txt = zh(`${o.caption ? o.caption + ' · ' : ''}摄影：${meta.artist} · ${meta.license}`);
    const lw = Math.min(w - 0.24, cjkWidth(txt, 8) + 0.3);
    label = s.text(txt, {
      x: x + (o.labelDx || 0.12), y: o.labelTop ? y + 0.14 : y + h - 0.36, w: lw, h: 0.24, shape: 'roundRect', rectRadius: 0.12,
      fill: { color: '1F3550', transparency: 35 }, color: C.white, fontSize: 8, align: 'center', valign: 'middle',
    });
  }
  return { pic, label };
}

function tile(s, x, y, w, h, num, unit, label, color, delay, o = {}) {
  const c = card(s, x, y, w, h);
  const t = s.text(
    [
      { text: num, options: { fontFace: FONT.num, fontSize: o.numSize || 30, bold: true, color } },
      { text: unit ? ' ' + unit : '', options: { fontSize: o.unitSize || 13, bold: true, color } },
    ],
    { x: x + 0.22, y: y + 0.14, w: w - 0.4, h: h * 0.48, valign: 'middle' }
  );
  const l = s.text(label, { x: x + 0.22, y: y + 0.14 + h * 0.48, w: w - 0.4, h: h * 0.52 - 0.24, fontSize: o.labelSize || 11, color: C.sub, valign: 'top', lineSpacingMultiple: 1.15 });
  s.anim(c, 'zoom', delay, 450);
  s.anim(t, 'fade', delay + 120, 400);
  s.anim(l, 'fade', delay + 200, 400);
  return c;
}

function bubble(s, text, x, y, w, h, o = {}) {
  const name = bubbleBox(s, text, x, y, w, h, o);
  s.spec.callouts[name] = o.tail || [-45000, 62000];
  return name;
}

function bubbleBox(s, text, x, y, w, h, o) {
  return s.text(text, {
    x, y, w, h, shape: 'wedgeRoundRectCallout',
    fill: { color: o.fill || C.white }, line: { color: o.line || C.blueSoft, width: 1.25 },
    fontSize: o.size || 12, color: o.color || C.navy, bold: !!o.bold, align: o.align || 'center', valign: 'middle', margin: 8,
    lineSpacingMultiple: 1.15, shadow: shadow(),
  });
}

const drift = (s, target, dx, dur) => s.anim(target, 'path', 0, dur, { path: `M 0 0 L ${dx} 0 E`, autoRev: true, repeat: 'forever', ease: false });
const bob = (s, target, dy, dur, delay = 0) => s.anim(target, 'path', delay, dur, { path: `M 0 0 L 0 ${dy} E`, autoRev: true, repeat: 'forever', ease: false });

/** Smooth route through points (inches): returns an SVG path in pixels and a motion path. */
function route(pts, box, end) {
  const seg = (map) => {
    let d = '';
    pts.forEach((p, i) => {
      const [x, y] = map(p);
      if (i === 0) {
        d = `M ${x} ${y}`;
        return;
      }
      const [px, py] = map(pts[i - 1]);
      const mx = (px + x) / 2;
      d += ` C ${mx} ${py} ${mx} ${y} ${x} ${y}`;
    });
    return d;
  };
  const r = (v) => Math.round(v * 10000) / 10000;
  return {
    svg: seg(([x, y]) => [Math.round((x - box.x) * PX), Math.round((y - box.y) * PX)]),
    motion: seg(([x, y]) => [r((x - end[0]) / W), r((y - end[1]) / H)]) + ' E',
  };
}

// ---------- slides ----------------------------------------------------------------------

async function build() {
  fs.rmSync(IMG, { recursive: true, force: true });
  fs.mkdirSync(IMG, { recursive: true });
  fs.mkdirSync(OUT_DIR, { recursive: true });

  await makeArt('bgMain', A.bgMain(), 1920);
  await makeArt('bgCover', A.bgCover(), 1920);
  await makeArt('bgPink', A.bgPink(), 1920);
  await makeArt('plane', A.plane(), 900);
  await makeArt('planeOpen', A.plane('open'), 900);
  await makeArt('planeWink', A.plane('wink'), 900);
  await makeArt('cloud', A.cloud(), 500);
  await makeArt('cloudSmile', A.cloud('smile'), 600);
  await makeArt('lotus', A.lotus(), 800);
  await makeArt('leaf', A.leaf(), 500);
  await makeArt('bud', A.bud(), 240);
  await makeArt('suitcase', A.suitcase(), 440);
  await makeArt('tower', A.tower(), 400);
  await makeArt('star', A.star(), 160);
  await makeArt('starPink', A.star(A.P.pink), 160);
  await makeArt('heart', A.heart(), 160);
  await makeArt('sun', A.sun(), 500);
  const iconSet = {
    metro: A.P.blue, train: A.P.pink, bus: A.P.green, car: A.P.yellow, parking: A.P.blueDeep, plane: A.P.blue, globe: A.P.green,
    box: A.P.yellow, pin: A.P.pink, seat: A.P.pink, flag: A.P.green, medal: A.P.pink, leafIcon: A.P.green, clock: A.P.blue, ruler: A.P.yellow,
  };
  for (const [name, color] of Object.entries(iconSet)) await makeArt('i-' + name, A.icon(name, color), 240);

  // ===== 1 封面 ==========================================================================
  {
    const s = newSlide('fade', 'bgCover');
    const sun = s.art('sun', 11.6, 0.28, 1.2);
    const c1 = s.art('cloudSmile', 9.35, 0.3, 1.6);
    const c2 = s.art('cloud', 0.5, 0.3, 1.25);
    const trail = s.image(await trailArt('cover', 420, 110, 'M 8 96 Q 190 6 412 58'), 3.75, 0.5, 2.8, 0.733, { alt: '航迹' });
    const plane = s.art('plane', 6.35, 0.42, 1.9);

    const tag = s.text('期末汇报  ·  带你认识一座机场', {
      x: 0.85, y: 1.12, w: 3.45, h: 0.42, shape: 'roundRect', rectRadius: 0.21,
      fill: { color: C.blue }, color: C.white, bold: true, fontSize: 13, align: 'center', valign: 'middle',
    });
    const title = s.text('杭州萧山国际机场', { x: 0.82, y: 1.72, w: 7.3, h: 1.1, fontSize: 54, bold: true, color: C.navy, valign: 'middle' });
    const sub = s.text('从“接天莲叶”出发，认识浙江第一空中门户', { x: 0.86, y: 2.86, w: 7.2, h: 0.5, fontSize: 20, bold: true, color: C.blueDeep, valign: 'middle' });
    const en = s.text('HANGZHOU XIAOSHAN INTERNATIONAL AIRPORT  ·  HGH', { x: 0.86, y: 3.4, w: 7.2, h: 0.32, fontSize: 11.5, color: C.sub, charSpacing: 2, valign: 'middle' });

    const box = card(s, 0.85, 4.02, 6.3, 2.5, { radius: 0.3 });
    const rows = [
      ['汇报人', C.blueDeep, C.bluePale],
      ['制作人', C.pinkDeep, C.pinkPale],
      ['小组成员', C.greenDeep, C.greenPale],
    ];
    const fields = [];
    rows.forEach(([label, col, pale], i) => {
      const y = 4.3 + i * 0.7;
      const lab = s.text(label, {
        x: 1.15, y, w: 1.45, h: 0.46, shape: 'roundRect', rectRadius: 0.23,
        fill: { color: pale }, color: col, bold: true, fontSize: 15, align: 'center', valign: 'middle',
      });
      const field = s.text(' ', { x: 2.8, y, w: 4.05, h: 0.46, fontSize: 16, color: C.navy, valign: 'middle' });
      const line = s.shape('line', { x: 2.8, y: y + 0.5, w: 4.05, h: 0, line: { color: C.blueSoft, width: 1.25, dashType: 'dash' } });
      fields.push([lab, field, line]);
    });

    const { pic, label } = await photo(s, 'lotus', 8.35, 1.5, 4.25, 4.75, { adj: 14000, borderPt: 5, caption: 'T4 航站楼“荷花谷”', labelTop: true, labelDx: 0.2 });
    const leaf = s.art('leaf', 11.25, 5.7, 1.6);
    const lotus = s.art('lotus', 7.45, 5.42, 1.95);
    const bud = s.art('bud', 12.35, 4.95, 0.42);
    const st1 = s.art('star', 7.25, 2.05, 0.36);
    const st2 = s.art('starPink', 0.42, 1.78, 0.26);
    const st3 = s.art('star', 7.5, 3.45, 0.22);

    s.anim(tag, 'wipe', 200, 450, { dir: 'left' });
    s.anim(title, 'zoom', 450, 700);
    s.anim(sub, 'rise', 900, 550);
    s.anim(en, 'fade', 1250, 500);
    s.anim(pic, 'circle', 500, 900);
    s.anim(label, 'fade', 1400, 400);
    s.anim(box, 'rise', 1500, 550);
    fields.forEach(([lab, , line], i) => {
      s.anim(lab, 'zoom', 1750 + i * 180, 350);
      s.anim(line, 'wipe', 1850 + i * 180, 400, { dir: 'left' });
    });
    s.anim(lotus, 'zoom', 1500, 600);
    s.anim(leaf, 'zoom', 1700, 600);
    s.anim(bud, 'rise', 1900, 600);
    s.anim(trail, 'wipe', 700, 1300, { dir: 'left' });
    s.anim(plane, 'fade', 600, 300);
    s.anim(plane, 'path', 600, 1900, { path: 'M -0.34 0.07 C -0.22 -0.02 -0.1 -0.05 0 0 E' });
    bob(s, plane, -0.014, 1600, 2600);
    [st1, st2, st3].forEach((st, i) => {
      s.anim(st, 'zoom', 1300 + i * 250, 400);
      s.anim(st, 'pulse', 1800 + i * 300, 900, { by: 1.35, repeat: 'forever' });
    });
    s.anim(sun, 'pulse', 0, 1800, { by: 1.06, repeat: 'forever' });
    drift(s, c1, -0.03, 7000);
    drift(s, c2, 0.035, 8000);
    s.notes(
      '【开场】各位老师、同学大家好，我们小组今天要介绍的是杭州萧山国际机场。\n' +
        '它是浙江省的第一空中门户，2025 年旅客吞吐量突破 5000 万人次，排名全国第 7。\n' +
        '封面这张照片是 T4 航站楼里的“荷花谷”，待会儿会重点讲到它。\n' +
        '（提示：封面上“汇报人、制作人、小组成员”三处虚线上方可以直接点击填写。）'
    );
  }

  // ===== 2 目录 ==========================================================================
  {
    const s = newSlide('airplane');
    const pill = s.text('CONTENTS  ·  目录', {
      x: 0.6, y: 0.42, w: 2.2, h: 0.38, shape: 'roundRect', rectRadius: 0.19,
      fill: { color: C.pink }, color: C.white, bold: true, fontSize: 12, align: 'center', valign: 'middle',
    });
    const title = s.text('今天的航线图：一共飞六站', { x: 0.6, y: 0.9, w: 7.6, h: 0.72, fontSize: 30, bold: true, color: C.navy, valign: 'middle' });
    s.anim(pill, 'wipe', 0, 350, { dir: 'left' });
    s.anim(title, 'rise', 120, 500);

    const stops = [
      ['初识机场', '一张机场名片', 3, C.blue],
      ['成长之路', '大事记 · 客流曲线', 4, C.pink],
      ['莲叶下的 T4', '设计意象 · 三个看点', 6, C.green],
      ['枢纽实力', '跑道 · 航线 · 排名 · 货运', 8, C.yellowDeep],
      ['四通八达', '交通中心 · 轨道交通', 12, C.blue],
      ['亚运与未来', '亚运保障 · 发展蓝图', 14, C.pink],
    ];
    const pos = stops.map((_, i) => [1.6 + i * 2.05, i % 2 === 0 ? 4.9 : 3.3]);
    const hotspots = [];
    const end = [12.15, 1.2];
    const boxTop = 0.6;
    const r = route([[-0.5, 5.9], ...pos, end], { x: 0, y: boxTop }, end);
    const trail = s.image(await trailArt('route', Math.round(W * PX), Math.round(5.6 * PX), r.svg, A.P.blueSoft, 6, '3 20'), 0, boxTop, W, 5.6, { alt: '航线' });
    s.anim(trail, 'wipe', 300, 3200, { dir: 'left' });

    stops.forEach(([name, desc, target, color], i) => {
      const [cx, cy] = pos[i];
      const up = i % 2 === 1;
      const link = { slide: target, tooltip: zh('跳转到：' + name) };
      const dot = s.text(String(i + 1).padStart(2, '0'), {
        x: cx - 0.46, y: cy - 0.46, w: 0.92, h: 0.92, shape: 'ellipse', fill: { color }, line: { color: C.white, width: 4 },
        fontFace: FONT.num, fontSize: 24, bold: true, color: C.white, align: 'center', valign: 'middle', shadow: shadow(),
      });
      const ty = up ? cy - 1.42 : cy + 0.6;
      const t = s.text(name, { x: cx - 1.1, y: ty, w: 2.2, h: 0.42, fontSize: 17, bold: true, color: C.navy, align: 'center', valign: 'middle' });
      const d = s.text(desc, { x: cx - 1.15, y: ty + 0.42, w: 2.3, h: 0.34, fontSize: 10.5, color: C.sub, align: 'center', valign: 'middle' });
      const at = 600 + i * 480;
      s.anim(dot, 'zoom', at, 400);
      s.anim(t, 'rise', at + 120, 400);
      s.anim(d, 'fade', at + 240, 400);
      hotspots.push([cx - 1.15, up ? cy - 1.45 : cy - 0.5, link]);
    });

    const plane = s.art('planeOpen', end[0] - 0.75, end[1] - 0.48, 1.5);
    s.anim(plane, 'fade', 300, 250);
    s.anim(plane, 'path', 300, 3600, { path: r.motion, ease: false });
    const say = bubble(s, '我是向导“小荷号”，点一下站名就能直接飞过去～', 8.05, 0.42, 3.1, 0.82, { size: 11.5, tail: [58000, 30000] });
    s.anim(say, 'zoom', 3900, 450);
    bob(s, plane, -0.014, 1500, 4000);
    const leaf = s.art('leaf', 0.35, 6.2, 1.3);
    const lotus = s.art('lotus', 0.95, 5.95, 1.25);
    s.anim(leaf, 'zoom', 800, 500);
    s.anim(lotus, 'zoom', 1000, 500);
    s.text('“小荷号”为本组原创卡通形象', { x: 9.4, y: 7.04, w: 3.35, h: 0.28, fontSize: 9, color: C.mute, align: 'right', valign: 'middle' });
    // invisible click areas on top of every stop (number, name and description)
    hotspots.forEach(([x, y, link]) => s.shape('rect', { x, y, w: 2.3, h: 1.95, fill: { color: C.white, transparency: 100 }, hyperlink: link }));
    s.notes(
      '【目录】今天的汇报像一次航班，一共飞六站：先认识机场的基本情况，再看它 25 年的成长；' +
        '第三站走进最有特色的 T4 航站楼；第四站用数据看它的枢纽实力；第五站讲怎么去机场；最后是亚运会和未来规划。\n' +
        '（放映时点击站名或圆圈可以直接跳转到对应部分。）'
    );
  }

  // ===== 3 机场名片 ======================================================================
  {
    const s = newSlide('pageCurl');
    header(s, 1, '一张登机牌，认识萧山机场', 3);
    const { pic, label } = await photo(s, 'aerial', 0.6, 2.05, 5.45, 3.75, { caption: '空中俯瞰萧山机场', fy: 0.45 });
    const sticker = s.text('浙江省第一空中门户', {
      x: 0.42, y: 1.9, w: 2.55, h: 0.44, shape: 'roundRect', rectRadius: 0.22, rotate: -5,
      fill: { color: C.pink }, color: C.white, bold: true, fontSize: 13.5, align: 'center', valign: 'middle', shadow: shadow(),
    });
    s.anim(pic, 'zoom', 300, 650);
    s.anim(label, 'fade', 900, 400);
    s.anim(sticker, 'zoom', 1000, 400);

    const bx = 6.45;
    const bw = 6.3;
    const pass = card(s, bx, 2.05, bw, 3.75, { radius: 0.28 });
    const bandTop = s.shape('roundRect', { x: bx, y: 2.05, w: bw, h: 0.72, rectRadius: 0.28, fill: { color: C.blue }, line: { color: C.blue, width: 0 } });
    const bandFill = s.shape('rect', { x: bx, y: 2.45, w: bw, h: 0.32, fill: { color: C.blue }, line: { color: C.blue, width: 0 } });
    const bandText = s.text('BOARDING PASS  ·  登机牌', { x: bx + 0.3, y: 2.07, w: 3.6, h: 0.68, fontSize: 14, bold: true, color: C.white, charSpacing: 1, valign: 'middle' });
    const bandRight = s.text('航班 HGH-2025', { x: bx + bw - 2.5, y: 2.07, w: 2.2, h: 0.68, fontSize: 11, color: C.white, align: 'right', valign: 'middle' });
    const code1 = s.text('HGH', { x: bx + 0.3, y: 2.88, w: 2.1, h: 0.82, fontFace: FONT.num, fontSize: 46, bold: true, color: C.navy, valign: 'middle' });
    const lab1 = s.text('IATA 三字代码', { x: bx + 0.3, y: 3.68, w: 2.1, h: 0.26, fontSize: 10, color: C.sub, valign: 'middle' });
    const dash = s.shape('line', { x: bx + 2.3, y: 3.3, w: 1.55, h: 0, line: { color: C.blueSoft, width: 1.5, dashType: 'dash' } });
    const mini = s.art('plane', bx + 2.72, 3.02, 0.82);
    const code2 = s.text('ZSHC', { x: bx + 3.95, y: 2.88, w: 2.1, h: 0.82, fontFace: FONT.num, fontSize: 46, bold: true, color: C.navy, align: 'right', valign: 'middle' });
    const lab2 = s.text('ICAO 四字代码', { x: bx + 3.95, y: 3.68, w: 2.1, h: 0.26, fontSize: 10, color: C.sub, align: 'right', valign: 'middle' });
    const perf = s.shape('line', { x: bx + 0.25, y: 4.13, w: bw - 0.5, h: 0, line: { color: C.line, width: 1.5, dashType: 'sysDash' } });
    const n1 = s.shape('ellipse', { x: bx - 0.15, y: 3.98, w: 0.3, h: 0.3, fill: { color: 'F4FAFF' }, line: { color: C.line, width: 0.75 } });
    const n2 = s.shape('ellipse', { x: bx + bw - 0.15, y: 3.98, w: 0.3, h: 0.3, fill: { color: 'F4FAFF' }, line: { color: C.line, width: 0.75 } });
    s.anim(pass, 'rise', 500, 550);
    [bandTop, bandFill].forEach((b) => s.anim(b, 'fade', 650, 350));
    s.anim(bandText, 'fade', 800, 350);
    s.anim(bandRight, 'fade', 800, 350);
    s.anim(code1, 'zoom', 950, 450);
    s.anim(lab1, 'fade', 1100, 350);
    s.anim(dash, 'wipe', 1150, 500, { dir: 'left' });
    s.anim(mini, 'fade', 1150, 250);
    s.anim(mini, 'path', 1150, 900, { path: 'M -0.06 0 L 0 0 E' });
    s.anim(code2, 'zoom', 1500, 450);
    s.anim(lab2, 'fade', 1650, 350);
    [perf, n1, n2].forEach((p) => s.anim(p, 'fade', 1750, 300));

    const cells = [
      ['飞行区等级', '4F'],
      ['建成通航', '2000 年 12 月 28 日'],
      ['所在地', '杭州市萧山区'],
      ['距市中心', '约 27 公里'],
      ['跑道', '2 条'],
      ['在用航站楼', '3 号、4 号航站楼'],
    ];
    cells.forEach(([k, v], i) => {
      const x = bx + 0.3 + (i % 3) * 2.0;
      const y = 4.32 + Math.floor(i / 3) * 0.72;
      const kk = s.text(k, { x, y, w: 1.9, h: 0.24, fontSize: 9.5, color: C.sub, valign: 'middle' });
      const vv = s.text(v, { x, y: y + 0.24, w: 1.95, h: 0.36, fontSize: 13.5, bold: true, color: C.navy, valign: 'middle' });
      s.anim(kk, 'fade', 1900 + i * 110, 300);
      s.anim(vv, 'rise', 1950 + i * 110, 350);
    });

    const guide = s.art('planeWink', 0.65, 6.02, 1.25);
    const say = bubble(s, '一句话认识它：2025 年旅客吞吐量 5045.9 万人次、全国第 7，是浙江省的第一空中门户。', 2.2, 6.0, 8.9, 0.62, { size: 13, bold: true, tail: [-52500, 0] });
    s.anim(guide, 'fly', 2500, 500, { dir: 'left' });
    s.anim(say, 'zoom', 2850, 450);
    bob(s, guide, -0.012, 1500, 3200);
    footer(s, '杭州萧山国际机场官网（机场概况）；中国民航局《2025 年全国民用运输机场生产统计公报》');
    s.notes(
      '【机场名片】我们把基本信息做成了一张登机牌。\n' +
        '萧山机场的三字代码是 HGH，四字代码 ZSHC，飞行区等级 4F，是我国民用机场的最高等级。\n' +
        '它位于杭州市萧山区，距离市中心约 27 公里，2000 年 12 月 28 日建成通航，目前有两条跑道。\n' +
        '2025 年它的旅客吞吐量是 5045.9 万人次，排名全国第 7。'
    );
  }

  // ===== 4 成长之路 ======================================================================
  {
    const s = newSlide('gallery');
    header(s, 2, '25 年：从一条跑道到 5000 万人次', 4);
    const items = [
      ['2000.12.28', '建成通航', '民航航班从笕桥机场转场而来'],
      ['2004.3', '成为国际机场', '正式对外国籍飞机开放'],
      ['2007', '首破 1000 万', '全年旅客 1173.0 万人次'],
      ['2012.12.30', '二期工程投运', '第二跑道、T3 航站楼启用'],
      ['2016.12', '突破 3000 万', '三年前刚突破 2000 万'],
      ['2019.12.30', '突破 4000 万', '全年旅客 4010.8 万人次'],
      ['2022.9.22', '三期工程投运', 'T4 航站楼、交通中心启用'],
      ['2023.9', '亚运空中门户', '杭州亚运会官方出入境口岸'],
      ['2025.12.28', '突破 5000 万', '正好是通航 25 周年纪念日'],
    ];
    const colors = [C.blue, C.pink, C.green, C.yellowDeep];
    const lineY = 4.12;
    const bar = s.shape('roundRect', {
      x: 0.45, y: lineY - 0.06, w: 12.43, h: 0.12, rectRadius: 0.06,
      fill: { color: gradient(0, [{ pos: 0, color: C.blueSoft }, { pos: 55, color: C.blue }, { pos: 100, color: C.pink }]) }, line: { color: C.white, width: 0 },
    });
    s.anim(bar, 'wipe', 300, 4200, { dir: 'left' });
    const cw = 2.2;
    const ch = 1.28;
    items.forEach(([year, name, desc], i) => {
      const cx = 1.25 + i * 1.355;
      const top = i % 2 === 0;
      const color = colors[i % 4];
      const x = Math.max(0.45, Math.min(W - 0.45 - cw, cx - cw / 2));
      const y = top ? 2.15 : 4.85;
      const stem = s.shape('line', { x: cx, y: top ? y + ch : lineY + 0.17, w: 0, h: top ? lineY - 0.17 - (y + ch) : y - (lineY + 0.17), line: { color, width: 1.25, dashType: 'sysDot' } });
      const node = s.shape('ellipse', { x: cx - 0.17, y: lineY - 0.17, w: 0.34, h: 0.34, fill: { color }, line: { color: C.white, width: 3 }, shadow: shadow() });
      const c = card(s, x, y, cw, ch, { radius: 0.18 });
      const yr = s.text(year, { x: x + 0.16, y: y + 0.1, w: cw - 0.3, h: 0.34, fontFace: FONT.num, fontSize: 15, bold: true, color, valign: 'middle' });
      const nm = s.text(name, { x: x + 0.16, y: y + 0.44, w: cw - 0.3, h: 0.34, fontSize: 13.5, bold: true, color: C.navy, valign: 'middle' });
      const ds = s.text(desc, { x: x + 0.16, y: y + 0.8, w: cw - 0.2, h: 0.34, fontSize: 10, color: C.sub, valign: 'middle' });
      const at = 500 + i * 460;
      s.anim(node, 'zoom', at, 300);
      s.anim(stem, 'fade', at + 100, 300);
      s.anim(c, 'rise', at + 150, 400);
      [yr, nm, ds].forEach((t) => s.anim(t, 'fade', at + 300, 350));
    });
    const plane = s.art('plane', 0.05, lineY - 0.62, 0.9);
    s.anim(plane, 'fade', 300, 200);
    s.anim(plane, 'path', 300, 4600, { path: 'M 0 0 L 1.02 0 E', ease: false });
    footer(s, '杭州萧山国际机场官网（发展历程）；人民网 2004-03-18；浙江在线、新华社 2025-12-28；中国民航局历年统计公报');
    s.notes(
      '【成长之路】萧山机场 2000 年 12 月 28 日建成通航，杭州的民航航班从老的笕桥机场搬到了这里。\n' +
        '2004 年正式对外籍飞机开放，成为国际机场；2007 年旅客量第一次超过 1000 万。\n' +
        '2012 年二期工程投运，有了第二条跑道和 T3 航站楼；2016 年突破 3000 万，2019 年突破 4000 万。\n' +
        '2022 年 9 月 22 日三期工程投运，T4 航站楼启用；2023 年它是杭州亚运会的官方出入境口岸。\n' +
        '2025 年 12 月 28 日，正好是通航 25 周年这一天，年旅客量突破了 5000 万人次。'
    );
  }

  // ===== 5 客流曲线 ======================================================================
  {
    const s = newSlide('prism');
    header(s, 2, '一条向上的客流曲线', 5);
    const years = ['2001', '2005', '2007', '2010', '2013', '2016', '2019', '2020', '2021', '2022', '2023', '2024', '2025'];
    const vals = [298.1, 809.3, 1173.0, 1706.9, 2211.4, 3159.5, 4010.8, 2822.4, 2816.4, 2003.8, 4117.0, 4805.4, 5045.9];
    const barColors = years.map((y) => (['2020', '2021', '2022'].includes(y) ? 'B9C7D6' : y === '2025' ? C.pink : C.blue));
    const box = card(s, 0.6, 2.0, 8.2, 4.75);
    const cap = s.text('年旅客吞吐量（万人次）', { x: 0.9, y: 2.14, w: 4, h: 0.32, fontSize: 12, bold: true, color: C.navy, valign: 'middle' });
    const legend = s.text(
      [
        { text: '■ ', options: { color: 'B9C7D6' } },
        { text: '2020–2022 年受疫情影响   ', options: { color: C.sub } },
        { text: '■ ', options: { color: C.pink } },
        { text: '历史新高', options: { color: C.sub } },
      ],
      { x: 4.6, y: 2.14, w: 4.0, h: 0.32, fontSize: 10, align: 'right', valign: 'middle' }
    );
    const chart = s.chart(pptx.ChartType.bar, [{ name: '旅客吞吐量（万人次）', labels: years, values: vals }], {
      x: 0.75, y: 2.5, w: 7.9, h: 4.15, barDir: 'col', barGapWidthPct: 45, chartColors: barColors,
      showLegend: false, showTitle: false, showValue: true, dataLabelFontSize: 9, dataLabelColor: C.text, dataLabelFormatCode: '0.0', dataLabelPosition: 'outEnd',
      dataLabelFontFace: FONT.num, catAxisLabelFontSize: 10, catAxisLabelColor: C.sub, catAxisLabelFontFace: FONT.num, catAxisLineShow: true,
      valAxisHidden: true, valAxisMaxVal: 5600, valAxisMinVal: 0, valGridLine: { color: 'E8F0F7', width: 0.75 }, catGridLine: { style: 'none' },
    });
    s.anim(box, 'rise', 300, 500);
    s.anim(cap, 'fade', 600, 350);
    s.anim(legend, 'fade', 700, 350);
    s.anim(chart, 'wipe', 800, 1500, { dir: 'down' });

    const notes = [
      ['约 17 倍', '2001 年 298.1 万人次，2025 年 5045.9 万人次', C.blue, 'i-clock'],
      ['V 形反弹', '2022 年跌到 2003.8 万，2023 年回到 4117.0 万，超过 2019 年', C.yellowDeep, 'i-ruler'],
      ['首破 5000 万', '2025 年同比增长 5.0%，全国排名第 7', C.pinkDeep, 'i-medal'],
    ];
    notes.forEach(([big, small, color, ic], i) => {
      const y = 2.0 + i * 1.62;
      const c = card(s, 9.05, y, 3.7, 1.47);
      const icn = s.art(ic, 9.25, y + 0.2, 0.62);
      const b = s.text(big, { x: 10.02, y: y + 0.16, w: 2.6, h: 0.5, fontSize: 21, bold: true, color, valign: 'middle' });
      const t = s.text(small, { x: 10.02, y: y + 0.68, w: 2.58, h: 0.66, fontSize: 10.5, color: C.sub, valign: 'top', lineSpacingMultiple: 1.2 });
      const at = 1500 + i * 400;
      s.anim(c, 'slide', at, 450, { offset: 0.05 });
      s.anim(icn, 'zoom', at + 150, 350);
      s.anim(b, 'fade', at + 200, 350);
      s.anim(t, 'fade', at + 300, 350);
    });
    footer(s, '中国民航局历年《民航机场生产统计公报》（2001–2025 年）');
    s.notes(
      '【客流曲线】这张图是萧山机场历年的旅客吞吐量。\n' +
        '2001 年通航第一个完整年度只有 298.1 万人次，到 2025 年是 5045.9 万人次，大约是当年的 17 倍。\n' +
        '灰色的三年是受疫情影响的 2020 到 2022 年，2022 年跌到 2003.8 万；2023 年一年就回到 4117.0 万，超过了疫情前的 2019 年。\n' +
        '2025 年首次突破 5000 万，同比增长 5.0%，全国排名第 7。'
    );
  }

  // ===== 6 T4 ============================================================================
  {
    const s = newSlide('curtains', 'bgPink');
    header(s, 3, 'T4 航站楼：把西湖的荷花“种”进机场', 6);
    const { pic, label } = await photo(s, 'hall', 0.6, 2.05, 7.3, 4.7, { caption: 'T4 出发大厅的荷叶柱', fy: 0.6, labelDx: 1.45 });
    s.anim(pic, 'splitH', 300, 900);
    s.anim(pic, 'grow', 1200, 9000, { by: 1.03 });
    s.anim(label, 'fade', 1200, 400);
    const quote = card(s, 8.2, 2.05, 4.55, 1.5, { fill: C.pinkPale, line: 'FFD3DE' });
    const q1 = s.text('接天莲叶 · 出水芙蓉', { x: 8.45, y: 2.17, w: 4.1, h: 0.56, fontSize: 23, bold: true, color: C.pinkDeep, valign: 'middle' });
    const q2 = s.text('设计意象取自西湖荷花；2022 年 9 月 22 日随三期工程正式投运。', { x: 8.45, y: 2.75, w: 4.1, h: 0.68, fontSize: 11.5, color: C.text, valign: 'top', lineSpacingMultiple: 1.2 });
    s.anim(quote, 'rise', 700, 500);
    s.anim(q1, 'wipe', 950, 600, { dir: 'left' });
    s.anim(q2, 'fade', 1300, 450);
    const tiles = [
      ['71.9', '万㎡', '总建筑面积', C.greenDeep],
      ['40', '根', '荷叶柱：室内 32、室外 8', C.pinkDeep],
      ['53.5', '米', '荷花谷高度（−9 米层到屋顶）', C.blueDeep],
      ['5000', '万人次', '设计年旅客吞吐量', C.yellowDeep],
    ];
    tiles.forEach(([n, u, l, col], i) => {
      tile(s, 8.2 + (i % 2) * 2.33, 3.75 + Math.floor(i / 2) * 1.55, 2.22, 1.42, n, u, l, col, 1500 + i * 220, { numSize: 28, unitSize: 12, labelSize: 10 });
    });
    const lotus = s.art('lotus', 0.3, 5.95, 1.55);
    s.anim(lotus, 'zoom', 1400, 600);
    s.anim(lotus, 'pulse', 2200, 1400, { by: 1.06, repeat: 'forever' });
    footer(s, '杭州萧山国际机场官网（2022-09-22 三期投运通稿、2026 年招商公告）；浙江在线 2022-04-01');
    s.notes(
      '【T4 航站楼】第三站是最有特色的 T4 航站楼。\n' +
        '它的设计意象取自西湖荷花，叫“接天莲叶、出水芙蓉”，2022 年 9 月 22 日随三期工程投运。\n' +
        '总建筑面积约 71.9 万平方米；整座楼有 40 根“荷叶柱”，室内 32 根、室外 8 根，就是照片里这种像荷叶一样撑开的柱子。\n' +
        '中央的荷花谷从地下 9 米层一直通到屋顶，高 53.5 米。T4 的设计年旅客吞吐量是 5000 万人次。'
    );
  }

  // ===== 7 T4 三个看点 ===================================================================
  {
    const s = newSlide('ferris', 'bgPink');
    header(s, 3, '走进 T4：三个值得抬头看的地方', 7);
    const cards = [
      ['lotus', '荷花谷', '三根“荷花杆”自下而上，在顶部绽开成一朵“荷花”；它是视觉焦点，也是连接各楼层的换乘中庭。', 'T4“荷花谷”', C.pinkDeep, 0.5],
      ['checkin', '莲下泛舟', '6 座值机岛散落在“荷叶”之下；柱顶的菱形天窗把自然光引进来，白天不用开太多灯。', 'T4 值机岛', C.greenDeep, 0.5],
      ['concourse', '省心的小细节', '行李称重带与地面“0 高差”，箱子一推就上；T4 还获得了三星级绿色建筑设计标识。', 'T4 国际出发层指廊', C.blueDeep, 0.5],
    ];
    for (let i = 0; i < cards.length; i++) {
      const [key, name, desc, cap, color, fx] = cards[i];
      const x = 0.6 + i * 4.11;
      const y = 2.05;
      const c = card(s, x, y, 3.93, 4.7, { radius: 0.26 });
      const { pic, label } = await photo(s, key, x + 0.16, y + 0.16, 3.61, 2.55, { caption: cap, fx, shadow: false });
      const badge = s.text(String(i + 1), {
        x: x - 0.12, y: y - 0.14, w: 0.56, h: 0.56, shape: 'ellipse', fill: { color }, line: { color: C.white, width: 3 },
        fontFace: FONT.num, fontSize: 18, bold: true, color: C.white, align: 'center', valign: 'middle', shadow: shadow(),
      });
      const t = s.text(name, { x: x + 0.26, y: y + 2.86, w: 3.4, h: 0.44, fontSize: 18, bold: true, color, valign: 'middle' });
      const d = s.text(desc, { x: x + 0.26, y: y + 3.34, w: 3.42, h: 1.2, fontSize: 11.5, color: C.text, valign: 'top', lineSpacingMultiple: 1.3 });
      const at = 350 + i * 500;
      s.anim(c, 'rise', at, 500);
      s.anim(pic, 'fade', at + 200, 500);
      s.anim(label, 'fade', at + 500, 350);
      s.anim(badge, 'zoom', at + 350, 350);
      s.anim(t, 'wipe', at + 400, 400, { dir: 'left' });
      s.anim(d, 'fade', at + 550, 400);
    }
    const bag = s.art('suitcase', 12.2, 6.0, 0.78);
    s.anim(bag, 'fly', 2000, 600, { dir: 'right' });
    bob(s, bag, -0.012, 1400, 2700);
    footer(s, '浙江在线 2022-04-01；浙江省交通运输厅 2022-03-14；杭州萧山国际机场官网 2022-09-22');
    s.notes(
      '【三个看点】走进 T4，有三个地方值得抬头看。\n' +
        '第一是荷花谷：三根“荷花杆”自下而上，在顶部绽开成一朵荷花，它同时是连接各个楼层的换乘中庭。\n' +
        '第二是“莲下泛舟”：6 座值机岛散落在荷叶柱下面，柱顶有菱形天窗把自然光引进来，白天比较省电。\n' +
        '第三是一些省心的小细节：比如行李称重带和地面是“0 高差”的，行李箱一推就能上去；T4 还获得了三星级绿色建筑设计标识。'
    );
  }

  // ===== 8 飞行区 ========================================================================
  {
    const s = newSlide('conveyor');
    header(s, 4, '两条跑道，托起一座 4F 机场', 8);
    const field = card(s, 0.6, 2.05, 7.7, 4.7, { fill: 'E4F5EA', line: 'CDEBD9', radius: 0.3 });
    s.anim(field, 'fade', 300, 450);
    const runway = (x, y, w, name, len, labelAbove, at) => {
      const r = s.shape('roundRect', { x, y, w, h: 0.44, rectRadius: 0.08, fill: { color: '5F7084' }, line: { color: '5F7084', width: 0 } });
      const c = s.shape('line', { x: x + 0.3, y: y + 0.22, w: w - 0.6, h: 0, line: { color: C.white, width: 2, dashType: 'dash' } });
      const ly = labelAbove ? y - 0.42 : y + 0.5;
      const n = s.text(
        [
          { text: name + '  ', options: { fontSize: 13, bold: true, color: C.navy } },
          { text: len, options: { fontFace: FONT.num, fontSize: 15, bold: true, color: C.greenDeep } },
        ],
        { x: x + 0.1, y: ly, w: 3.6, h: 0.36, valign: 'middle' }
      );
      s.anim(r, 'wipe', at, 800, { dir: 'left' });
      s.anim(c, 'wipe', at + 300, 800, { dir: 'left' });
      s.anim(n, 'fade', at + 500, 400);
    };
    runway(1.0, 2.9, 6.5, '北跑道', '3400 m', true, 500);
    runway(1.0, 5.75, 6.9, '南跑道', '3600 m', false, 800);
    const apron = s.shape('roundRect', { x: 2.2, y: 3.75, w: 3.9, h: 1.5, rectRadius: 0.2, fill: { color: C.white }, line: { color: C.blueSoft, width: 1.5 } });
    const apT = s.text('航站区', { x: 2.2, y: 3.9, w: 3.9, h: 0.5, fontSize: 18, bold: true, color: C.blueDeep, align: 'center', valign: 'middle' });
    const apD = s.text('3 号、4 号航站楼 · 综合交通中心', { x: 2.2, y: 4.42, w: 3.9, h: 0.34, fontSize: 10.5, color: C.sub, align: 'center', valign: 'middle' });
    const apE = s.text('航站区夹在两条东西向的平行跑道之间', { x: 2.2, y: 4.78, w: 3.9, h: 0.3, fontSize: 9.5, color: C.mute, align: 'center', valign: 'middle' });
    const tower = s.art('tower', 6.55, 3.5, 1.05);
    s.anim(apron, 'zoom', 1200, 500);
    [apT, apD, apE].forEach((t, i) => s.anim(t, 'fade', 1400 + i * 120, 350));
    s.anim(tower, 'rise', 1500, 600);
    const jet = s.art('plane', 1.05, 5.3, 1.1);
    s.anim(jet, 'fade', 1900, 250);
    s.anim(jet, 'path', 2100, 3000, { path: 'M 0 0 L 0.25 0 C 0.32 0 0.38 -0.05 0.45 -0.2 E' });
    s.anim(jet, 'fadeOut', 4700, 400);

    const stats = [
      ['165', '个', '客机位', C.blueDeep],
      ['72', '座', '登机桥', C.pinkDeep],
      ['27', '个', '货机位', C.greenDeep],
    ];
    stats.forEach(([n, u, l, col], i) => tile(s, 8.55 + i * 1.43, 2.05, 1.34, 1.3, n, u, l, col, 900 + i * 200, { numSize: 26, unitSize: 11, labelSize: 10.5 }));
    const { pic, label } = await photo(s, 'aerial2', 8.55, 3.55, 4.2, 2.3, { caption: '跑道与航站区', fy: 0.5 });
    s.anim(pic, 'zoom', 1500, 600);
    s.anim(label, 'fade', 2000, 400);
    const say = bubble(s, '4F 是我国民用机场飞行区的最高等级，能起降空客 A380 这样的大型客机。', 8.95, 6.02, 3.8, 0.66, { size: 10.5, tail: [-47000, 62000] });
    const cl = s.art('cloudSmile', 8.3, 6.5, 0.85);
    s.anim(say, 'zoom', 2400, 450);
    s.anim(cl, 'zoom', 2300, 400);
    footer(s, '杭州萧山国际机场官网（保障能力）');
    s.notes(
      '【飞行区】第四站看枢纽实力，先看飞行区。\n' +
        '萧山机场有两条东西向的平行跑道，北跑道长 3400 米，南跑道长 3600 米，航站区夹在两条跑道中间。\n' +
        '飞行区等级是 4F，这是我国民用机场的最高等级，可以起降空客 A380 这样的大型客机。\n' +
        '全场现在有 165 个客机位、72 座登机桥，还有 27 个货机位。'
    );
  }

  // ===== 9 航线网络 ======================================================================
  {
    const s = newSlide('flythrough');
    header(s, 4, '从杭州出发，飞向五大洲', 9);
    const tiles = [
      ['223', '个', '通航点，遍布 5 大洲', C.blueDeep],
      ['35', '个', '通达的国家和地区', C.greenDeep],
      ['75', '家', '在这里运营的航空公司', C.yellowDeep],
      ['89.1', '%', '平均客座率：全国前十机场第一，已连续 24 个月', C.pinkDeep],
    ];
    tiles.forEach(([n, u, l, col], i) => tile(s, 0.6 + (i % 2) * 2.95, 2.05 + Math.floor(i / 2) * 1.62, 2.8, 1.47, n, u, l, col, 350 + i * 220, { labelSize: 10.5 }));
    const tag = s.text('截至 2025 年底', {
      x: 4.75, y: 1.82, w: 1.6, h: 0.34, shape: 'roundRect', rectRadius: 0.17, rotate: 4,
      fill: { color: C.yellow }, color: C.navy, bold: true, fontSize: 11, align: 'center', valign: 'middle', shadow: shadow(),
    });
    s.anim(tag, 'zoom', 1300, 400);
    const more = card(s, 0.6, 5.3, 5.75, 1.45, { fill: C.bluePale, line: 'C9E4F8' });
    const icn = s.art('i-globe', 0.8, 5.52, 0.62);
    const m1 = s.text('2025 年国际及地区旅客约 498 万人次', { x: 1.6, y: 5.42, w: 4.6, h: 0.4, fontSize: 13.5, bold: true, color: C.navy, valign: 'middle' });
    const m2 = s.text('国产大飞机也来了：2024 年 9 月 11 日，国航 C919 首次执飞杭州—北京首都航线。', { x: 1.6, y: 5.84, w: 4.6, h: 0.78, fontSize: 11, color: C.text, valign: 'top', lineSpacingMultiple: 1.25 });
    s.anim(more, 'rise', 1500, 500);
    s.anim(icn, 'zoom', 1700, 350);
    s.anim(m1, 'fade', 1750, 400);
    s.anim(m2, 'fade', 1900, 400);

    const p1 = await photo(s, 'loong', 6.6, 2.05, 6.15, 2.6, { caption: '总部设在杭州的长龙航空', fy: 0.55 });
    const p2 = await photo(s, 'c919', 6.6, 4.82, 3.0, 1.93, { caption: '国航 C919', fx: 0.45, fy: 0.5 });
    const p3 = await photo(s, 'b787', 9.75, 4.82, 3.0, 1.93, { caption: '南航波音 787', fy: 0.45 });
    s.anim(p1.pic, 'slide', 600, 600, { offset: 0.06 });
    s.anim(p1.label, 'fade', 1200, 350);
    s.anim(p2.pic, 'rise', 1100, 550);
    s.anim(p2.label, 'fade', 1650, 350);
    s.anim(p3.pic, 'rise', 1350, 550);
    s.anim(p3.label, 'fade', 1900, 350);
    footer(s, '每日商报 2026-01-07（杭州机场 2025 年运行数据）；中国民航局 2025 年统计公报；中国民航网 2024-09');
    s.notes(
      '【航线网络】截至 2025 年底，萧山机场一共通达 223 个航点，覆盖 5 大洲、35 个国家和地区，有 75 家航空公司在这里运营。\n' +
        '它的平均客座率是 89.1%，在全国旅客量前十的机场里排第一，而且已经连续 24 个月保持第一。\n' +
        '2025 年国际及地区旅客大约 498 万人次。\n' +
        '另外一个小亮点：2024 年 9 月 11 日，国航的国产大飞机 C919 第一次执飞杭州到北京首都的航线。右边大图是总部设在杭州的长龙航空。'
    );
  }

  // ===== 10 全国排名 =====================================================================
  {
    const s = newSlide('cube');
    header(s, 4, '放到全国看：2025 年排第 7', 10);
    const rank = [
      ['上海浦东', 8499.5], ['广州白云', 8358.3], ['北京首都', 7074.3], ['深圳宝安', 6648.5], ['成都天府', 5668.7],
      ['北京大兴', 5361.9], ['杭州萧山', 5045.9], ['上海虹桥', 5015.1], ['重庆江北', 5009.5], ['昆明长水', 4970.6],
    ];
    const rev = [...rank].reverse();
    const box = card(s, 0.6, 2.0, 8.2, 4.75);
    const cap = s.text('2025 年旅客吞吐量前十位机场（万人次）', { x: 0.9, y: 2.14, w: 6, h: 0.32, fontSize: 12, bold: true, color: C.navy, valign: 'middle' });
    const chart = s.chart(pptx.ChartType.bar, [{ name: '旅客吞吐量（万人次）', labels: rev.map((r, i) => `${10 - i}  ${r[0]}`), values: rev.map((r) => r[1]) }], {
      x: 0.75, y: 2.48, w: 7.9, h: 4.2, barDir: 'bar', barGapWidthPct: 38, chartColors: rev.map((r) => (r[0] === '杭州萧山' ? C.pink : C.blueSoft)),
      showLegend: false, showTitle: false, showValue: true, dataLabelFontSize: 10, dataLabelColor: C.text, dataLabelFormatCode: '0.0', dataLabelPosition: 'outEnd',
      dataLabelFontFace: FONT.num, catAxisLabelFontSize: 11, catAxisLabelColor: C.text, catAxisLabelFontFace: FONT.cn,
      valAxisHidden: true, valAxisMaxVal: 9600, valAxisMinVal: 0, valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
    });
    s.anim(box, 'rise', 300, 500);
    s.anim(cap, 'fade', 600, 350);
    s.anim(chart, 'wipe', 800, 1500, { dir: 'left' });

    const medal = s.shape('ellipse', { x: 9.75, y: 2.0, w: 2.3, h: 2.3, fill: { color: gradient(90, [{ pos: 0, color: 'FFB3C6' }, { pos: 100, color: C.pinkDeep }]) }, line: { color: C.white, width: 5 }, shadow: shadow() });
    const mt = s.text(
      [
        { text: '全国第', options: { fontSize: 14, bold: true, color: C.white, breakLine: true } },
        { text: '7', options: { fontFace: FONT.num, fontSize: 66, bold: true, color: C.white } },
      ],
      { x: 9.75, y: 2.1, w: 2.3, h: 2.1, align: 'center', valign: 'middle' }
    );
    const st = s.art('star', 11.85, 1.95, 0.45);
    const st2 = s.art('star', 9.55, 3.75, 0.3);
    s.anim(medal, 'zoom', 1500, 600);
    s.anim(mt, 'zoom', 1700, 500);
    s.anim(medal, 'pulse', 2600, 800, { by: 1.05, repeat: 2 });
    [st, st2].forEach((x, i) => {
      s.anim(x, 'zoom', 2000 + i * 200, 350);
      s.anim(x, 'spin', 2400, 4000, { turns: 1, repeat: 'forever' });
    });
    const facts = [
      ['旅客', '5045.9 万人次，同比 +5.0%，比 2024 年上升 1 位', C.pinkDeep],
      ['货邮', '79.27 万吨，同样排全国第 7', C.yellowDeep],
      ['起降', '32.94 万架次，全国第 10', C.blueDeep],
    ];
    facts.forEach(([k, v, col], i) => {
      const y = 4.55 + i * 0.75;
      const c = card(s, 9.05, y, 3.7, 0.66, { radius: 0.18 });
      const kk = s.text(k, { x: 9.17, y: y + 0.13, w: 0.72, h: 0.4, shape: 'roundRect', rectRadius: 0.2, fill: { color: col }, color: C.white, bold: true, fontSize: 11, align: 'center', valign: 'middle' });
      const vv = s.text(v, { x: 10.0, y: y + 0.05, w: 2.68, h: 0.56, fontSize: 10.5, color: C.text, valign: 'middle', lineSpacingMultiple: 1.1 });
      const at = 2100 + i * 250;
      s.anim(c, 'slide', at, 400, { offset: 0.05 });
      s.anim(kk, 'zoom', at + 150, 300);
      s.anim(vv, 'fade', at + 200, 350);
    });
    footer(s, '中国民航局《2025 年全国民用运输机场生产统计公报》');
    s.notes(
      '【全国排名】把它放到全国来看。这是 2025 年全国旅客吞吐量前十的机场。\n' +
        '杭州萧山排在第 7 位，5045.9 万人次，同比增长 5.0%，比 2024 年上升了 1 位，排在上海虹桥、重庆江北前面。\n' +
        '除了客运，货邮吞吐量 79.27 万吨也是全国第 7，起降 32.94 万架次，排全国第 10。'
    );
  }

  // ===== 11 货运 =========================================================================
  {
    const s = newSlide('origami');
    header(s, 4, '不只载客：全国第 7 的货运机场', 11);
    const p1 = await photo(s, 'sfhub', 0.6, 2.05, 6.0, 2.85, { caption: '机场内的顺丰物流枢纽与货机坪', fy: 0.55 });
    const p2 = await photo(s, 'sf', 0.6, 5.05, 6.0, 1.7, { caption: '停场的顺丰航空货机', fy: 0.62 });
    s.anim(p1.pic, 'wipe', 300, 800, { dir: 'left' });
    s.anim(p1.label, 'fade', 1100, 350);
    s.anim(p2.pic, 'wipe', 700, 800, { dir: 'left' });
    s.anim(p2.label, 'fade', 1500, 350);
    const rows = [
      ['i-box', '79.27', '万吨', '2025 年货邮吞吐量，同比增长 7.9%，全国第 7', C.yellowDeep],
      ['i-globe', '25.7', '万吨', '其中国际及地区货邮吞吐量', C.greenDeep],
      ['i-plane', '28', '条', '国际货运航线（截至 2025 年底）', C.blueDeep],
      ['i-flag', '150', '万吨', '规划目标：2035 年的年货邮吞吐量', C.pinkDeep],
    ];
    rows.forEach(([ic, n, u, l, col], i) => {
      const y = 2.05 + i * 1.2;
      const c = card(s, 6.9, y, 5.85, 1.05, { radius: 0.2 });
      const icn = s.art(ic, 7.1, y + 0.18, 0.69);
      const num = s.text(
        [
          { text: n, options: { fontFace: FONT.num, fontSize: 28, bold: true, color: col } },
          { text: ' ' + u, options: { fontSize: 12, bold: true, color: col } },
        ],
        { x: 8.0, y: y + 0.1, w: 1.95, h: 0.85, valign: 'middle' }
      );
      const lab = s.text(l, { x: 9.95, y: y + 0.1, w: 2.65, h: 0.85, fontSize: 11, color: C.text, valign: 'middle', lineSpacingMultiple: 1.2 });
      const at = 600 + i * 320;
      s.anim(c, 'slide', at, 450, { offset: 0.06 });
      s.anim(icn, 'wheel', at + 150, 500);
      s.anim(num, 'zoom', at + 200, 400);
      s.anim(lab, 'fade', at + 300, 400);
    });
    footer(s, '中国民航局 2025 年统计公报；每日商报 2026-01-07；中国民航报 2024-12-16');
    s.notes(
      '【货运】萧山机场不只是载客。\n' +
        '2025 年它的货邮吞吐量是 79.27 万吨，同比增长 7.9%，同样排全国第 7，其中国际及地区货邮 25.7 万吨。\n' +
        '截至 2025 年底有 28 条国际货运航线。照片里是机场内的顺丰物流枢纽和货机。\n' +
        '按照规划，到 2035 年它的年货邮吞吐量目标是 150 万吨。'
    );
  }

  // ===== 12 交通中心 =====================================================================
  {
    const s = newSlide('doors');
    header(s, 5, '下了飞机，5 分钟内换乘', 12);
    const { pic, label } = await photo(s, 'curb', 0.6, 2.05, 5.0, 4.7, { caption: 'T4 出发层车道边', fx: 0.75 });
    s.anim(pic, 'zoom', 300, 650);
    s.anim(label, 'fade', 900, 400);

    const cx = 9.35;
    const cy = 3.82;
    const R = 1.5;
    const ring = s.shape('ellipse', { x: cx - R, y: cy - R, w: R * 2, h: R * 2, fill: { color: C.white, transparency: 100 }, line: { color: C.yellow, width: 1.5, dashType: 'dash' } });
    s.anim(ring, 'wheel', 500, 1200);
    const hub = s.text(
      [
        { text: '综合交通中心', options: { fontSize: 14.5, bold: true, color: C.white, breakLine: true } },
        { text: '位于新老航站楼之间', options: { fontSize: 9, color: C.white } },
      ],
      { x: cx - 0.95, y: cy - 0.95, w: 1.9, h: 1.9, shape: 'ellipse', fill: { color: gradient(90, [{ pos: 0, color: 'FFD46B' }, { pos: 100, color: C.yellowDeep }]) }, line: { color: C.white, width: 4 }, align: 'center', valign: 'middle', shadow: shadow() }
    );
    s.anim(hub, 'zoom', 700, 550);
    s.anim(hub, 'pulse', 3600, 900, { by: 1.05, repeat: 2 });
    const modes = [
      ['i-plane', '航站楼'],
      ['i-metro', '地铁'],
      ['i-train', '高铁（在建）'],
      ['i-bus', '机场大巴'],
      ['i-car', '出租车 · 网约车'],
      ['i-parking', '停车库'],
    ];
    modes.forEach(([ic, name], i) => {
      const a = ((-90 + i * 60) * Math.PI) / 180;
      const x = cx + Math.cos(a) * R * 1.0;
      const y = cy + Math.sin(a) * R * 1.0;
      const icn = s.art(ic, x - 0.36, y - 0.36, 0.72);
      const side = Math.cos(a) > 0.3 ? 'right' : Math.cos(a) < -0.3 ? 'left' : 'mid';
      const lx = side === 'right' ? x + 0.44 : side === 'left' ? x - 0.44 - 1.6 : x - 0.8;
      const ly = side === 'mid' ? (Math.sin(a) < 0 ? y - 0.36 - 0.3 : y + 0.38) : y - 0.16;
      const t = s.text(name, { x: lx, y: ly, w: 1.6, h: 0.32, fontSize: 11.5, bold: true, color: C.navy, align: side === 'right' ? 'left' : side === 'left' ? 'right' : 'center', valign: 'middle' });
      const at = 1200 + i * 260;
      s.anim(icn, 'zoom', at, 400);
      s.anim(t, 'fade', at + 150, 350);
    });
    const chips = [
      ['48 万㎡', '总建筑面积'],
      ['地上 2 层 · 地下 4 层', '立体换乘'],
      ['4400 个', '停车位，含 524 个充电车位'],
    ];
    chips.forEach(([big, small], i) => {
      const x = 5.9 + i * 2.31;
      const c = card(s, x, 6.0, 2.2, 0.76, { radius: 0.2 });
      const b = s.text(big, { x: x + 0.1, y: 6.03, w: 2.0, h: 0.4, fontSize: i === 1 ? 12 : 15, bold: true, color: [C.yellowDeep, C.blueDeep, C.greenDeep][i], align: 'center', valign: 'middle' });
      const t = s.text(small, { x: x + 0.05, y: 6.42, w: 2.1, h: 0.28, fontSize: 9, color: C.sub, align: 'center', valign: 'middle' });
      const at = 2900 + i * 220;
      s.anim(c, 'rise', at, 400);
      s.anim(b, 'fade', at + 150, 300);
      s.anim(t, 'fade', at + 200, 300);
    });
    footer(s, '杭州萧山国际机场官网 2022-09-15、2022-09-22');
    s.notes(
      '【交通中心】第五站讲怎么到机场、怎么离开机场。\n' +
        '三期工程在新老航站楼之间建了一座综合交通中心，总建筑面积约 48 万平方米，地上 2 层、地下 4 层。\n' +
        '地铁、机场大巴、出租车、网约车和停车库都集中在这里，按机场的介绍，换乘时间都在 5 分钟以内；高铁站还在建设中。\n' +
        '停车库设计了 4400 个车位，其中 524 个是充电车位。'
    );
  }

  // ===== 13 轨道交通 =====================================================================
  {
    const s = newSlide('pan');
    header(s, 5, '三条地铁进机场，高铁正在路上', 13);
    const box = card(s, 0.6, 2.05, 7.95, 2.75);
    const h1 = s.text('19 号线 · 机场轨道快线', { x: 0.9, y: 2.2, w: 4, h: 0.4, fontSize: 16, bold: true, color: C.navy, valign: 'middle' });
    const chip = s.text('最高时速 120 km/h', { x: 6.3, y: 2.22, w: 1.95, h: 0.36, shape: 'roundRect', rectRadius: 0.18, fill: { color: C.yellowPale }, color: C.yellowDeep, bold: true, fontSize: 11, align: 'center', valign: 'middle' });
    s.anim(box, 'rise', 300, 500);
    s.anim(h1, 'fade', 550, 350);
    s.anim(chip, 'zoom', 700, 350);
    const lineY = 3.5;
    const x0 = 1.5;
    const x1 = 7.65;
    const track = s.shape('roundRect', { x: x0, y: lineY - 0.06, w: x1 - x0, h: 0.12, rectRadius: 0.06, fill: { color: C.blue }, line: { color: C.blue, width: 0 } });
    s.anim(track, 'wipe', 800, 1000, { dir: 'right' });
    const stations = [
      ['杭州西站', '约 45 分钟', x0],
      ['西湖文化广场', '约 31 分钟', x0 + (x1 - x0) * 0.36],
      ['杭州东站', '约 20 分钟', x0 + (x1 - x0) * 0.68],
      ['萧山国际机场', '从这里出发', x1],
    ];
    stations.forEach(([name, time, x], i) => {
      const last = i === stations.length - 1;
      const dot = s.shape('ellipse', { x: x - 0.16, y: lineY - 0.16, w: 0.32, h: 0.32, fill: { color: C.white }, line: { color: last ? C.pinkDeep : C.blue, width: 3 } });
      const n = s.text(name, { x: x - 0.95, y: lineY + 0.26, w: 1.9, h: 0.34, fontSize: 12.5, bold: true, color: C.navy, align: 'center', valign: 'middle' });
      const t = s.text(time, { x: x - 0.8, y: lineY + 0.62, w: 1.6, h: 0.34, shape: 'roundRect', rectRadius: 0.17, fill: { color: last ? C.pinkPale : C.bluePale }, color: last ? C.pinkDeep : C.blueDeep, bold: true, fontSize: 11, align: 'center', valign: 'middle' });
      const at = 1700 - i * 280 + 300;
      s.anim(dot, 'zoom', at, 300);
      s.anim(n, 'fade', at + 100, 300);
      s.anim(t, 'zoom', at + 200, 300);
    });
    const metro = s.art('i-metro', x1 - 0.3, lineY - 0.84, 0.6);
    s.anim(metro, 'zoom', 900, 300);
    s.anim(metro, 'path', 1300, 5000, { path: `M 0 0 L ${-Math.round(((x1 - x0) / W) * 10000) / 10000} 0 E`, autoRev: true, repeat: 'forever' });

    const lines = [
      ['1 号线、7 号线', '2020 年 12 月 30 日通到机场', C.pinkDeep, C.pinkPale],
      ['19 号线', '2022 年 9 月 22 日开通，全长约 59 公里', C.blueDeep, C.bluePale],
      ['机场高铁（在建）', '萧山机场站设在地下，规模为 2 台 4 线', C.greenDeep, C.greenPale],
    ];
    lines.forEach(([k, v, col, pale], i) => {
      const y = 5.0 + i * 0.6;
      const c = card(s, 0.6, y, 7.95, 0.52, { radius: 0.18 });
      const kk = s.text(k, { x: 0.72, y: y + 0.08, w: 2.1, h: 0.36, shape: 'roundRect', rectRadius: 0.18, fill: { color: pale }, color: col, bold: true, fontSize: 11.5, align: 'center', valign: 'middle' });
      const vv = s.text(v, { x: 3.0, y: y + 0.06, w: 5.4, h: 0.4, fontSize: 12, color: C.text, valign: 'middle' });
      const at = 2400 + i * 260;
      s.anim(c, 'slide', at, 400, { offset: -0.05 });
      s.anim(kk, 'zoom', at + 120, 300);
      s.anim(vv, 'fade', at + 200, 350);
    });
    const p1 = await photo(s, 'metro19', 8.8, 2.05, 3.95, 2.75, { caption: '19 号线机场站（开通当天）', fx: 0.3 });
    const p2 = await photo(s, 'sign', 8.8, 4.98, 3.95, 1.77, { caption: '地铁站名墙', fy: 0.5 });
    s.anim(p1.pic, 'zoom', 600, 600);
    s.anim(p1.label, 'fade', 1150, 350);
    s.anim(p2.pic, 'zoom', 1000, 600);
    s.anim(p2.label, 'fade', 1550, 350);
    footer(s, '杭州萧山国际机场官网 2022-09-22；杭州地铁、浙江省交通运输部门公开信息');
    s.notes(
      '【轨道交通】现在有三条地铁线通到萧山机场。\n' +
        '1 号线和 7 号线是 2020 年 12 月 30 日通到机场的；19 号线也叫机场轨道快线，2022 年 9 月 22 日开通，全长约 59 公里，最高时速 120 公里。\n' +
        '坐 19 号线，从机场到杭州东站大约 20 分钟，到西湖文化广场约 31 分钟，到杭州西站约 45 分钟。\n' +
        '另外，机场高铁正在建设，萧山机场站设在地下，规模是 2 台 4 线，建成后可以空铁联运。'
    );
  }

  // ===== 14 亚运 =========================================================================
  {
    const s = newSlide('glitter', 'bgPink');
    header(s, 6, '2023 年：把亚洲迎进家门', 14);
    const { pic, label } = await photo(s, 'columns', 0.6, 2.05, 5.2, 4.7, { caption: 'T4 出发层大厅', fx: 0.4 });
    s.anim(pic, 'circle', 300, 900);
    s.anim(label, 'fade', 1200, 400);
    const head = card(s, 6.05, 2.05, 6.7, 1.2, { fill: C.pinkPale, line: 'FFD3DE' });
    const icn = s.art('i-medal', 6.27, 2.3, 0.7);
    const h1 = s.text('杭州第 19 届亚运会官方出入境口岸', { x: 7.15, y: 2.15, w: 5.5, h: 0.48, fontSize: 18, bold: true, color: C.pinkDeep, valign: 'middle' });
    const h2 = s.text('三期工程（T4 航站楼、交通中心）本身就是亚运会的重点配套项目。', { x: 7.15, y: 2.63, w: 5.5, h: 0.5, fontSize: 11, color: C.text, valign: 'middle' });
    s.anim(head, 'rise', 600, 500);
    s.anim(icn, 'wheel', 800, 600);
    s.anim(h1, 'wipe', 850, 600, { dir: 'left' });
    s.anim(h2, 'fade', 1200, 400);
    const tiles = [
      ['4538', '架次', '保障亚运会、亚残运会客运航班', C.blueDeep],
      ['5.09', '万人次', '运送的赛事相关旅客', C.greenDeep],
      ['3200+', '人', '9 月 21 日单日入境高峰，来自 35 个代表团', C.yellowDeep],
      ['224', '匹', '运输的亚运赛马（货运航班 44 架次）', C.pinkDeep],
    ];
    tiles.forEach(([n, u, l, col], i) => tile(s, 6.05 + (i % 2) * 3.4, 3.42 + Math.floor(i / 2) * 1.5, 3.3, 1.38, n, u, l, col, 1300 + i * 240, { labelSize: 10.5 }));
    const say = bubble(s, '连马儿都是“飞”来杭州的！', 8.75, 6.42, 2.6, 0.42, { size: 11, bold: true, line: 'FFC2D1', tail: [58000, 0] });
    const guide = s.art('planeOpen', 11.66, 6.2, 1.05);
    s.anim(guide, 'fly', 2500, 500, { dir: 'right' });
    s.anim(say, 'zoom', 2900, 400);
    bob(s, guide, -0.012, 1500, 3100);
    footer(s, '杭州萧山国际机场官网 2023-11-30；新华社 2023-09-21；浙江省交通运输厅 2022-03-14');
    s.notes(
      '【亚运会】最后一站，先回顾 2023 年的杭州亚运会。\n' +
        '萧山机场是这届亚运会的官方出入境口岸，T4 航站楼和交通中心本身就是亚运会的重点配套项目。\n' +
        '根据机场公布的数据，亚运会和亚残运会期间一共保障客运航班 4538 架次，运送赛事相关旅客 5.09 万人次。\n' +
        '9 月 21 日是单日入境高峰，当天有 35 个代表团、3200 多名赛事人员抵达。\n' +
        '一个有趣的数字：参加马术比赛的 224 匹赛马，也是坐飞机从这里入境的。'
    );
  }

  // ===== 15 未来 =========================================================================
  {
    const s = newSlide('window');
    header(s, 6, '下一站：更大的萧山机场', 15);
    const stops = [
      ['计划 2026.12', 'T4 南区投运', 'T4 航站楼南区计划于 2026 年 12 月 1 日投运，东航将从 3 号航站楼转场入驻。', 'i-plane', C.blue, C.blueDeep, C.bluePale],
      ['建设中', '机场高铁', '萧山机场站设在地下，建成后高铁直接开进机场，实现空铁联运。', 'i-train', C.pink, C.pinkDeep, C.pinkPale],
      ['前期研究', '四期工程', '包含第三跑道在内的四期扩建工程，目前处于前期研究阶段。', 'i-ruler', C.yellow, C.yellowDeep, C.yellowPale],
      ['目标 2035', '更大的枢纽', '规划到 2035 年，年旅客吞吐量 7500 万人次、货邮吞吐量 150 万吨。', 'i-flag', C.green, C.greenDeep, C.greenPale],
    ];
    const end = [12.6, 2.2];
    const r = route([[-0.6, 2.75], [2.05, 2.55], [5.15, 2.85], [8.25, 2.5], [11.3, 2.75], end], { x: 0, y: 1.9 }, end);
    const trail = s.image(await trailArt('future', Math.round(W * PX), Math.round(1.3 * PX), r.svg, A.P.blueSoft, 6, '3 20'), 0, 1.9, W, 1.3, { alt: '航线' });
    s.anim(trail, 'wipe', 300, 2600, { dir: 'left' });
    stops.forEach(([when, name, desc, ic, color, deep, pale], i) => {
      const x = 0.6 + i * 3.08;
      const y = 3.3;
      const cw = 2.92;
      const c = card(s, x, y, cw, 2.75, { radius: 0.26 });
      const tag = s.text(when, { x: x + 0.24, y: y - 0.2, w: 1.5, h: 0.4, shape: 'roundRect', rectRadius: 0.2, fill: { color: deep }, color: C.white, bold: true, fontSize: 11.5, align: 'center', valign: 'middle', shadow: shadow() });
      const icn = s.art(ic, x + cw - 1.05, y + 0.32, 0.78);
      const t = s.text(name, { x: x + 0.26, y: y + 0.45, w: 1.7, h: 0.5, fontSize: 18, bold: true, color: C.navy, valign: 'middle' });
      const band = s.shape('roundRect', { x: x + 0.26, y: y + 1.05, w: 0.6, h: 0.07, rectRadius: 0.035, fill: { color }, line: { color, width: 0 } });
      const d = s.text(desc, { x: x + 0.26, y: y + 1.3, w: cw - 0.5, h: 1.3, fontSize: 12, color: C.text, valign: 'top', lineSpacingMultiple: 1.35 });
      const pin = s.shape('ellipse', { x: x + cw / 2 - 0.12, y: [2.55, 2.85, 2.5, 2.75][i] - 0.12, w: 0.24, h: 0.24, fill: { color }, line: { color: C.white, width: 2.5 } });
      const at = 700 + i * 500;
      s.anim(pin, 'zoom', at, 300);
      s.anim(c, 'rise', at + 100, 500);
      s.anim(tag, 'zoom', at + 300, 350);
      s.anim(icn, 'wheel', at + 350, 500);
      s.anim(t, 'fade', at + 350, 350);
      s.anim(band, 'wipe', at + 450, 300, { dir: 'left' });
      s.anim(d, 'fade', at + 500, 400);
    });
    const plane = s.art('plane', end[0] - 0.55, end[1] - 0.35, 1.1);
    s.anim(plane, 'fade', 300, 250);
    s.anim(plane, 'path', 300, 3000, { path: r.motion, ease: false });
    bob(s, plane, -0.012, 1500, 3400);

    // progress towards the 2035 passenger target: 5045.9 / 7500 = 67 %
    const bx = 2.55;
    const bw = 7.6;
    const share = 5045.9 / 7500;
    const pl = s.text('年旅客吞吐量', { x: 0.6, y: 6.33, w: 1.85, h: 0.44, fontSize: 12.5, bold: true, color: C.navy, valign: 'middle' });
    const back = s.shape('roundRect', { x: bx, y: 6.33, w: bw, h: 0.44, rectRadius: 0.22, fill: { color: C.white }, line: { color: C.line, width: 1 } });
    const fill = s.text('2025 年：5045.9 万人次（目标的 67%）', {
      x: bx, y: 6.33, w: bw * share, h: 0.44, shape: 'roundRect', rectRadius: 0.22,
      fill: { color: gradient(0, [{ pos: 0, color: C.blueSoft }, { pos: 100, color: C.pink }]) }, line: { color: C.white, width: 0 },
      color: C.white, bold: true, fontSize: 12, align: 'center', valign: 'middle',
    });
    const goal = s.text('2035 年目标：7500 万', { x: bx + bw + 0.12, y: 6.33, w: 2.5, h: 0.44, fontSize: 12.5, bold: true, color: C.greenDeep, valign: 'middle' });
    s.anim(pl, 'fade', 2800, 350);
    s.anim(back, 'fade', 2800, 350);
    s.anim(fill, 'wipe', 3000, 1300, { dir: 'left' });
    s.anim(goal, 'zoom', 4100, 400);
    footer(s, '杭州萧山国际机场、浙江省机场集团 2026 年招商公告；中国民航报 2024-12-16');
    s.notes(
      '【未来】萧山机场还在继续长大。\n' +
        '最近的一件事：T4 航站楼南区计划在 2026 年 12 月 1 日投运，东航会从 3 号航站楼搬过去。\n' +
        '机场高铁正在建设，车站设在机场地下，建成后可以空铁联运。\n' +
        '再往后是四期工程，包括第三条跑道，目前还在前期研究阶段。\n' +
        '按规划，到 2035 年它的目标是年旅客吞吐量 7500 万人次、货邮吞吐量 150 万吨。'
    );
  }

  // ===== 16 小结 =========================================================================
  {
    const s = newSlide('ripple');
    const pill = s.text('SUMMARY  ·  小结', {
      x: 0.6, y: 0.42, w: 2.2, h: 0.38, shape: 'roundRect', rectRadius: 0.19,
      fill: { color: C.green }, color: C.white, bold: true, fontSize: 12, align: 'center', valign: 'middle',
    });
    const title = s.text('三句话，带走萧山机场', { x: 0.6, y: 0.9, w: 9, h: 0.72, fontSize: 30, bold: true, color: C.navy, valign: 'middle' });
    s.anim(pill, 'wipe', 0, 350, { dir: 'left' });
    s.anim(title, 'rise', 120, 500);
    const cards = [
      ['lotus', 1.5, '一朵“莲”', 'T4 航站楼把西湖荷花做成了建筑：40 根荷叶柱，53.5 米高的荷花谷。', C.pinkDeep, C.pinkPale],
      ['i-globe', 1.0, '一张“网”', '223 个通航点、3 条地铁线，再加上在建的机场高铁，空中地面连成一张网。', C.blueDeep, C.bluePale],
      ['planeWink', 1.55, '一条“上升线”', '25 年间旅客量从 298.1 万增长到 5045.9 万人次，2025 年排名全国第 7。', C.greenDeep, C.greenPale],
    ];
    cards.forEach(([artName, aw, name, desc, color, pale], i) => {
      const x = 0.6 + i * 4.11;
      const y = 2.15;
      const c = card(s, x, y, 3.93, 3.75, { radius: 0.3 });
      const top = s.shape('roundRect', { x: x + 0.2, y: y + 0.2, w: 3.53, h: 1.45, rectRadius: 0.22, fill: { color: pale }, line: { color: pale, width: 0 } });
      const a = s.art(artName, x + 1.965 - aw / 2, y + 0.925 - aw / RATIO[artName] / 2, aw);
      const t = s.text(name, { x: x + 0.3, y: y + 1.82, w: 3.33, h: 0.52, fontSize: 22, bold: true, color, align: 'center', valign: 'middle' });
      const d = s.text(desc, { x: x + 0.35, y: y + 2.42, w: 3.23, h: 1.2, fontSize: 13, color: C.text, align: 'center', valign: 'top', lineSpacingMultiple: 1.35 });
      const at = 400 + i * 550;
      s.anim(c, 'rise', at, 500);
      s.anim(top, 'fade', at + 200, 350);
      s.anim(a, 'zoom', at + 300, 500);
      s.anim(a, 'pulse', at + 1500, 1300, { by: 1.07, repeat: 'forever' });
      s.anim(t, 'zoom', at + 400, 400);
      s.anim(d, 'fade', at + 550, 400);
    });
    const say = bubble(s, '下次去杭州坐飞机，记得抬头看看荷叶柱哦～', 3.1, 6.18, 5.6, 0.55, { size: 13, bold: true, tail: [53500, 0] });
    const cl = s.art('cloudSmile', 8.95, 6.2, 1.15);
    s.anim(cl, 'zoom', 2300, 450);
    s.anim(say, 'zoom', 2550, 450);
    bob(s, cl, -0.012, 1600, 3000);
    footer(s, '');
    s.notes(
      '【小结】最后用三句话总结。\n' +
        '第一，一朵“莲”：T4 航站楼把西湖的荷花做成了建筑，有 40 根荷叶柱和 53.5 米高的荷花谷。\n' +
        '第二，一张“网”：223 个通航点、3 条地铁线，加上在建的机场高铁，空中和地面连成了一张网。\n' +
        '第三，一条“上升线”：25 年里旅客量从 298.1 万增长到 5045.9 万人次，2025 年排名全国第 7。'
    );
  }

  // ===== 17 来源 =========================================================================
  {
    const s = newSlide('fade');
    const pill = s.text('SOURCES  ·  资料与图片来源', {
      x: 0.6, y: 0.42, w: 3.0, h: 0.38, shape: 'roundRect', rectRadius: 0.19,
      fill: { color: C.blue }, color: C.white, bold: true, fontSize: 12, align: 'center', valign: 'middle',
    });
    const title = s.text('数据都查得到：资料与图片来源', { x: 0.6, y: 0.9, w: 9, h: 0.72, fontSize: 30, bold: true, color: C.navy, valign: 'middle' });
    s.anim(pill, 'wipe', 0, 350, { dir: 'left' });
    s.anim(title, 'rise', 120, 500);
    const left = card(s, 0.6, 1.95, 6.1, 4.3);
    const lh = s.text('数据与文字资料', { x: 0.85, y: 2.05, w: 4, h: 0.38, fontSize: 14, bold: true, color: C.blueDeep, valign: 'middle' });
    const src = [
      '中国民用航空局：历年《民航机场生产统计公报》（2001–2025 年）',
      '杭州萧山国际机场官网 hzairport.com：机场概况、保障能力、发展历程',
      '杭州萧山国际机场官网新闻：2022-09-15、2022-09-22（三期投运）、2023-11-30（亚运保障数据）',
      '杭州萧山国际机场、浙江省机场集团招商公告（2026-09-18、2026-11-04）',
      '新华社 2023-09-21《杭州亚运会迎来代表团单日入境最高峰》',
      '浙江在线 2022-04-01；浙江省交通运输厅 2022-03-14（T4 设计）',
      '每日商报 2026-01-07（2025 年航线网络与客座率）',
      '中国民航报 2024-12-16（2035 年目标）；中国民航网 2024-09（C919）',
      '人民网 2004-03-18；浙江在线、新华社 2025-12-28（5000 万人次）',
    ];
    const lt = s.text(
      src.map((t) => ({ text: t, options: { bullet: { characterCode: '25CF' }, breakLine: true } })),
      { x: 0.85, y: 2.5, w: 5.65, h: 3.65, fontSize: 11, color: C.text, valign: 'top', paraSpaceAfter: 5, lineSpacingMultiple: 1.15 }
    );
    const right = card(s, 6.95, 1.95, 5.8, 4.3);
    const rh = s.text('图片（Wikimedia Commons，均已裁剪）', { x: 7.2, y: 2.05, w: 5.3, h: 0.38, fontSize: 14, bold: true, color: C.pinkDeep, valign: 'middle' });
    const used = [...new Set(USED_PHOTOS)];
    const byArtist = {};
    used.forEach((k) => {
      const p = PHOTOS[k];
      const id = `${p.artist} · ${p.license}`;
      (byArtist[id] = byArtist[id] || []).push(PHOTO_NAMES[k] || k);
    });
    const runs = [];
    Object.entries(byArtist).forEach(([who, names]) => {
      runs.push({ text: who + '：', options: { bold: true, color: C.navy } });
      runs.push({ text: names.join('、'), options: { color: C.text, breakLine: true } });
    });
    runs.push({ text: '许可协议：CC BY-SA 4.0（creativecommons.org/licenses/by-sa/4.0）。每张原图的文件名与链接见本页备注。', options: { color: C.sub, breakLine: true } });
    runs.push({ text: '卡通插画（小飞机“小荷号”、云朵、荷花、图标）均为本组原创绘制。', options: { color: C.sub } });
    const rt = s.text(runs, { x: 7.2, y: 2.5, w: 5.35, h: 3.65, fontSize: 11, valign: 'top', paraSpaceAfter: 6, lineSpacingMultiple: 1.15 });
    const guide = s.art('planeWink', 0.7, 6.38, 1.0);
    const say = bubble(s, '每个数字都有出处，欢迎大家按来源去核对～', 1.95, 6.42, 5.2, 0.5, { size: 12, bold: true, tail: [-53500, 0] });
    s.anim(guide, 'fly', 1100, 500, { dir: 'left' });
    s.anim(say, 'zoom', 1450, 400);
    bob(s, guide, -0.012, 1500, 1800);
    [left, lh, lt].forEach((x, i) => s.anim(x, i === 0 ? 'rise' : 'fade', 300 + i * 150, 450));
    [right, rh, rt].forEach((x, i) => s.anim(x, i === 0 ? 'rise' : 'fade', 600 + i * 150, 450));
    footer(s, '');
    s.notes(
      '【资料来源】这一页列出了我们用到的数据和图片来源。数据主要来自中国民航局的统计公报和杭州萧山国际机场官网，图片来自 Wikimedia Commons，按知识共享许可协议署名使用。\n\n' +
        '图片原始文件（Wikimedia Commons）：\n' +
        used.map((k) => `${PHOTO_NAMES[k]}：File:${PHOTOS[k].title}，作者 ${PHOTOS[k].artist}，${PHOTOS[k].license}，${PHOTOS[k].page}`).join('\n')
    );
  }

  // ===== 18 致谢 =========================================================================
  {
    const s = newSlide('prestige', 'bgCover');
    const sun = s.art('sun', 11.4, 0.4, 1.3);
    const c1 = s.art('cloudSmile', 1.0, 0.7, 1.9);
    const c2 = s.art('cloud', 5.7, 0.2, 1.3);
    const end = [10.75, 2.5];
    const r = route([[-0.8, 3.7], [1.4, 3.3], [4.0, 1.25], [8.3, 1.2], end], { x: 0, y: 0.9 }, end);
    const trail = s.image(await trailArt('thanks', Math.round(W * PX), Math.round(3.1 * PX), r.svg, A.P.blueSoft, 7, '3 22'), 0, 0.9, W, 3.1, { alt: '航迹' });
    const title = s.text('谢谢聆听', { x: 3.07, y: 1.75, w: 7.2, h: 1.5, fontSize: 80, bold: true, color: C.navy, align: 'center', valign: 'middle', charSpacing: 6 });
    const sub = s.text('THANK YOU  ·  欢迎老师和同学们提问', { x: 1.6, y: 3.3, w: 10.1, h: 0.55, fontSize: 20, bold: true, color: C.blueDeep, align: 'center', valign: 'middle' });
    const wish = s.text('祝大家每一次出发，都一路晴空', {
      x: 4.4, y: 4.1, w: 4.5, h: 0.5, shape: 'roundRect', rectRadius: 0.25,
      fill: { color: C.pink }, color: C.white, bold: true, fontSize: 15, align: 'center', valign: 'middle', shadow: shadow(),
    });
    const box = card(s, 4.05, 5.05, 5.2, 0.95, { radius: 0.3 });
    const lab = s.text('汇报人', { x: 4.3, y: 5.3, w: 1.3, h: 0.46, shape: 'roundRect', rectRadius: 0.23, fill: { color: C.bluePale }, color: C.blueDeep, bold: true, fontSize: 15, align: 'center', valign: 'middle' });
    s.text(' ', { x: 5.8, y: 5.28, w: 3.2, h: 0.46, fontSize: 16, color: C.navy, valign: 'middle' });
    const line = s.shape('line', { x: 5.8, y: 5.78, w: 3.2, h: 0, line: { color: C.blueSoft, width: 1.25, dashType: 'dash' } });
    const plane = s.art('planeOpen', end[0] - 0.9, end[1] - 0.57, 1.8);
    const leaf = s.art('leaf', 0.5, 5.9, 1.9);
    const lotus = s.art('lotus', 1.3, 5.55, 1.9);
    const bud = s.art('bud', 0.75, 5.0, 0.48);
    const bag = s.art('suitcase', 11.55, 5.55, 1.0);
    const h1 = s.art('heart', 9.2, 4.0, 0.3);
    const st1 = s.art('star', 3.05, 3.0, 0.36);
    const st2 = s.art('starPink', 10.1, 3.7, 0.3);

    s.anim(trail, 'wipe', 300, 2200, { dir: 'left' });
    s.anim(plane, 'fade', 300, 250);
    s.anim(plane, 'path', 300, 2600, { path: r.motion });
    bob(s, plane, -0.016, 1600, 3000);
    s.anim(title, 'zoom', 900, 800);
    s.anim(sub, 'rise', 1500, 550);
    s.anim(wish, 'zoom', 1900, 450);
    s.anim(box, 'rise', 2200, 500);
    s.anim(lab, 'zoom', 2400, 350);
    s.anim(line, 'wipe', 2500, 400, { dir: 'left' });
    [leaf, lotus, bud, bag].forEach((x, i) => s.anim(x, 'zoom', 1200 + i * 200, 500));
    s.anim(lotus, 'pulse', 2600, 1400, { by: 1.06, repeat: 'forever' });
    bob(s, bag, -0.012, 1400, 2600);
    [h1, st1, st2].forEach((x, i) => {
      s.anim(x, 'zoom', 2000 + i * 200, 350);
      s.anim(x, 'pulse', 2500 + i * 250, 900, { by: 1.35, repeat: 'forever' });
    });
    s.anim(sun, 'pulse', 0, 1800, { by: 1.06, repeat: 'forever' });
    drift(s, c1, 0.035, 8000);
    drift(s, c2, -0.03, 7000);
    s.notes('【致谢】我们的汇报到这里就结束了，谢谢各位老师和同学的聆听，欢迎大家提问和指正。\n（提示：“汇报人”右侧虚线上方可以直接点击填写。）');
  }

  // ---------- write ----------------------------------------------------------------------
  if (specs.length !== TOTAL) throw new Error(`expected ${TOTAL} slides, built ${specs.length}`);
  const raw = await pptx.write({ outputType: 'nodebuffer' });
  const zip = await JSZip.loadAsync(raw);
  for (let i = 0; i < specs.length; i++) {
    const name = `ppt/slides/slide${i + 1}.xml`;
    const xml = await zip.file(name).async('string');
    zip.file(name, decorateSlide(xml, specs[i]));
  }
  const out = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 9 } });
  fs.writeFileSync(OUT_FILE, out);
  console.log(`${path.relative(__dirname, OUT_FILE)}  ${specs.length} slides  ${(out.length / 1048576).toFixed(2)} MB`);
}

const trailCache = {};
async function trailArt(name, w, h, d, color, width, dash) {
  if (!trailCache[name]) {
    const file = path.join(IMG, `trail-${name}.png`);
    await sharp(Buffer.from(A.trail(w, h, d, color || A.P.blueSoft, width || 6, dash || '3 18')))
      .png({ compressionLevel: 9, palette: true })
      .toFile(file);
    trailCache[name] = file;
  }
  return trailCache[name];
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
