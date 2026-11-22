// Builds the student presentation 杭州萧山国际机场 (one airport, 20 slides, light "lotus" theme).
// PptxGenJS lays out the slides; lib/fx.js adds gradients, picture frames, transitions and
// animations; lib/art-hgh.js draws the mascot and the other decorations.
//
// Every figure on the slides is listed with its source in FACTS-hgh.md and on the "参考资料" slide.

const fs = require('fs');
const path = require('path');
const PptxGenJS = require('pptxgenjs');
const JSZip = require('jszip');
const sharp = require('sharp');
const { gradient, decorateSlide } = require('./lib/fx');
const art = require('./lib/art-hgh');

const ROOT = __dirname;
const PHOTOS = path.join(ROOT, 'photos');
const IMG = path.join(ROOT, 'img-hgh');
const OUT = path.join(ROOT, 'out');
const OUT_FILE = process.argv[2] ? path.resolve(process.argv[2]) : path.join(OUT, '杭州萧山国际机场-期末汇报.pptx');

const W = 13.333;
const H = 7.5;
const PX = 144;
const MX = 0.75;
const TOTAL = 20;
const TOC_PAGE = 2;

const C = {
  ink: '21494C',
  text: '3A5759',
  sub: '6B8888',
  muted: '9AB0AE',
  teal: '1FA79B',
  tealDark: '11736D',
  mint: 'BFEBE1',
  mintLight: 'E9F8F4',
  pink: 'FF8FAB',
  pinkDark: 'EA6189',
  pinkLight: 'FFE6EC',
  yellow: 'FFC94D',
  orange: 'FF9F45',
  cream: 'FFF5DD',
  sky: '8FD8F2',
  skyDark: '4BB6DE',
  skyLight: 'E2F4FB',
  white: 'FFFFFF',
  line: 'D3EBE6',
};
const FONT = { cn: 'Microsoft YaHei', light: 'Microsoft YaHei Light', num: 'Arial Rounded MT Bold' };

const photoInfo = Object.fromEntries(
  JSON.parse(fs.readFileSync(path.join(ROOT, 'photos-hgh.json'), 'utf8')).map((p) => [p.key, p])
);
const usedPhotos = [];

// ---------- pictures --------------------------------------------------------------------

/** Crops a source photo to w x h inches around a focus point (fx, fy in 0..1). */
async function crop(key, wIn, hIn, { fx = 0.5, fy = 0.5 } = {}) {
  if (!usedPhotos.includes(key)) usedPhotos.push(key);
  const outW = Math.round(wIn * PX);
  const outH = Math.round(hIn * PX);
  const src = path.join(PHOTOS, key + '.jpg');
  const meta = await sharp(src).metadata();
  const k = Math.max(outW / meta.width, outH / meta.height);
  const cw = Math.min(meta.width, Math.round(outW / k));
  const ch = Math.min(meta.height, Math.round(outH / k));
  const left = Math.round((meta.width - cw) * fx);
  const top = Math.round((meta.height - ch) * fy);
  const file = path.join(IMG, `${key}_${outW}x${outH}.jpg`);
  await sharp(src)
    .extract({ left, top, width: cw, height: ch })
    .resize(outW, outH, { kernel: 'lanczos3' })
    .sharpen({ sigma: 0.5 })
    .jpeg({ quality: 86, mozjpeg: true })
    .toFile(file);
  return file;
}

const drawn = {
  plane: (face = 'smile', flip = false) => art.render(IMG, `plane-${face}${flip ? '-l' : ''}`, art.plane({ face, flip }), 640),
  cloud: (face = false) => art.render(IMG, face ? 'cloud-face' : 'cloud', art.cloud({ face }), 420),
  lotus: () => art.render(IMG, 'lotus', art.lotus(), 480),
  bud: () => art.render(IMG, 'bud', art.lotusBud(), 240),
  osmanthus: () => art.render(IMG, 'osmanthus', art.osmanthus(), 450),
  sun: () => art.render(IMG, 'sun', art.sun(), 300),
  sparkle: (color = art.P.yellow) => art.render(IMG, 'sparkle-' + color.slice(1), art.sparkle(color), 120),
  heart: () => art.render(IMG, 'heart', art.heart(), 120),
  paper: () => art.render(IMG, 'paper', art.paperPlane(), 260),
  suitcase: () => art.render(IMG, 'suitcase', art.suitcase(), 260),
  bars: () => art.render(IMG, 'bars', art.barcode(), 460),
  icon: (name) => art.render(IMG, 'icon-' + name, art.icon(name), 220),
  bg: (kind) => art.render(IMG, 'bg-' + kind, art.background(kind), 1920),
};

// ---------- slide helpers ---------------------------------------------------------------

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: W, height: H });
pptx.layout = 'WIDE';
pptx.title = '杭州萧山国际机场';
pptx.subject = '期末作业汇报：机场介绍';
pptx.author = '';
pptx.company = '';

const specs = [];
const run = (text, options = {}) => ({ text, options });
const numRun = (text, options = {}) => run(text, { fontFace: FONT.num, ...options });
const pad2 = (n) => String(n).padStart(2, '0');
const CJK = /[\u2e80-\u9fff\uff00-\uffef]/;
// PptxGenJS rewrites the shadow object it is given, so every shape needs a fresh one.
const shadow = (opacity = 0.16, blur = 10, offset = 3) => ({ type: 'outer', color: '1F6F6A', opacity, blur, offset, angle: 90 });

/** Splits a string so digits and Latin letters use the rounded numeral face and Chinese uses YaHei. */
function mixed(text, options = {}, cjkOptions = {}) {
  const parts = text.match(/[\u2e80-\u9fff\uff00-\uffef]+|[^\u2e80-\u9fff\uff00-\uffef]+/g) || [];
  return parts.map((part) => (CJK.test(part) ? run(part, { ...options, fontFace: FONT.cn, ...cjkOptions }) : numRun(part, options)));
}

async function newSlide(transition, bg = 'main') {
  const slide = pptx.addSlide();
  slide.background = { path: await drawn.bg(bg) };
  const spec = { transition, effects: [], frames: {}, bulletColor: C.teal };
  specs.push(spec);
  let counter = 0;
  const auto = (prefix) => `${prefix} ${++counter}`;
  const api = {
    spec,
    shape(type, x, y, w, h, fill, o = {}) {
      const name = o.name || auto('Shape');
      const { name: _n, ...rest } = o;
      slide.addShape(type, {
        x, y, w, h,
        fill: typeof fill === 'string' ? { color: fill } : fill,
        ...rest,
        objectName: name,
      });
      return name;
    },
    rect(x, y, w, h, fill, o = {}) {
      return api.shape('rect', x, y, w, h, fill, o);
    },
    /** Rounded rectangle; radius in inches. */
    round(x, y, w, h, fill, o = {}) {
      const { radius = 0.18, ...rest } = o;
      return api.shape('roundRect', x, y, w, h, fill, { rectRadius: Math.min(radius, Math.min(w, h) / 2), ...rest });
    },
    /** White card with a soft shadow. */
    card(x, y, w, h, o = {}) {
      const { fill = C.white, ...rest } = o;
      return api.round(x, y, w, h, fill, { shadow: shadow(), ...rest });
    },
    dot(cx, cy, d, fill, o = {}) {
      return api.shape('ellipse', cx - d / 2, cy - d / 2, d, d, fill, o);
    },
    /** Straight line between two points; dashed by default. */
    line(x1, y1, x2, y2, color, o = {}) {
      const { name: given, width = 2, dash = 'dash', ...rest } = o;
      const name = given || auto('Line');
      slide.addShape('line', {
        x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
        flipV: (x2 - x1) * (y2 - y1) < 0,
        line: { color, width, dashType: dash },
        ...rest,
        objectName: name,
      });
      return name;
    },
    text(content, x, y, w, h, o = {}) {
      const { name: given, ...options } = o;
      const name = given || auto('Text');
      slide.addText(content, {
        x, y, w, h,
        margin: 0,
        fontFace: FONT.cn,
        fontSize: 12,
        color: C.text,
        valign: 'top',
        lang: 'zh-CN',
        ...options,
        objectName: name,
      });
      return name;
    },
    /** Rounded label with centred text. */
    pill(content, x, y, w, h, fill, o = {}) {
      const { color = C.white, fontSize = 11, ...rest } = o;
      return api.text(content, x, y, w, h, {
        shape: 'roundRect', rectRadius: h / 2, fill: { color: fill }, color, fontSize, bold: true, align: 'center', valign: 'middle', ...rest,
      });
    },
    image(file, x, y, w, h, name, alt, o = {}) {
      slide.addImage({ path: file, x, y, w, h, objectName: name, altText: alt, ...o });
      return name;
    },
    /** A cropped photo in a rounded white frame. */
    async photo(key, x, y, w, h, name, o = {}) {
      const { fx = 0.5, fy = 0.5, radius = 6, border = 4, frame = true, ...rest } = o;
      api.image(await crop(key, w, h, { fx, fy }), x, y, w, h, name, photoInfo[key].what, rest);
      if (frame) spec.frames[name] = { radius, border, shadow: true };
      return name;
    },
    async icon(iconName, x, y, size, name) {
      return api.image(await drawn.icon(iconName), x, y, size, size, name || auto('Icon'), '图标');
    },
    chart(type, data, options, name) {
      slide.addChart(type, data, { ...options, objectName: name });
      return name;
    },
    /** Invisible click target that jumps to another slide. */
    link(x, y, w, h, page, tooltip, name) {
      return api.rect(x, y, w, h, { color: 'FFFFFF', transparency: 100 }, { name, hyperlink: { slide: page, tooltip } });
    },
    anim(target, type, delay = 0, dur = 600, extra = {}) {
      spec.effects.push({ target, type, delay, dur, ...extra });
    },
    notes(text) {
      slide.addNotes(text);
    },
  };
  return api;
}

function footer(s, page, { onPhoto = false } = {}) {
  s.text('杭州萧山国际机场  ·  期末作业汇报', MX, 7.1, 4, 0.22, { fontSize: 9, color: C.sub, name: 'Footer title' });
  s.pill('返回目录', W - MX - 2.05, 7.05, 0.95, 0.28, C.white, {
    color: C.tealDark, fontSize: 9, name: 'Footer contents link', line: { color: C.teal, width: 0.75 },
    hyperlink: { slide: TOC_PAGE, tooltip: '返回目录' },
  });
  s.text(`${pad2(page)} / ${TOTAL}`, W - MX - 0.95, 7.1, 0.95, 0.22, {
    fontSize: 10, color: onPhoto ? C.white : C.tealDark, align: 'right', fontFace: FONT.num, name: 'Footer page',
  });
}

/** Section tag, title and the little mascot that bobs beside it. */
async function heading(s, no, part, title, { face = 'smile', mascot = true } = {}) {
  s.pill(no ? `${no}  ${part}` : part, MX, 0.52, (no ? 0.5 : 0.36) + 0.2 * [...part].length, 0.36, C.teal, { fontSize: 11.5, name: 'Kicker' });
  s.text(title, MX, 1.02, W - 2 * MX - (mascot ? 1.9 : 0), 0.62, { fontSize: 28, bold: true, color: C.ink, name: 'Title' });
  s.round(MX, 1.69, 0.9, 0.07, C.pink, { radius: 0.035, name: 'Title bar' });
  s.dot(MX + 1.07, 1.725, 0.07, C.yellow, { name: 'Title dot' });
  s.anim('Kicker', 'pop', 0, 500);
  s.anim('Title', 'rise', 150, 650);
  s.anim('Title bar', 'wipe', 450, 450);
  s.anim('Title dot', 'pop', 800, 350);
  if (mascot) {
    s.image(await drawn.plane(face, true), W - MX - 1.55, 0.42, 1.55, 1.107, 'Mascot', '小飞机吉祥物');
    s.anim('Mascot', 'fly', 200, 900, { dir: 'right' });
    s.anim('Mascot', 'float', 1200, 1800, { dy: -0.018 });
  }
}

function creditText(keys) {
  const people = [];
  for (const key of [].concat(keys)) {
    const p = photoInfo[key];
    const label = `${p.artist}（${p.license}）`;
    if (!people.includes(label)) people.push(label);
  }
  return '图：' + people.join('、');
}

function credit(s, keys, x, y, w, o = {}) {
  return s.text(creditText(keys), x, y, w, 0.2, { fontSize: 8, color: C.muted, name: o.name || 'Photo credit', align: o.align || 'left' });
}

/** Icon, bold label and one or two lines of text on a white card. */
async function infoCard(s, name, iconName, label, lines, x, y, w, h, delay, { tint = C.white } = {}) {
  s.card(x, y, w, h, { name: name + ' card', fill: tint });
  await s.icon(iconName, x + 0.2, y + (h - 0.7) / 2, 0.7, name + ' icon');
  s.text(label, x + 1.08, y + 0.17, w - 1.25, 0.26, { fontSize: 11, color: C.teal, bold: true, name: name + ' label' });
  s.text(lines, x + 1.08, y + 0.45, w - 1.25, h - 0.55, { fontSize: 13.5, color: C.ink, bold: true, lineSpacingMultiple: 1.08, name: name + ' text' });
  s.anim(name + ' card', 'rise', delay, 500);
  s.anim(name + ' icon', 'pop', delay + 150, 450);
  s.anim(name + ' label', 'fade', delay + 200, 400);
  s.anim(name + ' text', 'fade', delay + 300, 450);
}

/** Speech bubble for the mascot's small facts. */
function bubble(s, name, text, x, y, w, h, { fill = C.cream, tail = 'left', fontSize = 12 } = {}) {
  s.round(x, y, w, h, fill, { radius: 0.22, name: name + ' bubble', line: { color: C.yellow, width: 1.25 } });
  if (tail !== 'none') {
    const tx = tail === 'left' ? x - 0.12 : x + w - 0.1;
    s.shape('triangle', tx, y + h / 2 - 0.11, 0.22, 0.22, fill, { name: name + ' tail', rotate: tail === 'left' ? 270 : 90, line: { color: fill, width: 0 } });
  }
  s.text(text, x + 0.22, y + 0.1, w - 0.44, h - 0.2, { fontSize, color: C.ink, valign: 'middle', lineSpacingMultiple: 1.12, name: name + ' text' });
}

/** Photo on a white card with a caption, tilted like an instant photo. */
async function polaroid(s, name, key, caption, cx, cy, w, angle, { fx = 0.5, fy = 0.5, ratio = 0.72 } = {}) {
  const m = 0.13;
  const pw = w - 2 * m;
  const ph = pw * ratio;
  const capH = 0.46;
  const h = m + ph + capH;
  const a = (angle * Math.PI) / 180;
  const at = (dy) => [cx - dy * Math.sin(a), cy + dy * Math.cos(a)];
  const rot = (angle + 360) % 360;
  s.rect(cx - w / 2, cy - h / 2, w, h, C.white, { name: name + ' card', rotate: rot, shadow: shadow(0.22, 12, 4) });
  const [px, py] = at(-(capH - m) / 2);
  s.image(await crop(key, pw, ph, { fx, fy }), px - pw / 2, py - ph / 2, pw, ph, name + ' photo', photoInfo[key].what, { rotate: rot });
  const [tx, ty] = at(h / 2 - capH / 2 - 0.02);
  s.text(caption, tx - pw / 2, ty - 0.15, pw, 0.3, { fontSize: 11, color: C.ink, bold: true, align: 'center', valign: 'middle', rotate: rot, name: name + ' caption' });
  return [name + ' card', name + ' photo', name + ' caption'];
}

// ---------- slides ----------------------------------------------------------------------

async function slideCover() {
  const s = await newSlide('fade', 'cover');
  s.image(await drawn.sun(), 11.75, 0.3, 1.05, 1.05, 'Sun', '太阳');
  s.image(await drawn.cloud(true), 4.55, 0.22, 1.4, 0.79, 'Cloud a', '云朵');
  s.image(await drawn.cloud(), 0.2, 2.75, 1.2, 0.68, 'Cloud b', '云朵');

  s.pill('期末作业汇报  ·  机场介绍', MX, 0.72, 2.9, 0.4, C.teal, { fontSize: 12, name: 'Tag' });
  s.text('杭州萧山国际机场', MX, 1.32, 7.6, 1.0, { fontSize: 54, bold: true, color: C.ink, name: 'Cover title' });
  s.text('HANGZHOU XIAOSHAN INTERNATIONAL AIRPORT', MX + 0.04, 2.42, 7.4, 0.32, {
    fontSize: 13, color: C.teal, fontFace: FONT.num, charSpacing: 4, name: 'Cover title en',
  });
  s.text('从西湖荷花里起飞的“浙江第一空中门户”', MX, 2.92, 7.4, 0.46, { fontSize: 21, color: C.text, name: 'Cover subtitle' });

  // Boarding pass with the blanks the presenter fills in: type straight after the colon.
  const bx = MX, by = 3.82, bw = 6.95, bh = 2.3, stub = 1.75;
  s.card(bx, by, bw, bh, { name: 'Pass card', radius: 0.22 });
  s.round(bx, by, bw, 0.5, C.teal, { radius: 0.22, name: 'Pass band' });
  s.rect(bx, by + 0.3, bw, 0.2, C.teal, { name: 'Pass band square' });
  s.text([numRun('BOARDING PASS', { charSpacing: 3 }), run('   登机牌')], bx + 0.3, by + 0.1, 3.6, 0.3, { fontSize: 11.5, color: C.white, bold: true, valign: 'middle', name: 'Pass band text' });
  s.text('期末汇报专用航班', bx + bw - stub - 2.3, by + 0.1, 2.1, 0.3, { fontSize: 10, color: 'D9F5F0', align: 'right', valign: 'middle', name: 'Pass band note' });
  s.line(bx + bw - stub, by + 0.62, bx + bw - stub, by + bh - 0.14, C.mint, { width: 1.5, name: 'Pass tear line' });
  const fields = [
    ['汇报人：', bx + 0.32, by + 0.78, 2.25],
    ['制作人：', bx + 2.75, by + 0.78, 2.2],
    ['小组成员：', bx + 0.32, by + 1.5, 4.63],
  ];
  fields.forEach(([label, x, y, w], i) => {
    const key = label.replace('：', '');
    const labelW = 0.2 * [...label].length + 0.04;
    s.text(label, x, y, w, 0.36, { fontSize: 14, color: C.ink, bold: true, valign: 'middle', name: `Field ${key}` });
    s.rect(x + labelW, y + 0.38, w - labelW, 0.014, C.muted, { name: `Field line ${i + 1}` });
    s.anim(`Field ${key}`, 'fade', 2300 + i * 150, 450);
    s.anim(`Field line ${i + 1}`, 'wipe', 2350 + i * 150, 450);
  });
  const sx = bx + bw - stub;
  s.text('HGH', sx, by + 0.66, stub, 0.62, { fontSize: 32, color: C.teal, fontFace: FONT.num, bold: true, align: 'center', valign: 'middle', name: 'Pass code' });
  s.text('杭州萧山', sx, by + 1.26, stub, 0.26, { fontSize: 11, color: C.sub, align: 'center', name: 'Pass code cn' });
  // Decorative stripes only; they encode nothing.
  s.image(await drawn.bars(), sx + 0.3, by + 1.62, stub - 0.6, 0.42, 'Pass bars', '装饰条纹');

  // Photo of the lotus skylight in a rounded frame.
  const px = 8.3, py = 1.1, pw = 4.3, ph = 4.75;
  await s.photo('hgh_lotus', px, py, pw, ph, 'Cover photo', { fx: 0.5, fy: 0.4, radius: 9, border: 6 });
  credit(s, 'hgh_lotus', px, py + ph + 0.1, pw, { align: 'right' });
  s.image(await drawn.lotus(), px - 0.75, py + ph - 1.0, 1.6, 1.25, 'Lotus', '荷花');
  s.image(await drawn.sparkle(), px + pw - 0.2, py - 0.3, 0.5, 0.5, 'Sparkle a', '星光');
  s.image(await drawn.sparkle(art.P.pink), px - 0.42, py + 0.5, 0.34, 0.34, 'Sparkle b', '星光');
  s.image(await drawn.sparkle(), 7.35, 3.1, 0.28, 0.28, 'Sparkle c', '星光');

  // The mascot flies in from the left along a dotted trail and then bobs in place.
  const mx = 6.85, my = 0.2, mw = 2.3, mh = mw * (300 / 420);
  s.line(0.2, my + mh * 0.62, mx + 0.25, my + mh * 0.62, C.teal, { width: 2.25, dash: 'sysDash', name: 'Trail' });
  s.image(await drawn.plane('happy'), mx, my, mw, mh, 'Mascot', '小飞机吉祥物');

  s.anim('Sun', 'pop', 200, 700);
  s.anim('Sun', 'teeter', 1000, 1600, { deg: 10 });
  s.anim('Cloud a', 'fade', 0, 800);
  s.anim('Cloud a', 'float', 800, 4200, { dx: 0.02, dy: 0 });
  s.anim('Cloud b', 'fade', 200, 800);
  s.anim('Cloud b', 'float', 900, 5000, { dx: 0.025, dy: 0 });
  s.anim('Tag', 'pop', 300, 550);
  s.anim('Cover title', 'rise', 600, 800);
  s.anim('Cover title en', 'wipe', 1100, 900);
  s.anim('Cover subtitle', 'fade', 1500, 700);
  s.anim('Cover photo', 'zoom', 500, 900);
  s.anim('Photo credit', 'fade', 1300, 500);
  s.anim('Lotus', 'pop', 1400, 600);
  s.anim('Sparkle a', 'pop', 1700, 400);
  s.anim('Sparkle a', 'pulse', 2200, 700, { by: 1.35, repeat: 'forever' });
  s.anim('Sparkle b', 'pop', 1900, 400);
  s.anim('Sparkle b', 'pulse', 2500, 800, { by: 1.35, repeat: 'forever' });
  s.anim('Sparkle c', 'pop', 2100, 400);
  s.anim('Sparkle c', 'pulse', 2800, 900, { by: 1.4, repeat: 'forever' });
  for (const n of ['Pass card', 'Pass band', 'Pass band square']) s.anim(n, 'rise', 1800, 600);
  for (const n of ['Pass band text', 'Pass band note', 'Pass tear line']) s.anim(n, 'fade', 2150, 400);
  s.anim('Pass code', 'pop', 2500, 500);
  s.anim('Pass code cn', 'fade', 2700, 400);
  s.anim('Pass bars', 'wipe', 2800, 500);
  s.anim('Trail', 'wipe', 3000, 1500);
  s.anim('Mascot', 'fly', 3000, 1700, { dir: 'left' });
  s.anim('Trail', 'fadeOut', 4800, 700);
  s.anim('Mascot', 'float', 4800, 1700, { dy: -0.025 });
  s.notes('各位老师、同学好，我们小组汇报的题目是《杭州萧山国际机场》。它是浙江省第一空中门户，最新的 T4 航站楼把西湖荷花的形象建进了大厅里，所以我们的副标题叫“从西湖荷花里起飞”。');
}

const CONTENTS = [
  ['01', '初识机场', '机场名片 · 区位与定位', 3, C.teal],
  ['02', '成长足迹', '时间线 · 三次扩建', 5, C.pink],
  ['03', '荷花航站楼', '设计 · 实景 · 绿色智慧', 7, C.orange],
  ['04', '数据看实力', '成绩单 · 客流 · 航线 · 货运', 10, C.skyDark],
  ['05', '四通八达', '地铁 · 交通中心 · 高铁', 14, C.tealDark],
  ['06', '服务与未来', '亚运门户 · 暖心服务 · 规划', 15, C.pinkDark],
];

async function slideContents() {
  const s = await newSlide('airplane');
  s.pill('CONTENTS', MX, 0.52, 1.5, 0.36, C.teal, { fontSize: 11.5, fontFace: FONT.num, charSpacing: 3, name: 'Kicker' });
  s.text('目录  ·  本次航程的六个经停站', MX, 1.02, 9, 0.62, { fontSize: 28, bold: true, color: C.ink, name: 'Title' });
  s.round(MX, 1.69, 0.9, 0.07, C.pink, { radius: 0.035, name: 'Title bar' });
  s.dot(MX + 1.07, 1.725, 0.07, C.yellow, { name: 'Title dot' });
  s.anim('Kicker', 'pop', 0, 500);
  s.anim('Title', 'rise', 150, 650);
  s.anim('Title bar', 'wipe', 450, 450);
  s.anim('Title dot', 'pop', 800, 350);

  const x0 = 1.65, step = (W - 2 * 1.65) / 5;
  const nodes = CONTENTS.map((_, i) => [x0 + i * step, i % 2 === 0 ? 3.15 : 4.55]);
  for (let i = 0; i < nodes.length - 1; i++) {
    const n = s.line(nodes[i][0], nodes[i][1], nodes[i + 1][0], nodes[i + 1][1], C.teal, { width: 2.25, dash: 'sysDash', name: `Route ${i + 1}` });
    s.anim(n, 'wipe', 1100 + i * 380, 380);
  }
  for (const [i, [no, title, desc, page, color]] of CONTENTS.entries()) {
    const [cx, cy] = nodes[i];
    const up = i % 2 === 0;
    const delay = 900 + i * 380;
    s.dot(cx, cy, 1.02, C.white, { name: `Stop ${no} ring`, shadow: shadow(0.2, 8, 2) });
    s.text(no, cx - 0.41, cy - 0.41, 0.82, 0.82, {
      shape: 'ellipse', fill: { color }, color: C.white, fontSize: 24, bold: true, fontFace: FONT.num, align: 'center', valign: 'middle', name: `Stop ${no} number`,
    });
    const ty = up ? cy + 0.68 : cy - 1.52;
    s.text(title, cx - 1.1, ty, 2.2, 0.4, { fontSize: 19, bold: true, color: C.ink, align: 'center', name: `Stop ${no} title` });
    s.text(desc, cx - 1.1, ty + 0.44, 2.2, 0.4, { fontSize: 10.5, color: C.sub, align: 'center', name: `Stop ${no} text` });
    s.link(cx - 1.1, Math.min(cy - 0.55, ty), 2.2, 1.95, page, `跳转到：${title}`, `Stop ${no} link`);
    s.anim(`Stop ${no} ring`, 'pop', delay, 450);
    s.anim(`Stop ${no} number`, 'pop', delay + 60, 450);
    s.anim(`Stop ${no} title`, 'rise', delay + 200, 500);
    s.anim(`Stop ${no} text`, 'fade', delay + 350, 450);
  }

  // The mascot flies the whole route once.
  const mw = 1.5, mh = mw * (300 / 420);
  const start = [0.98, 2.42];
  s.image(await drawn.plane('smile'), start[0] - mw / 2, start[1] - mh / 2, mw, mh, 'Mascot', '小飞机吉祥物');
  const hops = nodes.map(([x, y]) => [x, y - 0.75]).concat([[W - 0.9, 3.3]]);
  const pathText = 'M 0 0 ' + hops.map(([x, y]) => `L ${((x - start[0]) / W).toFixed(4)} ${((y - start[1]) / H).toFixed(4)}`).join(' ') + ' E';
  s.anim('Mascot', 'fly', 500, 700, { dir: 'left' });
  s.anim('Mascot', 'path', 1300, 5200, { path: pathText });

  s.image(await drawn.osmanthus(), W - 2.05, 5.05, 1.5, 1.1, 'Osmanthus', '桂花枝');
  s.anim('Osmanthus', 'fade', 2600, 700);
  s.pill('放映时点击任意一站，可以直接跳到对应章节', (W - 4.6) / 2, 5.98, 4.6, 0.34, C.cream, { color: C.ink, fontSize: 10.5, bold: false, name: 'Hint' });
  s.anim('Hint', 'fade', 3400, 600);
  s.text(`${pad2(2)} / ${TOTAL}`, W - MX - 0.95, 7.1, 0.95, 0.22, { fontSize: 10, color: C.tealDark, align: 'right', fontFace: FONT.num, name: 'Footer page' });
  s.text('杭州萧山国际机场  ·  期末大作业汇报', MX, 7.1, 4, 0.22, { fontSize: 9, color: C.sub, name: 'Footer title' });
  s.notes('这次汇报我们把它设计成一趟航程，一共六站：先认识机场，再看它二十多年的成长足迹，然后重点参观像荷花一样的 T4 航站楼，接着用数据看它的实力，再讲地面交通，最后是服务亮点和未来规划。');
}

async function slideProfile() {
  const s = await newSlide('push');
  await heading(s, '01', '初识机场', '机场名片：一分钟认识萧山机场');
  const pw = 4.75, ph = 3.5, px = MX, py = 2.1;
  await s.photo('hgh_aerial', px, py, pw, ph, 'Profile photo', { fx: 0.45, fy: 0.5 });
  credit(s, 'hgh_aerial', px, py + ph + 0.08, pw);
  s.pill([run('IATA  '), numRun('HGH')], px + 0.25, py + ph - 0.62, 1.55, 0.42, C.teal, { fontSize: 13, name: 'Code IATA' });
  s.pill([run('ICAO  '), numRun('ZSHC')], px + 1.95, py + ph - 0.62, 1.7, 0.42, C.pink, { fontSize: 13, name: 'Code ICAO' });
  s.text('从空中看：两条平行跑道之间就是航站区', px, py + ph + 0.34, pw, 0.3, { fontSize: 11, color: C.sub, name: 'Profile caption' });
  s.anim('Profile photo', 'zoom', 400, 800);
  s.anim('Photo credit', 'fade', 1100, 400);
  s.anim('Code IATA', 'pop', 1200, 450);
  s.anim('Code ICAO', 'pop', 1400, 450);
  s.anim('Profile caption', 'fade', 1500, 500);

  const cards = [
    ['pin', '在哪里', '杭州市萧山区\n距市中心 27 公里'],
    ['calendar', '什么时候通航', '2000 年 12 月 28 日\n一期工程建成通航'],
    ['runway', '飞行区', '4F 级 · 2 条跑道\n长 3600 米 / 3400 米'],
    ['terminal', '航站楼', '4 座（T1—T4）\n总面积 99 万平方米'],
    ['star', '定位', '浙江省第一空中门户\n区域航空枢纽'],
    ['people', '2025 年旅客吞吐量', '5045.9 万人次\n全国第 7 位'],
  ];
  const gx = px + pw + 0.45, gw = (W - MX - gx - 0.25) / 2, gh = 1.22;
  for (const [i, [iconName, label, text]] of cards.entries()) {
    const x = gx + (i % 2) * (gw + 0.25);
    const y = 2.1 + Math.floor(i / 2) * (gh + 0.2);
    await infoCard(s, `Fact ${i + 1}`, iconName, label, text, x, y, gw, gh, 700 + i * 180);
  }
  footer(s, 3);
  s.notes('先看机场名片。萧山机场位于杭州市萧山区，距市中心 27 公里，2000 年 12 月 28 日建成通航。飞行区等级是最高的 4F 级，有两条跑道，分别长 3600 米和 3400 米。现在有 T1 到 T4 四座航站楼，总面积 99 万平方米。它是浙江省第一空中门户，2025 年旅客吞吐量 5045.9 万人次，排全国第 7。');
}

async function slideRole() {
  const s = await newSlide('pageCurl');
  await heading(s, '01', '初识机场', '区位与定位：它的三个“身份”', { face: 'wink' });
  const roles = [
    [C.teal, '浙江省第一空中门户', '中国重要的干线机场，2025 年旅客吞吐量 5045.9 万人次，居全国第 7 位'],
    [C.pink, '华东第二大航空枢纽', '三期投运后规模仅次于上海浦东；2025 年华东地区旅客吞吐量排名第 2'],
    [C.orange, '对外开放的一类航空口岸', '国际定期航班机场、国际航班备降机场，航线通达 5 大洲 35 个国家和地区'],
  ];
  const lw = 6.55;
  roles.forEach(([color, title, text], i) => {
    const y = 2.08 + i * 1.42;
    const n = `Role ${i + 1}`;
    s.card(MX, y, lw, 1.22, { name: n + ' card' });
    s.text(String(i + 1), MX + 0.24, y + 0.21, 0.8, 0.8, {
      shape: 'ellipse', fill: { color }, color: C.white, fontSize: 28, bold: true, fontFace: FONT.num, align: 'center', valign: 'middle', name: n + ' number',
    });
    s.text(title, MX + 1.28, y + 0.2, lw - 1.5, 0.38, { fontSize: 18, bold: true, color: C.ink, name: n + ' title' });
    s.text(text, MX + 1.28, y + 0.63, lw - 1.5, 0.5, { fontSize: 12, color: C.text, lineSpacingMultiple: 1.1, name: n + ' text' });
    const d = 600 + i * 350;
    s.anim(n + ' card', 'slide', d, 550, { offset: -0.05 });
    s.anim(n + ' number', 'pop', d + 150, 500);
    s.anim(n + ' title', 'fade', d + 250, 450);
    s.anim(n + ' text', 'fade', d + 400, 450);
  });

  const px = MX + lw + 0.5, pw = W - MX - px, ph = 2.75;
  await s.photo('hgh_curb', px, 2.08, pw, ph, 'Role photo', { fx: 0.4, fy: 0.55 });
  credit(s, 'hgh_curb', px, 2.08 + ph + 0.08, pw, { align: 'right' });
  s.anim('Role photo', 'zoom', 500, 800);
  s.anim('Photo credit', 'fade', 1200, 400);

  // Distance strip: city centre to the airport.
  const dy = 5.42, dh = 1.08;
  s.card(px, dy, pw, dh, { name: 'Distance card', fill: C.mintLight });
  await s.icon('flower', px + 0.3, dy + 0.14, 0.52, 'Distance icon a');
  await s.icon('plane', px + pw - 0.82, dy + 0.14, 0.52, 'Distance icon b');
  s.line(px + 0.98, dy + 0.4, px + pw - 0.98, dy + 0.4, C.teal, { width: 2, dash: 'sysDash', name: 'Distance line' });
  s.text('杭州市中心', px + 0.06, dy + 0.72, 1.0, 0.24, { fontSize: 9.5, color: C.sub, align: 'center', name: 'Distance from' });
  s.text('萧山机场', px + pw - 1.06, dy + 0.72, 1.0, 0.24, { fontSize: 9.5, color: C.sub, align: 'center', name: 'Distance to' });
  s.text([run('约 '), numRun('27', { fontSize: 20, color: C.tealDark }), run(' 公里')], px + 1.0, dy + 0.5, pw - 2.0, 0.44, { fontSize: 12, color: C.ink, bold: true, align: 'center', valign: 'middle', name: 'Distance text' });
  s.anim('Distance card', 'rise', 1500, 500);
  s.anim('Distance icon a', 'pop', 1700, 400);
  s.anim('Distance line', 'wipe', 1800, 700);
  s.anim('Distance icon b', 'pop', 2400, 400);
  for (const n of ['Distance from', 'Distance to', 'Distance text']) s.anim(n, 'fade', 2500, 450);
  footer(s, 4);
  s.notes('萧山机场有三个身份。第一，它是浙江省第一空中门户，也是中国重要的干线机场。第二，三期工程投运后，它成为华东地区仅次于上海浦东的第二大航空枢纽，2025 年华东地区旅客吞吐量排第 2。第三，它是对外开放的一类航空口岸，航线通达五大洲 35 个国家和地区。机场离杭州市中心大约 27 公里。');
}

const MILESTONES = [
  ['1997.7', '一期动工', '机场一期工程\n正式开工建设', C.teal],
  ['2000.12', '建成通航', '12 月 28 日\n一期工程建成通航', C.pink],
  ['2006.12', '牵手香港机场', '成为内地首家整体\n对外合资的机场', C.orange],
  ['2012.12', '二期投运', '国际航站楼、T3\n和第二跑道全部投运', C.skyDark],
  ['2016', 'G20 杭州峰会', '完成航空运输保障\n年旅客突破 3000 万', C.tealDark],
  ['2019', '跨越 4000 万', '年旅客吞吐量\n突破 4000 万人次', C.pinkDark],
  ['2022.9', '三期投运', 'T4 航站楼和交通中心\n9 月 22 日启用', C.orange],
  ['2025.12', '跨越 5000 万', '年旅客吞吐量\n首次突破 5000 万人次', C.teal],
];

async function slideTimeline() {
  const s = await newSlide('gallery');
  await heading(s, '02', '成长足迹', '成长足迹：从动工到年旅客 5000 万', { mascot: false });
  const axisY = 4.2;
  const x0 = MX + 1.2, step = (W - 2 * MX - 2.4) / (MILESTONES.length - 1);
  s.round(MX, axisY - 0.045, W - 2 * MX, 0.09, C.mint, { radius: 0.045, name: 'Axis' });
  s.round(MX, axisY - 0.045, W - 2 * MX, 0.09, C.teal, { radius: 0.045, name: 'Axis progress' });
  s.anim('Axis', 'fade', 400, 400);
  s.anim('Axis progress', 'wipe', 900, 4600);

  const cw = 2.35, chh = 1.42;
  for (const [i, [when, title, text, color]] of MILESTONES.entries()) {
    const cx = x0 + i * step;
    const up = i % 2 === 0;
    const n = `Stage ${i + 1}`;
    const cy = up ? axisY - 0.56 - chh : axisY + 0.36;
    const left = Math.max(MX, Math.min(W - MX - cw, cx - cw / 2));
    s.line(cx, up ? cy + chh : axisY + 0.14, cx, up ? axisY - 0.14 : cy, color, { width: 1.5, dash: 'sysDot', name: n + ' stem' });
    s.dot(cx, axisY, 0.3, C.white, { name: n + ' ring', line: { color, width: 3 } });
    s.card(left, cy, cw, chh, { name: n + ' card', radius: 0.16 });
    s.pill(when, left + 0.18, cy + 0.16, 0.34 + 0.105 * when.length, 0.28, color, { fontSize: 10.5, fontFace: FONT.num, name: n + ' when' });
    s.text(title, left + 0.2, cy + 0.5, cw - 0.4, 0.3, { fontSize: 14, bold: true, color: C.ink, valign: 'middle', name: n + ' title' });
    s.text(text, left + 0.2, cy + 0.84, cw - 0.4, 0.5, { fontSize: 10.5, color: C.text, lineSpacingMultiple: 1.1, name: n + ' text' });
    const d = 900 + i * 600;
    s.anim(n + ' ring', 'pop', d, 400);
    s.anim(n + ' stem', 'wipe', d + 100, 300, { dir: up ? 'down' : 'up' });
    s.anim(n + ' card', up ? 'drop' : 'rise', d + 250, up ? 700 : 500, { offset: 0.05 });
    s.anim(n + ' when', 'pop', d + 500, 400);
    s.anim(n + ' title', 'fade', d + 550, 400);
    s.anim(n + ' text', 'fade', d + 650, 400);
  }

  // The mascot taxis along the axis as the years go by.
  const mw = 0.74, mh = mw * (300 / 420);
  s.image(await drawn.plane('happy'), MX - 0.15, axisY - mh - 0.02, mw, mh, 'Runner', '小飞机吉祥物');
  s.anim('Runner', 'fade', 500, 400);
  s.anim('Runner', 'path', 900, 4600, { path: `M 0 0 L ${((W - 2 * MX - mw + 0.3) / W).toFixed(4)} 0 E`, ease: false });

  const [card, photo, caption] = await polaroid(s, 'Memo', 'hgh_4000', '4000 万纪念装置', 11.36, 1.1, 1.95, 5, { fx: 0.55, fy: 0.55, ratio: 0.62 });
  for (const n of [card, photo, caption]) s.anim(n, 'drop', 4300, 800, { offset: 0.1 });
  credit(s, 'hgh_4000', 7.9, 1.62, 2.3, { align: 'right' });
  s.anim('Photo credit', 'fade', 5000, 400);
  footer(s, 5);
  s.notes('这是机场的成长时间线。1997 年 7 月一期工程动工，2000 年 12 月 28 日建成通航。2006 年和香港国际机场合资，成为内地首家整体对外合资的机场。2012 年底二期工程全部投运。2016 年完成 G20 杭州峰会的航空保障，这一年旅客突破 3000 万。2019 年突破 4000 万，右上角的照片就是当时航站楼里的纪念装置。2022 年 9 月三期投运，T4 航站楼启用。2025 年 12 月，年旅客吞吐量首次突破 5000 万人次。');
}

async function slidePhases() {
  const s = await newSlide('ferris');
  await heading(s, '02', '成长足迹', '三次扩建：一步一个台阶', { face: 'happy' });
  const phases = [
    ['一期', '1997—2000', C.teal, 'hgh_aerial2', 0.5, 0.5, ['1997 年 7 月动工，占地 7260 余亩', '建成首座航站楼和第一条跑道', '2000 年 12 月 28 日建成通航']],
    ['二期', '2007—2012', C.pink, 'hgh_t3', 0.5, 0.5, ['2007 年 11 月 8 日开工', '新建国际航站楼、第二国内航站楼和第二跑道', '2012 年 12 月 30 日全部投运']],
    ['三期', '2018—2022', C.orange, 'hgh_hall2', 0.5, 0.45, ['T4 航站楼约 72 万平方米，交通中心 48 万平方米', '三期工程总投资 270 亿元', '2022 年 9 月 22 日正式投运']],
  ];
  const gap = 0.3, cw = (W - 2 * MX - 2 * gap) / 3, bottom = 6.5;
  for (const [i, [tag, years, color, key, fx, fy, lines]] of phases.entries()) {
    const x = MX + i * (cw + gap);
    const top = 3.05 - i * 0.48;
    const ph = 1.3 + i * 0.42;
    const n = `Phase ${i + 1}`;
    s.card(x, top, cw, bottom - top, { name: n + ' card', radius: 0.2 });
    await s.photo(key, x + 0.14, top + 0.14, cw - 0.28, ph, n + ' photo', { fx, fy, radius: 7, border: 0, frame: true });
    s.spec.frames[n + ' photo'].shadow = false;
    s.pill(tag, x + 0.26, top + ph - 0.08, 0.86, 0.38, color, { fontSize: 13, name: n + ' tag', shadow: shadow(0.2, 5, 2) });
    s.text(years, x + 1.25, top + ph + 0.24, cw - 1.45, 0.34, { fontSize: 17, color, fontFace: FONT.num, bold: true, valign: 'middle', name: n + ' years' });
    s.text(lines.map((l) => run(l, { bullet: { characterCode: '25CF', indent: 13 }, paraSpaceAfter: 5, breakLine: true })),
      x + 0.26, top + ph + 0.72, cw - 0.5, bottom - top - ph - 0.85, { fontSize: 12, color: C.text, lineSpacingMultiple: 1.1, name: n + ' text' });
    s.text('图：' + ['2023 年航拍的跑道与航站区', 'T3 航站楼候机厅', 'T4 航站楼出发大厅'][i], x + 0.26, bottom - 0.34, cw - 0.5, 0.2, { fontSize: 9, color: C.muted, name: n + ' caption' });
    const d = 600 + i * 550;
    s.anim(n + ' card', 'rise', d, 600, { offset: 0.08 });
    s.anim(n + ' photo', 'fade', d + 250, 500);
    s.anim(n + ' tag', 'pop', d + 450, 450);
    s.anim(n + ' years', 'wipe', d + 550, 450);
    lines.forEach((_, k) => s.anim(n + ' text', 'fade', d + 700 + k * 160, 400, { para: k }));
    s.anim(n + ' caption', 'fade', d + 1200, 400);
  }
  bubble(s, 'Tip', '三期新建建筑面积超过 150 万平方米，是浙江迄今最大的机场建设项目', MX, 2.08, cw + 0.15, 0.78, { tail: 'none', fontSize: 11.5 });
  for (const n of ['Tip bubble', 'Tip text']) s.anim(n, 'pop', 2600, 500);
  s.text(creditText(['hgh_aerial2', 'hgh_t3', 'hgh_hall2']), MX, 6.6, W - 2 * MX, 0.2, { fontSize: 8, color: C.muted, name: 'Photo credit' });
  s.anim('Photo credit', 'fade', 2400, 400);
  footer(s, 6);
  s.notes('机场一共经历了三次大的建设，像上台阶一样越建越大。一期 1997 年动工、2000 年通航，建成了第一座航站楼和第一条跑道。二期 2007 年开工、2012 年底全部投运，新建了国际航站楼、第二国内航站楼和第二跑道。三期 2018 年启动、2022 年 9 月 22 日投运，建成了约 72 万平方米的 T4 航站楼和 48 万平方米的交通中心，总投资 270 亿元，是浙江迄今最大的机场建设项目。');
}

async function slideDesign() {
  const s = await newSlide('ripple');
  const pw = 5.5;
  s.image(await crop('hgh_hall', pw, H, { fx: 0.5, fy: 0.5 }), W - pw, 0, pw, H, 'Design photo', photoInfo.hgh_hall.what);
  s.rect(W - pw, 0, 1.5, H, gradient(0, [
    { pos: 0, color: 'F6FCFB', alpha: 100 },
    { pos: 100, color: 'F6FCFB', alpha: 0 },
  ]), { name: 'Design blend' });
  s.rect(W - pw, H - 0.9, pw, 0.9, gradient(90, [
    { pos: 0, color: '0F3B3A', alpha: 0 },
    { pos: 100, color: '0F3B3A', alpha: 70 },
  ]), { name: 'Design shade' });
  s.text(creditText('hgh_hall'), W - MX - 3, 6.72, 3, 0.2, { fontSize: 8, color: 'E6F4F1', align: 'right', name: 'Photo credit' });
  s.anim('Design photo', 'grow', 0, 12000, { by: 1.05 });

  const tw = W - pw - MX - 0.1;
  s.pill('03  荷花航站楼', MX, 0.52, 1.7, 0.36, C.teal, { fontSize: 11.5, name: 'Kicker' });
  s.text('T4 航站楼：把西湖荷花“种”进大厅', MX, 1.02, tw + 0.6, 0.62, { fontSize: 27, bold: true, color: C.ink, name: 'Title' });
  s.round(MX, 1.69, 0.9, 0.07, C.pink, { radius: 0.035, name: 'Title bar' });
  s.dot(MX + 1.07, 1.725, 0.07, C.yellow, { name: 'Title dot' });
  s.anim('Kicker', 'pop', 0, 500);
  s.anim('Title', 'rise', 150, 650);
  s.anim('Title bar', 'wipe', 450, 450);
  s.anim('Title dot', 'pop', 800, 350);

  s.card(MX, 2.08, tw, 1.3, { name: 'Quote card', fill: C.pinkLight });
  s.image(await drawn.lotus(), MX + 0.2, 2.2, 1.35, 1.055, 'Quote lotus', '荷花');
  s.text('“接天莲叶  ·  出水芙蓉”', MX + 1.7, 2.24, tw - 1.85, 0.52, { fontSize: 22, bold: true, color: C.pinkDark, valign: 'middle', name: 'Quote text' });
  s.text('T4 的设计意象取自西湖荷花，2022 年 9 月 22 日启用', MX + 1.72, 2.82, tw - 1.85, 0.4, { fontSize: 12, color: C.text, name: 'Quote note' });
  s.anim('Quote card', 'rise', 600, 600);
  s.anim('Quote lotus', 'pop', 900, 600);
  s.anim('Quote text', 'wipe', 1100, 900);
  s.anim('Quote note', 'fade', 1700, 500);

  const rows = [
    ['flower', '荷花形立柱天窗', '大厅的立柱像出水的荷花，柱顶的天窗把自然光引进室内'],
    ['bulb', '中心“荷花谷”', '位于航站楼主楼中心，和立柱天窗一起带来良好的通风与采光'],
    ['terminal', '约 72 万平方米', '相当于 T1、T2、T3 面积总和的两倍，设计年旅客吞吐量 5000 万人次'],
  ];
  for (const [i, [iconName, title, text]] of rows.entries()) {
    const y = 3.62 + i * 0.98;
    const n = `Point ${i + 1}`;
    await s.icon(iconName, MX + 0.05, y, 0.72, n + ' icon');
    s.text(title, MX + 1.0, y - 0.02, tw - 1.0, 0.34, { fontSize: 15.5, bold: true, color: C.ink, name: n + ' title' });
    s.text(text, MX + 1.0, y + 0.34, tw - 1.0, 0.5, { fontSize: 12, color: C.text, lineSpacingMultiple: 1.1, name: n + ' text' });
    const d = 2000 + i * 400;
    s.anim(n + ' icon', 'pop', d, 450);
    s.anim(n + ' title', 'slide', d + 120, 450, { offset: -0.03 });
    s.anim(n + ' text', 'fade', d + 280, 450);
  }
  footer(s, 7, { onPhoto: true });
  s.notes('第三站是这次汇报的重点：T4 航站楼。它的设计意象取自西湖荷花，用的是“接天莲叶、出水芙蓉”这八个字。大家看右边的照片，大厅里的立柱像一朵朵出水的荷花，柱顶是天窗，可以把自然光引进来。航站楼主楼中心还有一个“荷花谷”，和这些立柱天窗一起，让室内有良好的通风和采光。T4 总建筑面积约 72 万平方米，相当于 T1、T2、T3 加起来的两倍，设计年旅客吞吐量 5000 万人次。');
}

async function slideGallery() {
  const s = await newSlide('prism');
  await heading(s, '03', '荷花航站楼', '实景相册：抬头就能看到“荷花”', { face: 'wow' });
  const shots = [
    ['hgh_lotus', '“荷花谷”采光天窗', 2.3, 4.05, -4, 0.5, 0.35, art.P.pink],
    ['hgh_checkin', '值机岛上方的天窗', 5.22, 4.3, 3, 0.4, 0.5, art.P.yellow],
    ['hgh_concourse', '国际出发层的指廊', 8.12, 4.02, -2.5, 0.5, 0.45, art.P.sky],
    ['hgh_china', '出发大厅的主题景观', 11.03, 4.3, 4, 0.45, 0.6, art.P.leaf],
  ];
  for (const [i, [key, caption, cx, cy, angle, fx, fy, tape]] of shots.entries()) {
    const n = `Shot ${i + 1}`;
    const parts = await polaroid(s, n, key, caption, cx, cy, 2.85, angle, { fx, fy, ratio: 0.8 });
    const a = (angle * Math.PI) / 180;
    const h = 0.13 + (2.85 - 0.26) * 0.8 + 0.46;
    const tx = cx + (h / 2) * Math.sin(a), ty = cy - (h / 2) * Math.cos(a);
    s.rect(tx - 0.48, ty - 0.13, 0.96, 0.26, { color: tape.slice(1), transparency: 22 }, { name: n + ' tape', rotate: (angle - 4 + 360) % 360 });
    const d = 600 + i * 500;
    for (const p of parts) s.anim(p, 'drop', d, 800, { offset: 0.14 });
    s.anim(n + ' tape', 'pop', d + 650, 350);
  }
  s.pill('T4 把荷花元素用在了立柱、屋顶和天窗上', (W - 4.9) / 2, 5.9, 4.9, 0.38, C.cream, { color: C.ink, fontSize: 12, bold: false, name: 'Gallery note' });
  s.anim('Gallery note', 'fade', 2900, 600);
  s.text(creditText(shots.map((x) => x[0])), MX, 6.4, W - 2 * MX, 0.2, { fontSize: 8, color: C.muted, align: 'center', name: 'Photo credit' });
  s.anim('Photo credit', 'fade', 3000, 400);
  footer(s, 8);
  s.notes('这一页是 T4 的实景照片。第一张是“荷花谷”的采光天窗；第二张是值机岛，上方也开了天窗；第三张是国际出发层的指廊；第四张是出发大厅里的主题景观。可以看到，荷花元素出现在立柱、屋顶和天窗上，抬头就能看到。');
}

async function slideGreenSmart() {
  const s = await newSlide('origami');
  await heading(s, '03', '荷花航站楼', '不只好看：绿色又聪明的航站楼');
  const cols = [
    ['leaf', '绿色航站楼', 'EEF9EF', C.teal, [
      ['三星级绿色建筑', '2021 年 1 月获三星级绿色建筑设计评价标识'],
      ['会“降温”的屋顶', '金属屋面覆盖辐射制冷膜，反射太阳辐射'],
      ['自然通风和采光', '窗户可自动控制，天窗引入自然光'],
      ['高效能源中心', '分布式能源系统，能源综合利用效率约 90%'],
    ]],
    ['chip', '智慧出行', 'E9F5FC', C.skyDark, [
      ['托运“0 高差”', '行李称重带与地面齐平，箱子不用往上抬'],
      ['智能安检', '安检通道装有“自出筐”回筐系统'],
      ['刷脸登机', '登机口人脸识别，刷脸即可通过'],
      ['智慧云数据中心', '运行一图指挥，服务一触即达'],
    ]],
  ];
  const cw = 4.1, gap = 0.26, top = 2.08, ch = 4.4;
  for (const [c, [iconName, title, tint, color, rows]] of cols.entries()) {
    const x = MX + c * (cw + gap);
    const n = `Col ${c + 1}`;
    s.card(x, top, cw, ch, { name: n + ' card', fill: tint });
    await s.icon(iconName, x + 0.24, top + 0.2, 0.62, n + ' icon');
    s.text(title, x + 1.02, top + 0.2, cw - 1.2, 0.62, { fontSize: 19, bold: true, color: C.ink, valign: 'middle', name: n + ' title' });
    const d = 600 + c * 900;
    s.anim(n + ' card', 'rise', d, 600);
    s.anim(n + ' icon', 'pop', d + 200, 450);
    s.anim(n + ' title', 'fade', d + 300, 450);
    rows.forEach(([head, text], r) => {
      const y = top + 1.02 + r * 0.82;
      const rn = `${n} row ${r + 1}`;
      s.dot(x + 0.4, y + 0.15, 0.16, color, { name: rn + ' dot' });
      s.text(head, x + 0.64, y, cw - 0.85, 0.3, { fontSize: 13.5, bold: true, color: C.ink, name: rn + ' head' });
      s.text(text, x + 0.64, y + 0.31, cw - 0.85, 0.42, { fontSize: 11, color: C.text, name: rn + ' text' });
      const rd = d + 500 + r * 220;
      s.anim(rn + ' dot', 'pop', rd, 300);
      s.anim(rn + ' head', 'slide', rd + 60, 400, { offset: -0.03 });
      s.anim(rn + ' text', 'fade', rd + 200, 400);
    });
  }
  const px = MX + 2 * (cw + gap), pw = W - MX - px;
  await s.photo('hgh_checkin2', px, top, pw, 3.0, 'Smart photo', { fx: 0.55, fy: 0.5 });
  credit(s, 'hgh_checkin2', px, top + 3.08, pw, { align: 'right' });
  s.anim('Smart photo', 'zoom', 800, 800);
  s.anim('Photo credit', 'fade', 1500, 400);
  s.card(px, top + 3.42, pw, 0.98, { name: 'Energy card', fill: C.cream });
  s.text([numRun('90', { fontSize: 30, color: C.orange, bold: true }), numRun('%', { fontSize: 18, color: C.orange, bold: true })], px + 0.2, top + 3.5, 1.25, 0.82, { valign: 'middle', align: 'center', name: 'Energy value' });
  s.text('能源中心的能源\n综合利用效率', px + 1.5, top + 3.5, pw - 1.65, 0.82, { fontSize: 12, color: C.ink, bold: true, valign: 'middle', lineSpacingMultiple: 1.15, name: 'Energy label' });
  s.anim('Energy card', 'rise', 2500, 500);
  s.anim('Energy value', 'pop', 2750, 550);
  s.anim('Energy label', 'fade', 2950, 450);
  footer(s, 9);
  s.notes('T4 不只是好看。在绿色方面，它在 2021 年 1 月拿到了三星级绿色建筑设计评价标识；金属屋面上覆盖了辐射制冷膜，可以反射太阳辐射给室内降温；窗户能自动控制，配合天窗做自然通风和采光；配套能源中心的能源综合利用效率约 90%。在智慧方面，行李托运是“0 高差”设计，箱子不用往上抬；安检通道有自动回筐系统；还可以刷脸登机。');
}

/** Big rounded number with a small unit after it. */
const bigValue = (value, unit, color, size = 34) => [
  numRun(value, { fontSize: size, color, bold: true }),
  run(' ' + unit, { fontSize: Math.round(size * 0.36), color, bold: true }),
];

async function slideReport() {
  const s = await newSlide('curtains');
  await heading(s, '04', '数据看实力', '2025 成绩单：三项指标都在全国前十', { face: 'happy' });
  const stats = [
    ['people', '旅客吞吐量', '5045.9', '万人次', '全国第 7', '同比 +5.0%', C.teal],
    ['box', '货邮吞吐量', '79.27', '万吨', '全国第 7', '同比 +7.9%', C.orange],
    ['plane', '飞机起降', '32.94', '万架次', '全国第 10', '同比 +2.8%', C.skyDark],
  ];
  const gap = 0.3, cw = (W - 2 * MX - 2 * gap) / 3, top = 2.05, ch = 2.35;
  for (const [i, [iconName, label, value, unit, rank, yoy, color]] of stats.entries()) {
    const x = MX + i * (cw + gap);
    const n = `Stat ${i + 1}`;
    s.card(x, top, cw, ch, { name: n + ' card' });
    s.round(x, top, cw, 0.12, color, { radius: 0.06, name: n + ' strip' });
    await s.icon(iconName, x + 0.26, top + 0.34, 0.6, n + ' icon');
    s.text(label, x + 1.02, top + 0.34, cw - 1.2, 0.6, { fontSize: 15, bold: true, color: C.ink, valign: 'middle', name: n + ' label' });
    s.text(bigValue(value, unit, color, 40), x + 0.26, top + 1.02, cw - 0.4, 0.74, { valign: 'middle', name: n + ' value' });
    s.pill(rank, x + 0.26, top + 1.84, 1.15, 0.3, color, { fontSize: 11, name: n + ' rank' });
    s.pill(yoy, x + 1.53, top + 1.84, 1.25, 0.3, C.mintLight, { fontSize: 11, color: C.tealDark, name: n + ' yoy' });
    const d = 600 + i * 450;
    s.anim(n + ' card', 'rise', d, 600);
    s.anim(n + ' strip', 'wipe', d + 200, 500);
    s.anim(n + ' icon', 'pop', d + 250, 450);
    s.anim(n + ' label', 'fade', d + 300, 400);
    s.anim(n + ' value', 'zoom', d + 450, 600);
    s.anim(n + ' rank', 'pop', d + 800, 400);
    s.anim(n + ' yoy', 'pop', d + 900, 400);
  }

  // A round "first time over 50 million" stamp on the passenger card.
  const sx = MX + cw - 0.62, sy = top + 0.7;
  s.dot(sx, sy, 1.02, C.white, { name: 'Stamp ring', line: { color: C.pinkDark, width: 2.5, dashType: 'sysDash' }, rotate: 345 });
  s.text([run('首破', { breakLine: true }), run('5000 万')], sx - 0.5, sy - 0.32, 1.0, 0.64, { fontSize: 12.5, bold: true, color: C.pinkDark, align: 'center', valign: 'middle', rotate: 345, name: 'Stamp text' });
  s.anim('Stamp ring', 'zoom', 2500, 450);
  s.anim('Stamp text', 'zoom', 2500, 450);
  s.anim('Stamp ring', 'teeter', 3000, 500, { deg: 6, repeat: 2 });
  s.anim('Stamp text', 'teeter', 3000, 500, { deg: 6, repeat: 2 });

  const chips = [
    ['89.1', '%', '平均客座率', '连续 24 个月居全国前十大机场首位', C.pinkDark],
    ['223', '个', '客货通航点', '航线覆盖 5 大洲 35 个国家及地区', C.teal],
    ['75', '家', '入驻航空公司', '截至 2025 年底', C.orange],
    ['530', '万+', '出入境人员（人次）', '同比增长 11.6%', C.skyDark],
  ];
  const cg = 0.24, chw = (W - 2 * MX - 3 * cg) / 4, cy = 4.66, chh = 1.5;
  for (const [i, [value, unit, label, note, color]] of chips.entries()) {
    const x = MX + i * (chw + cg);
    const n = `Chip ${i + 1}`;
    s.card(x, cy, chw, chh, { name: n + ' card', fill: 'FAFEFD' });
    s.text(bigValue(value, unit, color, 26), x + 0.22, cy + 0.14, chw - 0.4, 0.52, { valign: 'middle', name: n + ' value' });
    s.text(label, x + 0.22, cy + 0.72, chw - 0.4, 0.28, { fontSize: 12.5, bold: true, color: C.ink, name: n + ' label' });
    s.text(note, x + 0.22, cy + 1.03, chw - 0.4, 0.4, { fontSize: 10, color: C.sub, name: n + ' note' });
    const d = 2300 + i * 250;
    s.anim(n + ' card', 'rise', d, 500);
    s.anim(n + ' value', 'pop', d + 200, 450);
    s.anim(n + ' label', 'fade', d + 300, 400);
    s.anim(n + ' note', 'fade', d + 400, 400);
  }
  s.text('数据来源：民航局《2025 年全国民用运输机场生产统计公报》；潮新闻 2026 年 1 月 6 日报道', MX, 6.36, W - 2 * MX, 0.22, { fontSize: 9, color: C.muted, name: 'Source' });
  s.anim('Source', 'fade', 3500, 400);
  footer(s, 10);
  s.notes('第四站看数据。2025 年，萧山机场旅客吞吐量 5045.9 万人次，全国第 7，这是第一次突破 5000 万；货邮吞吐量 79.27 万吨，也是全国第 7；飞机起降 32.94 万架次，全国第 10。三项指标都在全国前十。另外，平均客座率 89.1%，连续 24 个月在全国前十大机场里排第一；通航点 223 个，覆盖 5 大洲 35 个国家及地区；入驻航空公司 75 家；出入境人员超过 530 万人次。');
}

async function slideGrowth() {
  const s = await newSlide('cover');
  await heading(s, '04', '数据看实力', '客流十年：从 3000 万到 5000 万');
  const years = ['2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'];
  const values = [3159.5, 3557.0, 3824.2, 4010.8, 2822.4, 2816.4, 2003.8, 4117.0, 4805.4, 5045.9];
  const pale = 'B5DDD6';
  const colors = [C.teal, C.teal, C.teal, C.teal, pale, pale, pale, C.teal, C.teal, C.pink];
  const cx = MX, cy = 2.02, cw = 8.05, ch = 4.3;
  s.card(cx, cy, cw, ch, { name: 'Chart card' });
  s.text('年旅客吞吐量（万人次）', cx + 0.3, cy + 0.2, 4, 0.28, { fontSize: 12.5, bold: true, color: C.ink, name: 'Chart title' });
  s.chart(pptx.charts.BAR, [{ name: '旅客吞吐量（万人次）', labels: years, values }], {
    x: cx + 0.15, y: cy + 0.5, w: cw - 0.3, h: ch - 0.6,
    barDir: 'col',
    chartColors: colors,
    barGapWidthPct: 45,
    catAxisLabelColor: C.sub, catAxisLabelFontFace: FONT.num, catAxisLabelFontSize: 11,
    catAxisLineShow: true, catAxisLineColor: C.line,
    valAxisHidden: true, valAxisLineShow: false, valAxisMinVal: 0, valAxisMaxVal: 5800,
    valGridLine: { style: 'none' },
    catGridLine: { style: 'none' },
    showValue: true, dataLabelColor: C.ink, dataLabelFontFace: FONT.num, dataLabelFontSize: 10,
    dataLabelFormatCode: '0.0', dataLabelPosition: 'outEnd',
    showLegend: false, showTitle: false,
    altText: '杭州萧山国际机场 2016—2025 年旅客吞吐量柱状图',
  }, 'Traffic chart');
  s.anim('Chart card', 'rise', 500, 600);
  s.anim('Chart title', 'fade', 800, 400);
  s.anim('Traffic chart', 'wipe', 1000, 1800, { dir: 'down' });

  const steps = [
    ['2016', '突破 3000 万人次', C.teal],
    ['2019', '突破 4000 万人次', C.teal],
    ['2020—2022', '受新冠疫情影响，客流回落', '8FBFB8'],
    ['2023', '恢复到 4117.0 万，超过 2019 年', C.teal],
    ['2025', '首次突破 5000 万，全国第 7', C.pinkDark],
  ];
  const sx = cx + cw + 0.3, sw = W - MX - sx, rowH = 0.78;
  s.line(sx + 0.14, cy + 0.2, sx + 0.14, cy + 0.2 + rowH * (steps.length - 1), C.mint, { width: 2.5, dash: 'solid', name: 'Step line' });
  s.anim('Step line', 'wipe', 1500, 1800, { dir: 'up' });
  for (const [i, [when, text, color]] of steps.entries()) {
    const y = cy + 0.2 + i * rowH;
    const n = `Step ${i + 1}`;
    s.dot(sx + 0.14, y, 0.26, C.white, { name: n + ' dot', line: { color, width: 3 } });
    s.text(when, sx + 0.46, y - 0.2, sw - 0.46, 0.34, { fontSize: 15, bold: true, color, fontFace: FONT.num, valign: 'middle', name: n + ' when' });
    s.text(text, sx + 0.46, y + 0.16, sw - 0.46, 0.3, { fontSize: 11.5, color: C.text, name: n + ' text' });
    const d = 1500 + i * 400;
    s.anim(n + ' dot', 'pop', d, 350);
    s.anim(n + ' when', 'slide', d + 80, 400, { offset: -0.03 });
    s.anim(n + ' text', 'fade', d + 200, 400);
  }
  s.image(await drawn.paper(), sx + sw - 1.0, cy + ch - 0.56, 0.72, 0.56, 'Paper plane', '纸飞机');
  s.anim('Paper plane', 'fly', 3300, 800, { dir: 'left' });
  s.anim('Paper plane', 'float', 4200, 2000, { dy: -0.02 });
  s.text('数据来源：民航局历年《全国民用运输机场生产统计公报》', MX, 6.4, 8, 0.22, { fontSize: 9, color: C.muted, name: 'Source' });
  s.anim('Source', 'fade', 3000, 400);
  footer(s, 11);
  s.notes('这张图是 2016 到 2025 年的年旅客吞吐量，数据来自民航局每年的统计公报。2016 年突破 3000 万，2019 年突破 4000 万。浅色的三根柱子是 2020 到 2022 年，受新冠疫情影响客流回落，2022 年只有 2003.8 万。2023 年恢复到 4117 万，超过了 2019 年；2024 年 4805.4 万；2025 年 5045.9 万，首次突破 5000 万。');
}

async function slideNetwork() {
  const s = await newSlide('flythrough');
  await heading(s, '04', '数据看实力', '航线网络：223 个通航点，连接 5 大洲', { face: 'wink' });
  const lw = 5.95, top = 2.05;
  const tiles = [
    ['globe', '223', '个', '客货通航点', C.teal],
    ['pin', '35', '个', '国家及地区，覆盖 5 大洲', C.pinkDark],
    ['plane', '75', '家', '入驻航空公司', C.orange],
    ['people', '93.6', '万', '中转旅客（人次）', C.skyDark],
  ];
  const tg = 0.22, tw = (lw - tg) / 2, th = 1.08;
  for (const [i, [iconName, value, unit, label, color]] of tiles.entries()) {
    const x = MX + (i % 2) * (tw + tg), y = top + Math.floor(i / 2) * (th + tg);
    const n = `Tile ${i + 1}`;
    s.card(x, y, tw, th, { name: n + ' card' });
    await s.icon(iconName, x + 0.18, y + 0.23, 0.62, n + ' icon');
    s.text(bigValue(value, unit, color, 24), x + 0.95, y + 0.12, tw - 1.05, 0.5, { valign: 'middle', name: n + ' value' });
    s.text(label, x + 0.95, y + 0.64, tw - 1.05, 0.3, { fontSize: 11, color: C.text, name: n + ' label' });
    const d = 600 + i * 220;
    s.anim(n + ' card', 'rise', d, 500);
    s.anim(n + ' icon', 'pop', d + 150, 400);
    s.anim(n + ' value', 'pop', d + 250, 450);
    s.anim(n + ' label', 'fade', d + 350, 400);
  }

  const by = top + 2 * th + 2 * tg, bh = 1.3;
  s.card(MX, by, lw, bh, { name: 'Base card', fill: C.mintLight });
  s.text('基地航空公司', MX + 0.24, by + 0.16, 2.5, 0.28, { fontSize: 13, bold: true, color: C.ink, name: 'Base title' });
  s.anim('Base card', 'rise', 1600, 500);
  s.anim('Base title', 'fade', 1800, 400);
  const bases = [['中国国际航空', C.pinkDark, 1.4], ['厦门航空', C.skyDark, 1.02], ['长龙航空', C.orange, 1.02]];
  let bx = MX + 0.24;
  for (const [i, [name, color, w]] of bases.entries()) {
    s.pill(name, bx, by + 0.62, w, 0.42, color, { fontSize: 11.5, name: `Base ${i + 1}` });
    s.anim(`Base ${i + 1}`, 'pop', 1900 + i * 180, 400);
    bx += w + 0.12;
  }
  s.line(bx + 0.1, by + 0.2, bx + 0.1, by + bh - 0.2, C.mint, { width: 1.5, dash: 'solid', name: 'Base divider' });
  s.text('货运基地航司', bx + 0.32, by + 0.16, 1.6, 0.28, { fontSize: 13, bold: true, color: C.ink, name: 'Cargo base title' });
  s.pill('圆通货运航空', bx + 0.32, by + 0.62, 1.45, 0.42, C.tealDark, { fontSize: 11.5, name: 'Cargo base' });
  s.anim('Base divider', 'fade', 2400, 300);
  s.anim('Cargo base title', 'fade', 2450, 400);
  s.anim('Cargo base', 'pop', 2550, 400);
  s.round(MX, by + bh + 0.2, lw, 0.5, C.cream, { radius: 0.25, name: 'News pill' });
  s.text('2025 年：恢复河内、雅加达等客运航线，阿联酋航空进驻', MX + 0.2, by + bh + 0.2, lw - 0.4, 0.5, { fontSize: 11.5, color: C.ink, valign: 'middle', align: 'center', name: 'News text' });
  s.anim('News pill', 'wipe', 2800, 600);
  s.anim('News text', 'fade', 3100, 500);

  const px = MX + lw + 0.3, pw = W - MX - px, ph = 2.08;
  const shots = [['hgh_loong', '长龙航空 A320neo', 0.5, 0.55], ['hgh_c919', '国航 C919 降落杭州', 0.5, 0.5]];
  for (const [i, [key, caption, fx, fy]] of shots.entries()) {
    const y = top + i * (ph + 0.22);
    const n = `Fleet ${i + 1}`;
    await s.photo(key, px, y, pw, ph, n + ' photo', { fx, fy });
    s.pill(caption, px + 0.2, y + ph - 0.5, 0.5 + 0.17 * [...caption].length, 0.32, C.white, { color: C.ink, fontSize: 10.5, name: n + ' caption' });
    s.anim(n + ' photo', i === 0 ? 'fly' : 'fly', 900 + i * 500, 800, { dir: 'right' });
    s.anim(n + ' caption', 'pop', 1700 + i * 500, 400);
  }
  credit(s, ['hgh_loong', 'hgh_c919'], px, top + 2 * ph + 0.3, pw, { align: 'right' });
  s.anim('Photo credit', 'fade', 2400, 400);
  footer(s, 12);
  s.notes('航线网络方面，到 2025 年底，萧山机场的客货通航点有 223 个，航线覆盖 5 大洲 35 个国家及地区，入驻航空公司 75 家。2025 年中转旅客 93.6 万人次，其中国际中转 23.8 万人次。这里的基地航空公司有国航、厦门航空和长龙航空，货运基地航司是圆通货运航空。右边两张照片，上面是长龙航空的 A320neo，下面是国航的国产大飞机 C919 降落杭州。2025 年机场恢复了河内、雅加达等客运航线，还引入了阿联酋航空。');
}

async function slideCargo() {
  const s = await newSlide('conveyor');
  await heading(s, '04', '数据看实力', '航空货运：把“浙江制造”送往全球');
  const top = 2.05, pw = 4.9;
  await s.photo('hgh_sf', MX, top, pw, 2.3, 'Cargo photo a', { fx: 0.5, fy: 0.55 });
  s.pill('停放在机场的顺丰航空货机', MX + 0.2, top + 1.8, 2.75, 0.32, C.white, { color: C.ink, fontSize: 10.5, name: 'Cargo caption a' });
  await s.photo('hgh_sfhub', MX, top + 2.5, pw, 1.72, 'Cargo photo b', { fx: 0.5, fy: 0.5 });
  s.pill('机场内的物流枢纽与货机坪', MX + 0.2, top + 3.72, 2.75, 0.32, C.white, { color: C.ink, fontSize: 10.5, name: 'Cargo caption b' });
  credit(s, ['hgh_sf', 'hgh_sfhub'], MX, top + 4.3, pw);
  s.anim('Cargo photo a', 'fly', 600, 800, { dir: 'left' });
  s.anim('Cargo caption a', 'pop', 1400, 400);
  s.anim('Cargo photo b', 'fly', 900, 800, { dir: 'left' });
  s.anim('Cargo caption b', 'pop', 1700, 400);
  s.anim('Photo credit', 'fade', 1900, 400);

  const rx = MX + pw + 0.3, rw = W - MX - rx;
  const stats = [
    ['79.27', '万吨', '2025 年货邮吞吐量', '全国第 7，同比 +7.9%', C.orange],
    ['25.7', '万吨', '国际及地区货量', '同比 +10.4%', C.teal],
    ['28', '条', '国际货运航线', '每周约 120 架次货机', C.skyDark],
  ];
  const sg = 0.2, sw = (rw - 2 * sg) / 3, sh = 1.5;
  for (const [i, [value, unit, label, note, color]] of stats.entries()) {
    const x = rx + i * (sw + sg);
    const n = `Cargo stat ${i + 1}`;
    s.card(x, top, sw, sh, { name: n + ' card' });
    s.text(bigValue(value, unit, color, 26), x + 0.18, top + 0.14, sw - 0.3, 0.52, { valign: 'middle', name: n + ' value' });
    s.text(label, x + 0.18, top + 0.72, sw - 0.3, 0.28, { fontSize: 12, bold: true, color: C.ink, name: n + ' label' });
    s.text(note, x + 0.18, top + 1.03, sw - 0.3, 0.36, { fontSize: 10, color: C.sub, name: n + ' note' });
    const d = 1100 + i * 280;
    s.anim(n + ' card', 'drop', d, 700, { offset: 0.08 });
    s.anim(n + ' value', 'pop', d + 400, 450);
    s.anim(n + ' label', 'fade', d + 500, 400);
    s.anim(n + ' note', 'fade', d + 600, 400);
  }

  const ly = top + sh + 0.22, lh = 1.86;
  s.card(rx, ly, rw, lh, { name: 'Moves card', fill: 'FFF8EC' });
  await s.icon('box', rx + 0.22, ly + 0.18, 0.5, 'Moves icon');
  s.text('2025 年的新动作', rx + 0.86, ly + 0.18, rw - 1.1, 0.5, { fontSize: 14.5, bold: true, color: C.ink, valign: 'middle', name: 'Moves title' });
  const moves = ['新增东京、塔什干、俄斯特拉发 3 条全货机航线', '“杭州—墨西哥城”货运航线加密到每周 4 班', '温哥华海外货站投入运营'];
  s.text(moves.map((l) => run(l, { bullet: { characterCode: '25CF', indent: 13 }, paraSpaceAfter: 4, breakLine: true })),
    rx + 0.3, ly + 0.8, rw - 0.55, lh - 0.9, { fontSize: 12, color: C.text, name: 'Moves text' });
  s.anim('Moves card', 'rise', 2100, 500);
  s.anim('Moves icon', 'pop', 2300, 400);
  s.anim('Moves title', 'fade', 2350, 400);
  moves.forEach((_, k) => s.anim('Moves text', 'wipe', 2550 + k * 300, 450, { para: k }));

  const by = ly + lh + 0.2;
  s.image(await drawn.suitcase(), rx + 0.05, by - 0.04, 0.62, 0.62, 'Cargo suitcase', '行李箱');
  bubble(s, 'Record', '小知识：货邮吞吐量的最高纪录是 2021 年的 91.4 万吨', rx + 0.86, by, rw - 0.86, 0.54, { tail: 'left', fontSize: 11.5 });
  s.anim('Cargo suitcase', 'pop', 3500, 450);
  s.anim('Cargo suitcase', 'teeter', 4000, 600, { deg: 8 });
  for (const n of ['Record bubble', 'Record tail', 'Record text']) s.anim(n, 'pop', 3700, 450);
  footer(s, 13);
  s.notes('除了客运，萧山机场的货运也很强。2025 年货邮吞吐量 79.27 万吨，全国第 7，同比增长 7.9%；其中国际及地区货量 25.7 万吨，增长 10.4%。国际货运航线有 28 条，每周大约 120 架次货机。2025 年新增了东京、塔什干、俄斯特拉发 3 条全货机航线，去墨西哥城的货运航线加密到每周 4 班，还在温哥华投运了海外货站。补充一个小知识：货邮吞吐量的历史最高纪录是 2021 年的 91.4 万吨。');
}

async function slideTransport() {
  const s = await newSlide('pan');
  await heading(s, '05', '四通八达', '四通八达：换乘不超过 5 分钟的交通中心', { face: 'happy' });
  const hubX = 3.85, hubY = 4.2;
  const nodes = [
    ['metro', '地铁 1、7、19 号线', '2020 年底接入地铁\n19 号线 2022 年 9 月开通', MX, 2.1],
    ['rail', '机场高铁（建设中）', '萧山机场站设在\n交通中心南侧，为地下站', 4.6, 2.1],
    ['bus', '机场大巴', '专线车直达\n27 个周边城市', MX, 5.2],
    ['park', '停车库', '上下 4 层\n设计停车位 4400 个', 4.6, 5.2],
  ];
  const nw = 2.9, nh = 1.2;
  for (const [i, [, , , x, y]] of nodes.entries()) {
    s.line(hubX, hubY, x + nw / 2, y + nh / 2, C.teal, { width: 1.75, dash: 'sysDash', name: `Spoke ${i + 1}` });
    s.anim(`Spoke ${i + 1}`, 'fade', 1200 + i * 250, 400);
  }
  s.dot(hubX, hubY, 2.0, C.mint, { name: 'Hub halo' });
  s.dot(hubX, hubY, 1.66, C.teal, { name: 'Hub', shadow: shadow(0.25, 10, 3) });
  s.text([run('交通中心', { fontSize: 15, bold: true, breakLine: true }), numRun('48', { fontSize: 22, bold: true }), run(' 万㎡', { fontSize: 11, bold: true })],
    hubX - 0.8, hubY - 0.5, 1.6, 1.0, { color: C.white, align: 'center', valign: 'middle', name: 'Hub text' });
  s.anim('Hub halo', 'zoom', 500, 600);
  s.anim('Hub', 'zoom', 600, 600);
  s.anim('Hub text', 'fade', 900, 400);
  s.anim('Hub halo', 'pulse', 3600, 1200, { by: 1.08, repeat: 3 });
  for (const [i, [iconName, title, text, x, y]] of nodes.entries()) {
    const n = `Mode ${i + 1}`;
    s.card(x, y, nw, nh, { name: n + ' card' });
    await s.icon(iconName, x + 0.16, y + 0.16, 0.56, n + ' icon');
    s.text(title, x + 0.86, y + 0.14, nw - 0.95, 0.3, { fontSize: 13, bold: true, color: C.ink, name: n + ' title' });
    s.text(text, x + 0.86, y + 0.5, nw - 0.95, 0.6, { fontSize: 10.5, color: C.text, lineSpacingMultiple: 1.12, name: n + ' text' });
    const d = 1300 + i * 250;
    s.anim(n + ' card', 'pop', d, 500);
    s.anim(n + ' icon', 'pop', d + 150, 400);
    s.anim(n + ' title', 'fade', d + 250, 400);
    s.anim(n + ' text', 'fade', d + 350, 400);
  }

  const px = 8.0, pw = W - MX - px, top = 2.1;
  await s.photo('hgh_metro19', px, top, pw, 2.2, 'Metro photo', { fx: 0.5, fy: 0.55 });
  s.pill('19 号线萧山国际机场站（开通当天）', px + 0.2, top + 1.7, 3.15, 0.32, C.white, { color: C.ink, fontSize: 10.5, name: 'Metro caption' });
  credit(s, 'hgh_metro19', px, top + 2.28, pw, { align: 'right' });
  s.anim('Metro photo', 'zoom', 800, 800);
  s.anim('Metro caption', 'pop', 1600, 400);
  s.anim('Photo credit', 'fade', 1700, 400);

  // Line 19 as a little route strip with travel times from the airport.
  const ry = top + 2.66, rh = 1.64;
  s.card(px, ry, pw, rh, { name: 'Route card', fill: C.skyLight });
  s.text('坐地铁 19 号线，从机场出发', px + 0.24, ry + 0.14, pw - 0.5, 0.28, { fontSize: 12.5, bold: true, color: C.ink, name: 'Route title' });
  const stops = [['机场', '', C.teal], ['杭州东站', '20 分钟', C.skyDark], ['西湖文化广场', '31 分钟', C.skyDark], ['杭州西站', '45 分钟', C.skyDark]];
  const lx0 = px + 0.5, lx1 = px + pw - 0.6, lyy = ry + 0.82;
  s.line(lx0, lyy, lx1, lyy, C.skyDark, { width: 3, dash: 'solid', name: 'Route line' });
  s.anim('Route card', 'rise', 2300, 500);
  s.anim('Route title', 'fade', 2500, 400);
  s.anim('Route line', 'wipe', 2700, 1200);
  for (const [i, [stop, time, color]] of stops.entries()) {
    const x = lx0 + ((lx1 - lx0) * i) / (stops.length - 1);
    const n = `Stop ${i + 1}`;
    s.dot(x, lyy, 0.24, C.white, { name: n + ' dot', line: { color, width: 3 } });
    s.text(stop, x - 0.7, lyy + 0.2, 1.4, 0.24, { fontSize: 10.5, bold: true, color: C.ink, align: 'center', name: n + ' name' });
    if (time) s.text(time, x - 0.7, lyy + 0.45, 1.4, 0.24, { fontSize: 10.5, color: C.tealDark, align: 'center', fontFace: FONT.cn, bold: true, name: n + ' time' });
    const d = 2700 + i * 380;
    s.anim(n + ' dot', 'pop', d, 350);
    s.anim(n + ' name', 'fade', d + 100, 350);
    if (time) s.anim(n + ' time', 'fade', d + 200, 350);
  }
  s.image(await drawn.plane('smile'), lx0 - 0.35, lyy - 0.55, 0.62, 0.443, 'Route rider', '小飞机吉祥物');
  s.anim('Route rider', 'fade', 2600, 300);
  s.anim('Route rider', 'path', 2700, 1500, { path: `M 0 0 L ${((lx1 - lx0) / W).toFixed(4)} 0 E` });
  footer(s, 14);
  s.notes('第五站是交通。T4 航站楼旁边就是 48 万平方米的交通中心，地上 2 层、地下 4 层，官方说法是各种交通方式的换乘时间都在 5 分钟之内。地铁方面，2020 年底 1 号线和 7 号线接入机场，2022 年 9 月 19 号线开通；坐 19 号线到杭州东站 20 分钟，到西湖文化广场 31 分钟，到杭州西站 45 分钟。机场大巴的专线车可以直达 27 个周边城市。停车库上下 4 层，设计车位 4400 个。另外，杭州机场高铁正在建设，萧山机场站是地下站，就设在交通中心南侧。');
}

async function slideGateway() {
  const s = await newSlide('doors');
  await heading(s, '06', '服务与未来', '城市门户：从 G20 峰会到杭州亚运会', { face: 'wow' });
  const lw = 6.75, top = 2.05;
  const marks = [
    ['2016', 'G20 杭州峰会', '完成航空运输保障', C.teal],
    ['2022.9', '三期工程投运', '亚运会重要基础配套项目', C.orange],
    ['2023.9', '杭州第 19 届亚运会', '9 月 23 日开幕', C.pinkDark],
  ];
  const mg = 0.18, mw = (lw - 2 * mg) / 3, mh = 1.18;
  for (const [i, [when, title, note, color]] of marks.entries()) {
    const x = MX + i * (mw + mg);
    const n = `Mark ${i + 1}`;
    s.card(x, top, mw, mh, { name: n + ' card' });
    s.pill(when, x + 0.16, top + 0.15, 0.34 + 0.105 * when.length, 0.28, color, { fontSize: 10.5, fontFace: FONT.num, name: n + ' when' });
    s.text(title, x + 0.18, top + 0.5, mw - 0.3, 0.3, { fontSize: 12.5, bold: true, color: C.ink, name: n + ' title' });
    s.text(note, x + 0.18, top + 0.82, mw - 0.3, 0.26, { fontSize: 10, color: C.sub, name: n + ' note' });
    const d = 600 + i * 350;
    s.anim(n + ' card', 'fly', d, 600, { dir: 'left' });
    s.anim(n + ' when', 'pop', d + 400, 350);
    s.anim(n + ' title', 'fade', d + 450, 400);
    s.anim(n + ' note', 'fade', d + 550, 400);
  }

  const ty = top + mh + 0.26;
  s.text('亚运会期间，机场累计保障', MX, ty, lw, 0.3, { fontSize: 13, bold: true, color: C.ink, name: 'Games title' });
  s.anim('Games title', 'fade', 1800, 400);
  const tiles = [
    ['plane', '4538', '架次', '涉亚运客运航班', C.teal],
    ['people', '5.09', '万人次', '涉亚运旅客', C.pinkDark],
    ['suitcase', '8.7', '万件', '涉亚运行李', C.orange],
    ['horse', '224', '匹', '参赛马匹（货运航班 44 架次）', C.skyDark],
  ];
  const tg = 0.2, tw = (lw - tg) / 2, th = 1.04, gy = ty + 0.42;
  for (const [i, [iconName, value, unit, label, color]] of tiles.entries()) {
    const x = MX + (i % 2) * (tw + tg), y = gy + Math.floor(i / 2) * (th + 0.18);
    const n = `Games ${i + 1}`;
    s.card(x, y, tw, th, { name: n + ' card' });
    await s.icon(iconName, x + 0.18, y + 0.22, 0.6, n + ' icon');
    s.text(bigValue(value, unit, color, 23), x + 0.95, y + 0.1, tw - 1.05, 0.5, { valign: 'middle', name: n + ' value' });
    s.text(label, x + 0.95, y + 0.62, tw - 1.05, 0.3, { fontSize: 10.5, color: C.text, name: n + ' label' });
    const d = 1900 + i * 230;
    s.anim(n + ' card', 'rise', d, 500);
    s.anim(n + ' icon', 'pop', d + 150, 400);
    s.anim(n + ' value', 'pop', d + 250, 450);
    s.anim(n + ' label', 'fade', d + 350, 400);
  }
  s.anim('Games 4 icon', 'teeter', 3600, 500, { deg: 10 });

  const px = MX + lw + 0.3, pw = W - MX - px;
  await s.photo('hgh_entrance', px, top, pw, 1.9, 'Gate photo', { fx: 0.5, fy: 0.5 });
  s.pill('T4 航站楼国际出发入口', px + 0.2, top + 1.42, 2.3, 0.32, C.white, { color: C.ink, fontSize: 10.5, name: 'Gate caption' });
  credit(s, 'hgh_entrance', px, top + 1.98, pw, { align: 'right' });
  s.anim('Gate photo', 'zoom', 900, 800);
  s.anim('Gate caption', 'pop', 1700, 400);
  s.anim('Photo credit', 'fade', 1800, 400);

  const vy = top + 2.32, vh = 2.1;
  s.card(px, vy, pw, vh, { name: 'Visa card', fill: C.pinkLight });
  await s.icon('passport', px + 0.2, vy + 0.18, 0.56, 'Visa icon');
  s.text([numRun('240', { fontSize: 20, color: C.pinkDark, bold: true }), run(' 小时过境免签', { fontSize: 14.5, bold: true, color: C.ink })], px + 0.9, vy + 0.18, pw - 1.0, 0.56, { valign: 'middle', name: 'Visa title' });
  const visa = ['2024 年 12 月 17 日起，停留时间延长到 240 小时', '杭州萧山机场口岸适用，停留范围为浙江全省', '2025 年入境外籍旅客 45.9 万人次，免签入境占 54.9%'];
  s.text(visa.map((l) => run(l, { bullet: { characterCode: '25CF', indent: 12 }, paraSpaceAfter: 4, breakLine: true })),
    px + 0.26, vy + 0.9, pw - 0.45, vh - 1.0, { fontSize: 10.5, color: C.text, name: 'Visa text' });
  s.anim('Visa card', 'rise', 3000, 500);
  s.anim('Visa icon', 'pop', 3200, 400);
  s.anim('Visa title', 'fade', 3250, 400);
  visa.forEach((_, k) => s.anim('Visa text', 'fade', 3450 + k * 250, 400, { para: k }));
  footer(s, 15);
  s.notes('最后一站讲服务和未来。萧山机场是杭州的城市门户：2016 年完成了 G20 杭州峰会的航空运输保障；2022 年 9 月投运的三期工程，本身就是杭州亚运会的重要基础配套项目。亚运会期间，机场累计保障涉亚运客运航班 4538 架次、旅客 5.09 万人次、行李 8.7 万件，还用 44 架次货运航班运来了 224 匹参赛马匹。口岸政策方面，从 2024 年 12 月 17 日起，过境免签停留时间延长到 240 小时，杭州萧山机场口岸适用，停留范围是浙江全省。2025 年从这里入境的外籍旅客有 45.9 万人次，其中免签入境占 54.9%。');
}

async function slideService() {
  const s = await newSlide('glitter');
  await heading(s, '06', '服务与未来', '暖心服务：小细节里的温度', { face: 'wink' });
  const cards = [
    ['paw', '宠物候机厅', '2025 年上线，是华东地区首个宠物候机厅', C.pinkLight, ['华东首个'], C.pinkDark],
    ['clock', '“即刻杭行”', '南航杭州—广州快线、没有托运行李的旅客，最晚起飞前 15 分钟到登机口即可', C.cream, ['杭州—广州快线', '提前 15 分钟'], C.orange],
    ['face', '无陪儿童智慧出行', '创新无陪儿童出行服务，还推出“码上查”“轻松游”行李服务', C.skyLight, ['码上查', '轻松游'], C.skyDark],
    ['passport', '国际服务中心', '集成支付、出入境、通信、文旅等功能，支持外卡支付和多币种兑换', C.mintLight, ['外卡支付', '多币种兑换'], C.teal],
  ];
  const lw = 7.5, top = 2.08, g = 0.24, cw = (lw - g) / 2, ch = 1.95;
  for (const [i, [iconName, title, text, tint, tags, tagColor]] of cards.entries()) {
    const x = MX + (i % 2) * (cw + g), y = top + Math.floor(i / 2) * (ch + g);
    const n = `Care ${i + 1}`;
    s.card(x, y, cw, ch, { name: n + ' card', fill: tint });
    await s.icon(iconName, x + 0.24, y + 0.24, 0.72, n + ' icon');
    s.text(title, x + 1.12, y + 0.24, cw - 1.3, 0.72, { fontSize: 16, bold: true, color: C.ink, valign: 'middle', name: n + ' title' });
    s.text(text, x + 0.26, y + 1.04, cw - 0.52, 0.5, { fontSize: 11, color: C.text, lineSpacingMultiple: 1.12, name: n + ' text' });
    let tx = x + 0.26;
    tags.forEach((tag, k) => {
      const w = 0.4 + 0.165 * [...tag].length;
      s.pill(tag, tx, y + ch - 0.48, w, 0.3, C.white, { color: tagColor, fontSize: 10, name: `${n} tag ${k + 1}` });
      tx += w + 0.12;
    });
    const d = 600 + i * 380;
    s.anim(n + ' card', 'pop', d, 550);
    s.anim(n + ' icon', 'pop', d + 250, 450);
    s.anim(n + ' title', 'fade', d + 350, 400);
    s.anim(n + ' text', 'fade', d + 450, 450);
    tags.forEach((_, k) => s.anim(`${n} tag ${k + 1}`, 'pop', d + 650 + k * 150, 350));
    s.anim(n + ' icon', 'teeter', 3200 + i * 300, 500, { deg: 9, repeat: 2 });
  }

  const cx = MX + lw + (W - MX - (MX + lw)) / 2 + 0.12;
  const parts = await polaroid(s, 'Snack', 'hgh_zwg', '国际出发区的知味观', cx, 3.72, 3.5, 4, { fx: 0.5, fy: 0.5, ratio: 0.78 });
  for (const p of parts) s.anim(p, 'drop', 1500, 800, { offset: 0.14 });
  bubble(s, 'Snack', '候机时还能尝到杭州老字号', cx - 1.55, 5.62, 3.1, 0.5, { tail: 'none', fontSize: 11.5 });
  for (const n of ['Snack bubble', 'Snack text']) s.anim(n, 'pop', 2500, 450);
  s.text(creditText('hgh_zwg'), cx - 1.75, 6.2, 3.5, 0.2, { fontSize: 8, color: C.muted, align: 'center', name: 'Photo credit' });
  s.anim('Photo credit', 'fade', 2600, 400);
  const hearts = [[cx - 1.95, 2.2, 0.34], [cx + 1.6, 2.0, 0.26], [cx + 1.72, 5.3, 0.3]];
  for (const [i, [x, y, size]] of hearts.entries()) {
    s.image(await drawn.heart(), x, y, size, size, `Heart ${i + 1}`, '爱心');
    s.anim(`Heart ${i + 1}`, 'pop', 2700 + i * 200, 400);
    s.anim(`Heart ${i + 1}`, 'float', 3200 + i * 200, 1500 + i * 200, { dy: -0.02 });
  }
  footer(s, 16);
  s.notes('萧山机场还有不少让人觉得暖心的小细节。第一，2025 年上线了华东地区首个宠物候机厅。第二，推出了“即刻杭行”服务，坐南航杭州到广州快线、没有托运行李的旅客，最晚在起飞前 15 分钟到登机口就可以。第三，创新了无陪儿童智慧出行，还有“码上查”“轻松游”这样的行李服务。第四，国际服务中心把支付、出入境、通信、文旅这些功能放在了一起，支持外卡支付和多币种兑换。右边这张照片是国际出发区的知味观，候机的时候还能尝到杭州老字号。');
}

async function slideFuture() {
  const s = await newSlide('cube');
  await heading(s, '06', '服务与未来', '未来展望：下一站，更大的枢纽', { face: 'happy' });
  const plans = [
    ['runway', '四期扩建：第三跑道', '规划新建第三跑道及滑行道系统等设施；2024 年 7 月，预可行性研究报告通过民航局组织的评估', C.teal],
    ['rail', '高铁开进机场', '杭州机场高铁正在建设，全长约 85 公里，设计时速 350 公里，萧山机场站设在交通中心南侧', C.pinkDark],
    ['chart', '总体规划（2020 年版）', '近期按年旅客吞吐量 9000 万人次、货邮吞吐量 180 万吨规划，共 4 条跑道、4 座航站楼', C.orange],
  ];
  const lw = 6.5, top = 2.08, ph = 1.32, pg = 0.2;
  s.line(MX + 0.5, top + 0.4, MX + 0.5, top + 2 * (ph + pg) + 0.6, C.mint, { width: 2.5, dash: 'sysDash', name: 'Plan route' });
  s.anim('Plan route', 'wipe', 600, 1500, { dir: 'up' });
  for (const [i, [iconName, title, text, color]] of plans.entries()) {
    const y = top + i * (ph + pg);
    const n = `Plan ${i + 1}`;
    s.card(MX + 1.1, y, lw - 1.1, ph, { name: n + ' card' });
    await s.icon(iconName, MX + 0.12, y + (ph - 0.76) / 2, 0.76, n + ' icon');
    s.text(title, MX + 1.34, y + 0.16, lw - 1.6, 0.32, { fontSize: 14.5, bold: true, color, name: n + ' title' });
    s.text(text, MX + 1.34, y + 0.54, lw - 1.6, 0.7, { fontSize: 11, color: C.text, lineSpacingMultiple: 1.15, name: n + ' text' });
    const d = 700 + i * 500;
    s.anim(n + ' icon', 'pop', d, 450);
    s.anim(n + ' card', 'fly', d + 150, 600, { dir: 'left' });
    s.anim(n + ' title', 'fade', d + 500, 400);
    s.anim(n + ' text', 'fade', d + 600, 400);
  }

  const rx = MX + lw + 0.3, rw = W - MX - rx;
  s.card(rx, top, rw, 2.72, { name: 'Target card', fill: 'FAFEFD' });
  s.text('年旅客吞吐量：现在与未来', rx + 0.26, top + 0.18, rw - 0.5, 0.3, { fontSize: 13, bold: true, color: C.ink, name: 'Target title' });
  s.anim('Target card', 'rise', 1200, 500);
  s.anim('Target title', 'fade', 1400, 400);
  const bars = [
    ['2025 年实际', 5045.9, '5045.9 万', C.teal],
    ['2035 年预测', 7500, '7500 万', C.pink],
    ['总体规划近期目标', 9000, '9000 万', C.orange],
  ];
  const maxW = rw - 0.52;
  for (const [i, [label, value, shown, color]] of bars.entries()) {
    const y = top + 0.62 + i * 0.68;
    const n = `Bar ${i + 1}`;
    const bw = (maxW * value) / 9000;
    s.text(label, rx + 0.26, y, 2.4, 0.24, { fontSize: 10.5, color: C.sub, name: n + ' label' });
    s.round(rx + 0.26, y + 0.27, maxW, 0.28, 'EEF6F4', { radius: 0.14, name: n + ' track' });
    s.round(rx + 0.26, y + 0.27, bw, 0.28, color, { radius: 0.14, name: n + ' fill' });
    s.text(shown + '人次', rx + 0.26, y + 0.27, bw - 0.14, 0.28, { fontSize: 10.5, bold: true, color: C.white, align: 'right', valign: 'middle', name: n + ' value' });
    const d = 1500 + i * 450;
    s.anim(n + ' label', 'fade', d, 350);
    s.anim(n + ' track', 'fade', d, 350);
    s.anim(n + ' fill', 'wipe', d + 150, 900);
    s.anim(n + ' value', 'fade', d + 900, 350);
  }
  await s.photo('hgh_apron', rx, top + 2.94, rw, 1.3, 'Future photo', { fx: 0.5, fy: 0.62 });
  s.pill('航站楼前停靠的客机', rx + 0.2, top + 3.78, 2.0, 0.32, C.white, { color: C.ink, fontSize: 10.5, name: 'Future caption' });
  credit(s, 'hgh_apron', rx, top + 4.32, rw, { align: 'right' });
  s.anim('Future photo', 'zoom', 2800, 700);
  s.anim('Future caption', 'pop', 3400, 400);
  s.anim('Photo credit', 'fade', 3500, 400);
  footer(s, 17);
  s.notes('再看未来。第一，四期扩建规划新建第三跑道和配套的滑行道系统，2024 年 7 月，预可行性研究报告已经通过了民航局组织的评估。第二，杭州机场高铁正在建设，全长约 85 公里，设计时速 350 公里，萧山机场站就设在交通中心南侧，以后可以在机场直接换乘高铁。第三，按照 2020 年版的总体规划，机场近期按年旅客 9000 万人次、货邮 180 万吨来规划，一共 4 条跑道、4 座航站楼。右边这组条形图：2025 年实际是 5045.9 万人次，预可研预测 2035 年是 7500 万人次，总体规划的近期目标是 9000 万人次。');
}

async function slideSummary() {
  const s = await newSlide('reveal');
  await heading(s, '', '总结', '总结：三个关键词，记住萧山机场', { mascot: false });
  const keys = [
    ['globe', '门户', C.teal, C.mintLight, ['浙江省第一空中门户', '2025 年旅客吞吐量全国第 7', '华东第二大航空枢纽']],
    ['flower', '荷花', C.pinkDark, C.pinkLight, ['T4 航站楼取意西湖荷花', '荷花形立柱撑起 72 万平方米', '绿色节能，也更智慧']],
    ['chart', '成长', C.orange, C.cream, ['通航 25 年，经历三次扩建', '2007 年破千万，2025 年破 5000 万', '高铁在建，第三跑道在规划']],
  ];
  const g = 0.3, cw = (W - 2 * MX - 2 * g) / 3, top = 2.1, ch = 3.25;
  for (const [i, [iconName, word, color, tint, lines]] of keys.entries()) {
    const x = MX + i * (cw + g);
    const n = `Key ${i + 1}`;
    s.card(x, top, cw, ch, { name: n + ' card', fill: tint });
    await s.icon(iconName, x + 0.3, top + 0.3, 0.86, n + ' icon');
    s.text(word, x + 1.36, top + 0.26, cw - 1.5, 0.94, { fontSize: 38, bold: true, color, valign: 'middle', name: n + ' word' });
    s.round(x + 0.3, top + 1.38, cw - 0.6, 0.03, color, { radius: 0.015, name: n + ' rule' });
    s.text(lines.map((l) => run(l, { bullet: { characterCode: '25CF', indent: 13 }, paraSpaceAfter: 7, breakLine: true })),
      x + 0.32, top + 1.62, cw - 0.6, ch - 1.8, { fontSize: 12.5, color: C.text, name: n + ' text' });
    const d = 600 + i * 600;
    s.anim(n + ' card', 'drop', d, 800, { offset: 0.1 });
    s.anim(n + ' icon', 'pop', d + 500, 450);
    s.anim(n + ' word', 'zoom', d + 600, 550);
    s.anim(n + ' rule', 'wipe', d + 800, 450);
    lines.forEach((_, k) => s.anim(n + ' text', 'fade', d + 950 + k * 180, 400, { para: k }));
  }
  const mw = 1.5, mh = mw * (300 / 420);
  s.image(await drawn.plane('happy'), MX + 0.15, 5.52, mw, mh, 'Summary mascot', '小飞机吉祥物');
  bubble(s, 'Line', '一句话总结：一座从“荷花”里起飞、还在继续长大的机场', MX + 1.95, 5.72, 7.0, 0.62, { tail: 'left', fontSize: 14 });
  s.anim('Summary mascot', 'fly', 3000, 800, { dir: 'left' });
  s.anim('Summary mascot', 'float', 3900, 1800, { dy: -0.018 });
  for (const n of ['Line bubble', 'Line tail', 'Line text']) s.anim(n, 'pop', 3600, 500);
  s.image(await drawn.lotus(), W - MX - 1.7, 5.4, 1.5, 1.172, 'Summary lotus', '荷花');
  s.anim('Summary lotus', 'pop', 3900, 600);
  footer(s, 18);
  s.notes('最后用三个关键词总结。第一个是“门户”：它是浙江省第一空中门户，2025 年旅客吞吐量 5045.9 万人次，全国第 7，也是华东第二大航空枢纽。第二个是“荷花”：T4 航站楼的设计取意西湖荷花，约 72 万平方米，既绿色节能，又很智慧。第三个是“成长”：通航 25 年，经历了三次扩建，2007 年旅客破千万，2025 年破 5000 万；现在高铁在建，第三跑道也在规划中。一句话：这是一座从“荷花”里起飞、还在继续长大的机场。');
}

const REFERENCES = [
  '杭州萧山国际机场官网：机场概况、保障能力（hzairport.com）',
  '中国民用航空局：历年《全国民用运输机场生产统计公报》（2017—2025 年，caac.gov.cn）',
  '杭州萧山国际机场官网新闻：三期项目正式投入运营（2022-09-22）',
  '潮新闻 / 浙江在线：杭州机场 2025 年运行情况报道（2026-01-06）',
  '国家移民管理局：过境免签政策新闻发布会（2024-12-17）',
  '中国民航网：杭州萧山国际机场总体规划（2020 年版）报道（2020-11-13）',
  '浙江省机场集团：四期扩建工程预可行性研究评估报道（2024-07-15）',
  '杭州发布、科技日报：杭州机场高铁建设进展报道（2026 年）',
  '联合资信：杭州萧山国际机场有限公司主体长期信用评级报告（2026-06-08）',
];

async function slideSources() {
  const s = await newSlide('split');
  await heading(s, '', '附录', '参考资料与图片来源', { mascot: false });
  const top = 2.05, lw = 6.9, hh = 4.3;
  s.card(MX, top, lw, hh, { name: 'Refs card' });
  s.pill('参考资料', MX + 0.26, top + 0.2, 1.15, 0.32, C.teal, { fontSize: 11.5, name: 'Refs tag' });
  s.text(REFERENCES.map((l, i) => run(`${i + 1}. ${l}`, { paraSpaceAfter: 4.5, breakLine: true })),
    MX + 0.28, top + 0.68, lw - 0.55, hh - 0.8, { fontSize: 10.5, color: C.text, name: 'Refs text' });
  s.anim('Refs card', 'rise', 500, 600);
  s.anim('Refs tag', 'pop', 800, 400);
  s.anim('Refs text', 'fade', 1000, 600);

  const rx = MX + lw + 0.3, rw = W - MX - rx;
  s.card(rx, top, rw, hh, { name: 'Credits card', fill: 'FAFEFD' });
  s.pill('图片来源', rx + 0.26, top + 0.2, 1.15, 0.32, C.pinkDark, { fontSize: 11.5, name: 'Credits tag' });
  const byArtist = new Map();
  for (const key of usedPhotos) {
    const p = photoInfo[key];
    const id = `${p.artist}（${p.license}）`;
    byArtist.set(id, (byArtist.get(id) || 0) + 1);
  }
  const lines = [
    run(`照片共 ${usedPhotos.length} 张，全部来自维基共享资源（Wikimedia Commons），按作者署名：`, { paraSpaceAfter: 5, breakLine: true }),
    ...[...byArtist].map(([id, count]) => run(`${id}  ${count} 张`, { bullet: { characterCode: '25CF', indent: 12 }, paraSpaceAfter: 3, breakLine: true })),
    run('照片按版面需要做了裁剪，每页照片旁都标了作者和许可协议。', { paraSpaceAfter: 5, breakLine: true }),
  ];
  s.text(lines, rx + 0.28, top + 0.68, rw - 0.55, 2.5, { fontSize: 10.5, color: C.text, name: 'Credits text' });
  s.round(rx + 0.26, top + hh - 1.04, rw - 0.52, 0.8, C.cream, { radius: 0.16, name: 'Art note box' });
  s.image(await drawn.plane('wink'), rx + 0.36, top + hh - 0.95, 0.86, 0.614, 'Art note mascot', '小飞机吉祥物');
  s.text('小飞机、云朵、荷花、桂花、图标和背景插画都是自己画的', rx + 1.32, top + hh - 1.04, rw - 1.72, 0.8, { fontSize: 10.5, color: C.ink, valign: 'middle', name: 'Art note text' });
  s.anim('Credits card', 'rise', 800, 600);
  s.anim('Credits tag', 'pop', 1100, 400);
  s.anim('Credits text', 'fade', 1300, 600);
  s.anim('Art note box', 'rise', 1900, 500);
  s.anim('Art note mascot', 'pop', 2100, 450);
  s.anim('Art note text', 'fade', 2200, 450);
  footer(s, 19);
  s.notes('这一页是参考资料和图片来源。数据主要来自机场官网和民航局的统计公报，照片全部来自维基共享资源，都按要求标注了作者和许可协议。小飞机、云朵、荷花这些插画是我们自己画的。');
}

async function slideThanks() {
  const s = await newSlide('vortex', 'end');
  s.image(await drawn.sun(), 0.9, 0.5, 1.0, 1.0, 'Sun', '太阳');
  s.image(await drawn.cloud(true), 10.4, 0.7, 1.9, 1.07, 'Cloud a', '云朵');
  s.image(await drawn.cloud(), 2.5, 1.5, 1.3, 0.73, 'Cloud b', '云朵');
  s.image(await drawn.cloud(), 9.3, 3.9, 1.1, 0.62, 'Cloud c', '云朵');
  s.anim('Sun', 'spin', 0, 14000, { turns: 1 });
  s.anim('Cloud a', 'float', 0, 5200, { dx: -0.03 });
  s.anim('Cloud b', 'float', 0, 6400, { dx: 0.035 });
  s.anim('Cloud c', 'float', 0, 5800, { dx: -0.025 });

  const mw = 3.0, mh = mw * (300 / 420);
  s.image(await drawn.plane('happy'), (W - mw) / 2, 0.55, mw, mh, 'Mascot', '小飞机吉祥物');
  s.anim('Mascot', 'fly', 200, 1000, { dir: 'left' });
  s.anim('Mascot', 'float', 1300, 1900, { dy: -0.022 });

  s.text('感谢聆听', 0, 2.72, W, 1.2, { fontSize: 66, bold: true, color: C.ink, align: 'center', valign: 'middle', charSpacing: 8, name: 'Thanks title' });
  s.text('THANKS FOR LISTENING', 0, 3.94, W, 0.4, { fontSize: 15, color: C.pinkDark, align: 'center', fontFace: FONT.num, charSpacing: 6, name: 'Thanks english' });
  s.text('欢迎老师和同学们批评指正', 0, 4.42, W, 0.4, { fontSize: 16, color: C.text, align: 'center', name: 'Thanks note' });
  s.anim('Thanks title', 'zoom', 700, 800);
  s.anim('Thanks english', 'wipe', 1300, 700);
  s.anim('Thanks note', 'fade', 1700, 600);

  const bw = 4.6, bx = (W - bw) / 2, by = 5.08, bh = 0.74;
  s.card(bx, by, bw, bh, { name: 'Speaker card', radius: 0.37 });
  s.text('汇报人：', bx + 0.5, by, 1.1, bh, { fontSize: 15, bold: true, color: C.ink, valign: 'middle', name: 'Speaker label' });
  s.line(bx + 1.55, by + 0.52, bx + bw - 0.5, by + 0.52, C.muted, { width: 1, dash: 'solid', name: 'Speaker blank' });
  s.anim('Speaker card', 'rise', 2000, 600);
  s.anim('Speaker label', 'fade', 2300, 400);
  s.anim('Speaker blank', 'wipe', 2400, 500);

  const hearts = [[3.3, 3.0, 0.4], [9.6, 2.7, 0.34], [4.2, 4.6, 0.26], [8.9, 5.0, 0.3]];
  for (const [i, [x, y, size]] of hearts.entries()) {
    s.image(await drawn.heart(), x, y, size, size, `Heart ${i + 1}`, '爱心');
    s.anim(`Heart ${i + 1}`, 'pop', 2200 + i * 200, 400);
    s.anim(`Heart ${i + 1}`, 'float', 2700 + i * 200, 1500 + i * 250, { dy: -0.025 });
  }
  const sparkles = [[2.2, 4.0, 0.3, art.P.yellow], [11.0, 3.2, 0.34, art.P.pink], [4.55, 2.15, 0.24, art.P.yellow]];
  for (const [i, [x, y, size, color]] of sparkles.entries()) {
    s.image(await drawn.sparkle(color), x, y, size, size, `Sparkle ${i + 1}`, '星光');
    s.anim(`Sparkle ${i + 1}`, 'pop', 2400 + i * 250, 400);
    s.anim(`Sparkle ${i + 1}`, 'pulse', 3000 + i * 250, 900, { by: 1.3, repeat: 6 });
  }
  s.notes('我的汇报到这里就结束了，谢谢大家！欢迎老师和同学们批评指正。');
}

// ---------- build -----------------------------------------------------------------------

const SLIDES = [slideCover, slideContents, slideProfile, slideRole, slideTimeline, slidePhases, slideDesign, slideGallery, slideGreenSmart, slideReport, slideGrowth, slideNetwork, slideCargo, slideTransport,
  slideGateway, slideService, slideFuture, slideSummary, slideSources, slideThanks,
];

(async () => {
  fs.mkdirSync(IMG, { recursive: true });
  fs.mkdirSync(path.dirname(OUT_FILE), { recursive: true });
  for (const build of SLIDES) await build();
  if (specs.length !== TOTAL) throw new Error(`expected ${TOTAL} slides, built ${specs.length}`);

  const raw = await pptx.write({ outputType: 'nodebuffer' });
  const zip = await JSZip.loadAsync(raw);
  let effects = 0;
  for (let i = 0; i < specs.length; i++) {
    const name = `ppt/slides/slide${i + 1}.xml`;
    const xml = await zip.file(name).async('string');
    zip.file(name, decorateSlide(xml, specs[i]));
    effects += specs[i].effects.length;
  }
  const out = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 9 } });
  fs.writeFileSync(OUT_FILE, out);
  console.log(`${OUT_FILE}\n${specs.length} slides, ${effects} animation effects, ${(out.length / 1048576).toFixed(2)} MB`);
  console.log('transitions: ' + specs.map((sp) => sp.transition).join(', '));
  console.log('photos used: ' + usedPhotos.join(', '));
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
