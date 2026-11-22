// Hand-drawn artwork for the Hangzhou Xiaoshan deck: the little plane mascot, clouds, lotus,
// osmanthus, the West Lake strip, slide backgrounds and the round icons. Everything is
// written here as SVG and rendered to PNG with sharp, so no outside clip-art is needed.

const path = require('path');
const sharp = require('sharp');

const P = {
  ink: '#2E4F54',
  teal: '#1FA79B',
  tealDark: '#11736D',
  mint: '#BFEBE1',
  mintLight: '#E3F6F1',
  pink: '#FF8FAB',
  pinkLight: '#FFD3DE',
  pinkDark: '#EA6189',
  yellow: '#FFC94D',
  orange: '#FF9F45',
  sky: '#8FD8F2',
  skyLight: '#DDF3FB',
  leaf: '#7CCB8E',
  leafDark: '#3FA36B',
  cream: '#FFF8EE',
  white: '#FFFFFF',
};

const svg = (w, h, body, viewBox = `0 0 ${w} ${h}`) =>
  `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="${viewBox}">${body}</svg>`;

// ---------- mascot ----------------------------------------------------------------------

const FACES = {
  smile:
    `<ellipse cx="304" cy="150" rx="9" ry="12" fill="${P.ink}"/><circle cx="307.5" cy="145" r="3.6" fill="#fff"/>` +
    `<ellipse cx="344" cy="152" rx="9" ry="12" fill="${P.ink}"/><circle cx="347.5" cy="147" r="3.6" fill="#fff"/>` +
    `<path d="M315 170 Q324 181 334 171" fill="none" stroke="${P.ink}" stroke-width="5" stroke-linecap="round"/>`,
  happy:
    `<path d="M294 153 Q304 139 314 153" fill="none" stroke="${P.ink}" stroke-width="6" stroke-linecap="round"/>` +
    `<path d="M334 155 Q344 141 354 155" fill="none" stroke="${P.ink}" stroke-width="6" stroke-linecap="round"/>` +
    `<path d="M313 167 Q324 190 336 168 Z" fill="${P.pinkDark}" stroke="${P.ink}" stroke-width="4.5" stroke-linejoin="round"/>`,
  wink:
    `<ellipse cx="304" cy="150" rx="9" ry="12" fill="${P.ink}"/><circle cx="307.5" cy="145" r="3.6" fill="#fff"/>` +
    `<path d="M334 154 Q344 144 355 152" fill="none" stroke="${P.ink}" stroke-width="6" stroke-linecap="round"/>` +
    `<path d="M315 170 Q324 181 334 171" fill="none" stroke="${P.ink}" stroke-width="5" stroke-linecap="round"/>`,
  wow:
    `<ellipse cx="304" cy="150" rx="9" ry="12" fill="${P.ink}"/><circle cx="307.5" cy="145" r="3.6" fill="#fff"/>` +
    `<ellipse cx="344" cy="152" rx="9" ry="12" fill="${P.ink}"/><circle cx="347.5" cy="147" r="3.6" fill="#fff"/>` +
    `<ellipse cx="325" cy="175" rx="6.5" ry="8" fill="${P.pinkDark}" stroke="${P.ink}" stroke-width="4"/>`,
};

/** The mascot: a chubby little airliner with a lotus-pink tail. Faces right. */
function planeBody(face = 'smile') {
  const o = `stroke="${P.ink}" stroke-width="7" stroke-linejoin="round" stroke-linecap="round"`;
  const body = 'M62 160 C62 118 120 100 210 100 C306 100 372 124 384 160 C390 180 372 198 338 206 C296 216 240 220 190 220 C112 220 62 204 62 160 Z';
  return (
    `<defs><clipPath id="b"><path d="${body}"/></clipPath></defs>` +
    `<path d="M196 106 L170 62 Q165 52 177 52 L202 52 Q214 52 221 61 L254 104 Z" fill="${P.tealDark}" ${o}/>` +
    `<path d="M92 132 L54 54 Q49 42 62 42 L94 42 Q108 42 116 54 L160 118 Z" fill="${P.pink}" ${o}/>` +
    `<path d="M92 76 q8 -14 0 -26 q-8 12 0 26 M92 76 q-16 -4 -20 -16 q14 0 20 16 M92 76 q16 -4 20 -16 q-14 0 -20 16" fill="#fff" opacity="0.92"/>` +
    `<path d="M76 152 L24 140 Q10 138 15 152 L24 172 Q30 184 44 180 L108 166 Z" fill="${P.teal}" ${o}/>` +
    `<path d="${body}" fill="#fff"/>` +
    `<path d="M40 190 C130 214 300 210 400 166 L400 240 L40 240 Z" fill="${P.mint}" clip-path="url(#b)"/>` +
    `<path d="${body}" fill="none" ${o}/>` +
    [150, 186, 222, 258].map((x) => `<circle cx="${x}" cy="146" r="10.5" fill="${P.sky}" stroke="${P.ink}" stroke-width="4.5"/>`).join('') +
    `<ellipse cx="290" cy="174" rx="12" ry="7" fill="${P.pink}" opacity="0.75"/><ellipse cx="358" cy="175" rx="12" ry="7" fill="${P.pink}" opacity="0.75"/>` +
    FACES[face] +
    `<path d="M188 178 L130 250 Q121 263 137 263 L172 263 Q187 263 197 253 L272 186 Z" fill="${P.teal}" ${o}/>` +
    `<ellipse cx="204" cy="236" rx="25" ry="14" fill="#fff" ${o}/><ellipse cx="222" cy="236" rx="6" ry="9" fill="${P.ink}"/>`
  );
}

function plane({ face = 'smile', flip = false } = {}) {
  const g = flip ? `<g transform="translate(420 0) scale(-1 1)">${planeBody(face)}</g>` : planeBody(face);
  return svg(420, 300, g);
}

// ---------- small pieces ----------------------------------------------------------------

function cloud({ face = false, tint = P.skyLight } = {}) {
  const shape = 'M58 150 C24 150 10 124 22 102 C30 86 48 80 62 84 C66 52 96 34 126 42 C142 20 182 18 200 46 C232 36 262 58 258 90 C284 94 294 124 278 140 C270 148 260 150 250 150 Z';
  const eyes = face
    ? `<path d="M112 108 q9 9 18 0 M166 108 q9 9 18 0" fill="none" stroke="${P.ink}" stroke-width="5" stroke-linecap="round"/>` +
      `<ellipse cx="100" cy="122" rx="10" ry="6" fill="${P.pink}" opacity="0.6"/><ellipse cx="196" cy="122" rx="10" ry="6" fill="${P.pink}" opacity="0.6"/>` +
      `<path d="M141 120 q7 7 14 0" fill="none" stroke="${P.ink}" stroke-width="4.5" stroke-linecap="round"/>`
    : '';
  return svg(300, 170,
    `<defs><clipPath id="c"><path d="${shape}"/></clipPath></defs>` +
    `<path d="${shape}" fill="#fff"/>` +
    `<path d="M0 132 C80 118 200 150 300 122 L300 170 L0 170 Z" fill="${tint}" clip-path="url(#c)"/>` + eyes);
}

const PETAL = 'M0 0 C-34 -38 -32 -96 0 -128 C32 -96 34 -38 0 0 Z';
function lotusFlower(cx, cy, s = 1) {
  const petal = (rot, fill, k = 1) =>
    `<path d="${PETAL}" transform="translate(${cx} ${cy}) rotate(${rot}) scale(${s * k})" fill="${fill}" stroke="${P.pinkDark}" stroke-width="${3 / (s * k)}" stroke-linejoin="round"/>`;
  return (
    petal(-78, P.pinkLight, 0.86) + petal(78, P.pinkLight, 0.86) +
    petal(-50, P.pink, 0.95) + petal(50, P.pink, 0.95) +
    petal(-24, P.pinkLight) + petal(24, P.pinkLight) +
    petal(0, P.pink, 1.04) +
    `<ellipse cx="${cx}" cy="${cy - 8 * s}" rx="${20 * s}" ry="${9 * s}" fill="${P.yellow}" stroke="${P.orange}" stroke-width="2.5"/>`
  );
}

function lotusLeaf(cx, cy, rx, ry, fill = P.leaf) {
  return (
    `<path d="M${cx} ${cy} L${cx + rx} ${cy - ry * 0.28} A${rx} ${ry} 0 1 1 ${cx + rx * 0.82} ${cy - ry * 0.62} Z" fill="${fill}"/>` +
    `<path d="M${cx} ${cy} L${cx - rx * 0.8} ${cy - ry * 0.3} M${cx} ${cy} L${cx - rx * 0.5} ${cy + ry * 0.7} M${cx} ${cy} L${cx + rx * 0.4} ${cy + ry * 0.8} M${cx} ${cy} L${cx - rx * 0.1} ${cy - ry * 0.9}" stroke="${P.leafDark}" stroke-width="2.5" opacity="0.55" stroke-linecap="round"/>`
  );
}

/** A lotus in bloom on its leaf. */
function lotus() {
  return svg(320, 250,
    lotusLeaf(110, 206, 104, 34) + lotusLeaf(236, 218, 74, 24, P.leafDark) +
    `<path d="M160 214 C160 196 160 190 160 176" stroke="${P.leafDark}" stroke-width="7" stroke-linecap="round"/>` +
    lotusFlower(160, 178, 1.02));
}

/** A lotus bud with a tall stalk, for corners. */
function lotusBud() {
  return svg(160, 260,
    `<path d="M80 250 C84 190 74 150 80 108" stroke="${P.leafDark}" stroke-width="7" fill="none" stroke-linecap="round"/>` +
    lotusLeaf(74, 232, 66, 22) +
    `<path d="M80 112 C52 86 56 44 80 20 C104 44 108 86 80 112 Z" fill="${P.pink}" stroke="${P.pinkDark}" stroke-width="3"/>` +
    `<path d="M80 112 C66 90 68 58 80 36 C92 58 94 90 80 112 Z" fill="${P.pinkLight}"/>`);
}

/** A sprig of osmanthus, Hangzhou's city flower. */
function osmanthus() {
  const flower = (x, y, r = 9) =>
    [0, 90, 180, 270].map((a) => `<ellipse cx="${x}" cy="${y - r * 0.8}" rx="${r * 0.55}" ry="${r * 0.8}" transform="rotate(${a + 45} ${x} ${y})" fill="${P.yellow}" stroke="${P.orange}" stroke-width="1.6"/>`).join('') +
    `<circle cx="${x}" cy="${y}" r="${r * 0.3}" fill="${P.orange}"/>`;
  const leaf = (x, y, rot, s = 1) =>
    `<g transform="translate(${x} ${y}) rotate(${rot}) scale(${s})"><path d="M0 0 C18 -22 58 -22 78 0 C58 22 18 22 0 0 Z" fill="${P.leaf}" stroke="${P.leafDark}" stroke-width="3" stroke-linejoin="round"/><path d="M6 0 L70 0" stroke="${P.leafDark}" stroke-width="2.4" opacity="0.6"/></g>`;
  return svg(300, 220,
    `<path d="M20 200 C80 170 150 120 270 40" fill="none" stroke="#9A7B5B" stroke-width="7" stroke-linecap="round"/>` +
    leaf(84, 164, -64) + leaf(120, 142, 20, 0.95) + leaf(176, 104, -70, 0.9) + leaf(214, 80, 14, 0.85) + leaf(256, 50, -40, 0.7) +
    flower(66, 170) + flower(84, 186, 8) + flower(52, 190, 7) +
    flower(146, 118) + flower(166, 132, 8) + flower(132, 136, 7) + flower(152, 100, 7) +
    flower(228, 62) + flower(246, 76, 8) + flower(216, 46, 7));
}

function sun() {
  const rays = Array.from({ length: 12 }, (_, i) => `<rect x="96" y="6" width="8" height="26" rx="4" fill="${P.yellow}" transform="rotate(${i * 30} 100 100)"/>`).join('');
  return svg(200, 200,
    rays + `<circle cx="100" cy="100" r="54" fill="${P.yellow}" stroke="${P.orange}" stroke-width="5"/>` +
    `<path d="M80 96 q7 -9 14 0 M106 96 q7 -9 14 0" fill="none" stroke="${P.ink}" stroke-width="5" stroke-linecap="round"/>` +
    `<path d="M88 112 q12 13 24 0" fill="none" stroke="${P.ink}" stroke-width="5" stroke-linecap="round"/>` +
    `<ellipse cx="72" cy="112" rx="9" ry="5.5" fill="${P.pink}" opacity="0.7"/><ellipse cx="128" cy="112" rx="9" ry="5.5" fill="${P.pink}" opacity="0.7"/>`);
}

function sparkle(color = P.yellow) {
  return svg(100, 100, `<path d="M50 4 C54 34 66 46 96 50 C66 54 54 66 50 96 C46 66 34 54 4 50 C34 46 46 34 50 4 Z" fill="${color}"/>`);
}

function heart(color = P.pink) {
  return svg(100, 100, `<path d="M50 88 C10 60 6 36 18 22 C30 8 46 14 50 28 C54 14 70 8 82 22 C94 36 90 60 50 88 Z" fill="${color}"/>`);
}

function paperPlane() {
  return svg(200, 160,
    `<path d="M10 78 L190 10 L120 150 L92 100 Z" fill="#fff" stroke="${P.ink}" stroke-width="6" stroke-linejoin="round"/>` +
    `<path d="M92 100 L190 10 L72 86 Z" fill="${P.mint}" stroke="${P.ink}" stroke-width="6" stroke-linejoin="round"/>` +
    `<path d="M92 100 L96 138 L112 116" fill="${P.teal}" stroke="${P.ink}" stroke-width="6" stroke-linejoin="round"/>`);
}

function suitcase() {
  const o = `stroke="${P.ink}" stroke-width="6" stroke-linejoin="round" stroke-linecap="round"`;
  return svg(200, 230,
    `<path d="M72 60 L72 26 Q72 14 84 14 L116 14 Q128 14 128 26 L128 60" fill="none" ${o}/>` +
    `<rect x="26" y="58" width="148" height="140" rx="22" fill="${P.pink}" ${o}/>` +
    `<path d="M26 104 H174" ${o}/><rect x="84" y="92" width="32" height="24" rx="6" fill="${P.yellow}" ${o}/>` +
    `<circle cx="62" cy="208" r="12" fill="${P.ink}"/><circle cx="138" cy="208" r="12" fill="${P.ink}"/>` +
    `<circle cx="64" cy="150" r="15" fill="#fff" opacity="0.9"/><path d="M58 150 l5 5 l9 -10" fill="none" stroke="${P.teal}" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"/>` +
    `<path d="M112 142 q14 -12 28 0 q-14 12 -28 0 Z" fill="#fff" opacity="0.9"/>`);
}

/** Stripes for the boarding pass. Decoration only: they encode nothing. */
function barcode() {
  const pattern = [3, 1, 2, 1, 3, 2, 1, 1, 3, 1, 2, 3, 1, 2, 1, 3, 1, 2, 2, 1];
  let x = 0;
  const rects = pattern.map((b) => { const r = `<rect x="${x}" y="0" width="${b}" height="10"/>`; x += b + 1; return r; }).join('');
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${x * 8}" height="120" viewBox="0 0 ${x - 1} 10" preserveAspectRatio="none"><g fill="${P.ink}">${rects}</g></svg>`;
}

// ---------- West Lake strip -------------------------------------------------------------

/** Hills, pagoda, arched bridge and lotus along the bottom edge. Transparent above the hills. */
function lakeStrip({ w = 1920, h = 250, soft = false, flowers = true } = {}) {
  const op = soft ? 0.5 : 1;
  const pagoda = (x, y, s) => {
    let out = '';
    for (let i = 0; i < 5; i++) {
      const fw = (44 - i * 6) * s;
      const fy = y - i * 20 * s;
      out += `<rect x="${x - fw * 0.36}" y="${fy - 14 * s}" width="${fw * 0.72}" height="${14 * s}" fill="#8CCFC4"/>` +
        `<path d="M${x - fw * 0.62} ${fy - 14 * s} Q${x} ${fy - 24 * s} ${x + fw * 0.62} ${fy - 14 * s} L${x + fw * 0.5} ${fy - 19 * s} Q${x} ${fy - 27 * s} ${x - fw * 0.5} ${fy - 19 * s} Z" fill="#5DB5A8"/>`;
    }
    return out + `<path d="M${x} ${y - 100 * s - 14 * s} l${-5 * s} ${-10 * s} l${5 * s} ${-18 * s} l${5 * s} ${18 * s} Z" fill="#5DB5A8"/>`;
  };
  const bridge = (x, y, bw) =>
    `<path d="M${x} ${y} Q${x + bw / 2} ${y - bw * 0.3} ${x + bw} ${y} L${x + bw} ${y + 16} L${x + bw * 0.7} ${y + 16} A${bw * 0.2} ${bw * 0.15} 0 0 0 ${x + bw * 0.3} ${y + 16} L${x} ${y + 16} Z" fill="#fff" stroke="#9FD9CF" stroke-width="3"/>`;
  const water = h - 96;
  return svg(w, h,
    `<g opacity="${op}">` +
    `<path d="M0 ${water} C${w * 0.08} ${water - 96} ${w * 0.17} ${water - 110} ${w * 0.26} ${water - 46} C${w * 0.33} ${water - 96} ${w * 0.4} ${water - 84} ${w * 0.47} ${water - 18} L${w * 0.47} ${water} Z" fill="#CDEEE7"/>` +
    `<path d="M${w * 0.5} ${water} C${w * 0.58} ${water - 74} ${w * 0.66} ${water - 92} ${w * 0.74} ${water - 40} C${w * 0.82} ${water - 110} ${w * 0.92} ${water - 100} ${w} ${water - 52} L${w} ${water} Z" fill="#CDEEE7"/>` +
    `<path d="M0 ${water} C${w * 0.05} ${water - 46} ${w * 0.12} ${water - 60} ${w * 0.2} ${water - 20} C${w * 0.26} ${water - 40} ${w * 0.3} ${water - 30} ${w * 0.36} ${water} Z" fill="#A9E0D5"/>` +
    `<path d="M${w * 0.62} ${water} C${w * 0.7} ${water - 54} ${w * 0.8} ${water - 60} ${w * 0.88} ${water - 22} C${w * 0.93} ${water - 40} ${w * 0.97} ${water - 34} ${w} ${water - 14} L${w} ${water} Z" fill="#A9E0D5"/>` +
    (soft ? '' : pagoda(w * 0.165, water - 44, 0.9)) +
    bridge(w * 0.4, water - 4, w * 0.13) +
    `<rect x="0" y="${water}" width="${w}" height="${h - water}" fill="#D9F3F5"/>` +
    [18, 44, 70].map((dy, i) => `<path d="M0 ${water + dy} ${Array.from({ length: 16 }, (_, k) => `q${w / 64} ${i % 2 ? 8 : -8} ${w / 32} 0 t${w / 32} 0`).join(' ')}" fill="none" stroke="#fff" stroke-width="4" opacity="0.9"/>`).join('') +
    (flowers
      ? lotusLeaf(w * 0.06, h - 26, 70, 20) + lotusLeaf(w * 0.105, h - 12, 46, 14, P.leafDark) +
        lotusLeaf(w * 0.9, h - 22, 76, 22) + lotusLeaf(w * 0.95, h - 40, 44, 14, P.leafDark) + lotusLeaf(w * 0.72, h - 14, 50, 15) +
        lotusFlower(w * 0.075, h - 44, 0.42) + lotusFlower(w * 0.915, h - 44, 0.46) + lotusFlower(w * 0.955, h - 54, 0.3)
      : '') +
    `</g>`);
}

// ---------- backgrounds -----------------------------------------------------------------

function softCloud(x, y, s, opacity) {
  return `<g transform="translate(${x} ${y}) scale(${s})" opacity="${opacity}"><path d="M58 150 C24 150 10 124 22 102 C30 86 48 80 62 84 C66 52 96 34 126 42 C142 20 182 18 200 46 C232 36 262 58 258 90 C284 94 294 124 278 140 C270 148 260 150 250 150 Z" fill="#fff"/></g>`;
}

function background(kind) {
  const w = 1920, h = 1080;
  if (kind === 'cover' || kind === 'end') {
    const top = kind === 'cover' ? '#CFEFFA' : '#FFE3EA';
    const mid = kind === 'cover' ? '#EAF8FB' : '#FFF1E6';
    return svg(w, h,
      `<defs><linearGradient id="g" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="${top}"/><stop offset="0.62" stop-color="${mid}"/><stop offset="1" stop-color="#FFFDF8"/></linearGradient>` +
      `<radialGradient id="s" cx="0.5" cy="0.5" r="0.5"><stop offset="0" stop-color="#FFF6CF" stop-opacity="0.95"/><stop offset="1" stop-color="#FFF6CF" stop-opacity="0"/></radialGradient></defs>` +
      `<rect width="${w}" height="${h}" fill="url(#g)"/>` +
      `<circle cx="${kind === 'cover' ? 1640 : 300}" cy="170" r="330" fill="url(#s)"/>` +
      softCloud(120, 70, 1.5, 0.75) + softCloud(840, 30, 1.0, 0.6) + softCloud(1450, 250, 1.3, 0.55) + softCloud(520, 300, 0.8, 0.5) +
      `<g transform="translate(0 ${h - 250})">${lakeStrip({ w, h: 250 }).replace(/^<svg[^>]*>|<\/svg>$/g, '')}</g>`);
  }
  // Content slides: nearly white, with a hint of sky at the top and a faint lake strip.
  return svg(w, h,
    `<defs><linearGradient id="g" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#E4F5FB"/><stop offset="0.3" stop-color="#F6FCFB"/><stop offset="1" stop-color="#FFFFFF"/></linearGradient></defs>` +
    `<rect width="${w}" height="${h}" fill="url(#g)"/>` +
    softCloud(1500, -40, 1.25, 0.8) + softCloud(1120, 20, 0.7, 0.65) + softCloud(-60, 40, 0.9, 0.6) +
    `<g transform="translate(0 ${h - 132})">${lakeStrip({ w, h: 132, soft: true, flowers: false }).replace(/^<svg[^>]*>|<\/svg>$/g, '')}</g>`);
}

// ---------- round icons -----------------------------------------------------------------

const ICONS = {
  pin: [P.pink, `<path d="M100 44 C74 44 58 64 58 86 C58 116 100 158 100 158 C100 158 142 116 142 86 C142 64 126 44 100 44 Z" fill="#fff"/><circle cx="100" cy="86" r="17" fill="${P.pink}"/>`],
  plane: [P.teal, `<path d="M40 112 L150 64 Q168 58 164 74 L150 96 L112 108 L96 148 L80 148 L88 114 L60 122 L50 138 L40 138 L44 118 Z" fill="#fff"/>`],
  terminal: [P.sky, `<path d="M40 112 Q100 44 160 112 Z" fill="#fff"/><rect x="46" y="112" width="108" height="34" rx="4" fill="#fff"/><path d="M62 122 h16 v24 h-16 Z M92 122 h16 v24 h-16 Z M122 122 h16 v24 h-16 Z" fill="${P.sky}"/>`],
  runway: [P.tealDark, `<path d="M78 156 L92 44 L108 44 L122 156 Z" fill="#fff"/><path d="M100 54 v16 M100 82 v18 M100 112 v20" stroke="${P.tealDark}" stroke-width="6" stroke-linecap="round"/>`],
  calendar: [P.orange, `<rect x="50" y="58" width="100" height="92" rx="14" fill="#fff"/><rect x="50" y="58" width="100" height="28" rx="14" fill="#FFE2C2"/><path d="M74 46 v24 M126 46 v24" stroke="#fff" stroke-width="10" stroke-linecap="round"/><path d="M72 106 h14 M94 106 h14 M116 106 h14 M72 128 h14 M94 128 h14" stroke="${P.orange}" stroke-width="9" stroke-linecap="round"/>`],
  globe: [P.teal, `<circle cx="100" cy="100" r="54" fill="none" stroke="#fff" stroke-width="9"/><ellipse cx="100" cy="100" rx="24" ry="54" fill="none" stroke="#fff" stroke-width="7"/><path d="M48 100 h104 M56 74 h88 M56 126 h88" stroke="#fff" stroke-width="7"/>`],
  box: [P.orange, `<path d="M100 44 L154 70 L154 132 L100 158 L46 132 L46 70 Z" fill="#fff"/><path d="M46 70 L100 96 L154 70 M100 96 V158" stroke="${P.orange}" stroke-width="7" fill="none" stroke-linejoin="round"/><path d="M72 57 L126 83 V104" stroke="#FFD9AE" stroke-width="9" fill="none"/>`],
  metro: [P.sky, `<rect x="58" y="44" width="84" height="98" rx="20" fill="#fff"/><rect x="70" y="58" width="60" height="38" rx="8" fill="${P.sky}"/><circle cx="78" cy="118" r="8" fill="${P.sky}"/><circle cx="122" cy="118" r="8" fill="${P.sky}"/><path d="M74 146 L62 164 M126 146 L138 164" stroke="#fff" stroke-width="9" stroke-linecap="round"/>`],
  rail: [P.pinkDark, `<path d="M40 126 C40 86 76 62 128 62 L150 62 Q162 62 162 74 L162 126 Z" fill="#fff"/><path d="M66 96 C74 80 96 76 118 76 L118 96 Z" fill="${P.pinkDark}"/><rect x="128" y="76" width="22" height="20" rx="4" fill="${P.pinkDark}"/><path d="M36 142 H166" stroke="#fff" stroke-width="9" stroke-linecap="round"/>`],
  leaf: [P.leafDark, `<path d="M52 148 C44 84 86 46 152 48 C156 112 120 156 62 148 Z" fill="#fff"/><path d="M56 150 C80 112 104 88 136 68" stroke="${P.leafDark}" stroke-width="7" fill="none" stroke-linecap="round"/>`],
  chip: [P.teal, `<rect x="62" y="62" width="76" height="76" rx="14" fill="#fff"/><rect x="82" y="82" width="36" height="36" rx="6" fill="${P.teal}"/><path d="M80 44 v16 M100 44 v16 M120 44 v16 M80 140 v16 M100 140 v16 M120 140 v16 M44 80 h16 M44 100 h16 M44 120 h16 M140 80 h16 M140 100 h16 M140 120 h16" stroke="#fff" stroke-width="8" stroke-linecap="round"/>`],
  bus: [P.orange, `<rect x="48" y="56" width="104" height="80" rx="16" fill="#fff"/><rect x="60" y="68" width="80" height="30" rx="6" fill="${P.orange}"/><circle cx="72" cy="117" r="7" fill="${P.orange}"/><circle cx="128" cy="117" r="7" fill="${P.orange}"/><rect x="62" y="134" width="18" height="18" rx="5" fill="#fff"/><rect x="120" y="134" width="18" height="18" rx="5" fill="#fff"/>`],
  park: [P.teal, `<rect x="52" y="52" width="96" height="96" rx="22" fill="#fff"/><path d="M84 128 V74 H106 Q126 74 126 92 Q126 110 106 110 H84" fill="none" stroke="${P.teal}" stroke-width="13" stroke-linecap="round" stroke-linejoin="round"/>`],
  paw: [P.pink, `<ellipse cx="100" cy="122" rx="30" ry="25" fill="#fff"/><ellipse cx="62" cy="92" rx="13" ry="17" fill="#fff"/><ellipse cx="138" cy="92" rx="13" ry="17" fill="#fff"/><ellipse cx="86" cy="66" rx="12" ry="16" fill="#fff"/><ellipse cx="114" cy="66" rx="12" ry="16" fill="#fff"/>`],
  suitcase: [P.orange, `<rect x="52" y="70" width="96" height="78" rx="14" fill="#fff"/><path d="M80 70 V56 Q80 48 88 48 H112 Q120 48 120 56 V70" fill="none" stroke="#fff" stroke-width="9"/><path d="M52 100 H148" stroke="${P.orange}" stroke-width="7"/>`],
  star: [P.yellow, `<path d="M100 40 L117 80 L160 84 L127 112 L138 154 L100 131 L62 154 L73 112 L40 84 L83 80 Z" fill="#fff"/>`],
  clock: [P.tealDark, `<circle cx="100" cy="100" r="54" fill="#fff"/><path d="M100 66 V100 L124 114" stroke="${P.tealDark}" stroke-width="9" fill="none" stroke-linecap="round" stroke-linejoin="round"/>`],
  passport: [P.pinkDark, `<rect x="60" y="42" width="80" height="116" rx="12" fill="#fff"/><circle cx="100" cy="88" r="22" fill="none" stroke="${P.pinkDark}" stroke-width="6"/><path d="M78 88 h44 M100 66 v44" stroke="${P.pinkDark}" stroke-width="5"/><path d="M78 132 h44" stroke="${P.pinkDark}" stroke-width="7" stroke-linecap="round"/>`],
  heart: [P.pink, `<path d="M100 152 C52 118 46 88 60 70 C74 54 94 60 100 78 C106 60 126 54 140 70 C154 88 148 118 100 152 Z" fill="#fff"/>`],
  people: [P.teal, `<circle cx="78" cy="78" r="20" fill="#fff"/><path d="M44 148 C44 118 60 106 78 106 C96 106 112 118 112 148 Z" fill="#fff"/><circle cx="128" cy="86" r="16" fill="#fff" opacity="0.85"/><path d="M104 148 C106 124 114 112 128 112 C144 112 156 124 156 148 Z" fill="#fff" opacity="0.85"/>`],
  chart: [P.sky, `<rect x="52" y="108" width="22" height="42" rx="5" fill="#fff"/><rect x="89" y="82" width="22" height="68" rx="5" fill="#fff"/><rect x="126" y="52" width="22" height="98" rx="5" fill="#fff"/>`],
  horse: [P.orange, `<path d="M62 150 C62 104 70 74 100 60 L96 44 L116 54 C140 58 154 80 154 100 L138 112 L122 98 C116 120 118 136 124 150 Z" fill="#fff"/><circle cx="126" cy="78" r="5" fill="${P.orange}"/>`],
  flower: [P.yellow, [0, 72, 144, 216, 288].map((a) => `<ellipse cx="100" cy="68" rx="19" ry="27" fill="#fff" transform="rotate(${a} 100 100)"/>`).join('') + `<circle cx="100" cy="100" r="15" fill="${P.orange}"/>`],
  bulb: [P.yellow, `<path d="M100 44 C70 44 54 66 54 88 C54 106 66 116 74 128 L126 128 C134 116 146 106 146 88 C146 66 130 44 100 44 Z" fill="#fff"/><rect x="78" y="134" width="44" height="12" rx="6" fill="#fff"/><rect x="86" y="150" width="28" height="10" rx="5" fill="#fff"/>`],
  bolt: [P.orange, `<path d="M112 38 L60 110 H96 L86 162 L142 86 H104 Z" fill="#fff"/>`],
  face: [P.teal, `<circle cx="100" cy="100" r="52" fill="none" stroke="#fff" stroke-width="9"/><circle cx="82" cy="90" r="7" fill="#fff"/><circle cx="118" cy="90" r="7" fill="#fff"/><path d="M78 114 Q100 136 122 114" fill="none" stroke="#fff" stroke-width="8" stroke-linecap="round"/>`],
};

function icon(name) {
  const [color, glyph] = ICONS[name];
  return svg(200, 200, `<circle cx="100" cy="100" r="96" fill="${color}"/><circle cx="100" cy="100" r="96" fill="none" stroke="#fff" stroke-width="6" opacity="0.5"/>${glyph}`);
}

// ---------- rendering -------------------------------------------------------------------

const cache = new Map();

/** Renders an SVG string to a PNG file once and returns the path. */
async function render(dir, name, svgText, width) {
  const file = path.join(dir, name + '.png');
  if (!cache.has(file)) {
    await sharp(Buffer.from(svgText), { density: 144 }).resize({ width }).png({ compressionLevel: 9 }).toFile(file);
    cache.set(file, true);
  }
  return file;
}

module.exports = {
  P, ICONS, plane, cloud, lotus, lotusBud, osmanthus, sun, sparkle, heart, paperPlane, suitcase, barcode, lakeStrip, background, icon, render,
};
