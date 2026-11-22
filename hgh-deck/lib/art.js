// Self-drawn illustrations for the deck (SVG, rasterised with sharp). No text inside the
// drawings, so the result does not depend on the fonts of the machine that builds it.

const P = {
  blue: '#4BA3E3',
  blueDeep: '#2F6FB5',
  blueSoft: '#9AD0F5',
  bluePale: '#D9EEFC',
  pink: '#FF8FAB',
  pinkDeep: '#F0668D',
  pinkSoft: '#FFC2D1',
  pinkPale: '#FFE5EC',
  green: '#63C7A0',
  greenDeep: '#3FA57F',
  greenPale: '#DDF4E8',
  yellow: '#FFC94D',
  yellowPale: '#FFF1C9',
  ink: '#2D3E55',
  white: '#FFFFFF',
};

const svg = (w, h, body, extra = '') =>
  `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" ${extra}>${body}</svg>`;

const face = (cx, cy, s = 1, mood = 'smile') => {
  const eye = (x) =>
    `<ellipse cx="${cx + x * s}" cy="${cy}" rx="${7 * s}" ry="${9.5 * s}" fill="${P.ink}"/>` +
    `<circle cx="${cx + x * s + 2.4 * s}" cy="${cy - 3.5 * s}" r="${2.8 * s}" fill="#fff"/>`;
  const wink = (x) =>
    `<path d="M${cx + x * s - 8 * s} ${cy + 2 * s} Q${cx + x * s} ${cy - 8 * s} ${cx + x * s + 8 * s} ${cy + 2 * s}" stroke="${P.ink}" stroke-width="${4.5 * s}" fill="none" stroke-linecap="round"/>`;
  const mouth =
    mood === 'open'
      ? `<path d="M${cx - 9 * s} ${cy + 15 * s} Q${cx} ${cy + 33 * s} ${cx + 9 * s} ${cy + 15 * s} Z" fill="${P.pinkDeep}" stroke="${P.ink}" stroke-width="${3.5 * s}" stroke-linejoin="round"/>`
      : `<path d="M${cx - 10 * s} ${cy + 16 * s} Q${cx} ${cy + 27 * s} ${cx + 10 * s} ${cy + 16 * s}" stroke="${P.ink}" stroke-width="${4.5 * s}" fill="none" stroke-linecap="round"/>`;
  return (
    (mood === 'wink' ? wink(-26) : eye(-26)) +
    eye(26) +
    `<ellipse cx="${cx - 44 * s}" cy="${cy + 17 * s}" rx="${11 * s}" ry="${6.5 * s}" fill="${P.pinkSoft}"/>` +
    `<ellipse cx="${cx + 44 * s}" cy="${cy + 17 * s}" rx="${11 * s}" ry="${6.5 * s}" fill="${P.pinkSoft}"/>` +
    mouth
  );
};

/** The guide character: a chubby little plane flying to the right. */
function plane(mood = 'smile') {
  return svg(
    520,
    330,
    `<path d="M262 138 L224 44 Q219 30 234 31 L268 34 Q281 35 288 46 L352 136 Z" fill="${P.blueSoft}"/>
     <path d="M110 140 L66 52 Q60 37 76 39 L116 44 Q129 46 135 58 L176 136 Z" fill="${P.blue}"/>
     <path d="M88 180 L36 166 Q21 163 27 177 L40 203 Q46 214 59 212 L118 202 Z" fill="${P.blueSoft}"/>
     <path d="M70 186 Q64 128 176 122 L346 122 Q470 126 484 190 Q478 246 346 252 L170 252 Q78 250 70 186 Z" fill="#fff" stroke="${P.blueDeep}" stroke-width="7" stroke-linejoin="round"/>
     <path d="M80 208 Q210 224 476 204 Q462 244 346 249 L170 249 Q92 247 80 208 Z" fill="${P.bluePale}"/>
     <path d="M78 206 Q210 222 478 202" stroke="${P.pink}" stroke-width="8" fill="none" stroke-linecap="round"/>
     <circle cx="168" cy="166" r="12" fill="${P.blueSoft}"/><circle cx="214" cy="166" r="12" fill="${P.blueSoft}"/><circle cx="260" cy="166" r="12" fill="${P.blueSoft}"/>
     <path d="M244 222 L198 300 Q192 314 208 314 L246 312 Q259 311 268 301 L350 224 Z" fill="${P.blue}" stroke="${P.blueDeep}" stroke-width="5" stroke-linejoin="round"/>
     <rect x="252" y="262" width="62" height="32" rx="16" fill="${P.blueDeep}"/><circle cx="304" cy="278" r="9" fill="${P.blueSoft}"/>
     ${face(396, 168, 1.05, mood)}`
  );
}

function cloud(mood) {
  return svg(
    360,
    210,
    `<g fill="#fff"><circle cx="92" cy="128" r="56"/><circle cx="160" cy="88" r="72"/><circle cx="244" cy="106" r="60"/><circle cx="290" cy="140" r="44"/><rect x="60" y="126" width="260" height="60" rx="30"/></g>
     <path d="M70 176 Q180 196 316 172 Q300 188 290 186 L90 186 Q78 186 70 176 Z" fill="${P.bluePale}"/>
     ${mood ? face(184, 128, 1.0, mood) : ''}`
  );
}

const petal = (angle, len, wid, fill, stroke) =>
  `<path transform="rotate(${angle})" d="M0 0 C ${-wid} ${-len * 0.28} ${-wid * 0.82} ${-len * 0.74} 0 ${-len} C ${wid * 0.82} ${-len * 0.74} ${wid} ${-len * 0.28} 0 0 Z" fill="${fill}" stroke="${stroke}" stroke-width="3" stroke-linejoin="round"/>`;

function lotus(withFace = false) {
  const back = [-78, -46, -15, 15, 46, 78].map((a) => petal(a, 118, 44, P.pinkSoft, P.pink)).join('');
  const front = [-58, -29, 0, 29, 58].map((a) => petal(a, 132, 46, a === 0 ? '#FFD6E0' : P.pinkPale, P.pink)).join('');
  return svg(
    340,
    250,
    `<ellipse cx="170" cy="218" rx="150" ry="26" fill="${P.green}"/>
     <ellipse cx="170" cy="212" rx="150" ry="24" fill="#8ADBB8"/>
     <g transform="translate(170 206)">${back}${front}
       <ellipse cx="0" cy="-16" rx="26" ry="16" fill="${P.yellow}"/>
       <g fill="#F2A93B"><circle cx="-12" cy="-18" r="3"/><circle cx="0" cy="-22" r="3"/><circle cx="12" cy="-18" r="3"/><circle cx="-6" cy="-12" r="3"/><circle cx="6" cy="-12" r="3"/></g>
     </g>${withFace ? face(170, 150, 0.72) : ''}`
  );
}

function leaf() {
  const veins = [20, 55, 90, 125, 160, 200, 240, 280, 320]
    .map((a) => {
      const r = (a * Math.PI) / 180;
      return `<line x1="150" y1="100" x2="${150 + 122 * Math.cos(r)}" y2="${100 + 78 * Math.sin(r)}" stroke="${P.greenDeep}" stroke-width="3" stroke-linecap="round" opacity=".55"/>`;
    })
    .join('');
  return svg(
    300,
    200,
    `<path d="M150 100 L288 60 A140 92 0 1 0 290 128 Z" fill="${P.green}" transform="rotate(-8 150 100)"/>
     <path d="M150 100 L276 66 A128 82 0 1 0 278 124 Z" fill="#8ADBB8" transform="rotate(-8 150 100)"/>${veins}
     <circle cx="150" cy="100" r="7" fill="${P.greenDeep}"/>`
  );
}

function bud() {
  return svg(
    120,
    260,
    `<path d="M60 250 Q52 180 60 96" stroke="${P.greenDeep}" stroke-width="7" fill="none" stroke-linecap="round"/>
     <g transform="translate(60 112)">${petal(-16, 96, 30, P.pinkSoft, P.pink)}${petal(16, 96, 30, P.pinkSoft, P.pink)}${petal(0, 104, 30, P.pinkPale, P.pink)}</g>`
  );
}

function suitcase(color = P.pink, deep = P.pinkDeep) {
  return svg(
    220,
    280,
    `<path d="M78 70 V34 Q78 20 92 20 H128 Q142 20 142 34 V70" fill="none" stroke="${P.ink}" stroke-width="10" stroke-linecap="round"/>
     <rect x="24" y="66" width="172" height="178" rx="28" fill="${color}" stroke="${deep}" stroke-width="6"/>
     <rect x="24" y="118" width="172" height="20" fill="${deep}" opacity=".35"/>
     <rect x="24" y="178" width="172" height="20" fill="${deep}" opacity=".35"/>
     <circle cx="64" cy="258" r="13" fill="${P.ink}"/><circle cx="156" cy="258" r="13" fill="${P.ink}"/>
     ${face(110, 144, 0.8)}
     <path d="M158 84 l6 12 13 2 -9.500 9 2.300 13 -11.800 -6.200 -11.800 6.200 2.300 -13 -9.500 -9 13 -2 Z" fill="${P.yellow}"/>`
  );
}

function tower() {
  return svg(
    200,
    330,
    `<path d="M78 320 L88 150 H112 L122 320 Z" fill="#fff" stroke="${P.blueDeep}" stroke-width="6" stroke-linejoin="round"/>
     <rect x="84" y="196" width="32" height="12" fill="${P.pinkSoft}"/><rect x="82" y="246" width="36" height="12" fill="${P.pinkSoft}"/>
     <path d="M52 150 H148 L164 96 H36 Z" fill="${P.blue}" stroke="${P.blueDeep}" stroke-width="6" stroke-linejoin="round"/>
     <path d="M50 104 H150 L143 134 H57 Z" fill="${P.bluePale}"/>
     <line x1="84" y1="104" x2="86" y2="134" stroke="${P.blueDeep}" stroke-width="4"/><line x1="116" y1="104" x2="114" y2="134" stroke="${P.blueDeep}" stroke-width="4"/>
     <rect x="40" y="76" width="120" height="22" rx="11" fill="#fff" stroke="${P.blueDeep}" stroke-width="6"/>
     <line x1="100" y1="76" x2="100" y2="34" stroke="${P.blueDeep}" stroke-width="6" stroke-linecap="round"/>
     <circle cx="100" cy="28" r="10" fill="${P.pink}"/>`
  );
}

function star(color = P.yellow) {
  return svg(
    120,
    120,
    `<path d="M60 8 C64 40 80 56 112 60 C80 64 64 80 60 112 C56 80 40 64 8 60 C40 56 56 40 60 8 Z" fill="${color}"/>`
  );
}

function heart(color = P.pink) {
  return svg(
    120,
    110,
    `<path d="M60 100 C16 70 6 46 10 32 C16 10 44 6 60 28 C76 6 104 10 110 32 C114 46 104 70 60 100 Z" fill="${color}"/>`
  );
}

function sun() {
  const rays = Array.from({ length: 12 }, (_, i) => `<rect x="146" y="6" width="8" height="34" rx="4" fill="${P.yellow}" transform="rotate(${i * 30} 150 150)"/>`).join('');
  return svg(300, 300, `${rays}<circle cx="150" cy="150" r="84" fill="${P.yellow}"/><circle cx="150" cy="150" r="84" fill="#FFE08A" opacity=".5"/>${face(150, 142, 1.05)}`);
}

// ---- round icons (white pictogram on a coloured disc) -----------------------------------

const disc = (color, body) => svg(200, 200, `<circle cx="100" cy="100" r="96" fill="${color}"/>${body}`);

const ICONS = {
  metro: `<rect x="56" y="40" width="88" height="104" rx="22" fill="#fff"/><rect x="68" y="56" width="64" height="38" rx="8" fill="COLOR"/><circle cx="76" cy="118" r="8" fill="COLOR"/><circle cx="124" cy="118" r="8" fill="COLOR"/><path d="M70 146 L56 168 M130 146 L144 168" stroke="#fff" stroke-width="10" stroke-linecap="round"/>`,
  train: `<path d="M40 128 Q40 74 110 70 H160 V128 Z" fill="#fff"/><path d="M62 104 Q70 84 104 82 H118 V104 Z" fill="COLOR"/><rect x="128" y="82" width="22" height="22" rx="4" fill="COLOR"/><rect x="34" y="136" width="132" height="10" rx="5" fill="#fff"/><circle cx="72" cy="128" r="9" fill="COLOR" stroke="#fff" stroke-width="5"/><circle cx="134" cy="128" r="9" fill="COLOR" stroke="#fff" stroke-width="5"/>`,
  bus: `<rect x="42" y="52" width="116" height="84" rx="16" fill="#fff"/><rect x="54" y="64" width="40" height="30" rx="5" fill="COLOR"/><rect x="104" y="64" width="42" height="30" rx="5" fill="COLOR"/><circle cx="70" cy="140" r="13" fill="#fff"/><circle cx="130" cy="140" r="13" fill="#fff"/><circle cx="70" cy="140" r="5" fill="COLOR"/><circle cx="130" cy="140" r="5" fill="COLOR"/>`,
  car: `<path d="M36 124 Q36 104 54 100 L70 70 Q74 62 84 62 H118 Q128 62 134 70 L150 100 Q164 104 164 124 V134 H36 Z" fill="#fff"/><path d="M76 98 L86 76 H116 L128 98 Z" fill="COLOR"/><circle cx="68" cy="136" r="14" fill="#fff"/><circle cx="132" cy="136" r="14" fill="#fff"/><circle cx="68" cy="136" r="6" fill="COLOR"/><circle cx="132" cy="136" r="6" fill="COLOR"/>`,
  parking: `<path d="M72 150 V52 H112 Q146 52 146 86 Q146 120 112 120 H94 V150 Z M94 72 V100 H110 Q124 100 124 86 Q124 72 110 72 Z" fill="#fff" fill-rule="evenodd"/>`,
  plane: `<path d="M100 34 Q110 34 112 54 L114 86 L164 116 V132 L114 116 L112 146 L128 158 V168 L100 160 L72 168 V158 L88 146 L86 116 L36 132 V116 L86 86 L88 54 Q90 34 100 34 Z" fill="#fff"/>`,
  globe: `<circle cx="100" cy="100" r="58" fill="none" stroke="#fff" stroke-width="10"/><ellipse cx="100" cy="100" rx="26" ry="58" fill="none" stroke="#fff" stroke-width="8"/><line x1="42" y1="100" x2="158" y2="100" stroke="#fff" stroke-width="8"/><path d="M52 70 Q100 86 148 70 M52 130 Q100 114 148 130" stroke="#fff" stroke-width="7" fill="none"/>`,
  box: `<path d="M100 40 L156 68 V132 L100 160 L44 132 V68 Z" fill="#fff"/><path d="M44 68 L100 96 L156 68 M100 96 V160" stroke="COLOR" stroke-width="8" fill="none" stroke-linejoin="round"/><path d="M72 54 L128 82 V104" stroke="COLOR" stroke-width="7" fill="none"/>`,
  pin: `<path d="M100 36 Q142 36 142 80 Q142 112 100 166 Q58 112 58 80 Q58 36 100 36 Z" fill="#fff"/><circle cx="100" cy="80" r="18" fill="COLOR"/>`,
  seat: `<circle cx="100" cy="60" r="20" fill="#fff"/><path d="M58 150 Q58 96 100 96 Q142 96 142 150 Z" fill="#fff"/>`,
  flag: `<line x1="64" y1="40" x2="64" y2="164" stroke="#fff" stroke-width="10" stroke-linecap="round"/><path d="M70 46 Q96 34 118 48 Q138 60 150 50 V104 Q138 114 118 102 Q96 88 70 100 Z" fill="#fff"/>`,
  ruler: `<rect x="34" y="78" width="132" height="44" rx="8" fill="#fff"/><path d="M58 78 V98 M82 78 V106 M106 78 V98 M130 78 V106 M150 78 V98" stroke="COLOR" stroke-width="7" stroke-linecap="round"/>`,
  horse: `<path d="M60 160 V118 Q52 96 70 78 L84 50 L96 66 Q132 62 146 96 Q152 116 146 132 V160 H126 V134 Q104 142 82 132 V160 Z" fill="#fff"/><circle cx="84" cy="80" r="5" fill="COLOR"/>`,
  medal: `<path d="M70 36 L100 86 L130 36" stroke="#fff" stroke-width="12" fill="none" stroke-linejoin="round"/><circle cx="100" cy="122" r="40" fill="#fff"/><path d="M100 98 l7 15 16 2 -12 11 3 16 -14 -8 -14 8 3 -16 -12 -11 16 -2 Z" fill="COLOR"/>`,
  leafIcon: `<path d="M48 150 Q40 60 152 48 Q160 150 62 152 Z" fill="#fff"/><path d="M56 148 Q90 110 132 70" stroke="COLOR" stroke-width="7" fill="none" stroke-linecap="round"/>`,
  clock: `<circle cx="100" cy="100" r="58" fill="#fff"/><path d="M100 64 V102 L126 118" stroke="COLOR" stroke-width="10" fill="none" stroke-linecap="round" stroke-linejoin="round"/>`,
};

function icon(name, color) {
  if (!ICONS[name]) throw new Error('unknown icon ' + name);
  return disc(color, ICONS[name].replace(/COLOR/g, color));
}

// ---- backgrounds ------------------------------------------------------------------------

const softCloud = (x, y, s, o) =>
  `<g transform="translate(${x} ${y}) scale(${s})" fill="#fff" opacity="${o}"><circle cx="92" cy="128" r="56"/><circle cx="160" cy="88" r="72"/><circle cx="244" cy="106" r="60"/><circle cx="290" cy="140" r="44"/><rect x="60" y="126" width="260" height="60" rx="30"/></g>`;

function bgMain() {
  return svg(
    1920,
    1080,
    `<defs><linearGradient id="g" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#E3F3FF"/><stop offset=".55" stop-color="#F6FBFF"/><stop offset="1" stop-color="#FFF4EA"/></linearGradient></defs>
     <rect width="1920" height="1080" fill="url(#g)"/>
     ${softCloud(1480, -40, 1.5, 0.75)}${softCloud(-120, 20, 1.1, 0.5)}${softCloud(1560, 900, 1.3, 0.8)}${softCloud(-60, 930, 1.6, 0.8)}${softCloud(760, 990, 1.2, 0.55)}`
  );
}

function bgCover() {
  return svg(
    1920,
    1080,
    `<defs><linearGradient id="g" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#BFE3FF"/><stop offset=".5" stop-color="#E6F4FF"/><stop offset="1" stop-color="#FFE9F0"/></linearGradient></defs>
     <rect width="1920" height="1080" fill="url(#g)"/>
     ${softCloud(-140, 760, 2.4, 0.85)}${softCloud(520, 880, 1.9, 0.9)}${softCloud(1180, 820, 2.3, 0.85)}${softCloud(1640, 930, 1.5, 0.95)}${softCloud(60, 60, 1.0, 0.6)}${softCloud(1250, -70, 1.4, 0.55)}`
  );
}

function bgPink() {
  return svg(
    1920,
    1080,
    `<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#FFE9F0"/><stop offset=".5" stop-color="#FFF8F2"/><stop offset="1" stop-color="#E3F3FF"/></linearGradient></defs>
     <rect width="1920" height="1080" fill="url(#g)"/>
     ${softCloud(1500, -30, 1.4, 0.8)}${softCloud(-100, 900, 1.7, 0.85)}${softCloud(1500, 930, 1.3, 0.8)}`
  );
}

/** A dashed flight trail; `d` is an SVG path in a w x h pixel box. */
function trail(w, h, d, color = P.blue, width = 7, dash = '4 22') {
  return svg(w, h, `<path d="${d}" fill="none" stroke="${color}" stroke-width="${width}" stroke-linecap="round" stroke-dasharray="${dash}"/>`);
}

module.exports = { P, plane, cloud, lotus, leaf, bud, suitcase, tower, star, heart, sun, icon, bgMain, bgCover, bgPink, trail };
