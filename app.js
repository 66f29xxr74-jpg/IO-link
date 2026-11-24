/* 도미노 경제 · 산업연관 파급 게임 — 계산 엔진과 게임 로직
   금액 단위: 백만 원 (생산자가격, 2023년 당해년 가격). 1조 원 = 1,000,000 */

const SRC = {
  large: ['io-2023-large.xlsx', '대분류 33부문'],
  mid:   ['io-2023-mid.xlsx',   '중분류 83부문'],
  small: ['io-2023-small.xlsx', '소분류 165부문'],
  basic: ['io-2023-basic.xlsx', '기본부문 380부문']
};

const MISSIONS = [
  { id: 'm1', t: '생산 파워',     d: '수요 투입의 생산유발 승수 2.0배 이상', ok: r => r.kind == 'demand' && r.mult >= 2 },
  { id: 'm2', t: '부가가치 장인', d: '수요 투입의 부가가치 유발률 40% 이상', ok: r => r.kind == 'demand' && r.va >= 0.4 },
  { id: 'm3', t: '국산화 챔피언', d: '국산화율 50% 이상으로 국산화 효과 실행', ok: r => r.kind == 'local' && r.alpha >= 50 },
  { id: 'm4', t: '누출 제로',     d: '수요 투입의 수입 유발률 10% 이하', ok: r => r.kind == 'demand' && r.imp <= 0.1 },
  { id: 'm5', t: '도미노 연쇄',   d: '원가 충격을 일으켜 물가 파급을 측정', ok: r => r.kind == 'price' }
];

const $ = s => document.querySelector(s);
const $$ = s => [...document.querySelectorAll(s)];
const sum = a => a.reduce((p, c) => p + c, 0);
const nm = s => String(s ?? '').replace(/\s/g, '');
const pct = v => (v * 100).toFixed(1) + '%';
const won = v => {
  const a = Math.abs(v), s = v < 0 ? '-' : '';
  return s + (a >= 1e6 ? (a / 1e6).toFixed(2) + ' 조' : a >= 100 ? Math.round(a / 100).toLocaleString('ko-KR') + ' 억' : Math.round(a).toLocaleString('ko-KR') + ' 백만') + ' 원';
};

let P = { xp: 0, done: {} };
try { P = JSON.parse(localStorage.getItem('rp_v1')) || P; } catch (e) {}
const save = () => { try { localStorage.setItem('rp_v1', JSON.stringify(P)); } catch (e) {} };

let D = null, cache = {}, cur = 0, mode = 'demand', amt = 1, alp = 30, shk = 10, sheet = '국산투입계수(Ad)';

/* ---------- 엑셀 파싱: 머리글 아래 첫 숫자 행부터 부문 행을 읽음 ---------- */
function parse(rows) {
  const st = rows.findIndex(r => r && r[0] != null && r[1] != null && typeof r[2] === 'number');
  const hdr = rows[st - 1] || [];
  let N = 0;
  for (let c = 2; c < hdr.length; c++) {
    const v = nm(hdr[c]);
    if (!v || /^(중간수요계|중간투입계|총투입계)$/.test(v)) break;
    N++;
  }
  const names = [], M = [];
  for (let r = st; r < rows.length && M.length < N; r++) {
    const row = rows[r];
    if (row && row[0] != null && row[1] != null && typeof row[2] === 'number') {
      names.push(String(row[1]).trim());
      M.push(Array.from({ length: N }, (_, j) => Number(row[2 + j]) || 0));
    }
  }
  if (N < 2 || M.length !== N) throw new Error('투입산출표 구조를 인식하지 못했습니다');
  return { names, M, N };
}

/* ---------- 가우스-조르당 역행렬 ---------- */
function inv(A) {
  const n = A.length;
  const a = A.map((r, i) => [...r, ...Array.from({ length: n }, (_, j) => +(i == j))]);
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(a[r][c]) > Math.abs(a[p][c])) p = r;
    [a[c], a[p]] = [a[p], a[c]];
    const d = a[c][c];
    for (let k = 0; k < 2 * n; k++) a[c][k] /= d;
    for (let r = 0; r < n; r++) if (r != c) {
      const f = a[r][c];
      if (f) for (let k = 0; k < 2 * n; k++) a[r][k] -= f * a[c][k];
    }
  }
  return a.map(r => r.slice(n));
}

/* ---------- 데이터 로드와 모형 구축 ---------- */
async function loadLevel(lv) {
  if (cache[lv]) return cache[lv];
  const res = await fetch('data/' + SRC[lv][0]);
  if (!res.ok) throw new Error('data/' + SRC[lv][0] + ' 파일을 찾을 수 없습니다 (' + res.status + ')');
  const wb = XLSX.read(new Uint8Array(await res.arrayBuffer()), { type: 'array', sheets: ['국산투입계수(Ad)', '수입투입계수(Am)'] });
  const rowsOf = n => {
    if (!wb.Sheets[n]) throw new Error('시트를 찾을 수 없습니다: ' + n);
    return XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: null });
  };
  const a = parse(rowsOf('국산투입계수(Ad)')), b = parse(rowsOf('수입투입계수(Am)'));
  const N = a.N, Ad = a.M, Am = b.M;
  const L = inv(Ad.map((r, i) => r.map((x, j) => (i == j) - x)));       // (I − Ad)⁻¹
  const mult = Array.from({ length: N }, (_, j) => sum(L.map(r => r[j]))); // 생산유발계수 = 열합
  const m = Array.from({ length: N }, (_, j) => sum(Am.map(r => r[j])));   // 수입의존도
  const v = Array.from({ length: N }, (_, j) => 1 - sum(Ad.map(r => r[j])) - m[j]); // 부가가치율
  return (cache[lv] = { lv, N, names: a.names, Ad, Am, L, mult, m, v });
}

/* ---------- 시나리오 계산 ---------- */
// 수요 투입: x = L·f, f = F·e_k
function runDemand(k, F) {
  const x = D.L.map(r => r[k] * F);
  const tot = sum(x), va = sum(x.map((z, i) => z * D.v[i])), imp = sum(x.map((z, i) => z * D.m[i]));
  return { kind: 'demand', k, F, x, tot, mult: tot / F, va: va / tot, imp: imp / tot, vaAbs: va, impAbs: imp };
}

// 국산화: 산업 k의 수입 중간재 α를 국산으로 대체 (Sherman–Morrison, 역행렬 재계산 없음)
function runLocal(k, F, a) {
  const x0 = D.L.map(r => r[k] * F);
  const u = D.Am.map(r => a * r[k]);                       // 바뀌는 국산투입 증분 (k열)
  const s = sum(D.L[k].map((c, j) => c * u[j]));           // e_kᵀ L u
  const w = D.L.map(r => sum(r.map((c, j) => c * u[j])));  // L u
  const g = D.L[k][k] * F / (1 - s);
  const x = x0.map((z, i) => z + w[i] * g);
  const m1 = D.m.map((z, j) => (j == k ? z * (1 - a) : z));
  const tot0 = sum(x0), tot = sum(x);
  const imp0 = sum(x0.map((z, i) => z * D.m[i])), imp1 = sum(x.map((z, i) => z * m1[i]));
  return { kind: 'local', k, F, alpha: a * 100, x, x0, tot, mult: tot / F, mult0: tot0 / F, extra: tot - tot0, saved: imp0 - imp1, imp: imp1 / tot, va: sum(x.map((z, i) => z * D.v[i])) / tot };
}

// 원가 충격: p_k 고정, 나머지는 p_j = Σ_i A_ij p_i (A = Ad + Am) 반복
function runPrice(k, s) {
  const N = D.N, p = Array(N).fill(0);
  p[k] = s / 100;
  for (let it = 0; it < 300; it++)
    for (let j = 0; j < N; j++) if (j != k) {
      let t = 0;
      for (let i = 0; i < N; i++) t += (D.Ad[i][j] + D.Am[i][j]) * p[i];
      p[j] = t;
    }
  const o = [...Array(N).keys()].filter(j => j != k);
  return {
    kind: 'price', k, s, p,
    avg: sum(o.map(j => p[j])) / o.length,
    top: [...o].sort((a, b) => p[b] - p[a]).slice(0, 6),
    n05: o.filter(j => p[j] >= 0.005).length,
    n1: o.filter(j => p[j] >= 0.01).length
  };
}

/* ---------- 화면: 타일 격자 ---------- */
function renderGrid() {
  const mx = Math.max(...D.mult), mn = Math.min(...D.mult), span = (mx - mn) || 1;
  $('#grid').innerHTML = D.names.map((n, i) =>
    `<button class="t" data-i="${i}" style="--t:${((D.mult[i] - mn) / span).toFixed(3)}" aria-label="${n}, 생산유발 ${D.mult[i].toFixed(2)}배">` +
    `<span class="tn">${n}</span><span class="tm">생산유발 ×${D.mult[i].toFixed(2)}</span></button>`).join('');
  $('#q').value = '';
  filterGrid('');
  pick(cur);
}

function filterGrid(q) {
  const k = nm(q);
  let n = 0;
  $$('.t').forEach(b => {
    const show = !k || nm(D.names[b.dataset.i]).includes(k);
    b.hidden = !show;
    if (show) n++;
  });
  $('#count').textContent = `${n}/${D.N}개 부문 · 진할수록 생산유발 효과가 큽니다`;
}

function pick(i) {
  cur = i;
  $$('.t').forEach(b => b.classList.toggle('sel', +b.dataset.i === i));
  $('#pname').textContent = D.names[i];
  $('#pmeta').textContent = `생산유발계수 ×${D.mult[i].toFixed(2)} · 부가가치율 ${pct(D.v[i])} · 수입의존도 ${pct(D.m[i])}`;
  renderCtrl();
  if (!$('#explore').hidden) renderExplore();
}

function renderCtrl() {
  const sl = (id, lab, min, max, st, val, unit) =>
    `<label class="sl"><span>${lab}<b id="${id}v">${val}${unit}</b></span><input type="range" id="${id}" min="${min}" max="${max}" step="${st}" value="${val}"></label>`;
  $('#ctrl').innerHTML = mode == 'price'
    ? sl('shk', '원가 상승률 (충격 산업)', 1, 30, 1, shk, '%')
    : sl('amt', '최종수요 투입 금액', 0.1, 10, 0.1, amt.toFixed(1), ' 조 원') +
      (mode == 'local' ? sl('alp', '국산화율 (수입 중간재를 국산으로 대체)', 10, 100, 5, alp, '%') : '');
}

/* ---------- 게임 실행 ---------- */
function go() {
  if (!D) return;
  const k = cur;
  const r = mode == 'demand' ? runDemand(k, amt * 1e6)
          : mode == 'local'  ? runLocal(k, amt * 1e6, alp / 100)
          : runPrice(k, shk);
  renderRes(r);
  animate(r);
  award(r);
}

function renderRes(r) {
  const N = D.names, name = N[r.k];
  const rank = arr => [...arr.keys()].sort((a, b) => arr[b] - arr[a]).slice(0, 6);
  let title, sub, kp, items, say, stars;

  if (r.kind == 'demand') {
    title = `${name} 수요 ${won(r.F)} 투입`; sub = '직접·간접 생산 파급 결과';
    const t = rank(r.x);
    kp = [['총 생산유발', won(r.tot), `승수 ${r.mult.toFixed(2)}배`],
          ['부가가치 유발', won(r.vaAbs), `유발률 ${pct(r.va)}`],
          ['수입 유발', won(r.impAbs), `유발률 ${pct(r.imp)}`],
          ['가장 큰 파급', N[t[0]], won(r.x[t[0]])]];
    items = t.map(j => [N[j], r.x[j], won(r.x[j])]);
    say = `${name}에 ${won(r.F)}을 투입하면 전체 경제에서 ${won(r.tot)}의 생산이 일어나고, 그중 ${won(r.vaAbs)}이 부가가치로 남습니다.`;
    stars = r.mult >= 2.2 ? '★★★' : r.mult >= 1.9 ? '★★' : '★';
  } else if (r.kind == 'local') {
    const diff = r.x.map((z, i) => z - r.x0[i]), t = rank(diff);
    title = `${name} 수입 중간재 ${r.alpha}% 국산화`; sub = '수입을 국내 공급으로 대체했을 때의 순증 효과';
    kp = [['추가 생산', won(r.extra), `승수 ${r.mult0.toFixed(2)} → ${r.mult.toFixed(2)}배`],
          ['수입 절감', won(r.saved), '해외로 빠지던 금액'],
          ['국산화율', r.alpha + '%', `${name} 기준`],
          ['가장 큰 효과', N[t[0]], '+' + won(diff[t[0]])]];
    items = t.map(j => [N[j], diff[j], '+' + won(diff[j])]);
    say = `${name}의 수입 중간재 ${r.alpha}%를 국산으로 바꾸면 생산은 ${won(r.extra)} 늘고 수입은 ${won(r.saved)} 줄어듭니다.`;
    stars = r.saved / r.tot >= 0.05 ? '★★★' : r.saved / r.tot >= 0.02 ? '★★' : '★';
  } else {
    title = `${name} 원가 ${r.s}% 상승`; sub = '비용 전가를 통한 산업별 물가 파급';
    kp = [['다른 산업 평균 전가', (r.avg * 100).toFixed(2) + '%', '단순 평균'],
          ['0.5% 이상 상승', r.n05 + '개 산업', `전체 ${N.length - 1}개 중`],
          ['1% 이상 상승', r.n1 + '개 산업', '원가 부담 집중'],
          ['가장 큰 파급', N[r.top[0]], '+' + (r.p[r.top[0]] * 100).toFixed(2) + '%']];
    items = r.top.map(j => [N[j], r.p[j] * 100, (r.p[j] * 100).toFixed(2) + '%']);
    say = `${name} 원가가 ${r.s}% 오르면 1% 이상 오르는 산업이 ${r.n1}개입니다. 가장 크게 영향받는 곳은 ${N[r.top[0]]}입니다.`;
    stars = r.n1 >= 10 ? '★★★' : r.n1 >= 3 ? '★★' : '★';
  }

  const mx = Math.max(...items.map(i => Math.abs(i[1])), 1e-9);
  $('#res').innerHTML =
    `<h3>${title}</h3><div class="sub2">${sub}</div>` +
    `<div class="kp">${kp.map(k => `<div><span>${k[0]}</span><b>${k[1]}</b><small>${k[2]}</small></div>`).join('')}</div>` +
    `<div class="bars">${items.map(i => `<div class="br"><span>${i[0]}</span><div class="tk"><i style="width:${Math.abs(i[1]) / mx * 100}%"></i></div><em>${i[2]}</em></div>`).join('')}</div>` +
    `<p class="say">${say}</p>` +
    `<p class="stars" aria-label="파급 등급">${stars}</p>` +
    `<p class="sub2" style="margin:8px 0 0">출처: 한국은행 2023년 산업연관표(연장표) · 선형 모형 추정치</p>`;
}

function animate(r) {
  const score = r.kind == 'demand' ? r.x : r.kind == 'local' ? r.x.map((z, i) => z - r.x0[i]) : r.p;
  [...score.keys()].sort((a, b) => score[b] - score[a]).slice(0, 40).forEach((j, n) => setTimeout(() => {
    const el = $(`.t[data-i="${j}"]`);
    if (!el) return;
    el.classList.remove('hit'); void el.offsetWidth; el.classList.add('hit');
  }, n * 32));
}

function award(r) {
  P.xp += 10;
  const got = [];
  MISSIONS.forEach(m => { if (!P.done[m.id] && m.ok(r)) { P.done[m.id] = 1; P.xp += 50; got.push(m.t); } });
  save(); renderXP(); renderMissions();
  if (got.length) toast(`미션 달성 · ${got.join(', ')} (+50 XP)`);
}

/* ---------- 미션, XP, 알림 ---------- */
function renderXP() {
  $('#lv').textContent = 'Lv.' + (Math.floor(P.xp / 100) + 1);
  $('#xpb').style.width = (P.xp % 100) + '%';
  $('#xpb').parentElement.title = P.xp + ' XP';
}

function renderMissions() {
  $('#mlist').innerHTML = MISSIONS.map((m, i) => {
    const ok = !!P.done[m.id];
    return `<article class="mc${ok ? ' ok' : ''}"><div class="mi">${ok ? '✓' : i + 1}</div><div>` +
      `<h3>${m.t}</h3><p>${m.d}</p>` +
      `<p style="margin-top:6px;font-size:12.5px;color:${ok ? 'var(--gold)' : 'var(--mut)'}">${ok ? '달성 완료 · +50 XP' : '미달성'}</p></div></article>`;
  }).join('');
}

function toast(t) {
  const el = document.createElement('div');
  el.className = 'toast';
  el.textContent = t;
  document.body.appendChild(el);
  setTimeout(() => el.remove(), 2600);
}

function setStatus(t, err) {
  const s = $('#status');
  s.textContent = t;
  s.classList.toggle('err', !!err);
}

/* ---------- 데이터 탐색 ---------- */
const MATS = {
  '국산투입계수(Ad)': (i, j) => D.Ad[i][j],
  '수입투입계수(Am)': (i, j) => D.Am[i][j],
  '총투입계수(A=Ad+Am)': (i, j) => D.Ad[i][j] + D.Am[i][j],
  '생산유발계수(L)': (i, j) => D.L[i][j]
};

function renderSheets() {
  $('#sheets').innerHTML = Object.keys(MATS).map(s => `<button data-s="${s}" aria-pressed="${s == sheet}">${s}</button>`).join('');
}

function renderExplore() {
  if (!D) return;
  const f = MATS[sheet], q = nm($('#eq').value);
  const rows = D.names.map((n, i) => ({ n, v: f(i, cur) }))
    .filter(r => !q || nm(r.n).includes(q))
    .sort((a, b) => Math.abs(b.v) - Math.abs(a.v));
  const mx = Math.max(...rows.map(r => Math.abs(r.v)), 1e-12);
  const unit = sheet.startsWith('생산유발') ? '배 (최종수요 1원당 생산유발, 무차원)' : '계수 (무차원, 열 산업 생산 1원당 행 산업 투입액)';
  $('#exmeta').textContent = `열 기준: ${D.names[cur]} · 표시 ${Math.min(rows.length, 300)}/${rows.length}개 · 단위: ${unit}`;
  $('#extb').innerHTML = rows.slice(0, 300).map((r, k) =>
    `<tr><td>${k + 1}</td><td>${r.n}</td><td>${r.v.toFixed(4)}</td><td><div class="tk"><i style="width:${Math.abs(r.v) / mx * 100}%"></i></div></td></tr>`).join('');
}

/* ---------- 부팅과 이벤트 ---------- */
async function boot(lv) {
  setStatus(`${SRC[lv][1]} 자료를 불러오는 중입니다… 기본부문(12MB)은 수십 초 걸릴 수 있어요.`);
  $('#go').disabled = true;
  try {
    D = await loadLevel(lv);
    const k = D.names.findIndex(n => /^건설/.test(nm(n)));
    cur = k >= 0 ? k : 0;
    setStatus(`${SRC[lv][1]} 자료 ${D.N}개 부문을 불러왔습니다. 단위: 백만 원 (생산자가격, 2023년 당해년 가격)`);
    $('#go').disabled = false;
    renderGrid();
    renderExplore();
  } catch (e) {
    setStatus('자료를 불러오지 못했습니다. ' + e.message + ' · data 폴더와 파일 이름, GitHub Pages 주소로 접속했는지 확인하세요.', true);
  }
}

$$('.tabs button').forEach(b => b.onclick = () => {
  $$('.tabs button').forEach(x => x.setAttribute('aria-selected', x == b));
  $$('.tab').forEach(s => s.hidden = s.id != b.dataset.tab);
  if (b.dataset.tab == 'explore') renderExplore();
});
$$('[data-mode]').forEach(b => b.onclick = () => {
  mode = b.dataset.mode;
  $$('[data-mode]').forEach(x => x.setAttribute('aria-pressed', x == b));
  renderCtrl();
});
$('#load').onclick = () => boot($('#level').value);
$('#level').onchange = () => boot($('#level').value);
$('#grid').onclick = e => { const b = e.target.closest('.t'); if (b) pick(+b.dataset.i); };
$('#go').onclick = go;
$('#sheets').onclick = e => {
  const b = e.target.closest('[data-s]');
  if (!b) return;
  sheet = b.dataset.s;
  renderSheets();
  renderExplore();
};
$('#reset').onclick = () => {
  if (confirm('미션과 XP를 초기화할까요?')) { P = { xp: 0, done: {} }; save(); renderXP(); renderMissions(); }
};
document.addEventListener('input', e => {
  const t = e.target;
  if (t.id == 'amt') { amt = +t.value; $('#amtv').textContent = amt.toFixed(1) + ' 조 원'; }
  else if (t.id == 'alp') { alp = +t.value; $('#alpv').textContent = alp + '%'; }
  else if (t.id == 'shk') { shk = +t.value; $('#shkv').textContent = shk + '%'; }
  else if (t.id == 'q' && D) filterGrid(t.value);
  else if (t.id == 'eq') renderExplore();
});

renderXP();
renderMissions();
renderSheets();
boot('large');
