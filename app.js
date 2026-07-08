/* 쿠팡 로켓그로스 매출·이익 계산기 — 전량 브라우저 내 처리 */
"use strict";

/* ---------- 유틸 ---------- */
const $ = (s) => document.querySelector(s);
const norm = (s) => (s == null ? "" : String(s).replace(/\s+/g, "").trim());
function idStr(v) {
  if (v == null || v === "") return "";
  if (typeof v === "number") return String(Math.trunc(v));
  let s = String(v).trim();
  if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, "");
  return s;
}
function num(v) {
  if (v == null || v === "") return 0;
  if (typeof v === "number") return v;
  const n = parseFloat(String(v).replace(/[, ]/g, ""));
  return isNaN(n) ? 0 : n;
}
const won = (n) => Math.round(n).toLocaleString("ko-KR");
function pct(n) {
  if (!isFinite(n)) return "-";
  return (n * 100).toFixed(1) + "%";
}

/* 컬럼 정의 & 계산식 (웹 툴팁 + 엑셀 메모 공용) */
const COLUMN_DEFS = {
  date: "날짜\n· 광고 리포트 파일명에서 읽은 판매 일자입니다.\n· 예: ..._20260705_...xlsx → 2026-07-05",
  group: "상품구분\n· 마진표의 '상품 구분' 값을 그대로 표시합니다.",
  name: "상품명\n· 마진표의 '등록 상품명'을 표시합니다. (없으면 인사이트 상품명)",
  regId: "등록상품ID\n· 쿠팡 등록상품ID. 세 파일(인사이트·광고·마진)을 연결하는 매칭 기준 키입니다.",
  qty: "판매수량\n· 인사이트의 옵션ID별 '판매량'을 등록상품ID 기준으로 합산한 값입니다.\n· = Σ 인사이트 판매량 (취소 반영된 순 판매량)",
  costPerUnit: "구매당비용\n· 이 상품 1개를 팔기 위해 실제로 쓴 광고비입니다. (손익분기 판단용)\n· = 사용광고비 ÷ 판매수량\n· 개당마진보다 크면(빨강) 그 상품은 적자입니다.\n· 판매수량이 0이면 '-'로 표시합니다.",
  marginUnit: "개당마진(손익분기)\n· 1개 팔 때 남는 마진, 곧 손익분기점입니다.\n· = 마진(광고전) ÷ 판매수량 (마진표의 개당마진)\n· 구매당비용이 이 값을 넘으면 적자입니다.",
  insightRev: "인사이트매출\n· 쿠팡 인사이트의 '매출(원)' 합계입니다. (판매가 기준 총매출)\n· = Σ 인사이트 매출(원)",
  settleRev: "정산매출\n· 쿠팡이 실제로 우리에게 정산해주는 금액 기준 매출입니다.\n· = 결제받는단가 × 판매수량",
  adCost: "사용광고비\n· 광고 리포트의 '광고비'를, 광고집행 옵션ID를 인사이트의 옵션ID→등록상품ID 매핑으로 상품에 붙여 합산한 값입니다.\n· = Σ 광고비 (해당 상품의 모든 광고옵션·캠페인)",
  marginBefore: "마진(광고전)\n· 광고비를 빼기 전, 상품 판매로 얻은 마진입니다.\n· = 개당마진 × 판매수량\n· (개당마진 = 결제받는단가 − 수입원가, 마진표 값)",
  netProfit: "순이익\n· 광고비까지 반영한 최종 이익금입니다.\n· = 마진(광고전) − 사용광고비",
  netRate: "순이익률\n· 정산매출 대비 순이익 비율입니다.\n· = 순이익 ÷ 정산매출",
};

/* 시트 배열에서 지정 헤더들을 모두 포함하는 시트를 찾아 {header,rows} 반환 */
function findSheet(wb, required) {
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    if (!rows.length) continue;
    // 헤더가 첫 행이 아닐 수 있으니 상위 5행 탐색
    for (let hr = 0; hr < Math.min(5, rows.length); hr++) {
      const header = (rows[hr] || []).map(norm);
      if (required.every((req) => header.includes(norm(req)))) {
        return { header: rows[hr], rows: rows.slice(hr + 1) };
      }
    }
  }
  return null;
}
function colIndex(header, candidates) {
  const H = header.map(norm);
  for (const c of candidates) {
    const i = H.indexOf(norm(c));
    if (i >= 0) return i;
  }
  return -1;
}

/* ---------- 파일 종류 자동 인식 ---------- */
function classify(wb) {
  // 마진표를 먼저 판별 (마진 워크북 안에 인사이트 형태 시트가 섞여 있을 수 있음)
  if (findSheet(wb, ["결제받는단가", "마진"])) return "margin";
  // 이 앱이 만든 결과(누적) 엑셀 — 다시 넣으면 이어서 기록 (상세 시트: 등록상품ID 포함)
  if (findSheet(wb, ["등록상품ID", "정산매출", "마진(광고전)"])) return "result";
  if (findSheet(wb, ["광고집행 옵션ID", "광고비"])) return "ad";
  if (findSheet(wb, ["옵션 ID", "등록상품ID", "매출(원)"])) return "insight";
  return "unknown";
}

/* 예전에 다운로드한 결과(누적) 엑셀을 다시 읽어 일자별 데이터로 복원 */
function parseResult(wb) {
  const s = findSheet(wb, ["등록상품ID", "정산매출", "마진(광고전)"]);
  if (!s) throw new Error("결과 엑셀 형식을 인식하지 못했습니다.");
  const H = s.header;
  const ci = {
    date: colIndex(H, ["날짜"]), group: colIndex(H, ["상품구분"]),
    name: colIndex(H, ["상품명"]), regId: colIndex(H, ["등록상품ID"]),
    qty: colIndex(H, ["판매수량"]), insightRev: colIndex(H, ["인사이트매출"]),
    settleRev: colIndex(H, ["정산매출"]), adCost: colIndex(H, ["사용광고비"]),
    marginBefore: colIndex(H, ["마진(광고전)"]), netProfit: colIndex(H, ["순이익"]),
    marginUnit: colIndex(H, ["개당마진"]),
  };
  const days = {};
  for (const r of s.rows) {
    const date = r[ci.date];
    if (!date) continue;
    const ds = String(date).trim();
    if (ds === "합계" || !/^\d{4}-\d{2}-\d{2}$/.test(ds)) continue; // 합계행·잡행 제외
    if (!days[ds]) days[ds] = { rows: [], soldNoMargin: [], unmappedAd: [] };
    const settleRev = num(r[ci.settleRev]);
    const netProfit = num(r[ci.netProfit]);
    const qty = num(r[ci.qty]);
    const marginBefore = num(r[ci.marginBefore]);
    // 개당마진: 저장 컬럼 우선, 없으면(옛 파일) 마진(광고전)÷수량으로 복원
    const marginUnit = ci.marginUnit >= 0 ? num(r[ci.marginUnit]) : (qty > 0 ? marginBefore / qty : NaN);
    days[ds].rows.push({
      date: ds, regId: idStr(r[ci.regId]),
      group: r[ci.group] || "", name: r[ci.name] || "",
      qty, insightRev: num(r[ci.insightRev]),
      settleRev, adCost: num(r[ci.adCost]),
      marginBefore, marginUnit, netProfit,
      netRate: settleRev > 0 ? netProfit / settleRev : NaN,
    });
  }
  return days;
}

/* ---------- 파서 ---------- */
function parseMargin(wb) {
  const s = findSheet(wb, ["결제받는단가", "마진"]);
  if (!s) throw new Error("마진표 형식을 인식하지 못했습니다.");
  const H = s.header;
  const cId = colIndex(H, ["등록 상품 아이디", "등록상품아이디", "등록상품ID"]);
  const cSettle = colIndex(H, ["결제받는단가"]);
  const cMargin = colIndex(H, ["마진"]);
  const cGroup = colIndex(H, ["상품 구분", "상품구분"]);
  const cName = colIndex(H, ["등록 상품명", "등록상품명", "상품명"]);
  const map = {};
  for (const r of s.rows) {
    const id = idStr(r[cId]);
    if (!id) continue;
    const mv = r[cMargin];
    if (mv == null || mv === "") continue; // 마진 미기입 행은 제외
    map[id] = {
      group: cGroup >= 0 ? (r[cGroup] || "") : "",
      name: cName >= 0 ? (r[cName] || "") : "",
      settle: num(r[cSettle]),
      margin: num(mv),
    };
  }
  return map; // regId -> {group,name,settle,margin(개당)}
}

function parseInsight(wb) {
  const s = findSheet(wb, ["옵션 ID", "등록상품ID", "매출(원)"]);
  if (!s) throw new Error("인사이트 파일 형식을 인식하지 못했습니다.");
  const H = s.header;
  const cOpt = colIndex(H, ["옵션 ID", "옵션ID"]);
  const cReg = colIndex(H, ["등록상품ID", "등록상품아이디"]);
  const cName = colIndex(H, ["상품명"]);
  const cQty = colIndex(H, ["판매량"]);
  const cRev = colIndex(H, ["매출(원)"]);
  const opt2reg = {};
  const regAgg = {}; // regId -> {qty,rev,name}
  for (const r of s.rows) {
    const opt = idStr(r[cOpt]);
    const reg = idStr(r[cReg]);
    if (opt && reg) opt2reg[opt] = reg;
    if (!reg) continue;
    if (!regAgg[reg]) regAgg[reg] = { qty: 0, rev: 0, name: cName >= 0 ? (r[cName] || "") : "" };
    regAgg[reg].qty += num(r[cQty]);
    regAgg[reg].rev += num(r[cRev]);
  }
  return { opt2reg, regAgg };
}

function parseAd(wb) {
  const s = findSheet(wb, ["광고집행 옵션ID", "광고비"]);
  if (!s) throw new Error("광고 리포트 형식을 인식하지 못했습니다.");
  const H = s.header;
  const cOpt = colIndex(H, ["광고집행 옵션ID", "광고집행옵션ID"]);
  const cCost = colIndex(H, ["광고비"]);
  const byOpt = {};
  for (const r of s.rows) {
    const opt = idStr(r[cOpt]);
    if (!opt) continue;
    byOpt[opt] = (byOpt[opt] || 0) + num(r[cCost]);
  }
  return byOpt; // 광고집행 옵션ID -> 광고비합
}

/* 파일명에서 날짜(YYYYMMDD) 추출 → YYYY-MM-DD
   계정번호 등 8자리 숫자가 앞에 있을 수 있으므로 '유효한 날짜'만 채택 */
function dateFromName(fname) {
  const all = fname.match(/\d{8}/g) || [];
  for (const d of all) {
    const y = +d.slice(0, 4), m = +d.slice(4, 6), day = +d.slice(6, 8);
    if (y >= 2000 && y <= 2099 && m >= 1 && m <= 12 && day >= 1 && day <= 31) {
      return `${d.slice(0, 4)}-${d.slice(4, 6)}-${d.slice(6, 8)}`;
    }
  }
  return null;
}

/* ---------- 계산 ---------- */
function computeDay(date, marginMap, insight, adByOpt) {
  const { opt2reg, regAgg } = insight;

  // 광고비를 등록상품ID로 매핑
  const adByReg = {};
  const unmappedAd = []; // {opt,cost}
  for (const [opt, cost] of Object.entries(adByOpt)) {
    const reg = opt2reg[opt];
    if (!reg) {
      if (cost > 0) unmappedAd.push({ opt, cost });
      continue;
    }
    adByReg[reg] = (adByReg[reg] || 0) + cost;
  }

  // 결과 행: 마진 기입 상품 중 (판매 있음 or 광고비 있음)
  const rows = [];
  const regs = new Set([...Object.keys(marginMap)]);
  for (const reg of regs) {
    const m = marginMap[reg];
    const agg = regAgg[reg] || { qty: 0, rev: 0, name: "" };
    const adCost = adByReg[reg] || 0;
    if (agg.qty === 0 && adCost === 0) continue; // 판매·광고 모두 없음 → 생략
    const qty = agg.qty;
    const insightRev = agg.rev;
    const settleRev = m.settle * qty;
    const marginBefore = m.margin * qty;
    const netProfit = marginBefore - adCost;
    rows.push({
      date, regId: reg,
      group: m.group || "",
      name: m.name || agg.name || "",
      qty, insightRev, settleRev, adCost, marginBefore, netProfit,
      marginUnit: m.margin, // 개당마진(손익분기) — 판매수량 0이어도 유효
      netRate: settleRev > 0 ? netProfit / settleRev : NaN,
    });
  }

  // 경고: 판매됐으나 마진 미기입
  const soldNoMargin = [];
  for (const [reg, agg] of Object.entries(regAgg)) {
    if (agg.qty > 0 && !marginMap[reg]) {
      soldNoMargin.push({ regId: reg, name: agg.name, qty: agg.qty, rev: agg.rev });
    }
  }
  soldNoMargin.sort((a, b) => b.rev - a.rev);
  rows.sort((a, b) => b.settleRev - a.settleRev);
  return { rows, soldNoMargin, unmappedAd };
}

/* ---------- 일별 요약 ---------- */
// 한 날짜의 상세 rows → 요약 배열 [날짜,수량,인사이트매출,정산매출,광고비,마진(광고전),순이익,순이익률]
function summaryRow(date, rows) {
  const s = (k) => rows.reduce((a, r) => a + (r[k] || 0), 0);
  const settle = s("settleRev"), net = s("netProfit");
  return [date, Math.round(s("qty")), Math.round(s("insightRev")), Math.round(settle),
    Math.round(s("adCost")), Math.round(s("marginBefore")), Math.round(net),
    settle > 0 ? +(net / settle).toFixed(4) : ""];
}
// 개당마진: 저장된 marginUnit 우선, 없으면(옛 누적본) 마진(광고전)÷수량으로 복원
function muOf(r) {
  if (isFinite(r.marginUnit)) return r.marginUnit;
  return r.qty > 0 ? r.marginBefore / r.qty : NaN;
}

// 상세 rows → 시트/엑셀용 배열
function detailRow(r) {
  const cpu = r.qty > 0 ? Math.round(r.adCost / r.qty) : "";      // 구매당비용(파생)
  const mu = isFinite(muOf(r)) ? Math.round(muOf(r)) : "";          // 개당마진
  return [r.date, r.group, r.name, r.regId, Math.round(r.qty), cpu, mu,
    Math.round(r.insightRev), Math.round(r.settleRev), Math.round(r.adCost),
    Math.round(r.marginBefore), Math.round(r.netProfit),
    isFinite(r.netRate) ? +r.netRate.toFixed(4) : ""];
}

/* ---------- 상태 (localStorage) ---------- */
const LS_MARGIN = "cm_margin_v1";
const LS_DAYS = "cm_days_v1";
const LS_GSHEET = "cm_gsheet_v1";
function loadMargin() {
  try { return JSON.parse(localStorage.getItem(LS_MARGIN) || "null"); } catch { return null; }
}
function saveMargin(map, meta) {
  localStorage.setItem(LS_MARGIN, JSON.stringify({ map, meta }));
}
function loadDays() {
  try { return JSON.parse(localStorage.getItem(LS_DAYS) || "{}"); } catch { return {}; }
}
function saveDays(days) { localStorage.setItem(LS_DAYS, JSON.stringify(days)); }
function loadGSheet() {
  try { return JSON.parse(localStorage.getItem(LS_GSHEET) || "null"); } catch { return null; }
}
function saveGSheet(cfg) { localStorage.setItem(LS_GSHEET, JSON.stringify(cfg)); }

/* 구글시트(Apps Script 웹앱)로 전송 — no-cors(fire-and-forget) */
async function pushToSheet(daysObj) {
  const cfg = loadGSheet();
  if (!cfg || !cfg.url) return { ok: false, reason: "미연결" };
  const payload = { token: cfg.token || "", days: {} };
  for (const [date, res] of Object.entries(daysObj)) {
    const rows = (res.rows || []);
    payload.days[date] = { summary: summaryRow(date, rows), rows: rows.map(detailRow) };
  }
  try {
    await fetch(cfg.url, {
      method: "POST", mode: "no-cors",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
    });
    return { ok: true, count: Object.keys(payload.days).length };
  } catch (e) {
    return { ok: false, reason: e.message };
  }
}

/* ---------- UI ---------- */
let staged = []; // {file, wb, kind}
let marginState = loadMargin(); // {map, meta}

function renderMarginBadge() {
  const badge = $("#marginBadge"), info = $("#marginInfo");
  if (marginState && marginState.map) {
    const n = Object.keys(marginState.map).length;
    badge.textContent = `기억됨 · ${n}개 상품`;
    badge.className = "badge ok";
    info.textContent = marginState.meta ? `파일: ${marginState.meta.name}` : "";
  } else {
    badge.textContent = "없음";
    badge.className = "badge";
    info.textContent = "마진표를 아직 올리지 않았습니다.";
  }
}

async function readWorkbook(file) {
  const buf = await file.arrayBuffer();
  return XLSX.read(buf, { type: "array" });
}

async function handleMarginFile(file) {
  try {
    const wb = await readWorkbook(file);
    const map = parseMargin(wb);
    if (!Object.keys(map).length) { alert("마진이 기입된 행을 찾지 못했습니다."); return; }
    marginState = { map, meta: { name: file.name } };
    saveMargin(map, marginState.meta);
    renderMarginBadge();
  } catch (e) { alert("마진표 처리 오류: " + e.message); }
}

async function stageDaily(files) {
  for (const file of files) {
    try {
      const wb = await readWorkbook(file);
      let kind = classify(wb);
      if (kind === "margin") { await handleMarginFile(file); continue; }
      if (kind === "result") { importResult(wb, file.name); continue; }
      staged.push({ file, wb, kind });
    } catch (e) { alert(`${file.name} 읽기 오류: ${e.message}`); }
  }
  renderStaged();
}

/* 결과 엑셀을 누적 데이터에 병합 */
function importResult(wb, fname) {
  let imported;
  try { imported = parseResult(wb); }
  catch (e) { alert("불러오기 오류: " + e.message); return; }
  const dates = Object.keys(imported);
  if (!dates.length) { alert(`${fname}: 불러올 날짜 데이터가 없습니다.`); return; }
  const days = loadDays();
  for (const d of dates) days[d] = imported[d]; // 같은 날짜는 파일 내용으로 갱신
  saveDays(days);
  renderResults();
  const sorted = dates.slice().sort();
  alert(`이전 누적본에서 ${dates.length}일치를 불러왔습니다. (${sorted[0]} ~ ${sorted[sorted.length - 1]})`);
}

function renderStaged() {
  const box = $("#stagedFiles");
  box.innerHTML = "";
  const tagText = { insight: "인사이트", ad: "광고", unknown: "인식 실패" };
  staged.forEach((s, i) => {
    const div = document.createElement("div");
    div.className = "item";
    div.innerHTML = `<span class="tag ${s.kind}">${tagText[s.kind]}</span>
      <span>${s.file.name}</span><span class="x" data-i="${i}">✕</span>`;
    box.appendChild(div);
  });
  box.querySelectorAll(".x").forEach((x) =>
    x.addEventListener("click", () => { staged.splice(+x.dataset.i, 1); renderStaged(); }));

  const ins = staged.find((s) => s.kind === "insight");
  const ad = staged.find((s) => s.kind === "ad");
  const dateEl = $("#dateLabel");
  let date = ad ? dateFromName(ad.file.name) : null;
  if (date) { dateEl.hidden = false; dateEl.textContent = `날짜: ${date}`; }
  else dateEl.hidden = true;

  $("#calcBtn").disabled = !(ins && ad && marginState && marginState.map);
}

function calculate() {
  const insFile = staged.find((s) => s.kind === "insight");
  const adFile = staged.find((s) => s.kind === "ad");
  if (!insFile || !adFile) { alert("인사이트와 광고 리포트 파일이 모두 필요합니다."); return; }
  if (!marginState || !marginState.map) { alert("먼저 마진표를 올려주세요."); return; }

  let date = dateFromName(adFile.file.name) || dateFromName(insFile.file.name) ||
    prompt("파일명에서 날짜를 찾지 못했습니다. 날짜를 입력하세요 (YYYY-MM-DD):", "");
  if (!date) return;

  let insight, adByOpt;
  try { insight = parseInsight(insFile.wb); adByOpt = parseAd(adFile.wb); }
  catch (e) { alert("계산 오류: " + e.message); return; }

  const res = computeDay(date, marginState.map, insight, adByOpt);

  // 누적 저장
  const days = loadDays();
  days[date] = res;
  saveDays(days);

  staged = [];
  renderStaged();
  renderResults();

  // 구글시트 자동 저장
  const cfg = loadGSheet();
  if (cfg && cfg.url && cfg.auto) {
    setGMsg(`구글시트로 ${date} 전송 중...`);
    pushToSheet({ [date]: res }).then((r) => {
      setGMsg(r.ok ? `구글시트에 ${date} 저장 요청 완료. 시트에서 확인하세요.` : `구글시트 전송 실패: ${r.reason}`);
    });
  }
}
function setGMsg(t) { const el = $("#gsheetMsg"); if (el) el.textContent = t; }

function renderResults() {
  const days = loadDays();
  const dates = Object.keys(days).sort();
  if (!dates.length) { $("#resultCard").hidden = true; return; }
  $("#resultCard").hidden = false;

  // 누적 현황
  const acc = $("#accStatus");
  if (acc) {
    acc.innerHTML = `📅 현재 누적 <b>${dates.length}일</b> · ${dates[0]} ~ ${dates[dates.length - 1]}
      <span class="acc-hint">— 매일 파일을 넣고 계산하면 자동으로 쌓입니다. 다운로드한 엑셀을 다시 드롭존에 넣으면 다른 PC에서도 이어서 기록됩니다.</span>`;
  }

  // 모든 행 합치기
  let allRows = [];
  let soldNoMargin = [];
  let unmappedAd = [];
  for (const d of dates) {
    allRows = allRows.concat(days[d].rows || []);
    (days[d].soldNoMargin || []).forEach((x) => soldNoMargin.push({ ...x, date: d }));
    (days[d].unmappedAd || []).forEach((x) => unmappedAd.push({ ...x, date: d }));
  }
  allRows.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : b.settleRev - a.settleRev));

  // 경고
  const warn = $("#warnings");
  warn.innerHTML = "";
  if (soldNoMargin.length) {
    const items = soldNoMargin.slice(0, 30).map((x) =>
      `<li>[${x.date}] ${x.name} (ID ${x.regId}) · ${x.qty}개 · 매출 ${won(x.rev)}원</li>`).join("");
    warn.innerHTML += `<div class="warn-box warn"><h4>⚠️ 판매됐지만 마진 미기입 상품 ${soldNoMargin.length}건 (마진표에 추가하면 이익 계산에 포함됩니다)</h4><ul>${items}</ul></div>`;
  }
  if (unmappedAd.length) {
    const tot = unmappedAd.reduce((s, x) => s + x.cost, 0);
    const items = unmappedAd.slice(0, 20).map((x) =>
      `<li>[${x.date}] 광고옵션ID ${x.opt} · 광고비 ${won(x.cost)}원</li>`).join("");
    warn.innerHTML += `<div class="warn-box warn"><h4>⚠️ 인사이트에 없어 매칭되지 않은 광고비 ${won(tot)}원</h4><ul>${items}</ul></div>`;
  }

  // 테이블
  const cols = [
    ["날짜", "date"], ["상품구분", "group"], ["상품명", "name"], ["등록상품ID", "regId"],
    ["판매수량", "qty"], ["구매당비용", "costPerUnit"], ["개당마진", "marginUnit"],
    ["인사이트매출", "insightRev"], ["정산매출", "settleRev"],
    ["사용광고비", "adCost"], ["마진(광고전)", "marginBefore"], ["순이익", "netProfit"], ["순이익률", "netRate"],
  ];
  const t = $("#resultTable");
  const thead = `<thead><tr>${cols.map((c) =>
    `<th title="${COLUMN_DEFS[c[1]] || ""}">${c[0]}<span class="qmark">?</span></th>`).join("")}</tr></thead>`;
  const money = new Set(["insightRev", "settleRev", "adCost", "marginBefore", "netProfit"]);
  const body = allRows.map((r) => {
    const tds = cols.map(([, k]) => {
      if (k === "netRate") return `<td>${pct(r.netRate)}</td>`;
      if (k === "qty") return `<td>${won(r.qty)}</td>`;
      if (k === "costPerUnit") {
        if (!(r.qty > 0)) return `<td class="muted">-</td>`;
        const cpu = r.adCost / r.qty, mu = muOf(r);
        const cls = isFinite(mu) ? (cpu >= mu ? "neg" : "pos") : "";
        return `<td class="${cls}">${won(cpu)}</td>`;
      }
      if (k === "marginUnit") return `<td>${isFinite(muOf(r)) ? won(muOf(r)) : "-"}</td>`;
      if (money.has(k)) {
        const cls = k === "netProfit" ? (r.netProfit >= 0 ? "pos" : "neg") : "";
        return `<td class="${cls}">${won(r[k])}</td>`;
      }
      return `<td>${r[k] == null ? "" : r[k]}</td>`;
    });
    return `<tr>${tds.join("")}</tr>`;
  }).join("");

  // 합계
  const sum = (k) => allRows.reduce((s, r) => s + (r[k] || 0), 0);
  const totNet = sum("netProfit"), totSettle = sum("settleRev");
  const totQty = sum("qty");
  const totalCells = cols.map(([, k]) => {
    if (k === "date") return `<td>합계 (${dates.length}일)</td>`;
    if (k === "group" || k === "name" || k === "regId") return `<td></td>`;
    if (k === "netRate") return `<td>${pct(totSettle > 0 ? totNet / totSettle : NaN)}</td>`;
    if (k === "qty") return `<td>${won(totQty)}</td>`;
    if (k === "costPerUnit") {
      if (!(totQty > 0)) return `<td class="muted">-</td>`;
      const cpu = sum("adCost") / totQty, mu = sum("marginBefore") / totQty;
      const cls = cpu >= mu ? "neg" : "pos";
      return `<td class="${cls}">${won(cpu)}</td>`;
    }
    if (k === "marginUnit") return `<td>${totQty > 0 ? won(sum("marginBefore") / totQty) : "-"}</td>`;
    const cls = k === "netProfit" ? (totNet >= 0 ? "pos" : "neg") : "";
    return `<td class="${cls}">${won(sum(k))}</td>`;
  }).join("");
  const foot = `<tr class="total">${totalCells}</tr>`;

  t.innerHTML = thead + `<tbody>${body}${foot}</tbody>`;
}

/* ---------- 엑셀 다운로드 ---------- */
function downloadExcel() {
  const days = loadDays();
  const dates = Object.keys(days).sort();
  if (!dates.length) return;
  let rows = [];
  for (const d of dates) rows = rows.concat(days[d].rows || []);
  rows.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : b.settleRev - a.settleRev));

  const header = ["날짜", "상품구분", "상품명", "등록상품ID", "판매수량", "구매당비용", "개당마진",
    "인사이트매출", "정산매출", "사용광고비", "마진(광고전)", "순이익", "순이익률"];
  const aoa = [header];
  for (const r of rows) {
    const cpu = r.qty > 0 ? Math.round(r.adCost / r.qty) : "";
    const mu = isFinite(muOf(r)) ? Math.round(muOf(r)) : "";
    aoa.push([r.date, r.group, r.name, r.regId, r.qty, cpu, mu,
      Math.round(r.insightRev), Math.round(r.settleRev), Math.round(r.adCost),
      Math.round(r.marginBefore), Math.round(r.netProfit),
      isFinite(r.netRate) ? +(r.netRate).toFixed(4) : ""]);
  }
  const sum = (k) => rows.reduce((s, r) => s + (r[k] || 0), 0);
  const tNet = sum("netProfit"), tSettle = sum("settleRev"), tQty = sum("qty");
  aoa.push(["합계", "", "", "", tQty,
    tQty > 0 ? Math.round(sum("adCost") / tQty) : "", tQty > 0 ? Math.round(sum("marginBefore") / tQty) : "",
    Math.round(sum("insightRev")), Math.round(tSettle), Math.round(sum("adCost")),
    Math.round(sum("marginBefore")), Math.round(tNet),
    tSettle > 0 ? +(tNet / tSettle).toFixed(4) : ""]);

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 11 }, { wch: 12 }, { wch: 34 }, { wch: 13 }, { wch: 9 }, { wch: 11 }, { wch: 11 },
    { wch: 13 }, { wch: 13 }, { wch: 12 }, { wch: 13 }, { wch: 13 }, { wch: 9 }];

  // 헤더 셀(1행)에 정의·계산식 메모 삽입 — 엑셀에서 셀에 마우스 올리면 표시
  const keysByCol = ["date", "group", "name", "regId", "qty", "costPerUnit", "marginUnit",
    "insightRev", "settleRev", "adCost", "marginBefore", "netProfit", "netRate"];
  keysByCol.forEach((k, c) => {
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    if (!ws[addr]) return;
    ws[addr].c = [{ a: "계산기", t: COLUMN_DEFS[k] || "" }];
    ws[addr].c.hidden = true;
  });
  // 숫자 서식 (구매당비용~순이익 = 열 5~11, 순이익률 = 열 12)
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let R = 1; R <= range.e.r; R++) {
    for (let C = 5; C <= 11; C++) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      if (cell && typeof cell.v === "number") cell.z = "#,##0";
    }
    const pc = ws[XLSX.utils.encode_cell({ r: R, c: 12 })];
    if (pc && typeof pc.v === "number") pc.z = "0.0%";
  }
  // 일별요약 시트
  const sHeader = ["날짜", "판매수량", "인사이트매출", "정산매출", "사용광고비", "마진(광고전)", "순이익", "순이익률"];
  const sAoa = [sHeader];
  for (const d of dates) sAoa.push(summaryRow(d, days[d].rows || []));
  const ssum = (i) => dates.reduce((a, d) => a + (summaryRow(d, days[d].rows || [])[i] || 0), 0);
  const stotSettle = ssum(3), stotNet = ssum(6);
  sAoa.push(["합계", ssum(1), ssum(2), stotSettle, ssum(4), ssum(5), stotNet,
    stotSettle > 0 ? +(stotNet / stotSettle).toFixed(4) : ""]);
  const wsSum = XLSX.utils.aoa_to_sheet(sAoa);
  wsSum["!cols"] = [{ wch: 12 }, { wch: 9 }, { wch: 13 }, { wch: 13 }, { wch: 12 }, { wch: 13 }, { wch: 13 }, { wch: 9 }];
  const sRange = XLSX.utils.decode_range(wsSum["!ref"]);
  for (let R = 1; R <= sRange.e.r; R++) {
    for (let C = 1; C <= 6; C++) {
      const cell = wsSum[XLSX.utils.encode_cell({ r: R, c: C })];
      if (cell && typeof cell.v === "number") cell.z = "#,##0";
    }
    const pc = wsSum[XLSX.utils.encode_cell({ r: R, c: 7 })];
    if (pc && typeof pc.v === "number") pc.z = "0.0%";
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsSum, "일별요약");
  XLSX.utils.book_append_sheet(wb, ws, "상세");
  const fname = dates.length === 1 ? `쿠팡이익_${dates[0]}.xlsx` : `쿠팡이익_${dates[0]}_${dates[dates.length - 1]}.xlsx`;
  XLSX.writeFile(wb, fname);
}

/* ---------- 이벤트 바인딩 ---------- */
function renderLegend() {
  const labels = {
    date: "날짜", group: "상품구분", name: "상품명", regId: "등록상품ID", qty: "판매수량",
    costPerUnit: "구매당비용", marginUnit: "개당마진",
    insightRev: "인사이트매출", settleRev: "정산매출", adCost: "사용광고비",
    marginBefore: "마진(광고전)", netProfit: "순이익", netRate: "순이익률",
  };
  const body = $("#legendBody");
  body.innerHTML = Object.keys(labels).map((k) => {
    const lines = COLUMN_DEFS[k].split("\n");
    const detail = lines.slice(1).map((l) => l.replace(/^·\s*/, "")).filter(Boolean)
      .map((l) => `<div>${l}</div>`).join("");
    return `<div class="legend-item"><b>${labels[k]}</b><div class="legend-detail">${detail}</div></div>`;
  }).join("");
}

function renderGSheet() {
  const cfg = loadGSheet();
  const badge = $("#gsheetBadge");
  if (cfg && cfg.url) {
    $("#gsheetUrl").value = cfg.url;
    $("#gsheetToken").value = cfg.token || "";
    $("#gsheetAuto").checked = cfg.auto !== false;
    badge.textContent = "연결됨"; badge.className = "badge ok";
  } else {
    badge.textContent = "미연결"; badge.className = "badge";
  }
}

function init() {
  renderMarginBadge();
  renderGSheet();
  renderLegend();
  renderResults();

  $("#gsheetSave").addEventListener("click", () => {
    const url = $("#gsheetUrl").value.trim();
    if (!/^https:\/\/script\.google\.com\/.*\/exec$/.test(url)) {
      alert("구글시트 웹 앱 URL 형식이 아닙니다. (https://script.google.com/.../exec)"); return;
    }
    saveGSheet({ url, token: $("#gsheetToken").value.trim(), auto: $("#gsheetAuto").checked });
    renderGSheet(); setGMsg("연결 정보를 저장했습니다.");
  });
  $("#gsheetAuto").addEventListener("change", () => {
    const cfg = loadGSheet(); if (cfg) { cfg.auto = $("#gsheetAuto").checked; saveGSheet(cfg); }
  });
  $("#gsheetPushAll").addEventListener("click", async () => {
    const days = loadDays();
    if (!Object.keys(days).length) { setGMsg("보낼 누적 데이터가 없습니다."); return; }
    if (!loadGSheet()?.url) { alert("먼저 구글시트 주소를 저장하세요."); return; }
    setGMsg("구글시트로 전체 전송 중...");
    const r = await pushToSheet(days);
    setGMsg(r.ok ? `${r.count}일치 전송 요청 완료. 구글시트에서 확인하세요.` : `전송 실패: ${r.reason}`);
  });

  $("#marginInput").addEventListener("change", (e) => {
    if (e.target.files[0]) handleMarginFile(e.target.files[0]);
    e.target.value = "";
  });
  $("#marginClear").addEventListener("click", () => {
    if (!confirm("기억된 마진표를 지울까요?")) return;
    localStorage.removeItem(LS_MARGIN); marginState = null; renderMarginBadge(); renderStaged();
  });

  const dz = $("#dropzone");
  ["dragenter", "dragover"].forEach((ev) => dz.addEventListener(ev, (e) => {
    e.preventDefault(); dz.classList.add("drag");
  }));
  ["dragleave", "drop"].forEach((ev) => dz.addEventListener(ev, (e) => {
    e.preventDefault(); dz.classList.remove("drag");
  }));
  dz.addEventListener("drop", (e) => {
    const files = [...(e.dataTransfer?.files || [])].filter((f) => /\.xlsx?$/i.test(f.name));
    if (files.length) stageDaily(files);
  });
  $("#dailyInput").addEventListener("change", (e) => {
    if (e.target.files.length) stageDaily([...e.target.files]);
    e.target.value = "";
  });

  $("#calcBtn").addEventListener("click", calculate);
  $("#downloadBtn").addEventListener("click", downloadExcel);
  $("#clearAllBtn").addEventListener("click", () => {
    if (!confirm("누적된 모든 날짜 결과를 지울까요? (마진표는 유지됩니다)")) return;
    localStorage.removeItem(LS_DAYS); renderResults();
  });
}
init();
