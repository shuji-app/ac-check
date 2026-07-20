// ブラウザ上でBabel事前コンパイル済みJSを実行するため、ES importではなくグローバルReactから取り出す
const {
  useState,
  useRef,
  useEffect,
  memo
} = React;
const INDOOR_FIELDS = [{
  code: "b1",
  label: "吸込温度",
  unit: "°C",
  step: 0.1,
  group: "indoor"
}, {
  code: "b2",
  label: "吹出温度",
  unit: "°C",
  step: 0.1,
  group: "indoor"
}, {
  code: "b3",
  label: "室内熱交・入口",
  unit: "°C",
  step: 0.1,
  group: "indoor"
}, {
  code: "b4",
  label: "室内熱交・中間",
  unit: "°C",
  step: 0.1,
  group: "indoor"
}, {
  code: "b5",
  label: "室内熱交・出口",
  unit: "°C",
  step: 0.1,
  group: "indoor"
}, {
  code: "b6",
  label: "膨張弁開度",
  unit: "pulse",
  step: 1,
  group: "indoor"
}, {
  code: "b8",
  label: "リモコン感知",
  unit: "°C",
  step: 0.1,
  group: "indoor"
}];
// 「エアコン定期点検.xls」の室外機シート「運転調整・データ採取」欄と同じ6項目・同じ順序（v9.28）
const OUTDOOR_FIELDS = [{
  code: "od1",
  label: "圧縮機電流",
  unit: "A",
  step: 0.1,
  group: "outdoor"
}, {
  code: "od2",
  label: "吐出ガス圧力",
  unit: "MPa",
  step: 0.01,
  group: "outdoor"
}, {
  code: "od3",
  label: "吐出ガス温度",
  unit: "°C",
  step: 0.1,
  group: "outdoor"
}, {
  code: "od4",
  label: "吸入ガス圧力",
  unit: "MPa",
  step: 0.01,
  group: "outdoor"
}, {
  code: "od5",
  label: "吸入ガス温度",
  unit: "°C",
  step: 0.1,
  group: "outdoor"
}, {
  code: "od6",
  label: "冷媒液温度",
  unit: "°C",
  step: 0.1,
  group: "outdoor"
}];
const ALL_FIELDS = [...INDOOR_FIELDS, ...OUTDOOR_FIELDS];
// 正常値範囲設定画面：テンキーで最小→最大→次の項目…と連続入力するための順序（点検データ入力画面のfocusSeqと同じ考え方）
const LIM_SEQ = ALL_FIELDS.flatMap(f => [{
  code: f.code,
  part: "min"
}, {
  code: f.code,
  part: "max"
}]);
const defVis = () => {
  const v = {};
  ALL_FIELDS.forEach(f => v[f.code] = true);
  return v;
};
const defLim = () => {
  const v = {};
  ALL_FIELDS.forEach(f => v[f.code] = {
    enabled: false,
    min: "",
    max: ""
  });
  return v;
};
const emptyVal = () => {
  const v = {};
  ALL_FIELDS.forEach(f => v[f.code] = "");
  return v;
};
const emptyForm = (inspector = "", date = "") => ({
  id: crypto.randomUUID(),
  // 保存前でもボイスメモ等を紐づけられるよう、フォーム作成時点で確定させる
  floor: "",
  room: "",
  managementNo: "",
  unitNo: "",
  inspectionDate: date || new Date().toISOString().slice(0, 10),
  inspector,
  preOperation: "",
  preMode: "",
  preWind: "",
  preSetTemp: "",
  values: emptyVal(),
  checks: {},
  remarks: ""
});

// フォーカス移動時の自動スクロールを、ブラウザ標準のsmoothスクロールに頼らず
// 自前でアニメーションさせるヘルパー（対象行の縦幅が変化してもスクロール先を
// 都度計算し直さず、開始時点のズレ量だけを一定時間でなめらかに詰めるため、
// ちらつき無く一定のスライドに見える）
function smoothScrollTo(container, targetTop, duration = 260) {
  if (!container) return;
  const startTop = container.scrollTop;
  const delta = targetTop - startTop;
  if (Math.abs(delta) < 1) return;
  const startTime = performance.now();
  const ease = t => 1 - Math.pow(1 - t, 3); // ease-out
  const step = now => {
    const t = Math.min(1, (now - startTime) / duration);
    container.scrollTop = startTop + delta * ease(t);
    if (t < 1) requestAnimationFrame(step);
  };
  requestAnimationFrame(step);
}
function isAbn(code, val, limits) {
  const v = parseFloat(val);
  if (val === "" || isNaN(v)) return false;
  const l = limits[code];
  if (!l || !l.enabled) return false;
  if (l.min !== "" && !isNaN(parseFloat(l.min)) && v < parseFloat(l.min)) return true;
  if (l.max !== "" && !isNaN(parseFloat(l.max)) && v > parseFloat(l.max)) return true;
  return false;
}
// 列名候補からキーを探す（大文字小文字・全半角を無視）
function findColKey(keys, ...candidates) {
  const norm = s => s.trim().replace(/\s/g, "").toLowerCase();
  for (const cand of candidates) {
    const found = keys.find(k => norm(k) === norm(cand));
    if (found) return found;
  }
  return null;
}
// _raw（列名→値）から floor/room/managementNo/unitNo を再計算する（parseDevRowsと同じ列検出ロジック）。
// 設定画面での機器の個別編集・新規追加で、_rawを直接編集した後にこの4項目を同期するために使う。
function deriveDevCore(raw, keys) {
  const fk = findColKey(keys, "階", "floor", "フロア", "エリア", "area");
  const rk = findColKey(keys, "部屋名", "room", "部屋", "室名");
  const mk = findColKey(keys, "管理番号", "managementno", "管理No", "管理no", "管理");
  const uk = findColKey(keys, "機器番号", "unitno", "機器No", "機器no", "機器", "unit");
  const fk2 = fk || keys[0] || null,
    rk2 = rk || keys[1] || null;
  const mk2 = mk || keys[2] || null,
    uk2 = uk || keys[3] || null;
  return {
    floor: String(fk2 && raw[fk2] || "").trim(),
    room: String(rk2 && raw[rk2] || "").trim(),
    managementNo: String(mk2 && raw[mk2] || "").trim(),
    unitNo: String(uk2 && raw[uk2] || "").trim()
  };
}
// JSON行配列（xlsxのsheet_to_json結果）をdevListに変換
function parseDevRows(rows, origKeys) {
  if (!rows || rows.length === 0) return [];
  const keys = origKeys || Object.keys(rows[0]);
  const fk = findColKey(keys, "階", "floor", "フロア", "エリア", "area");
  const rk = findColKey(keys, "部屋名", "room", "部屋", "室名");
  const mk = findColKey(keys, "管理番号", "managementno", "管理No", "管理no", "管理");
  const uk = findColKey(keys, "機器番号", "unitno", "機器No", "機器no", "機器", "unit");
  // 見つからない場合は列順で対応
  const fk2 = fk || keys[0] || null,
    rk2 = rk || keys[1] || null;
  const mk2 = mk || keys[2] || null,
    uk2 = uk || keys[3] || null;
  return rows.map(row => {
    const raw = {};
    keys.forEach(k => {
      raw[k] = String(row[k] || "");
    });
    return {
      floor: String(row[fk2] || "").trim(),
      room: String(row[rk2] || "").trim(),
      managementNo: String(row[mk2] || "").trim(),
      unitNo: String(row[uk2] || "").trim(),
      _raw: raw
    };
  }).filter(r => r.managementNo || r.unitNo);
}
// CSV文字列からdevListに変換
function parseDevCSV(text) {
  const lines = text.trim().split(/\r?\n/).filter(l => l.trim());
  if (lines.length < 2) return [];
  const origKeys = lines[0].split(",").map(s => s.trim().replace(/^"|"$/g, ""));
  const rows = lines.slice(1).map(line => {
    const c = line.split(",").map(s => s.trim().replace(/^"|"$/g, ""));
    const obj = {};
    origKeys.forEach((k, i) => {
      obj[k] = c[i] || "";
    });
    return obj;
  });
  return parseDevRows(rows, origKeys);
}
function parseInspCSV(text) {
  return text.trim().split(/\r?\n/).map(l => l.trim().replace(/^"|"$/g, "")).filter(Boolean);
}
function doExport(records, visibility, checkFields = []) {
  const vf = ALL_FIELDS.filter(f => visibility[f.code]);
  const rows = [["点検日", "点検者", "階", "部屋名", "管理番号", "機器番号", "運転", "モード", "風量", "設定温度", ...vf.map(f => f.code + "(" + f.unit + ")"), ...checkFields.map(f => f.label), "備考"], ...records.map(r => [r.inspectionDate, r.inspector, r.floor, r.room, r.managementNo, r.unitNo, r.preOperation || "", r.preMode || "", r.preWind || "", r.preSetTemp || "", ...vf.map(f => r.values[f.code]), ...checkFields.map(f => r.checks?.[f.code] || ""), r.remarks])];
  const tsv = rows.map(r => r.map(c => '"' + String(c ?? "").replace(/"/g, '""') + '"').join("\t")).join("\n");
  const blob = new Blob(["\uFEFF" + tsv], {
    type: "text/tab-separated-values;charset=utf-8;"
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "ac_check_" + new Date().toISOString().slice(0, 10) + ".tsv";
  a.click();
}
const PS = "@media print{body>*{display:none!important;}#print-area{display:block!important;position:static!important;}@page{size:A4 landscape;margin:10mm;}}";
// ─── ブラウザ保存ヘルパー（機器リスト等の「取込データ」永続化用）───
// アプリのコードを更新（再読込）しても、設定画面で読み込んだデータが消えないようにする。
// 対象：機器リスト・表示列設定・点検項目・点検者リスト・表示項目設定・正常値範囲
// 対象外：入力中の点検記録（indoorRecords/outdoorRecords）・セッション情報は従来通り
function lsGet(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw !== null ? JSON.parse(raw) : fallback;
  } catch (e) {
    return fallback;
  }
}
function lsSet(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch (e) {}
}
const LS_KEYS = {
  devList: "acDevList",
  devColumns: "acDevColumns",
  devVisibleCols: "acDevVisibleCols",
  inspList: "acInspList",
  checkFields: "acCheckFields",
  limits: "acLimits",
  vis: "acVis",
  cardLabels: "acCardLabels"
};
const defCardLabels = () => ({
  device: "機器リスト",
  inspector: "点検者リスト",
  criteria: "点検基準設定"
});

// ─── Firebase（複数端末間のリアルタイム共有）───
// 使い方：Firebaseコンソール（console.firebase.google.com）でプロジェクトを作成し、
// Realtime Databaseを有効化した上で、下記の値を自分のプロジェクトのものに置き換える。
// CDN読み込みに失敗した場合（オフライン等）はdb=nullとなり、以降の同期処理は
// すべて安全にスキップされ、アプリ自体はローカルのみで動作し続ける。
const firebaseConfig = {
  apiKey: "AIzaSyDu_e2tLEVu1aO2STneS6tNVNl9Wr3jjXs",
  authDomain: "ac-check-app.firebaseapp.com",
  databaseURL: "https://ac-check-app-default-rtdb.firebaseio.com",
  projectId: "ac-check-app",
  storageBucket: "ac-check-app.firebasestorage.app",
  messagingSenderId: "1089584979639",
  appId: "1:1089584979639:web:9ecd774294b912d6fbb732"
};
const fbApp = typeof window !== "undefined" && window.firebase && firebaseConfig.apiKey !== "YOUR_API_KEY" ? window.firebase.initializeApp(firebaseConfig) : null;
const db = fbApp ? window.firebase.database() : null;
function saveRecordRemote(roundId, mode, rec) {
  if (!db || !roundId || !rec.id) return;
  const {
    id,
    ...rest
  } = rec;
  db.ref("rounds/" + roundId + "/" + mode + "/" + id).set({
    ...rest,
    updatedAt: window.firebase.database.ServerValue.TIMESTAMP
  });
}
function deleteRecordRemote(roundId, mode, id) {
  if (!db || !roundId || !id) return;
  db.ref("rounds/" + roundId + "/" + mode + "/" + id).remove();
}

// ─── ボイスメモ（iPadのみ・端末内保存）───
// iPadOS13以降はSafari（および中身がSafariのChrome等）がUser-Agentを"Macintosh"と
// 偽装するため、UAでの判定は効かない。タッチ対応の有無で見分ける。
function isIPad() {
  if (typeof navigator === "undefined") return false;
  if (/iPad/.test(navigator.userAgent)) return true;
  return navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1;
}
const IS_IPAD = isIPad();
function pickMimeType() {
  const cands = ["audio/mp4", "audio/webm;codecs=opus", "audio/webm", "audio/ogg"];
  for (const c of cands) {
    if (window.MediaRecorder && MediaRecorder.isTypeSupported(c)) return c;
  }
  return "";
}
function vmOpenDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open("acVoiceMemos", 1);
    req.onupgradeneeded = () => {
      req.result.createObjectStore("memos", {
        keyPath: "id"
      });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
async function vmAdd(memo) {
  const dbx = await vmOpenDB();
  return new Promise((resolve, reject) => {
    const tx = dbx.transaction("memos", "readwrite");
    tx.objectStore("memos").add(memo);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}
async function vmDelete(id) {
  const dbx = await vmOpenDB();
  return new Promise((resolve, reject) => {
    const tx = dbx.transaction("memos", "readwrite");
    tx.objectStore("memos").delete(id);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}
async function vmListByScope(scope, recordId) {
  const dbx = await vmOpenDB();
  return new Promise((resolve, reject) => {
    const tx = dbx.transaction("memos", "readonly");
    const req = tx.objectStore("memos").getAll();
    req.onsuccess = () => {
      const all = req.result || [];
      const filtered = all.filter(m => m.scope === scope && (scope !== "record" || m.recordId === recordId));
      filtered.sort((a, b) => b.ts - a.ts);
      resolve(filtered);
    };
    req.onerror = () => reject(req.error);
  });
}
const C = {
  navy: "#1B3A6B",
  blue: "#2563B0",
  teal: "#0D7A6B",
  green: "#059669",
  red: "#DC2626",
  purple: "#7C3AED",
  g50: "#F8FAFC",
  g100: "#F1F5F9",
  g200: "#E2E8F0",
  g300: "#CBD5E1",
  g400: "#94A3B8",
  g500: "#64748B",
  g600: "#475569",
  g800: "#1E293B",
  white: "#FFFFFF",
  inp: "#F7F9FC"
};
const FieldRow = memo(function FR({
  f,
  isIn,
  idx,
  active,
  val,
  abn,
  dimmed,
  onClick,
  fRef
}) {
  const fill = val !== "";
  return /*#__PURE__*/React.createElement("tr", {
    ref: fRef,
    onClick: onClick,
    style: {
      background: active ? "linear-gradient(90deg,#2563B022,#2563B00A)" : abn ? "#FEF2F2" : dimmed ? "#F5F7FA" : idx % 2 === 0 ? C.white : C.g50,
      cursor: "pointer",
      outline: active ? "2px solid " + C.blue : "none",
      outlineOffset: -1,
      opacity: dimmed ? 0.45 : 1
    }
  }, /*#__PURE__*/React.createElement("td", {
    style: {
      width: 5,
      padding: 0,
      background: active ? C.blue : abn ? C.red : "transparent",
      borderBottom: "1px solid " + C.g100
    }
  }), /*#__PURE__*/React.createElement("td", {
    style: {
      padding: "9px 8px",
      fontFamily: "monospace",
      fontWeight: 700,
      fontSize: 13,
      color: active ? C.blue : isIn ? C.blue : C.teal,
      borderBottom: "1px solid " + C.g100,
      whiteSpace: "nowrap"
    }
  }, f.code), /*#__PURE__*/React.createElement("td", {
    style: {
      padding: "9px 8px",
      fontSize: 13,
      color: active ? C.navy : abn ? C.red : C.g600,
      fontWeight: active ? 700 : 400,
      borderBottom: "1px solid " + C.g100
    }
  }, f.label, abn && /*#__PURE__*/React.createElement("span", {
    style: {
      marginLeft: 4,
      fontSize: 9,
      background: C.red,
      color: C.white,
      padding: "1px 4px",
      borderRadius: 3,
      fontWeight: 700
    }
  }, "\u26A0\uFE0F")), /*#__PURE__*/React.createElement("td", {
    style: {
      padding: "9px 6px",
      fontSize: 11,
      color: C.g400,
      borderBottom: "1px solid " + C.g100,
      textAlign: "center",
      whiteSpace: "nowrap"
    }
  }, f.unit), /*#__PURE__*/React.createElement("td", {
    style: {
      padding: "4px 7px",
      borderBottom: "1px solid " + C.g100,
      width: 110
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "6px 10px",
      borderRadius: 8,
      fontSize: 14,
      fontFamily: "monospace",
      textAlign: "right",
      fontWeight: 800,
      border: "2px solid " + (active ? C.blue : fill ? C.green : C.g200),
      background: active ? "#EFF6FF" : fill ? "#F0FDF4" : C.white,
      color: abn ? C.red : fill ? C.g800 : C.g300,
      minHeight: 32,
      display: "flex",
      alignItems: "center",
      justifyContent: "flex-end"
    }
  }, fill ? val : active ? /*#__PURE__*/React.createElement("span", {
    style: {
      color: C.blue + "80",
      fontSize: 18
    }
  }, "\u2014") : "—")));
});
function Numpad({
  mode = "numeric",
  display,
  onPress,
  onConfirm,
  canConfirm,
  checkLabel,
  checkCategory,
  checkValue,
  onCheckPress,
  onPrev,
  onNext,
  canPrev,
  canNext,
  onSave,
  saveComplete,
  saveMissing
}) {
  const isCheck = mode === "check";
  const KEYS = [[7, 8, 9], [4, 5, 6], [1, 2, 3], [0, "."]];
  const kb = en => ({
    flex: 1,
    padding: 0,
    height: 56,
    borderRadius: 12,
    border: "none",
    cursor: en ? "pointer" : "not-allowed",
    fontWeight: 800,
    fontFamily: "monospace",
    fontSize: 22,
    background: en ? C.white : C.g100,
    color: en ? C.g800 : C.g300,
    boxShadow: en ? "0 2px 6px rgba(0,0,0,0.10)" : "none",
    display: "flex",
    alignItems: "center",
    justifyContent: "center"
  });
  const nb = en => ({
    flex: 1,
    height: 52,
    borderRadius: 12,
    border: "2px solid " + (en ? C.blue : C.g200),
    cursor: en ? "pointer" : "not-allowed",
    fontSize: 22,
    fontWeight: 800,
    background: en ? C.blue + "15" : C.g50,
    color: en ? C.blue : C.g300,
    display: "flex",
    alignItems: "center",
    justifyContent: "center"
  });
  return /*#__PURE__*/React.createElement("div", {
    style: {
      width: 208,
      flexShrink: 0,
      display: "flex",
      flexDirection: "column",
      background: C.g50,
      borderLeft: "2px solid " + C.g200,
      padding: "8px 8px 10px",
      gap: 6
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: C.white,
      borderRadius: 12,
      padding: "8px 12px",
      border: "2px solid " + (isCheck ? C.green : C.blue),
      height: 58,
      boxSizing: "border-box",
      display: "flex",
      flexDirection: "column",
      justifyContent: "center"
    }
  }, isCheck ? /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 10,
      fontWeight: 700,
      color: "#059669",
      letterSpacing: "0.04em",
      minHeight: 14
    }
  }, checkCategory || ""), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 16,
      fontWeight: 800,
      color: C.navy,
      lineHeight: 1.3,
      whiteSpace: "nowrap",
      overflow: "hidden",
      textOverflow: "ellipsis"
    }
  }, checkLabel || "—")) : /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 10,
      fontWeight: 700,
      color: C.g400,
      letterSpacing: "0.04em"
    }
  }, "\u5165\u529B\u5024"), /*#__PURE__*/React.createElement("div", {
    style: {
      fontFamily: "monospace",
      fontSize: 32,
      fontWeight: 800,
      color: display ? C.navy : C.g300,
      textAlign: "right",
      lineHeight: 1.1
    }
  }, display || /*#__PURE__*/React.createElement("span", {
    style: {
      color: C.g200
    }
  }, "\u2014")))), KEYS.map((row, ri) => /*#__PURE__*/React.createElement("div", {
    key: ri,
    style: {
      display: "flex",
      gap: 6
    }
  }, row.map(k => /*#__PURE__*/React.createElement("button", {
    key: k,
    onClick: () => !isCheck && onPress(k),
    disabled: isCheck,
    style: {
      ...kb(!isCheck),
      flex: ri === 3 && k === 0 ? 2 : 1
    }
  }, k)))), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 6
    }
  }, ["○", "×"].map(v => /*#__PURE__*/React.createElement("button", {
    key: v,
    onClick: () => isCheck && onCheckPress && onCheckPress(v),
    disabled: !isCheck,
    style: {
      flex: 1,
      height: 56,
      borderRadius: 12,
      border: "2px solid " + (isCheck && checkValue === v ? v === "○" ? C.green : C.red : C.g200),
      background: isCheck ? checkValue === v ? v === "○" ? C.green : C.red : C.white : C.g100,
      color: isCheck ? checkValue === v ? C.white : C.g800 : C.g300,
      fontWeight: 800,
      fontSize: 28,
      cursor: isCheck ? "pointer" : "not-allowed",
      transition: "all 0.1s",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      boxShadow: isCheck ? "0 2px 6px rgba(0,0,0,0.10)" : "none"
    }
  }, v))), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 6
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: onPrev,
    disabled: !canPrev,
    style: {
      ...nb(canPrev),
      flex: 1
    }
  }, "\u25B2"), /*#__PURE__*/React.createElement("button", {
    onClick: onNext,
    disabled: !canNext,
    style: {
      ...nb(canNext),
      flex: 1
    }
  }, "\u25BC")), onSave && /*#__PURE__*/React.createElement("button", {
    onClick: onSave,
    disabled: !saveComplete,
    style: {
      marginTop: 2,
      padding: "12px 8px",
      borderRadius: 12,
      border: "none",
      cursor: saveComplete ? "pointer" : "not-allowed",
      fontWeight: 800,
      fontSize: saveComplete ? 16 : 12,
      background: saveComplete ? "linear-gradient(135deg," + C.green + ",#047857)" : C.g200,
      color: saveComplete ? C.white : C.g400,
      boxShadow: saveComplete ? "0 4px 14px rgba(5,150,105,0.35)" : "none",
      lineHeight: 1.4,
      textAlign: "center"
    }
  }, saveComplete ? "💾 保存" : "⏳ あと" + (saveMissing || 0) + "項目"));
}
function VoiceMemoPanel({
  scope,
  recordId
}) {
  if (!IS_IPAD) return null;
  const [memos, setMemos] = useState([]);
  const [recording, setRecording] = useState(false);
  const [elapsed, setElapsed] = useState(0);
  const [error, setError] = useState("");
  const recRef = useRef(null);
  const chunksRef = useRef([]);
  const timerRef = useRef(null);
  const reload = () => {
    vmListByScope(scope, recordId).then(setMemos).catch(() => {});
  };
  useEffect(() => {
    reload();
  }, [scope, recordId]);
  const start = async () => {
    setError("");
    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        audio: true
      });
      const mimeType = pickMimeType();
      const rec = mimeType ? new MediaRecorder(stream, {
        mimeType
      }) : new MediaRecorder(stream);
      chunksRef.current = [];
      rec.ondataavailable = e => {
        if (e.data && e.data.size > 0) chunksRef.current.push(e.data);
      };
      rec.onstop = async () => {
        stream.getTracks().forEach(t => t.stop());
        const blob = new Blob(chunksRef.current, {
          type: mimeType || "audio/mp4"
        });
        await vmAdd({
          id: crypto.randomUUID(),
          scope,
          recordId,
          ts: Date.now(),
          mimeType: mimeType || "audio/mp4",
          blob
        });
        reload();
      };
      recRef.current = rec;
      rec.start();
      setRecording(true);
      setElapsed(0);
      timerRef.current = setInterval(() => setElapsed(p => p + 1), 1000);
    } catch (e) {
      setError("マイクにアクセスできませんでした");
    }
  };
  const stop = () => {
    if (recRef.current && recRef.current.state !== "inactive") recRef.current.stop();
    clearInterval(timerRef.current);
    setRecording(false);
  };
  const del = id => {
    vmDelete(id).then(reload);
  };
  const fmtTime = s => Math.floor(s / 60) + ":" + String(s % 60).padStart(2, "0");
  const fmtTs = ts => {
    const d = new Date(ts);
    return d.getMonth() + 1 + "/" + d.getDate() + " " + String(d.getHours()).padStart(2, "0") + ":" + String(d.getMinutes()).padStart(2, "0");
  };
  return /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 6,
      marginTop: 6
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: 8
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: recording ? stop : start,
    style: {
      padding: "7px 14px",
      borderRadius: 8,
      border: "none",
      cursor: "pointer",
      fontWeight: 700,
      fontSize: 12,
      background: recording ? "#DC2626" : "#2563B0",
      color: "#fff",
      whiteSpace: "nowrap"
    }
  }, recording ? "⏹️ 停止（" + fmtTime(elapsed) + "）" : "🎙️ 録音開始"), error && /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 11,
      color: "#DC2626"
    }
  }, error)), memos.length > 0 && /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 4,
      maxHeight: 160,
      overflowY: "auto"
    }
  }, memos.map(m => /*#__PURE__*/React.createElement("div", {
    key: m.id,
    style: {
      display: "flex",
      alignItems: "center",
      gap: 8,
      background: "#F8FAFC",
      borderRadius: 8,
      padding: "5px 8px"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 10,
      color: "#64748B",
      whiteSpace: "nowrap"
    }
  }, fmtTs(m.ts)), /*#__PURE__*/React.createElement("audio", {
    controls: true,
    src: URL.createObjectURL(m.blob),
    style: {
      height: 28,
      flex: 1
    }
  }), /*#__PURE__*/React.createElement("button", {
    onClick: () => del(m.id),
    style: {
      padding: "3px 7px",
      borderRadius: 6,
      border: "none",
      cursor: "pointer",
      background: "#DC2626",
      color: "#fff",
      fontSize: 10,
      fontWeight: 700
    }
  }, "\uD83D\uDDD1\uFE0F")))));
}
function RoundSelector({
  current,
  onSelect,
  onCancel
}) {
  const now = new Date();
  const [tab, setTab] = useState("new"); // "new" | "past"
  const [year, setYear] = useState(now.getFullYear());
  const [month, setMonth] = useState(now.getMonth() + 1);
  const [pastList, setPastList] = useState(null); // null=未取得
  const [error, setError] = useState("");
  useEffect(() => {
    if (tab !== "past" || !db) return;
    db.ref("roundsIndex").once("value").then(snap => {
      const d = snap.val() || {};
      const arr = Object.keys(d).map(id => ({
        id,
        ...d[id]
      }));
      arr.sort((a, b) => b.year - a.year || b.month - a.month);
      setPastList(arr);
    }).catch(() => setPastList([]));
  }, [tab]);
  const label = year + "年" + month + "月点検分";
  const startNew = async () => {
    setError("");
    const id = year + "-" + String(month).padStart(2, "0");
    try {
      if (db) {
        const snap = await db.ref("roundsIndex/" + id).once("value");
        if (!snap.val()) {
          await db.ref("roundsIndex/" + id).set({
            label,
            year,
            month,
            createdAt: window.firebase.database.ServerValue.TIMESTAMP
          });
        }
      }
      onSelect({
        id,
        label
      });
    } catch (e) {
      setError("開始できませんでした（通信状況をご確認ください）");
    }
  };
  const yearOpts = [];
  for (let y = now.getFullYear() - 1; y <= now.getFullYear() + 3; y++) yearOpts.push(y);
  return /*#__PURE__*/React.createElement("div", {
    style: {
      width: "100%",
      maxWidth: 440,
      margin: "0 auto",
      display: "flex",
      flexDirection: "column",
      gap: 16
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 8
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: () => setTab("new"),
    style: {
      flex: 1,
      padding: "10px",
      borderRadius: 9,
      border: "none",
      cursor: "pointer",
      fontWeight: 700,
      fontSize: 14,
      background: tab === "new" ? C.blue : C.g100,
      color: tab === "new" ? "#fff" : C.g600
    }
  }, "\uD83C\uDD95 \u65B0\u898F\u70B9\u691C\u958B\u59CB"), /*#__PURE__*/React.createElement("button", {
    onClick: () => setTab("past"),
    style: {
      flex: 1,
      padding: "10px",
      borderRadius: 9,
      border: "none",
      cursor: "pointer",
      fontWeight: 700,
      fontSize: 14,
      background: tab === "past" ? C.blue : C.g100,
      color: tab === "past" ? "#fff" : C.g600
    }
  }, "\uD83D\uDCC2 \u904E\u53BB\u30C7\u30FC\u30BF\u8AAD\u307F\u8FBC\u307F")), tab === "new" && /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 12,
      background: "#fff",
      borderRadius: 14,
      padding: 18,
      boxShadow: "0 1px 8px rgba(0,0,0,0.08)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 10
    }
  }, /*#__PURE__*/React.createElement("select", {
    value: year,
    onChange: e => setYear(Number(e.target.value)),
    style: {
      flex: 1,
      padding: "10px",
      borderRadius: 8,
      border: "1.5px solid " + C.g200,
      fontSize: 15
    }
  }, yearOpts.map(y => /*#__PURE__*/React.createElement("option", {
    key: y,
    value: y
  }, y, "\u5E74"))), /*#__PURE__*/React.createElement("select", {
    value: month,
    onChange: e => setMonth(Number(e.target.value)),
    style: {
      flex: 1,
      padding: "10px",
      borderRadius: 8,
      border: "1.5px solid " + C.g200,
      fontSize: 15
    }
  }, Array.from({
    length: 12
  }, (_, i) => i + 1).map(m => /*#__PURE__*/React.createElement("option", {
    key: m,
    value: m
  }, m, "\u6708")))), /*#__PURE__*/React.createElement("div", {
    style: {
      textAlign: "center",
      fontSize: 16,
      fontWeight: 800,
      color: C.navy
    }
  }, label), /*#__PURE__*/React.createElement("button", {
    onClick: startNew,
    style: {
      padding: "12px",
      borderRadius: 9,
      border: "none",
      cursor: "pointer",
      fontWeight: 800,
      fontSize: 15,
      background: C.blue,
      color: "#fff"
    }
  }, "\u3053\u306E\u5185\u5BB9\u3067\u958B\u59CB\u3059\u308B"), error && /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12,
      color: "#DC2626",
      textAlign: "center"
    }
  }, error)), tab === "past" && /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 8,
      background: "#fff",
      borderRadius: 14,
      padding: 18,
      boxShadow: "0 1px 8px rgba(0,0,0,0.08)",
      maxHeight: 340,
      overflowY: "auto"
    }
  }, pastList === null && /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 13,
      color: C.g500,
      textAlign: "center"
    }
  }, "\u8AAD\u307F\u8FBC\u307F\u4E2D..."), pastList !== null && pastList.length === 0 && /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 13,
      color: C.g500,
      textAlign: "center"
    }
  }, "\u307E\u3060\u70B9\u691C\u56DE\u304C\u3042\u308A\u307E\u305B\u3093"), pastList && pastList.map(r => /*#__PURE__*/React.createElement("button", {
    key: r.id,
    onClick: () => onSelect({
      id: r.id,
      label: r.label
    }),
    style: {
      padding: "12px 14px",
      borderRadius: 9,
      border: "1.5px solid " + (current?.id === r.id ? C.blue : C.g200),
      cursor: "pointer",
      fontSize: 14,
      fontWeight: 700,
      background: current?.id === r.id ? C.blue + "10" : "#fff",
      color: C.g800,
      textAlign: "left"
    }
  }, r.label, current?.id === r.id && /*#__PURE__*/React.createElement("span", {
    style: {
      marginLeft: 8,
      fontSize: 11,
      color: C.blue
    }
  }, "\uFF08\u73FE\u5728\u306E\u56DE\uFF09")))), onCancel && /*#__PURE__*/React.createElement("button", {
    onClick: onCancel,
    style: {
      padding: "10px",
      borderRadius: 9,
      border: "1.5px solid " + C.g200,
      cursor: "pointer",
      fontWeight: 700,
      fontSize: 13,
      background: "#fff",
      color: C.g600
    }
  }, "\u2715 \u30AD\u30E3\u30F3\u30BB\u30EB\uFF08", current?.label, "\u306E\u307E\u307E\u306B\u3059\u308B\uFF09"));
}
function S1Head({
  num,
  label,
  done,
  active,
  onClick
}) {
  return /*#__PURE__*/React.createElement("div", {
    onClick: onClick,
    style: {
      display: "flex",
      alignItems: "center",
      gap: 8,
      padding: "8px 14px",
      cursor: "pointer",
      background: done ? C.green : active ? C.blue + "10" : C.g50,
      borderBottom: "1px solid " + C.g100,
      transition: "background 0.15s"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      width: 22,
      height: 22,
      borderRadius: 11,
      flexShrink: 0,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontWeight: 700,
      fontSize: 11,
      background: done ? C.white : active ? C.blue + "30" : C.g200,
      color: done ? C.green : active ? C.blue : C.g500
    }
  }, done ? "✓" : num), /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 13,
      fontWeight: 700,
      color: done ? C.white : active ? C.blue : C.g600,
      flex: 1
    }
  }, label));
}
function Step2View({
  form,
  setInfo,
  handleSave,
  setStep,
  visIn,
  visOut,
  visFields,
  activeCode,
  numDisp,
  limits,
  onPress,
  onConfirm,
  onRowClick,
  moveActive,
  rowRefs,
  listRef,
  complete,
  missing,
  editIdx,
  ALL_FIELDS,
  vis,
  isAbn,
  setCheck,
  checkFields,
  inspectionMode,
  focusSeq,
  isCheckCode,
  setCheckAndAdvance,
  hideHeader,
  hideNumpad,
  outdoorLocked,
  requestUnlockOutdoor
}) {
  const isOutdoor = inspectionMode === "outdoor";
  const chkFields = checkFields || [];
  const outChkFields = chkFields.filter(f => f.group === "check_out");
  const ciFields = chkFields.filter(f => f.group === "check_in");
  // 室外機モード：チェック項目のみで完了判定
  const filled = isOutdoor ? outChkFields.filter(f => (form.checks?.[f.code] || "") !== "").length : visFields.filter(f => form.values[f.code] !== "").length;
  const total = isOutdoor ? outChkFields.length : visFields.length;
  const pct = total > 0 ? Math.round(filled / total * 100) : 0;
  const seq = focusSeq || visFields;
  const ai = seq.findIndex(f => f.code === activeCode);
  const activeIsCheck = !isOutdoor && !!activeCode && !!isCheckCode && isCheckCode(activeCode);
  const activeCheckField = activeIsCheck ? ciFields.find(f => f.code === activeCode) : null;
  return /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      display: "flex",
      flexDirection: "column",
      overflow: "hidden"
    }
  }, !hideHeader && /*#__PURE__*/React.createElement("div", {
    style: {
      flexShrink: 0,
      background: "linear-gradient(135deg," + C.navy + "," + C.blue + ")",
      padding: "8px 14px",
      boxShadow: "0 2px 8px rgba(0,0,0,0.2)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 14,
      flexWrap: "wrap",
      alignItems: "center"
    }
  }, [["点検日", form.inspectionDate], ["点検者", form.inspector], ["階", form.floor], ["部屋名", form.room], ["管理番号", form.managementNo], ["機器番号", form.unitNo]].map(([k, v]) => /*#__PURE__*/React.createElement("div", {
    key: k,
    style: {
      display: "flex",
      gap: 4,
      alignItems: "baseline"
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 9,
      fontWeight: 700,
      color: "rgba(255,255,255,0.55)",
      textTransform: "uppercase",
      letterSpacing: "0.05em"
    }
  }, k), /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 13,
      fontWeight: 700,
      color: C.white
    }
  }, v || "—"))), /*#__PURE__*/React.createElement("button", {
    onClick: () => setStep(1),
    style: {
      marginLeft: "auto",
      padding: "4px 12px",
      borderRadius: 6,
      border: "1.5px solid rgba(255,255,255,0.4)",
      background: "rgba(255,255,255,0.15)",
      color: C.white,
      cursor: "pointer",
      fontSize: 12,
      fontWeight: 700
    }
  }, "\u2190 \u4FEE\u6B63"))), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      display: "flex",
      overflow: "hidden"
    }
  }, isOutdoor && /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      overflowY: "auto",
      scrollbarGutter: "stable",
      background: C.white,
      padding: "8px"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12,
      fontWeight: 700,
      color: C.teal,
      marginBottom: 8,
      padding: "6px 10px",
      background: C.teal + "10",
      borderRadius: 8
    }
  }, "\uD83C\uDFED \u5BA4\u5916\u6A5F\u70B9\u691C\u30C1\u30A7\u30C3\u30AF\uFF08", form.floor, " ", form.room, " ", form.managementNo, "\uFF09"), (() => {
    const cats = [...new Set(outChkFields.map(f => f.category))];
    return cats.map(cat => /*#__PURE__*/React.createElement("div", {
      key: cat,
      style: {
        marginBottom: 12
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        fontSize: 11,
        fontWeight: 700,
        color: C.teal,
        marginBottom: 4,
        padding: "3px 8px",
        background: C.teal + "15",
        borderRadius: 5,
        display: "inline-block"
      }
    }, cat), outChkFields.filter(f => f.category === cat).map((f, i) => {
      const val = form.checks?.[f.code] || "";
      return /*#__PURE__*/React.createElement("div", {
        key: f.code,
        style: {
          display: "flex",
          alignItems: "center",
          gap: 8,
          padding: "8px 10px",
          borderBottom: "1px solid " + C.g100,
          background: i % 2 === 0 ? C.white : C.g50
        }
      }, /*#__PURE__*/React.createElement("span", {
        style: {
          flex: 1,
          fontSize: 14,
          color: C.g700
        }
      }, f.label), /*#__PURE__*/React.createElement("div", {
        style: {
          display: "flex",
          gap: 6
        }
      }, ["○", "×"].map(v => /*#__PURE__*/React.createElement("button", {
        key: v,
        onClick: () => setCheck(f.code, v),
        style: {
          width: 52,
          height: 44,
          borderRadius: 10,
          border: "2px solid " + (val === v ? v === "○" ? C.green : C.red : C.g200),
          background: val === v ? v === "○" ? C.green : C.red : C.white,
          color: val === v ? C.white : C.g400,
          fontWeight: 800,
          fontSize: 22,
          cursor: "pointer",
          transition: "all 0.1s"
        }
      }, v))));
    })));
  })()), !isOutdoor && /*#__PURE__*/React.createElement("div", {
    ref: listRef,
    style: {
      flex: 1,
      overflowY: "auto",
      scrollbarGutter: "stable",
      background: C.white,
      position: "relative"
    }
  }, /*#__PURE__*/React.createElement("table", {
    style: {
      width: "100%",
      borderCollapse: "collapse"
    }
  }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", {
    style: {
      background: C.g100,
      position: "sticky",
      top: 0,
      zIndex: 2
    }
  }, /*#__PURE__*/React.createElement("th", {
    style: {
      width: 5,
      padding: 0,
      borderBottom: "2px solid " + C.g200
    }
  }), [["コード", 62], ["項目名", null], ["単位", 48], ["測定値", 110]].map(([h, w]) => /*#__PURE__*/React.createElement("th", {
    key: h,
    style: {
      padding: "8px 8px",
      fontSize: 11,
      fontWeight: 700,
      color: C.g500,
      textAlign: h === "測定値" ? "right" : "center",
      borderBottom: "2px solid " + C.g200,
      width: w || undefined,
      whiteSpace: "nowrap"
    }
  }, h)))), /*#__PURE__*/React.createElement("tbody", null, visIn.length > 0 && /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("td", {
    colSpan: 5,
    style: {
      padding: "6px 10px",
      background: C.blue + "12",
      fontSize: 11,
      fontWeight: 700,
      color: C.blue,
      borderBottom: "1px solid " + C.blue + "20"
    }
  }, "\u5BA4\u5185\u6A5F\uFF08\u30A4\u30F3\u30C9\u30A2\uFF09")), visIn.map((f, i) => {
    const act = activeCode === f.code;
    const v = act ? numDisp : form.values[f.code];
    return /*#__PURE__*/React.createElement(FieldRow, {
      key: f.code,
      f: f,
      isIn: true,
      idx: i,
      active: act,
      val: v,
      abn: isAbn(f.code, v, limits),
      dimmed: !!activeCode && !act,
      onClick: () => onRowClick(f),
      fRef: el => {
        rowRefs.current[f.code] = el;
      }
    });
  })), visOut.length > 0 && /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("td", {
    colSpan: 5,
    style: {
      padding: "6px 10px",
      background: C.teal + "12",
      fontSize: 11,
      fontWeight: 700,
      color: C.teal,
      borderBottom: "1px solid " + C.teal + "20"
    }
  }, "\u5BA4\u5916\u6A5F\uFF08\u30A2\u30A6\u30C8\u30C9\u30A2\uFF09")), outdoorLocked ? /*#__PURE__*/React.createElement("tr", {
    onClick: () => requestUnlockOutdoor && requestUnlockOutdoor(),
    style: {
      cursor: "pointer",
      background: "#FFF7ED"
    }
  }, /*#__PURE__*/React.createElement("td", {
    colSpan: 5,
    style: {
      padding: "14px 10px",
      textAlign: "center",
      fontSize: 13,
      fontWeight: 700,
      color: "#92400E"
    }
  }, "\uD83D\uDD12 \u5BA4\u5916\u6A5F\u30C7\u30FC\u30BF\u306F\u5165\u529B\u6E08\u307F\u3067\u3059\u3002\u30BF\u30C3\u30D7\u3057\u3066\u5165\u529B\u3059\u308B")) : visOut.map((f, i) => {
    const act = activeCode === f.code;
    const v = act ? numDisp : form.values[f.code];
    return /*#__PURE__*/React.createElement(FieldRow, {
      key: f.code,
      f: f,
      isIn: false,
      idx: i,
      active: act,
      val: v,
      abn: isAbn(f.code, v, limits),
      dimmed: !!activeCode && !act,
      onClick: () => onRowClick(f),
      fRef: el => {
        rowRefs.current[f.code] = el;
      }
    });
  })))), ciFields.length > 0 && /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "14px 12px 20px",
      background: C.g50,
      borderTop: "2px solid " + C.g200
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12,
      fontWeight: 800,
      color: "#059669",
      marginBottom: 8
    }
  }, "\u2705 \u5BA4\u5185\u6A5F\u30C1\u30A7\u30C3\u30AF\u9805\u76EE"), /*#__PURE__*/React.createElement("table", {
    style: {
      width: "100%",
      borderCollapse: "collapse"
    }
  }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("th", {
    style: {
      width: 90,
      padding: "6px 8px",
      textAlign: "left",
      fontSize: 10,
      fontWeight: 700,
      color: C.g500,
      borderBottom: "2px solid " + C.g200
    }
  }, "\u9805\u76EE"), /*#__PURE__*/React.createElement("th", {
    style: {
      padding: "6px 8px",
      textAlign: "left",
      fontSize: 10,
      fontWeight: 700,
      color: C.g500,
      borderBottom: "2px solid " + C.g200
    }
  }, "\u70B9\u691C\u5185\u5BB9"), /*#__PURE__*/React.createElement("th", {
    style: {
      width: 110,
      padding: "6px 8px",
      textAlign: "center",
      fontSize: 10,
      fontWeight: 700,
      color: C.g500,
      borderBottom: "2px solid " + C.g200
    }
  }, "\u25CB\xD7"))), /*#__PURE__*/React.createElement("tbody", null, ciFields.map((f, i) => {
    const val = form.checks?.[f.code] || "";
    const act = activeCode === f.code;
    return /*#__PURE__*/React.createElement("tr", {
      key: f.code,
      ref: el => {
        if (rowRefs) rowRefs.current[f.code] = el;
      },
      onClick: () => onRowClick && onRowClick(f),
      style: {
        background: act ? C.blue + "0C" : i % 2 === 0 ? C.white : C.g50,
        boxShadow: act ? "inset 0 0 0 2px " + C.blue : "none",
        cursor: "pointer",
        transition: "box-shadow 0.1s"
      }
    }, /*#__PURE__*/React.createElement("td", {
      style: {
        padding: "9px 8px",
        fontSize: 11,
        fontWeight: 700,
        color: C.teal,
        borderBottom: "1px solid " + C.g100,
        verticalAlign: "middle"
      }
    }, f.category), /*#__PURE__*/React.createElement("td", {
      style: {
        padding: "9px 8px",
        fontSize: 14,
        color: C.g700,
        borderBottom: "1px solid " + C.g100,
        verticalAlign: "middle"
      }
    }, f.label), /*#__PURE__*/React.createElement("td", {
      style: {
        padding: "6px 8px",
        borderBottom: "1px solid " + C.g100,
        verticalAlign: "middle"
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        gap: 6,
        justifyContent: "center"
      },
      onClick: e => e.stopPropagation()
    }, ["○", "×"].map(v => /*#__PURE__*/React.createElement("button", {
      key: v,
      onClick: () => setCheckAndAdvance ? setCheckAndAdvance(f.code, v) : setCheck(f.code, v),
      style: {
        width: 48,
        height: 40,
        borderRadius: 9,
        border: "2px solid " + (val === v ? v === "○" ? C.green : C.red : C.g200),
        background: val === v ? v === "○" ? C.green : C.red : C.white,
        color: val === v ? C.white : C.g400,
        fontWeight: 800,
        fontSize: 19,
        cursor: "pointer",
        transition: "all 0.1s"
      }
    }, v)))));
  }))))), !isOutdoor && !hideNumpad && /*#__PURE__*/React.createElement(Numpad, {
    mode: activeIsCheck ? "check" : "numeric",
    display: numDisp,
    onPress: onPress,
    onConfirm: onConfirm,
    canConfirm: !!activeCode && !activeIsCheck && numDisp !== "",
    checkLabel: activeCheckField?.label,
    checkCategory: activeCheckField?.category,
    checkValue: form.checks?.[activeCode] || "",
    onCheckPress: v => setCheckAndAdvance && setCheckAndAdvance(activeCode, v),
    onPrev: () => moveActive(-1),
    onNext: () => moveActive(1),
    canPrev: ai > 0,
    canNext: ai >= 0
  })), /*#__PURE__*/React.createElement("div", {
    style: {
      flexShrink: 0,
      background: C.white,
      borderTop: "2px solid " + C.g200,
      padding: "10px 12px",
      display: "flex",
      gap: 10,
      alignItems: "stretch"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      display: "flex",
      flexDirection: "column",
      gap: 4
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12,
      fontWeight: 700,
      color: C.g500
    }
  }, "\uD83D\uDCDD \u5099\u8003\u30FB\u7279\u8A18\u4E8B\u9805"), /*#__PURE__*/React.createElement("textarea", {
    value: form.remarks,
    onChange: e => setInfo("remarks", e.target.value),
    placeholder: "\u7570\u5E38\u7B87\u6240\u3001\u7279\u8A18\u4E8B\u9805\u306A\u3069...",
    style: {
      flex: 1,
      width: "100%",
      padding: "10px 12px",
      borderRadius: 9,
      fontSize: 14,
      border: "1.5px solid " + C.g200,
      background: C.inp,
      outline: "none",
      boxSizing: "border-box",
      fontFamily: "inherit",
      resize: "none",
      minHeight: 100,
      lineHeight: 1.5
    }
  }), IS_IPAD && /*#__PURE__*/React.createElement(VoiceMemoPanel, {
    scope: "record",
    recordId: form.id
  }))));
}

// ─── セッション開始画面 ────────────────────────────────────────────────────
function SessionView({
  devList,
  inspList,
  sessionInfo,
  setSessionInfo,
  onStart,
  onSelectOutdoor,
  records,
  undoneOnly,
  setUndoneOnly,
  devColumns,
  devVisibleCols
}) {
  const today = new Date().toISOString().slice(0, 10);
  const [date, setDate] = useState(sessionInfo?.date || today);
  // sessionInfo.inspectorは複数人が「・」で結合済みの表示用文字列なので、この画面自身の再表示にはinspector1〜4（生の個別値）を使う
  const [inspector, setInspector] = useState(sessionInfo?.inspector1 || "");
  const [inspector2, setInspector2] = useState(sessionInfo?.inspector2 || "");
  const [inspector3, setInspector3] = useState(sessionInfo?.inspector3 || "");
  const [inspector4, setInspector4] = useState(sessionInfo?.inspector4 || "");
  const [access, setAccess] = useState(sessionInfo?.roomAccess || {});
  const [memoOpen, setMemoOpen] = useState({});
  const [floorSortAsc, setFloorSortAsc] = useState(false); // デフォルト降順（10F→1F）
  const [selectedBuildings, setSelectedBuildings] = useState(() => sessionInfo?.selectedBuildings || []); // 選択中の建物
  // 機器種別（機器リストの「分類」列）
  const categoryKey = devColumns.find(k => /^(分類|category|class)$/i.test(k.trim())) || null;
  const allCategories = categoryKey ? [...new Set(devList.map(d => d._raw?.[categoryKey]).map(v => String(v || "").trim()).filter(Boolean))] : [];
  const hasCategoryPanel = allCategories.length > 0;
  const [selectedType, setSelectedType] = useState(() => sessionInfo?.selectedType || allCategories.find(c => /室内機/.test(c)) || null);
  // 「室内機」が選択されている場合のみ点検エリア確認を表示（分類列が無い場合は従来通り常時表示）
  const showAccessCheck = !hasCategoryPanel || !!selectedType && /室内機/.test(selectedType);
  // 入室可否チェックは必須ではなく選択制：「入室可否チェックをする」を押すまでは一覧を表示しない
  const [accessCheckStarted, setAccessCheckStarted] = useState(!!sessionInfo && !sessionInfo?.skipAccessCheck);
  const skipAccessCheck = !accessCheckStarted;
  const isOutdoorSelected = !!selectedType && /室外機/.test(selectedType);
  // 点検エリア確認の機器リストは「室内機」のみを対象にする（分類列が無い場合は従来通り全件）
  const getCategory = d => categoryKey && d._raw ? String(d._raw[categoryKey] || "").trim() : "";
  const indoorDevList = categoryKey ? devList.filter(d => /室内機/.test(getCategory(d))) : devList;
  // 対象階（前回引き継ぎ、なければ全階選択）
  const allFloors = [...new Set(devList.map(d => String(d.floor || "").trim()).filter(s => s.length > 0))].sort((a, b) => a.localeCompare(b, undefined, {
    numeric: true,
    sensitivity: "base"
  }));
  const visibleFloors = allFloors;
  const [targetFloors, setTargetFloors] = useState(() => sessionInfo?.targetFloors || []);

  // 建物列を特定してallBuildings・建物→階マップを生成
  const buildingKey = devColumns.find(k => /建物|building|棟|ビル/i.test(k)) || null;
  // _rawから階を取得するヘルパー（d.floorが空の場合に備える）
  const floorKey4Room = devColumns.find(k => /^(階|floor|フロア|階数|階層)$/i.test(k.trim())) || null;
  const getFloor = d => {
    if (d.floor) return d.floor;
    if (floorKey4Room && d._raw) return String(d._raw[floorKey4Room] || "").trim();
    return "";
  };
  const allBuildings = buildingKey ? [...new Set(devList.map(d => d._raw?.[buildingKey]).filter(Boolean))] : [];
  // 建物ごとの階リスト（_rawから取得）
  const buildingFloorMap = buildingKey ? allBuildings.reduce((acc, b) => {
    acc[b] = [...new Set(devList.filter(d => d._raw?.[buildingKey] === b).map(d => getFloor(d)).filter(Boolean))];
    return acc;
  }, {}) : {};
  // 建物選択に応じて表示する階を絞り込む
  const floorsInSelectedBuildings = selectedBuildings.length > 0 ? [...new Set(selectedBuildings.flatMap(b => buildingFloorMap[b] || []))] : visibleFloors;
  const sortedFloors = floorSortAsc ? [...floorsInSelectedBuildings] : [...floorsInSelectedBuildings].reverse();
  // 部屋リストも建物選択で絞り込む（室内機のみのindoorDevListが元）
  const filteredDevList = selectedBuildings.length > 0 && buildingKey ? indoorDevList.filter(d => selectedBuildings.includes(d._raw?.[buildingKey])) : indoorDevList;
  const rooms = [...new Map(filteredDevList.map(d => {
    const fl = getFloor(d);
    return [fl + "__" + d.room, {
      floor: fl,
      room: d.room
    }];
  }).filter(([k, v]) => v.floor || v.room)).values()].sort((a, b) => a.floor.localeCompare(b.floor, undefined, {
    numeric: true
  }) || a.room.localeCompare(b.room));
  const key = (floor, room) => floor + "__" + room;
  const setAcc = (k, val) => setAccess(p => ({
    ...p,
    [k]: {
      ...p[k],
      ...val
    }
  }));
  const getAcc = k => access[k]?.status || "OK";
  const getMemo = k => access[k]?.memo || "";
  const toggleFloor = fl => setTargetFloors(p => p.includes(fl) ? p.filter(f => f !== fl) : [...p, fl]);
  // 対象階の部屋のみでNG件数カウント
  const targetRooms = targetFloors.length > 0 ? rooms.filter(r => targetFloors.includes(r.floor)) : rooms;
  const ngCount = targetRooms.filter(r => getAcc(key(r.floor, r.room)) === "NG").length;
  const undoneCount = indoorDevList.filter(d => {
    if (targetFloors && targetFloors.length > 0 && !targetFloors.includes(d.floor)) return false;
    return !records.some(r => r.managementNo === d.managementNo && r.unitNo === d.unitNo && Object.values(r.values).some(v => v !== ""));
  }).length;
  const canStart = !!date && !!inspector && (visibleFloors.length === 0 || targetFloors.length > 0);
  const startInspection = skip => {
    if (!canStart) return;
    const combinedInspector = [inspector, inspector2, inspector3, inspector4].filter(Boolean).join("・");
    // inspector1〜4は生の個別値（SessionView再表示時にこの画面自身の状態を正しく復元するため）。
    // inspectorは結合済みの表示用文字列（点検記録side等、他の箇所が期待する形）としてそのまま残す。
    const info = {
      date,
      inspector: combinedInspector,
      inspector1: inspector,
      inspector2,
      inspector3,
      inspector4,
      targetFloors,
      roomAccess: access,
      selectedBuildings,
      selectedType,
      skipAccessCheck: skip
    };
    try {
      localStorage.setItem("acSessionInfo", JSON.stringify(info));
    } catch (e) {}
    if (isOutdoorSelected) {
      onSelectOutdoor(info);
    } else {
      onStart(info);
    }
  };
  return /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      overflowY: "auto",
      scrollbarGutter: "stable",
      padding: "10px 12px",
      display: "flex",
      flexDirection: "column",
      gap: 10,
      maxWidth: 680,
      margin: "0 auto",
      width: "100%"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: "linear-gradient(135deg," + C.navy + "," + C.blue + ")",
      borderRadius: 14,
      padding: "10px 14px",
      color: C.white,
      boxShadow: "0 2px 10px rgba(27,58,107,0.25)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 15,
      fontWeight: 800,
      marginBottom: 2
    }
  }, "\uD83D\uDEAA \u70B9\u691C\u30A8\u30EA\u30A2\u78BA\u8A8D"), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 11,
      color: "rgba(255,255,255,0.75)"
    }
  }, "\u70B9\u691C\u65E5\u30FB\u70B9\u691C\u8005\u30FB\u5BFE\u8C61\u968E\u30FB\u70B9\u691C\u30A8\u30EA\u30A2\u3092\u8A2D\u5B9A\u3057\u3066\u304B\u3089\u70B9\u691C\u3092\u958B\u59CB\u3057\u3066\u304F\u3060\u3055\u3044")), /*#__PURE__*/React.createElement("div", {
    style: {
      background: C.white,
      borderRadius: 14,
      padding: "10px 12px",
      boxShadow: "0 1px 8px rgba(0,0,0,0.07)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12,
      fontWeight: 800,
      color: C.navy,
      marginBottom: 6
    }
  }, "\uD83D\uDCC5 \u70B9\u691C\u65E5"), /*#__PURE__*/React.createElement("input", {
    type: "date",
    value: date,
    onChange: e => setDate(e.target.value),
    onClick: e => {
      const el = e.target;
      try {
        el.showPicker && el.showPicker();
      } catch (err) {}
    },
    style: {
      width: "100%",
      boxSizing: "border-box",
      fontSize: 14,
      fontWeight: 700,
      padding: "7px 10px",
      border: "2px solid " + (date ? C.blue : C.g200),
      borderRadius: 8,
      outline: "none",
      color: C.g800,
      background: C.g50,
      cursor: "pointer"
    }
  })), /*#__PURE__*/React.createElement("div", {
    style: {
      background: C.white,
      borderRadius: 14,
      padding: "10px 12px",
      boxShadow: "0 1px 8px rgba(0,0,0,0.07)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12,
      fontWeight: 800,
      color: C.navy,
      marginBottom: 6
    }
  }, "\uD83D\uDC64 \u70B9\u691C\u8005"), inspList.length > 0 ? /*#__PURE__*/React.createElement("select", {
    value: inspector,
    onChange: e => setInspector(e.target.value),
    style: {
      width: "100%",
      boxSizing: "border-box",
      fontSize: 14,
      fontWeight: 700,
      padding: "7px 10px",
      border: "2px solid " + (inspector ? C.blue : C.g200),
      borderRadius: 8,
      outline: "none",
      color: inspector ? C.g800 : C.g400,
      background: C.g50,
      appearance: "auto",
      cursor: "pointer"
    }
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u2014 \u70B9\u691C\u8005\u3092\u9078\u629E \u2014"), inspList.map(name => /*#__PURE__*/React.createElement("option", {
    key: name,
    value: name
  }, name))) : /*#__PURE__*/React.createElement("input", {
    type: "text",
    value: inspector,
    onChange: e => setInspector(e.target.value),
    placeholder: "\u70B9\u691C\u8005\u540D\u3092\u5165\u529B",
    style: {
      width: "100%",
      boxSizing: "border-box",
      fontSize: 13,
      padding: "7px 10px",
      border: "2px solid " + (inspector ? C.blue : C.g200),
      borderRadius: 8,
      outline: "none",
      background: C.g50
    }
  }), [{
    label: "二人目",
    val: inspector2,
    set: v => {
      setInspector2(v);
      if (!v) {
        setInspector3("");
        setInspector4("");
      }
    },
    exclude: [inspector],
    disabled: !inspector
  }, ...(inspector2 ? [{
    label: "三人目",
    val: inspector3,
    set: v => {
      setInspector3(v);
      if (!v) setInspector4("");
    },
    exclude: [inspector, inspector2],
    disabled: false
  }] : []), ...(inspector2 && inspector3 ? [{
    label: "四人目",
    val: inspector4,
    set: setInspector4,
    exclude: [inspector, inspector2, inspector3],
    disabled: false
  }] : [])].map(({
    label,
    val,
    set,
    exclude,
    disabled
  }) => /*#__PURE__*/React.createElement("div", {
    key: label,
    style: {
      marginTop: 8,
      display: "flex",
      gap: 6,
      alignItems: "center"
    }
  }, inspList.length > 0 ? /*#__PURE__*/React.createElement("select", {
    value: val,
    onChange: e => set(e.target.value),
    disabled: disabled,
    style: {
      flex: 1,
      minWidth: 0,
      boxSizing: "border-box",
      fontSize: 14,
      fontWeight: 700,
      padding: "7px 10px",
      border: "2px solid " + (val ? C.blue : C.g200),
      borderRadius: 8,
      outline: "none",
      color: val ? C.g800 : C.g400,
      background: disabled ? C.g100 : C.g50,
      appearance: "auto",
      cursor: disabled ? "not-allowed" : "pointer",
      opacity: disabled ? 0.55 : 1
    }
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u2014 ", label, "\u306E\u70B9\u691C\u8005\u3092\u9078\u629E \u2014"), inspList.filter(n => !exclude.includes(n)).map(name => /*#__PURE__*/React.createElement("option", {
    key: name,
    value: name
  }, name))) : /*#__PURE__*/React.createElement("input", {
    type: "text",
    value: val,
    onChange: e => set(e.target.value),
    disabled: disabled,
    placeholder: label + "の点検者名を入力",
    style: {
      flex: 1,
      minWidth: 0,
      boxSizing: "border-box",
      fontSize: 13,
      padding: "7px 10px",
      border: "2px solid " + (val ? C.blue : C.g200),
      borderRadius: 8,
      outline: "none",
      background: disabled ? C.g100 : C.g50,
      opacity: disabled ? 0.55 : 1
    }
  }), /*#__PURE__*/React.createElement("button", {
    onClick: () => set(""),
    style: {
      padding: "7px 10px",
      borderRadius: 8,
      border: "1.5px solid " + C.g200,
      background: C.g50,
      color: C.g500,
      cursor: "pointer",
      fontSize: 12,
      fontWeight: 700,
      flexShrink: 0
    }
  }, "\u2715")))), allFloors.length === 0 ? /*#__PURE__*/React.createElement("div", {
    style: {
      background: "#FFF7ED",
      borderRadius: 14,
      padding: "10px 12px",
      border: "1.5px solid #F59E0B",
      boxShadow: "0 1px 8px rgba(0,0,0,0.07)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12,
      fontWeight: 800,
      color: "#92400E",
      marginBottom: 4
    }
  }, "\uD83C\uDFE2 \u672C\u65E5\u306E\u5BFE\u8C61\u30A8\u30EA\u30A2"), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 11,
      color: "#92400E"
    }
  }, "\u26A0\uFE0F \u6A5F\u5668\u30EA\u30B9\u30C8CSV\u304C\u8AAD\u307F\u8FBC\u307E\u308C\u3066\u3044\u307E\u305B\u3093\u3002", /*#__PURE__*/React.createElement("br", null), "\u8A2D\u5B9A\u753B\u9762\u304B\u3089\u6A5F\u5668\u30EA\u30B9\u30C8CSV\u3092\u8AAD\u307F\u8FBC\u3080\u3068\u3001\u30A8\u30EA\u30A2\u306E\u9078\u629E\u30FB\u70B9\u691C\u30A8\u30EA\u30A2\u78BA\u8A8D\u304C\u5229\u7528\u3067\u304D\u307E\u3059\u3002")) : /*#__PURE__*/React.createElement("div", {
    style: {
      background: C.white,
      borderRadius: 14,
      padding: "10px 12px",
      boxShadow: "0 1px 8px rgba(0,0,0,0.07)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: 8,
      marginBottom: 6
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12,
      fontWeight: 800,
      color: C.navy,
      flex: 1
    }
  }, "\uD83C\uDFE2 \u672C\u65E5\u306E\u5BFE\u8C61\u30A8\u30EA\u30A2")), allBuildings.length > 0 && /*#__PURE__*/React.createElement("div", {
    style: {
      marginBottom: 8
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 11,
      fontWeight: 700,
      color: C.g500,
      marginBottom: 5
    }
  }, "\uD83C\uDFD7\uFE0F \u5EFA\u7269"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexWrap: "wrap",
      gap: 5
    }
  }, allBuildings.map(b => {
    const sel = selectedBuildings.includes(b);
    return /*#__PURE__*/React.createElement("button", {
      key: b,
      onClick: () => {
        setSelectedBuildings(sel ? [] : [b]);
        setTargetFloors([]); // 建物変更時に階選択をリセット
      },
      style: {
        padding: "7px 14px",
        borderRadius: 8,
        border: "2px solid " + (sel ? C.navy : C.g200),
        background: sel ? "linear-gradient(135deg," + C.navy + ",#374151)" : C.white,
        color: sel ? C.white : C.g600,
        fontWeight: 700,
        fontSize: 13,
        cursor: "pointer",
        transition: "all 0.15s"
      }
    }, b, sel ? " ✓" : "");
  })), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 5,
      height: "1px",
      background: C.g200
    }
  })), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexWrap: "wrap",
      gap: 5
    }
  }, (() => {
    const allSel = floorsInSelectedBuildings.length > 0 && floorsInSelectedBuildings.every(f => targetFloors.includes(f));
    return /*#__PURE__*/React.createElement("button", {
      onClick: () => setTargetFloors(allSel ? [] : floorsInSelectedBuildings),
      style: {
        padding: "7px 14px",
        borderRadius: 8,
        border: "2px solid " + (allSel ? C.teal : C.g200),
        background: allSel ? "linear-gradient(135deg," + C.teal + ",#0D9488)" : C.white,
        color: allSel ? C.white : C.g600,
        fontWeight: 700,
        fontSize: 13,
        cursor: "pointer",
        transition: "all 0.15s",
        minWidth: 62,
        textAlign: "center"
      }
    }, "\u3059\u3079\u3066", allSel ? " ✓" : "");
  })(), sortedFloors.map(fl => {
    const sel = targetFloors.includes(fl);
    return /*#__PURE__*/React.createElement("button", {
      key: fl,
      onClick: () => toggleFloor(fl),
      style: {
        padding: "7px 14px",
        borderRadius: 8,
        border: "2px solid " + (sel ? C.blue : C.g200),
        background: sel ? "linear-gradient(135deg," + C.navy + "," + C.blue + ")" : C.white,
        color: sel ? C.white : C.g600,
        fontWeight: 700,
        fontSize: 13,
        cursor: "pointer",
        transition: "all 0.15s",
        minWidth: 50,
        textAlign: "center"
      }
    }, fl, sel ? " ✓" : "");
  })), targetFloors.length === 0 && /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 6,
      fontSize: 11,
      color: C.red,
      fontWeight: 700
    }
  }, "\u26A0\uFE0F \u5BFE\u8C61\u30A8\u30EA\u30A2\u30921\u3064\u4EE5\u4E0A\u9078\u629E\u3057\u3066\u304F\u3060\u3055\u3044"), targetFloors.length > 0 && indoorDevList.length > 0 && /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 8,
      display: "flex",
      gap: 6,
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: () => setUndoneOnly(false),
    style: {
      flex: 1,
      padding: "7px",
      borderRadius: 8,
      border: "2px solid " + (undoneOnly ? C.g200 : C.blue),
      background: undoneOnly ? C.white : "linear-gradient(135deg," + C.navy + "," + C.blue + ")",
      color: undoneOnly ? C.g500 : C.white,
      fontWeight: 700,
      fontSize: 12,
      cursor: "pointer",
      transition: "all 0.15s"
    }
  }, "\uD83D\uDCCB \u5168\u6A5F\u5668"), /*#__PURE__*/React.createElement("button", {
    onClick: () => setUndoneOnly(true),
    style: {
      flex: 1,
      padding: "7px",
      borderRadius: 8,
      border: "2px solid " + (undoneOnly ? C.red : C.g200),
      background: undoneOnly ? "#FEF2F2" : C.white,
      color: undoneOnly ? C.red : C.g500,
      fontWeight: 700,
      fontSize: 12,
      cursor: "pointer",
      transition: "all 0.15s"
    }
  }, "\u23F3 \u672A\u5165\u529B\u5206 ", undoneCount > 0 ? "(" + undoneCount + ")" : "(なし)")), allFloors.length > 1 && /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 8,
      display: "flex",
      justifyContent: "flex-end"
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: () => setFloorSortAsc(p => !p),
    style: {
      padding: "3px 9px",
      borderRadius: 6,
      border: "1.5px solid " + C.teal,
      background: C.teal + "18",
      color: C.teal,
      fontSize: 10,
      fontWeight: 700,
      cursor: "pointer"
    }
  }, "\u8868\u793A\u9806 ", floorSortAsc ? "▲ 昇順" : "▼ 降順"))), targetRooms.length > 0 && showAccessCheck && /*#__PURE__*/React.createElement("div", {
    style: {
      background: C.white,
      borderRadius: 14,
      padding: "10px 12px",
      boxShadow: "0 1px 8px rgba(0,0,0,0.07)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: 8,
      marginBottom: 6
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12,
      fontWeight: 800,
      color: C.navy,
      flex: 1
    }
  }, "\uD83D\uDEAA \u5165\u5BA4\u53EF\u5426\u30C1\u30A7\u30C3\u30AF"), accessCheckStarted && /*#__PURE__*/React.createElement("button", {
    onClick: () => setAccessCheckStarted(false),
    style: {
      padding: "5px 10px",
      borderRadius: 7,
      border: "1.5px solid " + C.g200,
      background: C.white,
      color: C.g500,
      fontSize: 11,
      fontWeight: 700,
      cursor: "pointer",
      whiteSpace: "nowrap"
    }
  }, "\u2190 \u623B\u308B")), !accessCheckStarted ? /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 8
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: () => setAccessCheckStarted(true),
    style: {
      flex: 1,
      padding: "12px",
      borderRadius: 9,
      border: "2px solid " + C.blue,
      background: C.blue + "0F",
      color: C.blue,
      fontWeight: 800,
      fontSize: 13,
      cursor: "pointer"
    }
  }, "\uD83D\uDEAA \u5165\u5BA4\u53EF\u5426\u30C1\u30A7\u30C3\u30AF\u3092\u3059\u308B"), /*#__PURE__*/React.createElement("button", {
    onClick: () => startInspection(true),
    disabled: !canStart,
    style: {
      flex: 1,
      padding: "12px",
      borderRadius: 9,
      border: "2px solid " + C.g300,
      background: canStart ? C.g50 : C.g100,
      color: canStart ? C.g600 : C.g400,
      fontWeight: 800,
      fontSize: 13,
      cursor: canStart ? "pointer" : "not-allowed"
    }
  }, "\u25B6\uFE0F \u3053\u306E\u307E\u307E\u70B9\u691C\u3059\u308B")) : /*#__PURE__*/React.createElement(React.Fragment, null, ngCount > 0 && /*#__PURE__*/React.createElement("div", {
    style: {
      marginBottom: 6,
      padding: "5px 10px",
      background: "#FEF2F2",
      border: "1.5px solid " + C.red,
      borderRadius: 7,
      fontSize: 11,
      fontWeight: 700,
      color: C.red
    }
  }, "\u26A0\uFE0F NG ", ngCount, "\u90E8\u5C4B \u2014 \u70B9\u691C\u30D5\u30A9\u30FC\u30E0\u3067\u30B0\u30EC\u30FC\u30A2\u30A6\u30C8\u3055\u308C\u307E\u3059"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 8
    }
  }, sortedFloors.filter(fl => targetFloors.includes(fl)).map(fl => {
    const flRooms = rooms.filter(r => r.floor === fl);
    const flNg = flRooms.filter(r => getAcc(key(r.floor, r.room)) === "NG").length;
    return /*#__PURE__*/React.createElement("div", {
      key: fl
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        alignItems: "center",
        gap: 8,
        marginBottom: 4
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        fontSize: 11,
        fontWeight: 800,
        color: C.g600,
        flex: 1,
        letterSpacing: "0.05em"
      }
    }, fl), flNg > 0 && /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 9,
        fontWeight: 700,
        color: C.red
      }
    }, "NG ", flNg)), /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        flexDirection: "column",
        gap: 3
      }
    }, flRooms.map(r => {
      const k = key(r.floor, r.room);
      const acc = getAcc(k);
      const isNG = acc === "NG";
      const memo = getMemo(k);
      const open = memoOpen[k];
      const ngType = access[k]?.ngType || "today";
      return /*#__PURE__*/React.createElement("div", {
        key: k,
        style: {
          borderRadius: 8,
          border: "1.5px solid " + (isNG ? C.red : C.g200),
          overflow: "hidden",
          transition: "border-color 0.15s"
        }
      }, /*#__PURE__*/React.createElement("div", {
        style: {
          display: "flex",
          alignItems: "center",
          gap: 7,
          padding: "7px 10px",
          background: isNG ? "#FFF0F0" : C.g50,
          cursor: isNG && !open ? "pointer" : "default"
        },
        onClick: isNG && !open ? () => setMemoOpen(p => ({
          ...p,
          [k]: true
        })) : undefined
      }, /*#__PURE__*/React.createElement("div", {
        style: {
          flex: 1,
          display: "flex",
          gap: 8,
          alignItems: "center",
          overflow: "hidden",
          minWidth: 0,
          flexWrap: "wrap"
        }
      }, (() => {
        const cols = devVisibleCols && devVisibleCols.length > 0 ? devVisibleCols : devColumns && devColumns.length > 0 ? devColumns : null;
        // devListからこの部屋の機器データを取得（_rawを持つ）
        const devItem = indoorDevList.find(d => d.floor === r.floor && d.room === r.room);
        if (cols) {
          return cols.map(col => {
            // _rawがあればそこから、なければfloor/room/managementNo/unitNoにフォールバック
            const val = devItem?._raw?.[col] || (col === "階" || col.toLowerCase() === "floor" ? r.floor : col === "部屋名" || col.toLowerCase() === "room" ? r.room : col === "管理番号" || col.toLowerCase() === "managementno" ? devItem?.managementNo : col === "機器番号" || col.toLowerCase() === "unitno" ? devItem?.unitNo : "");
            return val ? /*#__PURE__*/React.createElement("span", {
              key: col,
              style: {
                fontSize: 12,
                fontWeight: col === "部屋名" || col.toLowerCase() === "room" ? 700 : 400,
                color: isNG ? C.red : C.g800,
                whiteSpace: "nowrap"
              }
            }, val) : null;
          });
        }
        return /*#__PURE__*/React.createElement("span", {
          style: {
            fontSize: 12,
            fontWeight: 700,
            color: isNG ? C.red : C.g800
          }
        }, r.room);
      })(), isNG && !open && /*#__PURE__*/React.createElement("span", {
        style: {
          fontSize: 10,
          color: C.red,
          fontWeight: 600,
          overflow: "hidden",
          textOverflow: "ellipsis",
          whiteSpace: "nowrap",
          flexShrink: 0
        }
      }, "\uD83D\uDEAB ", [["today", "本日のみNG"], ["am", "午前のみNG"], ["pm", "午後のみNG"], ["discharge", "退院までNG"]].find(([v]) => v === ngType)?.[1] || "本日のみNG", memo && "　" + memo)), /*#__PURE__*/React.createElement("div", {
        style: {
          display: "flex",
          gap: 3
        },
        onClick: e => e.stopPropagation()
      }, ["OK", "NG"].map(s => /*#__PURE__*/React.createElement("button", {
        key: s,
        onClick: () => {
          setAcc(k, {
            status: s
          });
          setMemoOpen(p => ({
            ...p,
            [k]: s === "NG"
          }));
        },
        style: {
          padding: "5px 12px",
          borderRadius: 6,
          border: "2px solid " + (acc === s ? s === "OK" ? C.green : C.red : C.g200),
          background: acc === s ? s === "OK" ? C.green : C.red : C.white,
          color: acc === s ? C.white : C.g500,
          fontWeight: 800,
          fontSize: 11,
          cursor: "pointer",
          transition: "all 0.12s"
        }
      }, s)))), isNG && open && /*#__PURE__*/React.createElement("div", {
        style: {
          padding: "8px 10px",
          borderTop: "1.5px solid #FECACA",
          background: "#FEF2F2",
          display: "flex",
          flexDirection: "column",
          gap: 6
        }
      }, /*#__PURE__*/React.createElement("div", {
        style: {
          display: "flex",
          gap: 4,
          flexWrap: "wrap"
        }
      }, [["today", "本日のみNG"], ["am", "午前のみNG"], ["pm", "午後のみNG"], ["discharge", "退院までNG"]].map(([v, label]) => /*#__PURE__*/React.createElement("button", {
        key: v,
        onClick: () => {
          setAcc(k, {
            ngType: v
          });
          setMemoOpen(p => ({
            ...p,
            [k]: false
          }));
        },
        style: {
          padding: "4px 10px",
          borderRadius: 6,
          border: "1.5px solid " + (ngType === v ? C.red : C.g300),
          background: ngType === v ? C.red : C.white,
          color: ngType === v ? C.white : C.g600,
          fontSize: 11,
          fontWeight: 700,
          cursor: "pointer",
          whiteSpace: "nowrap",
          transition: "all 0.1s"
        }
      }, label))), /*#__PURE__*/React.createElement("textarea", {
        value: memo,
        onChange: e => setAcc(k, {
          memo: e.target.value
        }),
        placeholder: "\u5099\u8003\uFF08\u5165\u5BA4\u4E0D\u53EF\u7406\u7531\u306A\u3069\uFF09",
        inputMode: "text",
        style: {
          width: "100%",
          fontSize: 11,
          padding: "6px 9px",
          border: "1.5px solid #FECACA",
          borderRadius: 7,
          outline: "none",
          background: C.white,
          boxSizing: "border-box",
          resize: "none",
          minHeight: 66,
          lineHeight: 1.5,
          fontFamily: "inherit"
        }
      }), /*#__PURE__*/React.createElement("button", {
        onClick: () => setMemoOpen(p => ({
          ...p,
          [k]: false
        })),
        style: {
          alignSelf: "flex-end",
          padding: "4px 12px",
          borderRadius: 6,
          border: "1.5px solid " + C.g300,
          background: C.white,
          color: C.g500,
          fontSize: 10,
          fontWeight: 700,
          cursor: "pointer"
        }
      }, "\u9589\u3058\u308B")));
    })));
  })))), (!(targetRooms.length > 0 && showAccessCheck) || accessCheckStarted) && /*#__PURE__*/React.createElement("button", {
    disabled: !canStart,
    onClick: () => startInspection(skipAccessCheck),
    style: {
      padding: "13px",
      borderRadius: 12,
      border: "none",
      cursor: canStart ? "pointer" : "not-allowed",
      background: canStart ? "linear-gradient(135deg," + (isOutdoorSelected ? C.teal : C.navy) + "," + (isOutdoorSelected ? "#0D9488" : C.blue) + ")" : "#CBD5E1",
      color: C.white,
      fontWeight: 800,
      fontSize: 14,
      boxShadow: canStart ? "0 3px 12px rgba(37,99,176,0.35)" : "none",
      opacity: canStart ? 1 : 0.7,
      transition: "all 0.2s",
      flexShrink: 0
    }
  }, !date || !inspector ? "点検日・点検者を入力してください" : visibleFloors.length > 0 && targetFloors.length === 0 ? "対象エリアを選択してください" : visibleFloors.length > 0 ? (isOutdoorSelected ? "🏭 室外機点検を開始する（" : "✅ 点検を開始する（") + targetFloors.slice().sort().join("・") + ")" : isOutdoorSelected ? "🏭 室外機点検を開始する" : "✅ 点検を開始する"), /*#__PURE__*/React.createElement("div", {
    style: {
      height: 4
    }
  }));
}
function Step1View({
  form,
  setInfo,
  inspList,
  devList,
  devSearch,
  setDevSearch,
  s1DateDone,
  s1InspDone,
  s1DevDone,
  step1Valid,
  records,
  s1Focus,
  setS1Focus,
  goToStep2,
  handleStep1TmpSave,
  setForm,
  editIdx,
  lastInsp,
  lastDate,
  setStep,
  setView,
  sessionInfo,
  undoneOnly,
  setUndoneOnly,
  devColumns,
  devVisibleCols,
  inspectionMode,
  checkFields,
  setCheck,
  handleSave,
  complete,
  missing,
  visIn,
  visOut,
  visFields,
  activeCode,
  numDisp,
  limits,
  onPress,
  onConfirm,
  onRowClick,
  moveActive,
  rowRefs,
  listRef,
  ALL_FIELDS,
  vis,
  isAbn,
  focusSeq,
  isCheckCode,
  setCheckAndAdvance,
  onSwitchMode,
  outdoorLocked,
  requestUnlockOutdoor
}) {
  const allFloors = [...new Set(devList.map(d => d.floor).filter(Boolean))].sort((a, b) => a.localeCompare(b, undefined, {
    numeric: true,
    sensitivity: "base"
  }));
  // allFloorsが空の場合：_rawから階候補列を探して補完
  const allFloorsResolved = allFloors.length > 0 ? allFloors : (() => {
    const floorKey = devColumns.find(k => /^(階|floor|フロア|階数|F|階層)$/i.test(k.trim()));
    if (!floorKey) return [];
    return [...new Set(devList.map(d => String(d._raw?.[floorKey] || "").trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, undefined, {
      numeric: true,
      sensitivity: "base"
    }));
  })();
  const roomAccess = sessionInfo?.roomAccess || {};
  const targetFloors = sessionInfo?.targetFloors || null;
  const displayFloors = targetFloors && targetFloors.length > 0 ? targetFloors : allFloorsResolved;
  // 非対象階 or 入室NGならグレーアウト
  const isNG = (floor, room) => {
    if (targetFloors && targetFloors.length > 0 && !targetFloors.includes(floor)) return true;
    return (roomAccess[floor + "__" + room]?.status || "OK") === "NG";
  };
  const isNGReason = (floor, room) => {
    if (targetFloors && targetFloors.length > 0 && !targetFloors.includes(floor)) return "対象外";
    return (roomAccess[floor + "__" + room]?.status || "OK") === "NG" ? "入室NG" : null;
  };
  // 機器選択リストを、選択中のモード（室内機／室外機）かつ点検エリア確認で選択した建物・階のみに絞り込む
  const categoryKey = devColumns.find(k => /^(分類|category|class)$/i.test(k.trim())) || null;
  const buildingKey = devColumns.find(k => /建物|building|棟|ビル/i.test(k)) || null;
  const remarksKey = devColumns.find(k => /^(備考|remarks?|memo|note)$/i.test(k.trim())) || null;
  // 選択中の機器の備考。「リモコン無し」（サーミスタ計での実測が必要）は専用の注意文を、
  // それ以外の備考（「撤去」「運転off禁止」等）はそのままコメントとして表示する。点検前状態欄・データ入力欄の両方で使う
  const selDevForRemote = devList.find(d => d.managementNo === form.managementNo && d.unitNo === form.unitNo);
  const remarksValForRemote = remarksKey ? String(selDevForRemote?._raw?.[remarksKey] || "").trim() : "";
  const noRemote = /リモコン\s*無し/.test(remarksValForRemote);
  const otherRemark = remarksValForRemote && !noRemote ? remarksValForRemote : "";
  const selectedBuildings = sessionInfo?.selectedBuildings || [];
  const categoryPattern = inspectionMode === "outdoor" ? /室外機/ : /室内機/;
  const baseDevList = devList.filter(d => {
    if (categoryKey && !categoryPattern.test(String(d._raw?.[categoryKey] || "").trim())) return false;
    if (buildingKey && selectedBuildings.length > 0 && !selectedBuildings.includes(d._raw?.[buildingKey])) return false;
    if (targetFloors && targetFloors.length > 0 && !targetFloors.includes(d.floor)) return false;
    return true;
  });
  // 室外機チェック項目（画面下部リスト＋右側テンキーパネルの両方で使う）
  const outFields = (checkFields || []).filter(f => f.group === "check_out");
  const outTotal = outFields.length;
  const outFilled = outFields.filter(f => (form.checks?.[f.code] || "") !== "").length;
  const outPct = outTotal > 0 ? Math.round(outFilled / outTotal * 100) : 0;
  const isCheckFocus = outFields.some(f => f.code === s1Focus);
  const focusedCheckField = outFields.find(f => f.code === s1Focus);
  const outListRef = useRef();
  const outRowRefs = useRef({});
  // 入力欄が常に画面内に収まるよう、フォーカス項目が変わるたびにスクロール位置を調整する
  const scrollToCheckRow = (code, smooth = true) => {
    requestAnimationFrame(() => requestAnimationFrame(() => {
      const el = outRowRefs.current[code];
      const ct = outListRef.current;
      if (!el || !ct) return;
      let top = 0,
        node = el;
      while (node && node !== ct) {
        top += node.offsetTop;
        node = node.offsetParent;
      }
      const target = Math.max(0, top - ct.clientHeight / 2 + el.offsetHeight / 2);
      if (smooth) smoothScrollTo(ct, target);else ct.scrollTop = target;
    }));
  };
  const focusCheckRow = code => {
    setS1Focus(code);
    if (code) scrollToCheckRow(code);
  };
  // 点検日・点検者・建物・階・機器選択が揃ったら、Step2Viewと同じ1行バーに折りたたむ（修正ボタンで再展開）
  const [summaryExpanded, setSummaryExpanded] = useState(false);
  // 機器選択の検索窓：この画面（機器未選択の検索・一覧状態）になったら自動的にフォーカスする（室内機・室外機とも）
  const devSearchRef = useRef(null);
  useEffect(() => {
    if (!s1DevDone && devSearchRef.current) {
      devSearchRef.current.focus();
    }
  }, [s1DevDone, inspectionMode]);
  // 点検前状態：一次保存/スキップを押すと折りたたむ（測定データ入力はブロックしない。機器を切り替えたら自動的に展開状態に戻る）
  const [preStateConfirmedFor, setPreStateConfirmedFor] = useState(null);
  const deviceKey = (form.managementNo || "") + "|" + (form.unitNo || "");
  const preStateOpen = preStateConfirmedFor !== deviceKey;
  // 機器リストのソート：点検前状態・データ入力の済/未でスライドボタン切替（ONのとき未入力を先頭に表示）
  const [preStateSortDir, setPreStateSortDir] = useState(null); // null|"done"|"undone"
  const [dataSortDir, setDataSortDir] = useState(null); // null|"done"|"undone"
  const devPreStateDone = d => records.some(r => r.managementNo === d.managementNo && r.unitNo === d.unitNo && (r.preOperation || r.preMode || r.preWind || r.preSetTemp));
  const devDataDone = d => records.some(r => r.managementNo === d.managementNo && r.unitNo === d.unitNo && (inspectionMode === "outdoor" ? Object.values(r.checks || {}).some(v => v !== "") : Object.values(r.values).some(v => v !== "")));
  // 機器選択時の共通処理：機器を選ぶとすぐ下にデータ入力（測定値・チェック）が表示されるため、最初の項目に自動フォーカスする
  const selectDevice = dev => {
    setForm(p => ({
      ...p,
      floor: dev.floor,
      room: dev.room,
      managementNo: dev.managementNo,
      unitNo: dev.unitNo,
      preOperation: "",
      preMode: "",
      preWind: "",
      preSetTemp: ""
    }));
    setDevSearch("");
    setSummaryExpanded(false);
    if (inspectionMode === "outdoor" && outFields.length > 0) {
      focusCheckRow(outFields[0].code);
    } else if (inspectionMode === "indoor" && focusSeq && focusSeq.length > 0) {
      onRowClick && onRowClick(focusSeq[0]);
    } else {
      setS1Focus(null);
    }
  };
  // 建物・階などの絞り込みで候補が1件だけになっている場合は、タップしなくても自動的に選択する
  // （画面を開き直すたびに同じ機器を選び直す手間を無くすため）
  useEffect(() => {
    // summaryExpandedがtrue＝「修正」等でユーザーが自分から機器選択に戻った直後なので、勝手に選び直さない
    if (!summaryExpanded && !s1DevDone && !devSearch && s1InspDone && baseDevList.length === 1) {
      selectDevice(baseDevList[0]);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [baseDevList.length, s1DevDone, devSearch, s1InspDone, summaryExpanded]);
  // チェックの値を設定し、続けて次のチェック項目へフォーカスを進める（テンキーパネル・一覧の○×ボタン共通）
  // ※反転（同じ値を押すとクリア）はしない。何度押しても選んだ値のまま（クリアはCLRボタンで行う）
  const advanceCheck = (code, v) => {
    setCheck && setCheck(code, v);
    const idx = outFields.findIndex(f => f.code === code);
    const next = outFields[idx + 1];
    if (next) {
      setS1Focus(next.code);
      scrollToCheckRow(next.code);
    } else {
      setS1Focus(null);
    }
  };
  const topSummary = !step1Valid || summaryExpanded ? /*#__PURE__*/React.createElement(React.Fragment, null, (() => {
    const bKey = devColumns.find(k => /建物|building|棟|ビル/i.test(k));
    const sessionBuildings = sessionInfo?.selectedBuildings || [];
    const showBuildings = bKey ? sessionBuildings.length > 0 ? sessionBuildings : [...new Set(devList.map(d => d._raw?.[bKey]).filter(Boolean))] : [];
    const floorsSorted = [...displayFloors].sort((a, b) => b.localeCompare(a, undefined, {
      numeric: true
    }));
    return /*#__PURE__*/React.createElement("div", {
      style: {
        background: C.white,
        borderRadius: 12,
        padding: "10px 12px",
        boxShadow: "0 1px 6px rgba(0,0,0,0.06)",
        display: "flex",
        alignItems: "center",
        gap: 12
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        flex: 1,
        display: "flex",
        flexDirection: "column",
        gap: 6,
        minWidth: 0
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        gap: 16,
        rowGap: 6,
        flexWrap: "wrap"
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        alignItems: "baseline",
        gap: 6
      }
    }, /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 11,
        fontWeight: 700,
        color: C.green,
        whiteSpace: "nowrap"
      }
    }, "\u2713 \u70B9\u691C\u65E5"), /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 14,
        fontWeight: 700,
        color: C.navy,
        fontFamily: "monospace"
      }
    }, form.inspectionDate || "—")), /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        alignItems: "baseline",
        gap: 6
      }
    }, /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 11,
        fontWeight: 700,
        color: C.green,
        whiteSpace: "nowrap"
      }
    }, "\u2713 \u70B9\u691C\u8005"), /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 14,
        fontWeight: 700,
        color: C.navy
      }
    }, form.inspector || "—"))), /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        gap: 16,
        rowGap: 6,
        flexWrap: "wrap"
      }
    }, showBuildings.length > 0 && /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        alignItems: "baseline",
        gap: 6,
        flexWrap: "wrap"
      }
    }, /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 11,
        fontWeight: 700,
        color: C.green,
        whiteSpace: "nowrap"
      }
    }, "\u2713 \u5EFA\u7269"), showBuildings.map(b => /*#__PURE__*/React.createElement("span", {
      key: b,
      style: {
        fontSize: 13,
        fontWeight: 700,
        color: C.navy,
        background: C.navy + "10",
        padding: "2px 9px",
        borderRadius: 6
      }
    }, b))), allFloorsResolved.length > 0 && /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        alignItems: "baseline",
        gap: 6,
        flexWrap: "wrap"
      }
    }, /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 11,
        fontWeight: 700,
        color: C.green,
        whiteSpace: "nowrap"
      }
    }, "\u2713 \u968E"), floorsSorted.map(fl => /*#__PURE__*/React.createElement("span", {
      key: fl,
      style: {
        fontSize: 13,
        fontWeight: 700,
        color: C.blue,
        background: C.blue + "15",
        padding: "2px 9px",
        borderRadius: 6
      }
    }, fl))))), /*#__PURE__*/React.createElement("button", {
      onMouseDown: e => e.preventDefault(),
      onClick: () => setStep(0),
      style: {
        flexShrink: 0,
        padding: "5px 12px",
        borderRadius: 7,
        border: "1.5px solid " + C.g300,
        background: C.g50,
        color: C.g600,
        cursor: "pointer",
        fontSize: 12,
        fontWeight: 700,
        whiteSpace: "nowrap"
      }
    }, "\u4FEE\u6B63"));
  })(), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      background: C.g200,
      gap: "1px"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      width: 110,
      flexShrink: 0,
      background: C.green + "18",
      display: "flex",
      alignItems: "center",
      padding: "8px 12px"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: 7
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      width: 18,
      height: 18,
      borderRadius: 9,
      background: C.green,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontSize: 9,
      fontWeight: 700,
      color: C.white
    }
  }, "\u2713"), /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 12,
      fontWeight: 700,
      color: C.green
    }
  }, "\u6A5F\u5668\u7A2E\u5225"))), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      background: C.white,
      padding: "7px 10px",
      display: "flex",
      alignItems: "center",
      gap: 8
    }
  }, [["indoor", "🏠", "室内機", C.blue], ["outdoor", "🏭", "室外機", C.teal]].map(([mode, icon, label, color]) => {
    const sel = inspectionMode === mode;
    return /*#__PURE__*/React.createElement("button", {
      key: mode,
      onMouseDown: e => e.preventDefault(),
      onClick: () => onSwitchMode && onSwitchMode(mode),
      style: {
        padding: "7px 16px",
        borderRadius: 9,
        border: "2px solid " + (sel ? color : C.g200),
        cursor: "pointer",
        fontWeight: 700,
        fontSize: 13,
        background: sel ? "linear-gradient(135deg," + color + "," + color + "CC)" : C.white,
        color: sel ? C.white : C.g600,
        transition: "all 0.12s",
        whiteSpace: "nowrap"
      }
    }, icon, " ", label, sel ? " ✓" : "");
  }))), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      background: C.g200,
      gap: "1px",
      opacity: s1InspDone ? 1 : 0.45
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      width: 110,
      flexShrink: 0,
      background: s1DevDone ? C.green + "18" : C.g50,
      padding: "10px 12px"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: 7
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      width: 18,
      height: 18,
      borderRadius: 9,
      background: s1DevDone ? C.green : C.g300,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontSize: 9,
      fontWeight: 700,
      color: C.white
    }
  }, s1DevDone ? "✓" : "3"), /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 12,
      fontWeight: 700,
      color: s1DevDone ? C.green : C.g600
    }
  }, "\u6A5F\u5668\u9078\u629E"))), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      background: C.white,
      padding: "8px 10px",
      display: "flex",
      flexDirection: "column",
      gap: 5,
      pointerEvents: s1InspDone ? "auto" : "none"
    }
  }, false && null, devList.length > 0 ? /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 8,
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      position: "relative",
      flex: 1
    }
  }, /*#__PURE__*/React.createElement("input", {
    value: devSearch,
    ref: devSearchRef,
    onChange: e => {
      setDevSearch(e.target.value);
    },
    onFocus: () => setS1Focus("devSearch"),
    readOnly: IS_IPAD,
    inputMode: IS_IPAD ? "none" : undefined,
    placeholder: "\u7BA1\u7406\u756A\u53F7\u30FB\u90E8\u5C4B\u540D\u3067\u691C\u7D22\u2026",
    style: {
      width: "100%",
      padding: "7px 10px 7px 28px",
      borderRadius: 7,
      fontSize: 13,
      border: "1.5px solid " + (s1Focus === "devSearch" ? C.blue : C.g200),
      background: C.inp,
      outline: "none",
      boxSizing: "border-box",
      fontFamily: "inherit"
    }
  }), /*#__PURE__*/React.createElement("span", {
    style: {
      position: "absolute",
      left: 8,
      top: "50%",
      transform: "translateY(-50%)",
      fontSize: 12,
      color: C.blue,
      pointerEvents: "none"
    }
  }, "\uD83D\uDD0D"), devSearch && /*#__PURE__*/React.createElement("button", {
    onMouseDown: e => e.preventDefault(),
    onClick: () => setDevSearch(""),
    style: {
      position: "absolute",
      right: 6,
      top: "50%",
      transform: "translateY(-50%)",
      border: "none",
      background: "none",
      cursor: "pointer",
      fontSize: 14,
      color: C.g400,
      lineHeight: 1
    }
  }, "\u2715")), s1DevDone && /*#__PURE__*/React.createElement("button", {
    onMouseDown: e => e.preventDefault(),
    onClick: () => {
      setForm(p => ({
        ...p,
        floor: "",
        room: "",
        managementNo: "",
        unitNo: ""
      }));
      setDevSearch("");
    },
    style: {
      padding: "5px 10px",
      borderRadius: 7,
      border: "1.5px solid " + C.g200,
      background: C.white,
      color: C.g500,
      cursor: "pointer",
      fontSize: 12,
      fontWeight: 700,
      whiteSpace: "nowrap"
    }
  }, "\u5909\u66F4")), s1DevDone && !devSearch ? /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "8px 12px",
      borderRadius: 9,
      background: C.g50,
      border: "1.5px solid " + C.g200,
      display: "flex",
      flexWrap: "wrap",
      rowGap: 6,
      gap: 14,
      alignItems: "center"
    }
  }, [["階", form.floor], ["部屋名", form.room], ["管理番号", form.managementNo], ["機器番号", form.unitNo]].map(([k, v]) => /*#__PURE__*/React.createElement("div", {
    key: k,
    style: {
      display: "flex",
      alignItems: "baseline",
      gap: 6
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 11,
      fontWeight: 700,
      color: C.green,
      whiteSpace: "nowrap"
    }
  }, "\u2713 ", k), /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 13,
      fontWeight: 700,
      color: C.navy
    }
  }, v || "—"))), (() => {
    const selDev = devList.find(d => d.managementNo === form.managementNo && d.unitNo === form.unitNo);
    const remarksVal = remarksKey ? selDev?._raw?.[remarksKey] : "";
    return remarksVal ? /*#__PURE__*/React.createElement("span", {
      style: {
        flexBasis: "100%",
        fontSize: 11,
        color: C.g500,
        whiteSpace: "pre-wrap",
        wordBreak: "break-word"
      }
    }, "\uD83D\uDCDD ", remarksVal) : null;
  })()) : /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 10,
      flexWrap: "wrap",
      flexShrink: 0,
      alignItems: "center"
    }
  }, [{
    label: "点検前状態",
    dir: preStateSortDir,
    set: setPreStateSortDir
  }, {
    label: "データ入力",
    dir: dataSortDir,
    set: setDataSortDir
  }].map(({
    label,
    dir,
    set
  }) => /*#__PURE__*/React.createElement("div", {
    key: label,
    style: {
      display: "flex",
      alignItems: "center",
      gap: 4
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 11,
      fontWeight: 700,
      color: C.g500,
      whiteSpace: "nowrap"
    }
  }, label, "\uFF1A"), [["done", "入力済"], ["undone", "未入力"]].map(([v, txt]) => /*#__PURE__*/React.createElement("button", {
    key: v,
    onMouseDown: e => e.preventDefault(),
    onClick: () => set(p => p === v ? null : v),
    style: {
      padding: "4px 10px",
      borderRadius: 16,
      border: "1.5px solid " + (dir === v ? C.blue : C.g200),
      background: dir === v ? C.blue : C.white,
      color: dir === v ? C.white : C.g500,
      cursor: "pointer",
      fontSize: 11,
      fontWeight: 700,
      whiteSpace: "nowrap"
    }
  }, txt))))), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 3,
      overflowY: "auto",
      scrollbarGutter: "stable",
      maxHeight: "calc(100vh - 340px)",
      minHeight: 80,
      flexShrink: 0
    }
  }, (() => {
    // 対象階の表示順（sortedFloors=降順 or 昇順）を尊重
    const floorOrder = targetFloors && targetFloors.length > 0 ? targetFloors // セッション画面の選択順（sortedFloorsと同じ降順）
    : null;
    const sorted = [...baseDevList].sort((a, b) => {
      if (dataSortDir) {
        const da = devDataDone(a) ? 1 : 0,
          db = devDataDone(b) ? 1 : 0;
        const cmp = dataSortDir === "undone" ? da - db : db - da;
        if (da !== db) return cmp;
      }
      if (preStateSortDir) {
        const pa = devPreStateDone(a) ? 1 : 0,
          pb = devPreStateDone(b) ? 1 : 0;
        const cmp = preStateSortDir === "undone" ? pa - pb : pb - pa;
        if (pa !== pb) return cmp;
      }
      if (!floorOrder) return a.floor.localeCompare(b.floor) || a.room.localeCompare(b.room);
      const ai = floorOrder.indexOf(a.floor);
      const bi = floorOrder.indexOf(b.floor);
      // 対象階内は降順（sortedFloorsと同じ並び）
      const ar = ai < 0 ? 999 : ai;
      const br = bi < 0 ? 999 : bi;
      if (ar !== br) return ar - br;
      // 同じ階内は部屋名順
      return a.room.localeCompare(b.room);
    });
    const filtered = sorted.filter(d => {
      const n = s => s.replace(/-/g, "").toLowerCase();
      const digitsOnly = s => s.replace(/[^0-9]/g, "");
      if (devSearch) {
        const qRaw = devSearch.trim();
        const isNumericQuery = /^[0-9]+$/.test(qRaw);
        const q = n(qRaw);
        const qDigits = digitsOnly(qRaw);
        // 数字だけの検索語のときは、英字・記号を無視して数字だけで比較する
        // （例："AC-10A-19"→"1019"、"02-01"（リモコン番号）→"0201" のどちらでも見つかる）
        const matches = v => {
          const s = String(v || "");
          return isNumericQuery ? digitsOnly(s).includes(qDigits) : n(s).includes(q);
        };
        // devVisibleColsの全列＋基本4フィールド＋リモコン番号列（非表示でも検索対象に含める）でヒット判定
        const cols = devVisibleCols && devVisibleCols.length > 0 ? devVisibleCols : [];
        const remoteKey = devColumns.find(k => /リモコン|remote/i.test(k));
        const searchCols = remoteKey && !cols.includes(remoteKey) ? [...cols, remoteKey] : cols;
        const rawHit = searchCols.some(col => matches(d._raw ? d._raw[col] : ""));
        const basicHit = matches(d.managementNo) || matches(d.unitNo) || matches(d.room) || matches(d.floor);
        if (!rawHit && !basicHit) return false;
      }
      if (undoneOnly) {
        const hasMeas = records.some(r => r.managementNo === d.managementNo && r.unitNo === d.unitNo && (inspectionMode === "outdoor" ? Object.values(r.checks || {}).some(v => v !== "") : Object.values(r.values).some(v => v !== "")));
        if (hasMeas) return false;
      }
      return true;
    });
    if (filtered.length === 0) return /*#__PURE__*/React.createElement("div", {
      style: {
        fontSize: 12,
        color: C.g400,
        padding: "6px 4px"
      }
    }, "\u8A72\u5F53\u3059\u308B\u6A5F\u5668\u304C\u3042\u308A\u307E\u305B\u3093");
    return filtered.map((dev, i) => {
      const sel = form.managementNo === dev.managementNo && form.unitNo === dev.unitNo;
      const hasMeas = records.some(r => r.managementNo === dev.managementNo && r.unitNo === dev.unitNo && (inspectionMode === "outdoor" ? Object.values(r.checks || {}).some(v => v !== "") : Object.values(r.values).some(v => v !== "")));
      const ngRoom = isNG(dev.floor, dev.room);
      const ngReason = isNGReason(dev.floor, dev.room);
      if (ngRoom) return /*#__PURE__*/React.createElement("div", {
        key: i,
        style: {
          padding: "7px 10px",
          borderRadius: 7,
          border: "1.5px solid " + C.g200,
          background: C.g100,
          display: "flex",
          gap: 6,
          alignItems: "center",
          flexShrink: 0,
          opacity: 0.45,
          userSelect: "none",
          flexWrap: "wrap"
        }
      }, (() => {
        const cols = devVisibleCols && devVisibleCols.length > 0 ? devVisibleCols : devColumns && devColumns.length > 0 ? devColumns : null;
        const remarksVal = remarksKey ? dev._raw?.[remarksKey] : "";
        const remarksStyle = {
          flexBasis: "100%",
          fontSize: 10,
          color: C.g500,
          whiteSpace: "pre-wrap",
          wordBreak: "break-word"
        };
        if (cols) {
          return /*#__PURE__*/React.createElement(React.Fragment, null, cols.map(col => {
            const val = dev._raw ? dev._raw[col] : col === "階" ? dev.floor : col === "部屋名" ? dev.room : col === "管理番号" ? dev.managementNo : col === "機器番号" ? dev.unitNo : "";
            return val ? /*#__PURE__*/React.createElement("span", {
              key: col,
              style: {
                fontSize: 13,
                fontWeight: 700,
                color: C.g400,
                textDecoration: "line-through"
              }
            }, val) : null;
          }), /*#__PURE__*/React.createElement("span", {
            style: {
              fontSize: 10,
              fontWeight: 700,
              color: ngReason === "対象外" ? C.g400 : C.red,
              background: ngReason === "対象外" ? C.g200 : "#FEF2F2",
              padding: "1px 6px",
              borderRadius: 4,
              whiteSpace: "nowrap"
            }
          }, ngReason), remarksVal && /*#__PURE__*/React.createElement("span", {
            style: remarksStyle
          }, "\uD83D\uDCDD ", remarksVal));
        }
        return /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("span", {
          style: {
            fontSize: 13,
            fontWeight: 700,
            flex: 1,
            color: C.g400,
            textDecoration: "line-through"
          }
        }, dev.floor, "\u3000", dev.room), /*#__PURE__*/React.createElement("span", {
          style: {
            fontSize: 12,
            fontFamily: "monospace",
            opacity: 0.6,
            color: C.g400
          }
        }, dev.managementNo, " / ", dev.unitNo), /*#__PURE__*/React.createElement("span", {
          style: {
            fontSize: 10,
            fontWeight: 700,
            color: ngReason === "対象外" ? C.g400 : C.red,
            background: ngReason === "対象外" ? C.g200 : "#FEF2F2",
            padding: "1px 6px",
            borderRadius: 4,
            whiteSpace: "nowrap"
          }
        }, ngReason), remarksVal && /*#__PURE__*/React.createElement("span", {
          style: remarksStyle
        }, "\uD83D\uDCDD ", remarksVal));
      })());
      if (hasMeas && !sel) return /*#__PURE__*/React.createElement("button", {
        key: i,
        onMouseDown: e => e.preventDefault(),
        onClick: () => selectDevice(dev),
        style: {
          padding: "7px 10px",
          borderRadius: 7,
          border: "2px solid " + C.green,
          cursor: "pointer",
          textAlign: "left",
          background: C.green + "10",
          color: C.g800,
          display: "flex",
          gap: 6,
          alignItems: "center",
          flexShrink: 0,
          transition: "all 0.1s",
          flexWrap: "wrap"
        }
      }, (() => {
        const cols = devVisibleCols && devVisibleCols.length > 0 ? devVisibleCols : devColumns && devColumns.length > 0 ? devColumns : null;
        const remarksVal = remarksKey ? dev._raw?.[remarksKey] : "";
        const showExtraRemarks = remarksVal && !(cols && cols.includes(remarksKey));
        const remarksStyle = {
          flexBasis: "100%",
          fontSize: 10,
          fontWeight: 400,
          color: C.g500,
          textAlign: "left",
          whiteSpace: "pre-wrap",
          wordBreak: "break-word"
        };
        if (cols) {
          return /*#__PURE__*/React.createElement(React.Fragment, null, cols.map(col => {
            const val = dev._raw ? dev._raw[col] : col === "階" ? dev.floor : col === "部屋名" ? dev.room : col === "管理番号" ? dev.managementNo : col === "機器番号" ? dev.unitNo : "";
            return val ? /*#__PURE__*/React.createElement("span", {
              key: col,
              style: {
                fontSize: 13,
                fontWeight: 700
              }
            }, val) : null;
          }), /*#__PURE__*/React.createElement("span", {
            style: {
              fontSize: 10,
              fontWeight: 700,
              color: C.green,
              background: C.green + "20",
              padding: "1px 6px",
              borderRadius: 4,
              whiteSpace: "nowrap"
            }
          }, "\u5165\u529B\u6E08"), showExtraRemarks && /*#__PURE__*/React.createElement("span", {
            style: remarksStyle
          }, "\uD83D\uDCDD ", remarksVal));
        }
        return /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("span", {
          style: {
            fontSize: 13,
            fontWeight: 700,
            flex: 1
          }
        }, dev.floor, "\u3000", dev.room), /*#__PURE__*/React.createElement("span", {
          style: {
            fontSize: 12,
            fontFamily: "monospace",
            opacity: 0.8
          }
        }, dev.managementNo, " / ", dev.unitNo), /*#__PURE__*/React.createElement("span", {
          style: {
            fontSize: 10,
            fontWeight: 700,
            color: C.green,
            background: C.green + "20",
            padding: "1px 6px",
            borderRadius: 4,
            whiteSpace: "nowrap"
          }
        }, "\u5165\u529B\u6E08"), showExtraRemarks && /*#__PURE__*/React.createElement("span", {
          style: remarksStyle
        }, "\uD83D\uDCDD ", remarksVal));
      })());
      return /*#__PURE__*/React.createElement("button", {
        key: i,
        onMouseDown: e => e.preventDefault(),
        onClick: () => selectDevice(dev),
        style: {
          padding: "7px 10px",
          borderRadius: 7,
          border: "2px solid " + (sel ? C.blue : C.g200),
          cursor: "pointer",
          textAlign: "left",
          background: sel ? "linear-gradient(135deg," + C.navy + "," + C.blue + ")" : C.white,
          color: sel ? C.white : C.g800,
          display: "flex",
          gap: 6,
          alignItems: "center",
          flexShrink: 0,
          transition: "all 0.1s",
          flexWrap: "wrap"
        }
      }, (() => {
        const cols = devVisibleCols && devVisibleCols.length > 0 ? devVisibleCols : devColumns && devColumns.length > 0 ? devColumns : null;
        const remarksVal = remarksKey ? dev._raw?.[remarksKey] : "";
        const showExtraRemarks = remarksVal && !(cols && cols.includes(remarksKey));
        const remarksStyle = {
          flexBasis: "100%",
          fontSize: 10,
          fontWeight: 400,
          color: sel ? "rgba(255,255,255,0.85)" : C.g500,
          textAlign: "left",
          whiteSpace: "pre-wrap",
          wordBreak: "break-word"
        };
        if (cols) {
          return /*#__PURE__*/React.createElement(React.Fragment, null, cols.map(col => {
            const val = dev._raw ? dev._raw[col] : col === "階" ? dev.floor : col === "部屋名" ? dev.room : col === "管理番号" ? dev.managementNo : col === "機器番号" ? dev.unitNo : "";
            return val ? /*#__PURE__*/React.createElement("span", {
              key: col,
              style: {
                fontSize: 13,
                fontWeight: 700
              }
            }, val) : null;
          }), showExtraRemarks && /*#__PURE__*/React.createElement("span", {
            style: remarksStyle
          }, "\uD83D\uDCDD ", remarksVal));
        }
        return /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("span", {
          style: {
            fontSize: 13,
            fontWeight: 700,
            flex: 1
          }
        }, dev.floor, "\u3000", dev.room), /*#__PURE__*/React.createElement("span", {
          style: {
            fontSize: 12,
            fontFamily: "monospace",
            opacity: 0.8
          }
        }, dev.managementNo, " / ", dev.unitNo), showExtraRemarks && /*#__PURE__*/React.createElement("span", {
          style: remarksStyle
        }, "\uD83D\uDCDD ", remarksVal));
      })());
    });
  })()))) : /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gridTemplateColumns: "1fr 1fr 0.5fr 1.4fr",
      gap: 5
    }
  }, [{
    k: "managementNo",
    l: "管理番号",
    p: "A-001"
  }, {
    k: "unitNo",
    l: "機器番号",
    p: "IDU-001"
  }, {
    k: "floor",
    l: "階",
    p: "1F"
  }, {
    k: "room",
    l: "部屋名",
    p: "事務室"
  }].map(f => /*#__PURE__*/React.createElement("div", {
    key: f.k
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 10,
      fontWeight: 700,
      color: C.g500,
      marginBottom: 3
    }
  }, f.l), /*#__PURE__*/React.createElement("input", {
    value: form[f.k],
    placeholder: f.p,
    onChange: e => setInfo(f.k, e.target.value),
    style: {
      width: "100%",
      padding: "7px 9px",
      borderRadius: 7,
      fontSize: 13,
      border: "1.5px solid " + (form[f.k] ? C.green : C.g200),
      background: C.inp,
      outline: "none",
      boxSizing: "border-box",
      fontFamily: "inherit"
    }
  }))))))) :
  /*#__PURE__*/
  /* ── 折りたたみ済み：1行目=点検日・点検者・建物・階（1つの枠、修正ボタンも1つ）／2行目=機器選択（部屋名・管理番号・機器番号、別の修正ボタン） ── */
  React.createElement(React.Fragment, null, (() => {
    const bKey = devColumns.find(k => /建物|building|棟|ビル/i.test(k));
    const sessionBuildings = sessionInfo?.selectedBuildings || [];
    const showBuildings = bKey ? sessionBuildings.length > 0 ? sessionBuildings : [...new Set(devList.map(d => d._raw?.[bKey]).filter(Boolean))] : [];
    return /*#__PURE__*/React.createElement("div", {
      style: {
        background: C.white,
        borderRadius: 12,
        padding: "10px 12px",
        boxShadow: "0 1px 6px rgba(0,0,0,0.06)",
        display: "flex",
        alignItems: "center",
        gap: 12
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        flex: 1,
        display: "flex",
        flexDirection: "column",
        gap: 6,
        minWidth: 0
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        gap: 16,
        rowGap: 6,
        flexWrap: "wrap"
      }
    }, [["点検日", form.inspectionDate], ["点検者", form.inspector]].map(([k, v]) => /*#__PURE__*/React.createElement("div", {
      key: k,
      style: {
        display: "flex",
        alignItems: "baseline",
        gap: 6
      }
    }, /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 11,
        fontWeight: 700,
        color: C.green,
        whiteSpace: "nowrap"
      }
    }, "\u2713 ", k), /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 14,
        fontWeight: 700,
        color: C.navy
      }
    }, v || "—")))), /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        gap: 16,
        rowGap: 6,
        flexWrap: "wrap"
      }
    }, showBuildings.length > 0 && /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        alignItems: "baseline",
        gap: 6
      }
    }, /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 11,
        fontWeight: 700,
        color: C.green,
        whiteSpace: "nowrap"
      }
    }, "\u2713 \u5EFA\u7269"), /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 14,
        fontWeight: 700,
        color: C.navy
      }
    }, showBuildings.join("・"))), /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        alignItems: "baseline",
        gap: 6
      }
    }, /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 11,
        fontWeight: 700,
        color: C.green,
        whiteSpace: "nowrap"
      }
    }, "\u2713 \u968E"), /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 14,
        fontWeight: 700,
        color: C.navy
      }
    }, form.floor || "—")))), /*#__PURE__*/React.createElement("button", {
      onMouseDown: e => e.preventDefault(),
      onClick: () => setStep(0),
      style: {
        flexShrink: 0,
        padding: "5px 12px",
        borderRadius: 7,
        border: "1.5px solid " + C.g300,
        background: C.g50,
        color: C.g600,
        cursor: "pointer",
        fontSize: 12,
        fontWeight: 700,
        whiteSpace: "nowrap"
      }
    }, "\u4FEE\u6B63"));
  })(), /*#__PURE__*/React.createElement("div", {
    style: {
      background: C.white,
      borderRadius: 12,
      padding: "10px 12px",
      boxShadow: "0 1px 6px rgba(0,0,0,0.06)",
      marginTop: 6
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 16,
      rowGap: 8,
      flexWrap: "wrap",
      alignItems: "center"
    }
  }, [["部屋名", form.room], ["管理番号", form.managementNo], ["機器番号", form.unitNo]].map(([k, v]) => /*#__PURE__*/React.createElement("div", {
    key: k,
    style: {
      display: "flex",
      alignItems: "baseline",
      gap: 6
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 11,
      fontWeight: 700,
      color: C.green,
      whiteSpace: "nowrap"
    }
  }, "\u2713 ", k), /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 14,
      fontWeight: 700,
      color: C.navy
    }
  }, v || "—"))), /*#__PURE__*/React.createElement("button", {
    onMouseDown: e => e.preventDefault(),
    onClick: () => {
      setSummaryExpanded(true);
      setForm(p => ({
        ...p,
        floor: "",
        room: "",
        managementNo: "",
        unitNo: ""
      }));
      setDevSearch("");
    },
    style: {
      marginLeft: "auto",
      padding: "5px 12px",
      borderRadius: 7,
      border: "1.5px solid " + C.g300,
      background: C.g50,
      color: C.g600,
      cursor: "pointer",
      fontSize: 12,
      fontWeight: 700,
      whiteSpace: "nowrap"
    }
  }, "\u6A5F\u5668\u9078\u629E"))));
  const seqIndoor = focusSeq || [];
  const aiIndoor = seqIndoor.findIndex(f => f.code === activeCode);
  const activeIsCheckIndoor = !!activeCode && !!isCheckCode && isCheckCode(activeCode);
  const ciFieldsIndoor = (checkFields || []).filter(f => f.group === "check_in");
  const activeCheckFieldIndoor = activeIsCheckIndoor ? ciFieldsIndoor.find(f => f.code === activeCode) : null;
  return /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      overflow: "hidden",
      display: "flex",
      flexDirection: "row",
      gap: 0
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      overflow: "hidden",
      display: "flex",
      flexDirection: "column",
      padding: "10px 12px 12px",
      gap: 0
    }
  }, /*#__PURE__*/React.createElement("div", {
    ref: outListRef,
    style: {
      flex: 1,
      display: "flex",
      flexDirection: "column",
      gap: "1px",
      background: C.g200,
      borderRadius: 14,
      overflow: "auto",
      scrollbarGutter: "stable",
      boxShadow: "0 1px 8px rgba(0,0,0,0.08)",
      position: "relative"
    }
  }, topSummary, /*#__PURE__*/React.createElement(React.Fragment, null, inspectionMode === "outdoor" ? /*#__PURE__*/React.createElement("div", {
    style: {
      opacity: s1DevDone && !devSearch ? 1 : 0.45,
      pointerEvents: s1DevDone && !devSearch ? "auto" : "none"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "12px 12px 4px"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12,
      fontWeight: 800,
      color: C.teal,
      marginBottom: 8
    }
  }, "\uD83C\uDFED \u5BA4\u5916\u6A5F\u70B9\u691C\u30C1\u30A7\u30C3\u30AF"), /*#__PURE__*/React.createElement("table", {
    style: {
      width: "100%",
      borderCollapse: "collapse"
    }
  }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("th", {
    style: {
      width: 90,
      padding: "6px 8px",
      textAlign: "left",
      fontSize: 10,
      fontWeight: 700,
      color: C.g500,
      borderBottom: "2px solid " + C.g200
    }
  }, "\u9805\u76EE"), /*#__PURE__*/React.createElement("th", {
    style: {
      padding: "6px 8px",
      textAlign: "left",
      fontSize: 10,
      fontWeight: 700,
      color: C.g500,
      borderBottom: "2px solid " + C.g200
    }
  }, "\u70B9\u691C\u5185\u5BB9"), /*#__PURE__*/React.createElement("th", {
    style: {
      width: 110,
      padding: "6px 8px",
      textAlign: "center",
      fontSize: 10,
      fontWeight: 700,
      color: C.g500,
      borderBottom: "2px solid " + C.g200
    }
  }, "\u25CB\xD7"))), /*#__PURE__*/React.createElement("tbody", null, outFields.map((f, i) => {
    const val = form.checks?.[f.code] || "";
    const act = s1Focus === f.code;
    return /*#__PURE__*/React.createElement("tr", {
      key: f.code,
      ref: el => {
        outRowRefs.current[f.code] = el;
      },
      onClick: () => focusCheckRow(f.code),
      style: {
        background: act ? C.blue + "0C" : i % 2 === 0 ? C.white : C.g50,
        boxShadow: act ? "inset 0 0 0 2px " + C.blue : "none",
        cursor: "pointer",
        transition: "box-shadow 0.1s"
      }
    }, /*#__PURE__*/React.createElement("td", {
      style: {
        padding: "9px 8px",
        fontSize: 11,
        fontWeight: 700,
        color: C.teal,
        borderBottom: "1px solid " + C.g100,
        verticalAlign: "middle"
      }
    }, f.category), /*#__PURE__*/React.createElement("td", {
      style: {
        padding: "9px 8px",
        fontSize: 14,
        color: C.g700,
        borderBottom: "1px solid " + C.g100,
        verticalAlign: "middle"
      }
    }, f.label), /*#__PURE__*/React.createElement("td", {
      style: {
        padding: "6px 8px",
        borderBottom: "1px solid " + C.g100,
        verticalAlign: "middle"
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        gap: 6,
        justifyContent: "center"
      },
      onClick: e => e.stopPropagation()
    }, ["○", "×"].map(v => /*#__PURE__*/React.createElement("button", {
      key: v,
      onClick: () => advanceCheck(f.code, v),
      style: {
        width: 48,
        height: 40,
        borderRadius: 9,
        border: "2px solid " + (val === v ? v === "○" ? C.green : C.red : C.g200),
        background: val === v ? v === "○" ? C.green : C.red : C.white,
        color: val === v ? C.white : C.g400,
        fontWeight: 800,
        fontSize: 19,
        cursor: "pointer",
        transition: "all 0.1s"
      }
    }, v)))));
  }))))) : s1DevDone && !devSearch ? preStateOpen ? /*#__PURE__*/React.createElement(React.Fragment, null, (() => {
    const prevRec = records.find(r => r.managementNo === form.managementNo && r.unitNo === form.unitNo && (r.preOperation || r.preMode || r.preWind || r.preSetTemp));
    if (!prevRec) return null;
    const prevSummary = [prevRec.preOperation, prevRec.preMode, prevRec.preWind, prevRec.preSetTemp ? prevRec.preSetTemp + "°C" : null].filter(Boolean).join(" ・ ");
    return /*#__PURE__*/React.createElement("div", {
      style: {
        padding: "6px 10px",
        marginBottom: 4,
        background: "#FFF7ED",
        border: "1.5px solid #F59E0B",
        borderRadius: 8,
        fontSize: 11,
        color: "#92400E",
        fontWeight: 700
      }
    }, "\uD83D\uDCCB \u4E00\u6B21\u4FDD\u5B58\u6E08\u307F\uFF1A\u524D\u56DE\u306E\u70B9\u691C\u524D\u72B6\u614B\u300C", prevSummary, "\u300D");
  })(), /*#__PURE__*/React.createElement("div", {
    style: {
      position: "relative",
      display: "flex",
      background: C.g200,
      gap: "1px"
    }
  }, /*#__PURE__*/React.createElement("button", {
    onMouseDown: e => e.preventDefault(),
    onClick: () => {
      setSummaryExpanded(true);
      setForm(p => ({
        ...p,
        floor: "",
        room: "",
        managementNo: "",
        unitNo: ""
      }));
      setDevSearch("");
    },
    style: {
      position: "absolute",
      top: 6,
      right: 6,
      zIndex: 1,
      padding: "4px 10px",
      borderRadius: 7,
      border: "1.5px solid " + C.g300,
      background: C.white,
      color: C.g600,
      cursor: "pointer",
      fontSize: 11,
      fontWeight: 700,
      whiteSpace: "nowrap"
    }
  }, "\u2190 \u623B\u308B"), /*#__PURE__*/React.createElement("div", {
    style: {
      width: 110,
      flexShrink: 0,
      background: C.teal + "10",
      padding: "10px 12px"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: 7
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      width: 18,
      height: 18,
      borderRadius: 9,
      background: C.teal + "50",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontSize: 9,
      fontWeight: 700,
      color: C.teal
    }
  }, "4"), /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 12,
      fontWeight: 700,
      color: C.teal
    }
  }, "\u70B9\u691C\u524D\u72B6\u614B")), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 10,
      color: C.g400,
      marginTop: 2,
      paddingLeft: 25
    }
  }, "\u4EFB\u610F")), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      background: C.teal + "05",
      padding: "12px 12px",
      pointerEvents: s1DevDone ? "auto" : "none",
      display: "flex",
      gap: 8,
      alignItems: "flex-start",
      borderTop: "1px solid " + C.teal + "20",
      flexWrap: "nowrap",
      overflowX: "auto"
    }
  }, [{
    label: "運転",
    items: [["ON", "🟢 ON", C.blue], ["OFF", "⭕ OFF", C.blue]],
    key: "preOperation"
  }, {
    label: "運転モード",
    items: [["冷房", "❄️ 冷房", C.teal], ["暖房", "🔥 暖房", C.teal], ["送風", "💨 送風", C.teal], ["除湿", "💧 除湿", C.teal]],
    key: "preMode"
  }, {
    label: "風量",
    items: [["自動", "🔄 自動", "#7C3AED"], ["弱", "💨 弱", "#7C3AED"], ["強", "💨 強", "#7C3AED"], ["急風", "🌪️ 急風", "#7C3AED"]],
    key: "preWind"
  }].map(({
    label,
    items,
    key
  }) => /*#__PURE__*/React.createElement("div", {
    key: key,
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 4,
      flexShrink: 0
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 10,
      fontWeight: 700,
      color: C.g500,
      marginBottom: 2
    }
  }, label), items.map(([v, txt, col]) => {
    const sel = form[key] === v;
    return /*#__PURE__*/React.createElement("button", {
      key: v,
      onClick: () => setInfo(key, sel ? "" : v),
      style: {
        padding: "10px 16px",
        borderRadius: 10,
        border: "2px solid " + (sel ? col : C.g200),
        cursor: "pointer",
        fontWeight: 700,
        fontSize: 15,
        background: sel ? "linear-gradient(135deg," + col + "," + col + "BB)" : C.white,
        color: sel ? C.white : C.g600,
        transition: "all 0.12s",
        textAlign: "center",
        whiteSpace: "nowrap"
      }
    }, txt);
  }))), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 4,
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 10,
      fontWeight: 700,
      color: C.g500,
      marginBottom: 2
    }
  }, "\u8A2D\u5B9A\u6E29\u5EA6"), /*#__PURE__*/React.createElement("button", {
    onClick: () => setInfo("preSetTemp", Math.min(30, parseFloat(form.preSetTemp || 20) + 0.5).toFixed(1)),
    style: {
      width: 48,
      height: 40,
      borderRadius: 10,
      border: "1.5px solid " + C.g200,
      cursor: "pointer",
      fontSize: 20,
      fontWeight: 700,
      background: C.white,
      color: C.g600
    }
  }, "\uFF0B"), /*#__PURE__*/React.createElement("div", {
    onClick: () => setS1Focus(s1Focus === "preSetTemp" ? null : "preSetTemp"),
    style: {
      width: 80,
      padding: "8px 6px",
      borderRadius: 10,
      cursor: "pointer",
      border: "2px solid " + (s1Focus === "preSetTemp" ? C.blue : form.preSetTemp ? C.green : C.g200),
      background: s1Focus === "preSetTemp" ? "#EFF6FF" : C.inp,
      textAlign: "center",
      fontFamily: "monospace",
      fontSize: 22,
      fontWeight: 800,
      color: form.preSetTemp ? C.navy : C.g300,
      transition: "all 0.12s"
    }
  }, form.preSetTemp ? parseFloat(form.preSetTemp).toFixed(1) : "—", /*#__PURE__*/React.createElement("br", null), /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 9,
      fontWeight: 400,
      color: C.g400
    }
  }, "\xB0C")), /*#__PURE__*/React.createElement("button", {
    onClick: () => setInfo("preSetTemp", Math.max(16, parseFloat(form.preSetTemp || 20) - 0.5).toFixed(1)),
    style: {
      width: 48,
      height: 40,
      borderRadius: 10,
      border: "1.5px solid " + C.g200,
      cursor: "pointer",
      fontSize: 20,
      fontWeight: 700,
      background: C.white,
      color: C.g600
    }
  }, "\uFF0D")))), (() => {
    const preStateSet = !!(form.preOperation || form.preMode || form.preWind || form.preSetTemp);
    return /*#__PURE__*/React.createElement(React.Fragment, null, noRemote && /*#__PURE__*/React.createElement("div", {
      style: {
        marginTop: 6,
        padding: "8px 10px",
        background: "#FFF7ED",
        border: "1.5px solid #F59E0B",
        borderRadius: 8,
        fontSize: 12,
        fontWeight: 700,
        color: "#92400E"
      }
    }, "\u26A0\uFE0F \u3053\u306E\u5BA4\u5185\u6A5F\u306F\u30EA\u30E2\u30B3\u30F3\u304C\u3042\u308A\u307E\u305B\u3093\u3002\u30B5\u30FC\u30DF\u30B9\u30BF\u8A08\u3067\u5B9F\u6E2C\u5024\u3092\u5165\u529B\u3057\u3066\u304F\u3060\u3055\u3044\u3002"), otherRemark && /*#__PURE__*/React.createElement("div", {
      style: {
        marginTop: 6,
        padding: "8px 10px",
        background: "#FFF7ED",
        border: "1.5px solid #F59E0B",
        borderRadius: 8,
        fontSize: 12,
        fontWeight: 700,
        color: "#92400E"
      }
    }, "\uD83D\uDCDD \u5099\u8003\uFF1A", otherRemark), /*#__PURE__*/React.createElement("button", {
      onClick: () => {
        setPreStateConfirmedFor(deviceKey);
        if (focusSeq && focusSeq.length > 0) {
          onRowClick && onRowClick(focusSeq[0]);
        }
      },
      style: {
        marginTop: 6,
        display: "block",
        width: "100%",
        boxSizing: "border-box",
        padding: "10px",
        borderRadius: 10,
        border: "none",
        cursor: "pointer",
        fontWeight: 800,
        fontSize: 14,
        background: preStateSet ? "linear-gradient(135deg," + C.green + ",#047857)" : C.g200,
        color: preStateSet ? C.white : C.g600
      }
    }, preStateSet ? "💾 一次保存" : "⏭️ スキップ"));
  })()) :
  /*#__PURE__*/
  /* ── 折りたたみ済み：点検日等のバーと同じ書式（白背景＋✓アイコン）で1行表示（タップで再展開） ── */
  React.createElement("div", {
    style: {
      width: "100%",
      boxSizing: "border-box",
      background: C.white,
      borderRadius: 12,
      padding: "10px 12px",
      boxShadow: "0 1px 6px rgba(0,0,0,0.06)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 16,
      rowGap: 8,
      flexWrap: "wrap",
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "baseline",
      gap: 6
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 11,
      fontWeight: 700,
      color: C.green,
      whiteSpace: "nowrap"
    }
  }, "\u2713 \u70B9\u691C\u524D\u72B6\u614B"), /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 13,
      fontWeight: 700,
      color: C.navy
    }
  }, [form.preOperation, form.preMode, form.preWind, form.preSetTemp ? form.preSetTemp + "°C" : null].filter(Boolean).join(" ・ ") || "未入力（スキップ済み）")), /*#__PURE__*/React.createElement("button", {
    onMouseDown: e => e.preventDefault(),
    onClick: () => setPreStateConfirmedFor(null),
    style: {
      marginLeft: "auto",
      padding: "5px 12px",
      borderRadius: 7,
      border: "1.5px solid " + C.g300,
      background: C.g50,
      color: C.g600,
      cursor: "pointer",
      fontSize: 12,
      fontWeight: 700,
      whiteSpace: "nowrap"
    }
  }, "\u4FEE\u6B63"))) : null, inspectionMode === "indoor" && s1DevDone && !devSearch && !preStateOpen && noRemote && /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 6,
      padding: "8px 10px",
      background: "#FFF7ED",
      border: "1.5px solid #F59E0B",
      borderRadius: 8,
      fontSize: 12,
      fontWeight: 700,
      color: "#92400E"
    }
  }, "\u26A0\uFE0F \u3053\u306E\u5BA4\u5185\u6A5F\u306F\u30EA\u30E2\u30B3\u30F3\u304C\u3042\u308A\u307E\u305B\u3093\u3002\u30B5\u30FC\u30DF\u30B9\u30BF\u8A08\u3067\u5B9F\u6E2C\u5024\u3092\u5165\u529B\u3057\u3066\u304F\u3060\u3055\u3044\u3002"), inspectionMode === "indoor" && s1DevDone && !devSearch && !preStateOpen && otherRemark && /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 6,
      padding: "8px 10px",
      background: "#FFF7ED",
      border: "1.5px solid #F59E0B",
      borderRadius: 8,
      fontSize: 12,
      fontWeight: 700,
      color: "#92400E"
    }
  }, "\uD83D\uDCDD \u5099\u8003\uFF1A", otherRemark), inspectionMode === "indoor" && s1DevDone && !devSearch && /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      display: "flex",
      flexDirection: "column",
      overflow: "hidden",
      opacity: preStateOpen ? 0.4 : 1,
      pointerEvents: preStateOpen ? "none" : "auto",
      transition: "opacity 0.15s"
    }
  }, /*#__PURE__*/React.createElement(Step2View, {
    hideHeader: true,
    hideNumpad: true,
    form: form,
    setInfo: setInfo,
    handleSave: handleSave,
    setStep: () => setSummaryExpanded(true),
    visIn: visIn,
    visOut: visOut,
    visFields: visFields,
    activeCode: activeCode,
    numDisp: numDisp,
    limits: limits,
    onPress: onPress,
    onConfirm: onConfirm,
    onRowClick: onRowClick,
    moveActive: moveActive,
    rowRefs: rowRefs,
    listRef: listRef,
    complete: complete,
    missing: missing,
    editIdx: editIdx,
    ALL_FIELDS: ALL_FIELDS,
    vis: vis,
    isAbn: isAbn,
    setCheck: setCheck,
    checkFields: checkFields,
    inspectionMode: inspectionMode,
    focusSeq: focusSeq,
    isCheckCode: isCheckCode,
    setCheckAndAdvance: setCheckAndAdvance,
    outdoorLocked: outdoorLocked,
    requestUnlockOutdoor: requestUnlockOutdoor
  })))), inspectionMode === "outdoor" && /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 10,
      background: C.white,
      borderRadius: 12,
      padding: "10px 12px",
      display: "flex",
      gap: 10,
      alignItems: "stretch",
      boxShadow: "0 1px 8px rgba(0,0,0,0.08)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      display: "flex",
      flexDirection: "column",
      gap: 4
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12,
      fontWeight: 700,
      color: C.g500
    }
  }, "\uD83D\uDCDD \u5099\u8003\u30FB\u7279\u8A18\u4E8B\u9805"), /*#__PURE__*/React.createElement("textarea", {
    value: form.remarks,
    onChange: e => setInfo("remarks", e.target.value),
    placeholder: "\u7570\u5E38\u7B87\u6240\u3001\u7279\u8A18\u4E8B\u9805\u306A\u3069...",
    style: {
      flex: 1,
      width: "100%",
      padding: "10px 12px",
      borderRadius: 9,
      fontSize: 14,
      border: "1.5px solid " + C.g200,
      background: C.inp,
      outline: "none",
      boxSizing: "border-box",
      fontFamily: "inherit",
      resize: "none",
      minHeight: 100,
      lineHeight: 1.5
    }
  }), IS_IPAD && /*#__PURE__*/React.createElement(VoiceMemoPanel, {
    scope: "record",
    recordId: form.id
  })))), inspectionMode === "indoor" && s1DevDone && !devSearch ? /*#__PURE__*/React.createElement("div", {
    style: {
      opacity: preStateOpen ? 0.4 : 1,
      pointerEvents: preStateOpen ? "none" : "auto",
      transition: "opacity 0.15s",
      display: "flex"
    }
  }, /*#__PURE__*/React.createElement(Numpad, {
    mode: activeIsCheckIndoor ? "check" : "numeric",
    display: numDisp,
    onPress: onPress,
    onConfirm: onConfirm,
    canConfirm: !!activeCode && !activeIsCheckIndoor && numDisp !== "",
    checkLabel: activeCheckFieldIndoor?.label,
    checkCategory: activeCheckFieldIndoor?.category,
    checkValue: form.checks?.[activeCode] || "",
    onCheckPress: v => setCheckAndAdvance && setCheckAndAdvance(activeCode, v),
    onPrev: () => moveActive(-1),
    onNext: () => moveActive(1),
    canPrev: aiIndoor > 0,
    canNext: aiIndoor >= 0,
    onSave: () => handleSave && handleSave("next"),
    saveComplete: complete,
    saveMissing: missing
  })) : /*#__PURE__*/React.createElement(React.Fragment, null, inspectionMode === "outdoor" ? /*#__PURE__*/React.createElement("div", {
    style: {
      opacity: s1DevDone && !devSearch ? 1 : 0.45,
      pointerEvents: s1DevDone && !devSearch ? "auto" : "none",
      transition: "opacity 0.15s",
      display: "flex"
    }
  }, /*#__PURE__*/React.createElement(Numpad, {
    mode: "check",
    checkLabel: focusedCheckField?.label,
    checkCategory: focusedCheckField?.category,
    checkValue: form.checks?.[s1Focus] || "",
    onCheckPress: v => advanceCheck(s1Focus, v),
    onPrev: () => {
      const idx = outFields.findIndex(f => f.code === s1Focus);
      if (idx > 0) focusCheckRow(outFields[idx - 1].code);
    },
    onNext: () => {
      const idx = outFields.findIndex(f => f.code === s1Focus);
      const next = outFields[idx + 1];
      if (next) focusCheckRow(next.code);
    },
    canPrev: outFields.findIndex(f => f.code === s1Focus) > 0,
    canNext: (() => {
      const idx = outFields.findIndex(f => f.code === s1Focus);
      return idx >= 0 && idx < outFields.length - 1;
    })(),
    onSave: () => handleSave && handleSave("next"),
    saveComplete: complete,
    saveMissing: missing
  })) : /*#__PURE__*/React.createElement("div", {
    style: {
      width: 208,
      flexShrink: 0,
      display: "flex",
      flexDirection: "column",
      background: C.g50,
      borderLeft: "2px solid " + C.g200,
      padding: "10px 8px",
      gap: 6
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: C.white,
      borderRadius: 10,
      padding: "8px 10px",
      border: "2px solid " + (isCheckFocus ? C.green : s1Focus ? C.blue : C.g200),
      minHeight: 52,
      display: "flex",
      flexDirection: "column",
      justifyContent: "center",
      transition: "border-color 0.15s"
    }
  }, isCheckFocus ? /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 10,
      fontWeight: 700,
      color: "#059669",
      letterSpacing: "0.04em"
    }
  }, focusedCheckField?.category || ""), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 14,
      fontWeight: 800,
      color: C.navy,
      lineHeight: 1.3
    }
  }, focusedCheckField?.label || "—")) : /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 10,
      fontWeight: 700,
      color: C.g400,
      letterSpacing: "0.04em"
    }
  }, s1Focus === "preSetTemp" ? "設定温度を入力中" : s1Focus === "devSearch" ? "検索ワードを入力" : "検索窓または設定温度をタップ"), /*#__PURE__*/React.createElement("div", {
    style: {
      fontFamily: "monospace",
      fontSize: 22,
      fontWeight: 800,
      textAlign: "right",
      lineHeight: 1.2,
      color: s1Focus ? C.navy : C.g300,
      overflow: "hidden",
      textOverflow: "ellipsis",
      whiteSpace: "nowrap"
    }
  }, s1Focus === "preSetTemp" ? form.preSetTemp || /*#__PURE__*/React.createElement("span", {
    style: {
      color: C.g200
    }
  }, "\u2014") : s1Focus === "devSearch" ? devSearch || /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 13,
      color: C.g300
    }
  }, "\u5165\u529B\u3057\u3066\u304F\u3060\u3055\u3044") : /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 13
    }
  }, "\u2014")))), isCheckFocus ? /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 8,
      height: 226
    }
  }, ["○", "×"].map(v => {
    const val = form.checks?.[s1Focus] || "";
    return /*#__PURE__*/React.createElement("button", {
      key: v,
      onClick: () => advanceCheck(s1Focus, v),
      style: {
        flex: 1,
        borderRadius: 14,
        border: "2px solid " + (val === v ? v === "○" ? C.green : C.red : C.g200),
        background: val === v ? v === "○" ? C.green : C.red : C.white,
        color: val === v ? C.white : C.g400,
        fontWeight: 800,
        fontSize: 44,
        cursor: "pointer",
        transition: "all 0.1s",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        boxShadow: "0 2px 6px rgba(0,0,0,0.10)"
      }
    }, v);
  })) : [[7, 8, 9], [4, 5, 6], [1, 2, 3], [0, "."]].map((row, ri) => /*#__PURE__*/React.createElement("div", {
    key: ri,
    style: {
      display: "flex",
      gap: 5
    }
  }, row.map(k => /*#__PURE__*/React.createElement("button", {
    key: k,
    disabled: !s1Focus,
    onMouseDown: e => e.preventDefault(),
    onClick: () => {
      if (s1Focus === "preSetTemp") {
        const prev = String(form.preSetTemp || "");
        let next;
        if (k === ".") {
          next = prev.includes(".") ? prev : (prev || "0") + ".";
        } else if (prev === "0") {
          next = String(k);
        } else {
          next = prev.length < 4 ? prev + k : prev;
        }
        setInfo("preSetTemp", next);
      } else if (s1Focus === "devSearch") {
        setDevSearch(p => p + String(k));
      }
    },
    style: {
      flex: ri === 3 && k === 0 ? 2 : 1,
      height: 52,
      borderRadius: 10,
      border: "none",
      cursor: s1Focus ? "pointer" : "default",
      fontWeight: 800,
      fontFamily: "monospace",
      fontSize: 22,
      background: s1Focus ? C.white : C.g100,
      color: s1Focus ? C.g800 : C.g300,
      boxShadow: s1Focus ? "0 2px 5px rgba(0,0,0,0.09)" : "none",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      transition: "background 0.1s"
    }
  }, k)))), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 5
    }
  }, /*#__PURE__*/React.createElement("button", {
    disabled: !s1Focus || isCheckFocus,
    onMouseDown: e => e.preventDefault(),
    onClick: () => {
      if (s1Focus === "preSetTemp") setInfo("preSetTemp", String(form.preSetTemp || "").slice(0, -1));else if (s1Focus === "devSearch") setDevSearch(p => p.slice(0, -1));
    },
    style: {
      flex: 1,
      height: 48,
      borderRadius: 10,
      border: "none",
      cursor: s1Focus && !isCheckFocus ? "pointer" : "default",
      fontSize: 18,
      fontWeight: 800,
      background: s1Focus && !isCheckFocus ? C.g200 : C.g100,
      color: s1Focus && !isCheckFocus ? C.g600 : C.g300,
      display: "flex",
      alignItems: "center",
      justifyContent: "center"
    }
  }, "\u232B"), /*#__PURE__*/React.createElement("button", {
    disabled: !s1Focus,
    onMouseDown: e => e.preventDefault(),
    onClick: () => {
      if (s1Focus === "preSetTemp") {
        setInfo("preSetTemp", "");
        setS1Focus(null);
      } else if (s1Focus === "devSearch") {
        setDevSearch("");
        setS1Focus(null);
      } else if (isCheckFocus) {
        setCheck && setCheck(s1Focus, "");
      }
    },
    style: {
      flex: 1,
      height: 48,
      borderRadius: 10,
      border: "none",
      cursor: s1Focus ? "pointer" : "default",
      fontSize: 14,
      fontWeight: 800,
      background: s1Focus ? "#FEF2F2" : C.g100,
      color: s1Focus ? C.red : C.g300,
      display: "flex",
      alignItems: "center",
      justifyContent: "center"
    }
  }, "CLR")), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: "auto",
      padding: "4px",
      fontSize: 11,
      color: C.g400,
      textAlign: "center",
      lineHeight: 1.6
    }
  }, isCheckFocus ? /*#__PURE__*/React.createElement(React.Fragment, null, "\u30C1\u30A7\u30C3\u30AF\u9805\u76EE\u3092\u30BF\u30C3\u30D7", /*#__PURE__*/React.createElement("br", null), "\u3059\u308B\u3068\u9023\u7D9A\u5165\u529B\u3067\u304D\u307E\u3059") : /*#__PURE__*/React.createElement(React.Fragment, null, "\u691C\u7D22\u7A93\u307E\u305F\u306F", /*#__PURE__*/React.createElement("br", null), "\u8A2D\u5B9A\u6E29\u5EA6\u3092\u30BF\u30C3\u30D7")))));
}
function ACInspectionApp() {
  const [view, setView] = useState("form");
  const [step, setStep] = useState(0); // 0=点検エリア確認 1=機器選択 2=測定データ
  const [undoneOnly, setUndoneOnly] = useState(false); // 未入力分のみ表示
  const [sessionInfo, setSessionInfo] = useState(() => {
    try {
      const saved = localStorage.getItem("acSessionInfo");
      return saved ? JSON.parse(saved) : null;
    } catch (e) {
      return null;
    }
  }); // {date, inspector, targetFloors:[], roomAccess:{}}
  const [currentRound, setCurrentRoundState] = useState(() => {
    try {
      const saved = localStorage.getItem("acCurrentRound");
      return saved ? JSON.parse(saved) : null;
    } catch (e) {
      return null;
    }
  }); // {id:"YYYY-MM", label:"YYYY年M月点検分"} | null（未選択＝セレクター画面を表示）
  const setCurrentRound = r => {
    setCurrentRoundState(r);
    try {
      r ? localStorage.setItem("acCurrentRound", JSON.stringify(r)) : localStorage.removeItem("acCurrentRound");
    } catch (e) {}
  };
  const [showRoundSwitcher, setShowRoundSwitcher] = useState(false); // ヘッダーからの切替モーダル表示フラグ
  const [form, setForm] = useState(emptyForm());
  const [inspectionMode, setInspectionMode] = useState("indoor"); // "indoor" | "outdoor"
  const [indoorRecords, setIndoorRecords] = useState([]);
  const [outdoorRecords, setOutdoorRecords] = useState([]);
  // 現在のモードに応じたrecords/setRecords
  const records = inspectionMode === "indoor" ? indoorRecords : outdoorRecords;
  const setRecords = inspectionMode === "indoor" ? setIndoorRecords : setOutdoorRecords;
  const [editIdx, setEditIdx] = useState(null);
  const [flash, setFlash] = useState("");
  const [outdoorUnlockedFor, setOutdoorUnlockedFor] = useState(null); // 「室外機データは入力済」確認後にロック解除した機器キー（managementNo|unitNo）

  const [limits, setLimits] = useState(() => lsGet(LS_KEYS.limits, defLim()));
  const [tmpLim, setTmpLim] = useState(defLim());
  // 正常値範囲設定：点検データ入力画面と同じ「テンキーで連続入力」操作用state
  const [limActive, setLimActive] = useState(null); // {code,part:"min"|"max"} | null
  const [limNumDisp, setLimNumDisp] = useState("");
  const limOvr = useRef(true);
  const limFocus = (code, part) => {
    setLimActive({
      code,
      part
    });
    setLimNumDisp(tmpLim[code]?.[part] || "");
    limOvr.current = true;
  };
  const limOnPress = key => {
    setLimNumDisp(prev => {
      if (limOvr.current) {
        limOvr.current = false;
        return key === "." ? "0." : String(key);
      }
      if (key === ".") {
        if (prev.includes(".")) return prev;
        return (prev || "0") + ".";
      }
      if (prev === "0") return String(key);
      return prev + key;
    });
  };
  const limMove = dir => {
    if (!limActive) {
      if (LIM_SEQ.length > 0) limFocus(LIM_SEQ[0].code, LIM_SEQ[0].part);
      return;
    }
    if (limNumDisp !== "") setTmpLim(p => ({
      ...p,
      [limActive.code]: {
        ...p[limActive.code],
        [limActive.part]: limNumDisp
      }
    }));
    const idx = LIM_SEQ.findIndex(s => s.code === limActive.code && s.part === limActive.part);
    const next = LIM_SEQ[idx + dir];
    if (next) {
      limFocus(next.code, next.part);
    } else {
      setLimActive(null);
      setLimNumDisp("");
    }
  };
  const limIdx = limActive ? LIM_SEQ.findIndex(s => s.code === limActive.code && s.part === limActive.part) : -1;
  // defVis()を先に展開してから保存済み設定を上書きすることで、既存の保存設定に無い新規フィールド（例：v9.28で追加したod1〜od6）は
  // 常に表示ONをデフォルトにする（フィールド定義を後から追加・変更しても、既存ユーザーの保存済み設定に埋もれて非表示にならないようにする）
  const [vis, setVis] = useState(() => ({
    ...defVis(),
    ...lsGet(LS_KEYS.vis, {})
  }));
  const [tmpVis, setTmpVis] = useState(defVis());
  const [cardLabels, setCardLabels] = useState(() => lsGet(LS_KEYS.cardLabels, defCardLabels())); // 設定画面カードの表記（機器リスト・点検者リスト・点検基準設定）
  const [devList, setDevList] = useState(() => lsGet(LS_KEYS.devList, [])); // [{floor,room,managementNo,unitNo,_raw:{}}]
  const [devColumns, setDevColumns] = useState(() => lsGet(LS_KEYS.devColumns, [])); // CSVの全列名
  const [devVisibleCols, setDevVisibleCols] = useState(() => lsGet(LS_KEYS.devVisibleCols, [])); // 表示する列名（空=全列）
  // 設定画面：機器リストを検索して1件だけ編集・追加・削除するためのstate
  const [devEditSearch, setDevEditSearch] = useState("");
  const [devEditIdx, setDevEditIdx] = useState(null); // devListのindex／"new"／null（未選択）
  const [devEditDraft, setDevEditDraft] = useState(null); // 編集中データ {_raw:{...}}
  const [inspList, setInspList] = useState(() => lsGet(LS_KEYS.inspList, []));
  const [checkFields, setCheckFields] = useState(() => lsGet(LS_KEYS.checkFields, [
  // 室内機チェック項目
  {
    code: "ci1",
    label: "配管類支持異常の有無",
    category: "据付状態",
    group: "check_in"
  }, {
    code: "ci2",
    label: "異音・異常振動の有無",
    category: "フィルター点検・清掃",
    group: "check_in"
  }, {
    code: "ci3",
    label: "異音・異常振動の有無",
    category: "運転確認",
    group: "check_in"
  },
  // 室外機チェック項目
  {
    code: "co1",
    label: "防振装置異常の有無",
    category: "据付状態",
    group: "check_out"
  }, {
    code: "co2",
    label: "配管類支持異常の有無",
    category: "据付状態",
    group: "check_out"
  }, {
    code: "co3",
    label: "ガスリークテスト",
    category: "冷媒系統",
    group: "check_out"
  }, {
    code: "co4",
    label: "配管系統外観点検",
    category: "冷媒系統",
    group: "check_out"
  }, {
    code: "co5",
    label: "異音・異常振動の有無",
    category: "送排風機系統",
    group: "check_out"
  }, {
    code: "co6",
    label: "ドレン配管異常の有無",
    category: "排水系統",
    group: "check_out"
  }, {
    code: "co7",
    label: "フィン汚れの有無",
    category: "熱交換器系統",
    group: "check_out"
  }, {
    code: "co8",
    label: "異音・異常振動の有無",
    category: "熱交換器系統",
    group: "check_out"
  }, {
    code: "co9",
    label: "外面清掃",
    category: "作業終了時",
    group: "check_out"
  }])); // [{code,label,category,group}]
  const [lastInsp, setLastInsp] = useState("");
  const [lastDate, setLastDate] = useState(new Date().toISOString().slice(0, 10));
  const [devSearch, setDevSearch] = useState("");
  const [openSec, setOpenSec] = useState({
    device: false,
    inspector: false,
    cols: false,
    checkitems: false,
    vis: false,
    lim: false
  });
  const [modalSec, setModalSec] = useState(null); // 設定画面で開いているモーダルのid
  const [criteriaTab, setCriteriaTab] = useState("checkitems"); // 点検基準設定モーダル内のタブ（"checkitems"|"lim"）

  const [numDisp, setNumDisp] = useState("");
  const isOvr = useRef(false);
  const [activeCode, setActiveCode] = useState(null);
  const rowRefs = useRef({});
  const listRef = useRef();
  const devRef = useRef();
  const inspRef = useRef();
  const deletedIds = useRef(new Set()); // 削除直後に古いFirebase更新が届いて復活するのを防ぐ
  const settingsLoaded = useRef(false); // 初回のFirebase読み込みが終わるまでは設定を書き戻さない（初期値での上書き防止）

  useEffect(() => {
    if (view === "settings") {
      setTmpLim(JSON.parse(JSON.stringify(limits)));
      setTmpVis(JSON.parse(JSON.stringify(vis)));
    }
  }, [view]);

  // ─── 取込データの自動保存（ブラウザのlocalStorageへ）───
  // 機器リスト・表示列設定・点検項目・点検者リスト・表示項目設定・正常値範囲を
  // アプリ更新後も引き継げるようにする（入力中の点検記録は対象外）
  useEffect(() => {
    lsSet(LS_KEYS.devList, devList);
  }, [devList]);
  useEffect(() => {
    lsSet(LS_KEYS.devColumns, devColumns);
  }, [devColumns]);
  useEffect(() => {
    lsSet(LS_KEYS.devVisibleCols, devVisibleCols);
  }, [devVisibleCols]);
  useEffect(() => {
    lsSet(LS_KEYS.inspList, inspList);
  }, [inspList]);
  useEffect(() => {
    lsSet(LS_KEYS.checkFields, checkFields);
  }, [checkFields]);
  useEffect(() => {
    lsSet(LS_KEYS.limits, limits);
  }, [limits]);
  useEffect(() => {
    lsSet(LS_KEYS.vis, vis);
  }, [vis]);
  useEffect(() => {
    lsSet(LS_KEYS.cardLabels, cardLabels);
  }, [cardLabels]);

  // ─── Firebase：設定データの購読（他端末での変更を即座に反映）───
  // Firebase側にまだデータが無い（初回セットアップ直後）場合は、ローカルの設定を初期データとして書き込む。
  useEffect(() => {
    if (!db) return;
    const ref = db.ref("appSettings");
    const handler = snap => {
      const d = snap.val();
      if (d) {
        if (d.devList) setDevList(d.devList);
        if (d.devColumns) setDevColumns(d.devColumns);
        if (d.devVisibleCols) setDevVisibleCols(d.devVisibleCols);
        if (d.inspList) setInspList(d.inspList);
        if (d.checkFields) setCheckFields(d.checkFields);
        if (d.limits) setLimits(d.limits);
        // defVis()を土台にFirebase側の保存値を上書きすることで、保存済み設定に無い新規フィールド（例：v9.28のod1〜od6）は表示ONのままにする
        if (d.vis) setVis({
          ...defVis(),
          ...d.vis
        });
        if (d.cardLabels) setCardLabels(d.cardLabels);
      } else {
        ref.set({
          devList,
          devColumns,
          devVisibleCols,
          inspList,
          checkFields,
          limits,
          vis,
          cardLabels
        });
      }
      settingsLoaded.current = true;
    };
    ref.on("value", handler);
    return () => ref.off("value", handler);
  }, []);
  // 初回読み込み後にローカルで設定を変更したら、Firebaseにも書き戻して他端末に伝える
  useEffect(() => {
    if (db && settingsLoaded.current) db.ref("appSettings").update({
      devList
    });
  }, [devList]);
  useEffect(() => {
    if (db && settingsLoaded.current) db.ref("appSettings").update({
      devColumns
    });
  }, [devColumns]);
  useEffect(() => {
    if (db && settingsLoaded.current) db.ref("appSettings").update({
      devVisibleCols
    });
  }, [devVisibleCols]);
  useEffect(() => {
    if (db && settingsLoaded.current) db.ref("appSettings").update({
      inspList
    });
  }, [inspList]);
  useEffect(() => {
    if (db && settingsLoaded.current) db.ref("appSettings").update({
      checkFields
    });
  }, [checkFields]);
  useEffect(() => {
    if (db && settingsLoaded.current) db.ref("appSettings").update({
      limits
    });
  }, [limits]);
  useEffect(() => {
    if (db && settingsLoaded.current) db.ref("appSettings").update({
      vis
    });
  }, [vis]);
  useEffect(() => {
    if (db && settingsLoaded.current) db.ref("appSettings").update({
      cardLabels
    });
  }, [cardLabels]);

  // ─── Firebase：点検記録の購読（選択中の点検回・室内機・室外機それぞれ全件をリアルタイム反映）───
  useEffect(() => {
    setIndoorRecords([]);
    setOutdoorRecords([]); // 回の切替時、前の回のデータが一瞬混ざって見えないようにクリア
    if (!db || !currentRound?.id) return;
    const refIn = db.ref("rounds/" + currentRound.id + "/indoor");
    const refOut = db.ref("rounds/" + currentRound.id + "/outdoor");
    const toArr = snap => {
      const arr = [];
      snap.forEach(c => {
        if (!deletedIds.current.has(c.key)) arr.push({
          ...c.val(),
          id: c.key
        });
      });
      return arr;
    };
    const hIn = snap => setIndoorRecords(toArr(snap));
    const hOut = snap => setOutdoorRecords(toArr(snap));
    refIn.on("value", hIn);
    refOut.on("value", hOut);
    return () => {
      refIn.off("value", hIn);
      refOut.off("value", hOut);
    };
  }, [currentRound?.id]);
  const showFlash = msg => {
    setFlash(msg);
    setTimeout(() => setFlash(""), 2400);
  };
  // 同じ系統（機器番号の「-数字」を除いた部分が一致）の他の室内機に、すでに室外機データ（E/F/H等）が入力済みの場合、
  // 誤って別の値で上書きしないよう、確認して「はい」と答えるまで室外機欄の入力をロックする
  const parentOutdoorUnitNo = u => String(u || "").replace(/-\d+$/, "");
  const deviceKeyForOutdoorLock = (form.managementNo || "") + "|" + (form.unitNo || "");
  const siblingOutdoorRecord = inspectionMode === "indoor" && form.unitNo ? indoorRecords.find(r => r.unitNo !== form.unitNo && parentOutdoorUnitNo(r.unitNo) === parentOutdoorUnitNo(form.unitNo) && OUTDOOR_FIELDS.some(f => r.values?.[f.code])) : null;
  const outdoorLocked = inspectionMode === "indoor" && !!siblingOutdoorRecord && outdoorUnlockedFor !== deviceKeyForOutdoorLock;
  const requestUnlockOutdoor = () => {
    if (window.confirm("室外機データは入力済です。入力するとデータは上書きされます。入力しますか？")) {
      setOutdoorUnlockedFor(deviceKeyForOutdoorLock);
    }
  };
  const visIn = INDOOR_FIELDS.filter(f => vis[f.code]);
  const visOut = OUTDOOR_FIELDS.filter(f => vis[f.code]);
  const visFields = outdoorLocked ? ALL_FIELDS.filter(f => vis[f.code] && f.group !== "outdoor") : ALL_FIELDS.filter(f => vis[f.code]);
  const ciFields = checkFields.filter(f => f.group === "check_in");
  // 測定値項目＋室内機チェック項目を1本の連続シーケンスにして、同じ右側パネル・同じ手の位置でENTERまたは○×を続けて入力できるようにする
  const focusSeq = [...visFields, ...ciFields];
  const isCheckCode = code => ciFields.some(f => f.code === code);
  const outChkFieldsAll = checkFields.filter(f => f.group === "check_out");
  const complete = inspectionMode === "outdoor" ? outChkFieldsAll.length > 0 && outChkFieldsAll.every(f => (form.checks?.[f.code] || "") !== "") : visFields.every(f => form.values[f.code] !== "") && ciFields.every(f => (form.checks?.[f.code] || "") !== "");
  const missing = inspectionMode === "outdoor" ? outChkFieldsAll.filter(f => (form.checks?.[f.code] || "") === "").length : visFields.filter(f => form.values[f.code] === "").length + ciFields.filter(f => (form.checks?.[f.code] || "") === "").length;
  const setInfo = (k, v) => setForm(p => ({
    ...p,
    [k]: v
  }));
  const setVal = (code, v) => setForm(p => ({
    ...p,
    values: {
      ...p.values,
      [code]: v
    }
  }));
  const setCheck = (code, v) => setForm(p => ({
    ...p,
    checks: {
      ...p.checks,
      [code]: v
    }
  }));

  // Step1 state
  const s1DateDone = !!form.inspectionDate;
  const s1InspDone = !!form.inspector;
  const s1DevDone = !!(form.managementNo && form.unitNo && form.floor && form.room);
  const step1Valid = s1DateDone && s1InspDone && s1DevDone;
  const scrollToCenter = (code, smooth = true) => {
    requestAnimationFrame(() => requestAnimationFrame(() => {
      const el = rowRefs.current[code];
      const ct = listRef.current;
      if (!el || !ct) return;
      let top = 0,
        node = el;
      while (node && node !== ct) {
        top += node.offsetTop;
        node = node.offsetParent;
      }
      const target = Math.max(0, top - ct.clientHeight / 2 + el.offsetHeight / 2);
      if (smooth) smoothScrollTo(ct, target);else ct.scrollTop = target;
    }));
  };
  const onPress = key => {
    if (!activeCode || isCheckCode(activeCode)) return;
    setNumDisp(prev => {
      if (key === "C") {
        isOvr.current = false;
        return "";
      }
      if (key === "⌫") {
        isOvr.current = false;
        return prev.slice(0, -1);
      }
      if (isOvr.current) {
        isOvr.current = false;
        return key === "." ? "0." : String(key);
      }
      if (key === ".") {
        if (prev.includes(".")) return prev;
        return (prev || "0") + ".";
      }
      if (prev === "0") return String(key);
      return prev + key;
    });
  };
  const focusField = (code, smooth = true) => {
    setActiveCode(code);
    setNumDisp(code && !isCheckCode(code) ? form.values[code] || "" : "");
    isOvr.current = true;
    if (code) scrollToCenter(code, smooth);
  };
  const onConfirm = () => {
    if (!activeCode || isCheckCode(activeCode) || numDisp === "") return;
    setVal(activeCode, numDisp);
    const idx = focusSeq.findIndex(f => f.code === activeCode);
    const next = focusSeq[idx + 1];
    if (next) {
      focusField(next.code);
    } else {
      setActiveCode(null);
      setNumDisp("");
      isOvr.current = false;
    }
  };
  const moveActive = dir => {
    if (!activeCode) {
      if (focusSeq.length > 0) {
        focusField(focusSeq[0].code, true);
      }
      return;
    }
    if (!isCheckCode(activeCode) && numDisp !== "") setVal(activeCode, numDisp);
    const idx = focusSeq.findIndex(f => f.code === activeCode);
    const next = focusSeq[idx + dir];
    if (next) {
      focusField(next.code, false);
    } else {
      setActiveCode(null);
      setNumDisp("");
      isOvr.current = false;
    }
  };
  const onRowClick = f => {
    if (activeCode && !isCheckCode(activeCode) && numDisp !== "") setVal(activeCode, numDisp);
    focusField(f.code);
  };
  // ○×チェック項目：値を設定し、続けて次の項目（測定値/チェックの連続シーケンス上の次）へ自動的にフォーカスを移す
  // ※反転（同じ値を押すとクリア）はしない。何度押しても選んだ値のまま
  const setCheckAndAdvance = (code, v) => {
    setCheck(code, v);
    const idx = focusSeq.findIndex(f => f.code === code);
    const next = focusSeq[idx + 1];
    if (next) {
      focusField(next.code);
    } else {
      setActiveCode(null);
      setNumDisp("");
    }
  };
  const goToStep2 = () => {
    if (!step1Valid) return;
    const first = focusSeq[0];
    setStep(2);
    setActiveCode(first?.code || null);
    setNumDisp(first && !isCheckCode(first.code) ? form.values[first.code] || "" : "");
    isOvr.current = true;
    if (first) scrollToCenter(first.code);
  };
  // 機器選択画面内で室内機／室外機を切り替える（点検日・点検者・建物・階はそのまま、機器選択のみクリア）
  const onSwitchMode = mode => {
    if (mode === inspectionMode) return;
    setInspectionMode(mode);
    setForm(p => ({
      ...p,
      floor: "",
      room: "",
      managementNo: "",
      unitNo: "",
      values: emptyVal(),
      checks: {}
    }));
    setEditIdx(null);
    setActiveCode(null);
    setNumDisp("");
  };
  const [saveModal, setSaveModal] = useState(null); // 保存完了モーダル用データ
  const [measZoom, setMeasZoom] = useState(false); // 測定データ拡大表示
  const [tempWarningMode, setTempWarningMode] = useState(null); // 吸込・吹出温度差の確認待ち（null=非表示、値=確定時に使うmode）

  const doSave = mode => {
    setLastInsp(form.inspector);
    setLastDate(form.inspectionDate);
    const saved = {
      ...form,
      id: form.id || crypto.randomUUID()
    };
    if (editIdx !== null) {
      setRecords(p => p.map(r => r.id === saved.id ? saved : r));
      setEditIdx(null);
    } else {
      setRecords(p => [...p, saved]);
    }
    saveRecordRemote(currentRound?.id, inspectionMode, saved);
    setSaveModal({
      ...saved,
      _mode: mode
    }); // モーダル表示
  };
  const handleSave = (mode = "next") => {
    if (!complete) return;
    if (inspectionMode === "indoor") {
      const b1 = parseFloat(form.values.b1),
        b2 = parseFloat(form.values.b2);
      if (!isNaN(b1) && !isNaN(b2) && Math.abs(b1 - b2) < 5) {
        setTempWarningMode(mode);
        return;
      }
    }
    doSave(mode);
  };
  const confirmTempWarningYes = () => {
    const m = tempWarningMode;
    setTempWarningMode(null);
    doSave(m);
  };
  const confirmTempWarningNo = () => {
    setTempWarningMode(null);
    focusField("b1");
  };
  // ① 次へ：測定データ入力 → 保存後そのまま測定データ入力画面へ
  const closeSaveNext = () => {
    const insp = saveModal?.inspector || form.inspector;
    const date = saveModal?.inspectionDate || form.inspectionDate;
    setSaveModal(null);
    setForm(emptyForm(insp, date));
    setStep(inspectionMode === "outdoor" ? 1 : 1);
    setActiveCode(null);
    setNumDisp("");
    isOvr.current = false;
  };
  // ② 一次保存 → 点検日・点検者はそのまま、機器情報リセットして基本情報へ
  const closeSaveTmp = () => {
    const insp = saveModal?.inspector || form.inspector;
    const date = saveModal?.inspectionDate || form.inspectionDate;
    setSaveModal(null);
    setForm(p => ({
      ...emptyForm(insp, date),
      inspector: insp,
      inspectionDate: date
    }));
    setStep(1);
    setActiveCode(null);
    setNumDisp("");
    isOvr.current = false;
  };
  // ③ 保存後にデータ一覧へ戻る
  const closeSaveToList = () => {
    const insp = saveModal?.inspector || form.inspector;
    const date = saveModal?.inspectionDate || form.inspectionDate;
    setSaveModal(null);
    setForm(emptyForm(insp, date));
    setStep(1);
    setActiveCode(null);
    setNumDisp("");
    isOvr.current = false;
    setView("list");
  };
  const handleEdit = id => {
    const r = records.find(x => x.id === id);
    if (r) {
      setForm({
        ...r
      });
      setEditIdx(id);
      setStep(1);
      setView("form");
    }
  };
  const handleDel = id => {
    if (!window.confirm("削除しますか？")) return;
    deletedIds.current.add(id);
    setRecords(p => p.filter(r => r.id !== id));
    deleteRecordRemote(currentRound?.id, inspectionMode, id);
    if (editIdx === id) {
      setEditIdx(null);
      setForm(emptyForm(lastInsp, lastDate));
    }
  };
  // 一覧（室内機/室外機まとめ表示）用：対象の記録がどちらのモードかを指定して編集・削除する
  const handleEditFor = (rec, mode) => {
    setInspectionMode(mode);
    setForm({
      ...rec
    });
    setEditIdx(rec.id);
    setStep(1);
    setView("form");
  };
  const handleDelFor = (rec, mode) => {
    if (!window.confirm("削除しますか？")) return;
    deletedIds.current.add(rec.id);
    const setRec = mode === "indoor" ? setIndoorRecords : setOutdoorRecords;
    setRec(p => p.filter(r => r.id !== rec.id));
    deleteRecordRemote(currentRound?.id, mode, rec.id);
  };
  const handleInputFor = (row, mode) => {
    setInspectionMode(mode);
    setForm(p => ({
      ...emptyForm(lastInsp, lastDate),
      floor: row.floor,
      room: row.room,
      managementNo: row.managementNo,
      unitNo: row.unitNo,
      inspectionDate: p.inspectionDate,
      inspector: p.inspector
    }));
    const bKey = devColumns.find(k => /建物|building|棟|ビル/i.test(k));
    const rowBuilding = bKey && row._raw ? row._raw[bKey] : null;
    setSessionInfo(p => ({
      ...p,
      targetFloors: row.floor ? [row.floor] : p?.targetFloors || [],
      selectedBuildings: rowBuilding ? [rowBuilding] : p?.selectedBuildings || []
    }));
    setEditIdx(null);
    setStep(1);
    setView("form");
  };
  // 一覧の行タップ用：既存レコードがあれば編集、無ければ新規入力へ（室内機/室外機いずれか該当する方）
  const handleEditOrInputFor = (row, mode) => {
    const rec = mode === "indoor" ? row.indoorRecord : row.outdoorRecord;
    if (rec) handleEditFor(rec, mode);else handleInputFor(row, mode);
  };
  const loadXLSX = () => new Promise((resolve, reject) => {
    if (window.XLSX) {
      resolve(window.XLSX);
      return;
    }
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    s.onload = () => resolve(window.XLSX);
    s.onerror = reject;
    document.head.appendChild(s);
  });
  const readFile = (file, onParsed) => {
    const isXlsx = /\.xlsx?$/i.test(file.name);
    const reader = new FileReader();
    if (isXlsx) {
      reader.onload = async ev => {
        const XLSX = await loadXLSX();
        const wb = XLSX.read(ev.target.result, {
          type: "array"
        });
        const ws = wb.Sheets[wb.SheetNames[0]];
        // sheet_to_json でヘッダー行をキーとして取得（列名の揺れに対応）
        const json = XLSX.utils.sheet_to_json(ws, {
          defval: "",
          raw: false
        });
        if (json.length === 0) {
          onParsed(null, []);
          return;
        }
        // __EMPTYなどSheetJS内部列を除外
        const keys = Object.keys(json[0]).filter(k => !k.startsWith("__"));
        onParsed(json, keys); // JSON配列をそのまま渡す
      };
      reader.readAsArrayBuffer(file);
    } else {
      reader.onload = ev => {
        const bytes = new Uint8Array(ev.target.result);
        const hasUtf8Bom = bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF;
        const enc = hasUtf8Bom ? "UTF-8" : "Shift_JIS";
        const text = new TextDecoder(enc).decode(bytes).replace(/^\uFEFF/, "");
        const cols = text.trim().split(/\r?\n/)[0]?.split(",").map(s => s.trim().replace(/^"|"$/g, "")).filter(Boolean) || [];
        onParsed(text, cols);
      };
      reader.readAsArrayBuffer(file);
    }
  };
  const handleDevCSV = e => {
    const f = e.target.files[0];
    if (!f) return;
    readFile(f, (data, cols) => {
      let l;
      if (Array.isArray(data)) {
        // xlsx: JSON配列
        l = parseDevRows(data, cols);
      } else {
        // csv: 文字列
        l = parseDevCSV(data || "");
      }
      setDevList(l);
      if (cols && cols.length > 0) {
        const filteredCols = cols.filter(k => !k.startsWith("__"));
        setDevColumns(filteredCols);
        setDevVisibleCols(filteredCols);
      }
      showFlash("✅ 機器リスト " + l.length + "件 読込");
    });
    e.target.value = "";
  };
  const handleInspCSV = e => {
    const f = e.target.files[0];
    if (!f) return;
    readFile(f, (data, cols) => {
      let text;
      if (Array.isArray(data)) {
        // xlsx: 1列目の値を名前リストとして取得
        const nameKey = cols && cols.length > 0 ? cols[0] : Object.keys(data[0] || {})[0] || "";
        text = data.map(row => String(row[nameKey] || "").trim()).filter(Boolean).join("\n");
      } else {
        text = data || "";
      }
      const l = parseInspCSV(text);
      setInspList(l);
      showFlash("✅ 点検者 " + l.length + "名 読込");
    });
    e.target.value = "";
  };
  const toggleSec = k => setOpenSec(p => ({
    ...p,
    [k]: !p[k]
  }));
  const vf = ALL_FIELDS.filter(f => vis[f.code]);
  const [listFilter, setListFilter] = useState("all"); // "all"|"done"|"undone"
  const [sortCol, setSortCol] = useState(null); // ソート列
  const [hoverRow, setHoverRow] = useState(null); // 一覧テーブルのマウスオーバー中の行
  const [sortDir, setSortDir] = useState("asc");
  const [showStats, setShowStats] = useState(false); // 集計モーダル
  const [showVoiceMemo, setShowVoiceMemo] = useState(false); // ボイスメモモーダル（アプリ全体向け・iPadのみ）
  const [floorFilter, setFloorFilter] = useState(null); // 階フィルター（null=全階）

  // 一覧・集計・印刷では室内機/室外機の記録を両方まとめて1行にする（parentOutdoorUnitNoは上部で定義済み）
  const tRows = (() => {
    if (devList.length > 0) {
      return devList.map(d => ({
        ...d,
        indoorRecord: indoorRecords.find(r => r.managementNo === d.managementNo && r.unitNo === d.unitNo) || null,
        outdoorRecord: outdoorRecords.find(r => r.managementNo === d.managementNo && r.unitNo === d.unitNo) || null,
        // 同じ系統の室内機（機器番号が"この機器番号-数字"の形）で入力された室外機側の測定値（E/F/H等）を拾えるように
        linkedIndoorRecords: indoorRecords.filter(r => r.unitNo !== d.unitNo && parentOutdoorUnitNo(r.unitNo) === d.unitNo)
      }));
    }
    const map = new Map();
    indoorRecords.forEach(r => {
      map.set(r.managementNo + "|" + r.unitNo, {
        floor: r.floor,
        room: r.room,
        managementNo: r.managementNo,
        unitNo: r.unitNo,
        indoorRecord: r,
        outdoorRecord: null
      });
    });
    outdoorRecords.forEach(r => {
      const key = r.managementNo + "|" + r.unitNo;
      const ex = map.get(key);
      if (ex) ex.outdoorRecord = r;else map.set(key, {
        floor: r.floor,
        room: r.room,
        managementNo: r.managementNo,
        unitNo: r.unitNo,
        indoorRecord: null,
        outdoorRecord: r
      });
    });
    return [...map.values()].map(row => ({
      ...row,
      linkedIndoorRecords: indoorRecords.filter(r => r.unitNo !== row.unitNo && parentOutdoorUnitNo(r.unitNo) === row.unitNo)
    }));
  })();
  // 行から値を取り出す共通ヘルパー（室内機優先→室外機→同系統の室内機で入力された室外機データの順にフォールバック）
  // ※同系統の室内機からの値は「室外機（アウトドア）」グループの項目（E/F/H等）のみを対象にする（b1等の室内機データはその室内機自身の行にのみ表示する）
  const rowFieldVal = (row, code) => {
    const iv = row.indoorRecord?.values?.[code];
    if (iv) return iv;
    const ov = row.outdoorRecord?.values?.[code];
    if (ov) return ov;
    if (row.linkedIndoorRecords && ALL_FIELDS.find(f => f.code === code)?.group === "outdoor") {
      for (const r of row.linkedIndoorRecords) {
        const lv = r.values?.[code];
        if (lv) return lv;
      }
    }
    return "";
  };
  const rowMeta = (row, key) => row.indoorRecord?.[key] || row.outdoorRecord?.[key] || "";
  const rowRemarks = row => {
    const a = row.indoorRecord?.remarks || "",
      b = row.outdoorRecord?.remarks || "";
    if (a && b && a !== b) return a + " / " + b;
    return a || b || "";
  };
  const rowIndoorDone = row => !!row.indoorRecord && Object.values(row.indoorRecord.values).some(v => v !== "");
  const rowOutdoorDone = row => !!row.outdoorRecord && Object.values(row.outdoorRecord.checks || {}).some(v => v !== "");
  const rowDone = row => rowIndoorDone(row) || rowOutdoorDone(row);
  // 室内機の行は室内機チェック（check_in）、室外機の行は室外機チェック（check_out）の結果を表示する
  const rowChecksList = row => {
    if (row.indoorRecord) return checkFields.filter(f => f.group === "check_in").map(f => ({
      code: f.code,
      label: f.label,
      val: row.indoorRecord.checks?.[f.code] || ""
    }));
    if (row.outdoorRecord) return checkFields.filter(f => f.group === "check_out").map(f => ({
      code: f.code,
      label: f.label,
      val: row.outdoorRecord.checks?.[f.code] || ""
    }));
    return [];
  };
  const handleStep1TmpSave = () => {
    if (!step1Valid) return;
    setLastInsp(form.inspector);
    setLastDate(form.inspectionDate);
    let id = form.id;
    if (!id) {
      if (editIdx !== null) id = editIdx;else {
        const ex = records.find(r => r.managementNo === form.managementNo && r.unitNo === form.unitNo);
        id = ex ? ex.id : crypto.randomUUID();
      }
    }
    const saved = {
      ...form,
      values: emptyVal(),
      id
    };
    setRecords(p => p.some(r => r.id === id) ? p.map(r => r.id === id ? saved : r) : [...p, saved]);
    if (editIdx !== null) setEditIdx(null);
    saveRecordRemote(currentRound?.id, inspectionMode, saved);
    setForm(p => ({
      ...emptyForm(p.inspector, p.inspectionDate),
      inspector: p.inspector,
      inspectionDate: p.inspectionDate
    }));
    showFlash("💾 一次保存しました");
  };

  // ─── STEP1 ────────────────────────────────────────────
  // STEP1 テンキー対象フィールド
  const [s1Focus, setS1Focus] = useState(null); // 'date'|'inspector'|'device'|'preSetTemp'

  // Step1View → top-level component

  // フィルター適用
  const listBuildingKey = devColumns.find(k => /建物|building|棟|ビル/i.test(k)) || null;
  // 列の値を取得する共通ヘルパー（オートフィルタ・ソート両方で使用）
  const colValue = (row, col) => {
    if (col === "建物") return listBuildingKey ? String(row._raw?.[listBuildingKey] || "（未設定）") : "（未設定）";
    if (col === "階") return row.floor || "（未設定）";
    if (col === "部屋名") return row.room || "（未設定）";
    if (col === "管理番号") return row.managementNo || "（未設定）";
    if (col === "機器番号") return row.unitNo || "（未設定）";
    if (col === "室内機") return rowIndoorDone(row) ? "入力済" : "未入力";
    if (col === "室外機") return rowOutdoorDone(row) ? "入力済" : "未入力";
    if (col === "点検日") return rowMeta(row, "inspectionDate") || "（未設定）";
    if (col === "点検者") return rowMeta(row, "inspector") || "（未設定）";
    if (row._raw && row._raw[col] !== undefined) return String(row._raw[col] || "（未設定）");
    return "";
  };
  // 絞り込み：完了状況・機器種別・建物・階（パネル式、それぞれ単一選択）
  const statsCategoryKey = devColumns.find(k => /^(分類|category|class)$/i.test(k.trim())) || null;
  const rowStatsType = row => {
    if (statsCategoryKey && row._raw) {
      const c = String(row._raw[statsCategoryKey] || "");
      if (/室内機/.test(c)) return "indoor";
      if (/室外機/.test(c)) return "outdoor";
    }
    if (row.indoorRecord) return "indoor";
    if (row.outdoorRecord) return "outdoor";
    return "unknown";
  };
  const [typeFilter, setTypeFilter] = useState(null); // null|"indoor"|"outdoor"
  const [buildingFilter, setBuildingFilter] = useState(null); // null=すべての建物
  const filteredRows = (() => {
    const rows = tRows.filter(row => {
      if (floorFilter && row.floor !== floorFilter) return false;
      if (buildingFilter && colValue(row, "建物") !== buildingFilter) return false;
      if (typeFilter && rowStatsType(row) !== typeFilter) return false;
      if (listFilter === "done") return rowDone(row);
      if (listFilter === "undone") return !rowDone(row);
      return true;
    });
    if (!sortCol) return rows;
    return [...rows].sort((a, b) => {
      const va = colValue(a, sortCol),
        vb = colValue(b, sortCol);
      const cmp = va.localeCompare(vb, undefined, {
        numeric: true,
        sensitivity: "base"
      });
      return sortDir === "asc" ? cmp : -cmp;
    });
  })();

  // データ一覧の絞り込み状態をまとめて説明文にする（CSV出力前の確認ポップアップ用）
  const activeFilterDescriptions = () => {
    const list = [];
    if (listFilter !== "all") list.push("完了状況：" + (listFilter === "done" ? "入力済" : "未入力"));
    if (typeFilter) list.push("機器種別：" + (typeFilter === "indoor" ? "🏠 室内機" : "🏭 室外機"));
    if (buildingFilter) list.push("建物：" + buildingFilter);
    if (floorFilter) list.push("階：" + floorFilter);
    if (sortCol) list.push("並び替え：" + sortCol + "（" + (sortDir === "asc" ? "昇順" : "降順") + "）");
    return list;
  };
  // データ一覧に表示されている形（見えている列・行）のままCSV（TSV）出力する
  const exportListAsCSV = () => {
    const cols = devVisibleCols.length > 0 ? devVisibleCols : ["階", "部屋名", "管理番号", "機器番号"];
    const header = [...cols, "室内機", "室外機", "点検日", "点検者", "運転", "モード", "風量", "設定温度", ...vf.map(f => f.code + "(" + f.unit + ")"), ...checkFields.map(f => f.label), "備考"];
    const dataRows = filteredRows.map(row => {
      const ir = row.indoorRecord,
        or_ = row.outdoorRecord;
      const metaSrc = ir || or_;
      const colVals = cols.map(col => {
        const val = row._raw ? row._raw[col] : col === "階" ? row.floor : col === "部屋名" ? row.room : col === "管理番号" ? row.managementNo : col === "機器番号" ? row.unitNo : "";
        return val || "";
      });
      const checksVals = checkFields.map(f => {
        if (ir && f.group === "check_in") return ir.checks?.[f.code] || "";
        if (or_ && f.group === "check_out") return or_.checks?.[f.code] || "";
        return "";
      });
      return [...colVals, rowIndoorDone(row) ? "入力済" : "未入力", rowOutdoorDone(row) ? "入力済" : "未入力", metaSrc?.inspectionDate || "", metaSrc?.inspector || "", metaSrc?.preOperation || "", metaSrc?.preMode || "", metaSrc?.preWind || "", metaSrc?.preSetTemp || "", ...vf.map(f => rowFieldVal(row, f.code)), ...checksVals, rowRemarks(row) || ""];
    });
    const rows = [header, ...dataRows];
    const tsv = rows.map(r => r.map(c => '"' + String(c ?? "").replace(/"/g, '""') + '"').join("\t")).join("\n");
    const blob = new Blob(["\uFEFF" + tsv], {
      type: "text/tab-separated-values;charset=utf-8;"
    });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "ac_check_list_" + new Date().toISOString().slice(0, 10) + ".tsv";
    a.click();
  };
  // 絞り込み中の出力確認ポップアップ（null=非表示／{list,run,label}=絞り込み内容・実行する出力関数・表示ラベル）。CSV出力・点検表出力で共用。
  const [showExportConfirm, setShowExportConfirm] = useState(null);
  const handleCSVExportClick = () => {
    if (IS_IPAD) {
      showFlash("⚠️ iPadではこの機能は使えません");
      return;
    }
    if (filteredRows.length === 0) {
      showFlash("⚠️ データがありません");
      return;
    }
    const active = activeFilterDescriptions();
    if (active.length > 0) {
      setShowExportConfirm({
        list: active,
        run: exportListAsCSV,
        label: "CSV"
      });
    } else {
      exportListAsCSV();
    }
  };

  // ─── 点検表（Excel）出力 ─────────────────────────────────────
  // 「エアコン定期点検.xls」の室内機／室外機シートと同じレイアウトを、Excelが解釈できるHTMLテーブル（拡張子.xls）として出力する。
  // SheetJSの書き込み（コミュニティ版）は罫線・太字・背景色などのセル装飾を保存できないため、
  // 罫線・結合セル・見出し文字などレイアウトが本質的なこの点検表には、HTMLテーブル＋インラインCSSの方式を使う。
  const escapeHtml = v => String(v ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const fmtDateMD = iso => {
    const m = String(iso || "").match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    return m ? parseInt(m[2], 10) + "/" + parseInt(m[3], 10) : iso || "";
  };
  const deriveVerdict = checks => {
    const vals = Object.values(checks || {}).filter(v => v !== "");
    if (vals.length === 0) return "";
    return vals.includes("×") ? "否" : "合";
  };
  const criteriaTextFor = label => {
    const l = String(label || "");
    if (/リーク/.test(l)) return "漏れがないこと";
    if (/汚れ/.test(l)) return "汚れがないこと";
    if (/清掃/.test(l)) return "清掃実施のこと";
    return "異常がないこと";
  };
  const RPT_STY = {
    title: 'style="border:none;font-size:13pt;font-weight:bold;"',
    blank: 'style="border:none;"',
    hdr: 'style="border:1px solid #000;background:#DCE6F1;font-weight:bold;text-align:center;font-size:9pt;padding:3px 5px;white-space:nowrap;"',
    crit: 'style="border:1px solid #000;background:#F2F2F2;text-align:center;font-size:8pt;padding:3px 5px;white-space:nowrap;"',
    cell: 'style="border:1px solid #000;text-align:center;font-size:9pt;padding:3px 6px;white-space:nowrap;"',
    cellL: 'style="border:1px solid #000;text-align:left;font-size:9pt;padding:3px 6px;white-space:nowrap;"',
    legend: 'style="border:none;font-size:9pt;color:#444;padding-top:6px;"'
  };
  // 室内機シート：1行＝1台。既存点検表の3段見出し（点検日/部屋名/管理番号=3行結合、チェック項目見出し=2行結合＋3行目に管理値、
  // 運転調整・データ採取=4列結合＋2行目に運転モード/設定温度/吸込温度/吹出温度、点検者/判定/備考=3行結合）を再現する。
  const buildIndoorSheetHtml = (rows, buildingLabel) => {
    const ci = checkFields.filter(f => f.group === "check_in");
    const totalCols = 3 + ci.length + 4 + 3;
    let html = "<table>";
    html += '<tr><td colspan="' + totalCols + '" ' + RPT_STY.title + '>保守点検業務' + (buildingLabel ? " [" + escapeHtml(buildingLabel) + "]" : "") + '</td></tr>';
    html += '<tr><td colspan="' + totalCols + '" ' + RPT_STY.title + '>パッケージエアコン・ビルマルチエアコン定期点検（室内機）</td></tr>';
    html += '<tr><td colspan="' + totalCols + '" ' + RPT_STY.blank + '>&nbsp;</td></tr>';
    html += "<tr>";
    html += '<td rowspan="3" ' + RPT_STY.hdr + '>点検日</td>';
    html += '<td rowspan="3" ' + RPT_STY.hdr + '>部屋名</td>';
    html += '<td rowspan="3" ' + RPT_STY.hdr + '>管理番号</td>';
    ci.forEach(f => {
      html += '<td rowspan="2" ' + RPT_STY.hdr + '>' + escapeHtml(f.category) + '</td>';
    });
    html += '<td colspan="4" ' + RPT_STY.hdr + '>運転調整・データ採取</td>';
    html += '<td rowspan="3" ' + RPT_STY.hdr + '>点検者</td>';
    html += '<td rowspan="3" ' + RPT_STY.hdr + '>判定</td>';
    html += '<td rowspan="3" ' + RPT_STY.hdr + '>備考</td>';
    html += "</tr><tr>";
    ["運転モード", "設定温度", "吸込温度", "吹出温度"].forEach(t => {
      html += '<td ' + RPT_STY.hdr + '>' + t + '</td>';
    });
    html += "</tr><tr>";
    ci.forEach(f => {
      html += '<td ' + RPT_STY.crit + '>' + escapeHtml(criteriaTextFor(f.label)) + '</td>';
    });
    html += '<td ' + RPT_STY.crit + '></td><td ' + RPT_STY.crit + '></td>';
    html += '<td ' + RPT_STY.crit + '>15～30℃</td><td ' + RPT_STY.crit + '>5～50℃</td>';
    html += "</tr>";
    rows.forEach(row => {
      const ir = row.indoorRecord;
      html += "<tr>";
      html += '<td ' + RPT_STY.cell + '>' + escapeHtml(fmtDateMD(ir?.inspectionDate)) + '</td>';
      html += '<td ' + RPT_STY.cellL + '>' + escapeHtml(row.room) + '</td>';
      html += '<td ' + RPT_STY.cell + '>' + escapeHtml(row.managementNo) + '</td>';
      ci.forEach(f => {
        html += '<td ' + RPT_STY.cell + '>' + escapeHtml(ir?.checks?.[f.code] || "") + '</td>';
      });
      html += '<td ' + RPT_STY.cell + '>' + escapeHtml(ir?.preMode || "") + '</td>';
      html += '<td ' + RPT_STY.cell + '>' + escapeHtml(ir?.preSetTemp || "") + '</td>';
      html += '<td ' + RPT_STY.cell + '>' + escapeHtml(rowFieldVal(row, "b1")) + '</td>';
      html += '<td ' + RPT_STY.cell + '>' + escapeHtml(rowFieldVal(row, "b2")) + '</td>';
      html += '<td ' + RPT_STY.cell + '>' + escapeHtml(ir?.inspector || "") + '</td>';
      html += '<td ' + RPT_STY.cell + '>' + escapeHtml(deriveVerdict(ir?.checks)) + '</td>';
      html += '<td ' + RPT_STY.cellL + '>' + escapeHtml(ir?.remarks || "") + '</td>';
      html += "</tr>";
    });
    html += '<tr><td colspan="' + totalCols + '" ' + RPT_STY.legend + '>凡例（良好：○　注意要：△　不良：×　処置後良好：●）</td></tr>';
    html += "</table>";
    return html;
  };
  // 室外機シート：1列＝1台の横並び形式をそのまま再現する。チェック項目はcategory単位で行結合し、
  // 実測値6項目（od1〜od6）の値はv9.26の系統紐付け（同系統の室内機記録からのフォールバック）を含むrowFieldValをそのまま利用する。
  const buildOutdoorSheetHtml = (rows, buildingLabel) => {
    const coAll = checkFields.filter(f => f.group === "check_out");
    const coMain = coAll.filter(f => f.category !== "作業終了時");
    const coEnd = coAll.filter(f => f.category === "作業終了時");
    const totalCols = 3 + rows.length;
    let html = "<table>";
    html += '<tr><td colspan="' + totalCols + '" ' + RPT_STY.title + '>保守点検業務' + (buildingLabel ? " [" + escapeHtml(buildingLabel) + "]" : "") + '</td></tr>';
    html += '<tr><td colspan="' + totalCols + '" ' + RPT_STY.title + '>パッケージエアコン・ビルマルチエアコン定期点検（室外機）</td></tr>';
    html += '<tr><td colspan="' + totalCols + '" ' + RPT_STY.blank + '>&nbsp;</td></tr>';
    const labelRow = (label, vals) => {
      let r = '<tr><td colspan="3" ' + RPT_STY.hdr + '>' + escapeHtml(label) + '</td>';
      vals.forEach(v => {
        r += '<td ' + RPT_STY.cell + '>' + escapeHtml(v) + '</td>';
      });
      return r + "</tr>";
    };
    html += labelRow("点検日", rows.map(row => fmtDateMD(rowMeta(row, "inspectionDate"))));
    html += labelRow("管理番号", rows.map(row => row.managementNo || ""));
    html += labelRow("客先呼称", rows.map(row => row.room || ""));
    html += "<tr>";
    html += '<td colspan="2" ' + RPT_STY.hdr + '>点検項目</td><td ' + RPT_STY.hdr + '>管理値</td>';
    rows.forEach(() => {
      html += '<td ' + RPT_STY.hdr + '>結果・処置</td>';
    });
    html += "</tr>";
    const renderChecklistGroup = fields => {
      const cats = [];
      fields.forEach(f => {
        const last = cats[cats.length - 1];
        if (last && last.category === f.category) last.items.push(f);else cats.push({
          category: f.category,
          items: [f]
        });
      });
      let out = "";
      cats.forEach(({
        category,
        items
      }) => {
        items.forEach((f, i) => {
          out += "<tr>";
          if (i === 0) out += '<td rowspan="' + items.length + '" ' + RPT_STY.hdr + '>' + escapeHtml(category) + '</td>';
          out += '<td ' + RPT_STY.hdr + '>' + escapeHtml(f.label) + '</td>';
          out += '<td ' + RPT_STY.crit + '>' + escapeHtml(criteriaTextFor(f.label)) + '</td>';
          rows.forEach(row => {
            out += '<td ' + RPT_STY.cell + '>' + escapeHtml(row.outdoorRecord?.checks?.[f.code] || "") + '</td>';
          });
          out += "</tr>";
        });
      });
      return out;
    };
    html += renderChecklistGroup(coMain);
    OUTDOOR_FIELDS.forEach((f, i) => {
      html += "<tr>";
      if (i === 0) html += '<td rowspan="' + OUTDOOR_FIELDS.length + '" ' + RPT_STY.hdr + '>運転調整・データ採取</td>';
      html += '<td ' + RPT_STY.hdr + '>' + escapeHtml(f.label) + '</td><td ' + RPT_STY.crit + '></td>';
      rows.forEach(row => {
        const v = rowFieldVal(row, f.code);
        html += '<td ' + RPT_STY.cell + '>' + escapeHtml(v ? v + f.unit : "") + '</td>';
      });
      html += "</tr>";
    });
    html += renderChecklistGroup(coEnd);
    html += labelRow("点検者", rows.map(row => rowMeta(row, "inspector") || ""));
    html += labelRow("判定", rows.map(row => deriveVerdict(row.outdoorRecord?.checks)));
    html += labelRow("特記事項", rows.map(row => rowRemarks(row) || ""));
    html += '<tr><td colspan="' + totalCols + '" ' + RPT_STY.legend + '>凡例（良好：○　注意要：△　不良：×　処置後良好：●）</td></tr>';
    html += "</table>";
    return html;
  };
  const reportScopeRows = () => ({
    indoorRows: filteredRows.filter(row => rowStatsType(row) === "indoor" && rowIndoorDone(row)),
    outdoorRows: filteredRows.filter(row => rowStatsType(row) === "outdoor" && (rowOutdoorDone(row) || OUTDOOR_FIELDS.some(f => rowFieldVal(row, f.code))))
  });
  const exportInspectionReport = () => {
    const {
      indoorRows,
      outdoorRows
    } = reportScopeRows();
    if (indoorRows.length === 0 && outdoorRows.length === 0) {
      showFlash("⚠️ 出力できるデータがありません");
      return;
    }
    const buildings = [...new Set([...indoorRows, ...outdoorRows].map(row => colValue(row, "建物")))].filter(b => b && b !== "（未設定）");
    const buildingLabel = buildings.length === 1 ? buildings[0] : "";
    const indoorHtml = buildIndoorSheetHtml(indoorRows, buildingLabel);
    const outdoorHtml = buildOutdoorSheetHtml(outdoorRows, buildingLabel);
    // Excelの別シートタブ（室内機／室外機）に分けるMHTML（multipart/related）形式を試したが、
    // Excel COM自動化での検証で正しく開けなかった（2枚とも空になる）ため、確実に動作する
    // 「1枚のシートに室内機セクション→室外機セクションを縦に並べる」形式を採用する。
    const doc = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' + "<head><meta charset=\"UTF-8\">" + "<style>table{border-collapse:collapse;} td{font-family:'ＭＳ Ｐゴシック',sans-serif;}</style>" + "</head><body>" + indoorHtml + '<div style="height:24px;">&nbsp;</div>' + outdoorHtml + "</body></html>";
    const blob = new Blob(["﻿" + doc], {
      type: "application/vnd.ms-excel;charset=utf-8;"
    });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "点検表_" + (currentRound?.label || new Date().toISOString().slice(0, 10)) + ".xls";
    a.click();
  };
  const handleReportExportClick = () => {
    if (IS_IPAD) {
      showFlash("⚠️ iPadではこの機能は使えません");
      return;
    }
    const {
      indoorRows,
      outdoorRows
    } = reportScopeRows();
    if (indoorRows.length === 0 && outdoorRows.length === 0) {
      showFlash("⚠️ 出力できるデータがありません");
      return;
    }
    const active = activeFilterDescriptions();
    if (active.length > 0) {
      setShowExportConfirm({
        list: active,
        run: exportInspectionReport,
        label: "点検表"
      });
    } else {
      exportInspectionReport();
    }
  };

  // 集計（建物×階、室内機/室外機を分けて集計）
  const _floors = [...new Set(tRows.map(r => r.floor).filter(Boolean))].sort();
  const groupStats = (() => {
    const map = new Map();
    tRows.forEach(row => {
      const building = listBuildingKey ? String(row._raw?.[listBuildingKey] || "—") : "—";
      const key = building + "|" + (row.floor || "—");
      if (!map.has(key)) map.set(key, {
        building,
        floor: row.floor || "—",
        indoorRows: [],
        outdoorRows: []
      });
      const g = map.get(key);
      const t = rowStatsType(row);
      if (t === "indoor") g.indoorRows.push(row);else if (t === "outdoor") g.outdoorRows.push(row);
    });
    return [...map.values()].map(g => ({
      building: g.building,
      floor: g.floor,
      indoorTotal: g.indoorRows.length,
      indoorDone: g.indoorRows.filter(rowIndoorDone).length,
      outdoorTotal: g.outdoorRows.length,
      outdoorDone: g.outdoorRows.filter(rowOutdoorDone).length
    })).sort((a, b) => a.building.localeCompare(b.building, undefined, {
      numeric: true
    }) || a.floor.localeCompare(b.floor, undefined, {
      numeric: true
    }));
  })();
  const floorStats = _floors.map(f => {
    const rows = tRows.filter(r => r.floor === f);
    const done = rows.filter(rowDone).length;
    return {
      floor: f,
      total: rows.length,
      done,
      undone: rows.length - done
    };
  });
  const modeLabel = m => ({
    "冷房": "❄️ 冷房",
    "暖房": "🔥 暖房",
    "送風": "💨 送風",
    "除湿": "💧 除湿"
  })[m] || m || "—";
  const windLabel = w => ({
    "自動": "🔄 自動",
    "弱": "💨 弱",
    "強": "💨 強",
    "急風": "🌪️ 急風"
  })[w] || w || "—";
  const closeNext = saveModal && saveModal._mode === "tmp" ? closeSaveTmp : closeSaveNext;
  const closeTmp = saveModal && saveModal._mode === "tmp" ? closeSaveNext : closeSaveTmp;
  const nextLabel = saveModal && saveModal._mode === "tmp" ? "→ 基本情報へ" : "→ 次の機器へ";

  // ─── render ───────────────────────────────────────────
  if (!currentRound) {
    return /*#__PURE__*/React.createElement("div", {
      style: {
        fontFamily: "Hiragino Sans, Meiryo, Arial, sans-serif",
        height: "100%",
        display: "flex",
        flexDirection: "column",
        background: C.g100,
        color: C.g800,
        overflow: "auto",
        fontSize: 16
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        background: "linear-gradient(135deg," + C.navy + " 0%," + C.blue + " 100%)",
        color: C.white,
        padding: "18px 16px",
        flexShrink: 0
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        fontWeight: 700,
        fontSize: 17,
        letterSpacing: "0.04em"
      }
    }, "\uD83C\uDF21\uFE0F \u30A8\u30A2\u30B3\u30F3\u70B9\u691C"), /*#__PURE__*/React.createElement("div", {
      style: {
        fontSize: 12,
        opacity: 0.85,
        marginTop: 2
      }
    }, "\u70B9\u691C\u56DE\u3092\u9078\u629E\u3057\u3066\u304F\u3060\u3055\u3044")), /*#__PURE__*/React.createElement("div", {
      style: {
        flex: 1,
        padding: 20
      }
    }, /*#__PURE__*/React.createElement(RoundSelector, {
      current: null,
      onSelect: setCurrentRound
    })));
  }
  return /*#__PURE__*/React.createElement("div", {
    style: {
      fontFamily: "Hiragino Sans, Meiryo, Arial, sans-serif",
      height: "100%",
      display: "flex",
      flexDirection: "column",
      background: C.g100,
      color: C.g800,
      overflow: "hidden",
      fontSize: 16
    }
  }, /*#__PURE__*/React.createElement("style", null, PS), /*#__PURE__*/React.createElement("div", {
    style: {
      background: "linear-gradient(135deg," + C.navy + " 0%," + C.blue + " 100%)",
      color: C.white,
      flexShrink: 0,
      boxShadow: "0 2px 10px rgba(0,0,0,0.25)",
      display: "flex",
      flexDirection: "column"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "8px 16px 4px",
      fontWeight: 700,
      fontSize: 16,
      letterSpacing: "0.04em"
    }
  }, "\uD83C\uDF21\uFE0F \u30A8\u30A2\u30B3\u30F3\u70B9\u691C"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 6,
      padding: "6px 12px 8px",
      flexWrap: "wrap",
      justifyContent: "space-between",
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 6,
      flexWrap: "wrap"
    }
  }, [["form", "📝 点検データ入力"], ["list", "📋 データ一覧" + (tRows.length > 0 ? " (" + tRows.length + ")" : "")]].map(([v, label]) => /*#__PURE__*/React.createElement("button", {
    key: v,
    onClick: () => {
      setView(v);
      if (v === "form" && editIdx === null) setStep(inspectionMode === "outdoor" ? 1 : 0);
    },
    style: {
      padding: "6px 12px",
      borderRadius: 7,
      border: "none",
      cursor: "pointer",
      fontSize: 12,
      fontWeight: 700,
      whiteSpace: "nowrap",
      background: view === v ? C.white : "rgba(255,255,255,0.18)",
      color: view === v ? C.navy : C.white
    }
  }, label)), /*#__PURE__*/React.createElement("button", {
    onClick: () => setShowStats(true),
    disabled: tRows.length === 0,
    style: {
      padding: "6px 10px",
      borderRadius: 7,
      border: "1.5px solid rgba(255,255,255,0.4)",
      cursor: tRows.length > 0 ? "pointer" : "not-allowed",
      fontSize: 12,
      fontWeight: 700,
      background: "rgba(255,255,255,0.15)",
      color: C.white,
      opacity: tRows.length > 0 ? 1 : 0.5,
      whiteSpace: "nowrap"
    }
  }, "\uD83D\uDCCA \u96C6\u8A08")), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 6,
      flexWrap: "wrap"
    }
  }, /*#__PURE__*/React.createElement("button", {
    key: "roundswitch",
    onClick: () => setShowRoundSwitcher(true),
    style: {
      padding: "6px 12px",
      borderRadius: 7,
      border: "none",
      cursor: "pointer",
      fontSize: 12,
      fontWeight: 700,
      whiteSpace: "nowrap",
      background: "rgba(255,255,255,0.18)",
      color: C.white
    }
  }, "\uD83D\uDCC5 ", currentRound.label), IS_IPAD && /*#__PURE__*/React.createElement("button", {
    key: "voicememo",
    onClick: () => setShowVoiceMemo(true),
    style: {
      padding: "6px 12px",
      borderRadius: 7,
      border: "none",
      cursor: "pointer",
      fontSize: 12,
      fontWeight: 700,
      whiteSpace: "nowrap",
      background: "rgba(255,255,255,0.18)",
      color: C.white
    }
  }, "\uD83C\uDF99\uFE0F \u30E1\u30E2"), /*#__PURE__*/React.createElement("button", {
    key: "settings",
    onClick: () => setView("settings"),
    style: {
      padding: "6px 12px",
      borderRadius: 7,
      border: "none",
      cursor: "pointer",
      fontSize: 12,
      fontWeight: 700,
      whiteSpace: "nowrap",
      background: view === "settings" ? C.white : "rgba(255,255,255,0.18)",
      color: view === "settings" ? C.navy : C.white
    }
  }, "\u2699\uFE0F \u8A2D\u5B9A")))), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      overflow: "hidden",
      display: "flex",
      flexDirection: "column"
    }
  }, view === "form" && /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      overflow: "hidden",
      display: "flex",
      flexDirection: "column"
    }
  }, step === 0 && /*#__PURE__*/React.createElement(SessionView, {
    devList: devList,
    inspList: inspList,
    records: records,
    sessionInfo: sessionInfo,
    setSessionInfo: setSessionInfo,
    undoneOnly: undoneOnly,
    setUndoneOnly: setUndoneOnly,
    devColumns: devColumns,
    devVisibleCols: devVisibleCols,
    onStart: info => {
      setSessionInfo(info);
      setForm(p => ({
        ...p,
        inspectionDate: info.date,
        inspector: info.inspector
      }));
      setLastInsp(info.inspector);
      setLastDate(info.date);
      setStep(1);
    },
    onSelectOutdoor: info => {
      setInspectionMode("outdoor");
      setSessionInfo(p => ({
        ...(p || {}),
        date: info.date,
        inspector: info.inspector,
        inspector1: info.inspector1,
        inspector2: info.inspector2,
        inspector3: info.inspector3,
        inspector4: info.inspector4
      }));
      setForm(emptyForm(info.inspector, info.date));
      setLastInsp(info.inspector);
      setLastDate(info.date);
      setEditIdx(null);
      setStep(1);
    }
  }), step >= 1 && /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      overflow: "hidden",
      display: "flex",
      flexDirection: "column",
      paddingTop: 8
    }
  }, /*#__PURE__*/React.createElement(Step1View, {
    form: form,
    setInfo: setInfo,
    inspList: inspList,
    devList: devList,
    devSearch: devSearch,
    setDevSearch: setDevSearch,
    s1DateDone: s1DateDone,
    s1InspDone: s1InspDone,
    s1DevDone: s1DevDone,
    step1Valid: step1Valid,
    records: records,
    s1Focus: s1Focus,
    setS1Focus: setS1Focus,
    goToStep2: goToStep2,
    handleStep1TmpSave: handleStep1TmpSave,
    setForm: setForm,
    editIdx: editIdx,
    lastInsp: lastInsp,
    lastDate: lastDate,
    setStep: setStep,
    setView: setView,
    sessionInfo: sessionInfo,
    undoneOnly: undoneOnly,
    setUndoneOnly: setUndoneOnly,
    devColumns: devColumns,
    devVisibleCols: devVisibleCols,
    inspectionMode: inspectionMode,
    checkFields: checkFields,
    setCheck: setCheck,
    handleSave: handleSave,
    complete: complete,
    missing: missing,
    visIn: visIn,
    visOut: visOut,
    visFields: visFields,
    activeCode: activeCode,
    numDisp: numDisp,
    limits: limits,
    onPress: onPress,
    onConfirm: onConfirm,
    onRowClick: onRowClick,
    moveActive: moveActive,
    rowRefs: rowRefs,
    listRef: listRef,
    ALL_FIELDS: ALL_FIELDS,
    vis: vis,
    isAbn: isAbn,
    focusSeq: focusSeq,
    isCheckCode: isCheckCode,
    setCheckAndAdvance: setCheckAndAdvance,
    onSwitchMode: onSwitchMode,
    outdoorLocked: outdoorLocked,
    requestUnlockOutdoor: requestUnlockOutdoor
  })))), view === "list" && /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      display: "flex",
      flexDirection: "column",
      overflow: "hidden",
      padding: "10px 16px"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      flexShrink: 0,
      display: "flex",
      alignItems: "center",
      justifyContent: "space-between",
      marginBottom: 10,
      gap: 8
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 15,
      fontWeight: 800,
      color: C.navy
    }
  }, "\uD83D\uDCCB \u30C7\u30FC\u30BF\u4E00\u89A7"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 8
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: handleReportExportClick,
    title: IS_IPAD ? "iPadではこの機能は使えません" : undefined,
    style: {
      padding: "7px 14px",
      borderRadius: 8,
      border: "none",
      cursor: IS_IPAD ? "not-allowed" : "pointer",
      fontSize: 12,
      fontWeight: 700,
      whiteSpace: "nowrap",
      background: IS_IPAD ? C.g200 : C.navy,
      color: IS_IPAD ? C.g400 : C.white
    }
  }, "\uD83D\uDCC4 ", IS_IPAD ? "点検表出力（iPad不可）" : "点検表出力"), /*#__PURE__*/React.createElement("button", {
    onClick: handleCSVExportClick,
    title: IS_IPAD ? "iPadではこの機能は使えません" : undefined,
    style: {
      padding: "7px 14px",
      borderRadius: 8,
      border: "none",
      cursor: IS_IPAD ? "not-allowed" : "pointer",
      fontSize: 12,
      fontWeight: 700,
      whiteSpace: "nowrap",
      background: IS_IPAD ? C.g200 : filteredRows.length > 0 ? C.teal : C.g200,
      color: IS_IPAD ? C.g400 : filteredRows.length > 0 ? C.white : C.g400
    }
  }, "\uD83D\uDCBE ", IS_IPAD ? "CSV出力（iPad不可）" : "CSV出力"))), /*#__PURE__*/React.createElement("div", {
    style: {
      flexShrink: 0,
      display: "flex",
      alignItems: "center",
      gap: 6,
      marginBottom: 6,
      flexWrap: "wrap"
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 12,
      fontWeight: 700,
      color: C.g600,
      marginRight: 2
    }
  }, devList.length > 0 ? filteredRows.filter(rowDone).length + "/" + tRows.length + "台" : tRows.length + "件"), ["all", "done", "undone"].map(f => /*#__PURE__*/React.createElement("button", {
    key: f,
    onClick: () => setListFilter(f),
    style: {
      padding: "6px 12px",
      borderRadius: 8,
      border: "1.5px solid " + (listFilter === f ? C.blue : C.g200),
      cursor: "pointer",
      fontSize: 12,
      fontWeight: 700,
      background: listFilter === f ? C.blue : C.white,
      color: listFilter === f ? C.white : C.g500
    }
  }, f === "all" ? "すべて" : f === "done" ? "入力済" : "未入力")), /*#__PURE__*/React.createElement("span", {
    style: {
      width: 1,
      height: 20,
      background: C.g200,
      margin: "0 2px"
    }
  }), [["indoor", "🏠 室内機", C.blue], ["outdoor", "🏭 室外機", C.teal]].map(([v, label, col]) => /*#__PURE__*/React.createElement("button", {
    key: v,
    onClick: () => setTypeFilter(p => p === v ? null : v),
    style: {
      padding: "6px 12px",
      borderRadius: 8,
      border: "1.5px solid " + (typeFilter === v ? col : C.g200),
      cursor: "pointer",
      fontSize: 12,
      fontWeight: 700,
      background: typeFilter === v ? col : C.white,
      color: typeFilter === v ? C.white : C.g500,
      whiteSpace: "nowrap"
    }
  }, label))), (() => {
    const buildings = listBuildingKey ? [...new Set(tRows.map(r => colValue(r, "建物")).filter(b => b && b !== "（未設定）"))].sort((a, b) => a.localeCompare(b, undefined, {
      numeric: true
    })) : [];
    if (buildings.length === 0) return null;
    return /*#__PURE__*/React.createElement("div", {
      style: {
        flexShrink: 0,
        display: "flex",
        alignItems: "center",
        gap: 6,
        marginBottom: 6,
        flexWrap: "wrap"
      }
    }, /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 10,
        fontWeight: 700,
        color: C.g400,
        minWidth: 32
      }
    }, "\u5EFA\u7269\uFF1A"), /*#__PURE__*/React.createElement("button", {
      onClick: () => {
        setBuildingFilter(null);
        setFloorFilter(null);
      },
      style: {
        padding: "5px 11px",
        borderRadius: 7,
        border: "1.5px solid " + (!buildingFilter ? C.navy : C.g200),
        cursor: "pointer",
        fontSize: 11,
        fontWeight: 700,
        background: !buildingFilter ? C.navy : C.white,
        color: !buildingFilter ? C.white : C.g600
      }
    }, "\u3059\u3079\u3066"), buildings.map(b => /*#__PURE__*/React.createElement("button", {
      key: b,
      onClick: () => {
        setBuildingFilter(p => {
          const next = p === b ? null : b;
          setFloorFilter(null);
          return next;
        });
      },
      style: {
        padding: "5px 11px",
        borderRadius: 7,
        border: "1.5px solid " + (buildingFilter === b ? C.navy : C.g200),
        cursor: "pointer",
        fontSize: 11,
        fontWeight: 700,
        background: buildingFilter === b ? C.navy : C.white,
        color: buildingFilter === b ? C.white : C.g600,
        whiteSpace: "nowrap"
      }
    }, b)));
  })(), (() => {
    const floorsForBuilding = buildingFilter ? [...new Set(tRows.filter(r => colValue(r, "建物") === buildingFilter).map(r => r.floor).filter(Boolean))].sort() : _floors;
    if (floorsForBuilding.length === 0) return null;
    return /*#__PURE__*/React.createElement("div", {
      style: {
        flexShrink: 0,
        display: "flex",
        alignItems: "center",
        gap: 6,
        marginBottom: 8,
        flexWrap: "wrap"
      }
    }, /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 10,
        fontWeight: 700,
        color: C.g400,
        minWidth: 32
      }
    }, "\u968E\uFF1A"), /*#__PURE__*/React.createElement("button", {
      onClick: () => setFloorFilter(null),
      style: {
        padding: "5px 11px",
        borderRadius: 7,
        border: "1.5px solid " + (!floorFilter ? C.teal : C.g200),
        cursor: "pointer",
        fontSize: 11,
        fontWeight: 700,
        background: !floorFilter ? C.teal : C.white,
        color: !floorFilter ? C.white : C.g600
      }
    }, "\u3059\u3079\u3066"), floorsForBuilding.map(fl => /*#__PURE__*/React.createElement("button", {
      key: fl,
      onClick: () => setFloorFilter(floorFilter === fl ? null : fl),
      style: {
        padding: "5px 11px",
        borderRadius: 7,
        border: "1.5px solid " + (floorFilter === fl ? C.teal : C.g200),
        cursor: "pointer",
        fontSize: 11,
        fontWeight: 700,
        background: floorFilter === fl ? C.teal : C.white,
        color: floorFilter === fl ? C.white : C.g600
      }
    }, fl)));
  })(), (floorFilter || buildingFilter) && /*#__PURE__*/React.createElement("div", {
    style: {
      flexShrink: 0,
      display: "flex",
      alignItems: "center",
      gap: 8,
      marginBottom: 6,
      padding: "8px 14px",
      background: "linear-gradient(90deg," + C.teal + "18," + C.teal + "08)",
      border: "1.5px solid " + C.teal,
      borderRadius: 10
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 13,
      fontWeight: 800,
      color: C.teal,
      flex: 1
    }
  }, "\uD83C\uDFE2 ", [buildingFilter, floorFilter].filter(Boolean).join(" ・ "), " \u306E\u307F\u8868\u793A\u4E2D"), /*#__PURE__*/React.createElement("button", {
    onClick: () => {
      setBuildingFilter(null);
      setFloorFilter(null);
    },
    style: {
      padding: "5px 12px",
      borderRadius: 7,
      border: "1.5px solid " + C.teal,
      background: C.white,
      color: C.teal,
      cursor: "pointer",
      fontSize: 11,
      fontWeight: 700
    }
  }, "\u2715 \u89E3\u9664"), /*#__PURE__*/React.createElement("button", {
    onClick: () => setShowStats(true),
    style: {
      padding: "5px 12px",
      borderRadius: 7,
      border: "none",
      background: C.teal,
      color: C.white,
      cursor: "pointer",
      fontSize: 11,
      fontWeight: 700
    }
  }, "\uD83D\uDCCA \u96C6\u8A08\u3078\u623B\u308B")), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      overflowX: "auto",
      overflowY: "auto",
      scrollbarGutter: "stable",
      background: C.white,
      borderRadius: 12,
      boxShadow: "0 2px 10px rgba(0,0,0,0.07)"
    }
  }, filteredRows.length === 0 ? /*#__PURE__*/React.createElement("div", {
    style: {
      padding: 60,
      textAlign: "center",
      color: C.g400
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 36,
      marginBottom: 8
    }
  }, "\uD83D\uDCED"), /*#__PURE__*/React.createElement("div", {
    style: {
      fontWeight: 700,
      fontSize: 14
    }
  }, listFilter === "undone" ? "未入力の機器はありません" : "データがありません")) : /*#__PURE__*/React.createElement("table", {
    style: {
      width: "100%",
      borderCollapse: "collapse",
      fontSize: 12
    }
  }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", {
    style: {
      position: "sticky",
      top: 0,
      zIndex: 2
    }
  }, [...(devVisibleCols.length > 0 ? devVisibleCols : ["階", "部屋名", "管理番号", "機器番号"]), "入力状況", "点検日", "点検者", "運転", "モード", "風量", "設定温度"].map(h => {
    const isSort = sortCol === h;
    return /*#__PURE__*/React.createElement("th", {
      key: h,
      onClick: () => {
        if (sortCol === h) {
          setSortDir(d => d === "asc" ? "desc" : "asc");
        } else {
          setSortCol(h);
          setSortDir("asc");
        }
      },
      style: {
        background: isSort ? C.blue + "18" : C.g100,
        color: isSort ? C.blue : C.g600,
        padding: "6px 6px",
        textAlign: "center",
        fontWeight: 700,
        fontSize: 10,
        whiteSpace: "normal",
        wordBreak: "keep-all",
        borderBottom: "2px solid " + C.g200,
        cursor: "pointer",
        userSelect: "none",
        maxWidth: 64,
        lineHeight: 1.3,
        verticalAlign: "bottom"
      }
    }, h.length > 4 ? /*#__PURE__*/React.createElement(React.Fragment, null, h.slice(0, Math.ceil(h.length / 2)), /*#__PURE__*/React.createElement("br", null), h.slice(Math.ceil(h.length / 2))) : h, isSort ? sortDir === "asc" ? " ▲" : " ▼" : "");
  }), vf.map((f, fi) => /*#__PURE__*/React.createElement("th", {
    key: f.code,
    style: {
      background: C.g100,
      padding: "8px 6px",
      textAlign: "center",
      fontWeight: 700,
      fontSize: 10,
      whiteSpace: "nowrap",
      borderBottom: "2px solid " + C.g200,
      color: f.group === "indoor" ? C.blue : C.teal,
      borderLeft: fi === 0 ? "2px solid " + C.g200 : undefined
    }
  }, f.code, /*#__PURE__*/React.createElement("br", null), /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 8,
      fontWeight: 400
    }
  }, f.unit))), /*#__PURE__*/React.createElement("th", {
    style: {
      background: C.g100,
      color: C.g600,
      padding: "8px 6px",
      textAlign: "center",
      fontWeight: 700,
      fontSize: 10,
      whiteSpace: "nowrap",
      borderBottom: "2px solid " + C.g200,
      borderLeft: "2px solid " + C.g200
    }
  }, "\u30C1\u30A7\u30C3\u30AF\u7D50\u679C"), /*#__PURE__*/React.createElement("th", {
    style: {
      background: C.g100,
      color: C.g600,
      padding: "8px 8px",
      textAlign: "center",
      fontWeight: 700,
      fontSize: 10,
      borderBottom: "2px solid " + C.g200,
      borderLeft: "2px solid " + C.g200
    }
  }, "\u5099\u8003"))), /*#__PURE__*/React.createElement("tbody", null, filteredRows.map((row, i) => {
    const ir = row.indoorRecord,
      or_ = row.outdoorRecord;
    const indoorDone = rowIndoorDone(row),
      outdoorDone = rowOutdoorDone(row);
    const hasMeasure = indoorDone || outdoorDone;
    const bg = hoverRow === i ? "#EFF6FF" : !hasMeasure ? i % 2 === 0 ? "#FAFAFA" : "#F5F5F5" : i % 2 === 0 ? C.white : C.g50;
    const metaSrc = ir || or_;
    const rType = rowStatsType(row);
    const StatusCell = ({
      rec,
      done,
      mode,
      label
    }) => /*#__PURE__*/React.createElement("td", {
      style: {
        padding: "6px 8px",
        textAlign: "left",
        borderBottom: "1px solid " + C.g100,
        background: bg
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        flexDirection: "row",
        alignItems: "center",
        gap: 6,
        justifyContent: "flex-start"
      }
    }, done ? /*#__PURE__*/React.createElement("span", {
      style: {
        background: C.green + "22",
        color: C.green,
        fontWeight: 700,
        fontSize: 9,
        padding: "2px 6px",
        borderRadius: 5,
        whiteSpace: "nowrap",
        flexShrink: 0
      }
    }, "\u5165\u529B\u6E08") : /*#__PURE__*/React.createElement("span", {
      style: {
        background: "#FEF2F2",
        color: C.red,
        fontWeight: 700,
        fontSize: 9,
        padding: "2px 6px",
        borderRadius: 5,
        whiteSpace: "nowrap",
        flexShrink: 0
      }
    }, "\u672A\u5165\u529B"), rec ? /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        gap: 3,
        flexShrink: 0
      }
    }, /*#__PURE__*/React.createElement("button", {
      onClick: e => {
        e.stopPropagation();
        handleEditFor(rec, mode);
      },
      style: {
        padding: "2px 6px",
        borderRadius: 4,
        border: "none",
        cursor: "pointer",
        background: C.blue,
        color: C.white,
        fontSize: 9,
        fontWeight: 700
      }
    }, "\u7DE8\u96C6"), /*#__PURE__*/React.createElement("button", {
      onClick: e => {
        e.stopPropagation();
        handleDelFor(rec, mode);
      },
      style: {
        padding: "2px 6px",
        borderRadius: 4,
        border: "none",
        cursor: "pointer",
        background: C.red,
        color: C.white,
        fontSize: 9,
        fontWeight: 700
      }
    }, "\u524A\u9664")) : /*#__PURE__*/React.createElement("button", {
      onClick: e => {
        e.stopPropagation();
        handleInputFor(row, mode);
      },
      style: {
        padding: "2px 7px",
        borderRadius: 4,
        border: "1.5px solid " + C.blue,
        cursor: "pointer",
        background: "white",
        color: C.blue,
        fontSize: 9,
        fontWeight: 700,
        whiteSpace: "nowrap",
        flexShrink: 0
      }
    }, label, "\u5165\u529B")));
    const NACell = () => /*#__PURE__*/React.createElement("td", {
      style: {
        padding: "6px 8px",
        textAlign: "left",
        borderBottom: "1px solid " + C.g100,
        background: bg,
        color: C.g300,
        fontSize: 11
      }
    }, "\u2014");
    return /*#__PURE__*/React.createElement("tr", {
      key: i,
      style: {
        opacity: hasMeasure ? 1 : 0.65,
        cursor: "pointer"
      },
      onMouseEnter: () => setHoverRow(i),
      onMouseLeave: () => setHoverRow(p => p === i ? null : p),
      onClick: () => {
        if (rType === "outdoor") {
          handleEditOrInputFor(row, "outdoor");
        } else {
          handleEditOrInputFor(row, "indoor");
        }
      }
    }, (devVisibleCols.length > 0 ? devVisibleCols : ["階", "部屋名", "管理番号", "機器番号"]).map((col, j) => {
      const val = row._raw ? row._raw[col] : col === "階" ? row.floor : col === "部屋名" ? row.room : col === "管理番号" ? row.managementNo : col === "機器番号" ? row.unitNo : "";
      return /*#__PURE__*/React.createElement("td", {
        key: j,
        style: {
          padding: "7px 8px",
          textAlign: "center",
          borderBottom: "1px solid " + C.g100,
          background: bg,
          fontSize: 11,
          whiteSpace: "nowrap",
          fontWeight: 600
        }
      }, val || "—");
    }), rType === "outdoor" ? /*#__PURE__*/React.createElement(StatusCell, {
      rec: or_,
      done: outdoorDone,
      mode: "outdoor",
      label: "\uD83C\uDFED"
    }) : rType === "indoor" ? /*#__PURE__*/React.createElement(StatusCell, {
      rec: ir,
      done: indoorDone,
      mode: "indoor",
      label: "\uD83C\uDFE0"
    }) : /*#__PURE__*/React.createElement(NACell, null), [metaSrc?.inspectionDate, metaSrc?.inspector, metaSrc?.preOperation, metaSrc?.preMode, metaSrc?.preWind, metaSrc?.preSetTemp].map((v, j) => /*#__PURE__*/React.createElement("td", {
      key: j,
      style: {
        padding: "7px 8px",
        textAlign: "center",
        borderBottom: "1px solid " + C.g100,
        background: bg,
        fontSize: 11,
        color: v ? C.g800 : C.g300,
        whiteSpace: "nowrap",
        overflow: "hidden",
        textOverflow: "ellipsis",
        maxWidth: 90
      }
    }, v || "—")), vf.map((f, fi) => {
      const val = rowFieldVal(row, f.code);
      const ab = val ? isAbn(f.code, val, limits) : false;
      return /*#__PURE__*/React.createElement("td", {
        key: f.code,
        style: {
          padding: "7px 7px",
          textAlign: "right",
          borderBottom: "1px solid " + C.g100,
          background: ab ? "#FEF2F2" : bg,
          fontFamily: "monospace",
          fontSize: 11,
          color: ab ? C.red : val ? C.g800 : C.g300,
          borderLeft: fi === 0 ? "2px solid " + C.g200 : undefined
        }
      }, val || "—");
    }), /*#__PURE__*/React.createElement("td", {
      style: {
        padding: "6px 6px",
        textAlign: "center",
        borderBottom: "1px solid " + C.g100,
        background: bg,
        borderLeft: "2px solid " + C.g200,
        maxWidth: 180
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        flexWrap: "nowrap",
        gap: 2,
        justifyContent: "flex-start",
        overflowX: "auto",
        maxWidth: 170
      }
    }, rowChecksList(row).map(c => /*#__PURE__*/React.createElement("span", {
      key: c.code,
      title: c.label,
      style: {
        display: "inline-flex",
        flexShrink: 0,
        alignItems: "center",
        justifyContent: "center",
        width: 16,
        height: 16,
        borderRadius: 3,
        fontSize: 10,
        fontWeight: 800,
        background: c.val === "○" ? C.green + "22" : c.val === "×" ? C.red + "22" : C.g100,
        color: c.val === "○" ? C.green : c.val === "×" ? C.red : C.g300
      }
    }, c.val || "—")))), /*#__PURE__*/React.createElement("td", {
      style: {
        padding: "7px 8px",
        textAlign: "center",
        borderBottom: "1px solid " + C.g100,
        background: bg,
        fontSize: 11,
        borderLeft: "2px solid " + C.g200,
        color: rowRemarks(row) ? C.g800 : C.g300,
        whiteSpace: "nowrap",
        overflow: "hidden",
        textOverflow: "ellipsis",
        maxWidth: 150
      }
    }, rowRemarks(row) || "—"));
  }))))), view === "settings" && /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      overflowY: "auto",
      scrollbarGutter: "stable",
      padding: "14px 16px 28px",
      display: "flex",
      flexDirection: "column",
      gap: 12,
      maxWidth: 840,
      margin: "0 auto",
      width: "100%",
      boxSizing: "border-box"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: 10,
      padding: "10px 14px",
      background: C.teal + "10",
      border: "1.5px solid " + C.teal + "40",
      borderRadius: 12
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 18
    }
  }, "\uD83D\uDCBE"), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      fontSize: 11,
      color: C.g600,
      lineHeight: 1.5
    }
  }, "\u6A5F\u5668\u30EA\u30B9\u30C8\u30FB\u70B9\u691C\u8005\u30EA\u30B9\u30C8\u30FB\u70B9\u691C\u9805\u76EE\u30FB\u8868\u793A\u8A2D\u5B9A\u30FB\u6B63\u5E38\u5024\u7BC4\u56F2\u306F\u3001\u3053\u306E\u30D6\u30E9\u30A6\u30B6\u306B\u81EA\u52D5\u4FDD\u5B58\u3055\u308C\u307E\u3059\u3002", /*#__PURE__*/React.createElement("br", null), "\u30A2\u30D7\u30EA\u3092\u66F4\u65B0\uFF08\u518D\u8AAD\u8FBC\uFF09\u3057\u3066\u3082\u8AAD\u307F\u8FBC\u307F\u76F4\u3059\u5FC5\u8981\u306F\u3042\u308A\u307E\u305B\u3093\u3002"), /*#__PURE__*/React.createElement("button", {
    onClick: () => {
      if (!window.confirm("機器リスト・点検者リスト・点検項目・表示設定・正常値範囲の保存データを全て削除します。よろしいですか？")) return;
      Object.values(LS_KEYS).forEach(k => {
        try {
          localStorage.removeItem(k);
        } catch (e) {}
      });
      setDevList([]);
      setDevColumns([]);
      setDevVisibleCols([]);
      setInspList([]);
      setLimits(defLim());
      setVis(defVis());
      setCardLabels(defCardLabels());
      setCheckFields([{
        code: "ci1",
        label: "配管類支持異常の有無",
        category: "据付状態",
        group: "check_in"
      }, {
        code: "ci2",
        label: "異音・異常振動の有無",
        category: "フィルター点検・清掃",
        group: "check_in"
      }, {
        code: "ci3",
        label: "異音・異常振動の有無",
        category: "運転確認",
        group: "check_in"
      }, {
        code: "co1",
        label: "防振装置異常の有無",
        category: "据付状態",
        group: "check_out"
      }, {
        code: "co2",
        label: "配管類支持異常の有無",
        category: "据付状態",
        group: "check_out"
      }, {
        code: "co3",
        label: "ガスリークテスト",
        category: "冷媒系統",
        group: "check_out"
      }, {
        code: "co4",
        label: "配管系統外観点検",
        category: "冷媒系統",
        group: "check_out"
      }, {
        code: "co5",
        label: "異音・異常振動の有無",
        category: "送排風機系統",
        group: "check_out"
      }, {
        code: "co6",
        label: "ドレン配管異常の有無",
        category: "排水系統",
        group: "check_out"
      }, {
        code: "co7",
        label: "フィン汚れの有無",
        category: "熱交換器系統",
        group: "check_out"
      }, {
        code: "co8",
        label: "異音・異常振動の有無",
        category: "熱交換器系統",
        group: "check_out"
      }, {
        code: "co9",
        label: "外面清掃",
        category: "作業終了時",
        group: "check_out"
      }]);
      showFlash("🗑️ 保存データを削除しました");
    },
    style: {
      flexShrink: 0,
      padding: "6px 10px",
      borderRadius: 7,
      border: "1.5px solid " + C.g300,
      cursor: "pointer",
      background: C.white,
      color: C.g500,
      fontSize: 11,
      fontWeight: 700,
      whiteSpace: "nowrap"
    }
  }, "\uD83D\uDDD1\uFE0F \u4FDD\u5B58\u30C7\u30FC\u30BF\u3092\u524A\u9664")), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gridTemplateColumns: "repeat(3,1fr)",
      gap: 10
    }
  }, [{
    id: "device",
    title: cardLabels.device || "機器リスト",
    icon: "📂",
    color: C.navy,
    badge: devList.length > 0 ? devList.length + "件" : null,
    editable: true
  }, {
    id: "cols",
    title: "表示列設定",
    icon: "🗂️",
    color: C.teal,
    badge: devVisibleCols.length > 0 ? devVisibleCols.length + "列" : devColumns.length > 0 ? "全列" : null
  }, {
    id: "criteria",
    title: cardLabels.criteria || "点検基準設定",
    icon: "🎯",
    color: C.purple,
    badge: checkFields.length > 0 ? checkFields.length + "項目" : null,
    editable: true
  }, {
    id: "inspector",
    title: cardLabels.inspector || "点検者リスト",
    icon: "👤",
    color: C.blue,
    badge: inspList.length > 0 ? inspList.length + "名" : null,
    editable: true
  }, {
    id: "vis",
    title: "表示項目",
    icon: "👁️",
    color: C.teal,
    badge: null
  }].map(({
    id,
    title,
    icon,
    color,
    badge,
    editable
  }) => /*#__PURE__*/React.createElement("div", {
    key: id,
    style: {
      position: "relative"
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: () => setModalSec(id),
    style: {
      width: "100%",
      boxSizing: "border-box",
      background: C.white,
      borderRadius: 14,
      boxShadow: "0 2px 10px rgba(0,0,0,0.07)",
      border: "none",
      cursor: "pointer",
      padding: "18px 10px",
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      gap: 8,
      transition: "all 0.15s"
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 32
    }
  }, icon), /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 13,
      fontWeight: 700,
      color: color
    }
  }, title), badge && /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 11,
      background: color,
      color: C.white,
      padding: "2px 8px",
      borderRadius: 8,
      fontWeight: 700
    }
  }, badge)), editable && /*#__PURE__*/React.createElement("button", {
    onMouseDown: e => e.preventDefault(),
    onClick: e => {
      e.stopPropagation();
      const next = window.prompt("表示名を入力してください", title);
      if (next && next.trim()) setCardLabels(p => ({
        ...p,
        [id]: next.trim()
      }));
    },
    style: {
      position: "absolute",
      top: 6,
      right: 6,
      width: 26,
      height: 26,
      borderRadius: 13,
      border: "none",
      background: C.g100,
      color: C.g500,
      cursor: "pointer",
      fontSize: 12,
      display: "flex",
      alignItems: "center",
      justifyContent: "center"
    }
  }, "\u270F\uFE0F")))), modalSec && /*#__PURE__*/React.createElement("div", {
    style: {
      position: "fixed",
      inset: 0,
      background: "rgba(0,0,0,0.55)",
      zIndex: 10000,
      display: "flex",
      alignItems: "flex-end",
      justifyContent: "center"
    },
    onClick: () => setModalSec(null)
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: C.white,
      borderRadius: "20px 20px 0 0",
      width: "100%",
      maxWidth: 680,
      maxHeight: "80vh",
      display: "flex",
      flexDirection: "column",
      boxShadow: "0 -4px 30px rgba(0,0,0,0.2)"
    },
    onClick: e => e.stopPropagation()
  }, (() => {
    const sec = {
      device: {
        title: (cardLabels.device || "機器リスト") + " CSV読込",
        icon: "📂",
        color: C.navy
      },
      cols: {
        title: "表示列設定",
        icon: "🗂️",
        color: C.teal
      },
      criteria: {
        title: cardLabels.criteria || "点検基準設定",
        icon: "🎯",
        color: C.purple
      },
      inspector: {
        title: (cardLabels.inspector || "点検者リスト") + " CSV読込",
        icon: "👤",
        color: C.blue
      },
      vis: {
        title: "表示項目設定",
        icon: "👁️",
        color: C.teal
      }
    }[modalSec] || {};
    return /*#__PURE__*/React.createElement("div", {
      style: {
        background: "linear-gradient(135deg," + C.navy + "," + C.blue + ")",
        borderRadius: "20px 20px 0 0",
        padding: "14px 18px",
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        flexShrink: 0
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        alignItems: "center",
        gap: 10
      }
    }, /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 20
      }
    }, sec.icon), /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 15,
        fontWeight: 800,
        color: C.white
      }
    }, sec.title)), /*#__PURE__*/React.createElement("button", {
      onClick: () => setModalSec(null),
      style: {
        background: "rgba(255,255,255,0.2)",
        border: "none",
        color: C.white,
        borderRadius: 8,
        width: 32,
        height: 32,
        cursor: "pointer",
        fontSize: 18,
        fontWeight: 700,
        display: "flex",
        alignItems: "center",
        justifyContent: "center"
      }
    }, "\u2715"));
  })(), /*#__PURE__*/React.createElement("div", {
    style: {
      overflowY: "auto",
      scrollbarGutter: "stable",
      padding: "18px 20px",
      flex: 1
    }
  }, [{
    id: "device"
  }, {
    id: "cols"
  }, {
    id: "criteria"
  }, {
    id: "inspector"
  }, {
    id: "vis"
  }].map(({
    id
  }) => modalSec === id && /*#__PURE__*/React.createElement("div", {
    key: id
  }, id === "device" && /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 14
    }
  }, /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("p", {
    style: {
      fontSize: 13,
      color: C.g500,
      marginBottom: 10,
      lineHeight: 1.6
    }
  }, "1\u884C\u76EE\u3092\u9805\u76EE\u540D\u3068\u3057\u3066\u81EA\u52D5\u8A8D\u8B58\u3057\u307E\u3059\u3002"), /*#__PURE__*/React.createElement("input", {
    ref: devRef,
    type: "file",
    accept: ".csv,.xlsx,.xls",
    style: {
      display: "none"
    },
    onChange: handleDevCSV
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 10,
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: () => devRef.current.click(),
    style: {
      padding: "10px 20px",
      borderRadius: 9,
      border: "none",
      cursor: "pointer",
      background: C.blue,
      color: C.white,
      fontWeight: 700,
      fontSize: 13
    }
  }, "\uD83D\uDCC1 CSV\u3092\u9078\u629E"), devList.length > 0 && /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 13,
      color: C.green,
      fontWeight: 700
    }
  }, "\u2705 ", devList.length, "\u4EF6\u8AAD\u8FBC\u6E08"))), devList.length > 0 && /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 11,
      fontWeight: 700,
      color: C.g500,
      marginBottom: 6
    }
  }, "\u30D7\u30EC\u30D3\u30E5\u30FC\uFF08\u5148\u982D5\u4EF6\uFF09"), /*#__PURE__*/React.createElement("div", {
    style: {
      overflowX: "auto"
    }
  }, /*#__PURE__*/React.createElement("table", {
    style: {
      width: "100%",
      borderCollapse: "collapse",
      fontSize: 12
    }
  }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", null, (devVisibleCols.length > 0 ? devVisibleCols : devColumns).map(h => /*#__PURE__*/React.createElement("th", {
    key: h,
    style: {
      background: C.g100,
      padding: "6px 8px",
      textAlign: "center",
      fontWeight: 700,
      fontSize: 11,
      color: C.g500,
      borderBottom: "2px solid " + C.g200,
      whiteSpace: "nowrap"
    }
  }, h)))), /*#__PURE__*/React.createElement("tbody", null, devList.slice(0, 5).map((d, i) => {
    const cols = devVisibleCols.length > 0 ? devVisibleCols : devColumns;
    return /*#__PURE__*/React.createElement("tr", {
      key: i
    }, cols.map((col, j) => {
      const val = d._raw ? d._raw[col] : d[{
        階: "floor",
        部屋名: "room",
        管理番号: "managementNo",
        機器番号: "unitNo"
      }[col] || col] || "";
      return /*#__PURE__*/React.createElement("td", {
        key: j,
        style: {
          padding: "6px 8px",
          textAlign: "center",
          borderBottom: "1px solid " + C.g100,
          background: i % 2 === 0 ? C.white : C.g50
        }
      }, val);
    }));
  }), devList.length > 5 && /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("td", {
    colSpan: devVisibleCols.length || devColumns.length || 4,
    style: {
      padding: "6px",
      textAlign: "center",
      color: C.g400,
      fontSize: 11
    }
  }, "\u2026\u4ED6 ", devList.length - 5, "\u4EF6")))))), /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 11,
      fontWeight: 700,
      color: C.g500,
      marginBottom: 6
    }
  }, "1\u4EF6\u3092\u691C\u7D22\u3057\u3066\u7DE8\u96C6\u30FB\u8FFD\u52A0\u30FB\u524A\u9664"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 8,
      alignItems: "center",
      marginBottom: 8
    }
  }, /*#__PURE__*/React.createElement("input", {
    value: devEditSearch,
    onChange: e => setDevEditSearch(e.target.value),
    placeholder: "\u7BA1\u7406\u756A\u53F7\u30FB\u90E8\u5C4B\u540D\u306A\u3069\u3067\u691C\u7D22\u2026",
    style: {
      flex: 1,
      padding: "9px 12px",
      borderRadius: 8,
      border: "1.5px solid " + C.g200,
      fontSize: 13,
      outline: "none",
      boxSizing: "border-box"
    }
  }), /*#__PURE__*/React.createElement("button", {
    onClick: () => {
      const cols = devColumns.length > 0 ? devColumns : ["階", "部屋名", "管理番号", "機器番号"];
      setDevEditIdx("new");
      setDevEditDraft({
        _raw: Object.fromEntries(cols.map(c => [c, ""]))
      });
    },
    style: {
      padding: "9px 14px",
      borderRadius: 8,
      border: "none",
      background: C.green,
      color: C.white,
      cursor: "pointer",
      fontSize: 12,
      fontWeight: 700,
      whiteSpace: "nowrap"
    }
  }, "\uFF0B \u65B0\u898F\u8FFD\u52A0")), devEditSearch.trim() && (() => {
    const q = devEditSearch.trim().toLowerCase();
    const cols = devColumns.length > 0 ? devColumns : [];
    const matches = devList.map((d, i) => ({
      d,
      i
    })).filter(({
      d
    }) => {
      const hay = (cols.map(c => String(d._raw?.[c] || "")).join(" ") + " " + [d.floor, d.room, d.managementNo, d.unitNo].join(" ")).toLowerCase();
      return hay.includes(q);
    });
    if (matches.length === 0) return /*#__PURE__*/React.createElement("div", {
      style: {
        fontSize: 12,
        color: C.g400,
        padding: "8px 4px"
      }
    }, "\u8A72\u5F53\u3059\u308B\u6A5F\u5668\u304C\u3042\u308A\u307E\u305B\u3093");
    return /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        flexDirection: "column",
        gap: 3,
        maxHeight: 220,
        overflowY: "auto",
        scrollbarGutter: "stable",
        marginBottom: 10
      }
    }, matches.slice(0, 30).map(({
      d,
      i
    }) => /*#__PURE__*/React.createElement("button", {
      key: i,
      onClick: () => {
        setDevEditIdx(i);
        setDevEditDraft({
          _raw: {
            ...d._raw
          }
        });
      },
      style: {
        textAlign: "left",
        padding: "7px 10px",
        borderRadius: 7,
        border: "1.5px solid " + (devEditIdx === i ? C.blue : C.g200),
        background: devEditIdx === i ? "#EFF6FF" : C.white,
        cursor: "pointer",
        fontSize: 12,
        color: C.g700
      }
    }, d.floor, " ", d.room, " ", d.managementNo, "/", d.unitNo)), matches.length > 30 && /*#__PURE__*/React.createElement("div", {
      style: {
        fontSize: 11,
        color: C.g400,
        textAlign: "center",
        padding: "4px 0"
      }
    }, "\u307B\u304B ", matches.length - 30, "\u4EF6\u2026\u7D5E\u308A\u8FBC\u3093\u3067\u304F\u3060\u3055\u3044"));
  })(), devEditDraft && /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "12px 14px",
      background: C.g50,
      borderRadius: 10,
      border: "1.5px solid " + C.g200
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12,
      fontWeight: 700,
      color: C.navy,
      marginBottom: 10
    }
  }, devEditIdx === "new" ? "🆕 新規機器" : "✏️ 編集中"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 8,
      marginBottom: 12
    }
  }, (devColumns.length > 0 ? devColumns : ["階", "部屋名", "管理番号", "機器番号"]).map(col => /*#__PURE__*/React.createElement("div", {
    key: col,
    style: {
      display: "flex",
      alignItems: "center",
      gap: 8
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 12,
      color: C.g500,
      width: 90,
      flexShrink: 0
    }
  }, col), /*#__PURE__*/React.createElement("input", {
    value: devEditDraft._raw[col] || "",
    onChange: e => setDevEditDraft(p => ({
      ...p,
      _raw: {
        ...p._raw,
        [col]: e.target.value
      }
    })),
    style: {
      flex: 1,
      padding: "7px 10px",
      borderRadius: 7,
      border: "1.5px solid " + C.g200,
      fontSize: 13,
      outline: "none",
      boxSizing: "border-box"
    }
  })))), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 8
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: () => {
      const cols = devColumns.length > 0 ? devColumns : Object.keys(devEditDraft._raw);
      const core = deriveDevCore(devEditDraft._raw, cols);
      if (!core.managementNo && !core.unitNo) {
        showFlash("⚠️ 管理番号または機器番号を入力してください");
        return;
      }
      const finalDev = {
        ...core,
        _raw: {
          ...devEditDraft._raw
        }
      };
      if (devEditIdx === "new") {
        setDevList(p => [...p, finalDev]);
        if (devColumns.length === 0) setDevColumns(cols);
      } else {
        setDevList(p => p.map((d, i) => i === devEditIdx ? finalDev : d));
      }
      setDevEditIdx(null);
      setDevEditDraft(null);
      setDevEditSearch("");
      showFlash("✅ 保存しました");
    },
    style: {
      flex: 1,
      padding: "10px",
      borderRadius: 8,
      border: "none",
      background: C.green,
      color: C.white,
      cursor: "pointer",
      fontWeight: 700,
      fontSize: 13
    }
  }, "\uD83D\uDCBE \u4FDD\u5B58"), devEditIdx !== "new" && /*#__PURE__*/React.createElement("button", {
    onClick: () => {
      if (!window.confirm("この機器を削除しますか？")) return;
      setDevList(p => p.filter((_, i) => i !== devEditIdx));
      setDevEditIdx(null);
      setDevEditDraft(null);
      showFlash("🗑️ 削除しました");
    },
    style: {
      padding: "10px 16px",
      borderRadius: 8,
      border: "1.5px solid " + C.red,
      background: "#FEF2F2",
      color: C.red,
      cursor: "pointer",
      fontWeight: 700,
      fontSize: 13
    }
  }, "\uD83D\uDDD1\uFE0F \u524A\u9664"), /*#__PURE__*/React.createElement("button", {
    onClick: () => {
      setDevEditIdx(null);
      setDevEditDraft(null);
    },
    style: {
      padding: "10px 16px",
      borderRadius: 8,
      border: "1.5px solid " + C.g300,
      background: C.white,
      color: C.g500,
      cursor: "pointer",
      fontWeight: 700,
      fontSize: 13
    }
  }, "\u30AD\u30E3\u30F3\u30BB\u30EB"))))), id === "cols" && /*#__PURE__*/React.createElement("div", null, devColumns.length === 0 ? /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "12px 16px",
      background: "#FFF7ED",
      borderRadius: 8,
      fontSize: 13,
      color: "#92400E",
      border: "1.5px solid #F59E0B"
    }
  }, "\u26A0\uFE0F \u5148\u306B\u6A5F\u5668\u30EA\u30B9\u30C8CSV\u3092\u8AAD\u307F\u8FBC\u3093\u3067\u304F\u3060\u3055\u3044\u3002\u8AAD\u307F\u8FBC\u307F\u5F8C\u306B\u5217\u304C\u81EA\u52D5\u62BD\u51FA\u3055\u308C\u307E\u3059\u3002") : /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("p", {
    style: {
      fontSize: 12,
      color: C.g500,
      marginBottom: 10,
      lineHeight: 1.6
    }
  }, "\u30C1\u30A7\u30C3\u30AF\u3057\u305F\u5217\u3092\u300C\u5165\u5BA4\u53EF\u5426\u30C1\u30A7\u30C3\u30AF\u300D\u3068\u300C\u6A5F\u5668\u9078\u629E\u30EA\u30B9\u30C8\u300D\u306B\u8868\u793A\u3057\u307E\u3059\u3002", /*#__PURE__*/React.createElement("br", null), "\u203B \u300C\u672C\u65E5\u306E\u5BFE\u8C61\u30A8\u30EA\u30A2\u300D\u306E\u968E\u30DC\u30BF\u30F3\u306B\u306F\u5F71\u97FF\u3057\u307E\u305B\u3093\u3002"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 6,
      marginBottom: 10
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: () => setDevVisibleCols([...devColumns]),
    style: {
      padding: "5px 12px",
      borderRadius: 6,
      border: "1.5px solid " + C.teal,
      background: C.teal + "18",
      color: C.teal,
      fontSize: 11,
      fontWeight: 700,
      cursor: "pointer"
    }
  }, "\u5168\u9078\u629E"), /*#__PURE__*/React.createElement("button", {
    onClick: () => setDevVisibleCols([]),
    style: {
      padding: "5px 12px",
      borderRadius: 6,
      border: "1.5px solid " + C.g300,
      background: C.g100,
      color: C.g500,
      fontSize: 11,
      fontWeight: 700,
      cursor: "pointer"
    }
  }, "\u5168\u89E3\u9664")), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 4
    }
  }, devColumns.map(col => {
    const active = devVisibleCols.includes(col);
    return /*#__PURE__*/React.createElement("label", {
      key: col,
      style: {
        display: "flex",
        alignItems: "center",
        gap: 10,
        padding: "9px 14px",
        borderRadius: 9,
        background: active ? C.teal + "0A" : C.g50,
        border: "1.5px solid " + (active ? C.teal : C.g200),
        cursor: "pointer"
      }
    }, /*#__PURE__*/React.createElement("input", {
      type: "checkbox",
      checked: active,
      onChange: e => {
        setDevVisibleCols(p => e.target.checked ? [...p, col] : p.filter(c => c !== col));
      },
      style: {
        width: 16,
        height: 16,
        accentColor: C.teal,
        cursor: "pointer"
      }
    }), /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 13,
        fontWeight: active ? 700 : 400,
        color: active ? C.navy : C.g600,
        flex: 1
      }
    }, col), /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 11,
        color: C.g400
      }
    }, devList.filter(d => d._raw && d._raw[col]).length, "\u4EF6\u306B\u5024\u3042\u308A"));
  })))), id === "criteria" && /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 8,
      marginBottom: 16
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: () => setCriteriaTab("checkitems"),
    style: {
      flex: 1,
      padding: "10px",
      borderRadius: 9,
      border: "2px solid " + (criteriaTab === "checkitems" ? "#059669" : C.g200),
      background: criteriaTab === "checkitems" ? "#05966912" : C.white,
      color: criteriaTab === "checkitems" ? "#059669" : C.g500,
      fontWeight: 800,
      fontSize: 13,
      cursor: "pointer"
    }
  }, "\u2705 \u70B9\u691C\u9805\u76EE"), /*#__PURE__*/React.createElement("button", {
    onClick: () => setCriteriaTab("lim"),
    style: {
      flex: 1,
      padding: "10px",
      borderRadius: 9,
      border: "2px solid " + (criteriaTab === "lim" ? C.purple : C.g200),
      background: criteriaTab === "lim" ? C.purple + "12" : C.white,
      color: criteriaTab === "lim" ? C.purple : C.g500,
      fontWeight: 800,
      fontSize: 13,
      cursor: "pointer"
    }
  }, "\u2699\uFE0F \u6B63\u5E38\u5024\u7BC4\u56F2")), id === "criteria" && criteriaTab === "checkitems" && /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("p", {
    style: {
      fontSize: 13,
      color: C.g500,
      marginBottom: 10,
      lineHeight: 1.6
    }
  }, "Excel\u30D5\u30A1\u30A4\u30EB\u3092\u8AAD\u307F\u8FBC\u307F\u307E\u3059\u3002", /*#__PURE__*/React.createElement("br", null), "1\u5217\u76EE\u306B\u9805\u76EE\u540D\u30012\u5217\u76EE\u306B\u30AB\u30C6\u30B4\u30EA\uFF08\u7701\u7565\u53EF\uFF09\u3092\u7E26\u306B\u8A18\u8F09\u3057\u3066\u304F\u3060\u3055\u3044\u3002", /*#__PURE__*/React.createElement("br", null), "\u4F8B\uFF1A", /*#__PURE__*/React.createElement("code", {
    style: {
      background: C.g100,
      padding: "1px 6px",
      borderRadius: 4,
      fontSize: 12
    }
  }, "\u914D\u7BA1\u652F\u6301\u7570\u5E38\u306E\u6709\u7121 / \u5BA4\u5185\u6A5F")), /*#__PURE__*/React.createElement("input", {
    type: "file",
    accept: ".xlsx,.xls,.csv",
    style: {
      display: "none"
    },
    id: "checkItemsFile",
    onChange: e => {
      const f = e.target.files[0];
      if (!f) return;
      readFile(f, (data, cols) => {
        let rows = [];
        if (Array.isArray(data)) {
          rows = data;
        } else {
          const lines = (data || "").trim().split(/\r?\n/).filter(Boolean);
          rows = lines.map(l => {
            const c = l.split(",").map(s => s.trim().replace(/^"|"$/g, ""));
            return {
              label: c[0] || "",
              category: c[1] || ""
            };
          });
        }
        const fields = rows.map((row, i) => {
          const label = String(Array.isArray(data) ? Object.values(row)[0] : row.label || "").trim();
          const category = String(Array.isArray(data) ? Object.values(row)[1] || "" : row.category || "").trim();
          const group = /室内|indoor/i.test(category) ? "check_in" : "check_out";
          return label ? {
            code: "ck" + (i + 1),
            label,
            category: category || "点検",
            group
          } : null;
        }).filter(Boolean);
        setCheckFields(fields);
        showFlash("✅ 点検項目 " + fields.length + "件 読込");
      });
      e.target.value = "";
    }
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 10,
      alignItems: "center",
      marginBottom: 12
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: () => document.getElementById("checkItemsFile").click(),
    style: {
      padding: "10px 20px",
      borderRadius: 9,
      border: "none",
      cursor: "pointer",
      background: "#059669",
      color: C.white,
      fontWeight: 700,
      fontSize: 13
    }
  }, "\uD83D\uDCC1 Excel\u3092\u9078\u629E"), checkFields.length > 0 && /*#__PURE__*/React.createElement("button", {
    onClick: () => setCheckFields([]),
    style: {
      padding: "8px 14px",
      borderRadius: 8,
      border: "1.5px solid " + C.red,
      background: "#FEF2F2",
      color: C.red,
      cursor: "pointer",
      fontSize: 12,
      fontWeight: 700
    }
  }, "\uD83D\uDDD1\uFE0F \u30AF\u30EA\u30A2"), checkFields.length > 0 && /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 13,
      color: C.green,
      fontWeight: 700
    }
  }, "\u2705 ", checkFields.length, "\u9805\u76EE\u8AAD\u8FBC\u6E08")), /*#__PURE__*/React.createElement("p", {
    style: {
      fontSize: 12,
      color: C.g500,
      marginBottom: 8
    }
  }, "\u4E0B\u306E\u4E00\u89A7\u304B\u3089\u76F4\u63A5\u3001\u8FFD\u52A0\u30FB\u4FEE\u6B63\u30FB\u524A\u9664\u3082\u3067\u304D\u307E\u3059\u3002"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 2,
      maxHeight: 340,
      overflowY: "auto",
      scrollbarGutter: "stable"
    }
  }, [{
    group: "check_in",
    label: "室内機（インドア）",
    color: C.blue
  }, {
    group: "check_out",
    label: "室外機（アウトドア）",
    color: C.teal
  }].map(({
    group,
    label,
    color
  }) => /*#__PURE__*/React.createElement("div", {
    key: group,
    style: {
      marginBottom: 18
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontWeight: 700,
      fontSize: 13,
      color,
      borderBottom: "2px solid " + color,
      paddingBottom: 5,
      marginBottom: 6
    }
  }, label), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 4
    }
  }, checkFields.map((f, i) => f.group === group && /*#__PURE__*/React.createElement("div", {
    key: f.code,
    style: {
      display: "flex",
      gap: 6,
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement("input", {
    value: f.category,
    onChange: e => setCheckFields(p => p.map((x, j) => j === i ? {
      ...x,
      category: e.target.value
    } : x)),
    placeholder: "\u30AB\u30C6\u30B4\u30EA",
    style: {
      width: 110,
      padding: "7px 8px",
      borderRadius: 6,
      border: "1.5px solid " + C.g200,
      fontSize: 12,
      outline: "none",
      boxSizing: "border-box",
      flexShrink: 0
    }
  }), /*#__PURE__*/React.createElement("input", {
    value: f.label,
    onChange: e => setCheckFields(p => p.map((x, j) => j === i ? {
      ...x,
      label: e.target.value
    } : x)),
    placeholder: "\u70B9\u691C\u5185\u5BB9",
    style: {
      flex: 1,
      padding: "7px 8px",
      borderRadius: 6,
      border: "1.5px solid " + C.g200,
      fontSize: 13,
      outline: "none",
      boxSizing: "border-box",
      minWidth: 0
    }
  }), /*#__PURE__*/React.createElement("button", {
    onClick: () => {
      if (window.confirm("この項目を削除しますか？")) setCheckFields(p => p.filter((_, j) => j !== i));
    },
    style: {
      padding: "6px 10px",
      borderRadius: 6,
      border: "none",
      background: "#FEF2F2",
      color: C.red,
      cursor: "pointer",
      fontSize: 11,
      fontWeight: 700,
      flexShrink: 0
    }
  }, "\uD83D\uDDD1\uFE0F")))), /*#__PURE__*/React.createElement("button", {
    onClick: () => setCheckFields(p => [...p, {
      code: "ck" + Date.now(),
      label: "",
      category: "",
      group
    }]),
    style: {
      marginTop: 8,
      padding: "7px 12px",
      borderRadius: 7,
      border: "1.5px dashed " + C.g300,
      background: "none",
      color,
      cursor: "pointer",
      fontSize: 12,
      fontWeight: 700
    }
  }, "\uFF0B ", label, "\u306E\u9805\u76EE\u3092\u8FFD\u52A0"))))), id === "inspector" && /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("p", {
    style: {
      fontSize: 13,
      color: C.g500,
      marginBottom: 12
    }
  }, "\u30D5\u30A9\u30FC\u30DE\u30C3\u30C8\uFF1A1\u884C\u306B1\u540D\uFF08\u30D8\u30C3\u30C0\u30FC\u306A\u3057\uFF09\u3002\u4E0B\u306E\u4E00\u89A7\u304B\u3089\u76F4\u63A5\u3001\u8FFD\u52A0\u30FB\u4FEE\u6B63\u30FB\u524A\u9664\u3082\u3067\u304D\u307E\u3059\u3002"), /*#__PURE__*/React.createElement("input", {
    ref: inspRef,
    type: "file",
    accept: ".csv,.xlsx,.xls",
    style: {
      display: "none"
    },
    onChange: handleInspCSV
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 10,
      alignItems: "center",
      flexWrap: "wrap",
      marginBottom: 14
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: () => inspRef.current.click(),
    style: {
      padding: "11px 22px",
      borderRadius: 9,
      border: "none",
      cursor: "pointer",
      background: C.blue,
      color: C.white,
      fontWeight: 700,
      fontSize: 14
    }
  }, "\uD83D\uDCC1 CSV\u3092\u9078\u629E"), inspList.length > 0 && /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 13,
      color: C.green,
      fontWeight: 700
    }
  }, "\u2705 ", inspList.length, "\u540D")), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 6
    }
  }, inspList.map((name, i) => /*#__PURE__*/React.createElement("div", {
    key: i,
    style: {
      display: "flex",
      gap: 8,
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement("input", {
    value: name,
    onChange: e => setInspList(p => p.map((n, j) => j === i ? e.target.value : n)),
    placeholder: "\u70B9\u691C\u8005\u540D",
    style: {
      flex: 1,
      padding: "9px 12px",
      borderRadius: 8,
      border: "1.5px solid " + C.g200,
      fontSize: 14,
      outline: "none",
      boxSizing: "border-box"
    }
  }), /*#__PURE__*/React.createElement("button", {
    onClick: () => {
      if (window.confirm("この点検者を削除しますか？")) setInspList(p => p.filter((_, j) => j !== i));
    },
    style: {
      padding: "8px 12px",
      borderRadius: 7,
      border: "none",
      background: "#FEF2F2",
      color: C.red,
      cursor: "pointer",
      fontSize: 12,
      fontWeight: 700,
      flexShrink: 0
    }
  }, "\uD83D\uDDD1\uFE0F"))), /*#__PURE__*/React.createElement("button", {
    onClick: () => setInspList(p => [...p, ""]),
    style: {
      marginTop: 4,
      padding: "9px 14px",
      borderRadius: 8,
      border: "1.5px dashed " + C.g300,
      background: "none",
      color: C.blue,
      cursor: "pointer",
      fontSize: 13,
      fontWeight: 700
    }
  }, "\uFF0B \u70B9\u691C\u8005\u3092\u8FFD\u52A0"))), id === "vis" && /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("p", {
    style: {
      fontSize: 13,
      color: C.g500,
      marginBottom: 14
    }
  }, "\u30C1\u30A7\u30C3\u30AF\u3092\u5916\u3057\u305F\u9805\u76EE\u306F\u5165\u529B\u30D5\u30A9\u30FC\u30E0\u30FB\u4E00\u89A7\u304B\u3089\u975E\u8868\u793A\u306B\u306A\u308A\u307E\u3059\u3002"), [{
    fields: INDOOR_FIELDS,
    label: "室内機（インドア）",
    color: C.blue
  }, {
    fields: OUTDOOR_FIELDS,
    label: "室外機（アウトドア）",
    color: C.teal
  }].map(({
    fields,
    label,
    color
  }) => /*#__PURE__*/React.createElement("div", {
    key: label,
    style: {
      marginBottom: 18
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontWeight: 700,
      fontSize: 13,
      color,
      borderBottom: "2px solid " + color,
      paddingBottom: 5,
      marginBottom: 10
    }
  }, label), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 4
    }
  }, fields.map(f => /*#__PURE__*/React.createElement("label", {
    key: f.code,
    style: {
      display: "flex",
      alignItems: "center",
      gap: 14,
      padding: "10px 14px",
      borderRadius: 9,
      background: tmpVis[f.code] ? color + "0A" : C.g50,
      border: "1.5px solid " + (tmpVis[f.code] ? color : C.g200),
      cursor: "pointer"
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: tmpVis[f.code],
    onChange: e => setTmpVis(p => ({
      ...p,
      [f.code]: e.target.checked
    })),
    style: {
      width: 18,
      height: 18,
      accentColor: color,
      cursor: "pointer"
    }
  }), /*#__PURE__*/React.createElement("span", {
    style: {
      fontFamily: "monospace",
      fontWeight: 700,
      color,
      fontSize: 14,
      minWidth: 34
    }
  }, f.code), /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 15,
      color: C.g600,
      flex: 1
    }
  }, f.label), /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 11,
      color: C.g400,
      background: C.g100,
      padding: "2px 8px",
      borderRadius: 5
    }
  }, f.unit)))))), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 10,
      marginTop: 4
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: () => {
      setVis({
        ...tmpVis
      });
      showFlash("✅ 表示設定を保存しました");
    },
    style: {
      padding: "11px 26px",
      borderRadius: 9,
      border: "none",
      cursor: "pointer",
      background: C.teal,
      color: C.white,
      fontWeight: 700,
      fontSize: 14
    }
  }, "\u4FDD\u5B58\u3059\u308B"), /*#__PURE__*/React.createElement("button", {
    onClick: () => setTmpVis(defVis()),
    style: {
      padding: "11px 16px",
      borderRadius: 9,
      border: "1.5px solid " + C.g200,
      cursor: "pointer",
      background: C.white,
      color: C.g500,
      fontSize: 13
    }
  }, "\u5168\u9078\u629E"))), id === "criteria" && criteriaTab === "lim" && /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 12,
      alignItems: "flex-start"
    }
  }, /*#__PURE__*/React.createElement(Numpad, {
    mode: "numeric",
    display: limNumDisp,
    onPress: limOnPress,
    onPrev: () => limMove(-1),
    onNext: () => limMove(1),
    canPrev: limIdx > 0,
    canNext: limIdx >= 0
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      minWidth: 0
    }
  }, /*#__PURE__*/React.createElement("p", {
    style: {
      fontSize: 13,
      color: C.g500,
      marginBottom: 14
    }
  }, "\u6700\u5C0F\u30FB\u6700\u5927\u306F\u5E38\u306B\u5165\u529B\u3067\u304D\u307E\u3059\u3002\u2705 \u3092\u30C1\u30A7\u30C3\u30AF\u3057\u305F\u9805\u76EE\u3060\u3051\u7570\u5E38\u5024\u5224\u5B9A\uFF08\u8D64\u8272\u8868\u793A\uFF09\u306B\u4F7F\u308F\u308C\u307E\u3059\u3002\u6700\u5C0F\uFF0F\u6700\u5927\u306E\u6B04\u3092\u30BF\u30C3\u30D7\u3059\u308B\u3068\u5DE6\u306E\u30C6\u30F3\u30AD\u30FC\u3067\u5165\u529B\u3067\u304D\u307E\u3059\u3002"), [{
    fields: INDOOR_FIELDS,
    label: "室内機（インドア）",
    color: C.blue
  }, {
    fields: OUTDOOR_FIELDS,
    label: "室外機（アウトドア）",
    color: C.teal
  }].map(({
    fields,
    label,
    color
  }) => /*#__PURE__*/React.createElement("div", {
    key: label,
    style: {
      marginBottom: 18
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontWeight: 700,
      fontSize: 13,
      color,
      borderBottom: "2px solid " + color,
      paddingBottom: 5,
      marginBottom: 6
    }
  }, label), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 3
    }
  }, fields.map(f => {
    const isMinActive = limActive?.code === f.code && limActive?.part === "min";
    const isMaxActive = limActive?.code === f.code && limActive?.part === "max";
    const en = !!tmpLim[f.code]?.enabled;
    return /*#__PURE__*/React.createElement("div", {
      key: f.code,
      style: {
        display: "flex",
        alignItems: "center",
        gap: 8,
        padding: "7px 10px",
        borderRadius: 8,
        border: "1.5px solid " + (en ? C.purple : C.g200),
        background: en ? C.purple + "08" : C.g50
      }
    }, /*#__PURE__*/React.createElement("input", {
      type: "checkbox",
      checked: en,
      onChange: e => setTmpLim(p => ({
        ...p,
        [f.code]: {
          ...p[f.code],
          enabled: e.target.checked
        }
      })),
      style: {
        width: 16,
        height: 16,
        accentColor: C.purple,
        cursor: "pointer",
        flexShrink: 0
      }
    }), /*#__PURE__*/React.createElement("span", {
      style: {
        fontFamily: "monospace",
        fontWeight: 700,
        color,
        fontSize: 12,
        width: 28,
        flexShrink: 0
      }
    }, f.code), /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 12,
        color: C.g600,
        flex: 1,
        minWidth: 0,
        overflow: "hidden",
        textOverflow: "ellipsis",
        whiteSpace: "nowrap"
      }
    }, f.label), /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 10,
        color: C.g400,
        width: 38,
        flexShrink: 0,
        textAlign: "center"
      }
    }, f.unit), /*#__PURE__*/React.createElement("button", {
      onClick: () => limFocus(f.code, "min"),
      style: {
        width: 64,
        padding: "5px 6px",
        borderRadius: 6,
        border: "2px solid " + (isMinActive ? C.blue : C.g200),
        background: isMinActive ? "#EFF6FF" : C.white,
        color: (isMinActive ? limNumDisp : tmpLim[f.code]?.min) ? C.g800 : C.g300,
        fontSize: 13,
        fontFamily: "monospace",
        textAlign: "right",
        cursor: "pointer",
        flexShrink: 0
      }
    }, isMinActive ? limNumDisp || "—" : tmpLim[f.code]?.min || "—"), /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 11,
        color: C.g300,
        flexShrink: 0
      }
    }, "\u301C"), /*#__PURE__*/React.createElement("button", {
      onClick: () => limFocus(f.code, "max"),
      style: {
        width: 64,
        padding: "5px 6px",
        borderRadius: 6,
        border: "2px solid " + (isMaxActive ? C.blue : C.g200),
        background: isMaxActive ? "#EFF6FF" : C.white,
        color: (isMaxActive ? limNumDisp : tmpLim[f.code]?.max) ? C.g800 : C.g300,
        fontSize: 13,
        fontFamily: "monospace",
        textAlign: "right",
        cursor: "pointer",
        flexShrink: 0
      }
    }, isMaxActive ? limNumDisp || "—" : tmpLim[f.code]?.max || "—"));
  })))), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 10,
      marginTop: 4
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: () => {
      setLimits(JSON.parse(JSON.stringify(tmpLim)));
      showFlash("✅ 正常値範囲を保存しました");
    },
    style: {
      padding: "11px 26px",
      borderRadius: 9,
      border: "none",
      cursor: "pointer",
      background: C.purple,
      color: C.white,
      fontWeight: 700,
      fontSize: 14
    }
  }, "\u4FDD\u5B58\u3059\u308B"), /*#__PURE__*/React.createElement("button", {
    onClick: () => {
      setTmpLim(defLim());
      setLimActive(null);
      setLimNumDisp("");
    },
    style: {
      padding: "11px 16px",
      borderRadius: 9,
      border: "1.5px solid " + C.g200,
      cursor: "pointer",
      background: C.white,
      color: C.g500,
      fontSize: 13
    }
  }, "\u30EA\u30BB\u30C3\u30C8"))))))))))), /*#__PURE__*/React.createElement("div", {
    id: "print-area",
    style: {
      display: "none"
    }
  }, /*#__PURE__*/React.createElement("style", null, PS), /*#__PURE__*/React.createElement("div", {
    style: {
      fontFamily: "Hiragino Sans, Meiryo, sans-serif"
    }
  }, /*#__PURE__*/React.createElement("h2", {
    style: {
      textAlign: "center",
      fontSize: 14,
      marginBottom: 10
    }
  }, "\u30A8\u30A2\u30B3\u30F3\u70B9\u691C\u30C7\u30FC\u30BF\u4E00\u89A7"), /*#__PURE__*/React.createElement("table", {
    style: {
      width: "100%",
      borderCollapse: "collapse",
      fontSize: 8
    }
  }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", {
    style: {
      background: "#1B3A6B",
      color: "white"
    }
  }, ["階", "部屋名", "管理番号", "機器番号", "室内機", "室外機", "点検日", "点検者", "運転", "モード", "風量", "設定温度", ...vf.map(f => f.code + "(" + f.unit + ")"), "備考"].map((h, i) => /*#__PURE__*/React.createElement("th", {
    key: i,
    style: {
      border: "1px solid #ccc",
      padding: "3px 4px",
      textAlign: "center",
      fontSize: 7
    }
  }, h)))), /*#__PURE__*/React.createElement("tbody", null, tRows.map((row, i) => /*#__PURE__*/React.createElement("tr", {
    key: i,
    style: {
      background: i % 2 === 0 ? "white" : "#f8fafc"
    }
  }, [row.floor, row.room, row.managementNo, row.unitNo, rowIndoorDone(row) ? "入力済" : "未入力", rowOutdoorDone(row) ? "入力済" : "未入力", rowMeta(row, "inspectionDate"), rowMeta(row, "inspector"), rowMeta(row, "preOperation"), rowMeta(row, "preMode"), rowMeta(row, "preWind"), rowMeta(row, "preSetTemp")].map((v, j) => /*#__PURE__*/React.createElement("td", {
    key: j,
    style: {
      border: "1px solid #ddd",
      padding: "3px 4px",
      textAlign: "center",
      fontSize: 7
    }
  }, v)), vf.map(f => {
    const val = rowFieldVal(row, f.code);
    const ab = val ? isAbn(f.code, val, limits) : false;
    return /*#__PURE__*/React.createElement("td", {
      key: f.code,
      style: {
        border: "1px solid #ddd",
        padding: "3px 4px",
        textAlign: "right",
        fontSize: 7,
        color: ab ? "red" : undefined,
        fontWeight: ab ? "bold" : undefined
      }
    }, val || "");
  }), /*#__PURE__*/React.createElement("td", {
    style: {
      border: "1px solid #ddd",
      padding: "3px 4px",
      fontSize: 7
    }
  }, rowRemarks(row)))))), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 8,
      fontSize: 8,
      color: "#666"
    }
  }, "\u51FA\u529B\uFF1A", new Date().toLocaleString("ja-JP")))), showRoundSwitcher && /*#__PURE__*/React.createElement("div", {
    style: {
      position: "fixed",
      inset: 0,
      background: "rgba(0,0,0,0.55)",
      zIndex: 10000,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      padding: 20
    },
    onClick: () => setShowRoundSwitcher(false)
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: C.g100,
      borderRadius: 20,
      width: "100%",
      maxWidth: 480,
      boxShadow: "0 8px 40px rgba(0,0,0,0.3)",
      overflow: "hidden",
      maxHeight: "85vh",
      display: "flex",
      flexDirection: "column"
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: "linear-gradient(135deg," + C.navy + "," + C.blue + ")",
      padding: "16px 22px",
      display: "flex",
      alignItems: "center",
      justifyContent: "space-between",
      flexShrink: 0
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: 10
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 22
    }
  }, "\uD83D\uDCC5"), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 16,
      fontWeight: 800,
      color: C.white
    }
  }, "\u70B9\u691C\u56DE\u306E\u5207\u66FF")), /*#__PURE__*/React.createElement("button", {
    onClick: () => setShowRoundSwitcher(false),
    style: {
      background: "rgba(255,255,255,0.2)",
      border: "none",
      color: C.white,
      borderRadius: 8,
      width: 32,
      height: 32,
      cursor: "pointer",
      fontSize: 16,
      fontWeight: 700,
      display: "flex",
      alignItems: "center",
      justifyContent: "center"
    }
  }, "\u2715")), /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "18px 22px",
      overflowY: "auto"
    }
  }, /*#__PURE__*/React.createElement(RoundSelector, {
    current: currentRound,
    onSelect: r => {
      setCurrentRound(r);
      setShowRoundSwitcher(false);
    },
    onCancel: () => setShowRoundSwitcher(false)
  })))), showVoiceMemo && /*#__PURE__*/React.createElement("div", {
    style: {
      position: "fixed",
      inset: 0,
      background: "rgba(0,0,0,0.55)",
      zIndex: 10000,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      padding: 20
    },
    onClick: () => setShowVoiceMemo(false)
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: C.white,
      borderRadius: 20,
      width: "100%",
      maxWidth: 480,
      boxShadow: "0 8px 40px rgba(0,0,0,0.3)",
      overflow: "hidden",
      maxHeight: "85vh",
      display: "flex",
      flexDirection: "column"
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: "linear-gradient(135deg," + C.navy + "," + C.blue + ")",
      padding: "16px 22px",
      display: "flex",
      alignItems: "center",
      justifyContent: "space-between",
      flexShrink: 0
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: 10
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 22
    }
  }, "\uD83C\uDF99\uFE0F"), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 16,
      fontWeight: 800,
      color: C.white
    }
  }, "\u30DC\u30A4\u30B9\u30E1\u30E2\uFF08\u30A2\u30D7\u30EA\u306E\u4E0D\u5177\u5408\u30FB\u6539\u5584\u6848\u306A\u3069\uFF09")), /*#__PURE__*/React.createElement("button", {
    onClick: () => setShowVoiceMemo(false),
    style: {
      background: "rgba(255,255,255,0.2)",
      border: "none",
      color: C.white,
      borderRadius: 8,
      width: 32,
      height: 32,
      cursor: "pointer",
      fontSize: 16,
      fontWeight: 700,
      display: "flex",
      alignItems: "center",
      justifyContent: "center"
    }
  }, "\u2715")), /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "18px 22px",
      overflowY: "auto"
    }
  }, /*#__PURE__*/React.createElement(VoiceMemoPanel, {
    scope: "header",
    recordId: null
  })))), showStats && /*#__PURE__*/React.createElement("div", {
    style: {
      position: "fixed",
      inset: 0,
      background: "rgba(0,0,0,0.55)",
      zIndex: 10000,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      padding: 20
    },
    onClick: () => setShowStats(false)
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: C.white,
      borderRadius: 20,
      width: "100%",
      maxWidth: 560,
      boxShadow: "0 8px 40px rgba(0,0,0,0.3)",
      overflow: "hidden",
      maxHeight: "85vh",
      overflowY: "auto",
      scrollbarGutter: "stable"
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: "linear-gradient(135deg," + C.navy + "," + C.blue + ")",
      padding: "16px 22px",
      display: "flex",
      alignItems: "center",
      justifyContent: "space-between"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: 10
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 22
    }
  }, "\uD83D\uDCCA"), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 16,
      fontWeight: 800,
      color: C.white
    }
  }, "\u96C6\u8A08\uFF08\u5EFA\u7269\u30FB\u968E\u30FB\u5BA4\u5185\u6A5F\uFF0F\u5BA4\u5916\u6A5F\uFF09")), /*#__PURE__*/React.createElement("button", {
    onClick: () => setShowStats(false),
    style: {
      background: "rgba(255,255,255,0.2)",
      border: "none",
      color: C.white,
      borderRadius: 8,
      width: 32,
      height: 32,
      cursor: "pointer",
      fontSize: 16,
      fontWeight: 700,
      display: "flex",
      alignItems: "center",
      justifyContent: "center"
    }
  }, "\u2715")), /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "16px 20px",
      display: "flex",
      flexDirection: "column",
      gap: 10
    }
  }, (() => {
    const total = tRows.length;
    const totalDone = tRows.filter(rowDone).length;
    const indoorTotal = groupStats.reduce((s, g) => s + g.indoorTotal, 0);
    const indoorDone = groupStats.reduce((s, g) => s + g.indoorDone, 0);
    const outdoorTotal = groupStats.reduce((s, g) => s + g.outdoorTotal, 0);
    const outdoorDone = groupStats.reduce((s, g) => s + g.outdoorDone, 0);
    return /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
      style: {
        background: C.g50,
        borderRadius: 12,
        padding: "12px 16px",
        display: "flex",
        flexDirection: "column",
        gap: 6,
        border: "2px solid " + C.g200
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center"
      }
    }, /*#__PURE__*/React.createElement("span", {
      style: {
        fontWeight: 700,
        fontSize: 14,
        color: C.navy
      }
    }, "\u5408\u8A08"), /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        gap: 16,
        alignItems: "center"
      }
    }, /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 13,
        color: C.g500
      }
    }, total, "\u53F0"), /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 15,
        fontWeight: 800,
        color: C.green
      }
    }, totalDone, "\u2713"), total - totalDone > 0 && /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 13,
        fontWeight: 700,
        color: C.red
      }
    }, "\u672A", total - totalDone), /*#__PURE__*/React.createElement("div", {
      style: {
        width: 80,
        height: 8,
        background: C.g200,
        borderRadius: 4,
        overflow: "hidden"
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        width: (total > 0 ? Math.round(totalDone / total * 100) : 0) + "%",
        height: "100%",
        background: C.green,
        borderRadius: 4
      }
    })), /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 12,
        color: C.g500,
        minWidth: 34,
        textAlign: "right"
      }
    }, total > 0 ? Math.round(totalDone / total * 100) : 0, "%"))), /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        gap: 16,
        fontSize: 12,
        color: C.g600,
        paddingLeft: 2
      }
    }, /*#__PURE__*/React.createElement("span", null, "\uD83C\uDFE0 \u5BA4\u5185\u6A5F\uFF1A", indoorDone, "/", indoorTotal), /*#__PURE__*/React.createElement("span", null, "\uD83C\uDFED \u5BA4\u5916\u6A5F\uFF1A", outdoorDone, "/", outdoorTotal))), groupStats.length > 0 && /*#__PURE__*/React.createElement("div", {
      style: {
        overflowX: "auto"
      }
    }, /*#__PURE__*/React.createElement("table", {
      style: {
        width: "100%",
        borderCollapse: "collapse",
        fontSize: 12
      }
    }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("th", {
      style: {
        textAlign: "left",
        padding: "6px 8px",
        fontSize: 10,
        fontWeight: 700,
        color: C.g500,
        borderBottom: "2px solid " + C.g200
      }
    }, "\u5EFA\u7269"), /*#__PURE__*/React.createElement("th", {
      style: {
        textAlign: "left",
        padding: "6px 8px",
        fontSize: 10,
        fontWeight: 700,
        color: C.g500,
        borderBottom: "2px solid " + C.g200
      }
    }, "\u968E"), /*#__PURE__*/React.createElement("th", {
      style: {
        textAlign: "center",
        padding: "6px 8px",
        fontSize: 10,
        fontWeight: 700,
        color: C.blue,
        borderBottom: "2px solid " + C.g200
      }
    }, "\uD83C\uDFE0 \u5BA4\u5185\u6A5F"), /*#__PURE__*/React.createElement("th", {
      style: {
        textAlign: "center",
        padding: "6px 8px",
        fontSize: 10,
        fontWeight: 700,
        color: C.teal,
        borderBottom: "2px solid " + C.g200
      }
    }, "\uD83C\uDFED \u5BA4\u5916\u6A5F"))), /*#__PURE__*/React.createElement("tbody", null, groupStats.map((g, i) => /*#__PURE__*/React.createElement("tr", {
      key: i,
      onClick: () => {
        setBuildingFilter(g.building === "—" ? null : g.building);
        setFloorFilter(g.floor === "—" ? null : g.floor);
        setListFilter("all");
        setView("list");
        setShowStats(false);
      },
      style: {
        cursor: "pointer",
        background: i % 2 === 0 ? C.white : C.g50
      }
    }, /*#__PURE__*/React.createElement("td", {
      style: {
        padding: "7px 8px",
        borderBottom: "1px solid " + C.g100,
        fontWeight: 700,
        color: C.navy
      }
    }, g.building), /*#__PURE__*/React.createElement("td", {
      style: {
        padding: "7px 8px",
        borderBottom: "1px solid " + C.g100,
        fontWeight: 700,
        color: C.navy
      }
    }, g.floor), /*#__PURE__*/React.createElement("td", {
      style: {
        padding: "7px 8px",
        borderBottom: "1px solid " + C.g100,
        textAlign: "center"
      }
    }, g.indoorTotal > 0 ? /*#__PURE__*/React.createElement("span", {
      style: {
        color: g.indoorDone === g.indoorTotal ? C.green : C.g600,
        fontWeight: 700
      }
    }, g.indoorDone, "/", g.indoorTotal) : /*#__PURE__*/React.createElement("span", {
      style: {
        color: C.g300
      }
    }, "\u2014")), /*#__PURE__*/React.createElement("td", {
      style: {
        padding: "7px 8px",
        borderBottom: "1px solid " + C.g100,
        textAlign: "center"
      }
    }, g.outdoorTotal > 0 ? /*#__PURE__*/React.createElement("span", {
      style: {
        color: g.outdoorDone === g.outdoorTotal ? C.green : C.g600,
        fontWeight: 700
      }
    }, g.outdoorDone, "/", g.outdoorTotal) : /*#__PURE__*/React.createElement("span", {
      style: {
        color: C.g300
      }
    }, "\u2014"))))))), groupStats.length === 0 && /*#__PURE__*/React.createElement("div", {
      style: {
        textAlign: "center",
        padding: "24px 0",
        color: C.g400,
        fontSize: 13
      }
    }, "\u30C7\u30FC\u30BF\u304C\u3042\u308A\u307E\u305B\u3093"));
  })()))), tempWarningMode !== null && /*#__PURE__*/React.createElement("div", {
    style: {
      position: "fixed",
      inset: 0,
      background: "rgba(0,0,0,0.55)",
      zIndex: 10000,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      padding: 20
    },
    onClick: confirmTempWarningNo
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: C.white,
      borderRadius: 16,
      width: "100%",
      maxWidth: 380,
      boxShadow: "0 8px 40px rgba(0,0,0,0.3)",
      overflow: "hidden"
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: "linear-gradient(135deg,#F59E0B,#D97706)",
      padding: "14px 18px"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 14,
      fontWeight: 800,
      color: C.white
    }
  }, "\u26A0\uFE0F \u6E29\u5EA6\u5DEE\u304C\u5C0F\u3055\u3044\u30C7\u30FC\u30BF\u3067\u3059")), /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "16px 18px"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 13,
      color: C.g700,
      marginBottom: 14,
      lineHeight: 1.6
    }
  }, "\u5438\u8FBC\u6E29\u5EA6\u3068\u5439\u51FA\u6E29\u5EA6\u306E\u5DEE\u304C5\u2103\u672A\u6E80\u3067\u3059\u3002", /*#__PURE__*/React.createElement("br", null), "\u3053\u306E\u30C7\u30FC\u30BF\u3067\u3088\u308D\u3057\u3044\u3067\u3059\u304B\uFF1F"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 8
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: confirmTempWarningNo,
    style: {
      flex: 1,
      padding: "10px",
      borderRadius: 9,
      border: "1.5px solid " + C.g300,
      background: C.g50,
      color: C.g600,
      fontWeight: 700,
      fontSize: 13,
      cursor: "pointer"
    }
  }, "\u3044\u3044\u3048"), /*#__PURE__*/React.createElement("button", {
    onClick: confirmTempWarningYes,
    style: {
      flex: 1,
      padding: "10px",
      borderRadius: 9,
      border: "none",
      background: "#D97706",
      color: C.white,
      fontWeight: 700,
      fontSize: 13,
      cursor: "pointer"
    }
  }, "\u306F\u3044"))))), showExportConfirm && /*#__PURE__*/React.createElement("div", {
    style: {
      position: "fixed",
      inset: 0,
      background: "rgba(0,0,0,0.55)",
      zIndex: 10000,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      padding: 20
    },
    onClick: () => setShowExportConfirm(null)
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: C.white,
      borderRadius: 16,
      width: "100%",
      maxWidth: 380,
      boxShadow: "0 8px 40px rgba(0,0,0,0.3)",
      overflow: "hidden"
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: "linear-gradient(135deg,#F59E0B,#D97706)",
      padding: "14px 18px"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 14,
      fontWeight: 800,
      color: C.white
    }
  }, "\u26A0\uFE0F \u7D5E\u308A\u8FBC\u307F\u304C\u9078\u629E\u3055\u308C\u3066\u3044\u307E\u3059")), /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "16px 18px"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 13,
      color: C.g700,
      marginBottom: 10,
      lineHeight: 1.6
    }
  }, "\u30C7\u30FC\u30BF\u4E00\u89A7\u3067\u4EE5\u4E0B\u306E\u7D5E\u308A\u8FBC\u307F\u30FB\u4E26\u3073\u66FF\u3048\u304C\u9078\u629E\u3055\u308C\u305F\u72B6\u614B\u3067\u3059\u3002", /*#__PURE__*/React.createElement("br", null), "\u3053\u306E\u5185\u5BB9\u306E\u307E\u307E", showExportConfirm.label, "\u51FA\u529B\u3057\u3066\u3088\u308D\u3057\u3044\u3067\u3059\u304B\uFF1F"), /*#__PURE__*/React.createElement("div", {
    style: {
      background: C.g50,
      borderRadius: 10,
      padding: "10px 12px",
      marginBottom: 14,
      display: "flex",
      flexDirection: "column",
      gap: 5
    }
  }, showExportConfirm.list.map((desc, i) => /*#__PURE__*/React.createElement("div", {
    key: i,
    style: {
      fontSize: 12,
      fontWeight: 700,
      color: C.navy
    }
  }, "\u30FB", desc))), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 8
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: () => setShowExportConfirm(null),
    style: {
      flex: 1,
      padding: "10px",
      borderRadius: 9,
      border: "1.5px solid " + C.g300,
      background: C.g50,
      color: C.g600,
      fontWeight: 700,
      fontSize: 13,
      cursor: "pointer"
    }
  }, "\u3044\u3044\u3048"), /*#__PURE__*/React.createElement("button", {
    onClick: () => {
      showExportConfirm.run();
      setShowExportConfirm(null);
    },
    style: {
      flex: 1,
      padding: "10px",
      borderRadius: 9,
      border: "none",
      background: C.teal,
      color: C.white,
      fontWeight: 700,
      fontSize: 13,
      cursor: "pointer"
    }
  }, "\u306F\u3044\u30FB\u51FA\u529B\u3059\u308B"))))), flash && /*#__PURE__*/React.createElement("div", {
    style: {
      position: "fixed",
      bottom: 24,
      right: 24,
      background: C.green,
      color: C.white,
      padding: "12px 24px",
      borderRadius: 10,
      fontWeight: 700,
      fontSize: 15,
      boxShadow: "0 4px 16px rgba(0,0,0,0.2)",
      zIndex: 9999
    }
  }, flash), saveModal && /*#__PURE__*/React.createElement("div", {
    style: {
      position: "fixed",
      inset: 0,
      background: "rgba(0,0,0,0.55)",
      zIndex: 10000,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      padding: 20
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: C.white,
      borderRadius: 20,
      width: "100%",
      maxWidth: 560,
      boxShadow: "0 8px 40px rgba(0,0,0,0.3)",
      overflow: "hidden",
      maxHeight: "90vh",
      overflowY: "auto",
      scrollbarGutter: "stable"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: "linear-gradient(135deg," + C.green + ",#047857)",
      padding: "14px 18px",
      display: "flex",
      alignItems: "center",
      justifyContent: "space-between"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: 10
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 24
    }
  }, "\u2705"), /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 16,
      fontWeight: 800,
      color: C.white
    }
  }, "\u4FDD\u5B58\u3057\u307E\u3057\u305F"), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 11,
      color: "rgba(255,255,255,0.8)",
      marginTop: 1
    }
  }, saveModal.floor, "\u3000", saveModal.room, "\u3000", saveModal.managementNo, " / ", saveModal.unitNo))), /*#__PURE__*/React.createElement("button", {
    onClick: closeNext,
    style: {
      background: "rgba(255,255,255,0.2)",
      border: "none",
      color: C.white,
      borderRadius: 8,
      width: 34,
      height: 34,
      cursor: "pointer",
      fontSize: 18,
      fontWeight: 700,
      display: "flex",
      alignItems: "center",
      justifyContent: "center"
    }
  }, "\u2715")), /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "14px 18px",
      display: "flex",
      flexDirection: "column",
      gap: 12
    }
  }, /*#__PURE__*/React.createElement("div", {
    onClick: () => setMeasZoom(true),
    style: {
      background: C.g50,
      borderRadius: 10,
      padding: "10px 14px",
      cursor: "pointer",
      border: "2px solid " + C.g200,
      transition: "border-color 0.15s",
      userSelect: "none"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      justifyContent: "space-between",
      alignItems: "center",
      marginBottom: 8
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 11,
      fontWeight: 700,
      color: C.g500
    }
  }, "\uD83D\uDCCA \u6E2C\u5B9A\u30C7\u30FC\u30BF"), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 10,
      color: C.blue,
      fontWeight: 700
    }
  }, "\u30BF\u30C3\u30D7\u3067\u62E1\u5927 \u25B6")), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gridTemplateColumns: "repeat(auto-fill,minmax(170px,1fr))",
      gap: "4px 12px"
    }
  }, ALL_FIELDS.filter(f => vis[f.code] && saveModal.values[f.code] !== "").map(f => {
    const abn = isAbn(f.code, saveModal.values[f.code], limits);
    return /*#__PURE__*/React.createElement("div", {
      key: f.code,
      style: {
        display: "flex",
        alignItems: "baseline",
        gap: 5,
        padding: "3px 0",
        borderBottom: "1px solid " + C.g200
      }
    }, /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 11,
        fontFamily: "monospace",
        fontWeight: 700,
        color: f.group === "indoor" ? C.blue : C.teal,
        minWidth: 24
      }
    }, f.code), /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 10,
        color: C.g500,
        flex: 1,
        whiteSpace: "nowrap",
        overflow: "hidden",
        textOverflow: "ellipsis"
      }
    }, f.label), /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 13,
        fontWeight: 700,
        fontFamily: "monospace",
        color: abn ? C.red : C.navy,
        whiteSpace: "nowrap"
      }
    }, saveModal.values[f.code], /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 9,
        color: C.g400
      }
    }, " ", f.unit), abn && /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 9,
        color: C.red
      }
    }, " \u26A0\uFE0F")));
  })), saveModal.remarks && /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 8,
      fontSize: 12,
      color: C.g600,
      borderTop: "1px solid " + C.g200,
      paddingTop: 6
    }
  }, "\u5099\u8003: ", saveModal.remarks)), (() => {
    const chkGroup = inspectionMode === "outdoor" ? "check_out" : "check_in";
    const chkList = checkFields.filter(f => f.group === chkGroup);
    if (chkList.length === 0) return null;
    return /*#__PURE__*/React.createElement("div", {
      style: {
        background: C.g50,
        borderRadius: 10,
        padding: "10px 14px",
        border: "2px solid " + C.g200
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        fontSize: 11,
        fontWeight: 700,
        color: C.g500,
        marginBottom: 8
      }
    }, "\u2705 \u70B9\u691C\u30C1\u30A7\u30C3\u30AF\u9805\u76EE"), /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        flexDirection: "column",
        gap: 4
      }
    }, chkList.map(f => {
      const v = saveModal.checks?.[f.code] || "";
      return /*#__PURE__*/React.createElement("div", {
        key: f.code,
        style: {
          display: "flex",
          alignItems: "center",
          gap: 8,
          padding: "3px 0",
          borderBottom: "1px solid " + C.g200
        }
      }, /*#__PURE__*/React.createElement("span", {
        style: {
          fontSize: 10,
          fontWeight: 700,
          color: C.teal,
          minWidth: 70,
          whiteSpace: "nowrap",
          overflow: "hidden",
          textOverflow: "ellipsis"
        }
      }, f.category), /*#__PURE__*/React.createElement("span", {
        style: {
          fontSize: 12,
          color: C.g700,
          flex: 1
        }
      }, f.label), /*#__PURE__*/React.createElement("span", {
        style: {
          fontSize: 16,
          fontWeight: 800,
          color: v === "○" ? C.green : v === "×" ? C.red : C.g300,
          minWidth: 20,
          textAlign: "center"
        }
      }, v || "—"));
    })));
  })(), saveModal.preOperation || saveModal.preMode || saveModal.preWind || saveModal.preSetTemp ? /*#__PURE__*/React.createElement("div", {
    style: {
      background: "#FFF7ED",
      border: "2px solid #F59E0B",
      borderRadius: 12,
      padding: "14px 16px"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 13,
      fontWeight: 800,
      color: "#92400E",
      marginBottom: 10
    }
  }, "\u26A0\uFE0F \u30EA\u30E2\u30B3\u30F3\u3092\u70B9\u691C\u524D\u306E\u72B6\u614B\u306B\u623B\u3057\u3066\u304F\u3060\u3055\u3044"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 8,
      marginBottom: 12,
      flexWrap: "wrap"
    }
  }, [{
    show: !!saveModal.preOperation,
    label: "運転",
    val: saveModal.preOperation === "ON" ? "🟢 ON" : "⭕ OFF"
  }, {
    show: !!saveModal.preMode,
    label: "モード",
    val: modeLabel(saveModal.preMode)
  }, {
    show: !!saveModal.preWind,
    label: "風量",
    val: windLabel(saveModal.preWind)
  }, {
    show: !!saveModal.preSetTemp,
    label: "設定温度",
    val: saveModal.preSetTemp + "°C"
  }].filter(p => p.show).map(p => /*#__PURE__*/React.createElement("div", {
    key: p.label,
    style: {
      flex: 1,
      minWidth: 70,
      background: C.white,
      border: "2px solid #F59E0B",
      borderRadius: 10,
      padding: "8px 10px",
      textAlign: "center"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 10,
      fontWeight: 700,
      color: "#92400E",
      marginBottom: 4
    }
  }, p.label), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 15,
      fontWeight: 800,
      color: "#78350F",
      fontFamily: p.label === "設定温度" ? "monospace" : "inherit"
    }
  }, p.val)))), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 8
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: closeSaveToList,
    style: {
      padding: "16px 14px",
      borderRadius: 10,
      border: "1.5px solid " + C.g300,
      cursor: "pointer",
      background: C.white,
      color: C.g600,
      fontWeight: 700,
      fontSize: 14,
      whiteSpace: "nowrap"
    }
  }, "\uD83D\uDCCB \u30C7\u30FC\u30BF\u4E00\u89A7"), /*#__PURE__*/React.createElement("button", {
    onClick: closeNext,
    style: {
      flex: 1,
      padding: "16px",
      borderRadius: 10,
      border: "none",
      cursor: "pointer",
      background: "linear-gradient(135deg," + C.green + ",#047857)",
      color: C.white,
      fontWeight: 800,
      fontSize: 16,
      boxShadow: "0 3px 10px rgba(5,150,105,0.3)"
    }
  }, "\u2705 \u623B\u3057\u307E\u3057\u305F\u3000\u2192\u3000", nextLabel))) : /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 8
    }
  }, /*#__PURE__*/React.createElement("button", {
    onClick: closeSaveToList,
    style: {
      padding: "16px 14px",
      borderRadius: 10,
      border: "1.5px solid " + C.g300,
      cursor: "pointer",
      background: C.white,
      color: C.g600,
      fontWeight: 700,
      fontSize: 14,
      whiteSpace: "nowrap"
    }
  }, "\uD83D\uDCCB \u30C7\u30FC\u30BF\u4E00\u89A7"), /*#__PURE__*/React.createElement("button", {
    onClick: closeNext,
    style: {
      flex: 1,
      padding: "16px",
      borderRadius: 10,
      border: "none",
      cursor: "pointer",
      background: "linear-gradient(135deg," + C.navy + "," + C.blue + ")",
      color: C.white,
      fontWeight: 800,
      fontSize: 16,
      boxShadow: "0 3px 10px rgba(37,99,176,0.3)"
    }
  }, "\u2705 ", nextLabel))))), measZoom && saveModal && /*#__PURE__*/React.createElement("div", {
    style: {
      position: "fixed",
      inset: 0,
      background: "rgba(0,0,0,0.75)",
      zIndex: 10001,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      padding: 16
    },
    onClick: () => setMeasZoom(false)
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: C.white,
      borderRadius: 20,
      width: "100%",
      maxWidth: 600,
      boxShadow: "0 8px 40px rgba(0,0,0,0.4)",
      overflow: "hidden",
      maxHeight: "92vh",
      display: "flex",
      flexDirection: "column"
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: "linear-gradient(135deg," + C.navy + "," + C.blue + ")",
      padding: "12px 18px",
      display: "flex",
      alignItems: "center",
      justifyContent: "space-between",
      flexShrink: 0
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 15,
      fontWeight: 800,
      color: C.white
    }
  }, "\uD83D\uDCCA \u6E2C\u5B9A\u30C7\u30FC\u30BF\u3000", saveModal.floor, " ", saveModal.room), /*#__PURE__*/React.createElement("button", {
    onClick: () => setMeasZoom(false),
    style: {
      background: "rgba(255,255,255,0.2)",
      border: "none",
      color: C.white,
      borderRadius: 8,
      width: 34,
      height: 34,
      cursor: "pointer",
      fontSize: 18,
      fontWeight: 700,
      display: "flex",
      alignItems: "center",
      justifyContent: "center"
    }
  }, "\u2715")), /*#__PURE__*/React.createElement("div", {
    style: {
      overflowY: "auto",
      scrollbarGutter: "stable",
      padding: "12px 16px",
      display: "flex",
      flexDirection: "column",
      gap: 3
    }
  }, ALL_FIELDS.filter(f => vis[f.code] && saveModal.values[f.code] !== "").map(f => {
    const abn = isAbn(f.code, saveModal.values[f.code], limits);
    return /*#__PURE__*/React.createElement("div", {
      key: f.code,
      style: {
        display: "flex",
        alignItems: "center",
        gap: 10,
        padding: "10px 14px",
        borderRadius: 10,
        background: abn ? "#FEF2F2" : f.group === "indoor" ? C.blue + "08" : C.teal + "08",
        border: "1.5px solid " + (abn ? C.red : f.group === "indoor" ? C.blue + "30" : C.teal + "30")
      }
    }, /*#__PURE__*/React.createElement("span", {
      style: {
        fontFamily: "monospace",
        fontWeight: 800,
        fontSize: 18,
        color: f.group === "indoor" ? C.blue : C.teal,
        minWidth: 34
      }
    }, f.code), /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 15,
        color: C.g600,
        flex: 1
      }
    }, f.label), /*#__PURE__*/React.createElement("span", {
      style: {
        fontFamily: "monospace",
        fontWeight: 800,
        fontSize: 26,
        color: abn ? C.red : C.navy
      }
    }, saveModal.values[f.code]), /*#__PURE__*/React.createElement("span", {
      style: {
        fontSize: 13,
        color: C.g400,
        minWidth: 38
      }
    }, f.unit, abn && " ⚠️"));
  }), saveModal.remarks && /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 6,
      padding: "10px 14px",
      borderRadius: 10,
      background: C.g50,
      fontSize: 13,
      color: C.g600
    }
  }, "\u5099\u8003: ", saveModal.remarks)))));
}

// ─── 実機表示用ルート（iPad / Edge 等、実際のブラウザ画面いっぱいに表示） ─────────────────────────────────
function AppRoot() {
  return /*#__PURE__*/React.createElement("div", {
    style: {
      width: "100%",
      height: "100vh",
      background: "#fff",
      overflow: "hidden",
      fontFamily: "'Hiragino Sans','Hiragino Kaku Gothic ProN',Meiryo,sans-serif"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      width: "100%",
      height: "100%",
      display: "flex",
      flexDirection: "column"
    }
  }, /*#__PURE__*/React.createElement(ACInspectionApp, null)));
}
const rootEl = document.getElementById("root");
if (window.ReactDOM.createRoot) {
  window.ReactDOM.createRoot(rootEl).render(/*#__PURE__*/React.createElement(AppRoot, null));
} else {
  window.ReactDOM.render(/*#__PURE__*/React.createElement(AppRoot, null), rootEl);
}