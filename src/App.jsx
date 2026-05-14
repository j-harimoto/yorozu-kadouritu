import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

// ── 設定 ─────────────────────────────────────────────────────────────────────
const GAS_URL = "https://script.google.com/macros/s/AKfycbxuSYkH5QJGbrkMMdfifwn3aw_p9ncDDtpu8CDKQy40hoP1HXR9hQn0xmbL4QU9CMGV/exec";

const ACCENT       = "#D4A017";
const ACCENT_DARK  = "#B8860B";
const ACCENT_PALE  = "#FFFDE7";
const ACCENT_LIGHT = "#FFF8E1";

const NON_EMAIL  = ["来訪（対面）", "オンライン", "訪問（対面）", "電話"];
const VISIT_TYPE = "訪問（対面）";

// ── GAS API ──────────────────────────────────────────────────────────────────
async function gasGet(params) {
  const url = new URL(GAS_URL);
  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
  const res = await fetch(url.toString());
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
}

async function gasPost(body) {
  const res = await fetch(GAS_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain" },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
}

// ── データ解析 ────────────────────────────────────────────────────────────────
function parseRows(rawRows) {
  return rawRows
    .filter(r => r["対応者1"] && r["支援日"] && r["支援形態"])
    .map(r => {
      const dateRaw = r["支援日"];
      let dateStr = "";
      if (dateRaw instanceof Date) {
        dateStr = `${dateRaw.getMonth() + 1}/${dateRaw.getDate()}/${dateRaw.getFullYear()}`;
      } else {
        dateStr = String(dateRaw).trim();
      }
      return {
        staff:   String(r["対応者1"]).trim(),
        date:    dateStr,
        type:    String(r["支援形態"]).trim(),
        phase:   parseInt(r["フェーズ２"] ?? r["フェーズ2"] ?? 0) || 0,
        consult: parseInt(r["相談時間（分単位）"] ?? r["相談時間"] ?? 0) || 0,
        move:    parseInt(r["移動時間（数値）"]   ?? r["移動時間"] ?? 0) || 0,
      };
    })
    .filter(r => r.staff.length > 0);
}

function calcStats(rows, staff) {
  const data = staff ? rows.filter(d => d.staff === staff) : rows;
  const totalCount = data.length;
  const uniqueDays = new Set(data.map(d => d.date)).size;
  const rate = uniqueDays > 0 ? totalCount / uniqueDays : 0;
  let wsP1C = 0, wsP1M = 0, wsP2C = 0, wsP2M = 0, psC = 0, psM = 0;
  data.filter(d => NON_EMAIL.includes(d.type)).forEach(r => {
    if (r.type === VISIT_TYPE)   { psC  += r.consult; psM  += r.move; }
    else if (r.phase === 0)      { wsP1C += r.consult; wsP1M += r.move; }
    else                         { wsP2C += r.consult; wsP2M += r.move; }
  });
  const wsP1T = wsP1C + wsP1M;
  const wsP2T = wsP2C * 2 + wsP2M;
  return {
    totalCount, uniqueDays, rate,
    wsP1Consult: wsP1C, wsP1Move: wsP1M, wsP1Total: wsP1T,
    wsP2Consult: wsP2C * 2, wsP2Move: wsP2M, wsP2Total: wsP2T,
    wsGrand: wsP1T + wsP2T,
    psConsult: psC, psMove: psM,
    psGrand: Math.round((psC + psM) / 2),
  };
}

const fmt = n => Number(n).toLocaleString("ja-JP");

// ── メインコンポーネント ──────────────────────────────────────────────────────
export default function Dashboard() {
  const [tab, setTab]                     = useState("dashboard");
  const [rows, setRows]                   = useState([]);
  const [staffFilter, setStaffFilter]     = useState("");
  const [uploadFile, setUploadFile]       = useState(null);
  const [uploadYear, setUploadYear]       = useState(new Date().getFullYear());
  const [uploadMonth, setUploadMonth]     = useState(new Date().getMonth() + 1);
  const [statusType, setStatusType]       = useState("");   // "" | "loading" | "success" | "error"
  const [statusMsg, setStatusMsg]         = useState("");
  const [uploading, setUploading]         = useState(false);
  const [folders, setFolders]             = useState([]);
  const [selectedMonth, setSelectedMonth] = useState("");
  const [loadingFolders, setLoadingFolders] = useState(false);
  const [loadingData, setLoadingData]     = useState(false);
  const fileRef = useRef(null);

  useEffect(() => { fetchFolders(); }, []);

  // フォルダ一覧取得
  async function fetchFolders() {
    setLoadingFolders(true);
    try {
      const data = await gasGet({ action: "listFolders" });
      setFolders(data.folders || []);
    } catch (e) {
      console.error("フォルダ取得エラー:", e);
    }
    setLoadingFolders(false);
  }

  // 月次データ読み込み
  async function loadMonthData(folderId, folderName) {
    setSelectedMonth(folderName);
    setLoadingData(true);
    setRows([]);
    setStaffFilter("");
    try {
      const { files } = await gasGet({ action: "listFiles", folderId });
      const target = (files || []).find(f => f.name.endsWith(".xlsx") || f.name.endsWith(".csv"));
      if (!target) {
        alert("このフォルダに xlsx / csv ファイルが見つかりません");
        setLoadingData(false);
        return;
      }
      const { base64 } = await gasGet({ action: "getFile", fileId: target.id });
      const bytes = Uint8Array.from(atob(base64), c => c.charCodeAt(0));
      const wb = XLSX.read(bytes, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      setRows(parseRows(XLSX.utils.sheet_to_json(ws, { defval: "" })));
      setTab("dashboard");
    } catch (e) {
      alert("データ読み込みエラー: " + e.message);
    }
    setLoadingData(false);
  }

  // アップロード
  async function handleUpload() {
    if (!uploadFile) return;
    setUploading(true);
    setStatusType("loading");
    setStatusMsg("Google Drive にアップロード中...");
    try {
      const base64Data = await new Promise((res, rej) => {
        const reader = new FileReader();
        reader.onload  = () => res(reader.result.split(",")[1]);
        reader.onerror = rej;
        reader.readAsDataURL(uploadFile);
      });

      const folderName = `${uploadYear}-${String(uploadMonth).padStart(2, "0")}`;
      const result = await gasPost({
        action: "uploadFile",
        folderName,
        fileName: uploadFile.name,
        base64Data,
      });

      if (result.error) throw new Error(result.error);

      setStatusType("success");
      setStatusMsg(`${folderName} にアップロード完了しました`);
      await fetchFolders();

      // アップロードしたデータをそのまま表示
      const bytes = Uint8Array.from(atob(base64Data), c => c.charCodeAt(0));
      const wb = XLSX.read(bytes, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      setRows(parseRows(XLSX.utils.sheet_to_json(ws, { defval: "" })));
      setSelectedMonth(folderName);
      setStaffFilter("");
      setUploadFile(null);
      if (fileRef.current) fileRef.current.value = "";
    } catch (e) {
      setStatusType("error");
      setStatusMsg(e.message);
    }
    setUploading(false);
  }

  // ローカルプレビュー
  function handleLocalPreview(e) {
    const f = e.target.files[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const bytes = new Uint8Array(ev.target.result);
      const wb = XLSX.read(bytes, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      setRows(parseRows(XLSX.utils.sheet_to_json(ws, { defval: "" })));
      setSelectedMonth("ローカルプレビュー（未保存）");
      setStaffFilter("");
      setTab("dashboard");
    };
    reader.readAsArrayBuffer(f);
  }

  // 集計
  const staffList = [...new Set(rows.map(d => d.staff))].sort();
  const summary   = calcStats(rows, staffFilter);
  const avgRate   = staffList.length > 0
    ? staffList.reduce((s, p) => s + calcStats(rows, p).rate, 0) / staffList.length
    : 0;
  const maxRate = Math.max(...staffList.map(s => calcStats(rows, s).rate), 0.01);
  const currentYear = new Date().getFullYear();
  const years  = Array.from({ length: 5 }, (_, i) => currentYear - 2 + i);
  const months = Array.from({ length: 12 }, (_, i) => i + 1);

  return (
    <div style={{ fontFamily: "'Noto Sans JP', sans-serif", minHeight: "100vh", background: "#FAFAF5", color: "#1A1A0F" }}>

      {/* ヘッダー */}
      <div style={{ background: "#1A1A0F", borderBottom: `3px solid ${ACCENT}`, padding: "0 24px", position: "sticky", top: 0, zIndex: 100 }}>
        <div style={{ maxWidth: 1100, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", height: 54 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 6, height: 28, background: ACCENT, borderRadius: 2 }} />
            <span style={{ color: "#FFF", fontWeight: 700, fontSize: 14, letterSpacing: ".04em" }}>
              よろず支援拠点 管理ダッシュボード
            </span>
            {selectedMonth && (
              <span style={{ fontSize: 11, padding: "2px 10px", background: "#2A2A1A", color: ACCENT, borderRadius: 20, fontWeight: 600 }}>
                {selectedMonth}
              </span>
            )}
          </div>
          <div style={{ display: "flex", gap: 4 }}>
            {[["dashboard","ダッシュボード"],["upload","データ管理"]].map(([key, label]) => (
              <button key={key} onClick={() => setTab(key)} style={{
                padding: "6px 16px", fontSize: 12, fontWeight: 600,
                border: "none", cursor: "pointer", borderRadius: 4,
                background: tab === key ? ACCENT : "transparent",
                color:      tab === key ? "#1A1A0F" : "#999",
                transition: "all .15s",
              }}>{label}</button>
            ))}
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "20px 24px" }}>

        {/* ══ データ管理タブ ══ */}
        {tab === "upload" && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, alignItems: "start" }}>

            {/* アップロードパネル */}
            <div style={{ background: "#FFF", border: `1.5px solid ${ACCENT}`, borderRadius: 12, padding: 24 }}>
              <SectionTitle>月次データアップロード</SectionTitle>

              <Field label="対象年度">
                <select value={uploadYear} onChange={e => setUploadYear(+e.target.value)} style={selStyle}>
                  {years.map(y => <option key={y} value={y}>{y}年</option>)}
                </select>
              </Field>

              <Field label="対象月">
                <select value={uploadMonth} onChange={e => setUploadMonth(+e.target.value)} style={selStyle}>
                  {months.map(m => <option key={m} value={m}>{m}月</option>)}
                </select>
              </Field>

              <Field label="ファイル選択（.xlsx）">
                <div onClick={() => fileRef.current?.click()} style={{
                  border: `2px dashed ${uploadFile ? ACCENT : "#CCC"}`,
                  borderRadius: 8, padding: "22px 16px", textAlign: "center",
                  cursor: "pointer", background: uploadFile ? ACCENT_PALE : "#FAFAFA",
                  transition: "all .2s",
                }}>
                  <input ref={fileRef} type="file" accept=".xlsx,.csv" style={{ display: "none" }}
                    onChange={e => { setUploadFile(e.target.files[0] || null); setStatusType(""); }} />
                  {uploadFile
                    ? <><div style={{ fontSize: 22, marginBottom: 4 }}>📊</div>
                        <div style={{ fontSize: 13, fontWeight: 700, color: ACCENT_DARK }}>{uploadFile.name}</div>
                        <div style={{ fontSize: 11, color: "#AAA", marginTop: 2 }}>クリックで変更</div></>
                    : <><div style={{ fontSize: 28, color: "#CCC", marginBottom: 4 }}>⬆</div>
                        <div style={{ fontSize: 12, color: "#999" }}>クリックしてファイルを選択</div></>
                  }
                </div>
              </Field>

              <button onClick={handleUpload} disabled={!uploadFile || uploading} style={{
                width: "100%", padding: "12px", fontSize: 14, fontWeight: 700,
                background: uploadFile && !uploading ? ACCENT : "#DDD",
                color:      uploadFile && !uploading ? "#1A1A0F" : "#AAA",
                border: "none", borderRadius: 8,
                cursor: uploadFile && !uploading ? "pointer" : "not-allowed",
                letterSpacing: ".03em", transition: "all .15s", marginBottom: 10,
              }}>
                {uploading ? "⏳ アップロード中..." : `↑ ${uploadYear}年${uploadMonth}月 を Google Drive に保存`}
              </button>

              {statusType && (
                <div style={{
                  padding: "10px 14px", borderRadius: 6, fontSize: 12, fontWeight: 600,
                  background: statusType==="success"?"#E8F5E9": statusType==="error"?"#FFEBEE":ACCENT_PALE,
                  color:      statusType==="success"?"#2E7D32": statusType==="error"?"#C62828":ACCENT_DARK,
                }}>
                  {statusType==="success"?"✓ ": statusType==="error"?"✗ ":"⏳ "}
                  {statusMsg}
                </div>
              )}

              <div style={{ marginTop: 16, borderTop: "1px solid #EEE", paddingTop: 14 }}>
                <div style={{ fontSize: 11, color: "#AAA", marginBottom: 8 }}>ローカルプレビュー（Driveに保存せず確認のみ）</div>
                <label style={{
                  display: "inline-flex", alignItems: "center", gap: 6,
                  padding: "7px 14px", fontSize: 12, fontWeight: 600,
                  background: "#F5F5F5", border: "1px solid #DDD", borderRadius: 6, cursor: "pointer",
                }}>
                  📂 ファイルを開く
                  <input type="file" accept=".xlsx,.csv" style={{ display: "none" }} onChange={handleLocalPreview} />
                </label>
              </div>
            </div>

            {/* 保存済みデータ一覧 */}
            <div style={{ background: "#FFF", border: "1px solid #E8E0C8", borderRadius: 12, padding: 24 }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 18 }}>
                <SectionTitle style={{ marginBottom: 0 }}>保存済みデータ</SectionTitle>
                <button onClick={fetchFolders} style={{
                  fontSize: 11, padding: "5px 12px", background: "#F5F5F5",
                  border: "1px solid #DDD", borderRadius: 4, cursor: "pointer", fontWeight: 600,
                }}>
                  {loadingFolders ? "読込中..." : "↻ 更新"}
                </button>
              </div>

              {loadingFolders ? (
                <div style={{ textAlign: "center", padding: "40px 0", color: "#AAA", fontSize: 13 }}>読み込み中...</div>
              ) : folders.length === 0 ? (
                <div style={{ textAlign: "center", padding: "40px 0", color: "#AAA", fontSize: 13 }}>
                  <div style={{ fontSize: 36, marginBottom: 8 }}>📭</div>
                  データがまだありません
                </div>
              ) : folders.map(f => (
                <div key={f.id} onClick={() => loadMonthData(f.id, f.name)}
                  style={{
                    display: "flex", alignItems: "center", justifyContent: "space-between",
                    padding: "11px 14px", marginBottom: 6, borderRadius: 8, cursor: "pointer",
                    background: selectedMonth === f.name ? ACCENT_PALE : "#FAFAFA",
                    border: `1px solid ${selectedMonth === f.name ? ACCENT : "#EEE"}`,
                    transition: "all .15s",
                  }}
                  onMouseEnter={e => { if (selectedMonth !== f.name) e.currentTarget.style.background = ACCENT_LIGHT; }}
                  onMouseLeave={e => { e.currentTarget.style.background = selectedMonth === f.name ? ACCENT_PALE : "#FAFAFA"; }}
                >
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <span style={{ fontSize: 18 }}>📁</span>
                    <span style={{ fontSize: 13, fontWeight: selectedMonth === f.name ? 700 : 400 }}>{f.name}</span>
                  </div>
                  <span style={{ fontSize: 11, color: ACCENT_DARK, fontWeight: 600 }}>
                    {loadingData && selectedMonth === f.name ? "読込中..." : "表示 →"}
                  </span>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ══ ダッシュボードタブ ══ */}
        {tab === "dashboard" && (
          <>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 20, flexWrap: "wrap", gap: 12 }}>
              <div style={{ fontSize: 12, color: "#888" }}>
                {rows.length > 0
                  ? <span>全 <strong style={{ color: "#1A1A0F" }}>{rows.length}</strong> 件 ／ 担当者 <strong style={{ color: "#1A1A0F" }}>{staffList.length}</strong> 名</span>
                  : "「データ管理」タブからファイルをアップロード、または保存済みデータを選択してください"
                }
              </div>
              {rows.length > 0 && (
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <label style={{ fontSize: 12, color: "#666" }}>担当者：</label>
                  <select value={staffFilter} onChange={e => setStaffFilter(e.target.value)} style={selStyle}>
                    <option value="">全員</option>
                    {staffList.map(s => <option key={s} value={s}>{s}</option>)}
                  </select>
                </div>
              )}
            </div>

            {rows.length === 0 ? (
              <div style={{ textAlign: "center", padding: "80px 0", color: "#BBB" }}>
                <div style={{ fontSize: 56, marginBottom: 14 }}>📊</div>
                <div style={{ fontSize: 15, fontWeight: 700, marginBottom: 6, color: "#888" }}>データが読み込まれていません</div>
                <div style={{ fontSize: 12 }}>「データ管理」タブからファイルを選択してください</div>
                <button onClick={() => setTab("upload")} style={{
                  marginTop: 16, padding: "10px 24px", fontSize: 13, fontWeight: 700,
                  background: ACCENT, color: "#1A1A0F", border: "none", borderRadius: 8, cursor: "pointer",
                }}>データ管理へ →</button>
              </div>
            ) : (
              <>
                {/* KPIカード */}
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(175px, 1fr))", gap: 12, marginBottom: 18 }}>
                  {!staffFilter && <KpiCard dark label="全体平均稼働率" value={avgRate.toFixed(2)} sub={`${staffList.length}名の平均`} />}
                  {staffFilter  && <KpiCard dark label={`${staffFilter} 稼働率`} value={summary.rate.toFixed(2)} sub={`出勤 ${summary.uniqueDays}日`} />}
                  <KpiCard label="総件数"        value={fmt(summary.totalCount)} sub="メール含む" />
                  <KpiCard label="WS 需要量"      value={fmt(summary.wsGrand)}   sub="分（ワンストップ）" />
                  <KpiCard label="生産性C 需要量" value={fmt(summary.psGrand)}   sub="分（生産性センター）" />
                </div>

                {/* 需要量詳細 */}
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 18 }}>
                  <div style={{ background: "#FFF", border: `1.5px solid ${ACCENT}`, borderRadius: 12, overflow: "hidden" }}>
                    <div style={{ background: ACCENT, padding: "10px 18px" }}>
                      <span style={{ fontWeight: 700, fontSize: 13, color: "#1A1A0F" }}>ワンストップ 需要量</span>
                    </div>
                    <div style={{ padding: "14px 18px" }}>
                      <DRow label="フェーズ1 相談時間" value={fmt(summary.wsP1Consult)} />
                      <DRow label="フェーズ1 移動時間" value={fmt(summary.wsP1Move)} />
                      <SRow label="フェーズ1合計" value={fmt(summary.wsP1Total)} bg="#FFF8E1" color={ACCENT_DARK} />
                      <DRow label="フェーズ2 相談×2"  value={fmt(summary.wsP2Consult)} />
                      <DRow label="フェーズ2 移動時間" value={fmt(summary.wsP2Move)} />
                      <SRow label="フェーズ2合計" value={fmt(summary.wsP2Total)} bg="#FFF3E0" color="#E65100" />
                      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 15, fontWeight: 700, borderTop: `2px solid ${ACCENT}`, paddingTop: 10 }}>
                        <span>総計</span><span style={{ color: ACCENT_DARK }}>{fmt(summary.wsGrand)} 分</span>
                      </div>
                    </div>
                  </div>
                  <div style={{ background: "#FFF", border: "1.5px solid #C8C090", borderRadius: 12, overflow: "hidden" }}>
                    <div style={{ background: "#C8C090", padding: "10px 18px" }}>
                      <span style={{ fontWeight: 700, fontSize: 13, color: "#1A1A0F" }}>生産性センター 需要量</span>
                    </div>
                    <div style={{ padding: "14px 18px" }}>
                      <DRow label="相談時間合計" value={fmt(summary.psConsult)} />
                      <DRow label="移動時間合計" value={fmt(summary.psMove)} />
                      <div style={{ fontSize: 11, color: "#AAA", marginBottom: 10 }}>※（相談＋移動）÷ 2</div>
                      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 15, fontWeight: 700, borderTop: "2px solid #C8C090", paddingTop: 10 }}>
                        <span>合計（÷2）</span><span style={{ color: "#6D6030" }}>{fmt(summary.psGrand)} 分</span>
                      </div>
                    </div>
                  </div>
                </div>

                {/* 担当者テーブル */}
                <div style={{ background: "#FFF", border: "1px solid #E8E0C8", borderRadius: 12, overflow: "hidden" }}>
                  <div style={{ padding: "12px 18px", borderBottom: `2px solid ${ACCENT}`, background: "#FFFDF0", display: "flex", alignItems: "center", gap: 8 }}>
                    <div style={{ width: 4, height: 18, background: ACCENT, borderRadius: 2 }} />
                    <span style={{ fontWeight: 700, fontSize: 13 }}>担当者別 稼働率・需要量</span>
                  </div>
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                      <thead>
                        <tr style={{ background: "#F5F0DC" }}>
                          {["担当者","総件数","出勤日数","稼働率","フェーズ1（分）","フェーズ2（分）","生産性C（分）"].map(h => (
                            <th key={h} style={{ padding: "9px 12px", textAlign: h==="担当者"?"left":"right", fontWeight: 600, color: "#5A4A00", borderBottom: "1px solid #E0D890", whiteSpace: "nowrap" }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {(staffFilter ? [staffFilter] : staffList).map((s, i) => {
                          const c = calcStats(rows, s);
                          const barW = maxRate > 0 ? Math.round(c.rate / maxRate * 68) : 0;
                          return (
                            <tr key={s}
                              style={{ background: i%2===0?"#FFF":"#FFFDF5", borderBottom: "1px solid #F0EBD0", transition: "background .1s" }}
                              onMouseEnter={e => e.currentTarget.style.background = ACCENT_PALE}
                              onMouseLeave={e => e.currentTarget.style.background = i%2===0?"#FFF":"#FFFDF5"}
                            >
                              <td style={{ padding: "8px 12px", fontWeight: 600 }}>{s}</td>
                              <td style={{ padding: "8px 12px", textAlign: "right" }}>{c.totalCount}</td>
                              <td style={{ padding: "8px 12px", textAlign: "right" }}>{c.uniqueDays}</td>
                              <td style={{ padding: "8px 12px" }}>
                                <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                                  <div style={{ height: 6, width: barW, background: ACCENT, borderRadius: 3, flexShrink: 0 }} />
                                  <span style={{ fontWeight: 700, color: ACCENT_DARK, minWidth: 34 }}>{c.rate.toFixed(2)}</span>
                                </div>
                              </td>
                              <td style={{ padding: "8px 12px", textAlign: "right" }}>{fmt(c.wsP1Total)}</td>
                              <td style={{ padding: "8px 12px", textAlign: "right" }}>{fmt(c.wsP2Total)}</td>
                              <td style={{ padding: "8px 12px", textAlign: "right" }}>{fmt(c.psGrand)}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </>
            )}
          </>
        )}
      </div>

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;600;700&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        select, button, input { font-family: inherit; }
        select:focus, button:focus { outline: 2px solid ${ACCENT}; outline-offset: 2px; }
      `}</style>
    </div>
  );
}

// ── サブコンポーネント ─────────────────────────────────────────────────────────
function SectionTitle({ children }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 18 }}>
      <div style={{ width: 4, height: 18, background: ACCENT, borderRadius: 2 }} />
      <span style={{ fontWeight: 700, fontSize: 14 }}>{children}</span>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <label style={{ fontSize: 12, color: "#666", display: "block", marginBottom: 5 }}>{label}</label>
      {children}
    </div>
  );
}

function KpiCard({ label, value, sub, dark }) {
  return (
    <div style={{ background: dark?"#1A1A0F":"#FFF", border: `1.5px solid ${dark?ACCENT:"#E8E0C8"}`, borderRadius: 10, padding: "14px 18px" }}>
      <div style={{ fontSize: 11, color: dark?ACCENT:"#888", marginBottom: 4, fontWeight: 600, letterSpacing: ".03em" }}>{label}</div>
      <div style={{ fontSize: dark?28:22, fontWeight: 700, color: dark?ACCENT:"#1A1A0F", lineHeight: 1.1 }}>{value}</div>
      <div style={{ fontSize: 11, color: dark?"#666":"#AAA", marginTop: 3 }}>{sub}</div>
    </div>
  );
}

function DRow({ label, value }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "#666", marginBottom: 6 }}>
      <span>{label}</span>
      <span style={{ fontWeight: 600, color: "#333" }}>{value} 分</span>
    </div>
  );
}

function SRow({ label, value, bg, color }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, fontWeight: 700, marginBottom: 10, padding: "6px 10px", background: bg, borderRadius: 6 }}>
      <span>{label}</span>
      <span style={{ color }}>{value} 分</span>
    </div>
  );
}

const selStyle = {
  width: "100%", padding: "8px 12px",
  border: `1px solid ${ACCENT}`, borderRadius: 6,
  fontSize: 13, background: ACCENT_LIGHT, fontWeight: 600,
};
