import React, { useState, useCallback, useRef } from 'react';
import * as XLSX from 'xlsx';
import { Upload, Plus, Trash2, Download, Table2, AlertCircle, CheckCircle2, ChevronRight, FileSpreadsheet, ArrowRight, RefreshCw, Loader2, X, Wand2 } from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';

// ── XLSX helper (npm version) ─────────────────────────────────────────────────
const getXLSX = () => Promise.resolve(XLSX);

// ── Download helpers ──────────────────────────────────────────────────────────
const triggerDownload = (blobUrl, filename) => {
  const a = document.createElement('a');
  a.href = blobUrl; a.download = filename;
  a.style.cssText = 'position:fixed;top:-9999px;left:-9999px;';
  document.body.appendChild(a);
  a.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(blobUrl); }, 300);
};

const downloadCSV = (data, filename) => {
  if (!data.length) return;
  const cols = Object.keys(data[0]);
  const esc = v => { const s = String(v ?? ''); return (s.includes(',') || s.includes('"') || s.includes('\n')) ? `"${s.replace(/"/g, '""')}"` : s; };
  const csv = [cols.map(esc).join(','), ...data.map(r => cols.map(c => esc(r[c])).join(','))].join('\r\n');
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
  triggerDownload(URL.createObjectURL(blob), filename);
};

const downloadXLSXFile = async (data, filename) => {
  const XLSX = await getXLSX();
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  triggerDownload(URL.createObjectURL(blob), filename);
};

// ── CSV parser ────────────────────────────────────────────────────────────────
const parseCSV = (text) => {
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n').filter(l => l.trim());
  if (!lines.length) return [];
  const parseRow = (line) => {
    const res = []; let cur = ''; let inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') { if (inQ && line[i + 1] === '"') { cur += '"'; i++; } else inQ = !inQ; }
      else if (ch === ',' && !inQ) { res.push(cur.trim()); cur = ''; }
      else cur += ch;
    }
    res.push(cur.trim()); return res;
  };
  const headers = parseRow(lines[0]);
  return lines.slice(1).map(line => {
    const vals = parseRow(line); const row = {};
    headers.forEach((h, i) => { row[h] = vals[i] ?? ''; });
    return row;
  }).filter(row => Object.values(row).some(v => v !== ''));
};

const parseFile = async (file) => {
  if (file.name.toLowerCase().endsWith('.csv')) return parseCSV(await file.text());
  const XLSX = await getXLSX();
  const ab = await file.arrayBuffer();
  const wb = XLSX.read(new Uint8Array(ab), { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: '' });
};

// ── Helpers ───────────────────────────────────────────────────────────────────
const nk = v => (v === null || v === undefined) ? '' : String(v).trim().toLowerCase();

const autoDetectMaps = (t1cols, t2cols, k1, k2) => {
  const t2Set = new Set(t2cols);
  const excl  = new Set([k1, k2]);
  return t1cols.filter(c => t2Set.has(c) && !excl.has(c)).map(c => ({ from: c, to: c }));
};

// ── Styles ────────────────────────────────────────────────────────────────────
const SEL   = "w-full bg-gray-800 border border-gray-700 rounded-xl px-3 py-2.5 text-sm text-white focus:outline-none focus:border-blue-500 transition cursor-pointer";
const INPUT = "w-full bg-gray-800 border border-gray-700 rounded-xl px-3 py-2.5 text-sm text-white placeholder-gray-600 focus:outline-none focus:border-blue-500 transition";

// ── StepBar ───────────────────────────────────────────────────────────────────
const StepBar = ({ step }) => (
  <div className="flex items-center gap-2">
    {['上傳 & 設定比對', '欄位對應', '結果 & 下載'].map((s, i) => {
      const n = i + 1, done = step > n, active = step === n;
      return (
        <React.Fragment key={s}>
          <div className="flex items-center gap-1.5 min-w-0">
            <div className={`w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold flex-shrink-0 transition-all
              ${done ? 'bg-green-600' : active ? 'bg-blue-600' : 'bg-gray-800 text-gray-600 border border-gray-700'}`}>
              {done ? <CheckCircle2 size={11} /> : n}
            </div>
            <span className={`text-xs font-semibold whitespace-nowrap ${active ? 'text-white' : done ? 'text-green-400' : 'text-gray-600'}`}>{s}</span>
          </div>
          {i < 2 && <div className="flex-1 h-px bg-gray-800 min-w-1" />}
        </React.Fragment>
      );
    })}
  </div>
);

// ── FileRow ───────────────────────────────────────────────────────────────────
const FileRow = ({ num, label, info, busy, accent, onClear, onFile }) => {
  const ref = useRef(null);
  const [drag, setDrag] = useState(false);
  const iB = accent === 'blue';
  return (
    <div className={`rounded-xl border-2 transition-all p-4
      ${info ? (iB ? 'border-blue-500/40 bg-blue-500/5' : 'border-purple-500/40 bg-purple-500/5')
              : drag ? 'border-blue-500 bg-blue-500/5' : 'border-dashed border-gray-700 bg-gray-900/60'}`}
      onDragOver={e => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={e => { e.preventDefault(); setDrag(false); const f = e.dataTransfer.files[0]; if (f) onFile(f); }}>
      <div className="flex items-center gap-3">
        <div className={`w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold flex-shrink-0
          ${info ? (iB ? 'bg-blue-600' : 'bg-purple-600') : 'bg-gray-700 text-gray-500'}`}>{num}</div>
        <span className="text-sm font-semibold text-gray-300 flex-shrink-0">{label}</span>
        {info ? (
          <div className={`flex items-center gap-2 flex-1 min-w-0 ${iB ? 'text-blue-300' : 'text-purple-300'}`}>
            <FileSpreadsheet size={12} className="flex-shrink-0" />
            <span className="text-xs font-mono truncate">{info.name}</span>
            <span className="text-xs opacity-60 flex-shrink-0">{info.cols.length}欄·{info.data.length}筆</span>
          </div>
        ) : (
          <div onClick={() => ref.current?.click()}
            className="flex items-center gap-2 flex-1 cursor-pointer group">
            <input ref={ref} type="file" accept=".xlsx,.xls,.csv" className="hidden"
              onChange={e => { const f = e.target.files[0]; if (f) { onFile(f); e.target.value = ''; } }} />
            {busy
              ? <Loader2 size={13} className="text-blue-400 animate-spin" />
              : <span className="text-xs text-gray-600 group-hover:text-blue-400 transition">點擊或拖入 .xlsx / .csv</span>}
          </div>
        )}
        {info && <button onClick={onClear} className="text-gray-700 hover:text-red-400 transition flex-shrink-0"><RefreshCw size={13} /></button>}
      </div>
    </div>
  );
};

// ── Toggle ────────────────────────────────────────────────────────────────────
const Toggle = ({ on, onChange }) => (
  <button onClick={() => onChange(!on)}
    className={`relative w-10 h-5 rounded-full transition-colors flex-shrink-0 ${on ? 'bg-blue-600' : 'bg-gray-700'}`}>
    <div className={`absolute top-0.5 w-4 h-4 bg-white rounded-full transition-transform ${on ? 'translate-x-5' : 'translate-x-0.5'}`} />
  </button>
);

// ── Main ──────────────────────────────────────────────────────────────────────
export default function TableMapper() {
  const [step,           setStep]           = useState(1);
  const [t1,             setT1]             = useState(null);
  const [t2,             setT2]             = useState(null);
  const [busy1,          setBusy1]          = useState(false);
  const [busy2,          setBusy2]          = useState(false);
  const [key1,           setKey1]           = useState('');
  const [key2,           setKey2]           = useState('');
  const [maps,           setMaps]           = useState([]);
  const [customTo,       setCustomTo]       = useState({});
  const [autoCount,      setAutoCount]      = useState(0);
  const [inclMissing,    setInclMissing]    = useState(false);
  const [result,         setResult]         = useState(null);
  const [tab,            setTab]            = useState('filled');
  const [processing,     setProcessing]     = useState(false);
  const [dlState,        setDlState]        = useState('idle');
  const [error,          setError]          = useState('');

  const handleFile = useCallback(async (file, num) => {
    setError('');
    if (num === 1) setBusy1(true); else setBusy2(true);
    try {
      const data = await parseFile(file);
      const cols = data.length > 0 ? Object.keys(data[0]) : [];
      const info = { name: file.name, data, cols };
      if (num === 1) { setT1(info); setKey1(cols[0] ?? ''); }
      else           { setT2(info); setKey2(cols[0] ?? ''); }
    } catch (e) {
      setError('讀取失敗，請確認格式正確。');
    } finally {
      if (num === 1) setBusy1(false); else setBusy2(false);
    }
  }, []);

  const goToStep2 = () => {
    if (!t1 || !t2 || !key1 || !key2) return;
    const detected = autoDetectMaps(t1.cols, t2.cols, key1, key2);
    const t2Set = new Set(t2.cols);
    const excl  = new Set([key1, key2]);
    const extraCols = t1.cols.filter(c => !t2Set.has(c) && !excl.has(c));
    const extra = extraCols.slice(0, 5).map(c => ({ from: c, to: '__new__', _customTo: c }));
    const initMaps = detected.length > 0 || extra.length > 0
      ? [...detected, ...extra]
      : [{ from: '', to: '' }];
    setMaps(initMaps);
    const initCustomTo = {};
    extra.forEach((m, i) => { initCustomTo[detected.length + i] = m._customTo; });
    setCustomTo(initCustomTo);
    setAutoCount(detected.length);
    setStep(2);
  };

  const addMap = () => setMaps(p => [...p, { from: '', to: '' }]);
  const delMap = i => {
    setMaps(p => p.filter((_, j) => j !== i));
    setCustomTo(p => {
      const next = {};
      Object.entries(p).forEach(([k, v]) => {
        const ki = parseInt(k);
        if (ki < i) next[ki] = v;
        else if (ki > i) next[ki - 1] = v;
      });
      return next;
    });
  };
  const updMap = (i, k, v) => setMaps(p => p.map((m, j) => j === i ? { ...m, [k]: v } : m));

  const getToCol = (m, i) => {
    if (m.to === '__new__') return (customTo[i] ?? '').trim() || m.from;
    if (!m.to) return m.from;
    return m.to;
  };

  const run = async () => {
    if (!t1 || !t2 || !key1 || !key2) return;
    setProcessing(true); setError('');
    try {
      const lk = {};
      for (const r of t1.data) {
        const k = nk(r[key1]);
        if (k !== '') lk[k] = r;
      }
      const validMaps = maps.reduce((acc, m, i) => {
        const from = (m.from ?? '').trim();
        if (!from) return acc;
        const to = getToCol(m, i) || from;
        acc.push({ from, to });
        return acc;
      }, []);
      const t2KeySet = new Set(t2.data.map(r => nk(r[key2])));
      const missing = t1.data.filter(r => {
        const k = nk(r[key1]);
        return k !== '' && !t2KeySet.has(k);
      });
      let matched = 0;
      const updated = t2.data.map(row => {
        const k = nk(row[key2]);
        const src = lk[k];
        if (!src) return { ...row };
        matched++;
        const nr = { ...row };
        for (const m of validMaps) {
          if (m.from in src) nr[m.to] = src[m.from];
        }
        return nr;
      });
      const addedCols = [...new Set(validMaps.map(m => m.to).filter(c => c && !t2.cols.includes(c)))];
      const newCols   = [...t2.cols, ...addedCols];
      const finalOutput = inclMissing ? [...updated, ...missing] : updated;
      setResult({ missing, updated, finalOutput, matched, newCols, addedCols, validMapsCount: validMaps.length });
      setStep(3); setTab(matched > 0 ? 'filled' : 'missing'); setDlState('idle');
    } catch (e) {
      setError('比對時發生錯誤：' + e.message);
    } finally {
      setProcessing(false);
    }
  };

  const handleDownloadXLSX = async () => {
    if (!result || dlState === 'downloading') return;
    setDlState('downloading'); setError('');
    try {
      const base = t2?.name.replace(/\.[^.]+$/, '') ?? 'result';
      await downloadXLSXFile(result.finalOutput, `updated_${base}.xlsx`);
      setDlState('done'); setTimeout(() => setDlState('idle'), 2500);
    } catch (e) {
      setError('Excel 下載失敗，請改用 CSV。');
      setDlState('error'); setTimeout(() => setDlState('idle'), 2500);
    }
  };

  const handleDownloadCSV = () => {
    if (!result) return;
    const base = t2?.name.replace(/\.[^.]+$/, '') ?? 'result';
    try { downloadCSV(result.finalOutput, `updated_${base}.csv`); }
    catch (e) { setError('下載失敗。'); }
  };

  const missingCols = result?.missing.length > 0 ? Object.keys(result.missing[0]) : (t1?.cols ?? []);
  const filledCols  = result?.newCols ?? [];

  return (
    <div className="min-h-screen bg-gray-950 text-white font-sans">
      <header className="sticky top-0 z-10 border-b border-gray-800 bg-gray-950/90 backdrop-blur px-5 py-3 flex items-center gap-3">
        <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center"><Table2 size={15} /></div>
        <span className="font-bold tracking-tight">Table Mapper</span>
      </header>

      <div className="max-w-2xl mx-auto px-4 py-8 space-y-5">
        <StepBar step={step} />

        <AnimatePresence>
          {error && (
            <motion.div initial={{ opacity: 0, y: -6 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0 }}
              className="flex items-center gap-2 bg-red-500/10 border border-red-500/30 rounded-xl px-4 py-3 text-xs text-red-400">
              <AlertCircle size={13} className="flex-shrink-0" />{error}
              <button onClick={() => setError('')} className="ml-auto"><X size={12} /></button>
            </motion.div>
          )}
        </AnimatePresence>

        <AnimatePresence mode="wait">
          {step === 1 && (
            <motion.div key="s1" initial={{ opacity: 0, y: 12 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -12 }} className="space-y-4">
              <div className="bg-gray-900 border border-gray-800 rounded-2xl p-5 space-y-3">
                <p className="text-xs text-gray-500 font-semibold uppercase tracking-widest mb-1">① 上傳兩份表格</p>
                <FileRow num={1} label="表一（來源）" accent="blue"  info={t1} busy={busy1}
                  onClear={() => { setT1(null); setKey1(''); }} onFile={f => handleFile(f, 1)} />
                <FileRow num={2} label="表二（目標）" accent="purple" info={t2} busy={busy2}
                  onClear={() => { setT2(null); setKey2(''); }} onFile={f => handleFile(f, 2)} />
              </div>
              <div className={`bg-gray-900 border border-gray-800 rounded-2xl p-5 space-y-3 transition-opacity ${t1 && t2 ? 'opacity-100' : 'opacity-40'}`}>
                <p className="text-xs text-gray-500 font-semibold uppercase tracking-widest mb-1">② 選擇比對欄位</p>
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs text-gray-500 mb-1.5">表一比對欄位</label>
                    <select value={key1} onChange={e => setKey1(e.target.value)} className={SEL} disabled={!t1}>
                      {!t1 && <option>— 請先上傳 —</option>}
                      {(t1?.cols ?? []).map(c => <option key={c}>{c}</option>)}
                    </select>
                    {t1 && <p className="text-xs text-gray-600 mt-1">共 {t1.data.length} 筆</p>}
                  </div>
                  <div>
                    <label className="block text-xs text-gray-500 mb-1.5">表二比對欄位</label>
                    <select value={key2} onChange={e => setKey2(e.target.value)} className={SEL} disabled={!t2}>
                      {!t2 && <option>— 請先上傳 —</option>}
                      {(t2?.cols ?? []).map(c => <option key={c}>{c}</option>)}
                    </select>
                    {t2 && <p className="text-xs text-gray-600 mt-1">共 {t2.data.length} 筆</p>}
                  </div>
                </div>
                {t1 && t2 && key1 && key2 && (
                  <div className="flex items-center gap-2 bg-gray-800 rounded-xl px-3 py-2">
                    <span className="text-xs font-mono text-blue-300 bg-blue-500/15 px-2 py-0.5 rounded-lg">{key1}</span>
                    <ArrowRight size={11} className="text-gray-600" />
                    <span className="text-xs text-gray-500">比對</span>
                    <ArrowRight size={11} className="text-gray-600" />
                    <span className="text-xs font-mono text-purple-300 bg-purple-500/15 px-2 py-0.5 rounded-lg">{key2}</span>
                    <span className="ml-auto text-xs text-green-400 flex items-center gap-1">
                      <Wand2 size={11} /> 下一步自動偵測欄位對應
                    </span>
                  </div>
                )}
              </div>
              <button onClick={goToStep2} disabled={!t1 || !t2 || !key1 || !key2 || busy1 || busy2}
                className={`w-full flex items-center justify-center gap-2 py-3.5 rounded-xl text-sm font-bold transition-all
                  ${t1 && t2 && key1 && key2 && !busy1 && !busy2
                    ? 'bg-blue-600 hover:bg-blue-500 active:scale-95'
                    : 'bg-gray-800 text-gray-600 cursor-not-allowed'}`}>
                下一步：欄位對應設定 <ChevronRight size={15} />
              </button>
            </motion.div>
          )}

          {step === 2 && (
            <motion.div key="s2" initial={{ opacity: 0, y: 12 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -12 }} className="space-y-4">
              {autoCount > 0 ? (
                <div className="flex items-center gap-2 bg-green-500/10 border border-green-500/30 rounded-xl px-4 py-3">
                  <Wand2 size={13} className="text-green-400 flex-shrink-0" />
                  <p className="text-xs text-green-400">自動偵測到 <strong>{autoCount}</strong> 個同名欄位，已預填對應。請確認無誤後執行比對</p>
                </div>
              ) : (
                <div className="flex items-center gap-2 bg-amber-500/10 border border-amber-500/30 rounded-xl px-4 py-3">
                  <AlertCircle size={13} className="text-amber-400 flex-shrink-0" />
                  <p className="text-xs text-amber-400">未找到同名欄位，請手動設定對應關係</p>
                </div>
              )}
              <div className="bg-gray-900 border border-gray-800 rounded-2xl p-5 space-y-4">
                <div className="flex items-center justify-between">
                  <div>
                    <p className="text-sm font-semibold text-gray-200">欄位對應</p>
                    <p className="text-xs text-gray-500 mt-0.5">表一欄位 → 填入表二欄位</p>
                  </div>
                  <button onClick={addMap}
                    className="flex items-center gap-1.5 text-xs text-blue-400 border border-blue-500/30 bg-blue-500/10 hover:bg-blue-500/20 px-3 py-1.5 rounded-xl transition">
                    <Plus size={12} /> 新增
                  </button>
                </div>
                {maps.length === 0 && (
                  <div className="text-center py-5 text-xs text-gray-600 border border-dashed border-gray-800 rounded-xl">
                    點擊「新增」手動設定欄位對應
                  </div>
                )}
                <div className="space-y-2">
                  {maps.map((m, i) => (
                    <motion.div key={i} initial={{ opacity: 0, y: 4 }} animate={{ opacity: 1, y: 0 }} className="flex items-start gap-2">
                      <div className="flex-1">
                        {i === 0 && <label className="block text-xs text-gray-600 mb-1.5">來源（表一）</label>}
                        <select value={m.from} onChange={e => updMap(i, 'from', e.target.value)} className={SEL}>
                          <option value="">選擇欄位…</option>
                          {(t1?.cols ?? []).map(c => <option key={c}>{c}</option>)}
                        </select>
                      </div>
                      <div className={`flex-shrink-0 ${i === 0 ? 'mt-7' : 'mt-2.5'}`}>
                        <ArrowRight size={14} className="text-gray-600" />
                      </div>
                      <div className="flex-1">
                        {i === 0 && <label className="block text-xs text-gray-600 mb-1.5">目標（表二）</label>}
                        <select value={m.to} onChange={e => updMap(i, 'to', e.target.value)} className={SEL}>
                          <option value="">同名欄位（自動）</option>
                          {(t2?.cols ?? []).map(c => <option key={c}>{c}</option>)}
                          <option value="__new__">＋ 建立新欄位…</option>
                        </select>
                        {m.to === '__new__' && (
                          <input type="text" placeholder="新欄位名稱"
                            value={customTo[i] ?? ''}
                            onChange={e => setCustomTo(p => ({ ...p, [i]: e.target.value }))}
                            className={`${INPUT} mt-1.5`} />
                        )}
                      </div>
                      <button onClick={() => delMap(i)} className={`flex-shrink-0 text-gray-700 hover:text-red-400 transition ${i === 0 ? 'mt-7' : 'mt-2.5'}`}>
                        <Trash2 size={14} />
                      </button>
                    </motion.div>
                  ))}
                </div>
                {maps.length > 0 && (
                  <div className="bg-blue-500/5 border border-blue-500/20 rounded-xl px-4 py-3">
                    <p className="text-xs text-blue-300">
                      共 <strong>{maps.filter(m => m.from).length}</strong> 個有效對應，比對成功的列將把上述欄位值從表一複製到表二
                    </p>
                  </div>
                )}
              </div>
              <div className="bg-gray-900 border border-gray-800 rounded-2xl px-5 py-4 flex items-center justify-between gap-4">
                <div>
                  <p className="text-sm text-gray-300 font-medium">附加表一有、表二缺的行</p>
                  <p className="text-xs text-gray-600 mt-0.5">開啟後，缺漏行會附加在下載檔案末尾</p>
                </div>
                <Toggle on={inclMissing} onChange={setInclMissing} />
              </div>
              <div className="grid grid-cols-2 gap-3">
                <button onClick={() => setStep(1)}
                  className="flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-semibold text-gray-400 border border-gray-700 hover:bg-gray-800 transition">
                  上一步
                </button>
                <button onClick={run} disabled={processing}
                  className={`flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-bold transition-all
                    ${!processing ? 'bg-blue-600 hover:bg-blue-500 active:scale-95' : 'bg-gray-800 text-gray-500 cursor-not-allowed'}`}>
                  {processing ? <><Loader2 size={14} className="animate-spin" /> 比對中…</> : <>執行比對 <ChevronRight size={15} /></>}
                </button>
              </div>
            </motion.div>
          )}

          {step === 3 && result && (
            <motion.div key="s3" initial={{ opacity: 0, y: 12 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -12 }} className="space-y-4">
              {result.validMapsCount === 0 && (
                <div className="flex items-center gap-2 bg-amber-500/10 border border-amber-500/30 rounded-xl px-4 py-3">
                  <AlertCircle size={13} className="text-amber-400 flex-shrink-0" />
                  <p className="text-xs text-amber-400">未設定任何欄位對應，輸出與原表二相同。請返回設定</p>
                </div>
              )}
              {result.matched === 0 && result.validMapsCount > 0 && (
                <div className="flex items-start gap-2 bg-red-500/10 border border-red-500/30 rounded-xl px-4 py-3">
                  <AlertCircle size={13} className="text-red-400 flex-shrink-0 mt-0.5" />
                  <div>
                    <p className="text-xs text-red-400 font-semibold">比對欄位值完全沒有匹配</p>
                    <p className="text-xs text-red-300/70 mt-0.5">請確認兩份表格的 Key 欄格式一致（數字/文字/前導零）</p>
                  </div>
                </div>
              )}
              <div className="grid grid-cols-3 gap-3">
                {[
                  { label: '表二總筆數',   val: result.updated.length, color: 'text-white',     bg: 'bg-gray-800 border-gray-700' },
                  { label: '成功比對填入', val: result.matched,        color: 'text-green-400', bg: 'bg-green-500/10 border-green-500/30' },
                  { label: '表一有表二缺', val: result.missing.length, color: 'text-amber-400', bg: 'bg-amber-500/10 border-amber-500/30' },
                ].map(({ label, val, color, bg }) => (
                  <div key={label} className={`rounded-2xl border p-4 text-center ${bg}`}>
                    <p className={`text-2xl font-bold font-mono ${color}`}>{val}</p>
                    <p className="text-xs text-gray-500 mt-1 leading-tight">{label}</p>
                  </div>
                ))}
              </div>
              <div className="bg-gray-900 border border-gray-800 rounded-2xl overflow-hidden">
                <div className="flex border-b border-gray-800">
                  {[
                    { k: 'filled',  label: `填入結果 (${result.updated.length})`,  active: 'border-blue-500 text-blue-400 bg-blue-500/5' },
                    { k: 'missing', label: `缺漏行 (${result.missing.length})`,     active: 'border-amber-500 text-amber-400 bg-amber-500/5' },
                  ].map(({ k, label, active }) => (
                    <button key={k} onClick={() => setTab(k)}
                      className={`flex-1 py-3 text-xs font-semibold border-b-2 transition
                        ${tab === k ? active : 'border-transparent text-gray-600 hover:text-gray-400'}`}>
                      {label}
                    </button>
                  ))}
                </div>
                <div className="p-4 overflow-x-auto" style={{ maxHeight: '340px', overflowY: 'auto' }}>
                  {tab === 'filled' && (
                    <div className="space-y-2">
                      {result.addedCols.length > 0 && (
                        <div className="flex items-center gap-2 flex-wrap mb-2">
                          <CheckCircle2 size={12} className="text-green-400" />
                          <span className="text-xs text-green-400">新增欄位：</span>
                          {result.addedCols.map(c => (
                            <span key={c} className="text-xs font-mono text-green-300 bg-green-500/10 px-2 py-0.5 rounded-lg">{c}</span>
                          ))}
                        </div>
                      )}
                      <table className="w-full text-xs">
                        <thead>
                          <tr className="text-gray-600 border-b border-gray-800">
                            {filledCols.slice(0, 8).map(c => (
                              <th key={c} className={`text-left pb-2 pr-4 font-semibold whitespace-nowrap ${result.addedCols.includes(c) ? 'text-green-500' : ''}`}>
                                {c}{result.addedCols.includes(c) ? ' ✦' : ''}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {result.updated.slice(0, 50).map((row, i) => (
                            <tr key={i} className="border-b border-gray-800/40">
                              {filledCols.slice(0, 8).map((c, j) => (
                                <td key={j} className={`py-1.5 pr-4 font-mono whitespace-nowrap ${result.addedCols.includes(c) ? 'text-green-300/80' : 'text-gray-400'}`}>
                                  {String(row[c] ?? '')}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                      {result.updated.length > 50 && <p className="text-xs text-gray-600 text-center pt-2">僅顯示前 50 筆</p>}
                    </div>
                  )}
                  {tab === 'missing' && (
                    result.missing.length === 0 ? (
                      <div className="text-center py-8 text-xs text-green-400">
                        <CheckCircle2 size={20} className="mx-auto mb-2" />表二已涵蓋表一所有資料
                      </div>
                    ) : (
                      <table className="w-full text-xs">
                        <thead>
                          <tr className="text-gray-600 border-b border-gray-800">
                            {missingCols.slice(0, 8).map(c => <th key={c} className="text-left pb-2 pr-4 font-semibold whitespace-nowrap">{c}</th>)}
                          </tr>
                        </thead>
                        <tbody>
                          {result.missing.map((row, i) => (
                            <tr key={i} className="border-b border-gray-800/40">
                              {missingCols.slice(0, 8).map((c, j) => (
                                <td key={j} className="py-1.5 pr-4 font-mono text-amber-200/70 whitespace-nowrap">{String(row[c] ?? '')}</td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )
                  )}
                </div>
              </div>
              <div className="bg-gray-900 border border-gray-800 rounded-2xl p-5 space-y-3">
                <div className="flex items-center justify-between">
                  <p className="text-xs text-gray-400 font-semibold uppercase tracking-widest">下載結果</p>
                  <p className="text-xs text-gray-500">
                    {result.finalOutput.length} 筆
                    {inclMissing && result.missing.length > 0 && <span className="text-blue-400 ml-1">（含 {result.missing.length} 缺漏行）</span>}
                  </p>
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <button onClick={handleDownloadXLSX} disabled={dlState === 'downloading'}
                    className={`flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-bold transition-all
                      ${dlState === 'done'        ? 'bg-green-600' :
                        dlState === 'error'       ? 'bg-red-600/40 text-red-300 border border-red-500/30' :
                        dlState === 'downloading' ? 'bg-gray-800 text-gray-500 cursor-not-allowed' :
                                                    'bg-blue-600 hover:bg-blue-500 active:scale-95'}`}>
                    {dlState === 'downloading' ? <><Loader2 size={14} className="animate-spin" />處理中…</> :
                     dlState === 'done'        ? <><CheckCircle2 size={14} />完成</> :
                     dlState === 'error'       ? <><X size={14} />失敗，改用 CSV</> :
                                                 <><Download size={14} />下載 .xlsx</>}
                  </button>
                  <button onClick={handleDownloadCSV}
                    className="flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-bold border border-gray-700 text-gray-300 hover:bg-gray-800 hover:text-white active:scale-95 transition-all">
                    <Download size={14} />下載 .csv
                  </button>
                </div>
                <p className="text-xs text-gray-600 text-center">.xlsx 失敗時請用 .csv（Excel 可直接開啟）</p>
              </div>
              <button onClick={() => setStep(2)}
                className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-semibold text-gray-400 border border-gray-700 hover:bg-gray-800 transition">
                返回修改設定
              </button>
            </motion.div>
          )}
        </AnimatePresence>
      </div>

      <footer className="text-center py-6 text-xs text-gray-600">
        created by Tabris Yuen @2026
      </footer>
    </div>
  );
}
