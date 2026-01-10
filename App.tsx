
import React, { useState } from 'react';
import { FileUp, ArrowRightLeft, AlertCircle, CheckCircle2, Download, Table, Trash2, PlusCircle, Sparkles, Loader2, BarChart3, Calculator, ShieldCheck, FilePlus2, FileMinus2, Info } from 'lucide-react';
import { GoogleGenAI } from "@google/genai";
import { parseExcelFile, compareData, exportReport } from './services/excelService';
import { ComparisonResult, SaleRow } from './types';

type TabType = 'modified' | 'added' | 'removed';

// 小工具函數：處理 B 欄位顯示（如果是日期則格式化）
const getColBValue = (row: SaleRow) => {
  const val = row._raw[1];
  if (typeof val === 'number' && val > 30000 && val < 60000) {
    const date = new Date((val - 25569) * 86400 * 1000);
    return `${date.getFullYear()}/${String(date.getMonth() + 1).padStart(2, '0')}/${String(date.getDate()).padStart(2, '0')}`;
  }
  return String(val || "").trim();
};

const App: React.FC = () => {
  const [oldFile, setOldFile] = useState<File | null>(null);
  const [newFile, setNewFile] = useState<File | null>(null);
  const [result, setResult] = useState<ComparisonResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [aiAnalyzing, setAiAnalyzing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<TabType>('modified');

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>, type: 'old' | 'new') => {
    const file = e.target.files?.[0] || null;
    if (type === 'old') setOldFile(file);
    else setNewFile(file);
    setResult(null);
    setError(null);
  };

  const runComparison = async () => {
    if (!oldFile || !newFile) {
      setError("請先上傳需要比對的兩個 Excel 檔案。");
      return;
    }

    setLoading(true);
    setError(null);
    try {
      const oldData = await parseExcelFile(oldFile);
      const newData = await parseExcelFile(newFile);
      const diffResult = compareData(oldData, newData);
      setResult(diffResult);
      if (diffResult.modified.length > 0) setActiveTab('modified');
      else if (diffResult.added.length > 0) setActiveTab('added');
      else if (diffResult.removed.length > 0) setActiveTab('removed');
    } catch (err: any) {
      setError(err.message || "比對過程中發生錯誤，請檢查檔案格式。");
    } finally {
      setLoading(false);
    }
  };

  const runAiSummary = async () => {
    if (!result) return;
    setAiAnalyzing(true);
    try {
      const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
      const prompt = `
        我比對了兩個業務銷售 Excel 檔案。以下是差異摘要：
        - 修改項次: ${result.modified.length}
        - 新增項次: ${result.added.length}
        - 移除項次: ${result.removed.length}

        顯著變更 (前15筆):
        ${result.modified.slice(0, 15).map(m => `- [${m.id}] ${m.projectName || '月份小計列'}: ${m.changes.map(c => `${c.column} 從 ${c.oldValue} 變為 ${c.newValue}`).join(', ')}`).join('\n')}

        新增資料範例 (前5筆):
        ${result.added.slice(0, 5).map(a => `- [${a.id}] ${a.projectName} (${a.customer})`).join('\n')}

        請以資深商業分析師的角度，用繁體中文提供一份精簡的執行摘要。重點分析這些數據變動（如數量 Qty 的異動、狀態變更）對業務的潛在意義。
      `;

      const response = await ai.models.generateContent({
        model: 'gemini-3-flash-preview',
        contents: prompt,
      });

      setResult(prev => prev ? { ...prev, summary: response.text } : null);
    } catch (err) {
      console.error("AI 分析失敗", err);
      setError("AI 分析暫時無法使用，但數據比對已完成。");
    } finally {
      setAiAnalyzing(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 pb-20 font-sans selection:bg-indigo-100 selection:text-indigo-900">
      <header className="bg-white border-b sticky top-0 z-10 shadow-sm backdrop-blur-md bg-white/80">
        <div className="max-w-6xl mx-auto px-4 py-4 flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="bg-indigo-600 p-2.5 rounded-2xl shadow-indigo-200 shadow-lg rotate-3">
              <ArrowRightLeft className="text-white w-6 h-6" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-xl font-black text-slate-800 tracking-tight leading-none">Smart Comparer</h1>
                <span className="flex items-center gap-1 px-2 py-0.5 bg-emerald-100 text-emerald-700 text-[10px] font-black rounded-full border border-emerald-200 uppercase tracking-tighter">
                  <ShieldCheck className="w-3 h-3" />
                  v1.1 Advanced
                </span>
              </div>
              <p className="text-[10px] text-slate-400 font-bold uppercase tracking-[0.2em] mt-1">智慧化 Excel 銷售數據分析系統</p>
            </div>
          </div>
          {result && (
            <button
              onClick={() => exportReport(result)}
              className="flex items-center gap-2 px-6 py-2.5 bg-slate-900 hover:bg-black text-white rounded-xl transition-all font-bold text-sm shadow-lg active:scale-95"
            >
              <Download className="w-4 h-4" />
              匯出對照報告
            </button>
          )}
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-8">
        <div className="grid grid-cols-1 md:grid-cols-2 gap-8 mb-10">
          <FileDropZone label="基準檔案 (Baseline / 舊版)" file={oldFile} onChange={(e) => handleFileChange(e, 'old')} onClear={() => setOldFile(null)} />
          <FileDropZone label="更新檔案 (Updated / 新版)" file={newFile} onChange={(e) => handleFileChange(e, 'new')} onClear={() => setNewFile(null)} />
        </div>

        <div className="flex justify-center mb-16">
          <button
            onClick={runComparison}
            disabled={!oldFile || !newFile || loading}
            className={`
              relative group flex items-center gap-4 px-14 py-4 rounded-2xl text-lg font-black shadow-2xl transition-all
              ${loading ? 'bg-slate-200 text-slate-400 cursor-not-allowed' : 'bg-indigo-600 hover:bg-indigo-700 text-white hover:shadow-indigo-200 active:scale-95 overflow-hidden'}
            `}
          >
            {loading ? <Loader2 className="w-6 h-6 animate-spin" /> : <BarChart3 className="w-6 h-6" />}
            {loading ? '數據深度分析中...' : '開始自動對照'}
          </button>
        </div>

        {error && (
          <div className="mb-10 p-5 bg-rose-50 border-2 border-rose-100 rounded-3xl flex items-center gap-4 text-rose-700">
            <AlertCircle className="w-6 h-6 text-rose-400" />
            <p className="font-bold">{error}</p>
          </div>
        )}

        {result && (
          <div className="space-y-12 animate-in fade-in slide-in-from-bottom-8 duration-700">
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-8">
              <StatCard title="偵測到異動" value={result.modified.length} icon={<ArrowRightLeft />} color="indigo" active={activeTab === 'modified'} onClick={() => setActiveTab('modified')} />
              <StatCard title="全新增項次" value={result.added.length} icon={<PlusCircle />} color="emerald" active={activeTab === 'added'} onClick={() => setActiveTab('added')} />
              <StatCard title="已移除項次" value={result.removed.length} icon={<Trash2 />} color="rose" active={activeTab === 'removed'} onClick={() => setActiveTab('removed')} />
            </div>

            <section className="bg-white rounded-[2.5rem] border border-slate-200 overflow-hidden shadow-sm">
              <div className="p-8 border-b bg-slate-50/50 flex items-center justify-between">
                <div className="flex items-center gap-4">
                  <div className="p-3 bg-indigo-100 text-indigo-600 rounded-2xl">
                    <Sparkles className="w-6 h-6" />
                  </div>
                  <div>
                    <h3 className="font-black text-slate-800 text-lg">AI 智能商業摘要</h3>
                  </div>
                </div>
                {!result.summary && (
                  <button onClick={runAiSummary} disabled={aiAnalyzing} className="px-6 py-3 bg-indigo-600 text-white rounded-xl font-bold text-sm flex items-center gap-2">
                    {aiAnalyzing ? <Loader2 className="w-4 h-4 animate-spin" /> : <Sparkles className="w-4 h-4" />}
                    解讀數據價值
                  </button>
                )}
              </div>
              <div className="p-10">
                {aiAnalyzing ? (
                  <div className="flex flex-col items-center py-12 gap-5">
                    <Loader2 className="w-12 h-12 text-indigo-500 animate-spin" />
                    <p className="text-slate-400 font-bold">分析趨勢中...</p>
                  </div>
                ) : result.summary ? (
                  <div className="bg-indigo-50/30 p-10 rounded-[2rem] text-slate-700 leading-relaxed whitespace-pre-wrap border border-indigo-100/50">
                    <div className="prose prose-slate max-w-none font-medium italic text-lg opacity-90">{result.summary}</div>
                  </div>
                ) : (
                  <p className="text-center text-slate-400 font-bold py-6">點擊按鈕獲取數據異動的深度分析報告。</p>
                )}
              </div>
            </section>

            <section className="bg-white rounded-[2.5rem] border border-slate-200 shadow-sm overflow-hidden">
              <div className="px-10 py-8 border-b bg-white flex flex-col md:flex-row md:items-center justify-between gap-6">
                <div className="flex items-center gap-4">
                  <Table className="text-indigo-600 w-7 h-7" />
                  <h3 className="font-black text-slate-800 text-xl tracking-tight">差異比對詳細清單</h3>
                </div>
                
                <div className="flex bg-slate-100 p-1.5 rounded-2xl gap-1">
                  <TabButton active={activeTab === 'modified'} onClick={() => setActiveTab('modified')} count={result.modified.length} label="異動內容" icon={<ArrowRightLeft className="w-3.5 h-3.5" />} />
                  <TabButton active={activeTab === 'added'} onClick={() => setActiveTab('added')} count={result.added.length} label="新增項次" icon={<FilePlus2 className="w-3.5 h-3.5" />} />
                  <TabButton active={activeTab === 'removed'} onClick={() => setActiveTab('removed')} count={result.removed.length} label="移除項次" icon={<FileMinus2 className="w-3.5 h-3.5" />} />
                </div>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-50/80 text-slate-400 text-[11px] uppercase tracking-[0.2em] font-black">
                      <th className="px-10 py-6 w-48">識別碼</th>
                      <th className="px-10 py-6">業務資訊</th>
                      <th className="px-10 py-6">詳情</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {activeTab === 'modified' && (
                      result.modified.length === 0 ? <EmptyState msg="目前沒有偵測到資料修改。" /> :
                      result.modified.map((diff, i) => (
                        <tr key={i} className="hover:bg-slate-50/50 transition-colors group">
                          <td className="px-10 py-10 align-top">
                            <span className="bg-slate-900 text-white px-4 py-2 rounded-xl font-mono text-xs font-black inline-block shadow-md">{diff.id}</span>
                          </td>
                          <td className="px-10 py-10 align-top">
                            <div className="font-black text-slate-800 text-lg">{diff.projectName || '彙總統計列'}</div>
                            <div className="text-xs text-slate-400 font-bold mt-1 uppercase tracking-wider">{diff.customer || '---'}</div>
                            {diff.colBValue && (
                              <div className="flex items-center gap-1.5 mt-2 px-2 py-1 bg-amber-50 text-amber-700 border border-amber-100 rounded-lg w-fit text-[10px] font-black">
                                <Info className="w-3 h-3" />
                                B 欄: {diff.colBValue}
                              </div>
                            )}
                          </td>
                          <td className="px-10 py-10 space-y-4">
                            {diff.changes.map((change, j) => (
                              <div key={j} className="flex flex-col lg:flex-row lg:items-center gap-4">
                                <span className="text-[10px] font-black px-3 py-1 bg-indigo-50 text-indigo-600 rounded-full border border-indigo-100 w-24 text-center">{change.column}</span>
                                <div className="flex items-center gap-3 text-sm font-bold">
                                  <span className="text-slate-300 line-through">{change.oldValue}</span>
                                  <ArrowRightLeft className="w-4 h-4 text-slate-300" />
                                  <span className="bg-indigo-600 text-white px-4 py-1.5 rounded-xl shadow-lg shadow-indigo-100">{change.newValue}</span>
                                </div>
                              </div>
                            ))}
                          </td>
                        </tr>
                      ))
                    )}

                    {activeTab === 'added' && (
                      result.added.length === 0 ? <EmptyState msg="沒有任何新增的項次。" /> :
                      result.added.map((row, i) => (
                        <tr key={i} className="hover:bg-emerald-50/30 transition-colors group">
                          <td className="px-10 py-10 align-top">
                            <span className="bg-emerald-600 text-white px-4 py-2 rounded-xl font-mono text-xs font-black inline-block shadow-md">NEW</span>
                            <div className="mt-2 text-xs font-bold text-slate-400">{row.id}</div>
                          </td>
                          <td className="px-10 py-10 align-top">
                            <div className="font-black text-slate-800 text-lg">{row.projectName}</div>
                            <div className="text-xs text-slate-400 font-bold mt-1 uppercase tracking-wider">{row.customer}</div>
                            <div className="flex items-center gap-1.5 mt-2 px-2 py-1 bg-emerald-50 text-emerald-700 border border-emerald-100 rounded-lg w-fit text-[10px] font-black">
                              <Info className="w-3 h-3" />
                              B 欄: {getColBValue(row)}
                            </div>
                          </td>
                          <td className="px-10 py-10">
                            <div className="flex flex-wrap gap-2">
                              <Badge label="日" value={row._raw[row._colMap.date]} />
                              <Badge label="狀態" value={row._raw[row._colMap.status]} />
                              <Badge label="產品" value={row._raw[row._colMap.product]} />
                            </div>
                          </td>
                        </tr>
                      ))
                    )}

                    {activeTab === 'removed' && (
                      result.removed.length === 0 ? <EmptyState msg="沒有任何移除的項次。" /> :
                      result.removed.map((row, i) => (
                        <tr key={i} className="hover:bg-rose-50/30 transition-colors group grayscale opacity-60">
                          <td className="px-10 py-10 align-top">
                            <span className="bg-rose-600 text-white px-4 py-2 rounded-xl font-mono text-xs font-black inline-block shadow-md">DEL</span>
                            <div className="mt-2 text-xs font-bold text-slate-400">{row.id}</div>
                          </td>
                          <td className="px-10 py-10 align-top">
                            <div className="font-black text-slate-800 text-lg line-through">{row.projectName}</div>
                            <div className="text-xs text-slate-400 font-bold mt-1 uppercase tracking-wider">{row.customer}</div>
                            <div className="flex items-center gap-1.5 mt-2 px-2 py-1 bg-slate-50 text-slate-500 border border-slate-100 rounded-lg w-fit text-[10px] font-black">
                              <Info className="w-3 h-3" />
                              B 欄: {getColBValue(row)}
                            </div>
                          </td>
                          <td className="px-10 py-10">
                            <span className="text-xs font-bold text-rose-500 bg-rose-50 px-3 py-1 rounded-full border border-rose-100">此項次已從更新檔中移除</span>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </section>
          </div>
        )}
      </main>
    </div>
  );
};

const TabButton: React.FC<{ active: boolean; onClick: () => void; count: number; label: string; icon: React.ReactNode }> = ({ active, onClick, count, label, icon }) => (
  <button
    onClick={onClick}
    className={`
      flex items-center gap-3 px-6 py-2.5 rounded-xl font-black text-sm transition-all
      ${active ? 'bg-white text-indigo-600 shadow-sm ring-1 ring-slate-200' : 'text-slate-400 hover:text-slate-600'}
    `}
  >
    {icon}
    {label}
    <span className={`px-2 py-0.5 rounded-full text-[10px] ${active ? 'bg-indigo-100 text-indigo-600' : 'bg-slate-200 text-slate-500'}`}>{count}</span>
  </button>
);

const Badge: React.FC<{ label: string; value: any }> = ({ label, value }) => (
  <div className="flex items-center bg-slate-100 rounded-lg overflow-hidden border border-slate-200">
    <span className="text-[10px] font-black bg-slate-200 px-2 py-1 text-slate-600 uppercase tracking-tighter">{label}</span>
    <span className="px-3 py-1 text-xs font-bold text-slate-700">{value || '---'}</span>
  </div>
);

const EmptyState: React.FC<{ msg: string }> = ({ msg }) => (
  <tr>
    <td colSpan={3} className="px-10 py-32 text-center text-slate-400 font-bold text-lg italic">{msg}</td>
  </tr>
);

const FileDropZone: React.FC<{ label: string; file: File | null; onChange: (e: React.ChangeEvent<HTMLInputElement>) => void; onClear: () => void; }> = ({ label, file, onChange, onClear }) => (
  <div className="flex flex-col gap-4">
    <label className="text-xs font-black text-slate-500 uppercase tracking-[0.25em] ml-3">{label}</label>
    <div className={`relative border-2 border-dashed rounded-[2.5rem] p-10 h-56 flex flex-col items-center justify-center gap-6 group transition-all ${file ? 'border-indigo-400 bg-indigo-50/40 shadow-inner' : 'border-slate-200 bg-white hover:border-indigo-300'}`}>
      {file ? (
        <>
          <CheckCircle2 className="text-indigo-600 w-10 h-10 animate-in zoom-in duration-300" />
          <div className="text-center w-full px-8">
            <p className="font-black text-slate-800 text-base truncate">{file.name}</p>
            <p className="text-[10px] text-indigo-500 font-black uppercase mt-2 tracking-widest bg-indigo-100/50 px-3 py-1 rounded-full inline-block">Ready to process</p>
          </div>
          <button onClick={(e) => { e.stopPropagation(); onClear(); }} className="absolute top-6 right-6 p-3 text-slate-300 hover:text-rose-500 transition-colors"><Trash2 className="w-6 h-6" /></button>
        </>
      ) : (
        <>
          <FileUp className="w-12 h-12 text-slate-300 group-hover:text-indigo-400 transition-all duration-500" />
          <div className="text-center">
            <p className="text-base font-black text-slate-700">點擊或拖放上傳</p>
            <p className="text-[11px] text-slate-400 font-bold mt-2 uppercase tracking-widest">Excel / CSV 格式</p>
          </div>
          <input type="file" onChange={onChange} className="absolute inset-0 opacity-0 cursor-pointer" accept=".xlsx,.xls,.csv" />
        </>
      )}
    </div>
  </div>
);

const StatCard: React.FC<{ title: string; value: number; icon: React.ReactNode; color: 'indigo' | 'emerald' | 'rose'; active: boolean; onClick: () => void }> = ({ title, value, icon, color, active, onClick }) => {
  const styles = {
    indigo: active ? 'border-indigo-500 ring-4 ring-indigo-50 shadow-indigo-100' : 'border-slate-100 hover:border-indigo-200',
    emerald: active ? 'border-emerald-500 ring-4 ring-emerald-50 shadow-emerald-100' : 'border-slate-100 hover:border-emerald-200',
    rose: active ? 'border-rose-500 ring-4 ring-rose-50 shadow-rose-100' : 'border-slate-100 hover:border-rose-200'
  };
  const iconColors = {
    indigo: 'text-indigo-600 bg-indigo-50',
    emerald: 'text-emerald-600 bg-emerald-50',
    rose: 'text-rose-600 bg-rose-50'
  };

  return (
    <div onClick={onClick} className={`bg-white p-10 rounded-[2.5rem] border-2 flex items-center justify-between cursor-pointer transition-all duration-500 shadow-sm hover:shadow-xl ${styles[color]}`}>
      <div className="space-y-2">
        <div className="text-[11px] font-black uppercase tracking-[0.25em] text-slate-400">{title}</div>
        <div className="text-6xl font-black text-slate-900 tracking-tighter">{value}</div>
      </div>
      <div className={`p-6 rounded-3xl ${iconColors[color]} transition-transform duration-700 group-hover:scale-110`}>
        {React.isValidElement(icon) ? React.cloneElement(icon as React.ReactElement<any>, { className: 'w-10 h-10' }) : icon}
      </div>
    </div>
  );
};

export default App;
