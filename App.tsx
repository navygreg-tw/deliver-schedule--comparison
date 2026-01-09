
import React, { useState } from 'react';
import { FileUp, ArrowRightLeft, AlertCircle, CheckCircle2, Download, Table, Trash2, PlusCircle, Sparkles, Loader2, BarChart3, Calculator, ShieldCheck } from 'lucide-react';
import { GoogleGenAI } from "@google/genai";
import { parseExcelFile, compareData, exportReport } from './services/excelService';
import { ComparisonResult } from './types';

const App: React.FC = () => {
  const [oldFile, setOldFile] = useState<File | null>(null);
  const [newFile, setNewFile] = useState<File | null>(null);
  const [result, setResult] = useState<ComparisonResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [aiAnalyzing, setAiAnalyzing] = useState(false);
  const [error, setError] = useState<string | null>(null);

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
      // Create a new GoogleGenAI instance right before making an API call
      const ai = new GoogleGenAI({ apiKey: import.meta.env.VITE_GEMINI_API_KEY });
      const prompt = `
        我比對了兩個業務銷售 Excel 檔案。以下是差異摘要：
        - 修改項次: ${result.modified.length}
        - 新增項次: ${result.added.length}
        - 移除項次: ${result.removed.length}

        顯著變更 (前15筆):
        ${result.modified.slice(0, 15).map(m => `- [${m.id}] ${m.projectName || '月份小計列'}: ${m.changes.map(c => `${c.column} 從 ${c.oldValue} 變為 ${c.newValue}`).join(', ')}`).join('\n')}

        請以資深商業分析師的角度，用繁體中文提供一份精簡的執行摘要。重點分析這些數據變動（如數量 Qty 的異動、狀態變更）對業務的潛在意義。
      `;

      const response = await ai.models.generateContent({
        model: 'gemini-3-flash-preview',
        contents: prompt,
      });

      // The response.text is a property, not a method
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
            <div className="bg-indigo-600 p-2.5 rounded-2xl shadow-indigo-200 shadow-lg rotate-3 group-hover:rotate-0 transition-transform">
              <ArrowRightLeft className="text-white w-6 h-6" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-xl font-black text-slate-800 tracking-tight leading-none">Smart Comparer</h1>
                <span className="flex items-center gap-1 px-2 py-0.5 bg-emerald-100 text-emerald-700 text-[10px] font-black rounded-full border border-emerald-200 uppercase tracking-tighter">
                  <ShieldCheck className="w-3 h-3" />
                  v1.0 Stable
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
            {loading ? '大數據分析中...' : '開始自動對照'}
            {!loading && <div className="absolute inset-0 bg-white/10 translate-y-full group-hover:translate-y-0 transition-transform duration-300" />}
          </button>
        </div>

        {error && (
          <div className="mb-10 p-5 bg-rose-50 border-2 border-rose-100 rounded-3xl flex items-center gap-4 text-rose-700 animate-in fade-in slide-in-from-top-4 duration-300">
            <div className="bg-rose-100 p-2.5 rounded-xl">
              <AlertCircle className="w-6 h-6" />
            </div>
            <p className="font-bold text-base">{error}</p>
          </div>
        )}

        {result && (
          <div className="space-y-12 animate-in fade-in slide-in-from-bottom-8 duration-700">
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-8">
              <StatCard title="偵測到異動" value={result.modified.length} icon={<ArrowRightLeft />} color="indigo" />
              <StatCard title="全新增項次" value={result.added.length} icon={<PlusCircle />} color="emerald" />
              <StatCard title="已移除項次" value={result.removed.length} icon={<Trash2 />} color="rose" />
            </div>

            <section className="bg-white rounded-[2.5rem] border border-slate-200 overflow-hidden shadow-sm hover:shadow-md transition-shadow">
              <div className="p-8 border-b bg-slate-50/50 flex items-center justify-between">
                <div className="flex items-center gap-4">
                  <div className="p-3 bg-indigo-100 text-indigo-600 rounded-2xl">
                    <Sparkles className="w-6 h-6" />
                  </div>
                  <div>
                    <h3 className="font-black text-slate-800 text-lg">AI 智能商業報告</h3>
                    <p className="text-xs text-slate-500 font-medium tracking-wide">由 Gemini 3.0 提供深度數據洞察</p>
                  </div>
                </div>
                {!result.summary && (
                  <button
                    onClick={runAiSummary}
                    disabled={aiAnalyzing}
                    className="group px-6 py-3 bg-indigo-600 text-white rounded-xl font-bold text-sm hover:bg-indigo-700 transition-all flex items-center gap-2 shadow-lg shadow-indigo-100 active:scale-95"
                  >
                    {aiAnalyzing ? <Loader2 className="w-4 h-4 animate-spin" /> : <Sparkles className="w-4 h-4 group-hover:rotate-12 transition-transform" />}
                    產生摘要
                  </button>
                )}
              </div>
              <div className="p-10">
                {aiAnalyzing ? (
                  <div className="flex flex-col items-center py-12 gap-5">
                    <Loader2 className="w-12 h-12 text-indigo-500 animate-spin" />
                    <p className="text-slate-400 font-bold text-lg animate-pulse tracking-widest">正在解析商業價值...</p>
                  </div>
                ) : result.summary ? (
                  <div className="bg-indigo-50/30 p-10 rounded-[2rem] text-slate-700 leading-relaxed whitespace-pre-wrap border border-indigo-100/50 relative">
                    <div className="absolute top-0 left-0 w-2 h-full bg-indigo-500 rounded-l-[2rem]" />
                    <div className="prose prose-slate max-w-none font-medium italic text-lg opacity-90">
                      {result.summary}
                    </div>
                  </div>
                ) : (
                  <div className="text-center py-12">
                    <div className="inline-flex p-5 bg-slate-50 rounded-full mb-6 border border-slate-100">
                      <Sparkles className="w-10 h-10 text-slate-300" />
                    </div>
                    <p className="text-slate-400 font-bold text-lg">點擊上方按鈕，由 AI 自動解讀本次更新的關鍵變化。</p>
                  </div>
                )}
              </div>
            </section>

            <section className="bg-white rounded-[2.5rem] border border-slate-200 shadow-sm overflow-hidden">
              <div className="px-10 py-8 border-b flex items-center gap-4 bg-white">
                <Table className="text-indigo-600 w-7 h-7" />
                <h3 className="font-black text-slate-800 text-xl tracking-tight">差異比對詳細清單</h3>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-50/80 text-slate-400 text-[11px] uppercase tracking-[0.2em] font-black">
                      <th className="px-10 py-6 w-48">識別碼 / 月份</th>
                      <th className="px-10 py-6">業務資訊 (案名與客戶)</th>
                      <th className="px-10 py-6">具體異動細節</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {result.modified.length === 0 ? (
                      <tr>
                        <td colSpan={3} className="px-10 py-32 text-center">
                          <div className="flex flex-col items-center gap-6">
                            <div className="p-6 bg-emerald-50 rounded-full">
                              <CheckCircle2 className="w-16 h-16 text-emerald-400" />
                            </div>
                            <p className="text-slate-400 font-bold text-xl">目前兩個檔案完全一致，沒有任何異動。</p>
                          </div>
                        </td>
                      </tr>
                    ) : (
                      result.modified.map((diff, i) => (
                        <tr key={i} className="hover:bg-slate-50/50 transition-colors group">
                          <td className="px-10 py-10 align-top">
                            <div className="flex flex-col gap-3">
                              <span className="bg-slate-900 text-white px-4 py-2 rounded-xl font-mono text-xs font-black inline-block shadow-md group-hover:scale-105 transition-transform text-center tracking-wider">
                                {diff.id}
                              </span>
                              {!diff.projectName && (
                                <span className="flex items-center gap-1.5 text-[10px] font-black text-indigo-600 uppercase bg-indigo-50 px-3 py-1 rounded-lg self-start border border-indigo-100 shadow-sm">
                                  <Calculator className="w-3.5 h-3.5" />
                                  彙總統計列
                                </span>
                              )}
                            </div>
                          </td>
                          <td className="px-10 py-10 align-top">
                            {diff.projectName ? (
                              <div className="space-y-2">
                                <div className="font-black text-slate-800 text-lg leading-tight group-hover:text-indigo-600 transition-colors tracking-tight">{diff.projectName}</div>
                                <div className="flex items-center gap-3">
                                  <div className="w-1 h-5 bg-indigo-200 rounded-full" />
                                  <div className="text-xs text-slate-400 font-bold uppercase tracking-[0.1em]">{diff.customer || '客戶資訊未載明'}</div>
                                </div>
                              </div>
                            ) : (
                              <div className="py-4 px-6 bg-slate-50/80 rounded-2xl border border-dashed border-slate-200">
                                <span className="text-sm font-bold text-slate-400 italic flex items-center gap-3">
                                  <div className="w-2 h-2 bg-slate-200 rounded-full animate-pulse" />
                                  此筆為月份小計彙總，不包含單一案名。
                                </span>
                              </div>
                            )}
                          </td>
                          <td className="px-10 py-10 space-y-5">
                            {diff.changes.map((change, j) => (
                              <div key={j} className="flex flex-col lg:flex-row lg:items-center gap-5">
                                <div className={`
                                  text-[10px] font-black px-4 py-1.5 rounded-full uppercase tracking-[0.15em] w-28 text-center shadow-sm border
                                  ${change.column.includes('Qty') ? 'bg-amber-50 text-amber-700 border-amber-200' : 'bg-blue-50 text-blue-700 border-blue-200'}
                                `}>
                                  {change.column}
                                </div>
                                <div className="flex items-center gap-5 text-base font-bold">
                                  <span className="text-slate-300 line-through decoration-slate-200 decoration-2">{change.oldValue}</span>
                                  <ArrowRightLeft className="w-5 h-5 text-slate-400 group-hover:text-indigo-400 transition-colors" />
                                  <div className="bg-indigo-600 text-white px-5 py-2 rounded-2xl shadow-xl shadow-indigo-100 transform group-hover:scale-110 transition-transform tabular-nums">
                                    {change.newValue}
                                  </div>
                                </div>
                              </div>
                            ))}
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

const FileDropZone: React.FC<{ label: string; file: File | null; onChange: (e: React.ChangeEvent<HTMLInputElement>) => void; onClear: () => void; }> = ({ label, file, onChange, onClear }) => (
  <div className="flex flex-col gap-4">
    <label className="text-xs font-black text-slate-500 uppercase tracking-[0.25em] ml-3">{label}</label>
    <div className={`
      relative border-2 border-dashed rounded-[2.5rem] p-10 transition-all h-56 flex flex-col items-center justify-center gap-6 group overflow-hidden
      ${file ? 'border-indigo-400 bg-indigo-50/40 shadow-inner' : 'border-slate-200 bg-white hover:border-indigo-300 hover:bg-slate-50/50 shadow-sm'}
    `}>
      {file ? (
        <>
          <div className="bg-indigo-600 p-5 rounded-[1.5rem] shadow-2xl shadow-indigo-100 animate-in zoom-in spin-in-12 duration-500">
            <CheckCircle2 className="text-white w-8 h-8" />
          </div>
          <div className="text-center w-full px-8">
            <p className="font-black text-slate-800 text-base truncate tracking-tight">{file.name}</p>
            <p className="text-[10px] text-indigo-500 font-black uppercase mt-3 tracking-widest bg-indigo-100/50 px-3 py-1 rounded-full inline-block">檔案已成功就緒</p>
          </div>
          <button onClick={(e) => { e.stopPropagation(); onClear(); }} className="absolute top-6 right-6 p-3 text-slate-300 hover:text-rose-500 hover:bg-rose-50 rounded-2xl transition-all shadow-sm hover:shadow-md active:scale-90">
            <Trash2 className="w-6 h-6" />
          </button>
        </>
      ) : (
        <>
          <div className="bg-slate-50 p-6 rounded-[2rem] text-slate-300 group-hover:scale-110 group-hover:text-indigo-500 group-hover:bg-indigo-50 transition-all duration-700 border border-slate-100">
            <FileUp className="w-12 h-12" />
          </div>
          <div className="text-center">
            <p className="text-base font-black text-slate-700">點擊或拖放上傳</p>
            <p className="text-[11px] text-slate-400 font-bold mt-3 uppercase tracking-widest opacity-70">支援 .xlsx, .xls, .csv 檔案</p>
          </div>
          <input type="file" onChange={onChange} className="absolute inset-0 opacity-0 cursor-pointer" accept=".xlsx,.xls,.csv" />
        </>
      )}
    </div>
  </div>
);

const StatCard: React.FC<{ title: string; value: number; icon: React.ReactNode; color: 'indigo' | 'emerald' | 'rose' }> = ({ title, value, icon, color }) => {
  const styles = {
    indigo: 'border-slate-100 text-indigo-600',
    emerald: 'border-slate-100 text-emerald-600',
    rose: 'border-slate-100 text-rose-600'
  };
  const iconBg = {
    indigo: 'bg-indigo-50 text-indigo-600 border-indigo-100',
    emerald: 'bg-emerald-50 text-emerald-600 border-emerald-100',
    rose: 'bg-rose-50 text-rose-600 border-rose-100'
  };

  return (
    <div className={`bg-white p-10 rounded-[2.5rem] border-2 ${styles[color]} flex items-center justify-between group hover:shadow-2xl hover:border-indigo-100 transition-all duration-700`}>
      <div className="space-y-2">
        <div className="text-[11px] font-black uppercase tracking-[0.25em] text-slate-400">{title}</div>
        <div className="text-6xl font-black text-slate-900 leading-tight tabular-nums tracking-tighter">{value}</div>
      </div>
      <div className={`p-6 rounded-3xl border ${iconBg[color]} transition-all duration-1000 group-hover:rotate-[15deg] group-hover:scale-125 shadow-sm`}>
        {/* Fix for TypeScript error on line 338: Use React.isValidElement and cast to React.ReactElement<any> to satisfy property requirements of cloneElement */}
        {React.isValidElement(icon) ? React.cloneElement(icon as React.ReactElement<any>, { className: 'w-10 h-10' }) : icon}
      </div>
    </div>
  );
};

export default App;
