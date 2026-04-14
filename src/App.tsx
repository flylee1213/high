import React, { useState, useEffect } from 'react';
import { 
  FileSpreadsheet, 
  FileText, 
  Upload, 
  Download, 
  CheckCircle2, 
  AlertCircle, 
  Loader2, 
  X,
  FileUp,
  ArrowRightLeft,
  ShieldCheck,
  History,
  Info,
  ChevronDown,
  ChevronUp
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from './lib/utils';
import { processFilesLocally, MAPPING_CONFIG } from './lib/processor';

interface FileWithStatus {
  file: File;
  id: string;
}

export default function App() {
  const [dataSource, setDataSource] = useState<File | null>(null);
  const [templates, setTemplates] = useState<FileWithStatus[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);
  const [mapping, setMapping] = useState<Record<string, string>>(MAPPING_CONFIG);
  const [showMapping, setShowMapping] = useState(false);

  useEffect(() => {
    // Mapping is now loaded directly from processor.ts
  }, []);

  const handleDataSourceChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        setDataSource(file);
        setError(null);
      } else {
        setError('请上传有效的 Excel 数据源 (.xlsx, .xls)');
      }
    }
  };

  const handleTemplatesChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files) {
      const fileList = Array.from(files) as File[];
      const validFiles = fileList.filter(f => f.name.endsWith('.docx') || f.name.endsWith('.xlsx'));
      
      if (validFiles.length < fileList.length) {
        setError('部分文件已跳过。模板仅支持 .docx 和 .xlsx 格式。');
      }

      const newTemplates: FileWithStatus[] = validFiles.map(file => ({
        file,
        id: Math.random().toString(36).substring(7)
      }));
      
      setTemplates(prev => [...prev, ...newTemplates].slice(0, 4));
      if (validFiles.length > 0) setError(null);
    }
  };

  const removeTemplate = (id: string) => {
    setTemplates(prev => prev.filter(t => t.id !== id));
  };

  const processDocuments = async () => {
    if (!dataSource || templates.length === 0) {
      setError('请提供数据源和至少一个合同模板。');
      return;
    }

    setIsProcessing(true);
    setError(null);
    setSuccess(false);

    try {
      const blob = await processFilesLocally({
        dataSource,
        templates: templates.map(t => t.file)
      });

      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `contracts_${Date.now()}.zip`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      setSuccess(true);
    } catch (err: any) {
      console.error('Processing error:', err);
      setError(err.message || '合同生成失败');
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-900">
      {/* Navigation Rail / Header */}
      <nav className="bg-white border-b border-slate-200 px-8 py-4 flex items-center justify-between sticky top-0 z-50">
        <div className="flex items-center gap-3">
          <div className="bg-blue-600 p-2 rounded-lg shadow-blue-100 shadow-lg">
            <ShieldCheck className="w-6 h-6 text-white" />
          </div>
          <div>
            <h1 className="text-xl font-bold tracking-tight text-slate-800">ContractFlow AI</h1>
            <p className="text-[10px] text-slate-400 font-mono uppercase tracking-widest">Enterprise Automation</p>
          </div>
        </div>
        <div className="flex items-center gap-6 text-sm font-medium text-slate-500">
          <button className="hover:text-blue-600 transition-colors flex items-center gap-2">
            <History className="w-4 h-4" /> 生成记录
          </button>
          <button className="hover:text-blue-600 transition-colors flex items-center gap-2">
            <Info className="w-4 h-4" /> 使用帮助
          </button>
        </div>
      </nav>

      <main className="max-w-6xl mx-auto py-12 px-8">
        <div className="grid grid-cols-12 gap-8">
          {/* Left: Data Source */}
          <div className="col-span-12 lg:col-span-4 space-y-6">
            <section className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm">
              <div className="flex items-center gap-3 mb-6">
                <div className="p-2 bg-emerald-50 rounded-lg">
                  <FileSpreadsheet className="w-5 h-5 text-emerald-600" />
                </div>
                <h2 className="font-semibold text-slate-800">合同数据源</h2>
              </div>
              
              <label className={cn(
                "relative group cursor-pointer flex flex-col items-center justify-center border-2 border-dashed rounded-xl p-8 transition-all",
                dataSource ? "border-emerald-200 bg-emerald-50/30" : "border-slate-200 hover:border-blue-400 hover:bg-blue-50/30"
              )}>
                <input type="file" className="hidden" accept=".xlsx,.xls" onChange={handleDataSourceChange} />
                {dataSource ? (
                  <div className="text-center">
                    <CheckCircle2 className="w-10 h-10 text-emerald-500 mx-auto mb-3" />
                    <p className="text-sm font-medium text-slate-900 truncate max-w-[180px]">{dataSource.name}</p>
                    <button onClick={(e) => { e.preventDefault(); setDataSource(null); }} className="mt-2 text-xs text-red-500 hover:underline">移除</button>
                  </div>
                ) : (
                  <div className="text-center">
                    <Upload className="w-10 h-10 text-slate-300 mx-auto mb-3 group-hover:text-blue-500 transition-colors" />
                    <p className="text-sm font-medium text-slate-600">选择 Excel 数据源</p>
                    <p className="text-[10px] text-slate-400 mt-1 uppercase tracking-tighter">Support .xlsx, .xls</p>
                  </div>
                )}
              </label>
            </section>

            <section className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden transition-all">
              <button 
                onClick={() => setShowMapping(!showMapping)}
                className="w-full flex items-center justify-between p-6 hover:bg-slate-50 transition-colors"
              >
                <div className="flex items-center gap-3">
                  <div className="p-2 bg-blue-50 rounded-lg">
                    <ArrowRightLeft className="w-5 h-5 text-blue-600" />
                  </div>
                  <h2 className="font-semibold text-slate-800">映射规则预览</h2>
                  {Object.keys(mapping).length > 0 && (
                    <span className="px-2 py-0.5 bg-blue-100 text-blue-600 text-[10px] font-bold rounded-full">
                      {Object.keys(mapping).length}
                    </span>
                  )}
                </div>
                {showMapping ? (
                  <ChevronUp className="w-5 h-5 text-slate-400" />
                ) : (
                  <ChevronDown className="w-5 h-5 text-slate-400" />
                )}
              </button>
              
              <AnimatePresence>
                {showMapping && (
                  <motion.div 
                    initial={{ height: 0, opacity: 0 }}
                    animate={{ height: 'auto', opacity: 1 }}
                    exit={{ height: 0, opacity: 0 }}
                    className="px-6 pb-6"
                  >
                    <div className="space-y-2 pt-2 border-t border-slate-50">
                      {Object.entries(mapping).map(([col, keyword]) => (
                        <div key={col} className="flex items-center justify-between p-2 bg-slate-50 rounded-lg text-xs border border-slate-100">
                          <span className="text-slate-500 font-medium">{col}</span>
                          <div className="h-px flex-1 mx-3 bg-slate-200 border-t border-dashed" />
                          <span className="text-blue-600 font-mono font-bold">{keyword}</span>
                        </div>
                      ))}
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>
            </section>
          </div>

          {/* Right: Templates & Action */}
          <div className="col-span-12 lg:col-span-8 space-y-6">
            <section className="bg-white rounded-2xl p-8 border border-slate-200 shadow-sm">
              <div className="flex items-center justify-between mb-8">
                <div className="flex items-center gap-3">
                  <div className="p-2 bg-blue-50 rounded-lg">
                    <FileText className="w-5 h-5 text-blue-600" />
                  </div>
                  <h2 className="font-semibold text-slate-800 text-lg">文档模板 (Word/Excel)</h2>
                </div>
                <span className="text-xs text-slate-400 font-medium">最多支持 4 个模板</span>
              </div>

              <div className="grid grid-cols-2 gap-4 mb-6">
                <label className="col-span-2 md:col-span-1 relative group cursor-pointer flex flex-col items-center justify-center border-2 border-dashed border-slate-200 rounded-xl p-10 hover:border-blue-400 hover:bg-blue-50/30 transition-all">
                  <input type="file" className="hidden" multiple accept=".docx,.xlsx" onChange={handleTemplatesChange} />
                  <FileUp className="w-10 h-10 text-slate-300 mx-auto mb-3 group-hover:text-blue-500 transition-colors" />
                  <p className="text-sm font-medium text-slate-600">添加 Word/Excel 模板</p>
                  <p className="text-[10px] text-slate-400 mt-1 uppercase">Support .docx, .xlsx</p>
                </label>

                <div className="col-span-2 md:col-span-1 space-y-3 max-h-[220px] overflow-y-auto pr-2 custom-scrollbar">
                  <AnimatePresence initial={false}>
                    {templates.map((t) => (
                      <motion.div 
                        key={t.id}
                        initial={{ opacity: 0, x: 20 }}
                        animate={{ opacity: 1, x: 0 }}
                        exit={{ opacity: 0, scale: 0.95 }}
                        className="flex items-center justify-between p-4 bg-slate-50 rounded-xl border border-slate-100 group hover:border-blue-200 transition-colors"
                      >
                        <div className="flex items-center gap-3 overflow-hidden">
                          {t.file.name.endsWith('.docx') ? (
                            <FileText className="w-5 h-5 text-blue-500 flex-shrink-0" />
                          ) : (
                            <FileSpreadsheet className="w-5 h-5 text-emerald-500 flex-shrink-0" />
                          )}
                          <span className="text-sm font-medium text-slate-700 truncate">{t.file.name}</span>
                        </div>
                        <button onClick={() => removeTemplate(t.id)} className="p-1 text-slate-300 hover:text-red-500 transition-colors">
                          <X className="w-5 h-5" />
                        </button>
                      </motion.div>
                    ))}
                    {templates.length === 0 && (
                      <div className="h-full flex flex-col items-center justify-center text-slate-300 py-12">
                        <FileText className="w-12 h-12 opacity-20 mb-2" />
                        <p className="text-xs font-medium">暂未添加模板</p>
                      </div>
                    )}
                  </AnimatePresence>
                </div>
              </div>

              <div className="pt-8 border-t border-slate-100 flex flex-col items-center">
                <AnimatePresence mode="wait">
                  {error && (
                    <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} className="mb-6 p-4 bg-red-50 border border-red-100 rounded-xl flex items-center gap-3 text-red-700 w-full">
                      <AlertCircle className="w-5 h-5 flex-shrink-0" />
                      <p className="text-sm font-medium">{error}</p>
                    </motion.div>
                  )}
                  {success && (
                    <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} className="mb-6 p-4 bg-emerald-50 border border-emerald-100 rounded-xl flex items-center gap-3 text-emerald-700 w-full">
                      <CheckCircle2 className="w-5 h-5 flex-shrink-0" />
                      <p className="text-sm font-medium">合同生成成功！压缩包已开始下载。</p>
                    </motion.div>
                  )}
                </AnimatePresence>

                <button
                  onClick={processDocuments}
                  disabled={isProcessing || !dataSource || templates.length === 0}
                  className={cn(
                    "w-full md:w-auto px-16 py-4 rounded-xl font-bold text-lg transition-all shadow-xl flex items-center justify-center gap-3",
                    isProcessing 
                      ? "bg-slate-100 text-slate-400 cursor-not-allowed" 
                      : "bg-blue-600 text-white hover:bg-blue-700 hover:shadow-blue-200 active:scale-[0.98] disabled:opacity-50"
                  )}
                >
                  {isProcessing ? (
                    <>
                      <Loader2 className="w-6 h-6 animate-spin" />
                      正在生成合同...
                    </>
                  ) : (
                    <>
                      <Download className="w-6 h-6" />
                      批量生成文档
                    </>
                  )}
                </button>
                <p className="mt-4 text-xs text-slate-400 flex items-center gap-2 text-center">
                  <Info className="w-3 h-3" /> 提示：Excel 模板中的占位符请使用 [列名] 格式
                </p>
              </div>
            </section>
          </div>
        </div>
      </main>

      <style>{`
        .custom-scrollbar::-webkit-scrollbar { width: 4px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: transparent; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #e2e8f0; border-radius: 10px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #cbd5e1; }
      `}</style>
    </div>
  );
}
