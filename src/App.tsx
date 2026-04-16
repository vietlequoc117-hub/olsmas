import React, { useState } from 'react';
import { 
  FileSpreadsheet, 
  Download, 
  CheckCircle2, 
  AlertCircle, 
  Terminal, 
  FileCode2, 
  Upload, 
  Loader2, 
  FileUp
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

interface SyncStats {
  updatedCount: number;
  totalStudentsInTarget: number;
  totalSourceStudents: number;
  classes: Array<{
    name: string;
    count: number;
    avg: string;
    min: number;
    max: number;
  }>;
}

export default function App() {
  const [sources, setSources] = useState<File[]>([]);
  const [target, setTarget] = useState<File | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [resultUrl, setResultUrl] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [stats, setStats] = useState<SyncStats | null>(null);

  const handleSourceChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setSources(Array.from(e.target.files).slice(0, 3));
      setError(null);
    }
  };

  const handleTargetChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setTarget(e.target.files[0]);
      setError(null);
    }
  };

  const handleSync = async () => {
    if (sources.length === 0 || !target) {
      setError('Vui lòng chọn đầy đủ file nguồn và file đích.');
      return;
    }

    setIsProcessing(true);
    setError(null);
    setResultUrl(null);
    setStats(null);

    const formData = new FormData();
    sources.forEach(file => formData.append('sources', file));
    formData.append('target', target);

    try {
      const response = await fetch('/api/sync-grades', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const data = await response.json();
        throw new Error(data.error || 'Lỗi xử lý file.');
      }

      const updatedCount = response.headers.get('X-Updated-Count');
      const statsJson = response.headers.get('X-AI-Summary');
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      
      setResultUrl(url);
      if (statsJson) setStats(JSON.parse(statsJson));
      
      if (updatedCount && parseInt(updatedCount) === 0) {
        setError('Xử lý xong nhưng không tìm thấy học sinh nào khớp. Vui lòng kiểm tra lại tên lớp và định dạng file.');
      }
    } catch (err: any) {
      setError(err.message);
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-900 pb-20">
      {/* Hero Section */}
      <header className="bg-white border-b border-slate-200">
        <div className="max-w-5xl mx-auto px-6 py-12 flex flex-col items-center text-center">
          <motion.div 
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="flex items-center justify-center gap-4 mb-6"
          >
            <div className="p-3 bg-emerald-100 rounded-2xl">
              <FileSpreadsheet className="w-8 h-8 text-emerald-600" />
            </div>
            <h1 className="text-3xl font-bold tracking-tight text-slate-900">
              ĐỒNG BỘ ONLUYEN- SMAS
            </h1>
          </motion.div>
          <motion.p 
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.1 }}
            className="text-xl text-slate-600 max-w-2xl leading-relaxed"
          >
            Công cụ trực tuyến giúp đồng bộ điểm số từ các file bài thi vào Bảng điểm tổng hợp một cách tự động và chính xác.
          </motion.p>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-12 space-y-12">
        {/* Main Tool Section */}
        <section className="bg-white rounded-3xl border border-slate-200 shadow-xl overflow-hidden">
          <div className="p-8 border-b border-slate-100 bg-slate-50/50">
            <h2 className="text-xl font-bold flex items-center gap-2">
              <FileUp className="w-5 h-5 text-blue-600" />
              Bắt đầu đồng bộ điểm
            </h2>
          </div>
          
          <div className="p-8 grid md:grid-cols-2 gap-8">
            {/* Source Files Upload */}
            <div className="space-y-4">
              <label className="block font-semibold text-slate-700">
                1. Chọn các file nguồn (Tối đa 3 file)
                <span className="block text-xs font-normal text-slate-500 mt-1">
                  File kết quả từ hệ thống thi (Khối 10, 11, 12)
                </span>
              </label>
              <div className="relative group">
                <input 
                  type="file" 
                  multiple 
                  accept=".xlsx, .xls"
                  onChange={handleSourceChange}
                  className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                />
                <div className={`p-6 border-2 border-dashed rounded-2xl transition-all flex flex-col items-center justify-center gap-3 ${sources.length > 0 ? 'border-emerald-200 bg-emerald-50' : 'border-slate-200 group-hover:border-blue-300 group-hover:bg-blue-50'}`}>
                  <Upload className={`w-8 h-8 ${sources.length > 0 ? 'text-emerald-500' : 'text-slate-400'}`} />
                  <div className="text-center">
                    <p className="text-sm font-medium">
                      {sources.length > 0 ? `Đã chọn ${sources.length} file` : 'Kéo thả hoặc click để chọn file'}
                    </p>
                    <p className="text-xs text-slate-400 mt-1">Định dạng .xlsx</p>
                  </div>
                </div>
              </div>
              {sources.length > 0 && (
                <ul className="text-xs text-slate-500 space-y-1 pl-2">
                  {sources.map((f) => (
                    <li key={`${f.name}-${f.lastModified}`} className="flex items-center gap-2">
                      <CheckCircle2 className="w-3 h-3 text-emerald-500" /> {f.name}
                    </li>
                  ))}
                </ul>
              )}
            </div>

            {/* Target File Upload */}
            <div className="space-y-4">
              <label className="block font-semibold text-slate-700">
                2. Chọn file đích (Bảng điểm tổng hợp)
                <span className="block text-xs font-normal text-slate-500 mt-1">
                  File cần điền điểm vào cột ĐĐG GK
                </span>
              </label>
              <div className="relative group">
                <input 
                  type="file" 
                  accept=".xlsx, .xls"
                  onChange={handleTargetChange}
                  className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                />
                <div className={`p-6 border-2 border-dashed rounded-2xl transition-all flex flex-col items-center justify-center gap-3 ${target ? 'border-blue-200 bg-blue-50' : 'border-slate-200 group-hover:border-blue-300 group-hover:bg-blue-50'}`}>
                  <FileSpreadsheet className={`w-8 h-8 ${target ? 'text-blue-500' : 'text-slate-400'}`} />
                  <div className="text-center">
                    <p className="text-sm font-medium">
                      {target ? target.name : 'Kéo thả hoặc click để chọn file'}
                    </p>
                    <p className="text-xs text-slate-400 mt-1">Định dạng .xlsx</p>
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div className="p-8 bg-slate-50 border-t border-slate-100 flex flex-col items-center gap-6">
            <button 
              onClick={handleSync}
              disabled={isProcessing || sources.length === 0 || !target}
              className={`px-10 py-4 rounded-2xl font-bold text-white shadow-lg transition-all flex items-center gap-3 ${isProcessing || sources.length === 0 || !target ? 'bg-slate-300 cursor-not-allowed' : 'bg-emerald-600 hover:bg-emerald-700 hover:scale-105 active:scale-95'}`}
            >
              {isProcessing ? (
                <>
                  <Loader2 className="w-5 h-5 animate-spin" />
                  Đang xử lý...
                </>
              ) : (
                <>
                  <CheckCircle2 className="w-5 h-5" />
                  Bắt đầu đồng bộ ngay
                </>
              )}
            </button>

            <AnimatePresence mode="wait">
              {error && (
                <motion.div 
                  key="error-message"
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0 }}
                  className="flex items-center gap-2 text-red-600 bg-red-50 px-4 py-2 rounded-lg border border-red-100"
                >
                  <AlertCircle className="w-4 h-4" />
                  <span className="text-sm font-medium">{error}</span>
                </motion.div>
              )}

              {resultUrl && (
                <motion.div 
                  key="success-result"
                  initial={{ opacity: 0, scale: 0.9 }}
                  animate={{ opacity: 1, scale: 1 }}
                  className="flex flex-col items-center gap-6 p-8 bg-emerald-50 border border-emerald-100 rounded-3xl w-full max-w-2xl"
                >
                  <div className="flex flex-col items-center gap-4">
                    <div className="p-3 bg-white rounded-full shadow-sm">
                      <CheckCircle2 className="w-8 h-8 text-emerald-500" />
                    </div>
                    <div className="text-center">
                      <h3 className="font-bold text-emerald-900 text-lg">Đồng bộ thành công!</h3>
                      <p className="text-sm text-emerald-700 mt-1">
                        Đã cập nhật điểm cho <strong>{stats?.updatedCount}</strong> học sinh.
                      </p>
                    </div>
                  </div>

                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 w-full">
                    <div className="bg-white p-3 rounded-xl border border-emerald-100 text-center">
                      <p className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">Tổng khớp</p>
                      <p className="text-xl font-bold text-emerald-600">{stats?.updatedCount}</p>
                    </div>
                    <div className="bg-white p-3 rounded-xl border border-emerald-100 text-center">
                      <p className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">Nguồn</p>
                      <p className="text-xl font-bold text-blue-600">{stats?.totalSourceStudents}</p>
                    </div>
                    <div className="bg-white p-3 rounded-xl border border-emerald-100 text-center">
                      <p className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">Đích</p>
                      <p className="text-xl font-bold text-purple-600">{stats?.totalStudentsInTarget}</p>
                    </div>
                    <div className="bg-white p-3 rounded-xl border border-emerald-100 text-center">
                      <p className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">Lớp</p>
                      <p className="text-xl font-bold text-amber-600">{stats?.classes.length}</p>
                    </div>
                  </div>

                  <div className="flex flex-wrap justify-center gap-4">
                    <a 
                      href={resultUrl} 
                      download="BangDiem_KetQua_DongBo.xlsx"
                      className="flex items-center gap-2 px-8 py-4 bg-emerald-600 text-white rounded-2xl font-bold hover:bg-emerald-700 transition-all shadow-lg hover:scale-105 active:scale-95"
                    >
                      <Download className="w-5 h-5" />
                      Tải file kết quả
                    </a>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </div>
        </section>

        <footer className="text-center text-slate-400 text-sm pt-8">
          <p>© 2024 Excel Grade Sync Tool - Built for Teachers</p>
        </footer>
      </main>
    </div>
  );
}
