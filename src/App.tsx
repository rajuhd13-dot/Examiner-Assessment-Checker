/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, KeyboardEvent } from 'react';
import * as XLSX from 'xlsx';
import { 
  Search, Download, Plus, RefreshCw, 
  XCircle, CheckCircle2, FileSpreadsheet, X, LayoutGrid,
  MessageSquare, Send, Phone
} from 'lucide-react';

/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

declare global {
  const google: any;
}

// --- Types ---
type SubjectStats = {
  score: string;
  allowed: boolean;
};

type Examiner = {
  sl: string;
  name: string;
  tpin: string;
  inst: string;
  dept: string;
  batch: string;
  rm: string;
  mobile: string;
  alternate: string;
  nagad: string;
  campus: string;
  training: string;
  trainingDate: string;
  remarkedBy: string;
  hscGpa: string;
  homeDistrict: string;
  email: string;
  hscBoard: string;
  subjectsChoice: string;
  runningProgram: string;
  
  english: SubjectStats;
  bangla: SubjectStats;
  physics: SubjectStats;
  chemistry: SubjectStats;
  math: SubjectStats;
  biology: SubjectStats;
  ict: SubjectStats;
  
  remark: {
    count: number;
    show: boolean;
    body: string;
    byLine: string;
    dateLine: string;
  };
};

type RowStatus = 'idle' | 'loading' | 'found' | 'not-found';

type RowData = {
  id: string;
  inputValue: string;
  status: RowStatus;
  data: Examiner | null;
};

type ColDef = {
  key: string;
  label: string;
  defaultVisible: boolean;
  isScore?: boolean;
};

// --- Constants ---
const THRESHOLDS = { english: 55, bangla: 48, physics: 48, chemistry: 48, math: 48, biology: 48, ict: 48 };

const ALL_COLS: ColDef[] = [
  { key: 'sl', label: 'SL', defaultVisible: true },
  { key: 'name', label: 'Nick Name', defaultVisible: true },
  { key: 'tpin', label: 'T-Pin', defaultVisible: true },
  { key: 'inst', label: 'Institute', defaultVisible: true },
  { key: 'dept', label: 'Department', defaultVisible: true },
  { key: 'batch', label: 'HSC Batch', defaultVisible: true },
  { key: 'rm', label: 'RM', defaultVisible: true },
  { key: 'mobile', label: 'Mobile 1', defaultVisible: true },
  { key: 'alternate', label: 'Alternate', defaultVisible: true },
  { key: 'nagad', label: 'Nagad Number', defaultVisible: true },
  { key: 'english', label: 'English(%)', defaultVisible: true, isScore: true },
  { key: 'bangla', label: 'Bangla(%)', defaultVisible: true, isScore: true },
  { key: 'physics', label: 'Physics(%)', defaultVisible: true, isScore: true },
  { key: 'chemistry', label: 'Chemistry(%)', defaultVisible: true, isScore: true },
  { key: 'math', label: 'Math(%)', defaultVisible: true, isScore: true },
  { key: 'biology', label: 'Biology(%)', defaultVisible: true, isScore: true },
  { key: 'ict', label: 'ICT(%)', defaultVisible: true, isScore: true },
  { key: 'training', label: 'Training Report', defaultVisible: true },
  { key: 'trainingDate', label: 'Training Date', defaultVisible: true },
  { key: 'campus', label: 'Campus', defaultVisible: true },
];

// --- Backend Integration ---
const DEFAULT_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyesBP9GFM2tcfDdr_eBUGhUA-lLxF-9jSUHN4xpngdKyLb2vaeRJ9SbmxgiY5Zg-0-jg/exec';

const getScriptUrl = () => {
  const url = localStorage.getItem('examiner_script_url') || DEFAULT_SCRIPT_URL;
  return url.trim();
};

// Help fetch bypass CORS / sandbox / multi-login Google redirects in sandboxed iframes
const fetchWithFallback = async (targetUrl: string): Promise<Response> => {
  try {
    const res = await fetch(targetUrl);
    if (res.ok) return res;
    throw new Error(`HTTP ${res.status}`);
  } catch (err: any) {
    const errMsg = (err?.message || String(err)).toLowerCase();
    const isCorsOrNetwork = 
      errMsg.includes("fetch") || 
      err?.name === "TypeError" || 
      errMsg.includes("http 0") || 
      errMsg.includes("network") ||
      errMsg.includes("failed to") ||
      errMsg.includes("cors");
      
    if (isCorsOrNetwork) {
      console.warn("Direct fetch blocked or failed. Retrying through CORS Proxy Bridge 1...");
      try {
        const proxyUrl = `https://api.allorigins.win/raw?url=${encodeURIComponent(targetUrl)}`;
        const res = await fetch(proxyUrl);
        if (res.ok) return res;
      } catch (proxyErr) {
        console.warn("CORS Proxy 1 failed, trying CORS Proxy Bridge 2...", proxyErr);
      }

      try {
        const proxyUrl2 = `https://corsproxy.io/?${encodeURIComponent(targetUrl)}`;
        const res = await fetch(proxyUrl2);
        if (res.ok) return res;
      } catch (proxyErr2) {
        console.warn("CORS Proxy 2 failed. Request cannot be proxy bridged.");
      }
    }
    throw err;
  }
};

const lookupExaminer = async (query: string): Promise<{ success: boolean; found: boolean; data?: Examiner; error?: string }> => {
  const scriptBaseUrl = getScriptUrl();
  
  // If running inside Google Apps Script directly
  if (typeof google !== 'undefined' && google.script && google.script.run) {
    return new Promise((resolve) => {
      google.script.run
        .withSuccessHandler((res: any) => resolve(res))
        .withFailureHandler((err: any) => resolve({ success: false, found: false, error: err.message || String(err) }))
        .lookupExaminerByTPin(query);
    });
  }

  // If running externally (React Dev Server or API call)
  try {
    const separator = scriptBaseUrl.includes('?') ? '&' : '?';
    const url = `${scriptBaseUrl}${separator}action=lookup&query=${encodeURIComponent(query)}&_t=${Date.now()}`;
    const response = await fetchWithFallback(url);

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const result = await response.json();
    return result;
  } catch (error: any) {
    console.error("API Fetch Error:", error);
    let errorMessage = error.message || "Failed to connect to API";
    
    if (errorMessage === "Failed to fetch" || error.name === "TypeError") {
      errorMessage = "CORS error or network failure. Verify Google Apps Script Deployment: Web App > Anyone > Me.";
    }
    
    return { success: false, found: false, error: errorMessage };
  }
};

export default function App() {
  const [rows, setRows] = useState<RowData[]>([]);
  const [globalSearch, setGlobalSearch] = useState('');
  const [isColModalOpen, setIsColModalOpen] = useState(false);
  const [isErrorModalOpen, setIsErrorModalOpen] = useState(false);
  const [connectionStatus, setConnectionStatus] = useState<'connecting' | 'connected' | 'error'>('connecting');
  const [customUrl, setCustomUrl] = useState(getScriptUrl());
  const [lastError, setLastError] = useState<string | null>(null);
  const [localData, setLocalData] = useState<any[][]>(() => {
    try {
      const cached = localStorage.getItem('examiner_local_data');
      return cached ? JSON.parse(cached) : [];
    } catch { return []; }
  });
  const [thresholds, setThresholds] = useState(() => {
    try {
      const cached = localStorage.getItem('examiner_thresholds');
      return cached ? JSON.parse(cached) : THRESHOLDS;
    } catch { return THRESHOLDS; }
  });
  const [dbStats, setDbStats] = useState<{ rowCount: number | null }>(() => {
    try {
      const cached = localStorage.getItem('examiner_db_stats');
      return cached ? JSON.parse(cached) : { rowCount: null };
    } catch { return { rowCount: null }; }
  });
  const [isSyncing, setIsSyncing] = useState(false);
  const [addRowCount, setAddRowCount] = useState('100');
  const [lastSyncTime, setLastSyncTime] = useState<number>(() => {
    return Number(localStorage.getItem('examiner_last_sync') || 0);
  });

  // Helper to normalize strings for search (Adopted from user snippet)
  const norm = (v: any) => {
    let s = String(v || '').trim();
    if (!s) return '';
    // If it looks like a phone number, normalize phone formatting
    if (/^[\d+]+$/.test(s.replace(/[-\s]/g, '')) && s.length >= 10) {
      let d = s.replace(/\D/g, '');
      if (d) {
        if (d.length >= 12 && d.slice(0, 3) === '880') return d;
        if (d[0] === '0' && d.length === 11) return '88' + d;
        if (d[0] === '1' && d.length === 10) return '880' + d;
        return d;
      }
    }
    // For T-Pins or other text, just uppercase and trim
    return s.toUpperCase();
  };

  const padMobile = (v: any) => {
    const s = String(v || '').trim().replace(/\.0$/, '');
    if (/^\d{10}$/.test(s)) return '0' + s;
    return s;
  };

  const parseScore = (v: any, threshold: number): SubjectStats => {
    const raw = String(v || '').trim();
    if (!raw || raw === '—') return { score: '', allowed: false };
    
    // Handle multiple scores like "53/67" or "50/40"
    const scores = raw.split('/').map(s => parseFloat(s.trim()));
    const isAllowed = scores.some(s => Number.isFinite(s) && s >= threshold);
    
    return {
      score: raw,
      allowed: isAllowed
    };
  };

  const [visibleCols, setVisibleCols] = useState<Record<string, boolean>>(() => {
    const initial: Record<string, boolean> = {};
    ALL_COLS.forEach(c => initial[c.key] = c.defaultVisible);
    return initial;
  });

  const [selectedExaminer, setSelectedExaminer] = useState<Examiner | null>(null);

  // Helper to re-run connection check and SYNC
  const checkConnection = async (targetUrl?: string, forceManual?: boolean) => {
    const scriptBaseUrl = targetUrl || getScriptUrl();
    
    // FAST CONNECT: If we have cached data, we are "connected" instantly (0-1 seconds)
    if (localData && localData.length > 0) {
      setConnectionStatus('connected');
    } else {
      setConnectionStatus('connecting');
    }

    setLastError(null);
    setIsSyncing(true);
    
    try {
      const separator = scriptBaseUrl.includes('?') ? '&' : '?';
      
      // Step 1: Try action=sync first (modern server sync)
      const syncUrl = `${scriptBaseUrl}${separator}action=sync&_t=${Date.now()}`;
      let data: any = null;
      let syncAttemptSuccess = false;

      try {
        const res = await fetchWithFallback(syncUrl);
        
        if (res.ok) {
          const tempJson = await res.json();
          const isTempSuccess = tempJson && (tempJson.success || tempJson.ok);
          if (isTempSuccess) {
            data = tempJson;
            syncAttemptSuccess = true;
          } else {
            console.warn("action=sync returned inactive or error response:", tempJson);
          }
        }
      } catch (errSync) {
        console.warn("action=sync fetch failed, falling back to legacy...", errSync);
      }

      // Step 2: Fallback to action=filterOptions if action=sync was not successful or failed
      if (!syncAttemptSuccess || !data) {
        const fallbackUrl = `${scriptBaseUrl}${separator}action=filterOptions&_t=${Date.now()}`;
        const fallbackRes = await fetchWithFallback(fallbackUrl);
        
        if (!fallbackRes.ok) {
          throw new Error(`HTTP ${fallbackRes.status}. Could not sync with spreadsheet.`);
        }
        
        const fallbackJson = await fallbackRes.json();
        const isFallbackSuccess = fallbackJson && (fallbackJson.success || fallbackJson.ok);
        if (!isFallbackSuccess) {
          throw new Error(fallbackJson?.error || fallbackJson?.message || "Legacy filterOptions returned failure");
        }
        data = fallbackJson;
      }
      
      const isSuccess = data.success || data.ok;
      
      if (isSuccess) {
        setConnectionStatus('connected');
        if (targetUrl) localStorage.setItem('examiner_script_url', targetUrl);

        const now = Date.now();
        setLastSyncTime(now);
        localStorage.setItem('examiner_last_sync', String(now));

        // Rows can be returned as data.data (new sync) or data.rows (older sync)
        const fetchedRows = data.data || data.rows;
        if (fetchedRows && Array.isArray(fetchedRows) && fetchedRows.length > 0) {
          setLocalData(fetchedRows);
          try {
            localStorage.setItem('examiner_local_data', JSON.stringify(fetchedRows));
          } catch (storageErr) {
            console.warn("Storage quota exceeded, keeping in browser memory. Clear other domain cache.");
          }
          
          const recordCount = fetchedRows.length;
          const stats = { rowCount: recordCount };
          setDbStats(stats);
          localStorage.setItem('examiner_db_stats', JSON.stringify(stats));
        } else if (data.rowCount) {
          const recordCount = data.rowCount;
          const stats = { rowCount: recordCount };
          setDbStats(stats);
          localStorage.setItem('examiner_db_stats', JSON.stringify(stats));
        }
        
        if (data.allow) {
          const normalized: any = {};
          Object.keys(data.allow).forEach(k => {
            normalized[String(k).toLowerCase()] = data.allow[k];
          });
          setThresholds(normalized);
          localStorage.setItem('examiner_thresholds', JSON.stringify(normalized));
        }

        if (forceManual) {
          const count = fetchedRows ? fetchedRows.length : (data.rowCount || 0);
          alert(`Successfully synchronized ${count.toLocaleString()} examiners! Saved to browser cache.`);
        }
      } else {
        throw new Error(data.error || data.message || "Script reported failure during sync");
      }

    } catch (e: any) {
      console.error("Connectivity check and Sync failed:", e);
      let errMsg = e.message || String(e);
      if (errMsg === "Failed to fetch" || e.name === "TypeError") {
        errMsg = "Network Error (CORS or blocked request). Verify your script is deployed as Web App for 'Anyone'.";
      }

      if (!localData || localData.length === 0) {
        setLastError(errMsg);
        setConnectionStatus('error');
        setIsErrorModalOpen(true);
      } else {
        console.warn("Silent sync failed, utilizing offline cache:", errMsg);
        if (forceManual) {
          alert(`Could not sync with Google Sheets:\n${errMsg}\n\nUsing offline client cache with ${localData.length} records.`);
        }
      }
    } finally {
      setIsSyncing(false);
    }
  };

  const clearLocalCache = () => {
    if (window.confirm("Are you sure you want to clear your local database cache? You will need an active internet connection to download data again.")) {
      localStorage.removeItem('examiner_local_data');
      localStorage.removeItem('examiner_db_stats');
      localStorage.removeItem('examiner_last_sync');
      setLocalData([]);
      setDbStats({ rowCount: null });
      setLastSyncTime(0);
      setConnectionStatus('connecting');
      alert("Cache cleared successfully. Please click sync to download data again.");
    }
  };

  // Init
  useEffect(() => {
    const initialRows = Array.from({ length: 9 }).map(() => ({
      id: crypto.randomUUID(),
      inputValue: '',
      status: 'idle' as RowStatus,
      data: null
    }));
    setRows(initialRows);
    checkConnection();
  }, []);

  const resetToDefault = () => {
    localStorage.removeItem('examiner_script_url');
    setCustomUrl(DEFAULT_SCRIPT_URL);
    checkConnection(DEFAULT_SCRIPT_URL);
  };

  const handlePaste = (e: React.ClipboardEvent, startRowId: string) => {
    const pastedText = e.clipboardData.getData('text');
    const lines = pastedText.split(/[\n\r,]+/).map(l => l.trim()).filter(Boolean);
    
    if (lines.length <= 1) return;
    
    e.preventDefault();
    const startIndex = rows.findIndex(r => r.id === startRowId);
    if (startIndex === -1) return;

    // Fill current rows
    const updatedRows = [...rows];
    lines.forEach((line, i) => {
      const targetIdx = startIndex + i;
      if (targetIdx < updatedRows.length) {
        updatedRows[targetIdx] = {
          ...updatedRows[targetIdx],
          inputValue: line,
          status: 'idle',
          data: null
        };
        // Trigger search immediately
        executeSearch(updatedRows[targetIdx].id, line);
      }
    });
    setRows(updatedRows);
  };

  const handleRowInput = (id: string, val: string) => {
    setRows(prev => prev.map(r => r.id === id ? { ...r, inputValue: val, status: 'idle', data: null } : r));
  };

  const executeSearch = async (id: string, query: string) => {
    if (!query.trim()) return;
    setRows(prev => prev.map(r => r.id === id ? { ...r, status: 'loading', data: null, inputValue: query } : r));
    
    try {
      // 1. Local Search (Instant)
      if (localData.length > 0) {
        const q = norm(query);
        // Correct Mapping indices: 3: T-PIN, 9: Mobile, 10: Alternate
        const rowData = localData.find(r => norm(r[3]) === q || norm(r[9]) === q || norm(r[10]) === q);

        if (rowData) {
          const r = rowData;
          const curThresholds = thresholds || THRESHOLDS;

          const extractNum = (v: any) => {
            const m = String(v || '').match(/\d+/);
            return m ? Number(m[0]) : 0;
          };

          const parseRemark = (raw: string, rmNum: number) => {
            const show = (rmNum >= 4) || (raw.length > 0 && rmNum > 0);
            if (!show || !raw) return { count: rmNum, show: false, body: '', byLine: '', dateLine: '' };
            return { count: rmNum, show: true, body: raw, byLine: '', dateLine: '' };
          };

          const examiner: Examiner = {
            sl: String(r[0] || '').trim(),
            name: String(r[1] || '').trim(), // Usually Nick Name/Display Name
            tpin: String(r[3] || '').trim(),
            inst: String(r[4] || '').trim(),
            dept: String(r[5] || '').trim(),
            batch: String(r[6] || '').trim(),
            rm: String(r[7] || '').trim(),
            mobile: padMobile(r[9]),
            alternate: padMobile(r[10]),
            nagad: padMobile(r[11]),
            hscGpa: String(r[31] || '').trim(),
            homeDistrict: String(r[60] || '').trim(),
            email: String(r[22] || '').trim(),
            hscBoard: String(r[30] || '').trim(),
            subjectsChoice: [r[34], r[35], r[36], r[37], r[38]].filter(Boolean).join(', '),
            runningProgram: String(r[16] || '').trim(),
            english: parseScore(r[61], curThresholds.english || THRESHOLDS.english),
            bangla: parseScore(r[64], curThresholds.bangla || THRESHOLDS.bangla),
            physics: parseScore(r[67], curThresholds.physics || THRESHOLDS.physics), 
            chemistry: parseScore(r[70], curThresholds.chemistry || THRESHOLDS.chemistry),
            math: parseScore(r[73], curThresholds.math || THRESHOLDS.math),
            biology: parseScore(r[76], curThresholds.biology || THRESHOLDS.biology),
            ict: parseScore(r[79], curThresholds.ict || THRESHOLDS.ict),
            training: String(r[82] || '').trim(),
            trainingDate: String(r[83] || '').trim(),
            campus: String(r[88] || '').trim(),
            remarkedBy: String(r[8] || '').trim(),
            remark: parseRemark(String(r[92] || ''), extractNum(r[7]))
          };

          setRows(prev => prev.map(row => row.id === id ? { ...row, status: 'found', data: examiner } : row));
          return;
        }
      }

      // 2. API Fallback
      const res = await lookupExaminer(query);
      if (res.success && res.found && res.data) {
        // Map trainDate from API to trainingDate (ensure compatibility)
        const mappedData = {
          ...res.data,
          trainingDate: (res.data as any).trainDate || (res.data as any).trainingDate || ''
        } as Examiner;

        // Re-apply score thresholds logic in frontend to handle multi-scores like "53/67"
        const curThresholds = thresholds || THRESHOLDS;
        const getS = (obj: any) => (typeof obj === 'object' && obj !== null) ? obj.score : obj;
        
        mappedData.english = parseScore(getS(mappedData.english), curThresholds.english || THRESHOLDS.english);
        mappedData.bangla = parseScore(getS(mappedData.bangla), curThresholds.bangla || THRESHOLDS.bangla);
        mappedData.physics = parseScore(getS(mappedData.physics), curThresholds.physics || THRESHOLDS.physics);
        mappedData.chemistry = parseScore(getS(mappedData.chemistry), curThresholds.chemistry || THRESHOLDS.chemistry);
        mappedData.math = parseScore(getS(mappedData.math), curThresholds.math || THRESHOLDS.math);
        mappedData.biology = parseScore(getS(mappedData.biology), curThresholds.biology || THRESHOLDS.biology);
        mappedData.ict = parseScore(getS(mappedData.ict), curThresholds.ict || THRESHOLDS.ict);
        
        setRows(prev => prev.map(r => r.id === id ? { ...r, status: 'found', data: mappedData } : r));
      } else {
        setRows(prev => prev.map(r => r.id === id ? { ...r, status: 'not-found', data: null } : r));
      }
    } catch (err) {
      console.error("Search error:", err);
      setRows(prev => prev.map(r => r.id === id ? { ...r, status: 'not-found', data: null } : r));
    }
  };

  const handleRowKeyDown = (e: KeyboardEvent<HTMLInputElement>, row: RowData) => {
    if (e.key === 'Enter') {
      executeSearch(row.id, row.inputValue);
    }
  };

  const handleGlobalSearch = (e: KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter' && globalSearch.trim()) {
      const emptyRow = rows.find(r => !r.inputValue.trim() && !r.data);
      if (emptyRow) {
         executeSearch(emptyRow.id, globalSearch);
      } else {
         const newId = crypto.randomUUID();
         const newRow: RowData = { id: newId, inputValue: globalSearch, status: 'loading', data: null };
         setRows(prev => [...prev, newRow]);
         executeSearch(newId, globalSearch);
      }
      setGlobalSearch('');
    }
  };

  const addRows = (count: number) => {
    const newRows = Array.from({ length: count }).map(() => ({
      id: crypto.randomUUID(),
      inputValue: '',
      status: 'idle' as RowStatus,
      data: null
    }));
    setRows(prev => [...prev, ...newRows]);
  };

  const searchAll = () => {
    rows.forEach(r => {
      if (r.inputValue.trim() && !r.data && r.status !== 'loading') {
        executeSearch(r.id, r.inputValue);
      }
    });
  };

  const clearAll = () => {
    setRows(prev => prev.map(r => ({ ...r, inputValue: '', status: 'idle', data: null })));
  };

  const exportToExcel = () => {
    const dataRows = rows.filter(r => r.data && r.status === 'found');
    if (dataRows.length === 0) {
      alert("No data found to export. Please search for examiners first.");
      return;
    }
    
    // Map data to Excel format using ONLY the currently visible table columns
    const dataToExport = dataRows.map((r, index) => {
      const d = r.data!;
      const obj: Record<string, any> = {};

      // 1. Entry T-Pin / Mobile
      obj['Entry T-Pin / Mobile'] = r.inputValue;

      // 2. SL (if visible index)
      if (visibleCols.sl) {
        obj['SL'] = index + 1;
      }

      // 3. Other visible columns from ALL_COLS
      ALL_COLS.forEach(col => {
        if (col.key === 'sl' || col.key === 'entry' || col.key === 'action') return;
        if (visibleCols[col.key]) {
          const val = d[col.key as keyof Examiner];
          if (col.isScore) {
            const stats = val as SubjectStats;
            obj[col.label] = stats ? stats.score : '';
          } else {
            obj[col.label] = val;
          }
        }
      });

      return obj;
    });

    try {
      const worksheet = XLSX.utils.json_to_sheet(dataToExport);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Examiners");
      
      // Auto-size columns dynamically based on content length
      if (dataToExport.length > 0) {
        const firstRow = dataToExport[0];
        worksheet["!cols"] = Object.keys(firstRow).map(key => {
          const max_len = Math.max(
            key.length,
            ...dataToExport.map(r => String(r[key] || '').length)
          );
          return { wch: Math.min(Math.max(max_len + 2, 11), 50) };
        });
      }

      XLSX.writeFile(workbook, `Examiner_Results_${new Date().toLocaleDateString()}.xlsx`);
    } catch (err) {
      console.error("Export failed:", err);
      alert("Failed to export Excel file.");
    }
  };

  const toggleCol = (key: string) => {
    setVisibleCols(prev => ({ ...prev, [key]: !prev[key] }));
  };

  const setAllCols = (visible: boolean) => {
    const next: Record<string, boolean> = {};
    ALL_COLS.forEach(c => next[c.key] = visible);
    setVisibleCols(next);
  };

  const getScoreDisplay = (stats?: SubjectStats) => {
    if (!stats || stats.score === '') return <span className="text-slate-400">—</span>;
    if (stats.allowed) {
      return (
        <div className="inline-flex items-center gap-1.5 bg-emerald-50 text-emerald-700 px-2.5 py-0.5 rounded-md border border-emerald-200/50">
          <span className="font-semibold text-sm">{stats.score}</span>
          <CheckCircle2 className="w-3.5 h-3.5" />
        </div>
      );
    }
    return (
      <div className="inline-flex items-center gap-1.5 bg-rose-50 text-rose-700 px-2.5 py-0.5 rounded-md border border-rose-200/50">
        <span className="font-semibold text-sm">{stats.score}</span>
        <XCircle className="w-3.5 h-3.5" />
      </div>
    );
  };

  const isWorkingHours = () => {
    const now = new Date();
    const hr = now.getHours();
    return hr >= 8 && hr < 22; // 8 AM to 10 PM
  };

  const getNotificationMessage = (examiner: Examiner) => {
    const subjects = [
      { name: 'English', ...examiner.english },
      { name: 'Bangla', ...examiner.bangla },
      { name: 'Physics', ...examiner.physics },
      { name: 'Chemistry', ...examiner.chemistry },
      { name: 'Math', ...examiner.math },
      { name: 'Biology', ...examiner.biology },
      { name: 'ICT', ...examiner.ict },
    ];

    const scoredSubjects = subjects.filter(s => s.score !== '');
    const scoresSummary = scoredSubjects
      .map(s => `${s.name}: ${s.score}% (${s.allowed ? 'P' : 'F'})`)
      .join(', ');

    if (scoredSubjects.length === 0) {
      return `Hello ${examiner.name}. Your assessment is currently pending or no scores have been recorded yet. Please stay tuned for updates.`;
    }

    const passedAll = scoredSubjects.every(s => s.allowed);

    if (passedAll) {
      return `Congratulations ${examiner.name}! You have PASSED the assessment. Summary: ${scoresSummary}. Please wait for further instructions.`;
    } else {
      const failed = subjects.filter(s => s.score !== '' && !s.allowed).map(s => s.name).join(', ');
      return `Hello ${examiner.name}. Your assessment results are out. Summary: ${scoresSummary}. You did not meet the threshold in: ${failed}. Better luck next time.`;
    }
  };

  const sendWhatsApp = (examiner: Examiner) => {
    const msg = encodeURIComponent(getNotificationMessage(examiner));
    const phone = examiner.mobile.replace(/\D/g, '');
    const url = `https://wa.me/${phone.startsWith('88') ? phone : '88' + phone}?text=${msg}`;
    window.open(url, '_blank');
  };

  const sendTelegram = (examiner: Examiner) => {
    const msg = encodeURIComponent(getNotificationMessage(examiner));
    const phone = examiner.mobile.replace(/\D/g, '');
    const formattedPhone = phone.startsWith('88') ? phone : '88' + phone;
    const url = `https://t.me/+${formattedPhone}?text=${msg}`;
    window.open(url, '_blank');
  };

  return (
    <div className="min-h-screen bg-[#F8FAFC] text-slate-800 font-sans selection:bg-indigo-100 selection:text-indigo-900 flex flex-col relative">
      <header className="bg-indigo-700 text-white sticky top-0 z-40 shadow-md">
        <div className="max-w-[1600px] mx-auto px-6 py-2 flex flex-wrap items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <div className="bg-white/10 p-2 rounded-lg backdrop-blur-sm border border-white/5 flex items-center justify-center">
              <FileSpreadsheet className="w-5 h-5 text-indigo-50" />
            </div>
            <h1 className="text-[1.1rem] font-semibold tracking-wide">Examiner Assessment Checker</h1>
            <div className={`hidden sm:flex items-center gap-2 px-3 py-1 rounded-full text-xs font-medium border transition-colors ${
              connectionStatus === 'connected' ? 'bg-emerald-500/15 text-emerald-100 border-emerald-500/30' :
              connectionStatus === 'connecting' ? 'bg-amber-500/15 text-amber-100 border-amber-500/30' :
              'bg-rose-500/15 text-rose-100 border-rose-500/30 cursor-pointer hover:bg-rose-500/25'
            }`} onClick={() => setIsErrorModalOpen(true)}>
              <div className={`w-2 h-2 rounded-full ${
                connectionStatus === 'connected' ? (isSyncing ? 'bg-amber-400 animate-spin' : 'bg-emerald-400 animate-pulse box-shadow-glow') :
                connectionStatus === 'connecting' ? 'bg-amber-400 animate-pulse' :
                'bg-rose-400 font-bold'
              }`} />
              <LayoutGrid className="w-3 h-3 opacity-70" />
              {isSyncing ? 'Syncing Store...' : (connectionStatus === 'connected' ? (isWorkingHours() ? 'Instant Mode Active' : 'Ultra Search Active') : 
               connectionStatus === 'connecting' ? 'Connecting...' : 
               'Connection Error (Setup API)')}
            </div>
            {lastSyncTime > 0 && !isSyncing && (
              <div className="text-[10px] text-indigo-300 opacity-60 ml-2 italic">
                Last updated: {new Date(lastSyncTime).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
              </div>
            )}
          </div>
          
          <div className="flex items-center gap-4">
            {dbStats.rowCount !== null && (
              <div className="text-[11px] font-bold text-indigo-200 uppercase tracking-wider bg-indigo-800/50 px-2 py-1 rounded border border-indigo-600/30">
                {dbStats.rowCount.toLocaleString()} Records
              </div>
            )}
            <div className="relative group">
              <Search className="w-4 h-4 text-indigo-300 absolute left-3 top-1/2 -translate-y-1/2 transition-colors group-focus-within:text-indigo-200" />
              <input 
                className="bg-indigo-800/60 border border-indigo-600/50 text-sm text-white placeholder-indigo-300 rounded-lg pl-9 pr-4 py-2 focus:outline-none focus:ring-2 focus:ring-indigo-400 focus:bg-indigo-800 transition-all w-64 md:w-80 shadow-inner"
                placeholder="(T-Pin or Mobile)..."
                value={globalSearch}
                onChange={e => setGlobalSearch(e.target.value)}
                onKeyDown={handleGlobalSearch}
              />
              <div className="absolute right-2 top-1/2 -translate-y-1/2 flex items-center gap-1">
                 <kbd className="hidden md:inline-block text-[10px] bg-indigo-900/50 text-indigo-300 px-1.5 py-0.5 rounded border border-indigo-700">↵</kbd>
              </div>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-[1600px] mx-auto px-6 py-3 flex-1 w-full relative">
        <div className="flex flex-col md:flex-row items-start md:items-center justify-between mb-3 gap-4">
          <div className="flex items-center gap-3">
            <button 
              onClick={() => checkConnection(undefined, true)}
              disabled={isSyncing}
              className="flex items-center gap-2 bg-white border border-slate-200 shadow-sm text-indigo-600 px-3 py-1.5 rounded-xl text-sm font-semibold hover:bg-indigo-50 transition-all hover:shadow disabled:opacity-55 disabled:cursor-not-allowed cursor-pointer"
              title="Refresh local database"
            >
              <RefreshCw className={`w-4 h-4 ${isSyncing ? 'animate-spin' : ''}`} />
              {isSyncing ? 'Syncing...' : 'Sync Database'}
            </button>
            {localData && localData.length > 0 && (
              <button 
                onClick={clearLocalCache}
                className="flex items-center gap-1.5 bg-white border border-slate-200 shadow-sm text-rose-600 px-3 py-1.5 rounded-xl text-sm font-semibold hover:bg-rose-50 transition-all hover:shadow cursor-pointer"
                title="Clear cached data"
              >
                Clear Cache
              </button>
            )}
            <button 
              onClick={() => setIsColModalOpen(true)}
              className="flex items-center gap-2 bg-white border border-slate-200 shadow-sm text-slate-600 px-3 py-1.5 rounded-xl text-sm font-semibold hover:bg-slate-50 hover:text-indigo-600 transition-all hover:shadow cursor-pointer"
            >
              <LayoutGrid className="w-4 h-4" />
              Columns
            </button>
            <button 
              onClick={exportToExcel}
              className="flex items-center gap-2 bg-white border border-slate-200 shadow-sm text-slate-600 px-3 py-1.5 rounded-xl text-sm font-semibold hover:bg-slate-50 hover:text-indigo-600 transition-all hover:shadow active:scale-95 cursor-pointer"
              title="Download results as Excel (.xlsx)"
            >
              <Download className="w-4 h-4" />
              Export Excel
            </button>
            <button 
              onClick={clearAll}
              className="flex items-center gap-2 bg-white border border-slate-200 shadow-sm text-rose-600 px-3 py-1.5 rounded-xl text-sm font-semibold hover:bg-rose-50 hover:text-rose-700 transition-all hover:shadow cursor-pointer"
            >
               Clear Rows
            </button>
          </div>

          <div className="flex items-center gap-3">
            <div className="px-3 py-1 bg-indigo-50 text-indigo-700 rounded-lg text-sm font-semibold border border-indigo-100 mr-2">
               {rows.length} Rows
            </div>
            <div className="flex items-center bg-white border border-indigo-200 rounded-xl overflow-hidden focus-within:ring-2 focus-within:ring-indigo-500 focus-within:border-indigo-300 transition-all shadow-sm h-[34px] w-36">
              <input 
                type="number"
                min="1"
                max="500"
                value={addRowCount}
                onChange={(e) => setAddRowCount(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    const cnt = parseInt(addRowCount, 10);
                    if (!isNaN(cnt) && cnt > 0) addRows(cnt);
                  }
                }}
                className="w-16 h-full px-2 text-sm text-center font-bold text-indigo-700 bg-indigo-50/50 focus:outline-none placeholder-indigo-300 border-none"
                placeholder="Qty"
                title="Number of rows to add"
              />
              <button 
                onClick={() => {
                  const cnt = parseInt(addRowCount, 10);
                  if (!isNaN(cnt) && cnt > 0) addRows(cnt);
                }}
                className="flex items-center justify-center flex-1 h-full px-2 text-sm font-semibold text-indigo-600 hover:bg-indigo-50 transition-colors border-l border-indigo-100"
                title="Add Rows"
              >
                <Plus className="w-4 h-4 mr-1" />
                Add
              </button>
            </div>
            <button 
              onClick={searchAll}
              className="flex items-center gap-2 bg-indigo-600 text-white px-4 py-1.5 rounded-xl text-sm font-semibold hover:bg-indigo-700 hover:shadow-md transition-all shadow-sm"
            >
              <RefreshCw className="w-4 h-4" />
              Search Pending
            </button>
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden flex flex-col relative">
          <div className="overflow-auto max-h-[calc(100vh-280px)] min-h-[400px]">
            <table className="w-full text-left border-collapse whitespace-nowrap min-w-max">
              <thead className="sticky top-0 z-30">
                <tr className="bg-[#2462eb] border-b border-[#1d4ed8]">
                  <th className="px-4 py-1.5 text-center text-[12px] font-bold text-white tracking-wide w-[180px] min-w-[180px] sticky left-0 top-0 bg-[#2462eb] z-30 border-r border-[#437bf2]">
                    Entry<br/>T-Pin / Mobile
                  </th>
                  {visibleCols.sl && (
                    <th className="px-2 py-1.5 text-center text-[12px] font-bold text-white tracking-wide w-[60px] min-w-[60px] sticky left-[180px] top-0 bg-[#2462eb] z-30 border-r border-[#437bf2]">
                      SL
                    </th>
                  )}
                  
                  {ALL_COLS.map((col) => (col.key !== 'sl') && visibleCols[col.key] && (
                    <th key={col.key} className="px-4 py-1.5 text-center text-[12px] font-bold text-white tracking-wide min-w-[100px] whitespace-normal leading-tight align-middle border-r border-[#437bf2] sticky top-0 bg-[#2462eb] z-20">
                       {col.label}
                    </th>
                  ))}
                  
                  <th className="px-4 py-1.5 text-center text-[12px] font-bold text-white tracking-wide min-w-[100px] whitespace-normal leading-tight align-middle border-r border-[#437bf2] sticky top-0 bg-[#2462eb] z-20">
                     Action
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {rows.map((row, index) => {
                  const isEven = index % 2 === 0;
                  
                  let rowBg = '';
                  let hoverBg = '';
                  
                  if (row.status === 'found') {
                    rowBg = isEven ? 'bg-[#f4fbf7]' : 'bg-[#e7f7ed]';
                    hoverBg = isEven ? 'group-hover:bg-[#e7f7ed]' : 'group-hover:bg-[#dcfce7]';
                  } else if (row.status === 'error') {
                    rowBg = isEven ? 'bg-rose-50' : 'bg-rose-100/60';
                    hoverBg = isEven ? 'group-hover:bg-rose-100' : 'group-hover:bg-rose-200/60';
                  } else {
                    rowBg = isEven ? 'bg-white' : 'bg-[#f8fafc]';
                    hoverBg = isEven ? 'group-hover:bg-[#f1f5f9]' : 'group-hover:bg-[#e2e8f0]';
                  }

                  return (
                  <tr key={row.id} className={`group transition-colors border-b border-slate-200 ${rowBg}`}>
                    <td className={`px-3 py-1 sticky left-0 z-10 border-r border-slate-200 min-w-[200px] transition-colors ${rowBg} ${hoverBg}`}>
                      <div className="relative flex items-center gap-2">
                        <div className="relative flex-1">
                          <input 
                             type="text"
                             value={row.inputValue}
                             onChange={(e) => handleRowInput(row.id, e.target.value)}
                             onKeyDown={(e) => handleRowKeyDown(e, row)}
                             onPaste={(e) => handlePaste(e, row.id)}
                             placeholder="TPIN or Mobile"
                             className={`w-full text-[13px] rounded px-3 py-1 pr-8 focus:outline-none transition-all font-medium border
                               ${row.status === 'idle' ? 'bg-white border-slate-300 text-slate-800 focus:border-indigo-400 focus:ring-1 focus:ring-indigo-400' : ''}
                               ${row.status === 'loading' ? 'bg-amber-50 border-amber-400 text-amber-900 animate-pulse' : ''}
                               ${row.status === 'found' ? 'bg-white border-[#22c55e] text-slate-800 shadow-[0_0_2px_rgba(34,197,94,0.3)] focus:border-[#22c55e] focus:ring-1 focus:ring-[#22c55e]' : ''}
                               ${row.status === 'not-found' ? 'bg-white border-rose-300 text-rose-800 transition-none' : ''}
                             `}
                          />
                          <div className="absolute right-2 top-1/2 -translate-y-1/2 flex items-center gap-1">
                            {row.inputValue && row.status === 'idle' && (
                              <button onClick={() => executeSearch(row.id, row.inputValue)} className="p-1 text-slate-400 hover:text-indigo-600 transition-colors">
                                <Search className="w-3.5 h-3.5" />
                              </button>
                            )}
                            {row.status === 'loading' && <RefreshCw className="w-3 h-3 text-indigo-500 animate-spin" />}
                            {row.status === 'not-found' && <XCircle className="w-3.5 h-3.5 text-rose-500" />}
                            {row.status === 'found' && <CheckCircle2 className="w-3.5 h-3.5 text-emerald-500" />}
                          </div>
                        </div>
                        {row.inputValue && (
                           <button 
                             onClick={() => handleRowInput(row.id, '')}
                             className="p-1 text-slate-400 hover:text-rose-500 hover:bg-rose-50 rounded bg-white border border-slate-200"
                             title="Clear Row"
                           >
                             <X className="w-4 h-4" />
                           </button>
                        )}
                      </div>
                    </td>

                    {visibleCols.sl && (
                      <td className={`px-2 py-1 text-center text-[13px] font-medium text-slate-500 sticky left-[180px] z-10 border-r border-slate-200 transition-colors ${rowBg} ${hoverBg}`}>
                        {index + 1}
                      </td>
                    )}

                    {ALL_COLS.map((col) => {
                      if (col.key === 'sl' || col.key === 'entry' || col.key === 'action') return null;
                      if (!visibleCols[col.key]) return null;
                      
                      let cellContent: React.ReactNode = null;
                      
                      if (row.data) {
                        const val = row.data[col.key as keyof Examiner];
                        if (col.isScore) {
                           cellContent = getScoreDisplay(val as SubjectStats);
                        } else {
                           cellContent = <span className={`text-[13px] whitespace-nowrap ${index % 2 === 0 ? 'text-slate-700' : 'text-slate-800'}`}>{val as any}</span>;
                        }
                      } else if (row.status === 'loading') {
                          cellContent = <div className="h-4 w-12 bg-slate-100 rounded animate-pulse mx-auto" />;
                      }

                      return (
                        <td key={col.key} className="px-4 py-1 text-center align-middle border-r border-slate-200">
                          {cellContent}
                        </td>
                      );
                    })}

                    <td className="px-4 py-1 text-center align-middle border-r border-slate-200">
                       {row.data && (
                         <div className="flex items-center justify-center gap-2">
                           <button 
                             onClick={() => setSelectedExaminer(row.data)}
                             className="bg-indigo-50 text-indigo-600 p-1.5 rounded-md hover:bg-indigo-100 transition-colors border border-indigo-200/50"
                             title="View Full Profile"
                           >
                             <LayoutGrid className="w-4 h-4" />
                           </button>
                           <button 
                             onClick={() => sendWhatsApp(row.data!)}
                             className="bg-emerald-50 text-emerald-600 p-1.5 rounded-md hover:bg-emerald-100 transition-colors border border-emerald-200/50"
                             title="Notify via WhatsApp"
                           >
                             <MessageSquare className="w-4 h-4" />
                           </button>
                           <button 
                             onClick={() => sendTelegram(row.data!)}
                             className="bg-sky-50 text-sky-600 p-1.5 rounded-md hover:bg-sky-100 transition-colors border border-sky-200/50"
                             title="Notify via Telegram"
                           >
                             <Send className="w-4 h-4" />
                           </button>
                         </div>
                       )}
                    </td>
                  </tr>
                )})}
              </tbody>
            </table>
          </div>
          {rows.length === 0 && (
            <div className="flex-1 flex flex-col items-center justify-center p-12 text-slate-400">
              <FileSpreadsheet className="w-12 h-12 mb-4 opacity-20" />
              <p className="font-medium text-slate-500">No rows added yet</p>
              <button onClick={() => addRows(parseInt(addRowCount, 10) || 10)} className="mt-4 text-indigo-600 hover:underline text-sm font-medium">Add {parseInt(addRowCount, 10) || 10} rows to begin</button>
            </div>
          )}
        </div>
      </main>

      {isColModalOpen && (
        <div className="absolute inset-x-0 top-0 min-h-full bg-slate-900/40 backdrop-blur-sm z-50 flex justify-center items-start p-4 animate-in fade-in zoom-in duration-200 overflow-y-auto">
          <div className="bg-white rounded-lg shadow-2xl w-full max-w-4xl overflow-hidden border border-slate-200 mt-12 md:mt-24 mb-12">
            {/* Header */}
            <div className="bg-[#1a73e8] px-6 py-3 flex items-center justify-between text-white">
              <h3 className="text-lg font-bold">Column Settings</h3>
              <button 
                onClick={() => setIsColModalOpen(false)}
                className="hover:bg-white/20 p-1 rounded transition-colors"
              >
                <X className="w-6 h-6" />
              </button>
            </div>
            
            {/* Body */}
            <div className="p-8">
              <div className="grid grid-cols-4 gap-x-8 gap-y-4">
                {ALL_COLS.map(col => (
                  <label key={col.key} className="flex items-center gap-3 cursor-pointer group">
                    <div className="relative flex items-center">
                      <input 
                        type="checkbox" 
                        checked={visibleCols[col.key]} 
                        onChange={() => toggleCol(col.key)}
                        className="w-4 h-4 text-[#1a73e8] rounded border-slate-300 focus:ring-[#1a73e8] accent-[#1a73e8]"
                      />
                    </div>
                    <span className="text-[14px] text-slate-700 font-medium whitespace-nowrap group-hover:text-black">
                      {col.label}
                    </span>
                  </label>
                ))}
              </div>
            </div>
            
            {/* Footer */}
            <div className="px-6 py-4 bg-white border-t border-slate-100 flex justify-end gap-3">
              <button 
                onClick={() => setIsColModalOpen(false)}
                className="bg-[#5f6368] text-white px-6 py-2 rounded font-bold text-sm hover:bg-[#4d5156] transition-colors"
              >
                Close
              </button>
              <button 
                onClick={() => setIsColModalOpen(false)}
                className="bg-[#1a73e8] text-white px-6 py-2 rounded font-bold text-sm hover:bg-[#1557b0] transition-colors"
              >
                Apply & Save
              </button>
            </div>
          </div>
        </div>
      )}
      {isErrorModalOpen && (
        <div className="absolute inset-x-0 top-0 min-h-full bg-slate-900/60 backdrop-blur-md z-[60] flex justify-center items-start p-4 overflow-y-auto">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md overflow-hidden border border-slate-100 scale-in-center mt-12 md:mt-24 mb-12">
            <div className={`${connectionStatus === 'connected' ? 'bg-indigo-600' : 'bg-rose-600'} px-6 py-4 flex items-center gap-3 text-white`}>
              {connectionStatus === 'connected' ? <CheckCircle2 className="w-6 h-6 shrink-0" /> : <XCircle className="w-6 h-6 shrink-0" />}
              <div>
                <h3 className="font-bold text-lg leading-tight">
                  API Connection {connectionStatus === 'connected' ? 'Settings' : 'Failed'}
                </h3>
                <p className={`${connectionStatus === 'connected' ? 'text-indigo-100' : 'text-rose-100'} text-xs`}>
                  {connectionStatus === 'connected' ? 'Your app is live and linked' : 'Cannot reach Google Apps Script API'}
                </p>
              </div>
            </div>
            
            <div className="p-6 space-y-4">
              {connectionStatus !== 'connected' && (
                <div className="bg-rose-50 border border-rose-100 p-4 rounded-xl space-y-2">
                  <p className="text-sm font-bold text-rose-800">Potential Causes:</p>
                  <div className="text-[11px] text-rose-700 space-y-2">
                    <p>1. <span className="font-bold">Invalid URL</span>: The current URL below might be old or deleted.</p>
                    <p>2. <span className="font-bold">Not Deployed</span>: You must click <span className="bg-rose-200 px-1 rounded text-rose-900 border border-rose-300">Deploy &gt; New Deployment</span>.</p>
                    <p>3. <span className="font-bold">Access Denied</span>: Set "Who Has Access" to <span className="font-bold text-rose-900">Anyone</span>.</p>
                  </div>
                </div>
              )}

              {lastError && connectionStatus !== 'connected' && (
                <div className="bg-slate-900 p-3 rounded-lg text-[10px] font-mono text-rose-400 border border-slate-700 overflow-hidden">
                   <p className="text-slate-500 mb-1 font-sans">ERROR LOG:</p>
                   {lastError}
                </div>
              )}

              <div className="bg-indigo-50 border border-indigo-100 p-4 rounded-xl space-y-2">
                <p className="text-xs font-bold text-indigo-800 uppercase tracking-tighter">How to get your URL:</p>
                <ol className="text-[11px] text-indigo-700 list-decimal list-inside space-y-1">
                  <li>In Apps Script, click <span className="font-bold">Deploy</span> &gt; <span className="font-bold">New Deployment</span></li>
                  <li>Select type: <span className="font-bold">Web App</span></li>
                  <li>Execute as: <span className="font-bold">Me</span></li>
                  <li>Who has access: <span className="font-bold">Anyone</span></li>
                  <li>Copy the <span className="font-bold">Web App URL</span> and paste it below</li>
                </ol>
              </div>

              <div className="space-y-2">
                 <div className="flex justify-between items-center">
                   <p className="text-[13px] text-slate-600 font-medium leading-relaxed">
                     Web App URL:
                   </p>
                   <button 
                     onClick={resetToDefault}
                     className="text-[11px] text-indigo-600 hover:text-indigo-800 font-bold"
                   >
                     Reset to Default
                   </button>
                 </div>
                 <input 
                    type="text"
                    className="w-full bg-slate-50 border border-slate-300 p-3 rounded-xl text-[12px] break-all focus:ring-2 focus:ring-indigo-500 focus:outline-none font-mono text-indigo-700"
                    value={customUrl}
                    onChange={(e) => setCustomUrl(e.target.value)}
                    placeholder="https://script.google.com/macros/s/.../exec"
                 />
                 <p className="text-[10px] text-slate-400">
                    * Make sure you deployed as "Web App", Execute as "Me", and Access for "Anyone".
                 </p>
              </div>

              <div className="pt-2 flex flex-col gap-2">
                <button 
                  onClick={() => checkConnection(customUrl)}
                  className="w-full bg-indigo-600 text-white font-bold py-2.5 rounded-xl hover:bg-indigo-700 transition-colors shadow-md active:scale-[0.98] flex items-center justify-center gap-2"
                >
                  {connectionStatus === 'connecting' && <RefreshCw className="w-4 h-4 animate-spin" />}
                  Retry Connection
                </button>
                
                <div className="flex gap-2">
                  <button 
                    onClick={() => window.open(customUrl, '_blank')}
                    className="flex-1 bg-slate-100 text-slate-700 font-bold py-2 rounded-xl border border-slate-200 hover:bg-slate-200 text-xs text-center"
                  >
                    Test URL In Tab
                  </button>
                  <button 
                    onClick={() => setIsErrorModalOpen(false)}
                    className="flex-1 bg-slate-100 text-slate-700 font-bold py-2 rounded-xl border border-slate-200 hover:bg-slate-200 text-xs text-center"
                  >
                    Close (Offline)
                  </button>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {selectedExaminer && (
        <div className="absolute inset-x-0 top-0 min-h-full bg-slate-900/60 backdrop-blur-md z-[70] flex justify-center items-start p-4 overflow-y-auto animate-in fade-in duration-300">
          <div className="bg-white rounded-3xl shadow-2xl w-full max-w-4xl overflow-hidden border border-slate-100 scale-in-center mt-12 md:mt-24 mb-16">
             <div className="bg-indigo-700 px-8 py-6 text-white relative">
                <button 
                  onClick={() => setSelectedExaminer(null)}
                  className="absolute top-6 right-8 p-1.5 bg-white/10 rounded-full hover:bg-white/20 transition-colors"
                >
                  <X className="w-5 h-5" />
                </button>
                <div className="flex items-center gap-6">
                   <div className="w-20 h-20 bg-white/10 rounded-2xl flex items-center justify-center text-3xl font-bold border border-white/20">
                      {selectedExaminer.name.charAt(0)}
                   </div>
                   <div>
                      <h2 className="text-2xl font-bold">{selectedExaminer.name}</h2>
                      <div className="flex items-center gap-3 mt-1 text-indigo-100 text-sm">
                         <span className="bg-white/10 px-2 py-0.5 rounded border border-white/10">ID: {selectedExaminer.tpin}</span>
                         <span className="flex items-center gap-1"><RefreshCw className="w-3 h-3" /> {selectedExaminer.batch} Batch</span>
                         <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-tighter ${String(selectedExaminer.rm || '').toLowerCase().includes('0') ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/30' : 'bg-rose-500/20 text-rose-300 border border-rose-500/30'}`}>
                            RM: {selectedExaminer.rm}
                         </span>
                      </div>
                   </div>
                </div>
             </div>

             <div className="bg-slate-50">
               <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 p-8">
                  {/* Left Column: Personal info */}
                  <div className="lg:col-span-1 space-y-6">
                    <section className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100">
                      <h3 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-4">Personal Contact</h3>
                      <div className="space-y-4">
                        <div className="flex flex-col">
                           <span className="text-[10px] uppercase text-slate-400 font-bold">Mobile Number</span>
                           <span className="text-sm font-semibold text-slate-700">{selectedExaminer.mobile}</span>
                        </div>
                        <div className="flex flex-col">
                           <span className="text-[10px] uppercase text-slate-400 font-bold">Nagad Number</span>
                           <span className="text-sm font-semibold text-slate-700">{selectedExaminer.nagad}</span>
                        </div>
                        <div className="flex flex-col">
                           <span className="text-[10px] uppercase text-slate-400 font-bold">Email Address</span>
                           <span className="text-sm font-semibold text-indigo-600 truncate">{selectedExaminer.email || 'None'}</span>
                        </div>
                      </div>
                    </section>

                    <section className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100">
                      <h3 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-4">Academic</h3>
                      <div className="space-y-4">
                         <div className="flex flex-col">
                           <span className="text-[10px] uppercase text-slate-400 font-bold">Institute</span>
                           <span className="text-sm font-semibold text-slate-700">{selectedExaminer.inst}</span>
                         </div>
                         <div className="flex flex-col">
                           <span className="text-[10px] uppercase text-slate-400 font-bold">Home District</span>
                           <span className="text-sm font-semibold text-slate-700">{selectedExaminer.homeDistrict || '—'}</span>
                         </div>
                      </div>
                    </section>
                  </div>

                  {/* Middle Column: Scores & Remarks */}
                  <div className="lg:col-span-2 space-y-6">
                    <section className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
                      <h3 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-4">Subject Assessments</h3>
                      <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                        {ALL_COLS.filter(c => c.isScore).map(col => {
                           const stats = selectedExaminer[col.key as keyof Examiner] as SubjectStats;
                           return (
                             <div key={col.key} className={`p-4 rounded-xl border flex flex-col items-center justify-center gap-2 transition-all ${
                                stats.score !== '' ? (stats.allowed ? 'bg-emerald-50 border-emerald-100' : 'bg-rose-50 border-rose-100') : 'bg-slate-50 border-slate-200 opacity-60'
                             }`}>
                               <span className="text-[10px] font-bold text-slate-500 uppercase">{col.label}</span>
                               <span className={`text-xl font-black ${stats.score !== '' ? (stats.allowed ? 'text-emerald-700' : 'text-rose-700') : 'text-slate-300'}`}>
                                  {stats.score === '' ? '—' : stats.score}
                               </span>
                               {stats.score !== '' && (
                                 <span className={`text-[9px] font-bold px-1.5 py-0.5 rounded-full uppercase ${stats.allowed ? 'bg-emerald-500 text-white' : 'bg-rose-500 text-white'}`}>
                                    {stats.allowed ? 'Passed' : 'Failed'}
                                 </span>
                               )}
                             </div>
                           );
                        })}
                      </div>
                    </section>

                    {/* Remarks Card */}
                    <div className={`p-6 rounded-2xl border transition-all ${selectedExaminer.remark.show ? 'bg-rose-600 text-white border-rose-700 shadow-lg shadow-rose-200' : 'bg-white text-slate-500 border-slate-100'}`}>
                        <div className="flex items-center justify-between mb-4">
                           <h3 className={`text-xs font-bold uppercase tracking-widest ${selectedExaminer.remark.show ? 'text-rose-100' : 'text-slate-400'}`}>
                             Internal Remarks / Warning
                           </h3>
                           <div className={`px-2 py-1 rounded text-xs font-black ${selectedExaminer.remark.show ? 'bg-white text-rose-600' : 'bg-slate-100 text-slate-400'}`}>
                              Count: {selectedExaminer.remark.count}
                           </div>
                        </div>

                        {selectedExaminer.remark.show ? (
                          <div className="space-y-4">
                            <div className="bg-black/10 p-4 rounded-xl backdrop-blur-sm border border-white/5">
                               <p className="text-sm font-medium leading-relaxed whitespace-pre-line">
                                 {selectedExaminer.remark.body}
                               </p>
                            </div>
                            <div className="flex items-center gap-4 text-[11px] font-bold text-rose-100">
                               {selectedExaminer.remarkedBy && (
                                 <span className="flex items-center gap-1">
                                    <LayoutGrid className="w-3 h-3" /> User: {selectedExaminer.remarkedBy}
                                 </span>
                               )}
                            </div>
                          </div>
                        ) : (
                          <div className="flex flex-col items-center justify-center py-4 text-slate-300 italic">
                             <CheckCircle2 className="w-8 h-8 mb-2 opacity-20" />
                             <span className="text-sm">No critical remarks on record</span>
                          </div>
                        )}
                    </div>
                  </div>
               </div>
             </div>

             <div className="px-8 py-5 border-t border-slate-100 flex justify-between items-center bg-white">
                <div className="flex flex-wrap items-center gap-4 text-slate-400 text-xs font-medium">
                   <div className="flex items-center gap-2">
                     <LayoutGrid className="w-4 h-4" />
                     Campus: <span className="text-slate-700 font-bold">{selectedExaminer.campus || '—'}</span>
                   </div>
                   <div className="flex items-center gap-2 border-l border-slate-200 pl-4">
                     <RefreshCw className="w-3.5 h-3.5" />
                     Training Date: <span className="text-slate-700 font-bold">{selectedExaminer.trainingDate || '—'}</span>
                   </div>
                </div>
                 <div className="flex items-center gap-3">
                    <button 
                       onClick={() => sendWhatsApp(selectedExaminer!)}
                       className="flex items-center gap-2 bg-emerald-600 text-white font-bold px-5 py-2.5 rounded-xl hover:bg-emerald-700 transition-all shadow-md active:scale-95"
                    >
                       <MessageSquare className="w-4 h-4" />
                       WhatsApp
                    </button>
                    <button 
                       onClick={() => sendTelegram(selectedExaminer!)}
                       className="flex items-center gap-2 bg-sky-500 text-white font-bold px-5 py-2.5 rounded-xl hover:bg-sky-600 transition-all shadow-md active:scale-95"
                    >
                       <Send className="w-4 h-4" />
                       Telegram
                    </button>
                    <button 
                       onClick={() => setSelectedExaminer(null)}
                       className="bg-slate-100 text-slate-700 font-bold px-6 py-2.5 rounded-xl hover:bg-slate-200 transition-all border border-slate-200 active:scale-95"
                    >
                       Close Profile
                    </button>
                 </div>
             </div>
          </div>
        </div>
      )}
    </div>
  );
}

