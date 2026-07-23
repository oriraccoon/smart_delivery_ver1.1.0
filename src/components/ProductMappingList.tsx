import React, { useState, useRef } from 'react';
import { ProductMapping } from '../types';
import { 
  Search, Plus, Trash2, Edit2, Check, X, FileText, ClipboardList, 
  Filter, AlertTriangle, Download, Upload, FileSpreadsheet, 
  FileCode, ArrowLeft, CheckCircle2, Copy, Layers, List
} from 'lucide-react';
import * as XLSX from 'xlsx';

interface ProductMappingListProps {
  mappings: ProductMapping[];
  onChange: (updated: ProductMapping[]) => void;
  defaultRawNames?: string[];
}

type ViewMode = 'list' | 'single_add' | 'bulk_add';

export default function ProductMappingList({ mappings, onChange }: ProductMappingListProps) {
  // 뷰 모드 ('list': 목록보기 | 'single_add': 개별 추가 | 'bulk_add': 대량 추가)
  const [viewMode, setViewMode] = useState<ViewMode>('list');

  // 검색 및 필터
  const [searchQuery, setSearchQuery] = useState('');
  const [filterType, setFilterType] = useState<'all' | 'hasMapped' | 'noMapped'>('all');

  // 개별 추가 폼
  const [newRawName, setNewRawName] = useState('');
  const [newMappedName, setNewMappedName] = useState('');
  const [recentlyAdded, setRecentlyAdded] = useState<ProductMapping[]>([]);

  // 테이블 내 항목 수정
  const [editingIndex, setEditingIndex] = useState<number | null>(null);
  const [editRawName, setEditRawName] = useState('');
  const [editMappedName, setEditMappedName] = useState('');

  // 벌크 텍스트 & 파일 업로드
  const [bulkText, setBulkText] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  // 알림 & 모달
  const [toastMsg, setToastMsg] = useState<{ type: 'success' | 'error'; text: string } | null>(null);
  const [singleDeleteTarget, setSingleDeleteTarget] = useState<string | null>(null);
  const [selectedRawNames, setSelectedRawNames] = useState<Set<string>>(new Set());
  const [showDeleteModal, setShowDeleteModal] = useState<{
    isOpen: boolean;
    type: 'selected' | 'all';
    count: number;
  } | null>(null);

  // 벌크 / 파일 파싱 후 적용 확인 모달
  const [bulkImportConfirm, setBulkImportConfirm] = useState<{
    isOpen: boolean;
    parsed: ProductMapping[];
  } | null>(null);

  // 내보내기 드롭다운
  const [showExportMenu, setShowExportMenu] = useState(false);

  const showNotification = (type: 'success' | 'error', text: string) => {
    setToastMsg({ type, text });
    setTimeout(() => setToastMsg(null), 3500);
  };

  // 대치 목록 필터링
  const filteredMappings = mappings.filter(item => {
    const q = searchQuery.toLowerCase();
    const matchesSearch = (
      item.rawName.toLowerCase().includes(q) ||
      item.mappedName.toLowerCase().includes(q)
    );
    if (!matchesSearch) return false;

    if (filterType === 'hasMapped') return !!item.mappedName;
    if (filterType === 'noMapped') return !item.mappedName;
    return true;
  });

  // 개별 추가 핸들러
  const handleSingleAdd = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newRawName.trim()) return;

    const trimmedRaw = newRawName.trim();
    const trimmedMapped = newMappedName.trim();

    const exists = mappings.some(m => m.rawName.trim() === trimmedRaw);
    if (exists) {
      showNotification('error', "이미 등록된 대치 대상 품목명입니다.");
      return;
    }

    const newItem = { rawName: trimmedRaw, mappedName: trimmedMapped };
    const updated = [newItem, ...mappings];
    onChange(updated);

    setRecentlyAdded(prev => [newItem, ...prev].slice(0, 5));
    setNewRawName('');
    setNewMappedName('');
    showNotification('success', `'${trimmedRaw}' 대치 규칙이 추가되었습니다.`);
  };

  // 단일 항목 삭제
  const executeDelete = (rawName: string) => {
    const nextSelected = new Set(selectedRawNames);
    nextSelected.delete(rawName);
    setSelectedRawNames(nextSelected);

    onChange(mappings.filter(m => m.rawName !== rawName));
    setSingleDeleteTarget(null);
    showNotification('success', `'${rawName}' 규칙이 삭제되었습니다.`);
  };

  // 다중 선택 처리
  const handleSelectAll = (checked: boolean) => {
    if (checked) {
      setSelectedRawNames(new Set(filteredMappings.map(m => m.rawName)));
    } else {
      setSelectedRawNames(new Set());
    }
  };

  const handleSelectRow = (rawName: string, checked: boolean) => {
    const next = new Set(selectedRawNames);
    if (checked) next.add(rawName);
    else next.delete(rawName);
    setSelectedRawNames(next);
  };

  // 선택 삭제 / 전체 삭제
  const executeDeleteSelected = () => {
    const nextMappings = mappings.filter(m => !selectedRawNames.has(m.rawName));
    onChange(nextMappings);
    setSelectedRawNames(new Set());
    setShowDeleteModal(null);
    showNotification('success', `${selectedRawNames.size}개의 규칙이 삭제되었습니다.`);
  };

  const executeDeleteAll = () => {
    onChange([]);
    setSelectedRawNames(new Set());
    setShowDeleteModal(null);
    showNotification('success', "전체 대치 규칙이 삭제되었습니다.");
  };

  // 테이블 내 인라인 수정
  const startEdit = (globalIdx: number, item: ProductMapping) => {
    setEditingIndex(globalIdx);
    setEditRawName(item.rawName);
    setEditMappedName(item.mappedName);
  };

  const saveEdit = (globalIdx: number) => {
    if (!editRawName.trim()) return;

    const updated = [...mappings];
    updated[globalIdx] = {
      rawName: editRawName.trim(),
      mappedName: editMappedName.trim()
    };
    onChange(updated);
    setEditingIndex(null);
    showNotification('success', "대치 규칙이 수정되었습니다.");
  };

  // --- 텍스트 / 파일 파싱 로직 ---
  const parseRawText = (text: string): ProductMapping[] => {
    const lines = text.split('\n');
    const parsed: ProductMapping[] = [];
    const seen = new Set<string>();

    lines.forEach(line => {
      const trimmed = line.trim();
      if (!trimmed) return;

      let raw = '';
      let mapped = '';

      if (trimmed.includes(':')) {
        const parts = trimmed.split(':');
        raw = parts[0].replace(/['",]/g, '').trim();
        mapped = parts.slice(1).join(':').replace(/['",]/g, '').trim();
      } else if (trimmed.includes(',')) {
        const parts = trimmed.split(',');
        raw = parts[0].replace(/['"]/g, '').trim();
        mapped = parts.slice(1).join(',').replace(/['"]/g, '').trim();
      } else if (trimmed.includes('\t')) {
        const parts = trimmed.split('\t');
        raw = parts[0].replace(/['"]/g, '').trim();
        mapped = parts.slice(1).join('\t').replace(/['"]/g, '').trim();
      } else {
        raw = trimmed.replace(/['",]/g, '').trim();
      }

      if (raw && !seen.has(raw)) {
        seen.add(raw);
        parsed.push({ rawName: raw, mappedName: mapped });
      }
    });

    return parsed;
  };

  // 대량 텍스트 파싱 적용 버튼
  const handleApplyBulkText = () => {
    if (!bulkText.trim()) {
      showNotification('error', "입력된 텍스트가 없습니다.");
      return;
    }

    const parsed = parseRawText(bulkText);
    if (parsed.length === 0) {
      showNotification('error', "파싱 가능한 대치 규칙을 찾지 못했습니다. 형식을 확인해주세요.");
      return;
    }

    setBulkImportConfirm({ isOpen: true, parsed });
  };

  // 파일 업로드 처리
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const fileName = file.name.toLowerCase();
    const reader = new FileReader();

    if (fileName.endsWith('.json')) {
      reader.onload = (event) => {
        try {
          const content = event.target?.result as string;
          const json = JSON.parse(content);
          const parsed: ProductMapping[] = [];
          if (Array.isArray(json)) {
            json.forEach(item => {
              const raw = item.rawName || item['원본'] || item['원본상품명'] || item['raw'] || '';
              const mapped = item.mappedName || item['대치'] || item['표준상품명'] || item['mapped'] || '';
              if (raw) parsed.push({ rawName: String(raw).trim(), mappedName: String(mapped).trim() });
            });
          }
          if (parsed.length > 0) {
            setBulkImportConfirm({ isOpen: true, parsed });
          } else {
            showNotification('error', "JSON 파일에 유효한 대치 데이터가 없습니다.");
          }
        } catch {
          showNotification('error', "JSON 파일 형식이 올바르지 않습니다.");
        }
      };
      reader.readAsText(file);
    } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls') || fileName.endsWith('.csv')) {
      reader.onload = (event) => {
        try {
          const data = new Uint8Array(event.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const rows: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

          const parsed: ProductMapping[] = [];
          const seen = new Set<string>();

          rows.forEach((row, idx) => {
            if (!row || row.length === 0) return;
            const col0 = String(row[0] || '').trim();
            const col1 = String(row[1] || '').trim();

            // 헤더 행 건너뛰기
            if (idx === 0 && (col0.includes('원본') || col0.includes('raw') || col0.includes('상품명'))) {
              return;
            }

            if (col0 && !seen.has(col0)) {
              seen.add(col0);
              parsed.push({ rawName: col0, mappedName: col1 });
            }
          });

          if (parsed.length > 0) {
            setBulkImportConfirm({ isOpen: true, parsed });
          } else {
            showNotification('error', "엑셀/CSV 파일에서 대치 규칙을 읽지 못했습니다.");
          }
        } catch {
          showNotification('error', "파일을 읽는 중 오류가 발생했습니다.");
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      // txt 또는 텍스트 파일
      reader.onload = (event) => {
        const text = event.target?.result as string;
        setBulkText(text);
        const parsed = parseRawText(text);
        if (parsed.length > 0) {
          setBulkImportConfirm({ isOpen: true, parsed });
        } else {
          showNotification('error', "텍스트 파일에서 대치 규칙을 읽지 못했습니다.");
        }
      };
      reader.readAsText(file);
    }

    e.target.value = '';
  };

  // 대량 파싱 수락 결정을 목록에 반영
  const handleBulkConfirmDecision = (overwrite: boolean) => {
    if (!bulkImportConfirm) return;
    const { parsed } = bulkImportConfirm;

    if (overwrite) {
      onChange(parsed);
      showNotification('success', `총 ${parsed.length}개 대치 규칙으로 전체 교체되었습니다.`);
    } else {
      const merged = [...mappings];
      let addedCount = 0;
      parsed.forEach(p => {
        if (!merged.some(m => m.rawName === p.rawName)) {
          merged.push(p);
          addedCount++;
        }
      });
      onChange(merged);
      showNotification('success', `신규 ${addedCount}개 대치 규칙이 추가되었습니다.`);
    }

    setBulkImportConfirm(null);
    setBulkText('');
    setViewMode('list');
  };

  // --- 내보내기 (Export) 핸들러 ---
  const handleExportExcel = () => {
    if (mappings.length === 0) {
      showNotification('error', "내보낼 대치 규칙이 존재하지 않습니다.");
      return;
    }
    const data = mappings.map(m => ({
      '원본 상품명 키워드': m.rawName,
      '정제 후 표준 상품명': m.mappedName
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "품목대치사전");
    XLSX.writeFile(wb, `품목_대치_사전_${new Date().toISOString().slice(0, 10)}.xlsx`);
    setShowExportMenu(false);
    showNotification('success', "엑셀 파일(.xlsx) 내보내기가 완료되었습니다.");
  };

  const handleExportCSV = () => {
    if (mappings.length === 0) {
      showNotification('error', "내보낼 대치 규칙이 존재하지 않습니다.");
      return;
    }
    const data = mappings.map(m => ({
      '원본 상품명 키워드': m.rawName,
      '정제 후 표준 상품명': m.mappedName
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const csv = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob(["\uFEFF" + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `품목_대치_사전_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    setShowExportMenu(false);
    showNotification('success', "CSV 파일 내보내기가 완료되었습니다.");
  };

  const handleExportJSON = () => {
    if (mappings.length === 0) {
      showNotification('error', "내보낼 대치 규칙이 존재하지 않습니다.");
      return;
    }
    const jsonStr = JSON.stringify(mappings, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `품목_대치_사전_${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
    setShowExportMenu(false);
    showNotification('success', "JSON 파일 내보내기가 완료되었습니다.");
  };

  const handleCopyClipboardText = () => {
    if (mappings.length === 0) {
      showNotification('error', "복사할 대치 규칙이 없습니다.");
      return;
    }
    const text = mappings.map(m => `'${m.rawName}' : '${m.mappedName}',`).join('\n');
    navigator.clipboard.writeText(text).then(() => {
      showNotification('success', "전체 대치 규칙이 텍스트로 클립보드에 복사되었습니다.");
      setShowExportMenu(false);
    });
  };

  return (
    <div id="product-mapping-container" className="space-y-6">
      
      {/* 알림 토스트 배너 */}
      {toastMsg && (
        <div className={`p-3.5 rounded-xl border text-xs font-semibold flex items-center justify-between shadow-sm animate-fade-in ${
          toastMsg.type === 'error' 
            ? 'bg-rose-50 border-rose-200 text-rose-800' 
            : 'bg-emerald-50 border-emerald-200 text-emerald-800'
        }`}>
          <div className="flex items-center gap-2">
            {toastMsg.type === 'error' ? (
              <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
            ) : (
              <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
            )}
            <span>{toastMsg.text}</span>
          </div>
          <button onClick={() => setToastMsg(null)} className="text-slate-400 hover:text-slate-600">✕</button>
        </div>
      )}

      {/* 숨겨진 파일 선택 Input */}
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileUpload}
        accept=".xlsx,.xls,.csv,.json,.txt"
        className="hidden"
      />

      {/* 대량 추가 파싱 결과 확인 모달 */}
      {bulkImportConfirm && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center z-50 p-4 animate-fade-in">
          <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full p-6 space-y-4 border border-slate-100">
            <h3 className="font-bold text-slate-800 text-base flex items-center gap-2">
              <Layers className="w-5 h-5 text-indigo-600" />
              품목 대치 사전 대량 등록 확인
            </h3>
            <p className="text-xs text-slate-600 leading-relaxed">
              분석 결과 <strong className="text-indigo-600 font-bold">{bulkImportConfirm.parsed.length}개</strong>의 대치 규칙을 읽었습니다. 
              기존 목록에 어떻게 반영하시겠습니까?
            </p>
            <div className="grid grid-cols-2 gap-3 pt-2">
              <button
                onClick={() => handleBulkConfirmDecision(false)}
                className="px-4 py-2.5 bg-slate-50 hover:bg-slate-100 border border-slate-200 text-slate-700 rounded-xl text-xs font-semibold transition"
              >
                기존 목록에 추가
              </button>
              <button
                onClick={() => handleBulkConfirmDecision(true)}
                className="px-4 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-semibold transition shadow-md"
              >
                전체 덮어쓰기
              </button>
            </div>
            <div className="flex justify-end pt-1">
              <button
                onClick={() => setBulkImportConfirm(null)}
                className="text-xs text-slate-400 hover:text-slate-600"
              >
                취소
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 단일 항목 삭제 확인 모달 */}
      {singleDeleteTarget && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center z-50 p-4 animate-fade-in">
          <div className="bg-white rounded-2xl shadow-xl max-w-sm w-full p-6 space-y-4 border border-slate-100">
            <h3 className="font-bold text-slate-800 text-base flex items-center gap-1.5 text-rose-600">
              <AlertTriangle className="w-5 h-5 shrink-0" />
              대치 규칙 삭제
            </h3>
            <p className="text-xs text-slate-600 leading-relaxed">
              '<strong className="text-slate-800">{singleDeleteTarget}</strong>' 규칙을 삭제하시겠습니까?
            </p>
            <div className="flex justify-end gap-2 pt-2">
              <button
                onClick={() => setSingleDeleteTarget(null)}
                className="px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-semibold transition"
              >
                취소
              </button>
              <button
                onClick={() => executeDelete(singleDeleteTarget)}
                className="px-3.5 py-2 bg-rose-600 hover:bg-rose-700 text-white rounded-xl text-xs font-semibold transition shadow-md"
              >
                삭제
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 일괄/전체 삭제 재확인 모달 */}
      {showDeleteModal?.isOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center z-50 p-4 animate-fade-in">
          <div className="bg-white rounded-2xl shadow-xl max-w-sm w-full p-6 space-y-4 border border-slate-100">
            <h3 className="font-bold text-slate-800 text-base flex items-center gap-1.5 text-rose-600">
              <AlertTriangle className="w-5 h-5 shrink-0" />
              {showDeleteModal.type === 'all' ? '전체 규칙 삭제' : '선택 규칙 삭제'}
            </h3>
            <p className="text-xs text-slate-600 leading-relaxed">
              {showDeleteModal.type === 'all' 
                ? `주의! 현재 등록된 전체 ${mappings.length}개의 대치 규칙을 삭제하시겠습니까?`
                : `선택한 ${showDeleteModal.count}개의 대치 규칙을 삭제하시겠습니까?`
              }
            </p>
            <div className="flex justify-end gap-2 pt-2">
              <button
                onClick={() => setShowDeleteModal(null)}
                className="px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-semibold transition"
              >
                취소
              </button>
              <button
                onClick={showDeleteModal.type === 'all' ? executeDeleteAll : executeDeleteSelected}
                className="px-3.5 py-2 bg-rose-600 hover:bg-rose-700 text-white rounded-xl text-xs font-semibold transition shadow-md"
              >
                삭제 진행
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 페이지 상단 헤더 & 모드 전환 */}
      <div className="bg-white border border-slate-200 rounded-2xl p-5 shadow-sm space-y-4">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
              <ClipboardList className="w-5 h-5 text-indigo-600" />
              품목 대치 사전 관리
            </h2>
            <p className="text-xs text-slate-500 mt-1">
              원본 상품명의 긴 키워드를 정제하고, 요약 보고서에 표시할 심플한 표준 상품명으로 자동 대치합니다.
            </p>
          </div>

          {/* 우측 내보내기 버튼 */}
          <div className="flex items-center gap-2 shrink-0">
            {/* 내보내기 드롭다운 */}
            <div className="relative">
              <button
                onClick={() => setShowExportMenu(!showExportMenu)}
                className="flex items-center gap-1.5 px-3.5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-semibold transition shadow-sm"
              >
                <Download className="w-3.5 h-3.5" />
                목록 내보내기
              </button>

              {showExportMenu && (
                <div className="absolute right-0 mt-2 w-48 bg-white rounded-xl shadow-xl border border-slate-150 py-1.5 z-30 animate-fade-in">
                  <button
                    onClick={handleExportExcel}
                    className="w-full text-left px-3.5 py-2 hover:bg-slate-50 text-xs font-medium text-slate-700 flex items-center gap-2"
                  >
                    <FileSpreadsheet className="w-4 h-4 text-emerald-600" />
                    엑셀 파일 (.xlsx)
                  </button>
                  <button
                    onClick={handleExportCSV}
                    className="w-full text-left px-3.5 py-2 hover:bg-slate-50 text-xs font-medium text-slate-700 flex items-center gap-2"
                  >
                    <FileText className="w-4 h-4 text-blue-600" />
                    CSV 파일 (.csv)
                  </button>
                  <button
                    onClick={handleExportJSON}
                    className="w-full text-left px-3.5 py-2 hover:bg-slate-50 text-xs font-medium text-slate-700 flex items-center gap-2"
                  >
                    <FileCode className="w-4 h-4 text-purple-600" />
                    JSON 파일 (.json)
                  </button>
                  <div className="my-1 border-t border-slate-100" />
                  <button
                    onClick={handleCopyClipboardText}
                    className="w-full text-left px-3.5 py-2 hover:bg-slate-50 text-xs font-medium text-slate-700 flex items-center gap-2"
                  >
                    <Copy className="w-4 h-4 text-slate-500" />
                    텍스트 전체 복사
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* 메인 탭 모드 전환 버튼 (목록 보기 / 개별 추가 / 대량 추가) */}
        <div className="flex items-center gap-2 pt-2 border-t border-slate-100">
          <button
            onClick={() => setViewMode('list')}
            className={`flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold transition ${
              viewMode === 'list'
                ? 'bg-indigo-600 text-white shadow-sm'
                : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
            }`}
          >
            <List className="w-4 h-4" />
            대치 사전 목록 ({mappings.length}개)
          </button>

          <button
            onClick={() => setViewMode('single_add')}
            className={`flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold transition ${
              viewMode === 'single_add'
                ? 'bg-indigo-600 text-white shadow-sm'
                : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
            }`}
          >
            <Plus className="w-4 h-4" />
            개별 규칙 추가
          </button>

          <button
            onClick={() => setViewMode('bulk_add')}
            className={`flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold transition ${
              viewMode === 'bulk_add'
                ? 'bg-indigo-600 text-white shadow-sm'
                : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
            }`}
          >
            <Layers className="w-4 h-4" />
            대량 추가 / 파일 가져오기
          </button>
        </div>
      </div>

      {/* ========================================================= */}
      {/* 1. 개별 추가 모드 (viewMode === 'single_add') - 기존 목록 안보임 */}
      {/* ========================================================= */}
      {viewMode === 'single_add' && (
        <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm max-w-2xl mx-auto space-y-6 animate-fade-in">
          <div className="flex items-center justify-between border-b border-slate-100 pb-4">
            <div>
              <h3 className="font-bold text-slate-800 text-base flex items-center gap-2">
                <Plus className="w-5 h-5 text-indigo-600" />
                개별 대치 규칙 추가
              </h3>
              <p className="text-xs text-slate-500 mt-0.5">단일 상품명 키워드를 등록합니다.</p>
            </div>
            <button
              onClick={() => setViewMode('list')}
              className="flex items-center gap-1 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-xl transition"
            >
              <ArrowLeft className="w-3.5 h-3.5" />
              목록으로 돌아가기
            </button>
          </div>

          <form onSubmit={handleSingleAdd} className="space-y-4">
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">
                대치 대상 원본 상품명 (검색 키워드) <span className="text-rose-500">*</span>
              </label>
              <input
                type="text"
                required
                placeholder="예: 국내제조 비린맛 없는 절단 돌산 간장 꽃게장"
                value={newRawName}
                onChange={(e) => setNewRawName(e.target.value)}
                className="w-full text-xs border border-slate-200 rounded-xl p-3 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
              />
              <p className="text-[11px] text-slate-400 mt-1">※ 엑셀 업로드 시 이 단어가 포함된 상품명을 감지합니다.</p>
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">
                대치 후 표준 상품명 (보고서 정제명)
              </label>
              <input
                type="text"
                placeholder="예: 게장"
                value={newMappedName}
                onChange={(e) => setNewMappedName(e.target.value)}
                className="w-full text-xs border border-slate-200 rounded-xl p-3 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
              />
            </div>

            <div className="pt-2 flex gap-3">
              <button
                type="submit"
                className="flex-1 py-3 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold transition shadow-md"
              >
                사전에 규칙 추가
              </button>
              <button
                type="button"
                onClick={() => setViewMode('list')}
                className="px-5 py-3 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition"
              >
                완료 후 목록으로
              </button>
            </div>
          </form>

          {/* 이번 세션에서 새로 추가된 항목 내역 */}
          {recentlyAdded.length > 0 && (
            <div className="pt-4 border-t border-slate-100">
              <h4 className="text-xs font-bold text-slate-600 mb-2">최근 추가된 규칙</h4>
              <div className="space-y-1.5">
                {recentlyAdded.map((item, idx) => (
                  <div key={idx} className="flex items-center justify-between bg-slate-50 p-2.5 rounded-xl text-xs border border-slate-100">
                    <span className="font-semibold text-slate-800">{item.rawName}</span>
                    <span className="text-indigo-600 font-bold">→ {item.mappedName || '(동일)'}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}

      {/* ========================================================= */}
      {/* 2. 대량 추가 모드 (viewMode === 'bulk_add') - 기존 목록 안보임 */}
      {/* ========================================================= */}
      {viewMode === 'bulk_add' && (
        <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm space-y-6 animate-fade-in">
          <div className="flex items-center justify-between border-b border-slate-100 pb-4">
            <div>
              <h3 className="font-bold text-slate-800 text-base flex items-center gap-2">
                <Layers className="w-5 h-5 text-indigo-600" />
                대량 추가 (텍스트 / 파일)
              </h3>
              <p className="text-xs text-slate-500 mt-0.5">
                여러 대치 규칙을 한 번에 텍스트로 붙여넣거나 파일(.xlsx, .csv, .json, .txt)을 드래그해 업로드하세요.
              </p>
            </div>
            <button
              onClick={() => setViewMode('list')}
              className="flex items-center gap-1 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-xl transition"
            >
              <ArrowLeft className="w-3.5 h-3.5" />
              목록으로 돌아가기
            </button>
          </div>

          {/* 파일 드롭존 / 선택 영역 */}
          <div 
            onClick={() => fileInputRef.current?.click()}
            className="border-2 border-dashed border-indigo-200 hover:border-indigo-400 bg-indigo-50/30 hover:bg-indigo-50/60 rounded-2xl p-6 text-center cursor-pointer transition group"
          >
            <Upload className="w-8 h-8 text-indigo-500 group-hover:scale-110 transition-transform mx-auto mb-2" />
            <p className="text-xs font-bold text-slate-700">
              이곳을 클릭하여 엑셀(.xlsx), CSV(.csv), JSON(.json), TXT 파일을 선택하세요.
            </p>
            <p className="text-[11px] text-slate-400 mt-1">첫 번째 열: 원본 상품명 / 두 번째 열: 대치 후 표준 상품명</p>
          </div>

          {/* 텍스트 직접 입력 라벨 */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <label className="text-xs font-bold text-slate-700">텍스트 직접 붙여넣기</label>
              <span className="text-[11px] text-slate-400">형식 예시: '원본상품명' : '대치상품명', 또는 원본,대치</span>
            </div>
            <textarea
              value={bulkText}
              onChange={(e) => setBulkText(e.target.value)}
              placeholder={`'국내제조 비린맛 없는 절단 돌산 간장 꽃게장' : '게장',\n'명태품은 백명란' : '명란',`}
              rows={12}
              className="w-full text-xs font-mono p-4 bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
            />
          </div>

          <div className="flex justify-end gap-3 pt-2">
            <button
              onClick={() => {
                setBulkText('');
                setViewMode('list');
              }}
              className="px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition"
            >
              취소
            </button>
            <button
              onClick={handleApplyBulkText}
              className="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold transition shadow-md flex items-center gap-1.5"
            >
              <Check className="w-4 h-4" />
              대량 파싱 및 목록에 적용
            </button>
          </div>
        </div>
      )}

      {/* ========================================================= */}
      {/* 3. 목록 보기 모드 (viewMode === 'list') - 테이블 전체 표시 */}
      {/* ========================================================= */}
      {viewMode === 'list' && (
        <div className="space-y-4 animate-fade-in">
          
          {/* 검색 & 필터 & 선택 작업 바 */}
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 bg-white border border-slate-200 rounded-2xl p-3.5 shadow-sm">
            
            {/* 검색창 */}
            <div className="flex items-center gap-2 flex-1 border border-slate-200 rounded-xl px-3 py-2 bg-slate-50/50">
              <Search className="w-4 h-4 text-slate-400 shrink-0" />
              <input
                type="text"
                placeholder="대치 규칙 검색 (원본 또는 대치명)..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full text-xs border-none outline-none bg-transparent text-slate-700"
              />
              {searchQuery && (
                <button onClick={() => setSearchQuery('')} className="text-xs text-slate-400 hover:text-slate-600">
                  지우기
                </button>
              )}
            </div>

            {/* 필터 및 삭제 작업 */}
            <div className="flex items-center gap-2 shrink-0 flex-wrap">
              <div className="flex items-center gap-1.5 border border-slate-200 rounded-xl px-3 py-2 bg-slate-50/50">
                <Filter className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                <select
                  value={filterType}
                  onChange={(e) => {
                    setFilterType(e.target.value as any);
                    setSelectedRawNames(new Set());
                  }}
                  className="text-xs border-none bg-transparent focus:ring-0 text-slate-700 font-semibold py-0 pl-0 pr-6 outline-none"
                >
                  <option value="all">전체 ({mappings.length}개)</option>
                  <option value="hasMapped">대치어 지정됨 ({mappings.filter(m => m.mappedName).length})</option>
                  <option value="noMapped">동일 정제 ({mappings.filter(m => !m.mappedName).length})</option>
                </select>
              </div>

              {selectedRawNames.size > 0 && (
                <button
                  onClick={() => setShowDeleteModal({ isOpen: true, type: 'selected', count: selectedRawNames.size })}
                  className="flex items-center gap-1.5 px-3.5 py-2 bg-rose-50 hover:bg-rose-100 text-rose-600 border border-rose-150 rounded-xl text-xs font-semibold transition cursor-pointer"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                  선택 삭제 ({selectedRawNames.size})
                </button>
              )}

              {mappings.length > 0 && (
                <button
                  onClick={() => setShowDeleteModal({ isOpen: true, type: 'all', count: mappings.length })}
                  className="flex items-center gap-1.5 px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-600 border border-slate-200 rounded-xl text-xs font-semibold transition cursor-pointer"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                  전체 삭제
                </button>
              )}
            </div>
          </div>

          {/* 매핑 규칙 테이블 */}
          <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
            <div className="overflow-x-auto">
              <table className="w-full text-sm text-left text-slate-600">
                <thead className="bg-slate-50/80 text-slate-400 text-xs font-semibold uppercase border-b border-slate-150">
                  <tr>
                    <th className="px-4 py-3.5 w-10 text-center">
                      <input
                        type="checkbox"
                        checked={filteredMappings.length > 0 && filteredMappings.every(m => selectedRawNames.has(m.rawName))}
                        onChange={(e) => handleSelectAll(e.target.checked)}
                        className="w-3.5 h-3.5 text-indigo-600 border-slate-300 rounded focus:ring-indigo-500 cursor-pointer"
                        title="전체 선택 / 해제"
                      />
                    </th>
                    <th className="px-4 py-3.5">원본 상품명 키워드 (검색용)</th>
                    <th className="px-4 py-3.5">정제 후 표준 상품명</th>
                    <th className="px-4 py-3.5 text-right w-28">작업</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredMappings.length === 0 ? (
                    <tr>
                      <td colSpan={4} className="px-4 py-12 text-center text-xs text-slate-400 italic">
                        {searchQuery ? "검색 조건과 일치하는 대치 규칙이 존재하지 않습니다." : "등록된 대치 규칙이 없습니다. [개별 추가] 또는 [대량 추가] 버튼으로 등록해보세요."}
                      </td>
                    </tr>
                  ) : (
                    filteredMappings.map((item) => {
                      const globalIdx = mappings.findIndex(m => m.rawName === item.rawName);
                      const isEditing = editingIndex === globalIdx;
                      const isSelected = selectedRawNames.has(item.rawName);

                      return (
                        <tr 
                          key={item.rawName} 
                          className={`hover:bg-slate-50/60 transition-colors ${isSelected ? 'bg-indigo-50/20' : ''}`}
                        >
                          <td className="px-4 py-3.5 text-center">
                            <input
                              type="checkbox"
                              checked={isSelected}
                              onChange={(e) => handleSelectRow(item.rawName, e.target.checked)}
                              className="w-3.5 h-3.5 text-indigo-600 border-slate-300 rounded focus:ring-indigo-500 cursor-pointer"
                            />
                          </td>
                          <td className="px-4 py-3.5 font-medium text-slate-800">
                            {isEditing ? (
                              <input
                                type="text"
                                value={editRawName}
                                onChange={(e) => setEditRawName(e.target.value)}
                                className="w-full text-xs border border-slate-200 rounded-lg p-1.5 focus:outline-none focus:ring-2 focus:ring-indigo-500/20"
                              />
                            ) : (
                              item.rawName
                            )}
                          </td>
                          <td className="px-4 py-3.5 text-slate-700">
                            {isEditing ? (
                              <input
                                type="text"
                                value={editMappedName}
                                onChange={(e) => setEditMappedName(e.target.value)}
                                className="w-full text-xs border border-slate-200 rounded-lg p-1.5 focus:outline-none focus:ring-2 focus:ring-indigo-500/20"
                              />
                            ) : (
                              item.mappedName ? (
                                <span className="font-semibold text-indigo-600">{item.mappedName}</span>
                              ) : (
                                <span className="text-slate-400 italic text-xs">동일 (입력 안 됨)</span>
                              )
                            )}
                          </td>
                          <td className="px-4 py-3.5 text-right">
                            {isEditing ? (
                              <div className="flex items-center justify-end gap-1">
                                <button
                                  onClick={() => saveEdit(globalIdx)}
                                  className="p-1.5 text-emerald-600 hover:bg-emerald-50 rounded-lg"
                                  title="저장"
                                >
                                  <Check className="w-4 h-4" />
                                </button>
                                <button
                                  onClick={() => setEditingIndex(null)}
                                  className="p-1.5 text-rose-500 hover:bg-rose-50 rounded-lg"
                                  title="취소"
                                >
                                  <X className="w-4 h-4" />
                                </button>
                              </div>
                            ) : (
                              <div className="flex items-center justify-end gap-1">
                                <button
                                  onClick={() => startEdit(globalIdx, item)}
                                  className="p-1.5 text-slate-400 hover:text-indigo-600 hover:bg-slate-50 rounded-lg transition"
                                  title="수정"
                                >
                                  <Edit2 className="w-3.5 h-3.5" />
                                </button>
                                <button
                                  onClick={() => setSingleDeleteTarget(item.rawName)}
                                  className="p-1.5 text-slate-400 hover:text-rose-500 hover:bg-slate-50 rounded-lg transition"
                                  title="삭제"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </div>
                            )}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>

            <div className="px-4 py-3 bg-slate-50 border-t border-slate-100 flex items-center justify-between text-xs text-slate-500 font-medium">
              <span>총 {mappings.length}개의 대치 규칙 등록됨</span>
              {filteredMappings.length !== mappings.length && (
                <span>(검색 필터링 {filteredMappings.length}개 노출 중)</span>
              )}
            </div>
          </div>

        </div>
      )}

    </div>
  );
}
