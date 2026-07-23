import React, { useState } from 'react';
import { MatchingError, ProcessedOrder } from '../types';
import { AlertCircle, CheckCircle, HelpCircle, Layers, Package, Sparkles, FileText } from 'lucide-react';

interface WaybillReportDashboardProps {
  matchCount: number;
  softMatchCount: number;
  bundledCount?: number;
  bundledGroupCount?: number;
  unfilledErrors: MatchingError[];
  unusedErrors: MatchingError[];
  totalInputRows: number;
  outputOrders?: ProcessedOrder[];
}

export default function WaybillReportDashboard({
  matchCount,
  softMatchCount,
  bundledCount: propBundledCount,
  bundledGroupCount: propBundledGroupCount,
  unfilledErrors,
  unusedErrors,
  totalInputRows,
  outputOrders = []
}: WaybillReportDashboardProps) {
  // 주문 분류 필터 및 탭
  const [activeTab, setActiveTab] = useState<'all' | 'exact' | 'soft' | 'bundled' | 'unfilled' | 'unused'>(
    softMatchCount > 0 ? 'soft' : unfilledErrors.length > 0 ? 'unfilled' : 'all'
  );

  const exactMatchOrders = outputOrders.filter(o => o.matchType === 'exact');
  const softMatchOrders = outputOrders.filter(o => o.matchType === 'soft');
  const bundledOrders = outputOrders.filter(o => o.matchType === 'bundled');

  // 합포장 묶음(송장/고객 그룹) 수 산정
  const calculatedBundledGroupCount = new Set(
    bundledOrders.map(o => `${o.수령인}_${o.우편번호}_${o.trackingNumber || ''}`)
  ).size;

  const displayBundledGroupCount = propBundledGroupCount !== undefined ? propBundledGroupCount : calculatedBundledGroupCount;
  const displayBundledCount = propBundledCount !== undefined ? propBundledCount : bundledOrders.length;

  const totalSuccess = matchCount + softMatchCount + displayBundledCount;
  const successRate = totalInputRows > 0 ? Math.round((totalSuccess / totalInputRows) * 100) : 0;

  const getDisplayPlatform = (platformName?: string, fileName?: string) => {
    if (platformName && platformName.trim()) {
      return platformName.trim();
    }
    if (!fileName) return '플랫폼';
    const fn = fileName.toLowerCase();
    if (fn.includes('토스') || fn.includes('toss')) return '토스';
    if (fn.includes('스마트스토어') || fn.includes('네이버') || fn.includes('naver')) return '스마트스토어';
    if (fn.includes('쿠팡') || fn.includes('coupang')) return '쿠팡';
    if (fn.includes('11번가') || fn.includes('11st') || fn.includes('eleven')) return '11번가';
    if (fn.includes('옥션') || fn.includes('지마켓') || fn.includes('gmarket') || fn.includes('auction')) return 'G마켓/옥션';
    if (fn.includes('카카오') || fn.includes('kakao')) return '카카오스토어';
    if (fn.includes('자사몰') || fn.includes('카페24') || fn.includes('cafe24')) return '자사몰';

    const clean = fileName.replace(/\.[^/.]+$/, '').replace(/[\d\-_]/g, ' ').trim();
    return clean || '주문서';
  };

  return (
    <div id="waybill-report-dashboard" className="bg-slate-50 border border-slate-200 rounded-xl p-6 space-y-6">
      
      {/* 6가지 요약 비주얼 카테고리 버튼 */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
        
        {/* 1. 전체 주문 & 성공률 */}
        <button
          onClick={() => setActiveTab('all')}
          className={`p-3.5 rounded-xl border text-left transition flex flex-col justify-between cursor-pointer ${
            activeTab === 'all'
              ? 'bg-slate-800 border-slate-900 text-white shadow-md ring-2 ring-slate-400/30'
              : 'bg-white border-slate-200 shadow-sm hover:border-slate-300'
          }`}
        >
          <div className="flex items-center justify-between w-full">
            <span className={`text-xs font-semibold ${activeTab === 'all' ? 'text-slate-300' : 'text-slate-500'}`}>
              전체 주문
            </span>
            <div className={`p-1.5 rounded-md ${activeTab === 'all' ? 'bg-slate-700 text-emerald-400' : 'bg-slate-100 text-slate-600'}`}>
              <CheckCircle className="w-4 h-4" />
            </div>
          </div>
          <div className="mt-2">
            <div className={`text-lg font-bold ${activeTab === 'all' ? 'text-white' : 'text-slate-800'}`}>
              {totalInputRows}건
            </div>
            <div className={`text-[10px] mt-0.5 ${activeTab === 'all' ? 'text-emerald-400 font-medium' : 'text-emerald-600 font-semibold'}`}>
              성공률 {successRate}% ({totalSuccess}건)
            </div>
          </div>
        </button>

        {/* 2. 일반 완전 매칭 (100%) */}
        <button
          onClick={() => setActiveTab('exact')}
          className={`p-3.5 rounded-xl border text-left transition flex flex-col justify-between cursor-pointer ${
            activeTab === 'exact'
              ? 'bg-indigo-50/90 border-indigo-300 shadow-sm ring-2 ring-indigo-500/20'
              : 'bg-white border-slate-200 shadow-sm hover:border-indigo-200'
          }`}
        >
          <div className="flex items-center justify-between w-full">
            <span className="text-xs font-semibold text-slate-500">완전 매칭</span>
            <div className="p-1.5 bg-indigo-100 text-indigo-600 rounded-md">
              <Package className="w-4 h-4" />
            </div>
          </div>
          <div className="mt-2">
            <div className="text-lg font-bold text-slate-800">{matchCount}건</div>
            <div className="text-[10px] text-indigo-600 font-medium mt-0.5">100% 일치 매칭</div>
          </div>
        </button>

        {/* 3. 지능형 부분 매칭 */}
        <button
          onClick={() => setActiveTab('soft')}
          className={`p-3.5 rounded-xl border text-left transition flex flex-col justify-between cursor-pointer ${
            activeTab === 'soft'
              ? 'bg-amber-50/90 border-amber-300 shadow-sm ring-2 ring-amber-500/20'
              : 'bg-white border-slate-200 shadow-sm hover:border-amber-200'
          }`}
        >
          <div className="flex items-center justify-between w-full">
            <span className="text-xs font-semibold text-slate-500">지능형 부분매칭</span>
            <div className="p-1.5 bg-amber-100 text-amber-600 rounded-md">
              <Sparkles className="w-4 h-4" />
            </div>
          </div>
          <div className="mt-2">
            <div className="text-lg font-bold text-slate-800">{softMatchCount}건</div>
            <div className="text-[10px] text-amber-600 font-medium mt-0.5">유사 성명/대치 매칭</div>
          </div>
        </button>

        {/* 4. 합포장 (묶음배송) 카테고리 */}
        <button
          onClick={() => setActiveTab('bundled')}
          className={`p-3.5 rounded-xl border text-left transition flex flex-col justify-between cursor-pointer ${
            activeTab === 'bundled'
              ? 'bg-sky-50/90 border-sky-300 shadow-sm ring-2 ring-sky-500/20'
              : 'bg-white border-slate-200 shadow-sm hover:border-sky-200'
          }`}
        >
          <div className="flex items-center justify-between w-full">
            <span className="text-xs font-semibold text-slate-500">합포장 묶음</span>
            <div className="p-1.5 bg-sky-100 text-sky-700 rounded-md">
              <Layers className="w-4 h-4" />
            </div>
          </div>
          <div className="mt-2">
            <div className="text-lg font-bold text-sky-900">{displayBundledGroupCount}묶음</div>
            <div className="text-[10px] text-sky-600 font-medium mt-0.5">총 {displayBundledCount}개 품목 통합</div>
          </div>
        </button>

        {/* 5. 못 채운 값 (송장 누락/에러) */}
        <button
          onClick={() => setActiveTab('unfilled')}
          className={`p-3.5 rounded-xl border text-left transition flex flex-col justify-between cursor-pointer ${
            activeTab === 'unfilled'
              ? 'bg-rose-50/90 border-rose-300 shadow-sm ring-2 ring-rose-500/20'
              : 'bg-white border-slate-200 shadow-sm hover:border-rose-200'
          }`}
        >
          <div className="flex items-center justify-between w-full">
            <span className="text-xs font-semibold text-slate-500">못 채운 값</span>
            <div className="p-1.5 bg-rose-100 text-rose-600 rounded-md">
              <AlertCircle className="w-4 h-4" />
            </div>
          </div>
          <div className="mt-2">
            <div className="text-lg font-bold text-slate-800">{unfilledErrors.length}건</div>
            <div className="text-[10px] text-rose-600 font-medium mt-0.5">송장 누락/조건 불일치</div>
          </div>
        </button>

        {/* 6. 못 넣은 값 (미사용 송장) */}
        <button
          onClick={() => setActiveTab('unused')}
          className={`p-3.5 rounded-xl border text-left transition flex flex-col justify-between cursor-pointer ${
            activeTab === 'unused'
              ? 'bg-purple-50/90 border-purple-300 shadow-sm ring-2 ring-purple-500/20'
              : 'bg-white border-slate-200 shadow-sm hover:border-purple-200'
          }`}
        >
          <div className="flex items-center justify-between w-full">
            <span className="text-xs font-semibold text-slate-500">못 넣은 값</span>
            <div className="p-1.5 bg-purple-100 text-purple-600 rounded-md">
              <FileText className="w-4 h-4" />
            </div>
          </div>
          <div className="mt-2">
            <div className="text-lg font-bold text-slate-800">{unusedErrors.length}건</div>
            <div className="text-[10px] text-purple-600 font-medium mt-0.5">미사용 잔여 송장</div>
          </div>
        </button>

      </div>

      {/* 매칭 상세 리포트 메인 컨테이너 */}
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
        
        {/* 6가지 탭 메뉴 */}
        <div className="flex flex-wrap border-b border-slate-150 bg-slate-50/60">
          <button
            onClick={() => setActiveTab('all')}
            className={`py-3 px-4 text-xs font-semibold text-center border-b-2 transition ${
              activeTab === 'all'
                ? 'border-slate-800 text-slate-900 bg-white'
                : 'border-transparent text-slate-500 hover:text-slate-800'
            }`}
          >
            [1] 전체 주문 ({outputOrders.length}건)
          </button>
          <button
            onClick={() => setActiveTab('exact')}
            className={`py-3 px-4 text-xs font-semibold text-center border-b-2 transition ${
              activeTab === 'exact'
                ? 'border-indigo-600 text-indigo-700 bg-indigo-50/40'
                : 'border-transparent text-slate-500 hover:text-slate-800'
            }`}
          >
            [2] 완전 매칭 ({exactMatchOrders.length}건)
          </button>
          <button
            onClick={() => setActiveTab('soft')}
            className={`py-3 px-4 text-xs font-semibold text-center border-b-2 transition ${
              activeTab === 'soft'
                ? 'border-amber-500 text-amber-700 bg-amber-50/40'
                : 'border-transparent text-slate-500 hover:text-slate-800'
            }`}
          >
            [3] 지능형 부분 매칭 ({softMatchOrders.length}건) ✨
          </button>
          <button
            onClick={() => setActiveTab('bundled')}
            className={`py-3 px-4 text-xs font-semibold text-center border-b-2 transition ${
              activeTab === 'bundled'
                ? 'border-sky-600 text-sky-800 bg-sky-50/40'
                : 'border-transparent text-slate-500 hover:text-slate-800'
            }`}
          >
            [4] 📦 합포장 묶음 ({displayBundledGroupCount}묶음 / {displayBundledCount}개)
          </button>
          <button
            onClick={() => setActiveTab('unfilled')}
            className={`py-3 px-4 text-xs font-semibold text-center border-b-2 transition ${
              activeTab === 'unfilled'
                ? 'border-rose-500 text-rose-700 bg-rose-50/40'
                : 'border-transparent text-slate-500 hover:text-slate-800'
            }`}
          >
            [5] 못 채운 값 ({unfilledErrors.length}건) - 송장 누락
          </button>
          <button
            onClick={() => setActiveTab('unused')}
            className={`py-3 px-4 text-xs font-semibold text-center border-b-2 transition ${
              activeTab === 'unused'
                ? 'border-purple-600 text-purple-700 bg-purple-50/40'
                : 'border-transparent text-slate-500 hover:text-slate-800'
            }`}
          >
            [6] 못 넣은 값 ({unusedErrors.length}건) - 미사용 송장
          </button>
        </div>

        {/* 탭 테이블 콘텐츠 */}
        <div className="p-4">
          
          {/* 1. 전체 주문 탭 */}
          {activeTab === 'all' && (
            <div className="space-y-4">
              <p className="text-xs text-slate-600 flex items-center gap-1.5 bg-slate-100 p-3 rounded-lg border border-slate-200">
                <CheckCircle className="w-4 h-4 text-slate-700 shrink-0" />
                <span>
                  처리 대상 총 <strong>{outputOrders.length}건</strong>의 주문에 대해 운송장 매칭 및 합포장 연동 상태를 한눈에 조회합니다.
                </span>
              </p>

              <div className="overflow-x-auto border border-slate-150 rounded-lg">
                <table className="w-full text-xs text-left text-slate-600">
                  <thead className="bg-slate-50 text-slate-500 font-semibold uppercase border-b border-slate-150">
                    <tr>
                      <th className="px-3 py-2.5">플랫폼</th>
                      <th className="px-3 py-2.5">수령인</th>
                      <th className="px-3 py-2.5">우편번호</th>
                      <th className="px-3 py-2.5">수량</th>
                      <th className="px-3 py-2.5">주문 품목명</th>
                      <th className="px-3 py-2.5">부여된 운송장번호</th>
                      <th className="px-3 py-2.5">매칭 분류 및 근거</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {outputOrders.map((ord, i) => {
                      const isExact = ord.matchType === 'exact';
                      const isSoft = ord.matchType === 'soft';
                      const isBundled = ord.matchType === 'bundled';
                      const platformLabel = getDisplayPlatform(ord.platformName, ord.sourceFileName);

                      return (
                        <tr key={i} className="hover:bg-slate-50 transition">
                          <td className="px-3 py-2.5">
                            <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-slate-800 bg-slate-100 px-2 py-0.5 rounded border border-slate-200" title={ord.sourceFileName || ord.platformName}>
                              <FileText className="w-3 h-3 text-slate-500 shrink-0" />
                              {platformLabel}
                            </span>
                          </td>
                          <td className="px-3 py-2.5 font-bold text-slate-800">{ord.수령인}</td>
                          <td className="px-3 py-2.5 font-mono">{ord.우편번호}</td>
                          <td className="px-3 py-2.5 font-semibold text-slate-700">{ord.수량}</td>
                          <td className="px-3 py-2.5 max-w-xs truncate" title={ord.originalProductName || ord.상품명}>
                            {ord.originalProductName || ord.상품명}
                          </td>
                          <td className="px-3 py-2.5 font-mono font-bold text-slate-800">
                            {ord.trackingNumber ? (
                              <span className="px-2 py-1 rounded bg-slate-100 border border-slate-200">
                                {ord.trackingNumber}
                              </span>
                            ) : (
                              <span className="text-rose-500 font-semibold">미매칭 (누락)</span>
                            )}
                          </td>
                          <td className="px-3 py-2.5">
                            {isExact && (
                              <span className="inline-flex items-center gap-1 text-emerald-700 font-semibold bg-emerald-50 px-2 py-0.5 rounded border border-emerald-100">
                                <CheckCircle className="w-3 h-3" /> 완전 매칭 (100%)
                              </span>
                            )}
                            {isSoft && (
                              <span className="inline-flex items-center gap-1 text-amber-800 font-medium bg-amber-50 px-2 py-0.5 rounded border border-amber-200">
                                <Sparkles className="w-3 h-3 text-amber-600" /> {ord.softMatchReason || "지능형 부분 매칭"}
                              </span>
                            )}
                            {isBundled && (
                              <span className="inline-flex items-center gap-1 text-sky-800 font-semibold bg-sky-50 px-2 py-0.5 rounded border border-sky-200">
                                <Layers className="w-3 h-3 text-sky-600" /> {ord.softMatchReason || "합포장 묶음 배송"}
                              </span>
                            )}
                            {!isExact && !isSoft && !isBundled && (
                              <span className="text-rose-600 font-medium">
                                {ord.unmatchedDetail ? "매칭 조건 미충족 (송장 엑셀 누락)" : "운송장 미기입"}
                              </span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* 2. 일반 완전 매칭 탭 */}
          {activeTab === 'exact' && (
            <div className="space-y-4">
              <p className="text-xs text-slate-600 flex items-center gap-1.5 bg-indigo-50/80 p-3 rounded-lg border border-indigo-150">
                <Package className="w-4 h-4 text-indigo-600 shrink-0" />
                <span>
                  <strong>일반 완전 매칭</strong>은 성명, 우편번호, 연락처, 수량, 품목명이 100% 교차 대조되어 완벽하게 송장이 매칭된 주문 목록입니다.
                </span>
              </p>

              <div className="overflow-x-auto border border-slate-150 rounded-lg">
                <table className="w-full text-xs text-left text-slate-600">
                  <thead className="bg-slate-50 text-slate-500 font-semibold uppercase border-b border-slate-150">
                    <tr>
                      <th className="px-3 py-2.5">플랫폼</th>
                      <th className="px-3 py-2.5">수령인</th>
                      <th className="px-3 py-2.5">우편번호</th>
                      <th className="px-3 py-2.5">수량</th>
                      <th className="px-3 py-2.5">주문 품목명</th>
                      <th className="px-3 py-2.5">부여된 운송장번호</th>
                      <th className="px-3 py-2.5">상태</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {exactMatchOrders.map((ord, i) => (
                      <tr key={i} className="hover:bg-indigo-50/30 transition">
                        <td className="px-3 py-2.5">
                          <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-slate-800 bg-slate-100 px-2 py-0.5 rounded border border-slate-200" title={ord.sourceFileName || ord.platformName}>
                            <FileText className="w-3 h-3 text-slate-500 shrink-0" />
                            {getDisplayPlatform(ord.platformName, ord.sourceFileName)}
                          </span>
                        </td>
                        <td className="px-3 py-2.5 font-bold text-slate-800">{ord.수령인}</td>
                        <td className="px-3 py-2.5 font-mono">{ord.우편번호}</td>
                        <td className="px-3 py-2.5 font-semibold text-slate-700">{ord.수량}</td>
                        <td className="px-3 py-2.5 max-w-xs truncate" title={ord.originalProductName || ord.상품명}>
                          {ord.originalProductName || ord.상품명}
                        </td>
                        <td className="px-3 py-2.5 font-mono font-bold text-indigo-800 bg-indigo-50/50 rounded">
                          {ord.trackingNumber || "-"}
                        </td>
                        <td className="px-3 py-2.5 text-emerald-600 font-semibold flex items-center gap-1">
                          <CheckCircle className="w-3.5 h-3.5" />
                          100% 완전 일치
                        </td>
                      </tr>
                    ))}
                    {exactMatchOrders.length === 0 && (
                      <tr>
                        <td colSpan={7} className="px-3 py-8 text-center text-slate-400">
                          완전 매칭된 주문이 없습니다.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* 3. 지능형 부분 매칭 탭 */}
          {activeTab === 'soft' && (
            <div className="space-y-4">
              <p className="text-xs text-slate-600 flex items-center gap-1.5 bg-amber-50/80 p-3 rounded-lg border border-amber-150">
                <Sparkles className="w-4 h-4 text-amber-600 shrink-0" />
                <span>
                  <strong>지능형 부분 매칭</strong>은 수령인 성명 공백/짤림(80% 이상 유사), 품목 대치어 교차 대조 등을 통해 100% 문자가 일치하지 않더라도 정확한 대상을 안전하게 추론해 송장번호를 연결한 내역입니다.
                </span>
              </p>

              <div className="overflow-x-auto border border-slate-150 rounded-lg">
                <table className="w-full text-xs text-left text-slate-600">
                  <thead className="bg-slate-50 text-slate-500 font-semibold uppercase border-b border-slate-150">
                    <tr>
                      <th className="px-3 py-2.5">플랫폼</th>
                      <th className="px-3 py-2.5">수령인</th>
                      <th className="px-3 py-2.5">우편번호</th>
                      <th className="px-3 py-2.5">수량</th>
                      <th className="px-3 py-2.5">주문 품목명</th>
                      <th className="px-3 py-2.5">부여된 운송장번호</th>
                      <th className="px-3 py-2.5">지능형 매칭 유추 근거</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {softMatchOrders.map((ord, i) => (
                      <tr key={i} className="hover:bg-amber-50/30 transition">
                        <td className="px-3 py-2.5">
                          <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-slate-800 bg-slate-100 px-2 py-0.5 rounded border border-slate-200" title={ord.sourceFileName || ord.platformName}>
                            <FileText className="w-3 h-3 text-slate-500 shrink-0" />
                            {getDisplayPlatform(ord.platformName, ord.sourceFileName)}
                          </span>
                        </td>
                        <td className="px-3 py-2.5 font-bold text-slate-800">{ord.수령인}</td>
                        <td className="px-3 py-2.5 font-mono">{ord.우편번호}</td>
                        <td className="px-3 py-2.5 font-semibold text-slate-700">{ord.수량}</td>
                        <td className="px-3 py-2.5 max-w-xs truncate" title={ord.originalProductName || ord.상품명}>
                          {ord.originalProductName || ord.상품명}
                        </td>
                        <td className="px-3 py-2.5 font-mono font-bold text-amber-800 bg-amber-50/50 rounded">
                          {ord.trackingNumber || "-"}
                        </td>
                        <td className="px-3 py-2.5 text-amber-800 font-medium">
                          {ord.softMatchReason || "수령인 성명 유사 및 대치어 교차 매칭"}
                        </td>
                      </tr>
                    ))}
                    {softMatchOrders.length === 0 && (
                      <tr>
                        <td colSpan={7} className="px-3 py-8 text-center text-slate-400">
                          지능형 부분 매칭된 주문이 없습니다.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* 4. 합포장 (묶음배송) 탭 */}
          {activeTab === 'bundled' && (
            <div className="space-y-4">
              <p className="text-xs text-sky-900 flex items-center gap-2 bg-sky-50 p-3 rounded-lg border border-sky-200">
                <Layers className="w-4 h-4 text-sky-600 shrink-0" />
                <span>
                  <strong>합포장 묶음배송 내역:</strong> 동일한 수취인 및 배송지로 여러 상품이 주문되어 대표 운송장에 통합 동봉 출고된 건입니다.
                  현재 <strong>총 {displayBundledGroupCount}명의 수령인(묶음)</strong>에게 <strong>{displayBundledCount}개 품목</strong>이 합포장 발송 처리되었습니다.
                </span>
              </p>

              <div className="overflow-x-auto border border-sky-150 rounded-lg">
                <table className="w-full text-xs text-left text-slate-600">
                  <thead className="bg-sky-100/60 text-sky-900 font-semibold uppercase border-b border-sky-200">
                    <tr>
                      <th className="px-3 py-2.5">플랫폼</th>
                      <th className="px-3 py-2.5">수령인</th>
                      <th className="px-3 py-2.5">우편번호</th>
                      <th className="px-3 py-2.5">수량</th>
                      <th className="px-3 py-2.5">합포장 주문 품목명</th>
                      <th className="px-3 py-2.5">통합 부여 운송장번호</th>
                      <th className="px-3 py-2.5">합포장 안내</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-sky-100">
                    {bundledOrders.map((ord, i) => (
                      <tr key={i} className="hover:bg-sky-50/50 transition">
                        <td className="px-3 py-2.5">
                          <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-slate-800 bg-slate-100 px-2 py-0.5 rounded border border-slate-200" title={ord.sourceFileName || ord.platformName}>
                            <FileText className="w-3 h-3 text-slate-500 shrink-0" />
                            {getDisplayPlatform(ord.platformName, ord.sourceFileName)}
                          </span>
                        </td>
                        <td className="px-3 py-2.5 font-bold text-slate-800">{ord.수령인}</td>
                        <td className="px-3 py-2.5 font-mono">{ord.우편번호}</td>
                        <td className="px-3 py-2.5 font-semibold text-slate-700">{ord.수량}</td>
                        <td className="px-3 py-2.5 max-w-xs truncate font-medium text-slate-800" title={ord.originalProductName || ord.상품명}>
                          {ord.originalProductName || ord.상품명}
                        </td>
                        <td className="px-3 py-2.5 font-mono font-bold text-sky-900 bg-sky-100/80 rounded">
                          {ord.trackingNumber || "-"}
                        </td>
                        <td className="px-3 py-2.5 text-sky-800 font-medium">
                          {ord.softMatchReason || "동일 수취인 타 주문건과 대표 송장으로 합포장 출고됨"}
                        </td>
                      </tr>
                    ))}
                    {bundledOrders.length === 0 && (
                      <tr>
                        <td colSpan={7} className="px-3 py-8 text-center text-slate-400">
                          합포장 묶음 배송 처리된 주문건이 없습니다.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* 5. 못 채운 값 (송장 누락) 탭 */}
          {activeTab === 'unfilled' && (
            <div className="space-y-4">
              <p className="text-xs text-slate-600 flex items-center gap-1.5 bg-rose-50/80 p-3 rounded-lg border border-rose-150">
                <AlertCircle className="w-4 h-4 text-rose-500 shrink-0" />
                <span>
                  <strong>못 채운 값 (송장 누락):</strong> 주문서에는 존재하나 송장파일에서 매칭되는 운송장 데이터를 찾지 못한 미매칭 항목입니다. 아래의 <strong>[플랫폼]</strong> 열에서 어느 쇼핑몰의 주문인지 확인하실 수 있습니다.
                </span>
              </p>

              <div className="overflow-x-auto border border-slate-150 rounded-lg">
                <table className="w-full text-xs text-left text-slate-600">
                  <thead className="bg-slate-50 text-slate-500 font-semibold uppercase border-b border-slate-150">
                    <tr>
                      <th className="px-3 py-2.5">플랫폼</th>
                      <th className="px-3 py-2.5">수령인</th>
                      <th className="px-3 py-2.5">우편번호</th>
                      <th className="px-3 py-2.5">수량</th>
                      <th className="px-3 py-2.5">주문 품목명</th>
                      <th className="px-3 py-2.5">원인 진단 및 대조 내역</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {unfilledErrors.map((err, i) => (
                      <tr key={i} className="hover:bg-rose-50/30">
                        <td className="px-3 py-2.5">
                          <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-rose-800 bg-rose-50 px-2 py-1 rounded border border-rose-200" title={err.sourceFileName || err.platformName}>
                            <FileText className="w-3 h-3 text-rose-500 shrink-0" />
                            {getDisplayPlatform(err.platformName, err.sourceFileName)}
                          </span>
                        </td>
                        <td className="px-3 py-2.5 font-bold text-slate-800">{err.name}</td>
                        <td className="px-3 py-2.5 font-mono">{err.zipCode}</td>
                        <td className="px-3 py-2.5 font-semibold text-slate-700">{err.qty}</td>
                        <td className="px-3 py-2.5 max-w-xs truncate" title={err.prod}>{err.prod}</td>
                        <td className="px-3 py-2.5">
                          <div className="text-rose-600 font-semibold">{err.reason}</div>
                          {err.debugDetail && (
                            <div className="mt-1 text-[11px] p-2 rounded border border-rose-100 bg-rose-50/70 text-slate-700 whitespace-pre-wrap font-mono leading-normal">
                              {err.debugDetail}
                            </div>
                          )}
                        </td>
                      </tr>
                    ))}
                    {unfilledErrors.length === 0 && (
                      <tr>
                        <td colSpan={6} className="px-3 py-8 text-center text-slate-400">
                          누락된 주문이 없습니다. 모든 주문에 운송장 연결 완료! ✨
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* 6. 못 넣은 값 (미사용 송장) 탭 */}
          {activeTab === 'unused' && (
            <div className="space-y-4">
              <p className="text-xs text-purple-900 flex items-center gap-1.5 bg-purple-50 p-3 rounded-lg border border-purple-200">
                <FileText className="w-4 h-4 text-purple-600 shrink-0" />
                <span>
                  <strong>못 넣은 값 (미사용 송장):</strong> 소스(송장) 엑셀에는 운송장이 존재하나 주문서에 대상이 없어 사용되지 않고 남은 잔여 송장 내역입니다. 아래의 <strong>[송장 플랫폼]</strong> 열에서 출처를 확인하실 수 있습니다.
                </span>
              </p>

              <div className="overflow-x-auto border border-purple-150 rounded-lg">
                <table className="w-full text-xs text-left text-slate-600">
                  <thead className="bg-purple-100/60 text-purple-900 font-semibold uppercase border-b border-purple-200">
                    <tr>
                      <th className="px-3 py-2.5">송장 플랫폼</th>
                      <th className="px-3 py-2.5">송장 수령인</th>
                      <th className="px-3 py-2.5">우편번호</th>
                      <th className="px-3 py-2.5">수량</th>
                      <th className="px-3 py-2.5">송장 품목명</th>
                      <th className="px-3 py-2.5">상태 및 이유</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-purple-100">
                    {unusedErrors.map((err, i) => (
                      <tr key={i} className="hover:bg-purple-50/40">
                        <td className="px-3 py-2.5">
                          <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-purple-800 bg-purple-50 px-2 py-1 rounded border border-purple-200" title={err.sourceFileName}>
                            <FileText className="w-3 h-3 text-purple-500 shrink-0" />
                            {getDisplayPlatform(err.platformName, err.sourceFileName)}
                          </span>
                        </td>
                        <td className="px-3 py-2.5 font-bold text-slate-800">{err.name}</td>
                        <td className="px-3 py-2.5 font-mono">{err.zipCode}</td>
                        <td className="px-3 py-2.5 font-semibold">{err.qty}</td>
                        <td className="px-3 py-2.5 max-w-xs truncate" title={err.prod}>{err.prod}</td>
                        <td className="px-3 py-2.5 text-purple-700 font-medium">{err.reason}</td>
                      </tr>
                    ))}
                    {unusedErrors.length === 0 && (
                      <tr>
                        <td colSpan={6} className="px-3 py-8 text-center text-slate-400">
                          미사용 송장이 존재하지 않습니다. 모든 송장이 주문과 정확히 매칭되었습니다!
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

        </div>

      </div>
    </div>
  );
}

// 미려한 회전 새로고침 아이콘 컴포넌트
function RefreshCwIcon(props: React.SVGProps<SVGSVGElement>) {
  return (
    <svg
      {...props}
      xmlns="http://www.w3.org/2000/svg"
      width="24"
      height="24"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      <path d="M21 12a9 9 0 0 0-9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
      <path d="M3 3v5h5" />
      <path d="M3 12a9 9 0 0 0 9 9 9.75 9.75 0 0 0 6.74-2.74L21 16" />
      <path d="M16 16h5v5" />
    </svg>
  );
}
