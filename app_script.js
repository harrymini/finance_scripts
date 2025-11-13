/****************************************************
 * Global Liquidity Monitor v3.0 - 완전 통합 버전
 * 
 * 주요 기능:
 * 1. 미국 유동성 모니터링 (WALCL, TGA, ON RRP)
 * 2. 글로벌 유동성 추적 (중국 M2, BOJ, DXY)
 * 3. 신흥국 통화 모니터링
 * 4. 종합 유동성 점수 및 자동 알림
 * 5. 알림 설정/해제 기능
 * 6. 히스토리 자동 누적 (History, Global_History, Alert_History)
 ****************************************************/

const CONFIG = {
  SHEET_NAME: 'Live_Monitor',
  HISTORY_SHEET: 'History',
  GLOBAL_SHEET: 'Global_Liquidity',
  GLOBAL_HISTORY_SHEET: 'Global_History',
  ALERT_HISTORY_SHEET: 'Alert_History',
  CACHE_TIME: 300000, // 5분 캐시
  
  // 미국 지표
  FRED_IDS: {
    SOFR: 'SOFRINDEX',
    EFFR: 'EFFR',
    IORB: 'IORB',
    ON_RRP: 'RRPONTSYD',
    TGA: 'WTREGEN',
    WALCL: 'WALCL'
  },
  
  // 글로벌 지표
  GLOBAL_FRED_IDS: {
    // 달러 인덱스
    DXY: 'DTWEXBGS',
    
    // 중국 지표
    CHINA_M2_YOY: 'MABMM301CNM657S',
    CHINA_LOAN: 'QCNLOANTOPRIV',
    CHINA_RESERVES: 'TRESEGCNM052N',
    
    // 일본 지표
    USDJPY: 'DEXJPUS',
    JGB_10Y: 'IRLTLT01JPM156N',
    
    // 신흥국 통화
    USDKRW: 'DEXKOUS',
    USDBRL: 'DEXBZUS',
    USDMXN: 'DEXMXUS',
    
    // VIX
    VIX: 'VIXCLS'
  },
  
  FRED_BASE: 'https://fred.stlouisfed.org/graph/fredgraph.csv',
  SRF_API: 'https://markets.newyorkfed.org/api/operations/standing-repo-facility'
};

/** ===============================================
 * 1) FRED 데이터 수집 (기본 + 히스토리)
 * =============================================== */

function getFredData(fredId, useCache = true) {
  const cacheKey = `FRED_${fredId}`;
  const cache = CacheService.getScriptCache();
  
  if (useCache) {
    const cached = cache.get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }
  }
  
  try {
    const url = `${CONFIG.FRED_BASE}?id=${fredId}`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' },
      timeout: 15000
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`FRED API 오류: ${response.getResponseCode()}`);
    }
    
    const csv = response.getContentText();
    const lines = csv.trim().split('\n');
    
    if (lines.length < 2) {
      throw new Error(`FRED 데이터가 없음: ${fredId}`);
    }
    
    const lastLine = lines[lines.length - 1];
    const [date, value] = lastLine.split(',');
    
    const result = {
      date: date.trim(),
      value: parseFloat(value.trim()),
      timestamp: new Date().getTime(),
      fredId: fredId
    };
    
    cache.put(cacheKey, JSON.stringify(result), Math.floor(CONFIG.CACHE_TIME / 1000));
    
    return result;
  } catch (e) {
    Logger.log(`❌ FRED 수집 실패 [${fredId}]: ${e.message}`);
    return { value: null, error: e.message, fredId: fredId };
  }
}

function getFredDataHistorical(fredId, daysAgo) {
  try {
    const url = `${CONFIG.FRED_BASE}?id=${fredId}`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      timeout: 15000
    });
    
    if (response.getResponseCode() !== 200) {
      return { value: 0 };
    }
    
    const csv = response.getContentText();
    const lines = csv.trim().split('\n');
    
    const targetIndex = Math.max(lines.length - Math.ceil(daysAgo/5) - 1, 1);
    
    if (targetIndex < lines.length) {
      const [date, value] = lines[targetIndex].split(',');
      return {
        date: date.trim(),
        value: parseFloat(value.trim())
      };
    }
    
    return { value: 0 };
    
  } catch (e) {
    Logger.log(`❌ Historical 데이터 오류: ${e.message}`);
    return { value: 0 };
  }
}

/** ===============================================
 * 2) SRF 데이터 수집
 * =============================================== */

function getSRFData() {
  const cacheKey = 'SRF_LATEST';
  const cache = CacheService.getScriptCache();
  
  const cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
  
  try {
    const apiUrl = 'https://markets.newyorkfed.org/api/rp/all/all/results/latest/1.json';
    const response = UrlFetchApp.fetch(apiUrl, {
      muteHttpExceptions: true,
      headers: { 
        'User-Agent': 'Mozilla/5.0',
        'Accept': 'application/json'
      },
      timeout: 15000
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      
      if (data && data.repo && data.repo.operations && data.repo.operations.length > 0) {
        let srfOperation = null;
        
        for (const op of data.repo.operations) {
          if (op.operationType && 
              (op.operationType.includes('Standing') || 
               op.operationType.includes('SRF'))) {
            srfOperation = op;
            break;
          }
        }
        
        if (srfOperation) {
          const result = {
            date: srfOperation.operationDate || srfOperation.effectiveDate,
            amount: srfOperation.totalAmtAccepted || 0,
            rate: srfOperation.awardRate || 0,
            timestamp: new Date().getTime(),
            source: 'api_repo_operations'
          };
          
          if (result.amount > 0 && result.amount < 1000) {
            result.amount = result.amount * 1000;
          }
          
          cache.put(cacheKey, JSON.stringify(result), 86400);
          return result;
        }
      }
    }
  } catch (e) {
    Logger.log(`⚠️ SRF API 실패: ${e.message}`);
  }
  
  return { 
    amount: 0, 
    date: new Date().toISOString().split('T')[0],
    rate: 0,
    error: 'No data available',
    source: 'default'
  };
}

/** ===============================================
 * 3) 중국 유동성 모니터링
 * =============================================== */

function getChinaLiquidity() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'CHINA_LIQUIDITY';
    
    const cached = cache.get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }
    
    const m2_yoy = getFredData(CONFIG.GLOBAL_FRED_IDS.CHINA_M2_YOY, false);
    const loans = getFredData(CONFIG.GLOBAL_FRED_IDS.CHINA_LOAN, false);
    const reserves = getFredData(CONFIG.GLOBAL_FRED_IDS.CHINA_RESERVES, false);
    
    const result = {
      m2_growth: m2_yoy.value || 0,
      m2_date: m2_yoy.date,
      total_credit: loans.value || 0,
      fx_reserves: reserves.value || 0,
      liquidity_signal: determineChinaSignal(m2_yoy.value),
      timestamp: new Date().getTime()
    };
    
    cache.put(cacheKey, JSON.stringify(result), 3600);
    Logger.log(`✅ 중국 유동성 데이터: M2 YoY ${result.m2_growth}%`);
    
    return result;
    
  } catch (e) {
    Logger.log(`❌ 중국 데이터 오류: ${e.message}`);
    return { m2_growth: 0, total_credit: 0, liquidity_signal: 'NO DATA' };
  }
}

function determineChinaSignal(m2_growth) {
  if (m2_growth > 12) {
    return '🔴 과잉 유동성';
  } else if (m2_growth > 10) {
    return '✅ 적정 성장';
  } else if (m2_growth > 8) {
    return '⚖️ 중립';
  } else if (m2_growth > 6) {
    return '⚠️ 성장 둔화';
  } else {
    return '🔵 유동성 부족';
  }
}

/** ===============================================
 * 4) 일본/엔캐리 모니터링
 * =============================================== */

function getJapanLiquidity() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'JAPAN_LIQUIDITY';
    
    const cached = cache.get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }
    
    const usdjpy = getFredData(CONFIG.GLOBAL_FRED_IDS.USDJPY, false);
    const jgb10y = getFredData(CONFIG.GLOBAL_FRED_IDS.JGB_10Y, false);
    const us10y = getFredData('DGS10', false);
    
    const result = {
      usdjpy: usdjpy.value || 0,
      jgb_10y: jgb10y.value || 0,
      us_jpy_spread: (us10y.value || 0) - (jgb10y.value || 0),
      carry_risk: determineCarryRisk(usdjpy.value, (us10y.value || 0) - (jgb10y.value || 0)),
      timestamp: new Date().getTime()
    };
    
    cache.put(cacheKey, JSON.stringify(result), 3600);
    Logger.log(`✅ 일본 데이터: USDJPY ${result.usdjpy}`);
    
    return result;
    
  } catch (e) {
    Logger.log(`❌ 일본 데이터 오류: ${e.message}`);
    return { usdjpy: 0, carry_risk: 'NO DATA' };
  }
}

function determineCarryRisk(usdjpy, spread) {
  if (usdjpy > 150 && spread > 4) {
    return '🔴 극도의 리스크';
  } else if (usdjpy > 145 && spread > 3.5) {
    return '⚠️ 높은 리스크';
  } else if (usdjpy > 140) {
    return '⚖️ 중간 리스크';
  } else if (usdjpy < 130) {
    return '💨 언와인드 진행';
  } else {
    return '✅ 안정적';
  }
}

/** ===============================================
 * 5) TGA 상세 분석
 * =============================================== */

function getTGAAnalysis() {
  try {
    const tga = getFredData(CONFIG.FRED_IDS.TGA, false);
    const tga_1w = getFredDataHistorical(CONFIG.FRED_IDS.TGA, 7);
    const tga_1m = getFredDataHistorical(CONFIG.FRED_IDS.TGA, 30);
    
    const current = tga.value || 0;
    const weekAgo = tga_1w.value || current;
    const monthAgo = tga_1m.value || current;
    
    const weekChange = current - weekAgo;
    const monthChange = current - monthAgo;
    
    return {
      current: current,
      week_change: weekChange,
      month_change: monthChange,
      liquidity_impact: determineTGAImpact(weekChange, monthChange),
      debt_ceiling_risk: checkDebtCeilingRisk(current)
    };
    
  } catch (e) {
    Logger.log(`❌ TGA 분석 오류: ${e.message}`);
    return { current: 0, liquidity_impact: 'NO DATA' };
  }
}

function determineTGAImpact(weekChange, monthChange) {
  if (monthChange < -100000) {
    return '🚀 대규모 유동성 공급';
  } else if (monthChange < -50000) {
    return '✅ 유동성 공급중';
  } else if (monthChange > 50000) {
    return '⚠️ 유동성 흡수중';
  } else if (monthChange > 100000) {
    return '🔴 대규모 유동성 흡수';
  } else {
    return '⚖️ 중립';
  }
}

function checkDebtCeilingRisk(tga_balance) {
  if (tga_balance < 100000) {
    return '🔴 부채한도 리스크';
  } else if (tga_balance < 200000) {
    return '⚠️ 주의 필요';
  } else {
    return '✅ 충분';
  }
}

/** ===============================================
 * 6) 신흥국 통화 모니터링
 * =============================================== */

function getEmergingMarketsFX() {
  try {
    const usdkrw = getFredData(CONFIG.GLOBAL_FRED_IDS.USDKRW, false);
    const usdbrl = getFredData(CONFIG.GLOBAL_FRED_IDS.USDBRL, false);
    const usdmxn = getFredData(CONFIG.GLOBAL_FRED_IDS.USDMXN, false);
    
    const usdkrw_1w = getFredDataHistorical(CONFIG.GLOBAL_FRED_IDS.USDKRW, 7);
    const usdbrl_1w = getFredDataHistorical(CONFIG.GLOBAL_FRED_IDS.USDBRL, 7);
    const usdmxn_1w = getFredDataHistorical(CONFIG.GLOBAL_FRED_IDS.USDMXN, 7);
    
    const krw_change = ((usdkrw.value - usdkrw_1w.value) / usdkrw_1w.value) * 100;
    const brl_change = ((usdbrl.value - usdbrl_1w.value) / usdbrl_1w.value) * 100;
    const mxn_change = ((usdmxn.value - usdmxn_1w.value) / usdmxn_1w.value) * 100;
    
    const strength_index = -(krw_change + brl_change + mxn_change) / 3;
    
    return {
      usdkrw: usdkrw.value || 0,
      usdbrl: usdbrl.value || 0,
      usdmxn: usdmxn.value || 0,
      krw_change: krw_change,
      brl_change: brl_change,
      mxn_change: mxn_change,
      strength_index: strength_index,
      signal: strength_index > 1 ? '✅ EM 강세' : 
              strength_index < -1 ? '⚠️ EM 약세' : '⚖️ 중립'
    };
    
  } catch (e) {
    Logger.log(`❌ EM FX 오류: ${e.message}`);
    return { strength_index: 0, signal: 'NO DATA' };
  }
}

/** ===============================================
 * 7) 글로벌 유동성 종합 분석 + History 기록
 * =============================================== */

function analyzeGlobalLiquidity() {
  try {
    const ss = SpreadsheetApp.getActive();
    let globalSheet = ss.getSheetByName(CONFIG.GLOBAL_SHEET);
    
    if (!globalSheet) {
      globalSheet = ss.insertSheet(CONFIG.GLOBAL_SHEET);
      setupGlobalSheet(globalSheet);
    }
    
    // 데이터 수집
    const walcl = getFredData(CONFIG.FRED_IDS.WALCL);
    const walcl_1w = getFredDataHistorical(CONFIG.FRED_IDS.WALCL, 7);
    const tga = getTGAAnalysis();
    const onRrp = getFredData(CONFIG.FRED_IDS.ON_RRP);
    
    const dxy = getFredData(CONFIG.GLOBAL_FRED_IDS.DXY);
    const dxy_1w = getFredDataHistorical(CONFIG.GLOBAL_FRED_IDS.DXY, 7);
    const dxy_change = (dxy.value || 100) - (dxy_1w.value || 100);
    
    const china = getChinaLiquidity();
    const japan = getJapanLiquidity();
    const emFx = getEmergingMarketsFX();
    
    // WoW 계산
    const walcl_wow = (walcl.value || 0) - (walcl_1w.value || 0);
    
    // 종합 유동성 점수 계산
    let liquidityScore = 0;
    
    // 미국 요인 (40%)
    if (walcl_wow > 0) liquidityScore += 20;
    if (tga.week_change < -10000) liquidityScore += 10;
    if (onRrp.value < 200000) liquidityScore += 10;
    
    // 달러 요인 (20%)
    if (dxy_change < -1) liquidityScore += 20;
    else if (dxy_change > 1) liquidityScore -= 20;
    
    // 중국 요인 (20%)
    if (china.m2_growth > 10) liquidityScore += 20;
    else if (china.m2_growth < 8) liquidityScore -= 10;
    
    // 일본 요인 (10%)
    if (japan.usdjpy > 150) liquidityScore -= 10;
    
    // 신흥국 요인 (10%)
    if (emFx.strength_index > 0) liquidityScore += 10;
    
    // 최종 신호 결정
    let finalSignal = '';
    let recommendation = '';
    
    if (liquidityScore >= 60) {
      finalSignal = '🚀 EXTREME LIQUIDITY';
      recommendation = '성장주, 신흥국, 원자재 비중 확대';
    } else if (liquidityScore >= 30) {
      finalSignal = '✅ HIGH LIQUIDITY';
      recommendation = '위험자산 비중 유지/확대';
    } else if (liquidityScore >= 0) {
      finalSignal = '⚖️ NEUTRAL';
      recommendation = '포트폴리오 균형 유지';
    } else if (liquidityScore >= -30) {
      finalSignal = '⚠️ TIGHT';
      recommendation = '현금/채권 비중 증대';
    } else {
      finalSignal = '🔴 EXTREME TIGHT';
      recommendation = '방어적 포지션, 달러/금 선호';
    }
    
    // Global_Liquidity 시트 업데이트
    const timestamp = new Date().toLocaleString('ko-KR', {timeZone: 'America/New_York'});
    
    globalSheet.getRange(2, 1, 1, 19).setValues([[
      timestamp,
      walcl.value,
      walcl_wow,
      tga.current,
      tga.week_change,
      onRrp.value,
      dxy.value,
      dxy_change,
      china.m2_growth,
      china.total_credit,
      china.fx_reserves,
      japan.usdjpy,
      japan.jgb_10y,
      japan.us_jpy_spread,
      emFx.usdkrw,
      emFx.usdbrl,
      emFx.strength_index,
      liquidityScore,
      finalSignal
    ]]);
    
    // 추천사항 업데이트
    globalSheet.getRange('T2').setValue(recommendation);
    
    // 조건부 서식
    const signalCell = globalSheet.getRange('S2');
    if (liquidityScore >= 30) {
      signalCell.setBackground('#90EE90');
    } else if (liquidityScore >= 0) {
      signalCell.setBackground('#FFFFE0');
    } else {
      signalCell.setBackground('#FFB6C1');
    }
    
    Logger.log(`✅ 글로벌 유동성 분석 완료: Score ${liquidityScore}, ${finalSignal}`);
    
    return {
      score: liquidityScore,
      signal: finalSignal,
      recommendation: recommendation,
      timestamp: new Date(),
      details: {
        us: { walcl: walcl.value, walcl_wow: walcl_wow, tga: tga, onrrp: onRrp.value },
        dxy: { level: dxy.value, change: dxy_change },
        china: china,
        japan: japan,
        em: emFx
      }
    };
    
  } catch (e) {
    Logger.log(`❌ 글로벌 유동성 분석 오류: ${e.message}`);
    SpreadsheetApp.getUi().alert(`❌ 오류: ${e.message}`);
    return { score: 0, signal: 'ERROR', timestamp: new Date() };
  }
}

function setupGlobalSheet(sheet) {
  const headers = [
    '타임스탬프', 
    'WALCL(M$)', 'WALCL WoW', 
    'TGA(M$)', 'TGA WoW', 
    'ON RRP(M$)',
    'DXY', 'DXY WoW',
    '중국 M2(%)', '중국 신용', '중국 FX',
    'USD/JPY', 'JGB 10Y', 'US-JP 스프레드',
    'USD/KRW', 'USD/BRL', 'EM 강세지수',
    '유동성 점수', '신호', '투자 권장'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1f77b4')
    .setFontColor('white');
  
  sheet.autoResizeColumns(1, headers.length);
  
  Logger.log('✅ Global_Liquidity 시트 설정 완료');
}

/** ===============================================
 * 글로벌 유동성 히스토리 기록
 * =============================================== */

function logGlobalHistory(analysis) {
  try {
    const ss = SpreadsheetApp.getActive();
    let globalHistorySheet = ss.getSheetByName(CONFIG.GLOBAL_HISTORY_SHEET);
    
    // Global_History 시트가 없으면 생성
    if (!globalHistorySheet) {
      globalHistorySheet = ss.insertSheet(CONFIG.GLOBAL_HISTORY_SHEET);
      globalHistorySheet.appendRow([
        '타임스탬프',
        'WALCL(M$)', 'WALCL WoW',
        'TGA(M$)', 'TGA WoW',
        'ON RRP(M$)',
        'DXY', 'DXY WoW',
        '중국 M2(%)', '중국 신용', '중국 FX',
        'USD/JPY', 'JGB 10Y', 'US-JP 스프레드',
        'USD/KRW', 'USD/BRL', 'EM 강세지수',
        '유동성 점수', '신호', '투자 권장'
      ]);
      globalHistorySheet.getRange(1, 1, 1, 20).setFontWeight('bold')
        .setBackground('#1f77b4')
        .setFontColor('white');
      globalHistorySheet.setFrozenRows(1);
      globalHistorySheet.setColumnWidth(1, 150);
    }
    
    // 히스토리에 추가
    globalHistorySheet.appendRow([
      analysis.timestamp,
      analysis.details.us.walcl,
      analysis.details.us.walcl_wow,
      analysis.details.us.tga.current,
      analysis.details.us.tga.week_change,
      analysis.details.us.onrrp,
      analysis.details.dxy.level,
      analysis.details.dxy.change,
      analysis.details.china.m2_growth,
      analysis.details.china.total_credit,
      analysis.details.china.fx_reserves,
      analysis.details.japan.usdjpy,
      analysis.details.japan.jgb_10y,
      analysis.details.japan.us_jpy_spread,
      analysis.details.em.usdkrw,
      analysis.details.em.usdbrl,
      analysis.details.em.strength_index,
      analysis.score,
      analysis.signal,
      analysis.recommendation
    ]);
    
    Logger.log('✅ Global_History 기록 완료');
    
  } catch (e) {
    Logger.log(`❌ Global_History 기록 오류: ${e.message}`);
  }
}

/** ===============================================
 * 8) 기본 Live_Monitor 업데이트 + History 자동 누적
 * =============================================== */

function updateLiveMonitor() {
  try {
    const ss = SpreadsheetApp.getActive();
    const liveSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    let historySheet = ss.getSheetByName(CONFIG.HISTORY_SHEET);
    
    if (!liveSheet) {
      SpreadsheetApp.getUi().alert('❌ Live_Monitor 시트를 찾을 수 없습니다');
      return;
    }
    
    // History 시트가 없으면 생성
    if (!historySheet) {
      historySheet = ss.insertSheet(CONFIG.HISTORY_SHEET);
      historySheet.appendRow([
        '타임스탬프', '날짜', 'SOFR', 'EFFR', 'IORB', 'SOFR-IORB(bp)', 
        'ON RRP', 'TGA', 'WALCL', 'WoW', 'SRF', '신호'
      ]);
      historySheet.getRange(1, 1, 1, 12).setFontWeight('bold')
        .setBackground('#1f77b4')
        .setFontColor('white');
      historySheet.setFrozenRows(1);
      historySheet.setColumnWidth(1, 150);
    }
    
    // WALCL 2주 데이터
    const walclData = getWALCLWithHistory();
    
    if (!walclData) {
      Logger.log('❌ WALCL 데이터를 가져올 수 없습니다');
      return;
    }
    
    // 다른 데이터 수집
    const sofr = getFredData(CONFIG.FRED_IDS.SOFR);
    const effr = getFredData(CONFIG.FRED_IDS.EFFR);
    const iorb = getFredData(CONFIG.FRED_IDS.IORB);
    const onRrp = getFredData(CONFIG.FRED_IDS.ON_RRP);
    const tga = getFredData(CONFIG.FRED_IDS.TGA);
    const srf = getSRFData();
    
    // 값 계산
    const timestamp = new Date();
    const date_now = walclData.current.date;
    const sofr_now = sofr.value || 0;
    const effr_now = effr.value || 0;
    const iorb_now = iorb.value || 0;
    const sofr_iorb_now = (sofr_now - iorb_now) * 100;
    const on_rrp_now = onRrp.value || 0;
    const tga_now = tga.value || 0;
    const walcl_now = walclData.current.value;
    const srf_now = srf.amount || 0;
    const wowChange = walclData.wow;
    
    // 신호 판단
    const signal = determineSignal(sofr_iorb_now, on_rrp_now, wowChange, walcl_now);
    
    // Live_Monitor 업데이트
    const dataRow = 2;
    liveSheet.getRange(dataRow, 1, 1, 11).setValues([[
      date_now, sofr_now, effr_now, iorb_now, sofr_iorb_now,
      on_rrp_now, tga_now, walcl_now, wowChange, srf_now, signal
    ]]);
    
    // History에 타임스탬프와 함께 기록 (누적)
    historySheet.appendRow([
      timestamp,
      date_now, 
      sofr_now, 
      effr_now, 
      iorb_now, 
      sofr_iorb_now,
      on_rrp_now, 
      tga_now, 
      walcl_now, 
      wowChange, 
      srf_now, 
      signal
    ]);
    
    // 메모 추가
    const now = new Date().toLocaleString('ko-KR', {timeZone: 'America/New_York'});
    liveSheet.getRange('A2').setNote(`마지막 업데이트: ${now}`);
    
    Logger.log('✅ Live_Monitor 업데이트 완료 및 History 기록');
    
    // 글로벌 유동성도 업데이트 및 히스토리 기록
    const globalAnalysis = analyzeGlobalLiquidity();
    logGlobalHistory(globalAnalysis);
    
  } catch (e) {
    Logger.log(`❌ Live_Monitor 업데이트 오류: ${e.message}`);
    SpreadsheetApp.getUi().alert(`❌ 오류: ${e.message}`);
  }
}

function getWALCLWithHistory() {
  try {
    const url = `${CONFIG.FRED_BASE}?id=WALCL`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' },
      timeout: 15000
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`FRED API 오류: ${response.getResponseCode()}`);
    }
    
    const csv = response.getContentText();
    const lines = csv.trim().split('\n');
    
    if (lines.length < 3) {
      throw new Error('WALCL 데이터가 충분하지 않습니다');
    }
    
    const lastLine = lines[lines.length - 1];
    const secondLastLine = lines[lines.length - 2];
    
    const [currentDate, currentValue] = lastLine.split(',');
    const [weekAgoDate, weekAgoValue] = secondLastLine.split(',');
    
    const result = {
      current: {
        date: currentDate.trim(),
        value: parseFloat(currentValue.trim())
      },
      weekAgo: {
        date: weekAgoDate.trim(),
        value: parseFloat(weekAgoValue.trim())
      },
      wow: parseFloat(currentValue.trim()) - parseFloat(weekAgoValue.trim())
    };
    
    return result;
    
  } catch (e) {
    Logger.log(`❌ WALCL 히스토리 수집 실패: ${e.message}`);
    return null;
  }
}

function determineSignal(sofr_iorb, on_rrp, wowChange, walcl) {
  let tightScore = 0;
  let easingScore = 0;
  let excessScore = 0;
  
  if (sofr_iorb >= 10) {
    tightScore += 2;
  } else if (sofr_iorb < 5) {
    easingScore += 1;
  }
  
  if (on_rrp >= 300000) {
    excessScore += 2;
  } else if (on_rrp >= 200000) {
    tightScore += 1;
  } else {
    easingScore += 1;
  }
  
  if (wowChange < 0) {
    tightScore += 2;
  } else if (wowChange > 0) {
    easingScore += 2;
  }
  
  if (walcl < 6500000) {
    tightScore += 1;
  }
  
  if (excessScore >= 2) {
    return '🔴 Excess';
  } else if (tightScore >= easingScore && tightScore >= 3) {
    return '⚠️ Tight';
  } else if (easingScore > tightScore) {
    return '✅ Easing';
  } else {
    return '⚖️ Neutral';
  }
}

/** ===============================================
 * 9) 알림 시스템 (설정/해제 가능) + Alert History
 * =============================================== */

function setupGlobalAlerts() {
  const ui = SpreadsheetApp.getUi();
  
  // 현재 알림 상태 확인
  const triggers = ScriptApp.getProjectTriggers();
  const alertTrigger = triggers.find(t => t.getHandlerFunction() === 'checkGlobalAlerts');
  
  if (alertTrigger) {
    // 이미 설정되어 있음
    const result = ui.alert(
      '알림 관리',
      '현재 알림이 설정되어 있습니다.\n\n해제하시겠습니까?',
      ui.ButtonSet.YES_NO
    );
    
    if (result === ui.Button.YES) {
      disableAlerts();
    }
  } else {
    // 알림 설정
    const result = ui.alert(
      '알림 설정',
      '글로벌 유동성 알림을 설정하시겠습니까?\n\n2시간마다 자동으로 체크합니다.',
      ui.ButtonSet.YES_NO
    );
    
    if (result === ui.Button.YES) {
      enableAlerts();
    }
  }
}

function enableAlerts() {
  // 기존 트리거 제거
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkGlobalAlerts') {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  // 새 트리거 생성
  ScriptApp.newTrigger('checkGlobalAlerts')
    .timeBased()
    .everyHours(2)
    .create();
  
  SpreadsheetApp.getUi().alert('✅ 알림이 설정되었습니다.\n\n2시간마다 자동 체크합니다.');
  Logger.log('✅ 글로벌 알림 설정됨');
}

function disableAlerts() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = false;
  
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'checkGlobalAlerts') {
      ScriptApp.deleteTrigger(t);
      removed = true;
    }
  });
  
  if (removed) {
    SpreadsheetApp.getUi().alert('✅ 알림이 해제되었습니다.');
    Logger.log('✅ 알림 해제됨');
  } else {
    SpreadsheetApp.getUi().alert('ℹ️ 설정된 알림이 없습니다.');
  }
}

function checkGlobalAlerts() {
  try {
    const analysis = analyzeGlobalLiquidity();
    const alerts = [];
    
    // 극단적 신호
    if (analysis.score >= 60) {
      alerts.push({
        level: '🚀 OPPORTUNITY',
        message: '글로벌 유동성 급증',
        action: analysis.recommendation
      });
    } else if (analysis.score <= -30) {
      alerts.push({
        level: '🔴 WARNING',
        message: '글로벌 유동성 급감',
        action: analysis.recommendation
      });
    }
    
    // 중국 리스크
    if (analysis.details.china.m2_growth < 7) {
      alerts.push({
        level: '🇨🇳 CHINA RISK',
        message: '중국 유동성 경색',
        action: '신흥국/원자재 노출 축소'
      });
    }
    
    // 엔캐리 리스크
    if (analysis.details.japan.usdjpy > 155) {
      alerts.push({
        level: '🇯🇵 YEN RISK',
        message: '엔캐리 언와인드 임박',
        action: '변동성 헤지'
      });
    }
    
    // 달러 급변
    if (Math.abs(analysis.details.dxy.change) > 2) {
      alerts.push({
        level: '💵 DXY MOVE',
        message: `달러 ${analysis.details.dxy.change > 0 ? '급등' : '급락'} (${analysis.details.dxy.change.toFixed(2)})`,
        action: analysis.details.dxy.change > 0 ? 'Risk-OFF 준비' : 'Risk-ON 기회'
      });
    }
    
    if (alerts.length > 0) {
      sendGlobalAlert(alerts, analysis);
      logAlertHistory(alerts, analysis);
    }
    
  } catch (e) {
    Logger.log(`❌ 알림 체크 오류: ${e.message}`);
  }
}

/** ===============================================
 * Alert History 기록
 * =============================================== */

function logAlertHistory(alerts, analysis) {
  try {
    const ss = SpreadsheetApp.getActive();
    let alertHistorySheet = ss.getSheetByName(CONFIG.ALERT_HISTORY_SHEET);
    
    // Alert_History 시트가 없으면 생성
    if (!alertHistorySheet) {
      alertHistorySheet = ss.insertSheet(CONFIG.ALERT_HISTORY_SHEET);
      alertHistorySheet.appendRow([
        '타임스탬프',
        '유동성 점수',
        '신호',
        '알림 레벨',
        '메시지',
        '권장 조치'
      ]);
      alertHistorySheet.getRange(1, 1, 1, 6).setFontWeight('bold')
        .setBackground('#e74c3c')
        .setFontColor('white');
      alertHistorySheet.setFrozenRows(1);
      alertHistorySheet.setColumnWidth(1, 150);
      alertHistorySheet.setColumnWidth(4, 150);
      alertHistorySheet.setColumnWidth(5, 200);
      alertHistorySheet.setColumnWidth(6, 200);
    }
    
    const timestamp = new Date();
    
    // 각 알림을 별도 행으로 기록
    alerts.forEach(alert => {
      alertHistorySheet.appendRow([
        timestamp,
        analysis.score,
        analysis.signal,
        alert.level,
        alert.message,
        alert.action
      ]);
      
      // 마지막 행 서식 설정
      const lastRow = alertHistorySheet.getLastRow();
      
      // 알림 레벨에 따른 배경색
      if (alert.level.includes('OPPORTUNITY') || alert.level.includes('🚀')) {
        alertHistorySheet.getRange(lastRow, 1, 1, 6).setBackground('#d5f4e6');
      } else if (alert.level.includes('WARNING') || alert.level.includes('🔴')) {
        alertHistorySheet.getRange(lastRow, 1, 1, 6).setBackground('#fadbd8');
      } else if (alert.level.includes('RISK') || alert.level.includes('⚠️')) {
        alertHistorySheet.getRange(lastRow, 1, 1, 6).setBackground('#fff3cd');
      }
    });
    
    Logger.log(`✅ Alert_History 기록 완료: ${alerts.length}개 알림`);
    
  } catch (e) {
    Logger.log(`❌ Alert_History 기록 오류: ${e.message}`);
  }
}

function sendGlobalAlert(alerts, analysis) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const timestamp = new Date().toLocaleString('ko-KR');
    
    let emailBody = `
      <div style="font-family: Arial; background-color: #f5f5f5; padding: 20px;">
        <h2 style="color: #1f77b4;">🌐 글로벌 유동성 알림</h2>
        <p><strong>시간:</strong> ${timestamp}</p>
        <p><strong>유동성 점수:</strong> ${analysis.score} / 100</p>
        <p><strong>신호:</strong> ${analysis.signal}</p>
        
        <h3>📊 주요 지표</h3>
        <table style="border-collapse: collapse; width: 100%; background: white;">
          <tr>
            <td style="border: 1px solid #ddd; padding: 8px;"><strong>DXY:</strong></td>
            <td style="border: 1px solid #ddd; padding: 8px;">${analysis.details.dxy.level.toFixed(2)} (${analysis.details.dxy.change > 0 ? '+' : ''}${analysis.details.dxy.change.toFixed(2)})</td>
          </tr>
          <tr>
            <td style="border: 1px solid #ddd; padding: 8px;"><strong>WALCL WoW:</strong></td>
            <td style="border: 1px solid #ddd; padding: 8px;">${analysis.details.us.walcl_wow.toFixed(0)}M$</td>
          </tr>
          <tr>
            <td style="border: 1px solid #ddd; padding: 8px;"><strong>중국 M2:</strong></td>
            <td style="border: 1px solid #ddd; padding: 8px;">${analysis.details.china.m2_growth.toFixed(1)}%</td>
          </tr>
          <tr>
            <td style="border: 1px solid #ddd; padding: 8px;"><strong>USD/JPY:</strong></td>
            <td style="border: 1px solid #ddd; padding: 8px;">${analysis.details.japan.usdjpy.toFixed(2)}</td>
          </tr>
        </table>
        
        <h3>🚨 알림 내역</h3>
        <table style="border-collapse: collapse; width: 100%; margin: 20px 0;">
          <tr style="background-color: #d3d3d3;">
            <th style="border: 1px solid #999; padding: 10px;">레벨</th>
            <th style="border: 1px solid #999; padding: 10px;">메시지</th>
            <th style="border: 1px solid #999; padding: 10px;">권장 조치</th>
          </tr>
    `;
    
    alerts.forEach(a => {
      emailBody += `
        <tr style="background-color: white;">
          <td style="border: 1px solid #999; padding: 10px;"><strong>${a.level}</strong></td>
          <td style="border: 1px solid #999; padding: 10px;">${a.message}</td>
          <td style="border: 1px solid #999; padding: 10px;"><em>${a.action}</em></td>
        </tr>
      `;
    });
    
    emailBody += `
        </table>
        <hr style="border: 1px solid #ddd;">
        <p><a href="${SpreadsheetApp.getActive().getUrl()}" style="background-color: #1f77b4; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">📊 스프레드시트 보기</a></p>
      </div>
    `;
    
    GmailApp.sendEmail(userEmail, '🌐 글로벌 유동성 알림', '', {
      htmlBody: emailBody
    });
    
    Logger.log(`✉️ 글로벌 알림 발송: ${userEmail}`);
    
  } catch (e) {
    Logger.log(`❌ 이메일 발송 오류: ${e.message}`);
  }
}

/** ===============================================
 * 10) 개별 체크 함수들
 * =============================================== */

function checkChinaLiquidity() {
  const china = getChinaLiquidity();
  SpreadsheetApp.getUi().alert(
    `🇨🇳 중국 유동성 현황\n\n` +
    `M2 성장률: ${china.m2_growth.toFixed(1)}% YoY\n` +
    `총 신용: ${(china.total_credit/1000).toFixed(0)}조 위안\n` +
    `외환보유고: ${(china.fx_reserves/1000).toFixed(1)}조 달러\n\n` +
    `신호: ${china.liquidity_signal}`
  );
}

function checkJapanRisk() {
  const japan = getJapanLiquidity();
  SpreadsheetApp.getUi().alert(
    `🇯🇵 일본/엔캐리 리스크\n\n` +
    `USD/JPY: ${japan.usdjpy.toFixed(2)}\n` +
    `일본 10Y: ${japan.jgb_10y.toFixed(2)}%\n` +
    `미-일 금리차: ${japan.us_jpy_spread.toFixed(2)}%\n\n` +
    `리스크 평가: ${japan.carry_risk}`
  );
}

function checkTGADetail() {
  const tga = getTGAAnalysis();
  SpreadsheetApp.getUi().alert(
    `💵 TGA (재무부 계좌) 분석\n\n` +
    `현재 잔고: $${(tga.current/1000).toFixed(0)}B\n` +
    `주간 변화: $${(tga.week_change/1000).toFixed(0)}B\n` +
    `월간 변화: $${(tga.month_change/1000).toFixed(0)}B\n\n` +
    `유동성 영향: ${tga.liquidity_impact}\n` +
    `부채한도 리스크: ${tga.debt_ceiling_risk}`
  );
}

function checkDXYTrend() {
  const dxy = getFredData(CONFIG.GLOBAL_FRED_IDS.DXY);
  const dxy_1w = getFredDataHistorical(CONFIG.GLOBAL_FRED_IDS.DXY, 7);
  const dxy_1m = getFredDataHistorical(CONFIG.GLOBAL_FRED_IDS.DXY, 30);
  
  const weekChange = (dxy.value - dxy_1w.value).toFixed(2);
  const monthChange = (dxy.value - dxy_1m.value).toFixed(2);
  
  SpreadsheetApp.getUi().alert(
    `💵 달러 인덱스 (DXY) 추세\n\n` +
    `현재: ${dxy.value.toFixed(2)}\n` +
    `주간 변화: ${weekChange > 0 ? '+' : ''}${weekChange}\n` +
    `월간 변화: ${monthChange > 0 ? '+' : ''}${monthChange}\n\n` +
    `${Math.abs(weekChange) > 2 ? '⚠️ 급격한 변동 주의' : '✅ 정상 범위'}`
  );
}

/** ===============================================
 * 11) 대시보드 및 리포트
 * =============================================== */

function createGlobalDashboard() {
  try {
    const analysis = analyzeGlobalLiquidity();
    const ui = SpreadsheetApp.getUi();
    
    const html = HtmlService.createHtmlOutput(`
      <style>
        body { font-family: Arial; padding: 20px; }
        h2 { color: #1f77b4; }
        .score { font-size: 48px; font-weight: bold; margin: 20px 0; }
        .signal { font-size: 24px; margin: 15px 0; }
        .positive { color: green; }
        .negative { color: red; }
        .neutral { color: orange; }
        table { border-collapse: collapse; width: 100%; margin: 20px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .section { margin: 20px 0; padding: 15px; background: #f9f9f9; border-radius: 5px; }
      </style>
      
      <h2>🌐 글로벌 유동성 대시보드</h2>
      
      <div class="section">
        <h3>종합 점수</h3>
        <div class="score ${analysis.score >= 30 ? 'positive' : analysis.score <= -30 ? 'negative' : 'neutral'}">
          ${analysis.score} / 100
        </div>
        <div class="signal">${analysis.signal}</div>
        <p><strong>투자 권장:</strong> ${analysis.recommendation}</p>
      </div>
      
      <div class="section">
        <h3>주요 지표</h3>
        <table>
          <tr>
            <th>지표</th>
            <th>현재값</th>
            <th>변화</th>
            <th>신호</th>
          </tr>
          <tr>
            <td>DXY (달러지수)</td>
            <td>${analysis.details.dxy.level.toFixed(2)}</td>
            <td>${analysis.details.dxy.change > 0 ? '+' : ''}${analysis.details.dxy.change.toFixed(2)}</td>
            <td>${analysis.details.dxy.change < -1 ? '✅' : analysis.details.dxy.change > 1 ? '⚠️' : '⚖️'}</td>
          </tr>
          <tr>
            <td>WALCL (연준자산)</td>
            <td>${(analysis.details.us.walcl/1000000).toFixed(2)}T</td>
            <td>${analysis.details.us.walcl_wow > 0 ? '+' : ''}${(analysis.details.us.walcl_wow/1000).toFixed(1)}B</td>
            <td>${analysis.details.us.walcl_wow > 0 ? '✅' : '⚠️'}</td>
          </tr>
          <tr>
            <td>중국 M2 성장률</td>
            <td>${analysis.details.china.m2_growth.toFixed(1)}%</td>
            <td>YoY</td>
            <td>${analysis.details.china.liquidity_signal}</td>
          </tr>
          <tr>
            <td>USD/JPY</td>
            <td>${analysis.details.japan.usdjpy.toFixed(2)}</td>
            <td>금리차 ${analysis.details.japan.us_jpy_spread.toFixed(2)}%</td>
            <td>${analysis.details.japan.carry_risk}</td>
          </tr>
          <tr>
            <td>EM 통화</td>
            <td>지수 ${analysis.details.em.strength_index.toFixed(2)}</td>
            <td>-</td>
            <td>${analysis.details.em.signal}</td>
          </tr>
        </table>
      </div>
      
      <div class="section">
        <h3>리스크 요인</h3>
        <ul>
          ${analysis.details.china.m2_growth < 8 ? '<li>⚠️ 중국 유동성 둔화</li>' : ''}
          ${analysis.details.japan.usdjpy > 150 ? '<li>⚠️ 엔캐리 언와인드 리스크</li>' : ''}
          ${Math.abs(analysis.details.dxy.change) > 2 ? '<li>⚠️ 달러 급변동</li>' : ''}
          ${analysis.details.us.tga.current < 200000 ? '<li>⚠️ TGA 잔고 부족</li>' : ''}
        </ul>
      </div>
      
      <p style="text-align: center; margin-top: 30px;">
        <em>생성 시간: ${new Date().toLocaleString('ko-KR')}</em>
      </p>
    `).setWidth(600).setHeight(800);
    
    ui.showModalDialog(html, '글로벌 유동성 대시보드');
    
  } catch (e) {
    Logger.log(`❌ 대시보드 생성 오류: ${e.message}`);
    SpreadsheetApp.getUi().alert(`❌ 오류: ${e.message}`);
  }
}

/** ===============================================
 * 12) 메뉴 설정
 * =============================================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Global Liquidity')
    .addItem('🔄 전체 업데이트', 'updateLiveMonitor')
    .addItem('🌐 글로벌 유동성 분석', 'analyzeGlobalLiquidity')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🔍 개별 체크')
      .addItem('🇨🇳 중국 유동성', 'checkChinaLiquidity')
      .addItem('🇯🇵 엔캐리 리스크', 'checkJapanRisk')
      .addItem('💵 TGA 분석', 'checkTGADetail')
      .addItem('📈 DXY 추세', 'checkDXYTrend'))
    .addSeparator()
    .addItem('📊 종합 대시보드', 'createGlobalDashboard')
    .addItem('🔔 알림 설정/해제', 'setupGlobalAlerts')
    .addItem('⏰ 일일 자동갱신', 'createDailyTrigger')
    .addSeparator()
    .addItem('📋 캐시 초기화', 'clearAllCache')
    .addItem('❓ 도움말', 'showHelp')
    .addToUi();
}

/** ===============================================
 * 13) 유틸리티 함수
 * =============================================== */

function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'updateLiveMonitor') {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  ScriptApp.newTrigger('updateLiveMonitor')
    .timeBased()
    .everyDays(1)
    .atHour(17)
    .create();
  
  SpreadsheetApp.getUi().alert('✅ 일일 자동 업데이트가 설정되었습니다.\n\n매일 오후 5시(NY시간)에 실행됩니다.');
}

function clearAllCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll();
  SpreadsheetApp.getUi().alert('✅ 모든 캐시가 초기화되었습니다.');
}

function showHelp() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial; font-size: 12px; padding: 15px; }
      h3 { color: #1f77b4; margin-top: 15px; }
      code { background: #f5f5f5; padding: 3px 6px; border-radius: 3px; }
      li { margin: 8px 0; }
    </style>
    
    <h2>📊 Global Liquidity Monitor 도움말</h2>
    
    <h3>주요 기능</h3>
    <ul>
      <li><strong>전체 업데이트:</strong> 미국 + 글로벌 데이터 갱신 및 히스토리 누적</li>
      <li><strong>글로벌 분석:</strong> 종합 유동성 점수 계산</li>
      <li><strong>개별 체크:</strong> 중국, 일본, TGA, DXY 상세 분석</li>
      <li><strong>알림 설정:</strong> 2시간마다 자동 체크 (해제 가능)</li>
    </ul>
    
    <h3>히스토리 기록</h3>
    <ul>
      <li><strong>History:</strong> 미국 유동성 지표 타임시리즈</li>
      <li><strong>Global_History:</strong> 글로벌 유동성 분석 타임시리즈</li>
      <li><strong>Alert_History:</strong> 발생한 알림 전체 기록</li>
    </ul>
    
    <h3>유동성 점수</h3>
    <ul>
      <li><strong>60점 이상:</strong> 극도의 유동성 (Risk-ON)</li>
      <li><strong>30-60점:</strong> 높은 유동성</li>
      <li><strong>0-30점:</strong> 중립</li>
      <li><strong>-30-0점:</strong> 긴축</li>
      <li><strong>-30점 이하:</strong> 극도의 긴축 (Risk-OFF)</li>
    </ul>
    
    <h3>가중치</h3>
    <ul>
      <li>미국 요인: 40%</li>
      <li>달러 지수: 20%</li>
      <li>중국: 20%</li>
      <li>일본: 10%</li>
      <li>신흥국: 10%</li>
    </ul>
    
    <h3>시트 구성</h3>
    <ul>
      <li><strong>Live_Monitor:</strong> 미국 지표 최신값</li>
      <li><strong>Global_Liquidity:</strong> 글로벌 지표 최신값</li>
      <li><strong>History:</strong> 미국 지표 히스토리</li>
      <li><strong>Global_History:</strong> 글로벌 지표 히스토리</li>
      <li><strong>Alert_History:</strong> 알림 발생 기록</li>
    </ul>
  `).setWidth(500).setHeight(650);
  
  ui.showModelessDialog(html, '도움말');
}

/** ===============================================
 * 14) 테스트 함수
 * =============================================== */

function testAllSystems() {
  Logger.log('=== 전체 시스템 테스트 시작 ===');
  
  // 1. FRED 데이터
  Logger.log('\n--- FRED 데이터 테스트 ---');
  Object.entries(CONFIG.FRED_IDS).forEach(([name, id]) => {
    const data = getFredData(id, false);
    Logger.log(`${name}: ${data.value || 'ERROR'}`);
  });
  
  // 2. 글로벌 데이터
  Logger.log('\n--- 글로벌 데이터 테스트 ---');
  const china = getChinaLiquidity();
  Logger.log(`중국 M2: ${china.m2_growth}%`);
  
  const japan = getJapanLiquidity();
  Logger.log(`USD/JPY: ${japan.usdjpy}`);
  
  // 3. 종합 분석
  Logger.log('\n--- 종합 분석 테스트 ---');
  const analysis = analyzeGlobalLiquidity();
  Logger.log(`유동성 점수: ${analysis.score}`);
  Logger.log(`신호: ${analysis.signal}`);
  
  // 4. 히스토리 기록 테스트
  Logger.log('\n--- 히스토리 기록 테스트 ---');
  logGlobalHistory(analysis);
  
  Logger.log('\n=== 테스트 완료 ===');
}