/****************************************************
 * Global Liquidity Monitor v4.1 - DXY+CNY 조합 로직 + 가중치 조정
 *
 * 주요 기능:
 * 1. 미국 유동성 모니터링 (WALCL, TGA, ON RRP)
 * 2. 글로벌 유동성 추적 (USD/CNY, USD/JPY, DXY)
 * 3. 신흥국 통화 모니터링
 * 4. 종합 유동성 점수 (7단계 신호, 5단계 세분화)
 * 5. 알림 설정/해제 기능 (±50, ±80 임계값)
 * 6. 히스토리 자동 누적 (History, Global_History, Alert_History)
 * 7. 점수 계산 가이드 시트 자동 생성
 *
 * v4.0 신규 기능 (2026-01):
 * 8. 동적 가중치 시스템 (시장 국면별 가중치 자동 조정)
 * 9. Fed 정책 추적 강화 (QT/QE 상태, 금리 사이클, Reserve Balances)
 * 10. 맥락 기반 DXY 스코어링 (Fed 완화기 DXY 하락 = 긍정)
 * 11. 추세 기반 알림 시스템 (모멘텀/추세 변화 감지)
 *
 * v4.1 신규 기능 (2026-01):
 * 12. DXY + CNY 조합 로직 (중국 M2 대신 실시간 USD/CNY 사용)
 * 13. 가중치 재조정: US 45%, DXY 25%, China 5%, Japan 10%, EM 15%
 * 14. 중국 데이터 지연 문제 해결 (9개월 → 실시간)
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
    WALCL: 'WALCL',
    // v4.0: Fed 정책 추적용 추가 지표
    FED_FUNDS_UPPER: 'DFEDTARU',   // Fed Funds Target Upper Bound
    FED_FUNDS_LOWER: 'DFEDTARL',   // Fed Funds Target Lower Bound
    RESERVE_BALANCES: 'WRESBAL'    // Reserve Balances with Fed
  },
  
  // 글로벌 지표
  GLOBAL_FRED_IDS: {
    // 달러 인덱스
    DXY: 'DTWEXBGS',
    
    // 중국 지표 (실시간 환율 기반으로 변경 - 신용 데이터는 9개월 지연)
    USDCNY: 'DEXCHUS',                // USD/CNY 환율 (실시간)
    CHINA_RESERVES: 'TRESEGCNM052N',  // 외환보유고 (참고용)
    
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

/**
 * v4.1: 직전 발표 데이터 조회 (주간 데이터용)
 * - WALCL, TGA 등 주간 데이터의 WoW 계산에 사용
 * - 실행 날짜와 무관하게 항상 직전 발표값 반환
 */
function getPreviousRelease(fredId) {
  try {
    const url = `${CONFIG.FRED_BASE}?id=${fredId}`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      timeout: 15000
    });

    if (response.getResponseCode() !== 200) {
      return { value: 0, date: null };
    }

    const csv = response.getContentText();
    const lines = csv.trim().split('\n');

    // 마지막에서 두 번째 데이터 (직전 발표)
    if (lines.length >= 3) {
      const [date, value] = lines[lines.length - 2].split(',');
      return {
        date: date.trim(),
        value: parseFloat(value.trim())
      };
    }

    return { value: 0, date: null };

  } catch (e) {
    Logger.log(`❌ Previous Release 조회 오류 [${fredId}]: ${e.message}`);
    return { value: 0, date: null };
  }
}

/**
 * FRED에서 날짜 범위로 데이터 가져오기
 * @param {string} fredId - FRED 시리즈 ID
 * @param {Date} startDate - 시작 날짜
 * @returns {Object} 날짜를 키로 하는 데이터 맵
 */
function getFredDataRange(fredId, startDate) {
  try {
    const url = `${CONFIG.FRED_BASE}?id=${fredId}`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      timeout: 15000
    });

    if (response.getResponseCode() !== 200) {
      Logger.log(`❌ FRED API 오류 [${fredId}]: ${response.getResponseCode()}`);
      return {};
    }

    const csv = response.getContentText();
    const lines = csv.trim().split('\n');

    if (lines.length < 2) {
      Logger.log(`❌ FRED 데이터가 없음 [${fredId}]`);
      return {};
    }

    const dataMap = {};
    const startDateStr = Utilities.formatDate(startDate, 'GMT', 'yyyy-MM-dd');

    // 첫 번째 줄은 헤더이므로 건너뛰기
    for (let i = 1; i < lines.length; i++) {
      const [dateStr, valueStr] = lines[i].split(',');
      const date = dateStr.trim();
      const value = valueStr.trim();

      // 시작 날짜 이후 데이터만 포함
      if (date >= startDateStr && value !== '.' && value !== '') {
        dataMap[date] = parseFloat(value);
      }
    }

    Logger.log(`✅ ${fredId}: ${Object.keys(dataMap).length}개 데이터 포인트 로드됨`);
    return dataMap;

  } catch (e) {
    Logger.log(`❌ FRED Range 데이터 오류 [${fredId}]: ${e.message}`);
    return {};
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

function getChinaLiquidity(dxyChange) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'CHINA_LIQUIDITY';

    // 캐시는 dxyChange 없이 기본 데이터만 저장
    const cached = cache.get(cacheKey);
    let baseData;

    if (cached) {
      baseData = JSON.parse(cached);
    } else {
      // USD/CNY 환율 (일간 데이터 - 7일 전 비교)
      const usdcny = getFredData(CONFIG.GLOBAL_FRED_IDS.USDCNY, false);
      const usdcny_1w = getFredDataHistorical(CONFIG.GLOBAL_FRED_IDS.USDCNY, 7);
      const reserves = getFredData(CONFIG.GLOBAL_FRED_IDS.CHINA_RESERVES, false);

      // CNY 변화율 계산 (양수 = 위안 약세, 음수 = 위안 강세)
      let cnyChange = 0;
      if (usdcny.value && usdcny_1w.value && usdcny_1w.value > 0) {
        cnyChange = ((usdcny.value - usdcny_1w.value) / usdcny_1w.value) * 100;
      }

      baseData = {
        usdcny: usdcny.value || 0,
        usdcny_change: cnyChange,
        fx_reserves: reserves.value || 0,
        timestamp: new Date().getTime()
      };

      cache.put(cacheKey, JSON.stringify(baseData), 3600);
    }

    // DXY와 CNY 조합 로직으로 신호 결정
    const signal = determineChinaSignal(dxyChange, baseData.usdcny_change || 0);

    const result = {
      ...baseData,
      liquidity_signal: signal.signal,
      score: signal.score,
      interpretation: signal.interpretation
    };

    const usdcnyStr = baseData.usdcny ? baseData.usdcny.toFixed(4) : 'N/A';
    const changeStr = baseData.usdcny_change ? baseData.usdcny_change.toFixed(2) : '0.00';
    Logger.log(`✅ 중국 유동성: USD/CNY ${usdcnyStr}, 변화 ${changeStr}%, ${signal.interpretation}`);

    return result;

  } catch (e) {
    Logger.log(`❌ 중국 데이터 오류: ${e.message}`);
    return { usdcny: 0, usdcny_change: 0, fx_reserves: 0, liquidity_signal: 'NO DATA', score: 0 };
  }
}

/**
 * DXY + CNY 조합 로직
 * - DXY↓ + CNY 강세 → 확실한 긍정 (+10)
 * - DXY↑ + CNY 약세 → 확실한 부정 (-10)
 * - DXY↓ + CNY 약세 → 중국 자체 문제 (-5)
 * - DXY↑ + CNY 강세 → 중국 선방 (+5)
 */
function determineChinaSignal(dxyChange, cnyChange) {
  // dxyChange: 양수 = 달러 강세, 음수 = 달러 약세
  // cnyChange: 양수 = 위안 약세 (USD/CNY 상승), 음수 = 위안 강세 (USD/CNY 하락)

  const dxyWeak = dxyChange < -0.5;   // DXY 0.5% 이상 하락
  const dxyStrong = dxyChange > 0.5;  // DXY 0.5% 이상 상승
  const cnyStrong = cnyChange < -0.3; // CNY 0.3% 이상 강세
  const cnyWeak = cnyChange > 0.3;    // CNY 0.3% 이상 약세

  if (dxyWeak && cnyStrong) {
    return { signal: '🚀 위안 강세 + 달러 약세', score: 10, interpretation: '글로벌 유동성 최적' };
  } else if (dxyStrong && cnyWeak) {
    return { signal: '🔴 위안 약세 + 달러 강세', score: -10, interpretation: '글로벌 유동성 악화' };
  } else if (dxyWeak && cnyWeak) {
    return { signal: '⚠️ 위안 약세 (중국 문제)', score: -5, interpretation: '중국 자체 리스크' };
  } else if (dxyStrong && cnyStrong) {
    return { signal: '✅ 위안 선방', score: 5, interpretation: '중국 상대적 강세' };
  } else {
    return { signal: '⚖️ 중립', score: 0, interpretation: '뚜렷한 방향 없음' };
  }
}

/** ===============================================
 * 3-B) Google Finance 실시간 환율 조회
 * =============================================== */

/**
 * Google Finance를 사용하여 실시간 환율 조회 (~15분 지연)
 * @param {string} pair - 통화쌍 (예: 'USDKRW', 'USDJPY', 'USDBRL')
 * @returns {number} 환율 값
 */
function getGoogleFinanceFX(pair) {
  try {
    const ss = SpreadsheetApp.getActive();
    const tempSheet = ss.getSheetByName('_temp_fx') || ss.insertSheet('_temp_fx');

    // GOOGLEFINANCE 수식으로 실시간 환율 조회
    const cell = tempSheet.getRange('A1');
    cell.setFormula(`=GOOGLEFINANCE("CURRENCY:${pair}")`);
    SpreadsheetApp.flush();

    const value = cell.getValue();

    // 임시 시트 숨기기
    tempSheet.hideSheet();

    if (typeof value === 'number' && value > 0) {
      Logger.log(`✅ Google Finance ${pair}: ${value}`);
      return value;
    }

    throw new Error('Invalid value from Google Finance');

  } catch (e) {
    Logger.log(`⚠️ Google Finance ${pair} 실패, FRED fallback: ${e.message}`);
    return null; // fallback을 위해 null 반환
  }
}

/**
 * Google Finance로 환율 조회, 실패시 FRED fallback
 * @param {string} pair - 통화쌍 (예: 'USDKRW')
 * @param {string} fredId - FRED 시리즈 ID
 * @returns {number} 환율 값
 */
function getFXRate(pair, fredId) {
  // 먼저 Google Finance 시도
  const gfValue = getGoogleFinanceFX(pair);
  if (gfValue !== null) {
    return gfValue;
  }

  // 실패시 FRED fallback
  const fredData = getFredData(fredId, false);
  return fredData.value || 0;
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

    // Google Finance로 실시간 USD/JPY 조회 (FRED fallback)
    const usdjpy_value = getFXRate('USDJPY', CONFIG.GLOBAL_FRED_IDS.USDJPY);
    const jgb10y = getFredData(CONFIG.GLOBAL_FRED_IDS.JGB_10Y, false);
    const us10y = getFredData('DGS10', false);

    const result = {
      usdjpy: usdjpy_value,
      jgb_10y: jgb10y.value || 0,
      us_jpy_spread: (us10y.value || 0) - (jgb10y.value || 0),
      carry_risk: determineCarryRisk(usdjpy_value, (us10y.value || 0) - (jgb10y.value || 0)),
      timestamp: new Date().getTime()
    };

    // 캐시 시간을 5분으로 단축 (실시간 데이터이므로)
    cache.put(cacheKey, JSON.stringify(result), 300);
    Logger.log(`✅ 일본 데이터: USDJPY ${result.usdjpy} (Google Finance)`);

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
    const tga_prev = getPreviousRelease(CONFIG.FRED_IDS.TGA);  // v4.1: 직전 발표값
    const tga_1m = getFredDataHistorical(CONFIG.FRED_IDS.TGA, 30);

    const current = tga.value || 0;
    const prevRelease = tga_prev.value || current;  // 직전 주 발표값
    const monthAgo = tga_1m.value || current;

    const weekChange = current - prevRelease;  // v4.1: 직전 발표값 기준
    const monthChange = current - monthAgo;

    Logger.log(`✅ TGA: 현재 ${current}, 직전 ${prevRelease}, WoW ${weekChange}`);

    return {
      current: current,
      week_change: weekChange,
      month_change: monthChange,
      liquidity_impact: determineTGAImpact(weekChange, monthChange),
      debt_ceiling_risk: checkDebtCeilingRisk(current)
    };

  } catch (e) {
    Logger.log(`❌ TGA 분석 오류: ${e.message}`);
    return { current: 0, week_change: 0, liquidity_impact: 'NO DATA' };
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
 * 5-B) Fed 정책 추적 (v4.0 신규)
 * =============================================== */

/**
 * QT/QE 상태 감지
 * WALCL 월간 변화로 QT/QE 상태 판단
 * @returns {Object} QT/QE 상태 및 점수
 */
function getQTStatus() {
  try {
    const walcl = getFredData(CONFIG.FRED_IDS.WALCL, false);
    const walcl_1m = getFredDataHistorical(CONFIG.FRED_IDS.WALCL, 30);

    const current = walcl.value || 0;
    const monthAgo = walcl_1m.value || current;
    const monthChange = current - monthAgo;

    let status = 'NEUTRAL';
    let score = 0;
    let signal = '⚖️ 중립';

    if (monthChange > 50000) {        // $50B 이상 증가 (QE/RMP)
      status = 'QE';
      score = 30;
      signal = '🚀 QE/RMP 진행';
    } else if (monthChange > 10000) {  // $10B 이상 증가
      status = 'MILD_QE';
      score = 15;
      signal = '✅ 완화적';
    } else if (monthChange < -50000) { // $50B 이상 감소 (적극적 QT)
      status = 'AGGRESSIVE_QT';
      score = -25;
      signal = '🔴 적극적 QT';
    } else if (monthChange < -10000) { // $10B 이상 감소
      status = 'QT';
      score = -15;
      signal = '⚠️ QT 진행';
    }

    Logger.log(`✅ QT Status: ${status}, 월간변화: ${monthChange}M$`);

    return {
      status: status,
      score: score,
      signal: signal,
      monthChange: monthChange,
      current: current
    };

  } catch (e) {
    Logger.log(`❌ QT Status 오류: ${e.message}`);
    return { status: 'UNKNOWN', score: 0, signal: 'NO DATA', monthChange: 0 };
  }
}

/**
 * 금리 사이클 감지
 * Fed 기준금리 변화로 완화/긴축 사이클 판단
 * @returns {Object} 금리 사이클 상태 및 점수
 */
function getRateCycleStatus() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'RATE_CYCLE_STATUS';

    const cached = cache.get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }

    // Fed Funds Target Rate (Upper Bound)
    const rateNow = getFredData(CONFIG.FRED_IDS.FED_FUNDS_UPPER, false);
    const rate1m = getFredDataHistorical(CONFIG.FRED_IDS.FED_FUNDS_UPPER, 30);
    const rate3m = getFredDataHistorical(CONFIG.FRED_IDS.FED_FUNDS_UPPER, 90);

    const current = rateNow.value || 0;
    const monthAgo = rate1m.value || current;
    const quarterAgo = rate3m.value || current;

    // 변화량 (bp 단위로 환산)
    const change1m = (current - monthAgo) * 100;
    const change3m = (current - quarterAgo) * 100;

    let cycle = 'STABLE';
    let score = 0;
    let signal = '⚖️ 금리 동결';

    if (change3m <= -75) {          // 3개월간 75bp 이상 인하
      cycle = 'AGGRESSIVE_CUTTING';
      score = 30;
      signal = '🚀 적극적 완화';
    } else if (change3m <= -25) {   // 3개월간 25bp 이상 인하
      cycle = 'CUTTING';
      score = 20;
      signal = '✅ 완화 사이클';
    } else if (change1m < 0) {      // 1개월간 인하
      cycle = 'MILD_CUTTING';
      score = 10;
      signal = '✅ 완화 시작';
    } else if (change3m >= 75) {    // 3개월간 75bp 이상 인상
      cycle = 'AGGRESSIVE_HIKING';
      score = -25;
      signal = '🔴 적극적 긴축';
    } else if (change3m >= 25) {    // 3개월간 25bp 이상 인상
      cycle = 'HIKING';
      score = -15;
      signal = '⚠️ 긴축 사이클';
    }

    const result = {
      cycle: cycle,
      score: score,
      signal: signal,
      currentRate: current,
      change1m: change1m,
      change3m: change3m
    };

    cache.put(cacheKey, JSON.stringify(result), 3600); // 1시간 캐시
    Logger.log(`✅ Rate Cycle: ${cycle}, 현재금리: ${current}%, 3개월변화: ${change3m}bp`);

    return result;

  } catch (e) {
    Logger.log(`❌ Rate Cycle 오류: ${e.message}`);
    return { cycle: 'UNKNOWN', score: 0, signal: 'NO DATA', currentRate: 0 };
  }
}

/**
 * Reserve Balances 상태 확인
 * 은행 준비금 수준으로 유동성 상태 판단
 * @returns {Object} Reserve 상태 및 점수
 */
function getReserveBalancesStatus() {
  try {
    const reserves = getFredData(CONFIG.FRED_IDS.RESERVE_BALANCES, false);
    const reserves_1m = getFredDataHistorical(CONFIG.FRED_IDS.RESERVE_BALANCES, 30);

    const current = reserves.value || 0;
    const monthAgo = reserves_1m.value || current;
    const monthChange = current - monthAgo;

    // 임계값 (millions)
    const AMPLE_THRESHOLD = 2900000;    // $2.9T (ample threshold)
    const ABUNDANT_THRESHOLD = 3400000; // $3.4T (abundant)
    const SCARCE_THRESHOLD = 2500000;   // $2.5T (scarce)

    let status = 'AMPLE';
    let score = 0;
    let signal = '⚖️ Ample 범위';

    if (current < SCARCE_THRESHOLD) {
      status = 'SCARCE';
      score = -20;
      signal = '🔴 준비금 부족';
    } else if (current < AMPLE_THRESHOLD) {
      status = 'APPROACHING_SCARCE';
      score = -10;
      signal = '⚠️ 준비금 부족 우려';
    } else if (current > ABUNDANT_THRESHOLD) {
      status = 'ABUNDANT';
      score = 10;
      signal = '✅ 충분한 유동성';
    }

    // 추세 보정
    if (monthChange < -100000) {       // 월 $100B 이상 감소
      score -= 5;
      signal += ' (급감 중)';
    } else if (monthChange > 100000) { // 월 $100B 이상 증가
      score += 5;
      signal += ' (증가 중)';
    }

    Logger.log(`✅ Reserve Balances: ${(current/1000000).toFixed(2)}T, Status: ${status}`);

    return {
      status: status,
      score: score,
      signal: signal,
      current: current,
      monthChange: monthChange,
      currentTrillion: (current / 1000000).toFixed(2)
    };

  } catch (e) {
    Logger.log(`❌ Reserve Balances 오류: ${e.message}`);
    return { status: 'UNKNOWN', score: 0, signal: 'NO DATA', current: 0 };
  }
}

/**
 * Repo 스프레드 분석
 * SOFR-IORB 스프레드로 유동성 긴장 감지
 * @returns {Object} Repo 스프레드 상태 및 점수
 */
function getRepoSpreadStatus() {
  try {
    const sofr = getFredData(CONFIG.FRED_IDS.SOFR, false);
    const iorb = getFredData(CONFIG.FRED_IDS.IORB, false);

    const sofrRate = sofr.value || 0;
    const iorbRate = iorb.value || 0;
    const spread = (sofrRate - iorbRate) * 100; // bp 단위

    let status = 'NORMAL';
    let score = 0;
    let signal = '✅ 정상';

    if (spread > 15) {              // 15bp 초과 (긴장)
      status = 'TIGHT';
      score = -15;
      signal = '🔴 유동성 긴장';
    } else if (spread > 5) {        // 5-15bp (주의)
      status = 'ELEVATED';
      score = -5;
      signal = '⚠️ 스프레드 확대';
    } else if (spread < -5) {       // -5bp 미만 (과잉)
      status = 'EXCESS';
      score = 5;
      signal = '✅ 유동성 충분';
    }

    Logger.log(`✅ Repo Spread: ${spread.toFixed(1)}bp (SOFR: ${sofrRate}%, IORB: ${iorbRate}%)`);

    return {
      status: status,
      score: score,
      signal: signal,
      spread: spread,
      sofr: sofrRate,
      iorb: iorbRate
    };

  } catch (e) {
    Logger.log(`❌ Repo Spread 오류: ${e.message}`);
    return { status: 'UNKNOWN', score: 0, signal: 'NO DATA', spread: 0 };
  }
}

/**
 * 시장 국면 감지 (동적 가중치용)
 * DXY 변동성, Fed 정책 상태 기반으로 시장 국면 판단
 * @returns {Object} 시장 국면 및 가중치
 */
function detectMarketRegime() {
  try {
    // DXY 변동성 계산 (최근 30일 변화)
    const dxy = getFredData(CONFIG.GLOBAL_FRED_IDS.DXY, false);
    const dxy_1m = getFredDataHistorical(CONFIG.GLOBAL_FRED_IDS.DXY, 30);
    const dxy_1w = getFredDataHistorical(CONFIG.GLOBAL_FRED_IDS.DXY, 7);

    const dxyChange1w = Math.abs((dxy.value || 100) - (dxy_1w.value || 100));
    const dxyChange1m = Math.abs((dxy.value || 100) - (dxy_1m.value || 100));

    // Fed 정책 상태
    const rateCycle = getRateCycleStatus();
    const qtStatus = getQTStatus();

    const fedCutting = rateCycle.cycle.includes('CUTTING');
    const qtEnded = qtStatus.status === 'QE' || qtStatus.status === 'MILD_QE' || qtStatus.status === 'NEUTRAL';

    // v4.1: 새 기본 가중치 (US 45%, DXY 25%, China 5%, Japan 10%, EM 15%)
    let regime = 'NORMAL';
    let weights = {
      US: 0.45,
      DXY: 0.25,
      China: 0.05,
      Japan: 0.10,
      EM: 0.15
    };

    // 국면 판단 로직
    if (dxyChange1w > 2 || dxyChange1m > 5) {
      // DXY 고변동성 국면: DXY 영향력 증가
      regime = 'DXY_VOLATILE';
      weights = {
        US: 0.30,
        DXY: 0.40,
        China: 0.05,
        Japan: 0.10,
        EM: 0.15
      };
    } else if (fedCutting && qtEnded) {
      // Fed 적극 완화 국면: 미국 요인 지배
      regime = 'FED_EASING';
      weights = {
        US: 0.55,
        DXY: 0.15,
        China: 0.05,
        Japan: 0.10,
        EM: 0.15
      };
    } else if (fedCutting) {
      // 금리 인하 사이클: 미국 요인 강화
      regime = 'RATE_CUTTING';
      weights = {
        US: 0.50,
        DXY: 0.20,
        China: 0.05,
        Japan: 0.10,
        EM: 0.15
      };
    }

    Logger.log(`✅ Market Regime: ${regime}, DXY 변동성: 1w ${dxyChange1w.toFixed(2)}, 1m ${dxyChange1m.toFixed(2)}`);

    return {
      regime: regime,
      weights: weights,
      dxyVolatility1w: dxyChange1w,
      dxyVolatility1m: dxyChange1m,
      fedCutting: fedCutting,
      qtEnded: qtEnded
    };

  } catch (e) {
    Logger.log(`❌ Market Regime 오류: ${e.message}`);
    return {
      regime: 'NORMAL',
      weights: { US: 0.45, DXY: 0.25, China: 0.05, Japan: 0.10, EM: 0.15 }  // v4.1 가중치
    };
  }
}

/**
 * 맥락 기반 DXY 스코어링
 * Fed 정책 상태에 따라 DXY 변화 해석
 * @param {number} dxyChange - DXY 주간 변화
 * @returns {Object} 맥락 기반 점수
 */
function getContextualDXYScore(dxyChange) {
  const rateCycle = getRateCycleStatus();
  const fedCutting = rateCycle.cycle.includes('CUTTING');

  let score = 0;
  let interpretation = '';

  if (fedCutting) {
    // Fed 완화 중: DXY 하락 = 글로벌 유동성 증가 = 긍정적
    if (dxyChange < -2) {
      score = 25;
      interpretation = '🚀 달러 약세 + Fed 완화 = 강한 Risk-ON';
    } else if (dxyChange < -1) {
      score = 20;
      interpretation = '✅ 달러 약세 + Fed 완화 = Risk-ON';
    } else if (dxyChange < 0) {
      score = 10;
      interpretation = '✅ 약한 달러 약세 (긍정적)';
    } else if (dxyChange < 1) {
      score = 0;
      interpretation = '⚖️ 달러 안정';
    } else if (dxyChange < 2) {
      score = -10;
      interpretation = '⚠️ Fed 완화에도 달러 강세';
    } else {
      score = -15;
      interpretation = '🔴 비정상적 달러 강세 (경계)';
    }
  } else {
    // Fed 동결/긴축 중: 고변동성이 리스크
    if (Math.abs(dxyChange) > 2) {
      score = -15;
      interpretation = '🔴 달러 고변동성 (불안정)';
    } else if (Math.abs(dxyChange) > 1) {
      score = -5;
      interpretation = '⚠️ 달러 변동성';
    } else {
      score = 0;
      interpretation = '⚖️ 달러 안정';
    }
  }

  return {
    score: score,
    interpretation: interpretation,
    dxyChange: dxyChange,
    fedContext: fedCutting ? 'EASING' : 'NEUTRAL_OR_TIGHTENING'
  };
}

/**
 * 종합 Fed 정책 분석
 * QT/QE, 금리 사이클, Reserve, Repo 스프레드 통합
 * @returns {Object} 종합 Fed 정책 점수 및 상태
 */
function analyzeFedPolicy() {
  const qtStatus = getQTStatus();
  const rateCycle = getRateCycleStatus();
  const reserves = getReserveBalancesStatus();
  const repoSpread = getRepoSpreadStatus();

  // 종합 점수 (최대 +60, 최소 -60)
  const totalScore = qtStatus.score + rateCycle.score + reserves.score + repoSpread.score;

  // 종합 신호
  let overallSignal = '⚖️ 중립';
  if (totalScore >= 40) {
    overallSignal = '🚀 강한 완화';
  } else if (totalScore >= 20) {
    overallSignal = '✅ 완화적';
  } else if (totalScore <= -40) {
    overallSignal = '🔴 강한 긴축';
  } else if (totalScore <= -20) {
    overallSignal = '⚠️ 긴축적';
  }

  Logger.log(`✅ Fed Policy Analysis: Total Score ${totalScore}, ${overallSignal}`);

  return {
    totalScore: totalScore,
    overallSignal: overallSignal,
    components: {
      qtStatus: qtStatus,
      rateCycle: rateCycle,
      reserves: reserves,
      repoSpread: repoSpread
    }
  };
}

/** ===============================================
 * 6) 신흥국 통화 모니터링
 * =============================================== */

function getEmergingMarketsFX() {
  try {
    // Google Finance로 실시간 환율 조회 (FRED fallback)
    const usdkrw_value = getFXRate('USDKRW', CONFIG.GLOBAL_FRED_IDS.USDKRW);
    const usdbrl_value = getFXRate('USDBRL', CONFIG.GLOBAL_FRED_IDS.USDBRL);
    const usdmxn_value = getFXRate('USDMXN', CONFIG.GLOBAL_FRED_IDS.USDMXN);

    // 1주일 전 데이터는 FRED에서 (변화율 계산용)
    const usdkrw_1w = getFredDataHistorical(CONFIG.GLOBAL_FRED_IDS.USDKRW, 7);
    const usdbrl_1w = getFredDataHistorical(CONFIG.GLOBAL_FRED_IDS.USDBRL, 7);
    const usdmxn_1w = getFredDataHistorical(CONFIG.GLOBAL_FRED_IDS.USDMXN, 7);

    const krw_change = usdkrw_1w.value ? ((usdkrw_value - usdkrw_1w.value) / usdkrw_1w.value) * 100 : 0;
    const brl_change = usdbrl_1w.value ? ((usdbrl_value - usdbrl_1w.value) / usdbrl_1w.value) * 100 : 0;
    const mxn_change = usdmxn_1w.value ? ((usdmxn_value - usdmxn_1w.value) / usdmxn_1w.value) * 100 : 0;

    const strength_index = -(krw_change + brl_change + mxn_change) / 3;

    Logger.log(`✅ EM FX (Google Finance): KRW ${usdkrw_value}, BRL ${usdbrl_value}, MXN ${usdmxn_value}`);

    return {
      usdkrw: usdkrw_value,
      usdbrl: usdbrl_value,
      usdmxn: usdmxn_value,
      krw_change: krw_change,
      brl_change: brl_change,
      mxn_change: mxn_change,
      strength_index: strength_index,
      signal: strength_index > 1 ? '✅ EM 강세' :
              strength_index < -1 ? '⚠️ EM 약세' : '⚖️ 중립'
    };

  } catch (e) {
    Logger.log(`❌ EM FX 오류: ${e.message}`);
    return { usdkrw: 0, usdbrl: 0, usdmxn: 0, strength_index: 0, signal: 'NO DATA' };
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

    // ============= v4.0: 시장 국면 및 동적 가중치 감지 =============
    const marketRegime = detectMarketRegime();
    const fedPolicy = analyzeFedPolicy();

    // 데이터 수집 (v4.1: 주간 데이터는 직전 발표값과 비교)
    const walcl = getFredData(CONFIG.FRED_IDS.WALCL);
    const walcl_prev = getPreviousRelease(CONFIG.FRED_IDS.WALCL);  // 직전 주 발표값
    const tga = getTGAAnalysis();
    const onRrp = getFredData(CONFIG.FRED_IDS.ON_RRP);

    const dxy = getFredData(CONFIG.GLOBAL_FRED_IDS.DXY);
    const dxy_1w = getFredDataHistorical(CONFIG.GLOBAL_FRED_IDS.DXY, 7);  // 일간 데이터 - 7일 전
    const dxy_change = (dxy.value || 100) - (dxy_1w.value || 100);

    // 중국: DXY 변화와 함께 분석 (DXY+CNY 조합 로직)
    const china = getChinaLiquidity(dxy_change);
    const japan = getJapanLiquidity();
    const emFx = getEmergingMarketsFX();

    // WoW 계산 (v4.1: 직전 발표값 기준)
    const walcl_wow = (walcl.value || 0) - (walcl_prev.value || 0);

    // v4.1: WoW 디버깅 로그
    Logger.log(`📊 WALCL 계산: 현재 ${walcl.value} (${walcl.date}) - 직전 ${walcl_prev.value} (${walcl_prev.date}) = WoW ${walcl_wow}`);

    // ============= 요인별 원점수 계산 =============
    let usScore = 0;
    let dxyScore = 0;
    let chinaScore = 0;
    let japanScore = 0;
    let emScore = 0;

    // === 미국 요인 ===

    // 1. WALCL WoW (양방향 5단계 점수)
    if (walcl_wow > 50000) {
      usScore += 20;
    } else if (walcl_wow > 10000) {
      usScore += 10;
    } else if (walcl_wow < -50000) {
      usScore -= 20;
    } else if (walcl_wow < -10000) {
      usScore -= 10;
    }

    // 2. TGA 변화 (양방향 5단계 점수)
    if (tga.week_change < -100000) {
      usScore += 10;
    } else if (tga.week_change < -50000) {
      usScore += 5;
    } else if (tga.week_change > 100000) {
      usScore -= 10;
    } else if (tga.week_change > 50000) {
      usScore -= 5;
    }

    // 3. ON RRP (5단계 점수)
    if (onRrp.value > 500000) {
      usScore -= 15;
    } else if (onRrp.value > 300000) {
      usScore -= 10;
    } else if (onRrp.value > 200000) {
      usScore += 0;
    } else if (onRrp.value > 100000) {
      usScore += 5;
    } else {
      usScore += 10;
    }

    // 4. v4.0: Fed 정책 점수 추가
    usScore += fedPolicy.totalScore * 0.5;

    // === 달러 요인 (v4.0: 맥락 기반 스코어링) ===
    const contextualDxy = getContextualDXYScore(dxy_change);
    dxyScore = contextualDxy.score;

    // === 중국 요인 (v4.1: DXY+CNY 조합 로직, 5% 가중치) ===
    // getChinaLiquidity에서 이미 DXY와 CNY 조합으로 점수 계산됨
    chinaScore = china.score || 0;

    // === 일본 요인 ===
    if (japan.usdjpy > 155) {
      japanScore -= 15;
    } else if (japan.usdjpy > 150) {
      japanScore -= 10;
    } else if (japan.usdjpy > 145) {
      japanScore -= 5;
    } else if (japan.usdjpy < 130) {
      japanScore += 5;
    }

    // === 신흥국 요인 ===
    if (emFx.strength_index > 2) {
      emScore += 15;
    } else if (emFx.strength_index > 1) {
      emScore += 10;
    } else if (emFx.strength_index < -2) {
      emScore -= 15;
    } else if (emFx.strength_index < -1) {
      emScore -= 10;
    }

    // ============= v4.1: 새 가중치 (US 45%, DXY 25%, China 5%, Japan 10%, EM 15%) =============
    const weights = marketRegime.weights;
    let liquidityScore = Math.round(
      usScore * (weights.US / 0.45) +       // US 기본 45% 대비 조정
      dxyScore * (weights.DXY / 0.25) +     // DXY 기본 25% 대비 조정
      chinaScore * (weights.China / 0.05) + // China 기본 5% 대비 조정
      japanScore * (weights.Japan / 0.10) + // Japan 기본 10% 대비 조정
      emScore * (weights.EM / 0.15)         // EM 기본 15% 대비 조정
    );

    // 최종 신호 결정 (7단계 확장 범위)
    let finalSignal = '';
    let recommendation = '';

    if (liquidityScore >= 80) {
      finalSignal = '🚀🚀 SUPER LIQUIDITY';
      recommendation = '공격적 Risk-ON: 레버리지 ETF, 성장주, 비트코인, 신흥국 전면 확대';
    } else if (liquidityScore >= 50) {
      finalSignal = '🚀 EXTREME LIQUIDITY';
      recommendation = '적극적 Risk-ON: 성장주, 신흥국, 원자재 비중 확대';
    } else if (liquidityScore >= 20) {
      finalSignal = '✅ HIGH LIQUIDITY';
      recommendation = '위험자산 비중 유지/확대, 밸류/그로스 균형';
    } else if (liquidityScore >= -20) {
      finalSignal = '⚖️ NEUTRAL';
      recommendation = '포트폴리오 균형 유지, 관망';
    } else if (liquidityScore >= -50) {
      finalSignal = '⚠️ TIGHT';
      recommendation = '현금/채권 비중 증대, 방어주 선호';
    } else if (liquidityScore >= -80) {
      finalSignal = '🔴 EXTREME TIGHT';
      recommendation = '방어적 포지션, 달러/금/국채 선호';
    } else {
      finalSignal = '🔴🔴 CRISIS MODE';
      recommendation = '현금 확보, 손절 고려, 변동성 헤지 필수';
    }

    // Global_Liquidity 시트 업데이트
    const timestamp = new Date().toLocaleString('ko-KR', {timeZone: 'Asia/Seoul'});

    globalSheet.getRange(2, 1, 1, 19).setValues([[
      timestamp,
      walcl.value,
      walcl_wow,
      tga.current,
      tga.week_change,
      onRrp.value,
      dxy.value,
      dxy_change,
      china.usdcny,           // v4.1: USD/CNY 환율
      china.usdcny_change,    // v4.1: CNY 주간 변화율
      china.liquidity_signal, // v4.1: DXY+CNY 조합 신호
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

    // v4.0: 시장 국면 표시 (U2 셀)
    globalSheet.getRange('U2').setValue(`[${marketRegime.regime}] ${fedPolicy.overallSignal}`);

    // 조건부 서식 (7단계)
    const signalCell = globalSheet.getRange('S2');
    if (liquidityScore >= 80) {
      signalCell.setBackground('#00FF00').setFontWeight('bold');
    } else if (liquidityScore >= 50) {
      signalCell.setBackground('#90EE90');
    } else if (liquidityScore >= 20) {
      signalCell.setBackground('#D4EDDA');
    } else if (liquidityScore >= -20) {
      signalCell.setBackground('#FFFFE0');
    } else if (liquidityScore >= -50) {
      signalCell.setBackground('#FFE4B5');
    } else if (liquidityScore >= -80) {
      signalCell.setBackground('#FFB6C1');
    } else {
      signalCell.setBackground('#FF6B6B').setFontWeight('bold');
    }

    Logger.log(`✅ 글로벌 유동성 분석 완료: Score ${liquidityScore}, ${finalSignal}`);
    Logger.log(`   Market Regime: ${marketRegime.regime}, Fed Policy: ${fedPolicy.overallSignal}`);
    Logger.log(`   Weights - US:${(weights.US*100).toFixed(0)}% DXY:${(weights.DXY*100).toFixed(0)}% China:${(weights.China*100).toFixed(0)}%`);

    return {
      score: liquidityScore,
      signal: finalSignal,
      recommendation: recommendation,
      timestamp: new Date(),
      regime: marketRegime.regime,
      fedPolicy: fedPolicy,
      weights: weights,
      componentScores: {
        us: usScore,
        dxy: dxyScore,
        china: chinaScore,
        japan: japanScore,
        em: emScore
      },
      details: {
        us: { walcl: walcl.value, walcl_wow: walcl_wow, tga: tga, onrrp: onRrp.value },
        dxy: { level: dxy.value, change: dxy_change, contextual: contextualDxy },
        china: china,
        japan: japan,
        em: emFx
      }
    };

  } catch (e) {
    Logger.log(`❌ 글로벌 유동성 분석 오류: ${e.message}`);

    try {
      SpreadsheetApp.getUi().alert(`❌ 오류: ${e.message}`);
    } catch (uiError) {
      Logger.log('UI 알림 불가 (트리거 실행 중)');
    }

    return {
      score: 0,
      signal: 'ERROR',
      timestamp: new Date(),
      recommendation: '오류 발생 - 데이터 확인 필요',
      regime: 'UNKNOWN',
      fedPolicy: null,
      weights: { US: 0.45, DXY: 0.25, China: 0.05, Japan: 0.10, EM: 0.15 },  // v4.1 가중치
      componentScores: { us: 0, dxy: 0, china: 0, japan: 0, em: 0 },
      details: {
        us: { walcl: 0, walcl_wow: 0, tga: { current: 0, week_change: 0 }, onrrp: 0 },
        dxy: { level: 0, change: 0, contextual: null },
        china: { usdcny: 0, usdcny_change: 0, fx_reserves: 0, score: 0 },  // v4.1 구조
        japan: { usdjpy: 0, jgb_10y: 0, us_jpy_spread: 0 },
        em: { usdkrw: 0, usdbrl: 0, strength_index: 0 }
      }
    };
  }
}

function setupGlobalSheet(sheet) {
  if (!sheet) {
    Logger.log('❌ setupGlobalSheet: sheet가 undefined입니다');
    return;
  }

  const headers = [
    '타임스탬프',
    'WALCL(M$)', 'WALCL WoW',
    'TGA(M$)', 'TGA WoW',
    'ON RRP(M$)',
    'DXY', 'DXY WoW',
    'USD/CNY', 'CNY 변화(%)', '중국 신호',  // v4.1: DXY+CNY 조합
    'USD/JPY', 'JGB 10Y', 'US-JP 스프레드',
    'USD/KRW', 'USD/BRL', 'EM 강세지수',
    '유동성 점수', '신호', '투자 권장',
    '시장 국면'  // v4.0: 시장 국면 및 Fed 정책 상태
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
 * Global_History 마지막 두 행 비교 분석
 * =============================================== */

function getHistoryComparison() {
  try {
    const ss = SpreadsheetApp.getActive();
    const globalHistorySheet = ss.getSheetByName(CONFIG.GLOBAL_HISTORY_SHEET);

    if (!globalHistorySheet) {
      return null;
    }

    const lastRow = globalHistorySheet.getLastRow();
    if (lastRow < 3) {
      // 최소 헤더 + 2개 행이 필요
      return null;
    }

    // 마지막 두 행 가져오기 (20개 컬럼)
    const data = globalHistorySheet.getRange(lastRow - 1, 1, 2, 20).getValues();
    const prev = data[0];
    const curr = data[1];

    // 컬럼 인덱스 (0-based)
    // 0: 타임스탬프, 1: WALCL, 2: WALCL WoW, 3: TGA, 4: TGA WoW, 5: ON RRP
    // 6: DXY, 7: DXY WoW, 8: 중국 M2, 9: 중국 신용, 10: 중국 FX
    // 11: USD/JPY, 12: JGB 10Y, 13: US-JP 스프레드
    // 14: USD/KRW, 15: USD/BRL, 16: EM 강세지수
    // 17: 유동성 점수, 18: 신호, 19: 투자 권장

    return {
      prevTimestamp: prev[0],
      currTimestamp: curr[0],
      dxy: {
        prev: prev[6],
        curr: curr[6],
        change: curr[6] - prev[6]
      },
      walcl: {
        prev: prev[1],
        curr: curr[1],
        wow: curr[2]
      },
      tga: {
        prev: prev[3],
        curr: curr[3],
        wow: curr[4]
      },
      onrrp: {
        prev: prev[5],
        curr: curr[5],
        change: curr[5] - prev[5]
      },
      chinaM2: {
        prev: prev[8],
        curr: curr[8],
        change: curr[8] - prev[8]
      },
      usdjpy: {
        prev: prev[11],
        curr: curr[11],
        change: curr[11] - prev[11]
      },
      usdkrw: {
        prev: prev[14],
        curr: curr[14],
        change: curr[14] - prev[14]
      },
      emIndex: {
        prev: prev[16],
        curr: curr[16],
        change: curr[16] - prev[16]
      },
      score: {
        prev: prev[17],
        curr: curr[17],
        change: curr[17] - prev[17]
      }
    };

  } catch (e) {
    Logger.log(`❌ History 비교 분석 오류: ${e.message}`);
    return null;
  }
}

/** ===============================================
 * 글로벌 유동성 히스토리 기록
 * =============================================== */

function logGlobalHistory(analysis) {
  try {
    // 유효성 검사
    if (!analysis || !analysis.timestamp || !analysis.details) {
      Logger.log('❌ logGlobalHistory: analysis 객체가 유효하지 않습니다');
      return;
    }

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

    // 안전한 값 추출 (undefined 방지)
    const us = analysis.details.us || {};
    const dxy = analysis.details.dxy || {};
    const china = analysis.details.china || {};
    const japan = analysis.details.japan || {};
    const em = analysis.details.em || {};
    const tga = us.tga || {};

    // v4.1: WoW 값 로깅 (디버깅용)
    Logger.log(`📊 History 기록 - WALCL: ${us.walcl}, WoW: ${us.walcl_wow}`);
    Logger.log(`📊 History 기록 - TGA: ${tga.current}, WoW: ${tga.week_change}`);

    // 히스토리에 추가 (v4.1: 중국 데이터 구조 업데이트)
    globalHistorySheet.appendRow([
      analysis.timestamp,
      us.walcl || 0,
      us.walcl_wow || 0,
      tga.current || 0,
      tga.week_change || 0,
      us.onrrp || 0,
      dxy.level || 0,
      dxy.change || 0,
      china.usdcny || 0,           // v4.1: USD/CNY
      china.usdcny_change || 0,    // v4.1: CNY 변화율
      china.liquidity_signal || 'N/A',  // v4.1: 중국 신호
      japan.usdjpy || 0,
      japan.jgb_10y || 0,
      japan.us_jpy_spread || 0,
      em.usdkrw || 0,
      em.usdbrl || 0,
      em.strength_index || 0,
      analysis.score || 0,
      analysis.signal || 'N/A',
      analysis.recommendation || 'N/A'
    ]);

    Logger.log('✅ Global_History 기록 완료');

  } catch (e) {
    Logger.log(`❌ Global_History 기록 오류: ${e.message}`);
  }
}

/** ===============================================
 * 7-B) History 시트 일괄 업데이트 (올해 1월부터)
 * =============================================== */

/**
 * 가장 가까운 이전 날짜의 값을 찾는 헬퍼 함수
 * @param {Object} dataMap - 날짜:값 맵
 * @param {string} targetDate - 찾고자 하는 날짜
 * @returns {number} 값 또는 0
 */
function getClosestValue(dataMap, targetDate) {
  if (dataMap[targetDate] !== undefined) {
    return dataMap[targetDate];
  }

  // 이전 날짜들 중 가장 가까운 날짜 찾기
  const dates = Object.keys(dataMap).sort();
  for (let i = dates.length - 1; i >= 0; i--) {
    if (dates[i] <= targetDate) {
      return dataMap[dates[i]];
    }
  }

  return 0;
}

/**
 * History 시트를 올해 1월부터 현재까지 데이터로 채우기
 */
function populateHistoryFromJanuary() {
  try {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'History 데이터 업데이트',
      '올해 1월 1일부터 현재까지 데이터를 History 시트에 추가합니다.\n\n계속하시겠습니까?',
      ui.ButtonSet.YES_NO
    );

    if (result !== ui.Button.YES) {
      return;
    }

    Logger.log('=== History 시트 일괄 업데이트 시작 ===');

    const ss = SpreadsheetApp.getActive();
    let historySheet = ss.getSheetByName(CONFIG.HISTORY_SHEET);

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

    // 올해 1월 1일
    const startDate = new Date('2025-01-01');

    // 모든 지표의 데이터 가져오기
    Logger.log('📥 FRED 데이터 수집 중...');
    const walclData = getFredDataRange(CONFIG.FRED_IDS.WALCL, startDate);
    const sofrData = getFredDataRange(CONFIG.FRED_IDS.SOFR, startDate);
    const effrData = getFredDataRange(CONFIG.FRED_IDS.EFFR, startDate);
    const iorbData = getFredDataRange(CONFIG.FRED_IDS.IORB, startDate);
    const onRrpData = getFredDataRange(CONFIG.FRED_IDS.ON_RRP, startDate);
    const tgaData = getFredDataRange(CONFIG.FRED_IDS.TGA, startDate);

    // WALCL을 기준으로 날짜 목록 생성 (주간 데이터)
    const walclDates = Object.keys(walclData).sort();

    if (walclDates.length === 0) {
      ui.alert('❌ WALCL 데이터를 찾을 수 없습니다.');
      return;
    }

    Logger.log(`📊 ${walclDates.length}개 주간 데이터 포인트 처리 중...`);

    // 각 날짜별로 데이터 행 생성
    const rows = [];
    for (let i = 0; i < walclDates.length; i++) {
      const date = walclDates[i];
      const walcl = walclData[date];

      // WoW 계산 (이전 주 데이터와 비교)
      const walcl_prev = i > 0 ? walclData[walclDates[i-1]] : walcl;
      const wow = walcl - walcl_prev;

      // 각 지표의 가장 가까운 값 찾기
      const sofr = getClosestValue(sofrData, date);
      const effr = getClosestValue(effrData, date);
      const iorb = getClosestValue(iorbData, date);
      const onRrp = getClosestValue(onRrpData, date);
      const tga = getClosestValue(tgaData, date);

      // SOFR-IORB 스프레드 (bp)
      const sofr_iorb = (sofr - iorb) * 100;

      // 신호 판단
      const signal = determineSignal(sofr_iorb, onRrp, wow, walcl);

      // SRF는 historical 데이터가 없으므로 0으로 설정
      const srf = 0;

      // 타임스탬프는 해당 날짜의 자정으로 설정
      const timestamp = new Date(date);

      rows.push([
        timestamp,
        date,
        sofr,
        effr,
        iorb,
        sofr_iorb,
        onRrp,
        tga,
        walcl,
        wow,
        srf,
        signal
      ]);
    }

    // History 시트에 모든 행 추가
    if (rows.length > 0) {
      historySheet.getRange(historySheet.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
      Logger.log(`✅ ${rows.length}개 행이 History 시트에 추가됨`);

      ui.alert(
        '✅ 완료',
        `${rows.length}개 데이터 포인트가 History 시트에 추가되었습니다.\n\n기간: ${walclDates[0]} ~ ${walclDates[walclDates.length-1]}`,
        ui.ButtonSet.OK
      );
    }

    Logger.log('=== History 시트 업데이트 완료 ===');

  } catch (e) {
    Logger.log(`❌ History 업데이트 오류: ${e.message}`);
    SpreadsheetApp.getUi().alert(`❌ 오류: ${e.message}`);
  }
}

/**
 * Global_History 시트를 올해 1월부터 현재까지 데이터로 채우기
 */
function populateGlobalHistoryFromJanuary() {
  Logger.log('>>> populateGlobalHistoryFromJanuary 시작');
  try {
    Logger.log('>>> SpreadsheetApp.getActive() 호출');
    const ss = SpreadsheetApp.getActive();
    Logger.log('>>> SpreadsheetApp.getUi() 호출');
    const ui = SpreadsheetApp.getUi();
    Logger.log('>>> ui.alert() 호출 전');
    const result = ui.alert(
      'Global History 데이터 업데이트',
      '올해 1월 1일부터 현재까지 데이터를 Global_History 시트에 추가합니다.\n\n⚠️ 이 작업은 2-3분 소요됩니다. (13개 API 호출)\n\n진행 상황은 우측 하단 토스트에서 확인하세요.\n\n계속하시겠습니까?',
      ui.ButtonSet.YES_NO
    );
    Logger.log('>>> ui.alert() 응답: ' + result);

    if (result !== ui.Button.YES) {
      Logger.log('>>> 사용자가 취소함');
      return;
    }

    Logger.log('=== Global_History 시트 일괄 업데이트 시작 ===');
    ss.toast('🚀 Global History 업데이트 시작...', '진행 중', -1);

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
        'USD/CNY', 'CNY 변화(%)', '중국 신호',  // v4.1
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

    // 올해 1월 1일
    const startDate = new Date('2025-01-01');

    // 모든 지표의 데이터 가져오기 (진행 상황 표시)
    Logger.log('📥 FRED 데이터 수집 중...');

    // US 지표
    ss.toast('📥 1/11: WALCL 데이터 수집 중...', '진행 중', -1);
    const walclData = getFredDataRange(CONFIG.FRED_IDS.WALCL, startDate);

    ss.toast('📥 2/11: TGA 데이터 수집 중...', '진행 중', -1);
    const tgaData = getFredDataRange(CONFIG.FRED_IDS.TGA, startDate);

    ss.toast('📥 3/11: ON RRP 데이터 수집 중...', '진행 중', -1);
    const onRrpData = getFredDataRange(CONFIG.FRED_IDS.ON_RRP, startDate);

    // Global 지표
    ss.toast('📥 4/11: DXY 데이터 수집 중...', '진행 중', -1);
    const dxyData = getFredDataRange(CONFIG.GLOBAL_FRED_IDS.DXY, startDate);

    ss.toast('📥 5/11: USD/CNY 데이터 수집 중...', '진행 중', -1);
    const usdcnyData = getFredDataRange(CONFIG.GLOBAL_FRED_IDS.USDCNY, startDate);

    ss.toast('📥 6/11: USD/JPY 데이터 수집 중...', '진행 중', -1);
    const usdjpyData = getFredDataRange(CONFIG.GLOBAL_FRED_IDS.USDJPY, startDate);

    ss.toast('📥 7/11: JGB 10Y 데이터 수집 중...', '진행 중', -1);
    const jgb10yData = getFredDataRange(CONFIG.GLOBAL_FRED_IDS.JGB_10Y, startDate);

    ss.toast('📥 8/11: US 10Y 데이터 수집 중...', '진행 중', -1);
    const us10yData = getFredDataRange('DGS10', startDate);

    ss.toast('📥 9/11: USD/KRW 데이터 수집 중...', '진행 중', -1);
    const usdkrwData = getFredDataRange(CONFIG.GLOBAL_FRED_IDS.USDKRW, startDate);

    ss.toast('📥 10/11: USD/BRL 데이터 수집 중...', '진행 중', -1);
    const usdbrData = getFredDataRange(CONFIG.GLOBAL_FRED_IDS.USDBRL, startDate);

    ss.toast('📥 11/11: USD/MXN 데이터 수집 중...', '진행 중', -1);
    const usdmxnData = getFredDataRange(CONFIG.GLOBAL_FRED_IDS.USDMXN, startDate);

    ss.toast('📊 데이터 처리 중...', '진행 중', -1);

    // WALCL을 기준으로 날짜 목록 생성 (주간 데이터)
    const walclDates = Object.keys(walclData).sort();

    if (walclDates.length === 0) {
      ui.alert('❌ WALCL 데이터를 찾을 수 없습니다.');
      return;
    }

    Logger.log(`📊 ${walclDates.length}개 주간 데이터 포인트 처리 중...`);

    // 각 날짜별로 데이터 행 생성
    const rows = [];
    for (let i = 0; i < walclDates.length; i++) {
      const date = walclDates[i];

      // US 데이터
      const walcl = walclData[date];
      const walcl_prev = i > 0 ? walclData[walclDates[i-1]] : walcl;
      const walcl_wow = walcl - walcl_prev;

      const tga = getClosestValue(tgaData, date);
      const tga_prev = i > 0 ? getClosestValue(tgaData, walclDates[i-1]) : tga;
      const tga_wow = tga - tga_prev;

      const onRrp = getClosestValue(onRrpData, date);

      // Global 데이터
      const dxy = getClosestValue(dxyData, date);
      const dxy_prev = i > 0 ? getClosestValue(dxyData, walclDates[i-1]) : dxy;
      const dxy_wow = dxy - dxy_prev;

      // v4.1: USD/CNY 데이터 및 변화율 계산
      const usdcny = getClosestValue(usdcnyData, date);
      const usdcny_prev = i > 0 ? getClosestValue(usdcnyData, walclDates[i-1]) : usdcny;
      const usdcny_change = usdcny_prev !== 0 ? ((usdcny - usdcny_prev) / usdcny_prev) * 100 : 0;

      const usdjpy = getClosestValue(usdjpyData, date);
      const jgb10y = getClosestValue(jgb10yData, date);
      const us10y = getClosestValue(us10yData, date);
      const usJpSpread = us10y - jgb10y;

      const usdkrw = getClosestValue(usdkrwData, date);
      const usdbrl = getClosestValue(usdbrData, date);
      const usdmxn = getClosestValue(usdmxnData, date);

      // EM 강세 지수 계산
      const usdkrw_prev = i > 0 ? getClosestValue(usdkrwData, walclDates[i-1]) : usdkrw;
      const usdbrl_prev = i > 0 ? getClosestValue(usdbrData, walclDates[i-1]) : usdbrl;
      const usdmxn_prev = i > 0 ? getClosestValue(usdmxnData, walclDates[i-1]) : usdmxn;

      const krw_change = usdkrw_prev !== 0 ? ((usdkrw - usdkrw_prev) / usdkrw_prev) * 100 : 0;
      const brl_change = usdbrl_prev !== 0 ? ((usdbrl - usdbrl_prev) / usdbrl_prev) * 100 : 0;
      const mxn_change = usdmxn_prev !== 0 ? ((usdmxn - usdmxn_prev) / usdmxn_prev) * 100 : 0;

      const emStrengthIndex = -(krw_change + brl_change + mxn_change) / 3;

      // === 유동성 점수 계산 (analyzeGlobalLiquidity 로직과 동일) ===
      let liquidityScore = 0;

      // 미국 요인 (40%)
      if (walcl_wow > 50000) liquidityScore += 20;
      else if (walcl_wow > 10000) liquidityScore += 10;
      else if (walcl_wow < -50000) liquidityScore -= 20;
      else if (walcl_wow < -10000) liquidityScore -= 10;

      if (tga_wow < -100000) liquidityScore += 10;
      else if (tga_wow < -50000) liquidityScore += 5;
      else if (tga_wow > 100000) liquidityScore -= 10;
      else if (tga_wow > 50000) liquidityScore -= 5;

      if (onRrp > 500000) liquidityScore -= 15;
      else if (onRrp > 300000) liquidityScore -= 10;
      else if (onRrp > 200000) liquidityScore += 0;
      else if (onRrp > 100000) liquidityScore += 5;
      else liquidityScore += 10;

      // 달러 요인 (20%)
      if (dxy_wow < -2) liquidityScore += 25;
      else if (dxy_wow < -1) liquidityScore += 20;
      else if (dxy_wow > 2) liquidityScore -= 25;
      else if (dxy_wow > 1) liquidityScore -= 20;

      // v4.1: 중국 요인 (5%) - DXY+CNY 조합 로직
      const chinaSignal = determineChinaSignal(dxy_wow, usdcny_change);
      liquidityScore += chinaSignal.score;

      // 일본 요인 (10%)
      if (usdjpy > 155) liquidityScore -= 15;
      else if (usdjpy > 150) liquidityScore -= 10;
      else if (usdjpy > 145) liquidityScore -= 5;
      else if (usdjpy < 130) liquidityScore += 5;

      // 신흥국 요인 (10%)
      if (emStrengthIndex > 2) liquidityScore += 15;
      else if (emStrengthIndex > 1) liquidityScore += 10;
      else if (emStrengthIndex < -2) liquidityScore -= 15;
      else if (emStrengthIndex < -1) liquidityScore -= 10;

      // 신호 및 권장사항
      let signal = '';
      let recommendation = '';

      if (liquidityScore >= 80) {
        signal = '🚀🚀 SUPER LIQUIDITY';
        recommendation = '공격적 Risk-ON: 레버리지 ETF, 성장주, 비트코인, 신흥국 전면 확대';
      } else if (liquidityScore >= 50) {
        signal = '🚀 EXTREME LIQUIDITY';
        recommendation = '적극적 Risk-ON: 성장주, 신흥국, 원자재 비중 확대';
      } else if (liquidityScore >= 20) {
        signal = '✅ HIGH LIQUIDITY';
        recommendation = '위험자산 비중 유지/확대, 밸류/그로스 균형';
      } else if (liquidityScore >= -20) {
        signal = '⚖️ NEUTRAL';
        recommendation = '포트폴리오 균형 유지, 관망';
      } else if (liquidityScore >= -50) {
        signal = '⚠️ TIGHT';
        recommendation = '현금/채권 비중 증대, 방어주 선호';
      } else if (liquidityScore >= -80) {
        signal = '🔴 EXTREME TIGHT';
        recommendation = '방어적 포지션, 달러/금/국채 선호';
      } else {
        signal = '🔴🔴 CRISIS MODE';
        recommendation = '현금 확보, 손절 고려, 변동성 헤지 필수';
      }

      const timestamp = new Date(date);

      rows.push([
        timestamp,
        walcl,
        walcl_wow,
        tga,
        tga_wow,
        onRrp,
        dxy,
        dxy_wow,
        usdcny,                // v4.1: USD/CNY 환율
        usdcny_change,         // v4.1: CNY 변화율 (%)
        chinaSignal.signal,    // v4.1: 중국 신호
        usdjpy,
        jgb10y,
        usJpSpread,
        usdkrw,
        usdbrl,
        emStrengthIndex,
        liquidityScore,
        signal,
        recommendation
      ]);
    }

    // Global_History 시트에 모든 행 추가
    if (rows.length > 0) {
      ss.toast('💾 시트에 저장 중...', '진행 중', -1);
      globalHistorySheet.getRange(globalHistorySheet.getLastRow() + 1, 1, rows.length, 20).setValues(rows);
      Logger.log(`✅ ${rows.length}개 행이 Global_History 시트에 추가됨`);

      ss.toast(`✅ ${rows.length}개 데이터 저장 완료!`, '완료', 5);
      ui.alert(
        '✅ 완료',
        `${rows.length}개 데이터 포인트가 Global_History 시트에 추가되었습니다.\n\n기간: ${walclDates[0]} ~ ${walclDates[walclDates.length-1]}`,
        ui.ButtonSet.OK
      );
    }

    Logger.log('=== Global_History 시트 업데이트 완료 ===');

  } catch (e) {
    Logger.log(`❌ Global_History 업데이트 오류: ${e.message}`);
    try {
      SpreadsheetApp.getActive().toast(`❌ 오류: ${e.message}`, '오류', 10);
      SpreadsheetApp.getUi().alert(`❌ 오류: ${e.message}`);
    } catch (uiError) {
      Logger.log('UI 알림 불가 (트리거 실행 중)');
    }
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
    const now = new Date().toLocaleString('ko-KR', {timeZone: 'Asia/Seoul'});
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

    // Global_History에 현재 분석 결과 기록
    logGlobalHistory(analysis);

    const alerts = [];

    // v4.0: 추세 기반 알림을 위한 히스토리 비교
    const comparison = getHistoryComparison();
    let momentum = 0;
    let trend = 0;

    if (comparison && comparison.score) {
      momentum = analysis.score - (comparison.score.prev || 0);
      // 30일 추세 계산을 위해 7일 전 데이터 활용 (대략적)
      trend = analysis.score - (comparison.score.prev || 0);
    }

    // === 레벨 기반 알림 (기존) ===

    // 극단적 신호 (업데이트된 기준)
    if (analysis.score >= 80) {
      alerts.push({
        level: '🚀🚀 SUPER OPPORTUNITY',
        message: '슈퍼 유동성 폭발 - 역사적 기회',
        action: analysis.recommendation
      });
    } else if (analysis.score >= 50) {
      alerts.push({
        level: '🚀 EXTREME OPPORTUNITY',
        message: '극도의 유동성 급증',
        action: analysis.recommendation
      });
    } else if (analysis.score <= -80) {
      alerts.push({
        level: '🔴🔴 CRISIS ALERT',
        message: '위기 수준 유동성 경색',
        action: analysis.recommendation
      });
    } else if (analysis.score <= -50) {
      alerts.push({
        level: '🔴 EXTREME WARNING',
        message: '극도의 유동성 급감',
        action: analysis.recommendation
      });
    }

    // === v4.0: 추세/모멘텀 기반 알림 (신규) ===

    // 급격한 모멘텀 변화 감지
    if (momentum > 20) {
      alerts.push({
        level: '📈 MOMENTUM SURGE',
        message: `유동성 급증: 점수 +${momentum.toFixed(0)}pt 상승`,
        action: '강한 매수 신호 - Risk-ON 포지션 확대 검토'
      });
    } else if (momentum < -20) {
      alerts.push({
        level: '📉 MOMENTUM DROP',
        message: `유동성 급감: 점수 ${momentum.toFixed(0)}pt 하락`,
        action: '방어적 전환 검토 - 포지션 축소 고려'
      });
    }

    // 추세 전환 감지 (긍정적 추세 중 모멘텀 둔화)
    if (trend > 0 && momentum < -10 && analysis.score > 0) {
      alerts.push({
        level: '⚠️ TREND SLOWDOWN',
        message: '유동성 모멘텀 둔화 - 경계 구간 진입',
        action: '포지션 유지하되 신규 매수 자제'
      });
    }

    // 추세 전환 감지 (부정적 추세 중 반등 신호)
    if (trend < 0 && momentum > 10 && analysis.score < 0) {
      alerts.push({
        level: '💡 TREND REVERSAL',
        message: '유동성 반등 시작 가능성',
        action: '바닥 신호 주시 - 분할 매수 검토'
      });
    }

    // === v4.0: Fed 정책 전환 알림 (신규) ===

    if (analysis.fedPolicy) {
      const fedScore = analysis.fedPolicy.totalScore;

      if (fedScore >= 40) {
        alerts.push({
          level: '🏦 FED EASING',
          message: `Fed 강한 완화 신호 (정책점수: +${fedScore})`,
          action: '위험자산 비중 확대 적기'
        });
      } else if (fedScore <= -40) {
        alerts.push({
          level: '🏦 FED TIGHTENING',
          message: `Fed 강한 긴축 신호 (정책점수: ${fedScore})`,
          action: '방어적 포지션 전환 권장'
        });
      }
    }

    // === v4.0: 시장 국면 전환 알림 (신규) ===

    if (analysis.regime === 'DXY_VOLATILE') {
      alerts.push({
        level: '💵 REGIME: DXY VOLATILE',
        message: 'DXY 고변동성 국면 - 가중치 조정됨',
        action: '달러 움직임 주시, EM 노출 관리'
      });
    } else if (analysis.regime === 'FED_EASING') {
      alerts.push({
        level: '🏦 REGIME: FED EASING',
        message: 'Fed 적극 완화 국면 - 미국 요인 지배적',
        action: 'Fed 정책에 민감한 자산 선호'
      });
    }

    // === 기존 개별 리스크 알림 ===

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
      sendGlobalAlert(alerts, analysis, comparison);
      logAlertHistory(alerts, analysis);
    }

    Logger.log(`✅ 알림 체크 완료: ${alerts.length}개 알림 발생`);

  } catch (e) {
    Logger.log(`❌ 알림 체크 오류: ${e.message}`);
  }
}

/**
 * 알림 테스트 - 조건 없이 바로 테스트 알림 발송
 * 메뉴에서 실행: 📊 Global Liquidity → 🧪 알림 테스트
 */
function testGlobalAlert() {
  try {
    const analysis = analyzeGlobalLiquidity();

    // Global_History에 현재 분석 결과 기록
    logGlobalHistory(analysis);

    // 테스트용 알림 생성
    const alerts = [{
      level: '🧪 TEST ALERT',
      message: '알림 테스트 - 이메일 발송 확인',
      action: '테스트 완료 후 이 알림을 무시하세요'
    }];

    // History에서 이전 데이터와 비교
    const comparison = getHistoryComparison();
    sendGlobalAlert(alerts, analysis, comparison);

    SpreadsheetApp.getUi().alert('✅ 테스트 알림이 발송되었습니다.\n이메일을 확인해주세요.');
    Logger.log('✅ 테스트 알림 발송 완료');

  } catch (e) {
    Logger.log(`❌ 테스트 알림 오류: ${e.message}`);
    SpreadsheetApp.getUi().alert(`❌ 테스트 알림 실패: ${e.message}`);
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

/**
 * 이모지를 HTML numeric entity로 변환
 * Gmail에서 이모지가 깨지는 문제 해결
 */
function encodeEmojis(str) {
  return str.replace(/[\u{1F300}-\u{1F9FF}]|[\u{2600}-\u{26FF}]|[\u{2700}-\u{27BF}]|[\u{1F600}-\u{1F64F}]|[\u{1F680}-\u{1F6FF}]|[\u{1F1E0}-\u{1F1FF}]/gu,
    (char) => `&#${char.codePointAt(0)};`
  );
}

function sendGlobalAlert(alerts, analysis, comparison) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const timestamp = new Date().toLocaleString('ko-KR', {timeZone: 'Asia/Seoul'});

    // 변화량 포맷팅 헬퍼
    const formatChange = (value) => {
      if (value === null || value === undefined || isNaN(value)) return '';
      const sign = value >= 0 ? '+' : '';
      return ` (${sign}${value.toFixed(2)})`;
    };

    // 비교 데이터가 있는 경우 변화량 계산
    const dxyChange = comparison ? formatChange(comparison.dxy.change) : formatChange(analysis.details.dxy.change);
    const usdjpyChange = comparison ? formatChange(comparison.usdjpy.change) : '';
    const chinaM2Change = comparison ? formatChange(comparison.chinaM2.change) : '';
    const scoreChange = comparison ? formatChange(comparison.score.change) : '';

    let emailBody = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
      </head>
      <body>
      <div style="font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', Arial, sans-serif; background-color: #f5f5f5; padding: 20px;">
        <h2 style="color: #1f77b4;">🌐 글로벌 유동성 알림</h2>
        <p><strong>시간:</strong> ${timestamp}</p>
        <p><strong>유동성 점수:</strong> ${analysis.score} / 100${scoreChange}</p>
        <p><strong>신호:</strong> ${analysis.signal}</p>

        <h3>📊 주요 지표</h3>
        <table style="border-collapse: collapse; width: 100%; background: white;">
          <tr>
            <td style="border: 1px solid #ddd; padding: 8px;"><strong>DXY:</strong></td>
            <td style="border: 1px solid #ddd; padding: 8px;">${analysis.details.dxy.level.toFixed(2)}${dxyChange}</td>
          </tr>
          <tr>
            <td style="border: 1px solid #ddd; padding: 8px;"><strong>WALCL WoW:</strong></td>
            <td style="border: 1px solid #ddd; padding: 8px;">${analysis.details.us.walcl_wow.toFixed(0)}M$</td>
          </tr>
          <tr>
            <td style="border: 1px solid #ddd; padding: 8px;"><strong>TGA WoW:</strong></td>
            <td style="border: 1px solid #ddd; padding: 8px;">${analysis.details.us.tga.week_change.toFixed(0)}M$</td>
          </tr>
          <tr>
            <td style="border: 1px solid #ddd; padding: 8px;"><strong>ON RRP:</strong></td>
            <td style="border: 1px solid #ddd; padding: 8px;">${analysis.details.us.onrrp.toFixed(0)}M$</td>
          </tr>
          <tr>
            <td style="border: 1px solid #ddd; padding: 8px;"><strong>중국 M2:</strong></td>
            <td style="border: 1px solid #ddd; padding: 8px;">${analysis.details.china.m2_growth.toFixed(1)}%${chinaM2Change}</td>
          </tr>
          <tr>
            <td style="border: 1px solid #ddd; padding: 8px;"><strong>USD/JPY:</strong></td>
            <td style="border: 1px solid #ddd; padding: 8px;">${analysis.details.japan.usdjpy.toFixed(2)}${usdjpyChange}</td>
          </tr>
          <tr>
            <td style="border: 1px solid #ddd; padding: 8px;"><strong>USD/KRW:</strong></td>
            <td style="border: 1px solid #ddd; padding: 8px;">${analysis.details.em.usdkrw.toFixed(2)}${comparison ? formatChange(comparison.usdkrw.change) : ''}</td>
          </tr>
          <tr>
            <td style="border: 1px solid #ddd; padding: 8px;"><strong>EM 강세지수:</strong></td>
            <td style="border: 1px solid #ddd; padding: 8px;">${analysis.details.em.strength_index.toFixed(2)}${comparison ? formatChange(comparison.emIndex.change) : ''}</td>
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
      </body>
      </html>
    `;

    // 이모지를 HTML entity로 변환하여 깨짐 방지
    const encodedBody = encodeEmojis(emailBody);
    const encodedSubject = encodeEmojis('🌐 글로벌 유동성 알림');

    GmailApp.sendEmail(userEmail, encodedSubject, '', {
      htmlBody: encodedBody
    });

    Logger.log('✉️ 글로벌 알림 발송: ' + userEmail);

  } catch (e) {
    Logger.log('❌ 이메일 발송 오류: ' + e.message);
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

  // v4.0: 맥락 기반 DXY 해석 추가
  const contextualDxy = getContextualDXYScore(parseFloat(weekChange));

  SpreadsheetApp.getUi().alert(
    `💵 달러 인덱스 (DXY) 추세\n\n` +
    `현재: ${dxy.value.toFixed(2)}\n` +
    `주간 변화: ${weekChange > 0 ? '+' : ''}${weekChange}\n` +
    `월간 변화: ${monthChange > 0 ? '+' : ''}${monthChange}\n\n` +
    `맥락 해석: ${contextualDxy.interpretation}\n` +
    `(Fed 상태: ${contextualDxy.fedContext})`
  );
}

/**
 * v4.0: Fed 정책 상세 체크
 * QT/QE 상태, 금리 사이클, Reserve Balances, Repo 스프레드 분석
 */
function checkFedPolicy() {
  try {
    const fedPolicy = analyzeFedPolicy();
    const qt = fedPolicy.components.qtStatus;
    const rate = fedPolicy.components.rateCycle;
    const reserves = fedPolicy.components.reserves;
    const repo = fedPolicy.components.repoSpread;

    const message =
      `🏦 Fed 정책 분석 (v4.0)\n\n` +
      `=== 종합 점수: ${fedPolicy.totalScore} ===\n` +
      `신호: ${fedPolicy.overallSignal}\n\n` +
      `--- QT/QE 상태 ---\n` +
      `상태: ${qt.status} (${qt.signal})\n` +
      `월간 WALCL 변화: ${(qt.monthChange/1000).toFixed(0)}B$\n` +
      `점수 기여: ${qt.score > 0 ? '+' : ''}${qt.score}\n\n` +
      `--- 금리 사이클 ---\n` +
      `사이클: ${rate.cycle} (${rate.signal})\n` +
      `현재 금리: ${rate.currentRate}%\n` +
      `3개월 변화: ${rate.change3m}bp\n` +
      `점수 기여: ${rate.score > 0 ? '+' : ''}${rate.score}\n\n` +
      `--- Reserve Balances ---\n` +
      `상태: ${reserves.status} (${reserves.signal})\n` +
      `현재: ${reserves.currentTrillion}T$\n` +
      `점수 기여: ${reserves.score > 0 ? '+' : ''}${reserves.score}\n\n` +
      `--- Repo 스프레드 ---\n` +
      `상태: ${repo.status} (${repo.signal})\n` +
      `SOFR-IORB: ${repo.spread.toFixed(1)}bp\n` +
      `점수 기여: ${repo.score > 0 ? '+' : ''}${repo.score}`;

    SpreadsheetApp.getUi().alert(message);
    Logger.log('✅ Fed 정책 체크 완료');

  } catch (e) {
    Logger.log(`❌ Fed 정책 체크 오류: ${e.message}`);
    SpreadsheetApp.getUi().alert(`❌ 오류: ${e.message}`);
  }
}

/**
 * v4.0: 시장 국면 체크
 * 현재 시장 국면 및 동적 가중치 표시
 */
function checkMarketRegime() {
  try {
    const regime = detectMarketRegime();
    const weights = regime.weights;

    const message =
      `📊 시장 국면 분석 (v4.0)\n\n` +
      `=== 현재 국면: ${regime.regime} ===\n\n` +
      `--- DXY 변동성 ---\n` +
      `1주일: ${regime.dxyVolatility1w.toFixed(2)}pt\n` +
      `1개월: ${regime.dxyVolatility1m.toFixed(2)}pt\n\n` +
      `--- Fed 상태 ---\n` +
      `금리 인하 중: ${regime.fedCutting ? '✅ 예' : '❌ 아니오'}\n` +
      `QT 종료: ${regime.qtEnded ? '✅ 예' : '❌ 아니오'}\n\n` +
      `--- 동적 가중치 ---\n` +
      `미국 요인: ${(weights.US * 100).toFixed(0)}% (기본 40%)\n` +
      `DXY 요인: ${(weights.DXY * 100).toFixed(0)}% (기본 20%)\n` +
      `중국 요인: ${(weights.China * 100).toFixed(0)}% (기본 20%)\n` +
      `일본 요인: ${(weights.Japan * 100).toFixed(0)}% (기본 10%)\n` +
      `EM 요인: ${(weights.EM * 100).toFixed(0)}% (기본 10%)`;

    SpreadsheetApp.getUi().alert(message);
    Logger.log('✅ 시장 국면 체크 완료');

  } catch (e) {
    Logger.log(`❌ 시장 국면 체크 오류: ${e.message}`);
    SpreadsheetApp.getUi().alert(`❌ 오류: ${e.message}`);
  }
}

/** ===============================================
 * 10-B) Graph 시트 - 유동성 그래프 생성
 * =============================================== */

/**
 * Global_History 데이터로 유동성 그래프 생성
 */
function createLiquidityGraph() {
  try {
    const ss = SpreadsheetApp.getActive();
    const globalHistorySheet = ss.getSheetByName(CONFIG.GLOBAL_HISTORY_SHEET);

    if (!globalHistorySheet) {
      SpreadsheetApp.getUi().alert('❌ Global_History 시트를 찾을 수 없습니다.\n\n먼저 Global_History 데이터를 생성하세요.');
      return;
    }

    // 데이터가 있는지 확인
    const lastRow = globalHistorySheet.getLastRow();
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert('❌ Global_History 시트에 데이터가 없습니다.\n\n먼저 데이터를 채우세요.');
      return;
    }

    Logger.log('=== 유동성 그래프 생성 시작 ===');

    // Graph 시트 생성 또는 가져오기
    let graphSheet = ss.getSheetByName('Graph');
    if (graphSheet) {
      // 기존 차트 모두 삭제
      const charts = graphSheet.getCharts();
      charts.forEach(chart => graphSheet.removeChart(chart));
      graphSheet.clear();
    } else {
      graphSheet = ss.insertSheet('Graph');
    }

    // 타이틀 추가
    graphSheet.getRange('A1').setValue('📊 글로벌 유동성 추세 그래프')
      .setFontSize(16)
      .setFontWeight('bold')
      .setBackground('#1f77b4')
      .setFontColor('white');
    graphSheet.getRange('A1:F1').merge();

    // Global_History에서 전체 데이터 가져오기 (헤더 포함)
    const allData = globalHistorySheet.getRange(1, 1, lastRow, 20).getValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);

    // === 차트 1: 유동성 점수 ===
    const chart1StartRow = 3;
    const chart1Data = [
      [headers[0], headers[17]], // 타임스탬프, 유동성 점수
      ...dataRows.map(row => [row[0], row[17]])
    ];
    graphSheet.getRange(chart1StartRow, 1, chart1Data.length, 2).setValues(chart1Data);
    graphSheet.getRange(chart1StartRow, 1, 1, 2).setFontWeight('bold').setBackground('#f0f0f0');

    const mainChart = graphSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(graphSheet.getRange(chart1StartRow, 1, chart1Data.length, 2))
      .setPosition(chart1StartRow + chart1Data.length + 2, 1, 0, 0)
      .setOption('title', '유동성 점수 추세')
      .setOption('width', 1100)
      .setOption('height', 450)
      .setOption('hAxis', {
        title: '날짜',
        format: 'MMM dd',
        textStyle: { fontSize: 11 }
      })
      .setOption('vAxis', {
        title: '유동성 점수',
        textStyle: { fontSize: 11 },
        gridlines: { count: 7 }
      })
      .setOption('series', {
        0: {
          color: '#2E7D32',
          lineWidth: 4,
          pointSize: 5
        }
      })
      .setOption('legend', {
        position: 'top',
        textStyle: { fontSize: 14, bold: true }
      })
      .setOption('chartArea', { width: '80%', height: '70%' })
      .setOption('curveType', 'function')
      .build();

    graphSheet.insertChart(mainChart);

    // === 차트 2: 미국 요인 (WALCL WoW, TGA WoW) ===
    const chart2StartRow = chart1StartRow + chart1Data.length + 28;
    const chart2Data = [
      [headers[0], headers[2], headers[4]], // 타임스탬프, WALCL WoW, TGA WoW
      ...dataRows.map(row => [row[0], row[2], row[4]])
    ];
    graphSheet.getRange(chart2StartRow, 1, chart2Data.length, 3).setValues(chart2Data);
    graphSheet.getRange(chart2StartRow, 1, 1, 3).setFontWeight('bold').setBackground('#f0f0f0');

    const usChart = graphSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(graphSheet.getRange(chart2StartRow, 1, chart2Data.length, 3))
      .setPosition(chart2StartRow + chart2Data.length + 2, 1, 0, 0)
      .setOption('title', '미국 유동성 요인')
      .setOption('width', 650)
      .setOption('height', 380)
      .setOption('hAxis', {
        title: '날짜',
        format: 'MMM dd',
        textStyle: { fontSize: 10 }
      })
      .setOption('vAxis', {
        title: '변화량 (억$)',
        textStyle: { fontSize: 10 }
      })
      .setOption('series', {
        0: {
          color: '#1976D2',
          lineWidth: 2.5,
          pointSize: 3
        },
        1: {
          color: '#D32F2F',
          lineWidth: 2.5,
          pointSize: 3
        }
      })
      .setOption('legend', {
        position: 'top',
        textStyle: { fontSize: 13, bold: true }
      })
      .setOption('chartArea', { width: '80%', height: '70%' })
      .build();

    graphSheet.insertChart(usChart);

    // === 차트 3: 글로벌 요인 (DXY WoW, 중국 M2, EM 지수) ===
    const chart3StartRow = chart2StartRow + chart2Data.length + 28;
    const chart3Data = [
      [headers[0], headers[7], headers[8], headers[16]], // 타임스탬프, DXY WoW, 중국 M2, EM 강세지수
      ...dataRows.map(row => [row[0], row[7], row[8], row[16]])
    ];
    graphSheet.getRange(chart3StartRow, 1, chart3Data.length, 4).setValues(chart3Data);
    graphSheet.getRange(chart3StartRow, 1, 1, 4).setFontWeight('bold').setBackground('#f0f0f0');

    const globalChart = graphSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(graphSheet.getRange(chart3StartRow, 1, chart3Data.length, 4))
      .setPosition(chart3StartRow + chart3Data.length + 2, 1, 0, 0)
      .setOption('title', '글로벌 요인 (DXY WoW, 중국 M2, EM 지수)')
      .setOption('width', 650)
      .setOption('height', 380)
      .setOption('hAxis', {
        title: '날짜',
        format: 'MMM dd',
        textStyle: { fontSize: 10 }
      })
      .setOption('vAxis', {
        title: '지수값',
        textStyle: { fontSize: 10 }
      })
      .setOption('series', {
        0: {
          color: '#F57C00',
          lineWidth: 2.5,
          pointSize: 3
        },
        1: {
          color: '#C62828',
          lineWidth: 2.5,
          pointSize: 3
        },
        2: {
          color: '#6A1B9A',
          lineWidth: 2.5,
          pointSize: 3
        }
      })
      .setOption('legend', {
        position: 'top',
        textStyle: { fontSize: 13, bold: true }
      })
      .setOption('chartArea', { width: '80%', height: '70%' })
      .build();

    graphSheet.insertChart(globalChart);

    // === 차트 4: 일본 요인 (USD/JPY) ===
    const chart4StartRow = chart3StartRow + chart3Data.length + 28;
    const chart4Data = [
      [headers[0], headers[11]], // 타임스탬프, USD/JPY
      ...dataRows.map(row => [row[0], row[11]])
    ];
    graphSheet.getRange(chart4StartRow, 1, chart4Data.length, 2).setValues(chart4Data);
    graphSheet.getRange(chart4StartRow, 1, 1, 2).setFontWeight('bold').setBackground('#f0f0f0');

    const japanChart = graphSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(graphSheet.getRange(chart4StartRow, 1, chart4Data.length, 2))
      .setPosition(chart4StartRow + chart4Data.length + 2, 1, 0, 0)
      .setOption('title', '일본 요인 (USD/JPY)')
      .setOption('width', 650)
      .setOption('height', 380)
      .setOption('hAxis', {
        title: '날짜',
        format: 'MMM dd',
        textStyle: { fontSize: 10 }
      })
      .setOption('vAxis', {
        title: 'USD/JPY',
        textStyle: { fontSize: 10 }
      })
      .setOption('series', {
        0: {
          color: '#00796B',
          lineWidth: 2.5,
          pointSize: 3
        }
      })
      .setOption('legend', {
        position: 'top',
        textStyle: { fontSize: 13, bold: true }
      })
      .setOption('chartArea', { width: '80%', height: '70%' })
      .build();

    graphSheet.insertChart(japanChart);

    // === 차트 5: 통합 차트 - 모든 주요 요인 (정규화) ===
    const chart5StartRow = chart4StartRow + chart4Data.length + 28;

    // 정규화된 데이터 생성
    const cols = {
      score: 17,    // 유동성 점수
      walcl: 2,     // WALCL WoW
      dxy: 7,       // DXY WoW
      chinaM2: 8,   // 중국 M2
      usdjpy: 11,   // USD/JPY
      em: 16        // EM 지수
    };

    // 각 컬럼의 최소/최대값 찾기
    const ranges = {};
    for (const [key, idx] of Object.entries(cols)) {
      const values = dataRows.map(row => row[idx]);
      ranges[key] = {
        min: Math.min(...values),
        max: Math.max(...values)
      };
    }

    // 정규화 함수 (0-100 스케일)
    const normalize = (value, min, max) => {
      if (max === min) return 50;
      return ((value - min) / (max - min)) * 100;
    };

    // 정규화된 데이터 배열 생성 (헤더 포함)
    const chart5Data = [
      ['날짜', '유동성 점수', 'WALCL WoW', 'DXY WoW', '중국 M2', 'USD/JPY', 'EM 지수'], // 헤더
      ...dataRows.map(row => [
        row[0], // 날짜
        normalize(row[cols.score], ranges.score.min, ranges.score.max),
        normalize(row[cols.walcl], ranges.walcl.min, ranges.walcl.max),
        normalize(row[cols.dxy], ranges.dxy.min, ranges.dxy.max),
        normalize(row[cols.chinaM2], ranges.chinaM2.min, ranges.chinaM2.max),
        normalize(row[cols.usdjpy], ranges.usdjpy.min, ranges.usdjpy.max),
        normalize(row[cols.em], ranges.em.min, ranges.em.max)
      ])
    ];

    // 정규화된 데이터를 Graph 시트에 쓰기
    graphSheet.getRange(chart5StartRow, 1, chart5Data.length, 7).setValues(chart5Data);
    graphSheet.getRange(chart5StartRow, 1, 1, 7).setFontWeight('bold').setBackground('#f0f0f0');

    // 통합 차트 생성
    const integratedChart = graphSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(graphSheet.getRange(chart5StartRow, 1, chart5Data.length, 7))
      .setPosition(chart5StartRow + chart5Data.length + 2, 1, 0, 0)
      .setOption('title', '모든 요인 통합 뷰 (정규화 0-100)')
      .setOption('width', 1350)
      .setOption('height', 500)
      .setOption('hAxis', {
        title: '날짜',
        format: 'MMM dd',
        textStyle: { fontSize: 11 }
      })
      .setOption('vAxis', {
        title: '정규화 값 (0-100)',
        textStyle: { fontSize: 11 }
      })
      .setOption('series', {
        0: { // 유동성 점수
          color: '#2E7D32',
          lineWidth: 5,
          pointSize: 0
        },
        1: { // WALCL WoW
          color: '#1976D2',
          lineWidth: 1.5,
          pointSize: 0
        },
        2: { // DXY WoW
          color: '#F57C00',
          lineWidth: 1.5,
          pointSize: 0
        },
        3: { // 중국 M2
          color: '#C62828',
          lineWidth: 1.5,
          pointSize: 0
        },
        4: { // USD/JPY
          color: '#00796B',
          lineWidth: 1.5,
          pointSize: 0
        },
        5: { // EM 지수
          color: '#6A1B9A',
          lineWidth: 1.5,
          pointSize: 0
        }
      })
      .setOption('legend', {
        position: 'top',
        textStyle: { fontSize: 14, bold: true }
      })
      .setOption('chartArea', { width: '85%', height: '75%' })
      .setOption('curveType', 'function')
      .build();

    graphSheet.insertChart(integratedChart);

    // Graph 시트를 활성화
    ss.setActiveSheet(graphSheet);

    Logger.log('✅ 그래프 생성 완료');
    SpreadsheetApp.getUi().alert('✅ 그래프가 생성되었습니다!\n\n"Graph" 시트를 확인하세요.');

  } catch (e) {
    Logger.log(`❌ 그래프 생성 오류: ${e.message}`);
    SpreadsheetApp.getUi().alert(`❌ 오류: ${e.message}`);
  }
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
        <div class="score ${analysis.score >= 20 ? 'positive' : analysis.score <= -20 ? 'negative' : 'neutral'}">
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
        <em>생성 시간: ${new Date().toLocaleString('ko-KR', {timeZone: 'Asia/Seoul'})}</em>
      </p>
    `).setWidth(600).setHeight(800);
    
    ui.showModalDialog(html, '글로벌 유동성 대시보드');
    
  } catch (e) {
    Logger.log(`❌ 대시보드 생성 오류: ${e.message}`);
    SpreadsheetApp.getUi().alert(`❌ 오류: ${e.message}`);
  }
}

/** ===============================================
 * 점수 계산 가이드 시트 생성
 * =============================================== */

function createScoringGuide() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheetName = 'Scoring_Guide';

    // 기존 시트가 있으면 삭제
    let guideSheet = ss.getSheetByName(sheetName);
    if (guideSheet) {
      ss.deleteSheet(guideSheet);
    }

    // 새 시트 생성
    guideSheet = ss.insertSheet(sheetName);

    // 현재 행 추적
    let currentRow = 1;

    // ============= 타이틀 =============
    guideSheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue('📊 글로벌 유동성 점수 계산 가이드 v3.1')
      .setFontSize(16)
      .setFontWeight('bold')
      .setBackground('#1f77b4')
      .setFontColor('white')
      .setHorizontalAlignment('center');
    currentRow += 2;

    // ============= 개요 =============
    guideSheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue('📌 점수 계산 개요')
      .setFontSize(12)
      .setFontWeight('bold')
      .setBackground('#d0e0f0');
    currentRow++;

    guideSheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue('총 5개 요인을 분석하여 -120점 ~ +105점 범위의 종합 점수를 산출합니다.')
      .setWrap(true);
    currentRow += 2;

    // ============= 가중치 테이블 =============
    guideSheet.getRange(currentRow, 1, 1, 4).setValues([['요인', '가중치', '최대점수', '설명']])
      .setFontWeight('bold')
      .setBackground('#e6e6e6');
    currentRow++;

    const weights = [
      ['미국 요인 (WALCL + TGA + ON RRP)', '40%', '+40 / -45', 'Fed 자산, 재무부 계좌, 역레포'],
      ['달러 요인 (DXY)', '20%', '+25 / -25', '달러 인덱스 주간 변화'],
      ['중국 요인 (M2)', '20%', '+20 / -20', 'M2 통화 공급 성장률'],
      ['일본 요인 (USD/JPY)', '10%', '+5 / -15', '엔화 환율 및 캐리 리스크'],
      ['신흥국 요인 (EM Index)', '10%', '+15 / -15', '신흥국 통화 강세 지수']
    ];

    guideSheet.getRange(currentRow, 1, weights.length, 4).setValues(weights);
    currentRow += weights.length + 2;

    // ============= 미국 요인 (40%) =============
    guideSheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue('🇺🇸 미국 요인 (40% 가중치)')
      .setFontSize(12)
      .setFontWeight('bold')
      .setBackground('#d0e0f0');
    currentRow++;

    // 1. WALCL
    guideSheet.getRange(currentRow, 1).setValue('1. WALCL (연준 자산) 주간 변화')
      .setFontWeight('bold');
    currentRow++;

    const walclTable = [
      ['구간', '점수', '의미', '시장 영향'],
      ['> +500억$', '+20', '강한 확장 (QE 재개)', '🚀 Risk-ON'],
      ['+100억 ~ +500억$', '+10', '완만한 확장', '✅ 긍정적'],
      ['-100억 ~ +100억$', '0', '중립 (변화 없음)', '⚖️ 중립'],
      ['-500억 ~ -100억$', '-10', '완만한 긴축 (QT)', '⚠️ 주의'],
      ['< -500억$', '-20', '강한 긴축 (적극적 QT)', '🔴 Risk-OFF']
    ];

    guideSheet.getRange(currentRow, 1, walclTable.length, 4).setValues(walclTable);
    guideSheet.getRange(currentRow, 1, 1, 4).setFontWeight('bold').setBackground('#f0f0f0');
    currentRow += walclTable.length + 1;

    // 2. TGA
    guideSheet.getRange(currentRow, 1).setValue('2. TGA (재무부 계좌) 주간 변화')
      .setFontWeight('bold');
    currentRow++;

    const tgaTable = [
      ['구간', '점수', '의미', '시장 영향'],
      ['< -1000억$', '+10', '대규모 지출 (유동성 공급)', '🚀 Risk-ON'],
      ['-1000억 ~ -500억$', '+5', '중간 지출', '✅ 긍정적'],
      ['-500억 ~ +500억$', '0', '중립', '⚖️ 중립'],
      ['+500억 ~ +1000억$', '-5', '중간 축적 (채권 발행)', '⚠️ 주의'],
      ['> +1000억$', '-10', '대규모 축적 (유동성 흡수)', '🔴 Risk-OFF']
    ];

    guideSheet.getRange(currentRow, 1, tgaTable.length, 4).setValues(tgaTable);
    guideSheet.getRange(currentRow, 1, 1, 4).setFontWeight('bold').setBackground('#f0f0f0');
    currentRow += tgaTable.length + 1;

    // 3. ON RRP
    guideSheet.getRange(currentRow, 1).setValue('3. ON RRP (Overnight Reverse Repo) 잔고')
      .setFontWeight('bold');
    currentRow++;

    const rrpTable = [
      ['구간', '점수', '의미', '시장 영향'],
      ['< 1000억$', '+10', '완전 활용 (유동성 긴장)', '🚀 Risk-ON'],
      ['1000억 ~ 2000억$', '+5', '적정 수준', '✅ 건강'],
      ['2000억 ~ 3000억$', '0', '중립', '⚖️ 중립'],
      ['3000억 ~ 5000억$', '-10', '과잉 유동성 (리스크)', '⚠️ 버블 위험'],
      ['> 5000억$', '-15', '극도의 과잉', '🔴 시스템 리스크']
    ];

    guideSheet.getRange(currentRow, 1, rrpTable.length, 4).setValues(rrpTable);
    guideSheet.getRange(currentRow, 1, 1, 4).setFontWeight('bold').setBackground('#f0f0f0');
    currentRow += rrpTable.length + 2;

    // ============= 달러 요인 (20%) =============
    guideSheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue('💵 달러 요인 (20% 가중치)')
      .setFontSize(12)
      .setFontWeight('bold')
      .setBackground('#d0e0f0');
    currentRow++;

    guideSheet.getRange(currentRow, 1).setValue('DXY (달러 인덱스) 주간 변화')
      .setFontWeight('bold');
    currentRow++;

    const dxyTable = [
      ['구간', '점수', '의미', '시장 영향'],
      ['< -2.0 포인트', '+25', '급격한 달러 약세', '🚀🚀 강한 Risk-ON'],
      ['-2.0 ~ -1.0', '+20', '달러 약세', '✅ Risk-ON'],
      ['-1.0 ~ +1.0', '0', '중립', '⚖️ 중립'],
      ['+1.0 ~ +2.0', '-20', '달러 강세', '⚠️ Risk-OFF'],
      ['> +2.0 포인트', '-25', '급격한 달러 강세', '🔴🔴 강한 Risk-OFF']
    ];

    guideSheet.getRange(currentRow, 1, dxyTable.length, 4).setValues(dxyTable);
    guideSheet.getRange(currentRow, 1, 1, 4).setFontWeight('bold').setBackground('#f0f0f0');
    currentRow += dxyTable.length + 2;

    // ============= 중국 요인 (20%) =============
    guideSheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue('🇨🇳 중국 요인 (20% 가중치)')
      .setFontSize(12)
      .setFontWeight('bold')
      .setBackground('#d0e0f0');
    currentRow++;

    guideSheet.getRange(currentRow, 1).setValue('M2 (광의통화) YoY 성장률')
      .setFontWeight('bold');
    currentRow++;

    const chinaTable = [
      ['구간', '점수', '의미', '시장 영향'],
      ['> 12%', '+20', '과잉 확대 (부양 정책)', '🚀 강한 성장'],
      ['10% ~ 12%', '+15', '적정 확대 (건강한 성장)', '✅ 긍정적'],
      ['8% ~ 10%', '0', '중립 (정상 범위)', '⚖️ 중립'],
      ['6% ~ 8%', '-10', '성장 둔화', '⚠️ 경기 약화'],
      ['< 6%', '-20', '유동성 경색', '🔴 심각한 둔화']
    ];

    guideSheet.getRange(currentRow, 1, chinaTable.length, 4).setValues(chinaTable);
    guideSheet.getRange(currentRow, 1, 1, 4).setFontWeight('bold').setBackground('#f0f0f0');
    currentRow += chinaTable.length + 2;

    // ============= 일본 요인 (10%) =============
    guideSheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue('🇯🇵 일본 요인 (10% 가중치)')
      .setFontSize(12)
      .setFontWeight('bold')
      .setBackground('#d0e0f0');
    currentRow++;

    guideSheet.getRange(currentRow, 1).setValue('USD/JPY 환율 수준')
      .setFontWeight('bold');
    currentRow++;

    const japanTable = [
      ['구간', '점수', '의미', '시장 영향'],
      ['< 130', '+5', '언와인드 완료', '✅ 약한 호재'],
      ['130 ~ 145', '0', '안정 범위', '⚖️ 중립'],
      ['145 ~ 150', '-5', '주의 수준', '⚠️ 모니터링'],
      ['150 ~ 155', '-10', '고위험 (캐리 리스크)', '🔴 주의'],
      ['> 155', '-15', '극도의 캐리 리스크', '🔴🔴 언와인드 위험']
    ];

    guideSheet.getRange(currentRow, 1, japanTable.length, 4).setValues(japanTable);
    guideSheet.getRange(currentRow, 1, 1, 4).setFontWeight('bold').setBackground('#f0f0f0');
    currentRow += japanTable.length + 2;

    // ============= 신흥국 요인 (10%) =============
    guideSheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue('🌏 신흥국 요인 (10% 가중치)')
      .setFontSize(12)
      .setFontWeight('bold')
      .setBackground('#d0e0f0');
    currentRow++;

    guideSheet.getRange(currentRow, 1).setValue('EM 통화 강세 지수 (KRW, BRL, MXN 평균)')
      .setFontWeight('bold');
    currentRow++;

    const emTable = [
      ['구간', '점수', '의미', '시장 영향'],
      ['> +2.0%', '+15', '강한 EM 강세', '🚀 Risk-ON'],
      ['+1.0% ~ +2.0%', '+10', '완만한 EM 강세', '✅ 긍정적'],
      ['-1.0% ~ +1.0%', '0', '중립', '⚖️ 중립'],
      ['-2.0% ~ -1.0%', '-10', '완만한 EM 약세', '⚠️ 자금 유출'],
      ['< -2.0%', '-15', '강한 EM 약세', '🔴 위기 조짐']
    ];

    guideSheet.getRange(currentRow, 1, emTable.length, 4).setValues(emTable);
    guideSheet.getRange(currentRow, 1, 1, 4).setFontWeight('bold').setBackground('#f0f0f0');
    currentRow += emTable.length + 2;

    // ============= 최종 점수 해석 =============
    guideSheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue('🎯 최종 점수 해석 (7단계)')
      .setFontSize(12)
      .setFontWeight('bold')
      .setBackground('#d0e0f0');
    currentRow++;

    const signalTable = [
      ['점수 범위', '신호', '투자 권장', '역사적 사례'],
      ['80점 이상', '🚀🚀 슈퍼 유동성', '공격적 Risk-ON: 레버리지 ETF, 성장주, BTC', '2020년 3월 (코로나 QE)'],
      ['50 ~ 80점', '🚀 극도의 유동성', '적극적 Risk-ON: 성장주, 신흥국, 원자재', '2024년 4월 랠리'],
      ['20 ~ 50점', '✅ 높은 유동성', '위험자산 유지/확대, 밸류/그로스 균형', '2023년 하반기'],
      ['-20 ~ +20점', '⚖️ 중립', '포트폴리오 균형 유지, 관망', '2024년 상반기'],
      ['-50 ~ -20점', '⚠️ 긴축', '현금/채권 증대, 방어주 선호', '2022년 상반기 (금리인상)'],
      ['-80 ~ -50점', '🔴 극도의 긴축', '방어적 포지션, 달러/금/국채', '2022년 10월 (바닥)'],
      ['-80점 이하', '🔴🔴 위기 모드', '현금 확보, 손절 고려, 변동성 헤지', '2008년 9월 (리먼)']
    ];

    guideSheet.getRange(currentRow, 1, signalTable.length, 4).setValues(signalTable);
    guideSheet.getRange(currentRow, 1, 1, 4).setFontWeight('bold').setBackground('#e6e6e6');

    // 신호별 배경색
    for (let i = 0; i < signalTable.length - 1; i++) {
      const rowIdx = currentRow + i + 1;
      if (i === 0) guideSheet.getRange(rowIdx, 1, 1, 4).setBackground('#00FF00'); // 슈퍼
      else if (i === 1) guideSheet.getRange(rowIdx, 1, 1, 4).setBackground('#90EE90'); // 극도
      else if (i === 2) guideSheet.getRange(rowIdx, 1, 1, 4).setBackground('#D4EDDA'); // 높음
      else if (i === 3) guideSheet.getRange(rowIdx, 1, 1, 4).setBackground('#FFFFE0'); // 중립
      else if (i === 4) guideSheet.getRange(rowIdx, 1, 1, 4).setBackground('#FFE4B5'); // 긴축
      else if (i === 5) guideSheet.getRange(rowIdx, 1, 1, 4).setBackground('#FFB6C1'); // 극도긴축
      else if (i === 6) guideSheet.getRange(rowIdx, 1, 1, 4).setBackground('#FF6B6B'); // 위기
    }

    currentRow += signalTable.length + 2;

    // ============= 참고 사항 =============
    guideSheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue('📝 참고 사항')
      .setFontSize(12)
      .setFontWeight('bold')
      .setBackground('#d0e0f0');
    currentRow++;

    const notes = [
      ['• 최대 가능 점수: +105점 (모든 요인 극도로 긍정적)'],
      ['• 최소 가능 점수: -120점 (모든 요인 극도로 부정적)'],
      ['• 실시간 업데이트: "📊 Global Liquidity" 메뉴 → "🔄 전체 업데이트"'],
      ['• 알림 설정: "🔔 알림 설정/해제"에서 2시간마다 자동 체크 가능'],
      ['• 히스토리 확인: Global_History 시트에서 과거 점수 추이 확인'],
      ['• 문의 및 수정: v3.1 (2025-11-13) - 세밀한 5단계 로직 적용']
    ];

    guideSheet.getRange(currentRow, 1, notes.length, 6).setValues(notes.map(n => [n[0], '', '', '', '', '']));

    // 열 너비 조정
    guideSheet.setColumnWidth(1, 200);
    guideSheet.setColumnWidth(2, 100);
    guideSheet.setColumnWidth(3, 250);
    guideSheet.setColumnWidth(4, 200);

    // 시트를 맨 앞으로 이동
    ss.setActiveSheet(guideSheet);
    ss.moveActiveSheet(1);

    SpreadsheetApp.getUi().alert('✅ 점수 계산 가이드 시트가 생성되었습니다!\n\n"Scoring_Guide" 시트를 확인하세요.');
    Logger.log('✅ Scoring_Guide 시트 생성 완료');

  } catch (e) {
    Logger.log(`❌ 가이드 시트 생성 오류: ${e.message}`);
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
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📅 History 업데이트')
      .addItem('📈 History 시트 채우기 (1월~현재)', 'populateHistoryFromJanuary')
      .addItem('🌍 Global_History 시트 채우기 (1월~현재)', 'populateGlobalHistoryFromJanuary'))
    .addSeparator()
    .addItem('📉 유동성 그래프 생성', 'createLiquidityGraph')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🔍 개별 체크')
      .addItem('🇨🇳 중국 유동성', 'checkChinaLiquidity')
      .addItem('🇯🇵 엔캐리 리스크', 'checkJapanRisk')
      .addItem('💵 TGA 분석', 'checkTGADetail')
      .addItem('📈 DXY 추세', 'checkDXYTrend')
      .addSeparator()
      .addItem('🏦 Fed 정책 분석 (v4.0)', 'checkFedPolicy')
      .addItem('📊 시장 국면 분석 (v4.0)', 'checkMarketRegime'))
    .addSeparator()
    .addItem('📊 종합 대시보드', 'createGlobalDashboard')
    .addItem('🔔 알림 설정/해제', 'setupGlobalAlerts')
    .addItem('🧪 알림 테스트', 'testGlobalAlert')
    .addItem('⏰ 일일 자동갱신', 'createDailyTrigger')
    .addSeparator()
    .addItem('📋 캐시 초기화', 'clearAllCache')
    .addItem('📖 점수 계산 가이드', 'createScoringGuide')
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
  const keys = [
    'CHINA_LIQUIDITY',
    'FRED_WALCL',
    'FRED_WTREGEN',
    'FRED_RRPONTSYD',
    'FRED_DTWEXBGS',
    'FRED_DEXCHUS',
    'FRED_DEXJPUS',
    'FRED_DEXKOUS',
    'FRED_DEXBZUS',
    'FRED_DEXMXUS'
  ];
  cache.removeAll(keys);
  SpreadsheetApp.getUi().alert('✅ 모든 캐시가 초기화되었습니다.');
}

function showHelp() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial; font-size: 12px; padding: 15px; }
      h3 { color: #1f77b4; margin-top: 15px; }
      h4 { color: #2c5282; margin-top: 10px; }
      code { background: #f5f5f5; padding: 3px 6px; border-radius: 3px; }
      li { margin: 8px 0; }
      .new { color: #38a169; font-weight: bold; }
    </style>

    <h2>📊 Global Liquidity Monitor v4.0 도움말</h2>

    <h3>주요 기능</h3>
    <ul>
      <li><strong>전체 업데이트:</strong> 미국 + 글로벌 데이터 갱신 및 히스토리 누적</li>
      <li><strong>글로벌 분석:</strong> 종합 유동성 점수 계산</li>
      <li><strong>History 업데이트:</strong> 올해 1월부터 현재까지 데이터를 History/Global_History 시트에 일괄 추가</li>
      <li><strong>유동성 그래프 생성:</strong> Global_History 데이터로 유동성 점수 및 요인별 그래프 생성</li>
      <li><strong>개별 체크:</strong> 중국, 일본, TGA, DXY 상세 분석</li>
      <li><strong>알림 설정:</strong> 2시간마다 자동 체크 (해제 가능)</li>
    </ul>

    <h3 class="new">v4.0 신규 기능</h3>
    <ul>
      <li><strong>🏦 Fed 정책 분석:</strong> QT/QE 상태, 금리 사이클, Reserve Balances, Repo 스프레드 종합 분석</li>
      <li><strong>📊 시장 국면 분석:</strong> DXY 변동성, Fed 정책 상태 기반 시장 국면 자동 감지</li>
      <li><strong>동적 가중치:</strong> 시장 국면에 따라 요인별 가중치 자동 조정
        <ul>
          <li>DXY 고변동성 국면: DXY 가중치 35%로 상향</li>
          <li>Fed 적극 완화 국면: 미국 요인 50%로 상향</li>
        </ul>
      </li>
      <li><strong>맥락 기반 DXY 해석:</strong> Fed 완화 중 DXY 하락 = 긍정적 신호로 해석</li>
      <li><strong>추세 기반 알림:</strong> 모멘텀 급변, 추세 전환, Fed 정책 변화 감지</li>
    </ul>

    <h3>히스토리 기록</h3>
    <ul>
      <li><strong>History:</strong> 미국 유동성 지표 타임시리즈</li>
      <li><strong>Global_History:</strong> 글로벌 유동성 분석 타임시리즈</li>
      <li><strong>Alert_History:</strong> 발생한 알림 전체 기록</li>
    </ul>

    <h3>유동성 점수 (7단계)</h3>
    <ul>
      <li><strong>80점 이상:</strong> 🚀🚀 슈퍼 유동성 (공격적 Risk-ON)</li>
      <li><strong>50-80점:</strong> 🚀 극도의 유동성 (적극적 Risk-ON)</li>
      <li><strong>20-50점:</strong> ✅ 높은 유동성 (위험자산 선호)</li>
      <li><strong>-20~20점:</strong> ⚖️ 중립 (관망)</li>
      <li><strong>-50~-20점:</strong> ⚠️ 긴축 (방어주 선호)</li>
      <li><strong>-80~-50점:</strong> 🔴 극도의 긴축 (Risk-OFF)</li>
      <li><strong>-80점 이하:</strong> 🔴🔴 위기 모드 (현금 확보)</li>
    </ul>

    <h3>기본 가중치 (동적 조정됨)</h3>
    <ul>
      <li>미국 요인: 40% (Fed 완화시 최대 50%)</li>
      <li>달러 지수: 20% (고변동성시 최대 35%)</li>
      <li>중국: 20%</li>
      <li>일본: 10%</li>
      <li>신흥국: 10%</li>
    </ul>

    <h3 class="new">Fed 정책 지표 (v4.0)</h3>
    <ul>
      <li><strong>QT/QE 상태:</strong> WALCL 월간 변화로 양적완화/긴축 판단</li>
      <li><strong>금리 사이클:</strong> Fed Funds Rate 변화로 완화/긴축 사이클 감지</li>
      <li><strong>Reserve Balances:</strong> 은행 준비금 수준 모니터링 ($2.9T ample threshold)</li>
      <li><strong>Repo 스프레드:</strong> SOFR-IORB 스프레드로 유동성 긴장 감지</li>
    </ul>

    <h3>시트 구성</h3>
    <ul>
      <li><strong>Live_Monitor:</strong> 미국 지표 최신값</li>
      <li><strong>Global_Liquidity:</strong> 글로벌 지표 최신값 + 시장 국면</li>
      <li><strong>History:</strong> 미국 지표 히스토리</li>
      <li><strong>Global_History:</strong> 글로벌 지표 히스토리</li>
      <li><strong>Alert_History:</strong> 알림 발생 기록</li>
      <li><strong>Graph:</strong> 유동성 추세 그래프 (메인 + 요인별)</li>
      <li><strong>Scoring_Guide:</strong> 점수 계산 방법 가이드</li>
    </ul>

    <p style="margin-top: 20px; color: #666; font-size: 11px;">
      버전: 4.0 (2026-01) | 동적 가중치 + Fed 정책 추적
    </p>
  `).setWidth(550).setHeight(850);

  ui.showModelessDialog(html, 'Global Liquidity Monitor v4.0 도움말');
}

/** ===============================================
 * 14) 테스트 함수
 * =============================================== */

function testAllSystems() {
  Logger.log('=== 전체 시스템 테스트 시작 (v4.0) ===');

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

  // 3. v4.0: Fed 정책 분석 테스트
  Logger.log('\n--- Fed 정책 분석 테스트 (v4.0) ---');
  const fedPolicy = analyzeFedPolicy();
  Logger.log(`Fed 정책 종합 점수: ${fedPolicy.totalScore}`);
  Logger.log(`Fed 정책 신호: ${fedPolicy.overallSignal}`);
  Logger.log(`  - QT/QE 상태: ${fedPolicy.components.qtStatus.status}`);
  Logger.log(`  - 금리 사이클: ${fedPolicy.components.rateCycle.cycle}`);
  Logger.log(`  - Reserve 상태: ${fedPolicy.components.reserves.status}`);
  Logger.log(`  - Repo 스프레드: ${fedPolicy.components.repoSpread.status}`);

  // 4. v4.0: 시장 국면 감지 테스트
  Logger.log('\n--- 시장 국면 감지 테스트 (v4.0) ---');
  const regime = detectMarketRegime();
  Logger.log(`시장 국면: ${regime.regime}`);
  Logger.log(`Fed 금리 인하 중: ${regime.fedCutting}`);
  Logger.log(`QT 종료: ${regime.qtEnded}`);
  Logger.log(`동적 가중치:`);
  Logger.log(`  - US: ${(regime.weights.US * 100).toFixed(0)}%`);
  Logger.log(`  - DXY: ${(regime.weights.DXY * 100).toFixed(0)}%`);
  Logger.log(`  - China: ${(regime.weights.China * 100).toFixed(0)}%`);

  // 5. 종합 분석
  Logger.log('\n--- 종합 분석 테스트 ---');
  const analysis = analyzeGlobalLiquidity();
  Logger.log(`유동성 점수: ${analysis.score}`);
  Logger.log(`신호: ${analysis.signal}`);
  Logger.log(`시장 국면: ${analysis.regime}`);

  // 6. v4.0: 맥락 기반 DXY 해석 테스트
  if (analysis.details.dxy.contextual) {
    Logger.log('\n--- 맥락 기반 DXY 해석 (v4.0) ---');
    Logger.log(`DXY 변화: ${analysis.details.dxy.change.toFixed(2)}`);
    Logger.log(`맥락 해석: ${analysis.details.dxy.contextual.interpretation}`);
    Logger.log(`Fed 맥락: ${analysis.details.dxy.contextual.fedContext}`);
  }

  // 7. 히스토리 기록 테스트
  Logger.log('\n--- 히스토리 기록 테스트 ---');
  logGlobalHistory(analysis);

  Logger.log('\n=== 테스트 완료 (v4.0) ===');
}