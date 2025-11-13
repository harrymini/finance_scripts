# CLAUDE.md - AI Assistant Guide for Global Liquidity Monitor

## Project Overview

**Repository**: `finance_scripts`
**Type**: Google Apps Script (JavaScript)
**Purpose**: Automated global liquidity monitoring and alerting system
**Primary Language**: JavaScript (Google Apps Script API)
**UI Language**: Korean

### What This Project Does

This is a comprehensive financial data monitoring system that:
- Tracks US liquidity indicators (Federal Reserve data)
- Monitors global liquidity conditions (China, Japan, Emerging Markets)
- Calculates composite liquidity scores
- Sends automated alerts based on threshold conditions
- Maintains historical data in Google Sheets
- Provides investment recommendations based on liquidity conditions

---

## Repository Structure

```
finance_scripts/
├── .git/                    # Git repository metadata
├── app_script.js           # Main Google Apps Script file (1,375 lines)
└── CLAUDE.md              # This documentation file
```

### File Breakdown

**app_script.js** - Single-file architecture containing:
- Configuration constants (lines 13-56)
- Data fetching functions (FRED API, SRF API)
- Analysis functions (China, Japan, TGA, EM markets)
- Liquidity scoring algorithm
- Alert system
- Google Sheets management
- UI menu and helper functions

---

## Architecture & Key Components

### 1. Configuration (`CONFIG` object - lines 13-56)

```javascript
CONFIG = {
  SHEET_NAME: 'Live_Monitor',           // Main US liquidity dashboard
  HISTORY_SHEET: 'History',             // US liquidity history
  GLOBAL_SHEET: 'Global_Liquidity',     // Global liquidity dashboard
  GLOBAL_HISTORY_SHEET: 'Global_History', // Global liquidity history
  ALERT_HISTORY_SHEET: 'Alert_History', // Alert log
  CACHE_TIME: 300000,                   // 5-minute cache
  FRED_IDS: { ... },                    // US indicator IDs
  GLOBAL_FRED_IDS: { ... }              // Global indicator IDs
}
```

### 2. Core Function Groups

#### Data Collection (lines 62-212)
- `getFredData(fredId, useCache)` - Fetches latest data from FRED
- `getFredDataHistorical(fredId, daysAgo)` - Historical data retrieval
- `getSRFData()` - Standing Repo Facility data from NY Fed

#### Regional Analysis (lines 218-408)
- `getChinaLiquidity()` - M2 growth, credit, FX reserves
- `getJapanLiquidity()` - USD/JPY, JGB yields, carry trade risk
- `getTGAAnalysis()` - Treasury General Account analysis
- `getEmergingMarketsFX()` - EM currency strength index

#### Global Analysis (lines 414-543)
- `analyzeGlobalLiquidity()` - Composite liquidity scoring (0-100)
- Weighting: US 40%, DXY 20%, China 20%, Japan 10%, EM 10%

#### History Management (lines 573-628, 699-713, 950-1007)
- `logGlobalHistory(analysis)` - Records global metrics
- Auto-appends to History sheets on each update

#### Alert System (lines 820-1076)
- `setupGlobalAlerts()` - Enable/disable 2-hour alerts
- `checkGlobalAlerts()` - Evaluates conditions
- `sendGlobalAlert(alerts, analysis)` - Email notifications
- `logAlertHistory(alerts, analysis)` - Records alerts

#### Main Update Function (lines 634-728)
- `updateLiveMonitor()` - Primary update orchestrator
- Updates Live_Monitor sheet
- Appends to History
- Triggers global analysis

---

## Data Sources & APIs

### FRED (Federal Reserve Economic Data)
**Base URL**: `https://fred.stlouisfed.org/graph/fredgraph.csv`

**US Indicators**:
- `WALCL` - Fed Total Assets
- `WTREGEN` - Treasury General Account (TGA)
- `RRPONTSYD` - Overnight Reverse Repo (ON RRP)
- `SOFR`, `EFFR`, `IORB` - Interest rates
- `DGS10` - US 10Y Treasury

**Global Indicators**:
- `DTWEXBGS` - Dollar Index (DXY)
- `MABMM301CNM657S` - China M2 YoY
- `QCNLOANTOPRIV` - China Loans to Private Sector
- `TRESEGCNM052N` - China FX Reserves
- `DEXJPUS` - USD/JPY
- `IRLTLT01JPM156N` - JGB 10Y
- `DEXKOUS`, `DEXBZUS`, `DEXMXUS` - EM FX rates
- `VIXCLS` - VIX Index

### New York Fed API
**URL**: `https://markets.newyorkfed.org/api/rp/all/all/results/latest/1.json`
- Standing Repo Facility operations data

### Caching Strategy
- 5-minute cache for FRED data (`CacheService.getScriptCache()`)
- 1-hour cache for global indicators
- 24-hour cache for SRF data

---

## Liquidity Scoring Algorithm

### Composite Score Calculation (lines 442-462)

**Initialization**: `liquidityScore = 0`

**US Factors (40% weight)**:
- WALCL WoW increasing: +20
- TGA decreasing >$10B: +10
- ON RRP < $200B: +10

**Dollar Factors (20% weight)**:
- DXY WoW < -1: +20
- DXY WoW > +1: -20

**China Factors (20% weight)**:
- M2 growth > 10%: +20
- M2 growth < 8%: -10

**Japan Factors (10% weight)**:
- USD/JPY > 150: -10 (carry trade risk)

**EM Factors (10% weight)**:
- EM strength index > 0: +10

### Score Interpretation (lines 467-482)

| Score Range | Signal | Recommendation |
|------------|--------|----------------|
| ≥ 60 | 🚀 EXTREME LIQUIDITY | Growth stocks, EM, commodities |
| 30-59 | ✅ HIGH LIQUIDITY | Maintain/increase risk assets |
| 0-29 | ⚖️ NEUTRAL | Balanced portfolio |
| -29 to -1 | ⚠️ TIGHT | Increase cash/bonds |
| ≤ -30 | 🔴 EXTREME TIGHT | Defensive, prefer USD/gold |

---

## Google Sheets Structure

### Live_Monitor Sheet
**Purpose**: Latest US liquidity indicators
**Columns**: Date, SOFR, EFFR, IORB, SOFR-IORB spread, ON RRP, TGA, WALCL, WoW, SRF, Signal

### Global_Liquidity Sheet
**Purpose**: Latest global liquidity analysis
**Columns** (19 total): Timestamp, WALCL, WALCL WoW, TGA, TGA WoW, ON RRP, DXY, DXY WoW, China M2%, China Credit, China FX, USD/JPY, JGB 10Y, US-JP Spread, USD/KRW, USD/BRL, EM Index, Liquidity Score, Signal, Recommendation

### History Sheet
**Purpose**: Time-series of US indicators
**Updates**: Auto-appends on each `updateLiveMonitor()` call
**Retention**: Unlimited (manual cleanup required)

### Global_History Sheet
**Purpose**: Time-series of global liquidity analysis
**Updates**: Auto-appends on each global analysis
**Columns**: Same as Global_Liquidity + timestamp

### Alert_History Sheet
**Purpose**: Log of all triggered alerts
**Columns**: Timestamp, Liquidity Score, Signal, Alert Level, Message, Recommended Action
**Color Coding**: Green (opportunity), Red (warning), Yellow (risk)

---

## Development Workflows

### Common Tasks for AI Assistants

#### 1. Adding a New Data Source

**Location**: CONFIG object, data collection functions

**Steps**:
1. Add FRED ID to `CONFIG.FRED_IDS` or `CONFIG.GLOBAL_FRED_IDS`
2. Create getter function (follow pattern in lines 62-142)
3. Integrate into analysis function
4. Update sheet headers if new column needed
5. Test with `testAllSystems()`

**Example**:
```javascript
// 1. Add to CONFIG
GLOBAL_FRED_IDS: {
  NEW_INDICATOR: 'FRED_SERIES_ID',
  // ...
}

// 2. Create getter
function getNewIndicator() {
  const data = getFredData(CONFIG.GLOBAL_FRED_IDS.NEW_INDICATOR);
  return { value: data.value, signal: determineSignal(data.value) };
}

// 3. Integrate into analyzeGlobalLiquidity()
const newIndicator = getNewIndicator();
// Add to scoring logic
// Add to sheet output
```

#### 2. Modifying Scoring Logic

**Location**: `analyzeGlobalLiquidity()` function (lines 442-462)

**Guidelines**:
- Maintain total weight = 100%
- Document weight changes in comments
- Test edge cases (extreme values)
- Update help documentation

**Example**:
```javascript
// Add new factor (adjust other weights to compensate)
// Crypto volatility factor (5% weight)
if (btc_volatility > 80) liquidityScore -= 5;
else if (btc_volatility < 40) liquidityScore += 5;
```

#### 3. Customizing Alerts

**Location**: `checkGlobalAlerts()` function (lines 889-944)

**Alert Structure**:
```javascript
alerts.push({
  level: '🚨 LEVEL_NAME',      // Emoji + severity
  message: 'Clear description', // What happened
  action: 'Recommended action'  // What to do
});
```

**Trigger Thresholds**:
- Extreme liquidity: score ≥ 60
- Extreme tight: score ≤ -30
- China risk: M2 < 7%
- Yen carry risk: USD/JPY > 155
- Dollar volatility: |DXY change| > 2

#### 4. Updating Sheet Layouts

**Key Functions**:
- `setupGlobalSheet(sheet)` - Initialize Global_Liquidity
- Line 487 - Update Live data range
- Line 600 - Update Global_History append

**Important**: Column count must match data array length

---

## Key Conventions

### Code Style
- **Function naming**: camelCase, descriptive verbs
- **Constants**: UPPER_SNAKE_CASE in CONFIG
- **Comments**: Korean for business logic, English acceptable for code
- **Error handling**: Try-catch with Logger.log for all API calls

### Google Apps Script Specifics
- **Services used**:
  - `UrlFetchApp` - API calls
  - `SpreadsheetApp` - Sheet operations
  - `CacheService` - Data caching
  - `ScriptApp` - Trigger management
  - `GmailApp` - Email alerts
  - `Session` - User info
  - `HtmlService` - UI dialogs

- **Quotas**:
  - URL Fetch: 20,000 calls/day
  - Email: 100/day (consumer), 1,500/day (Google Workspace)
  - Script runtime: 6 min/execution

### Error Handling Pattern
```javascript
try {
  // API call or operation
  const result = someOperation();
  Logger.log('✅ Success message');
  return result;
} catch (e) {
  Logger.log(`❌ Error context: ${e.message}`);
  return { value: 0, error: 'ERROR' }; // Graceful fallback
}
```

### Emoji Usage in UI
- 🚀 - Extreme positive/opportunity
- ✅ - Positive/normal
- ⚖️ - Neutral
- ⚠️ - Warning/caution
- 🔴 - Critical/extreme negative
- 💵 - Dollar/USD related
- 🇨🇳 - China
- 🇯🇵 - Japan

---

## Deployment & Setup

### Prerequisites
1. Google account with Google Sheets access
2. Apps Script project linked to a Google Sheet
3. Time zone set to 'America/New_York' in script properties

### Initial Setup Steps

1. **Create Google Sheet**
   - Open Google Sheets, create new spreadsheet
   - Name it appropriately (e.g., "Global Liquidity Monitor")

2. **Add Apps Script**
   - Extensions → Apps Script
   - Delete default code
   - Paste contents of `app_script.js`
   - Save (Ctrl+S)

3. **First Run**
   - Run `onOpen()` function manually
   - Authorize required permissions:
     - Access spreadsheet
     - Connect to external services
     - Send email
     - Manage time-based triggers
   - Menu "📊 Global Liquidity" should appear in spreadsheet

4. **Initialize Sheets**
   - Run "🔄 전체 업데이트" from menu
   - Creates all required sheets automatically

5. **Optional: Set Triggers**
   - "⏰ 일일 자동갱신" - Daily at 5 PM ET
   - "🔔 알림 설정/해제" - 2-hour alert checks

### Configuration Customization

**Time Zone** (for timestamps):
```javascript
// Line 485, 715
.toLocaleString('ko-KR', {timeZone: 'America/New_York'})
```

**Cache Duration**:
```javascript
CONFIG.CACHE_TIME = 300000; // milliseconds (default: 5 min)
```

**Alert Email**:
Auto-sends to active user. To customize:
```javascript
// Line 1011 in sendGlobalAlert()
const userEmail = 'custom@email.com'; // Override Session.getActiveUser()
```

---

## Testing & Debugging

### Test Function
**Location**: Lines 1346-1375

**Run**: `testAllSystems()`

**Output**: Logs to Apps Script console
- FRED data fetch validation
- Global data retrieval
- Analysis calculation
- History logging

**Check**:
1. View → Executions (see runtime logs)
2. Check for "❌" error markers
3. Verify sheet updates

### Common Issues

**1. FRED API Rate Limits**
- **Symptom**: 429 HTTP errors
- **Solution**: Increase `CONFIG.CACHE_TIME`, reduce update frequency

**2. Missing Data**
- **Symptom**: `null` values in sheets
- **Cause**: FRED series discontinued or renamed
- **Solution**: Check FRED website, update ID in CONFIG

**3. Trigger Not Running**
- **Check**: Apps Script → Triggers panel
- **Fix**: Delete and recreate trigger

**4. Email Not Sending**
- **Quota**: Check Apps Script quotas
- **Permissions**: Re-authorize script
- **Spam**: Check Gmail spam folder

### Debugging Tools

**Logger Output**:
```javascript
Logger.log('Debug message');
// View: Ctrl+Enter or View → Logs
```

**Console Logging** (in custom HTML dialogs):
```javascript
console.log('Browser console message');
// View: F12 developer tools
```

**Execution Transcript**:
- Apps Script → Executions
- Shows runtime, errors, triggers

---

## API Reference - Key Functions

### Data Fetching

```javascript
getFredData(fredId, useCache = true)
// Returns: { date, value, timestamp, fredId }
// Throws: Error with message on failure

getFredDataHistorical(fredId, daysAgo)
// Returns: { date, value }
// Note: Approximate, uses row indexing

getSRFData()
// Returns: { date, amount, rate, source }
// Fallback: Returns zeros if API fails
```

### Analysis

```javascript
analyzeGlobalLiquidity()
// Returns: {
//   score: number,
//   signal: string,
//   recommendation: string,
//   timestamp: Date,
//   details: { us, dxy, china, japan, em }
// }

getChinaLiquidity()
// Returns: { m2_growth, total_credit, fx_reserves, liquidity_signal }

getJapanLiquidity()
// Returns: { usdjpy, jgb_10y, us_jpy_spread, carry_risk }

getTGAAnalysis()
// Returns: { current, week_change, month_change, liquidity_impact, debt_ceiling_risk }

getEmergingMarketsFX()
// Returns: { usdkrw, usdbrl, usdmxn, strength_index, signal }
```

### Main Operations

```javascript
updateLiveMonitor()
// Primary update function
// Side effects: Updates Live_Monitor, appends History, triggers global analysis

onOpen()
// Menu builder, runs automatically when sheet opens

setupGlobalAlerts()
// UI dialog for alert management

checkGlobalAlerts()
// Evaluates alert conditions, sends emails, logs history
```

---

## AI Assistant Guidelines

### When Modifying This Code

1. **Preserve Korean UI elements** - All user-facing text is in Korean
2. **Maintain sheet structure** - Column order changes break existing sheets
3. **Test with `testAllSystems()`** - Always validate after changes
4. **Update help documentation** - Modify `showHelp()` if adding features
5. **Respect API quotas** - Don't reduce cache times below 1 minute
6. **Graceful degradation** - Functions should return safe defaults on error

### Code Modification Checklist

- [ ] Update CONFIG if adding data sources
- [ ] Maintain total scoring weight = 100%
- [ ] Update sheet headers if adding columns
- [ ] Test edge cases (null data, API failures)
- [ ] Update `showHelp()` function
- [ ] Verify email template formatting
- [ ] Check Apps Script quota implications
- [ ] Add logging for debugging
- [ ] Update this CLAUDE.md file

### Best Practices

**DO**:
- Use descriptive variable names
- Add comments for complex logic
- Log success and failure states
- Handle API failures gracefully
- Test with `testAllSystems()`
- Cache expensive operations

**DON'T**:
- Hardcode email addresses
- Remove try-catch blocks
- Change column order without updating all references
- Make synchronous calls in loops (use batch operations)
- Ignore quota limits
- Remove error logging

### Performance Optimization

**Current bottlenecks**:
1. FRED API calls (15s timeout each)
2. Sheet operations (batch when possible)

**Optimization strategies**:
- Batch API calls where possible
- Use `getRange().setValues()` for multiple cells
- Increase cache times for stable data
- Consider background triggers for heavy updates

---

## Glossary

### Financial Terms
- **WALCL**: Weekly Assets Less Custody Liabilities (Fed balance sheet)
- **TGA**: Treasury General Account (US Treasury's Fed account)
- **ON RRP**: Overnight Reverse Repo Program
- **SOFR**: Secured Overnight Financing Rate
- **EFFR**: Effective Federal Funds Rate
- **IORB**: Interest on Reserve Balances
- **DXY**: US Dollar Index
- **M2**: Money Supply measure (cash + deposits)
- **JGB**: Japanese Government Bonds
- **EM**: Emerging Markets
- **Carry Trade**: Borrow in low-rate currency, invest in high-rate

### Technical Terms
- **WoW**: Week-over-Week change
- **YoY**: Year-over-Year change
- **bp**: Basis points (1/100th of 1%)
- **Spread**: Difference between two rates

### Script-Specific
- **Live_Monitor**: Real-time US data dashboard
- **Global_Liquidity**: Real-time global analysis dashboard
- **History**: Time-series storage
- **Liquidity Score**: Composite 0-100 metric

---

## Version History

**v3.0** (Current)
- Complete integration of US and global liquidity
- Automated history logging (3 sheets)
- Alert system with email notifications
- Dashboard and report generation
- Individual check functions for regions

**Features**: 5 main monitoring areas (US, Global, China, Japan, EM)

---

## Contact & Support

**Repository**: Git-based version control
**Platform**: Google Apps Script
**Documentation**: This file (CLAUDE.md)

**For AI Assistants**:
- Read this entire document before making changes
- When uncertain, preserve existing behavior
- Test thoroughly with `testAllSystems()`
- Document all modifications in git commits
- Update this CLAUDE.md if architecture changes

---

**Last Updated**: 2025-11-13
**Document Version**: 1.0
**Code Version**: 3.0
