import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

// Extract functions from the HTML file's script block
const html = readFileSync('lof_premium.html', 'utf-8');
const scriptMatch = html.match(/<script>([\s\S]*?)<\/script>/);
const scriptContent = scriptMatch[1];

// Remove DOM-dependent code and create a module-like environment
// DOM elements store for testing state management functions
const domElements = {
  'loading-indicator': { textContent: '', style: { display: 'none' } },
  'refresh-btn': { textContent: '', style: {}, disabled: false },
  'error-message': { textContent: '', style: { display: 'none' } },
  'empty-state': { textContent: '', style: { display: 'none' } },
  'last-updated': { textContent: '', style: {} },
  'sort-indicator': { textContent: '▼', style: {} },
  'table-body': { textContent: '', style: {}, innerHTML: '', appendChild: () => {} },
};

const wrappedCode = `
  // Stub document for code that references DOM
  const domElements = arguments[0];
  const document = {
    getElementById: (id) => domElements[id] || { textContent: '', style: {} },
    addEventListener: (event, handler) => { /* no-op in test */ }
  };
  const fetch = () => Promise.reject(new Error('no fetch in test'));
  const setTimeout = globalThis.setTimeout;
  const clearTimeout = globalThis.clearTimeout;
  const Date = globalThis.Date;
  const AbortController = globalThis.AbortController;
  ${scriptContent}
  return { parseData, filterPositivePremium, classifyMarket, formatPremiumRate, formatPrice, sortByPremium, toggleSort, state, showLoading, hideLoading, showError, updateTimestamp, refreshData };
`;
const mod = new Function(wrappedCode)(domElements);
const { parseData, filterPositivePremium, sortByPremium, toggleSort, state, showLoading, hideLoading, showError, updateTimestamp } = mod;

describe('parseData', () => {
  test('parses valid API response into FundRecord array', () => {
    const rawJson = {
      rows: [
        { id: '501029', cell: { fund_id: '501029', fund_nm: '华宝标普油气LOF', index_nm: '标普石油天然气上游股票指数', price: '0.571', discount_rt: '3.25%', apply_status: '限500', apply_fee: '1.20%' } }
      ]
    };
    const result = parseData(rawJson);
    assert.equal(result.length, 1);
    assert.equal(result[0].fundCode, '501029');
    assert.equal(result[0].fundName, '华宝标普油气LOF');
    assert.equal(result[0].indexName, '标普石油天然气上游股票指数');
    assert.equal(result[0].price, 0.571);
    assert.equal(result[0].premiumRate, 3.25);
    assert.equal(result[0].market, '欧美');
    assert.equal(result[0].applyStatus, '限500');
    assert.equal(result[0].applyFee, '1.20%');
  });

  test('returns empty array when rawJson is null/undefined', () => {
    assert.deepEqual(parseData(null), []);
    assert.deepEqual(parseData(undefined), []);
  });

  test('returns empty array when rows is missing', () => {
    assert.deepEqual(parseData({}), []);
    assert.deepEqual(parseData({ rows: 'not-array' }), []);
  });

  test('skips records with missing fund_id', () => {
    const rawJson = {
      rows: [
        { id: '1', cell: { fund_nm: 'Test', price: '1.0', discount_rt: '1.0%' } }
      ]
    };
    assert.deepEqual(parseData(rawJson), []);
  });

  test('skips records with missing fund_nm', () => {
    const rawJson = {
      rows: [
        { id: '1', cell: { fund_id: '123456', price: '1.0', discount_rt: '1.0%' } }
      ]
    };
    assert.deepEqual(parseData(rawJson), []);
  });

  test('skips records with non-numeric discount_rt', () => {
    const rawJson = {
      rows: [
        { id: '1', cell: { fund_id: '123456', fund_nm: 'Test', price: '1.0', discount_rt: 'abc%' } }
      ]
    };
    assert.deepEqual(parseData(rawJson), []);
  });

  test('skips records with non-numeric price', () => {
    const rawJson = {
      rows: [
        { id: '1', cell: { fund_id: '123456', fund_nm: 'Test', price: 'abc', discount_rt: '1.0%' } }
      ]
    };
    assert.deepEqual(parseData(rawJson), []);
  });

  test('skips records with null cell', () => {
    const rawJson = {
      rows: [{ id: '1', cell: null }, { id: '2' }]
    };
    assert.deepEqual(parseData(rawJson), []);
  });

  test('handles negative premium rates (parses them, does not filter)', () => {
    const rawJson = {
      rows: [
        { id: '1', cell: { fund_id: '123456', fund_nm: 'Test', index_nm: '', price: '1.0', discount_rt: '-2.50%' } }
      ]
    };
    const result = parseData(rawJson);
    assert.equal(result.length, 1);
    assert.equal(result[0].premiumRate, -2.5);
  });

  test('handles discount_rt without % sign', () => {
    const rawJson = {
      rows: [
        { id: '1', cell: { fund_id: '123456', fund_nm: 'Test', index_nm: '', price: '1.0', discount_rt: '3.25' } }
      ]
    };
    const result = parseData(rawJson);
    assert.equal(result.length, 1);
    assert.equal(result[0].premiumRate, 3.25);
  });

  test('defaults optional fields to empty string', () => {
    const rawJson = {
      rows: [
        { id: '1', cell: { fund_id: '123456', fund_nm: 'Test', price: '1.0', discount_rt: '1.0%' } }
      ]
    };
    const result = parseData(rawJson);
    assert.equal(result.length, 1);
    assert.equal(result[0].indexName, '');
    assert.equal(result[0].applyStatus, '');
    assert.equal(result[0].applyFee, '');
  });

  test('calls classifyMarket for market field', () => {
    const rawJson = {
      rows: [
        { id: '1', cell: { fund_id: '123456', fund_nm: 'Test', index_nm: '日经225指数', price: '1.0', discount_rt: '1.0%' } }
      ]
    };
    const result = parseData(rawJson);
    assert.equal(result[0].market, '亚洲');
  });
});

describe('filterPositivePremium', () => {
  test('filters records with premiumRate > 0', () => {
    const records = [
      { premiumRate: 3.25 },
      { premiumRate: -1.5 },
      { premiumRate: 0 },
      { premiumRate: 0.01 }
    ];
    const result = filterPositivePremium(records);
    assert.equal(result.length, 2);
    assert.equal(result[0].premiumRate, 3.25);
    assert.equal(result[1].premiumRate, 0.01);
  });

  test('returns empty array when no positive premiums', () => {
    const records = [
      { premiumRate: -1.5 },
      { premiumRate: 0 },
      { premiumRate: -0.01 }
    ];
    assert.deepEqual(filterPositivePremium(records), []);
  });

  test('returns empty array for empty input', () => {
    assert.deepEqual(filterPositivePremium([]), []);
  });

  test('returns empty array for non-array input', () => {
    assert.deepEqual(filterPositivePremium(null), []);
    assert.deepEqual(filterPositivePremium(undefined), []);
  });

  test('keeps all records when all have positive premiums', () => {
    const records = [
      { premiumRate: 1.0 },
      { premiumRate: 5.5 },
      { premiumRate: 0.001 }
    ];
    const result = filterPositivePremium(records);
    assert.equal(result.length, 3);
  });
});

describe('sortByPremium', () => {
  test('sorts descending (high to low) by default', () => {
    const records = [
      { premiumRate: 1.0 },
      { premiumRate: 3.25 },
      { premiumRate: 2.18 }
    ];
    const result = sortByPremium(records, 'desc');
    assert.equal(result[0].premiumRate, 3.25);
    assert.equal(result[1].premiumRate, 2.18);
    assert.equal(result[2].premiumRate, 1.0);
  });

  test('sorts ascending (low to high)', () => {
    const records = [
      { premiumRate: 3.25 },
      { premiumRate: 1.0 },
      { premiumRate: 2.18 }
    ];
    const result = sortByPremium(records, 'asc');
    assert.equal(result[0].premiumRate, 1.0);
    assert.equal(result[1].premiumRate, 2.18);
    assert.equal(result[2].premiumRate, 3.25);
  });

  test('does not mutate the original array', () => {
    const records = [
      { premiumRate: 3.0 },
      { premiumRate: 1.0 },
      { premiumRate: 2.0 }
    ];
    const original = [...records];
    sortByPremium(records, 'asc');
    assert.deepEqual(records, original);
  });

  test('returns empty array for empty input', () => {
    assert.deepEqual(sortByPremium([], 'desc'), []);
  });

  test('handles single element array', () => {
    const records = [{ premiumRate: 5.0 }];
    const result = sortByPremium(records, 'desc');
    assert.equal(result.length, 1);
    assert.equal(result[0].premiumRate, 5.0);
  });

  test('handles equal premium rates', () => {
    const records = [
      { premiumRate: 2.0, fundCode: 'A' },
      { premiumRate: 2.0, fundCode: 'B' }
    ];
    const result = sortByPremium(records, 'desc');
    assert.equal(result.length, 2);
    assert.equal(result[0].premiumRate, 2.0);
    assert.equal(result[1].premiumRate, 2.0);
  });
});

describe('toggleSort', () => {
  test('toggles from desc to asc', () => {
    state.sortDirection = 'desc';
    toggleSort();
    assert.equal(state.sortDirection, 'asc');
  });

  test('toggles from asc to desc', () => {
    state.sortDirection = 'asc';
    toggleSort();
    assert.equal(state.sortDirection, 'desc');
  });

  test('double toggle returns to original direction', () => {
    state.sortDirection = 'desc';
    toggleSort();
    toggleSort();
    assert.equal(state.sortDirection, 'desc');
  });
});


// Helper to reset DOM element state between tests
function resetDomElements() {
  domElements['loading-indicator'].style.display = 'none';
  domElements['refresh-btn'].disabled = false;
  domElements['error-message'].style.display = 'none';
  domElements['error-message'].textContent = '';
  domElements['empty-state'].style.display = 'none';
  domElements['last-updated'].textContent = '';
  state.isLoading = false;
  state.error = null;
  state.lastUpdated = null;
}

describe('showLoading', () => {
  test('sets state.isLoading to true', () => {
    resetDomElements();
    showLoading();
    assert.equal(state.isLoading, true);
  });

  test('shows loading indicator with display flex', () => {
    resetDomElements();
    showLoading();
    assert.equal(domElements['loading-indicator'].style.display, 'flex');
  });

  test('disables refresh button', () => {
    resetDomElements();
    showLoading();
    assert.equal(domElements['refresh-btn'].disabled, true);
  });

  test('hides error message', () => {
    resetDomElements();
    domElements['error-message'].style.display = 'block';
    showLoading();
    assert.equal(domElements['error-message'].style.display, 'none');
  });

  test('hides empty state', () => {
    resetDomElements();
    domElements['empty-state'].style.display = 'block';
    showLoading();
    assert.equal(domElements['empty-state'].style.display, 'none');
  });
});

describe('hideLoading', () => {
  test('sets state.isLoading to false', () => {
    resetDomElements();
    state.isLoading = true;
    hideLoading();
    assert.equal(state.isLoading, false);
  });

  test('hides loading indicator', () => {
    resetDomElements();
    domElements['loading-indicator'].style.display = 'flex';
    hideLoading();
    assert.equal(domElements['loading-indicator'].style.display, 'none');
  });

  test('enables refresh button', () => {
    resetDomElements();
    domElements['refresh-btn'].disabled = true;
    hideLoading();
    assert.equal(domElements['refresh-btn'].disabled, false);
  });
});

describe('showError', () => {
  test('sets state.error to the message', () => {
    resetDomElements();
    showError('网络错误');
    assert.equal(state.error, '网络错误');
  });

  test('sets error element textContent', () => {
    resetDomElements();
    showError('数据格式异常');
    assert.equal(domElements['error-message'].textContent, '数据格式异常');
  });

  test('shows error element with display block', () => {
    resetDomElements();
    showError('请求超时');
    assert.equal(domElements['error-message'].style.display, 'block');
  });
});

describe('updateTimestamp', () => {
  test('sets state.lastUpdated to a Date object', () => {
    resetDomElements();
    updateTimestamp();
    assert.ok(state.lastUpdated instanceof Date);
  });

  test('formats timestamp as "最后更新: YYYY-MM-DD HH:mm:ss"', () => {
    resetDomElements();
    updateTimestamp();
    const text = domElements['last-updated'].textContent;
    assert.match(text, /^最后更新: \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/);
  });

  test('timestamp is not in the future', () => {
    resetDomElements();
    const before = new Date();
    updateTimestamp();
    assert.ok(state.lastUpdated >= before);
    assert.ok(state.lastUpdated <= new Date());
  });
});
