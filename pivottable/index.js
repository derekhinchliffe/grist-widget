window.onerror = (err) => {
  console.trace();
  alert(String(err));
};

grist.ready({
  requiredAccess: 'full'
});

/**
 * Fetch a map of { colId -> label } for the currently selected table.
 * Falls back to colId if the REST API is unavailable.
 */
async function fetchColumnLabels() {
  try {
    const tokenInfo = await grist.getAccessToken({ readOnly: true });
    // Determine the current table id from the first record fetch or via viewApi.
    // We use fetchSelectedTable with keepEncoded to grab the raw column names.
    const rawTable = await grist.fetchSelectedTable({ keepEncoded: true });
    const colIds = Object.keys(rawTable).filter(k => k !== 'id');

    // Find which table we are attached to by inspecting the tableId.
    // We list tables and look for one whose columns include our colIds.
    const tablesRes = await fetch(
      `${tokenInfo.baseUrl}/tables?auth=${tokenInfo.token}`
    );
    const { tables } = await tablesRes.json();

    for (const table of tables) {
      const colsRes = await fetch(
        `${tokenInfo.baseUrl}/tables/${table.id}/columns?auth=${tokenInfo.token}`
      );
      const { columns } = await colsRes.json();
      const ids = columns.map(c => c.id);
      // Check if this table contains all our colIds (a reasonable heuristic).
      if (colIds.every(id => ids.includes(id))) {
        const map = {};
        for (const col of columns) {
          map[col.id] = col.fields.label || col.id;
        }
        return map;
      }
    }
  } catch (e) {
    console.warn('fetchColumnLabels failed, using raw column ids:', e);
  }
  return {};
}

/**
 * Given records (array of objects with colId keys) and a colId->label map,
 * return a new array of records with keys replaced by their labels.
 * Also returns the reverse map (label -> colId) for saving settings.
 */
function remapRecordKeys(records, labelMap) {
  if (!records.length || !Object.keys(labelMap).length) return records;
  return records.map(rec => {
    const out = {};
    for (const [k, v] of Object.entries(rec)) {
      out[labelMap[k] ?? k] = v;
    }
    return out;
  });
}

/**
 * Remap an array of column identifiers (as used in pivot settings)
 * through a mapping object. Missing keys pass through unchanged.
 */
function remapKeys(arr, map) {
  if (!arr) return arr;
  return arr.map(k => map[k] ?? k);
}

function wavg (n) {
  if (!n) { return; }
  n = n.filter(([note]) => typeof (note) === 'number');
  if (n.length) { return n.map(([note, coef]) => note * coef).reduce((a, b) => a + b) / n.map(([_note, coef]) => coef).reduce((a, b) => a + b); }
}

function weightedAverage ([val, coef]) {
  return (_data, _rowKey, _colKey) => ({
    values: [],
    push: function (rec) { this.values.push([rec[val], rec[coef]]); },
    value: function () { return wavg(this.values); },
    format: function (x) { return (Math.round(x * 100) / 100).toFixed(2); },
    numInputs: 2
  });
}

function sumOverSumPercentage([sumVal, sumOfVal]) {
  return (_data, _rowKey, _colKey) => ({
    sum: 0,
    sumOf: 0,
    push: function (rec) { this.sum += rec[sumVal]; this.sumOf += rec[sumOfVal]; },
    value: function () { return this.sumOf ? this.sum / this.sumOf : 0; },
    format: function (x) { return (Math.round(x * 10000) / 100).toFixed(0) + '%'; },
    numInputs: 2
  });
}

const aggregators = {
  'Weighted Average': weightedAverage,
  'Sum Over Sum Percentage': sumOverSumPercentage
};

const renderers = {
  'Table (min 75%)': function(data, opts) {
    var table = $.pivotUtilities.renderers['Table'](data, opts);
    // Search through all .pvtVal cells and set the text color to red if the value is below 75%
    $(table).find('.pvtVal').each(function() {
      var value = parseFloat($(this).text());
      if (value < 75) {
        $(this).css('color', 'red');
      }
    });
    return table;
  }
}

function transferComputedStyle(node) {
    var cs = getComputedStyle(node, null);
    var i;
    for (i = 0; i < cs.length; i++) {
        var s = cs[i] + "";
          node.style[s] = cs[s];
    }
}

function transferAllStyles(node) {
  transferComputedStyle(node);
    var all = node.querySelectorAll("*");
    var i;
    for (i = 0; i < all.length; i++) {
        transferComputedStyle(all[i]);
    }
}

function copyWithStyles(node) {
    var copy = node.cloneNode(true);
    //transferAllStyles(node);
    copy.style['font-family'] = 'Arial, sans-serif';
    copy.style['font-size'] = '8pt';
    
    return copy;
}

function copyContainerToClipboard(container) {
  if (!container) return;

  // Using the older select range method of copying to clipboard rather than the Clipboard API
  // as it seems to work better with pasting styles into Word.

  const range = document.createRange();
  range.selectNode(container);
  window.getSelection().removeAllRanges();
  window.getSelection().addRange(range);
  document.execCommand('copy');
  window.getSelection().removeAllRanges();
}


function syncCopyButtonState() {
  const copyButton = document.getElementById('copyPivotTable');
  if (!copyButton) return;

  const pivotTable = document.querySelector('#table table.pvtTable');
  copyButton.disabled = !pivotTable;
  copyButton.onclick = () => copyContainerToClipboard(pivotTable);
}

grist.onRecords(async rec => {
  const {
    rows, cols, vals, aggregatorName, rendererName, inclusions, exclusions
  } = await grist.getOption('settings') ?? {};

  // Build colId->label and label->colId maps, then remap record keys.
  const labelMap = await fetchColumnLabels();
  const idToLabel = labelMap;  // colId  -> label

  const remappedRec = remapRecordKeys(rec, idToLabel);

  // Saved settings store labels (what the user sees in the UI). If settings
  // were saved before labels were introduced they may contain colIds –
  // attempt to translate them to labels so the pivot restores correctly.
  const remappedRows = remapKeys(rows, idToLabel);
  const remappedCols = remapKeys(cols, idToLabel);
  const remappedVals = remapKeys(vals, idToLabel);

  // inclusions/exclusions are keyed by column name too.
  function remapIncExc(obj) {
    if (!obj) return obj;
    const out = {};
    for (const [k, v] of Object.entries(obj)) {
      out[idToLabel[k] ?? k] = v;
    }
    return out;
  }

  let initialRender = true;
  $('#table').pivotUI(
    remappedRec,
    {
      rows: remappedRows,
      cols: remappedCols,
      vals: remappedVals,
      onRefresh: function (config) {
        syncCopyButtonState();
        if (initialRender) {
          initialRender = false;
          return;
        }
        const {
          rows, cols, vals, aggregatorName, rendererName, inclusions, exclusions
        } = config;
        grist.setOption('settings', {
          rows, cols, vals, aggregatorName, rendererName, inclusions, exclusions
        });
      },
      aggregatorName,
      rendererName,
      inclusions: remapIncExc(inclusions),
      exclusions: remapIncExc(exclusions),
      aggregators: $.extend($.pivotUtilities.aggregators, aggregators),
      renderers: $.extend(
        $.pivotUtilities.renderers,
        $.pivotUtilities.plotly_renderers,
        $.pivotUtilities.d3_renderers,
        $.pivotUtilities.export_renderers,
        renderers
      ),
      rendererOptions: {
        heatmap: {
          // Add a colorScaleGenerator to show light red below 75%, through to dark red at 0%. Anything above 75% is not colored.
          colorScaleGenerator: function (values) {
            return function (value) {
              if (value >= 0.75) {
                return 'transparent';
              }
              // The following needs to be in #RRGGBB format to handle copy/paste to Word properly

              const intensity = 128 + Math.round(127 * (value / 0.75));
              return `#FF${intensity.toString(16).padStart(2, '0')}${intensity.toString(16).padStart(2, '0')}`;
            };
          }
        }
      }
    },
    true
  );
});
