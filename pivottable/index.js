window.onerror = (err) => {
  console.trace();
  alert(String(err));
};

grist.ready({
  requiredAccess: 'read table'
});

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
  let initialRender = true;
  $('#table').pivotUI(
    rec,
    {
      rows,
      cols,
      vals,
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
      inclusions,
      exclusions,
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
