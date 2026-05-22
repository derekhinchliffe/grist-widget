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

const aggregators = {
  'Weighted Average': weightedAverage
};

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
  //const container = document.querySelector('.copy-container');
  if (!container) return;

  // Get the HTML content of the container
  //const html = copyWithStyles(container).outerHTML;

  

  // Use the Clipboard API to write as HTML
  // if (navigator.clipboard && window.ClipboardItem) {
  //   const blob = new Blob([html], { type: 'text/html' });
  //   const item = new ClipboardItem({ 'text/html': blob });
  //   navigator.clipboard.write([item])
  //     .then(() => {
  //       console.log('Copied as HTML!');
  //     })
  //     .catch(err => {
  //       console.error('Clipboard write failed:', err);
  //     });
  // } else {
    // Fallback for older browsers
    const range = document.createRange();
    range.selectNode(container);
    window.getSelection().removeAllRanges();
    window.getSelection().addRange(range);
    document.execCommand('copy');
    window.getSelection().removeAllRanges();
    console.log('Copied using execCommand fallback.');
  // }
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
        $.pivotUtilities.export_renderers
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
