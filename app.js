async function loadData() {
  // Use inline data.js if available (works when opened via file://)
  if (window.DASHBOARD_DATA) return window.DASHBOARD_DATA;
  // Fall back to fetch (works when served over http)
  const res = await fetch('./data/data.json', { cache: 'no-store' });
  if (!res.ok) throw new Error(`Failed to load data: ${res.status}`);
  return res.json();
}

function drawChart(rows) {
  const labels = rows.map(r => (r.Date || '').slice(0, 10));
  const spValues = rows.map(r => Number(r['S&P']));
  const signalValues = rows.map(r => Number(r['Signal']));
  const numericSp = spValues.filter(v => !Number.isNaN(v));

  const ctx = document.getElementById('chart');
  new Chart(ctx, {
    type: 'line',
    data: {
      labels,
      datasets: [
        {
          label: 'S&P',
          data: spValues,
          yAxisID: 'y',
          borderColor: '#7aa2ff',
          backgroundColor: 'rgba(122,162,255,0.2)',
          tension: 0.2,
          pointRadius: 0,
          borderWidth: 2,
          fill: true
        },
        {
          label: 'Model Score (Signal)',
          data: signalValues,
          yAxisID: 'y1',
          borderColor: '#f6c343',
          backgroundColor: 'rgba(246,195,67,0.15)',
          tension: 0.2,
          pointRadius: 0,
          borderWidth: 2,
          fill: false
        }
      ]
    },
    options: {
      responsive: true,
      interaction: {
        mode: 'index',
        intersect: false
      },
      plugins: {
        legend: {
          display: true,
          labels: { color: '#ffffff', font: { size: 13 } }
        }
      },
      scales: {
        x: {
          ticks: {
            maxTicksLimit: 12,
            color: '#ffffff',
            font: { size: 13 }
          }
        },
        y: {
          type: 'linear',
          position: 'left',
          beginAtZero: false,
          title: { display: true, text: 'S&P', color: '#ffffff' },
          ticks: { color: '#ffffff', font: { size: 13 } }
        },
        y1: {
          type: 'linear',
          position: 'right',
          beginAtZero: false,
          min: 0,
          max: 1,
          grid: { drawOnChartArea: false },
          title: { display: true, text: 'Model Score', color: '#ffffff' },
          ticks: {
            color: '#ffffff',
            font: { size: 13 },
            callback: value => Math.round(value * 100) + '%'
          }
        }
      }
    }
  });

  return numericSp;
}

(async () => {
  const meta = document.getElementById('meta');
  try {
    const payload = await loadData();
    const rows = payload.rows || [];
    if (!rows.length) throw new Error('No rows found in data.');

    // Append live quote as today's point if newer than last row
    if (payload.spQuote != null || payload.modelScore != null) {
      const today = new Date().toISOString().slice(0, 10);
      const lastDate = (rows[rows.length - 1]?.Date || '').slice(0, 10);
      if (today > lastDate) {
        rows.push({
          Date: today + 'T00:00:00',
          'S&P': payload.spQuote ?? null,
          Signal: payload.modelScore ?? null
        });
      } else {
        // Update the last row with live values
        rows[rows.length - 1]['S&P'] = payload.spQuote ?? rows[rows.length - 1]['S&P'];
        rows[rows.length - 1]['Signal'] = payload.modelScore ?? rows[rows.length - 1]['Signal'];
      }
    }

    const values = drawChart(rows);
    const latest = values[values.length - 1];
    meta.textContent = `Sheet: ${payload.sheet} • Points: ${rows.length} • Latest S&P: ${latest}`;

    // Show portfolio weights
    const holdingsEl = document.getElementById('holdingsList');
    if (holdingsEl && payload.portfolioWeights && payload.portfolioWeights.length) {
      holdingsEl.innerHTML = payload.portfolioWeights.map(h =>
        `<div class="holding-item">
          <span class="holding-symbol">${h.symbol}</span>
          <span class="holding-pct">${(h.weight * 100).toFixed(1)}%</span>
        </div>`
      ).join('');
    }

    // Show index returns
    const indexEl = document.getElementById('indexReturnsList');
    if (indexEl && payload.indexReturns && payload.indexReturns.length) {
      indexEl.innerHTML = payload.indexReturns.map(h => {
        const pct = (h.return * 100).toFixed(2);
        const color = pct >= 0 ? '#4caf84' : '#f06292';
        return `<div class="holding-item">
          <span class="holding-symbol">${h.symbol}</span>
          <span class="holding-pct" style="color:${color}">${pct >= 0 ? '+' : ''}${pct}%</span>
        </div>`;
      }).join('');
    }

    // Show YTD return
    const ytdEl = document.getElementById('ytdReturn');
    if (ytdEl && payload.ytdReturn != null) {
      const pct = (payload.ytdReturn * 100).toFixed(2);
      ytdEl.textContent = (pct > 0 ? '+' : '') + pct + '%';
      ytdEl.style.color = pct >= 0 ? '#4caf84' : '#f06292';
    }

    // Show S&P quote
    const spEl = document.getElementById('spQuote');
    if (spEl && payload.spQuote != null) {
      spEl.textContent = Number(payload.spQuote).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    // Show daily returns
    const dailyIndexEl = document.getElementById('dailyIndexReturnsList');
    const dailyMyEl = document.getElementById('dailyMyReturn');
    if (payload.dailyReturns) {
      if (dailyIndexEl && payload.dailyReturns.indexDaily && payload.dailyReturns.indexDaily.length) {
        dailyIndexEl.innerHTML = payload.dailyReturns.indexDaily.map(h => {
          const pct = (h.return * 100).toFixed(2);
          const color = pct >= 0 ? '#4caf84' : '#f06292';
          return `<div class="holding-item">
            <span class="holding-symbol">${h.symbol}</span>
            <span class="holding-pct" style="color:${color}">${pct >= 0 ? '+' : ''}${pct}%</span>
          </div>`;
        }).join('');
      }
      if (dailyMyEl && payload.dailyReturns.myReturn != null) {
        const pct = (payload.dailyReturns.myReturn * 100).toFixed(2);
        dailyMyEl.textContent = (pct >= 0 ? '+' : '') + pct + '%';
        dailyMyEl.style.color = pct >= 0 ? '#4caf84' : '#f06292';
      }
    }

    // Show current model score
    const scoreEl = document.getElementById('modelScore');
    if (scoreEl && payload.modelScore != null) {
      const pct = (payload.modelScore * 100).toFixed(2);
      scoreEl.textContent = pct + '%';
      // Color: green if high, yellow if mid, red if low
      scoreEl.style.color = pct >= 50 ? '#4caf84' : pct >= 25 ? '#f6c343' : '#f06292';
    }
  } catch (err) {
    meta.textContent = `${err.message}. If you opened index.html directly, run a local web server first.`;
    console.error(err);
  }
})();
