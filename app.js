async function loadData() {
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
      plugins: { legend: { display: true } },
      scales: {
        x: { ticks: { maxTicksLimit: 12 } },
        y: {
          type: 'linear',
          position: 'left',
          beginAtZero: false,
          title: { display: true, text: 'S&P' }
        },
        y1: {
          type: 'linear',
          position: 'right',
          beginAtZero: false,
          grid: { drawOnChartArea: false },
          title: { display: true, text: 'Model Score' }
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

    const values = drawChart(rows);
    const latest = values[values.length - 1];
    meta.textContent = `Sheet: ${payload.sheet} • Points: ${rows.length} • Latest S&P: ${latest}`;
  } catch (err) {
    meta.textContent = `${err.message}. If you opened index.html directly, run a local web server first.`;
    console.error(err);
  }
})();
