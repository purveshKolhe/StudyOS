const form = document.getElementById('generate-form');
const topicInput = document.getElementById('topic');
const genBtn = document.getElementById('generate-btn');
const statusEl = document.getElementById('status');
const resultEl = document.getElementById('result');
const errorEl = document.getElementById('error');
const fileNameEl = document.getElementById('filename');
const downloadLink = document.getElementById('download-link');
const againBtn = document.getElementById('again');

function show(el) { el.classList.remove('hidden'); }
function hide(el) { el.classList.add('hidden'); }

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  const topic = topicInput.value.trim();
  if (!topic) return;

  hide(resultEl); hide(errorEl); show(statusEl);
  genBtn.disabled = true;

  try {
    const res = await fetch('/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ topic })
    });

    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Failed to generate');

    fileNameEl.textContent = data.filename;
    downloadLink.href = data.download_url;

    hide(statusEl); show(resultEl);
  } catch (err) {
    hide(statusEl); hide(resultEl);
    errorEl.textContent = err.message || 'Something went wrong';
    show(errorEl);
  } finally {
    genBtn.disabled = false;
  }
});

againBtn.addEventListener('click', () => {
  topicInput.value = '';
  hide(resultEl);
  topicInput.focus();
});
