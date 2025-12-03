const form = document.getElementById('generate-form');
const topicInput = document.getElementById('topic');
const genBtn = document.getElementById('generate-btn');
const statusEl = document.getElementById('status');
const resultEl = document.getElementById('result');
const errorEl = document.getElementById('error');
const fileNameEl = document.getElementById('filename');
const downloadLink = document.getElementById('download-link');
const againBtn = document.getElementById('again');
const statusMessageEl = document.getElementById('status-message');

const step1 = document.getElementById('step-1');
const step2 = document.getElementById('step-2');
const step3 = document.getElementById('step-3');

function show(el) { el.classList.remove('hidden'); }

function hide(el) { el.classList.add('hidden'); }

function goToStep(step) {
    hide(step1);
    hide(step2);
    hide(step3);
    if (step === 1) show(step1);
    if (step === 2) show(step2);
    if (step === 3) show(step3);
}

const statusMessages = [
    "Drafting content...",
    "Designing slides...",
    "Applying styles...",
    "Finalizing presentation...",
];

let statusInterval;

function cycleStatusMessages() {
    let i = 0;
    statusInterval = setInterval(() => {
        statusMessageEl.textContent = statusMessages[i];
        i = (i + 1) % statusMessages.length;
    }, 2000);
}

form.addEventListener('submit', async(e) => {
    e.preventDefault();
    const topic = topicInput.value.trim();
    if (!topic) return;

    goToStep(2);
    cycleStatusMessages();
    genBtn.disabled = true;
    hide(errorEl);

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

        clearInterval(statusInterval);
        goToStep(3);
    } catch (err) {
        clearInterval(statusInterval);
        goToStep(1);
        errorEl.textContent = err.message || 'Something went wrong';
        show(errorEl);
    } finally {
        genBtn.disabled = false;
    }
});

againBtn.addEventListener('click', () => {
    topicInput.value = '';
    goToStep(1);
    hide(errorEl);
    topicInput.focus();
});