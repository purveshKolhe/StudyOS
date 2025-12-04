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

const generateVideoBtn = document.getElementById("generate-video-btn");
const videoResult = document.getElementById("video-result");
const videoLoading = document.getElementById("video-loading");
const videoSuccess = document.getElementById("video-success");
const videoDownloadLink = document.getElementById("video-download-link");

let currentFilename = "";

form.addEventListener("submit", async(e) => {
    e.preventDefault();

    // Reset UI
    step1.classList.add("hidden");
    step2.classList.remove("hidden");
    step3.classList.add("hidden");
    errorEl.classList.add("hidden"); // Changed from errorDiv to errorEl
    videoResult.classList.add("hidden");

    const topic = document.getElementById("topic").value;
    // const description = document.getElementById("description").value; // Capture description if needed by backend later

    try {
        const response = await fetch("/generate", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            },
            body: JSON.stringify({ topic: topic }),
        });

        const data = await response.json();

        if (response.ok) {
            step2.classList.add("hidden");
            step3.classList.remove("hidden");

            fileNameEl.textContent = data.filename; // Changed from filenameDisplay to fileNameEl
            downloadLink.href = data.download_url;
            currentFilename = data.filename; // Store for video generation
        } else {
            throw new Error(data.error || "An error occurred");
        }
    } catch (err) {
        step2.classList.add("hidden");
        step1.classList.remove("hidden");
        errorEl.textContent = err.message; // Changed from errorDiv to errorEl
        errorEl.classList.remove("hidden"); // Changed from errorDiv to errorEl
    }
});

generateVideoBtn.addEventListener("click", async() => {
    if (!currentFilename) return;

    videoResult.classList.remove("hidden");
    videoLoading.classList.remove("hidden");
    videoSuccess.classList.add("hidden");
    generateVideoBtn.disabled = true;
    generateVideoBtn.textContent = "Generating...";

    try {
        const response = await fetch("/generate_video", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            },
            body: JSON.stringify({ filename: currentFilename }),
        });

        const data = await response.json();

        if (response.ok) {
            videoLoading.classList.add("hidden");
            videoSuccess.classList.remove("hidden");
            videoDownloadLink.href = data.video_url;
        } else {
            throw new Error(data.error || "Video generation failed");
        }
    } catch (err) {
        videoLoading.classList.add("hidden");
        alert(`Error: ${err.message}`);
    } finally {
        generateVideoBtn.disabled = false;
        generateVideoBtn.innerHTML = '<i data-lucide="video" width="20" height="20"></i> Generate Video';
        lucide.createIcons();
    }
});

againBtn.addEventListener("click", () => {
    step3.classList.add("hidden");
    step1.classList.remove("hidden");
    form.reset();
    videoResult.classList.add("hidden");
    currentFilename = "";
    topicInput.focus(); // Moved this line inside the event listener
});
const themeToggle = document.getElementById('theme-toggle');

themeToggle.addEventListener('click', () => {
    document.body.classList.toggle('dark-theme');
    const isDarkMode = document.body.classList.contains('dark-theme');
    themeToggle.innerHTML = isDarkMode ? '<i data-lucide="moon"></i>' : '<i data-lucide="sun"></i>';
    lucide.createIcons();
    localStorage.setItem('theme', isDarkMode ? 'dark' : 'light');
});

// Check for saved theme preference
document.addEventListener('DOMContentLoaded', () => {
    const savedTheme = localStorage.getItem('theme');
    if (savedTheme === 'dark') {
        document.body.classList.add('dark-theme');
        themeToggle.innerHTML = '<i data-lucide="moon"></i>';
        lucide.createIcons();
    }
})