async function initCameraCard(card) {
    const video = card.querySelector("video");
    const canvas = card.querySelector("canvas");
    const message = card.querySelector("[data-camera-message]");
    const startButton = card.querySelector("[data-camera-start]");
    const captureButton = card.querySelector("[data-camera-capture]");
    const autoButton = card.querySelector("[data-camera-auto]");
    const endpoint = card.dataset.endpoint;
    const autoScan = card.dataset.autoScan === "true";
    const scanIntervalMs = Number(card.dataset.scanIntervalMs || 4000);
    let stream;
    let monitorTimer = null;
    let isProcessing = false;
    let isMonitoring = false;

    function setMessage(text, isError = false) {
        message.textContent = text;
        message.classList.toggle("error-text", isError);
    }

    function setMonitoringState(active) {
        isMonitoring = active;
        if (autoButton) {
            autoButton.textContent = active ? "Stop Monitoring" : "Start Monitoring";
            autoButton.classList.toggle("button-danger", active);
            autoButton.classList.toggle("button-ghost", !active);
            autoButton.setAttribute("aria-pressed", String(active));
        }
    }

    async function startCamera() {
        if (video.srcObject) {
            return true;
        }

        try {
            stream = await navigator.mediaDevices.getUserMedia({ video: true, audio: false });
            video.srcObject = stream;
            setMessage(
                autoScan
                    ? "Camera started. Continuous attendance monitoring is ready."
                    : "Camera started. Keep one face centered in frame."
            );
            return true;
        } catch (error) {
            setMessage("Could not access the camera. Please allow camera permission and try again.", true);
            return false;
        }
    }

    function stopMonitoring() {
        if (monitorTimer) {
            window.clearInterval(monitorTimer);
            monitorTimer = null;
        }
        setMonitoringState(false);
    }

    async function captureFrame({ silent = false } = {}) {
        if (isProcessing) {
            return;
        }

        if (!video.srcObject) {
            if (silent) {
                return;
            }
            setMessage("Start the camera before capturing.", true);
            return;
        }

        canvas.width = video.videoWidth;
        canvas.height = video.videoHeight;
        const context = canvas.getContext("2d");
        context.drawImage(video, 0, 0, canvas.width, canvas.height);

        isProcessing = true;
        if (captureButton) {
            captureButton.disabled = true;
        }
        if (!silent) {
            setMessage("Processing face...");
        }

        try {
            const response = await fetch(endpoint, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ image: canvas.toDataURL("image/jpeg", 0.92) }),
            });
            const data = await response.json();
            if (!response.ok || !data.ok) {
                throw new Error(data.error || "Request failed.");
            }
            setMessage(data.message);
        } catch (error) {
            const errorMessage = error.message || "Something went wrong while processing the image.";
            const shouldBeQuiet =
                silent &&
                (errorMessage.includes("No face detected") ||
                    errorMessage.includes("Multiple faces detected") ||
                    errorMessage.includes("in cooldown"));

            if (!shouldBeQuiet) {
                setMessage(errorMessage, true);
            }
        } finally {
            isProcessing = false;
            if (captureButton) {
                captureButton.disabled = false;
            }
        }
    }

    async function startMonitoring() {
        const started = await startCamera();
        if (!started) {
            return;
        }

        stopMonitoring();
        setMonitoringState(true);
        setMessage("Continuous monitoring started. The scanner will keep checking for attendance.");
        monitorTimer = window.setInterval(() => {
            captureFrame({ silent: true });
        }, scanIntervalMs);
    }

    startButton?.addEventListener("click", startCamera);
    captureButton?.addEventListener("click", () => captureFrame());
    autoButton?.addEventListener("click", () => {
        if (isMonitoring) {
            stopMonitoring();
            setMessage("Continuous monitoring stopped.");
        } else {
            startMonitoring();
        }
    });
}

document.addEventListener("DOMContentLoaded", () => {
    document.querySelectorAll("[data-camera]").forEach((card) => {
        initCameraCard(card);
    });
});
