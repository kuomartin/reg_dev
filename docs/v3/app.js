// DOM Elements
const videoContainer = document.getElementById("video-container");
const qrReaderElem = document.getElementById("qr-reader");
const qrInput = document.getElementById("qr-input");
const qrcode = document.getElementById("qrcode");
const startBtn = document.getElementById("start-button");
const startBtnText = document.getElementById("start-btn-text");
const flipBtn = document.getElementById("flip-button");
const checkinBtn = document.getElementById("checkin-button");
const toastContainer = document.getElementById("custom-toast-container");
const toggleQrBtn = document.getElementById("toggle-qr-btn");
const qrPanel = document.getElementById("qr-panel");
const qrPanelQrcode = document.getElementById("qr-panel-qrcode");
const continuousScanToggle = document.getElementById("continuous-scan");
const sheetsTabsContainer = document.getElementById("sheets-tabs-container");
const sheetsTabsUl = document.getElementById("sheets-tabs");
const connectionStatus = document.getElementById("connection-status");
const connectionStatusIcon = document.getElementById("connection-status-icon");
const connectionStatusText = document.getElementById("connection-status-text");
const switchToOldLink = document.getElementById("switch-to-old-link");
const historyContainer = document.getElementById("history-container");
const historyList = document.getElementById("history-list");
const clearHistoryBtn = document.getElementById("clear-history");

let status = "idle"; // scanning, processing, idle
let currentCamera = "environment";
let activeSheetName = "學員報到";
let lastScannedCode = null;
let lastScanTimestamp = 0;
let dragged = false; // Add dragged to global state
let checkInHistory = []; // 存儲歷史紀錄
const MAX_HISTORY = 20; // 最大紀錄筆數
const SCAN_FPS = 12;
const SCAN_QRBOX = { width: 260, height: 260 };
const SCAN_FORMATS = [
  Html5QrcodeSupportedFormats.QR_CODE,
  Html5QrcodeSupportedFormats.CODE_128,
  Html5QrcodeSupportedFormats.CODE_39,
  Html5QrcodeSupportedFormats.EAN_13,
  Html5QrcodeSupportedFormats.EAN_8,
  Html5QrcodeSupportedFormats.UPC_A,
  Html5QrcodeSupportedFormats.UPC_E,
  Html5QrcodeSupportedFormats.ITF,
  Html5QrcodeSupportedFormats.PDF_417,
  Html5QrcodeSupportedFormats.DATA_MATRIX,
];
let html5Qrcode = null;
let isScannerRunning = false;

let activeToast = null;

function setConnectionStatus(type) {
  if (!connectionStatus || !connectionStatusIcon || !connectionStatusText)
    return;
  if (type === "connected") {
    connectionStatus.style.background = "#e8f5e9";
    connectionStatus.style.color = "#2e7d32";
    connectionStatusIcon.className = "bi bi-check-circle-fill";
    connectionStatusText.textContent = "Bridge 已連線";
    return;
  }
  if (type === "failed") {
    connectionStatus.style.background = "#ffebee";
    connectionStatus.style.color = "#c62828";
    connectionStatusIcon.className = "bi bi-x-circle-fill";
    connectionStatusText.textContent = "Bridge 連線失敗";
    return;
  }
  if (type === "no-id") {
    connectionStatus.style.background = "#eceff1";
    connectionStatus.style.color = "#455a64";
    connectionStatusIcon.className = "bi bi-link-45deg";
    connectionStatusText.textContent = "缺少部署 ID";
    return;
  }
  connectionStatus.style.background = "#fff3cd";
  connectionStatus.style.color = "#8a6d3b";
  connectionStatusIcon.className = "bi bi-arrow-repeat";
  connectionStatusText.textContent = "Bridge 連線中";
}

// ====== 網址參數解析與 Iframe 初始化 ======
const urlParams = new URLSearchParams(window.location.search);
const deployId = urlParams.get("id");
const gasIframe = document.getElementById("gas-bridge");
const urlWarning = document.getElementById("url-warning");
let gasUrl = null;
if (deployId) {
  setConnectionStatus("connecting");
  gasUrl = `https://script.google.com/macros/s/${deployId}/exec?bridge=1`;
  gasIframe.src = gasUrl;
} else {
  // 缺少 ID 時的警告與鎖定處理... (保留你原本寫的)
  if (urlWarning) urlWarning.style.display = "block";
  if (startBtn) startBtn.disabled = true;
  if (checkinBtn) checkinBtn.disabled = true;
  if (qrInput) qrInput.disabled = true;
  setConnectionStatus("no-id");
}

const appUrl = window.location.href;
document.getElementById("qr-url-input").value = appUrl;
document.getElementById("qr-panel-link").href = appUrl;
if (switchToOldLink) {
  switchToOldLink.href = `old.html${window.location.search || ""}`;
}
// =======================================

// ====== 秘密地道通訊邏輯 ======
let gasPort = null; // 用來儲存地道出口
let isGasReady = false; // 狀態標記
let gasReadyResolve = null;
const gasReadyPromise = new Promise((resolve) => {
  gasReadyResolve = resolve;
});

// 監聽 GAS 拋出來的地道出口
window.addEventListener("message", (event) => {
  if (event.data && event.data.action === "GAS_READY") {
    gasPort = event.ports[0]; // 成功接住 port2
    isGasReady = true;
    if (gasReadyResolve) gasReadyResolve();
    setConnectionStatus("connected");
    console.log("✅ GAS 秘密地道連線成功！");
    loadSheets();
  }
});

function loadSheets() {
  callGas("getCheckInSheets")
    .then((sheets) => {
      if (sheets && sheets.length > 0) {
        renderSheetSwitcher(sheets);
      }
    })
    .catch((err) => console.error("無法載入報到分頁:", err));
}

/**
 * 統一更新當前選中的報到分頁
 */
function updateActiveSheet(sheetName) {
  activeSheetName = sheetName;
  console.log("當前分頁更新為:", activeSheetName);

  // 1. 同步更新 Tabs 視覺狀態
  sheetsTabsUl.querySelectorAll("a").forEach((a) => {
    if (a.textContent === sheetName) {
      a.classList.add("active");
    } else {
      a.classList.remove("active");
    }
  });

  // 2. 同步更新 Select 選單狀態
  const sheetsSelect = document.getElementById("sheets-select");
  if (sheetsSelect && sheetsSelect.value !== sheetName) {
    sheetsSelect.value = sheetName;
    M.FormSelect.init(sheetsSelect);
  }
}

function renderSheetSwitcher(sheets) {
  const sheetsSelect = document.getElementById("sheets-select");
  const sheetsSelectContainer = document.getElementById(
    "sheets-select-container",
  );

  if (!sheets || sheets.length <= 1) {
    sheetsTabsContainer.style.display = "none";
    if (sheetsSelectContainer) sheetsSelectContainer.style.display = "none";
    activeSheetName = sheets && sheets.length === 1 ? sheets[0] : "學員報到";
    return;
  }

  sheetsTabsUl.innerHTML = "";
  if (sheetsSelect) sheetsSelect.innerHTML = "";

  sheets.forEach((sheet, index) => {
    // 1. 建立 Tab (UL)
    const li = document.createElement("li");
    li.className = "tab";
    const a = document.createElement("a");
    a.href = "#!";
    a.textContent = sheet;
    if (
      sheet === activeSheetName ||
      (index === 0 && !sheets.includes(activeSheetName))
    ) {
      a.classList.add("active");
      activeSheetName = sheet;
    }

    a.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (dragged) return;
      updateActiveSheet(sheet);
    };

    li.appendChild(a);
    sheetsTabsUl.appendChild(li);

    // 2. 建立 Option (Select)
    if (sheetsSelect) {
      const option = document.createElement("option");
      option.value = sheet;
      option.textContent = sheet;
      if (sheet === activeSheetName) option.selected = true;
      sheetsSelect.appendChild(option);
    }
  });

  sheetsTabsContainer.style.display = "block";
  if (sheetsSelect) {
    M.FormSelect.init(sheetsSelect);
    sheetsSelect.onchange = (e) => {
      updateActiveSheet(e.target.value);
    };
  }

  // 3. 拖曳滾動邏輯 (Mousedown/Mousemove)
  let isDown = false;
  let startX;
  let scrollLeft;

  if (!sheetsTabsUl.dataset.dragInit) {
    sheetsTabsUl.addEventListener("mousedown", (e) => {
      isDown = true;
      dragged = false;
      sheetsTabsUl.classList.add("dragging");
      startX = e.pageX - sheetsTabsUl.offsetLeft;
      scrollLeft = sheetsTabsUl.scrollLeft;
    });

    sheetsTabsUl.addEventListener("mouseleave", () => {
      isDown = false;
      sheetsTabsUl.classList.remove("dragging");
    });

    sheetsTabsUl.addEventListener("mouseup", () => {
      isDown = false;
      sheetsTabsUl.classList.remove("dragging");
      setTimeout(() => {
        dragged = false;
      }, 50);
    });

    sheetsTabsUl.addEventListener("mousemove", (e) => {
      if (!isDown) return;
      e.preventDefault();
      const x = e.pageX - sheetsTabsUl.offsetLeft;
      const walk = (x - startX) * 1.5;
      if (Math.abs(walk) > 5) dragged = true;
      sheetsTabsUl.scrollLeft = scrollLeft - walk;
    });
    sheetsTabsUl.dataset.dragInit = "true";
  }
}

function waitForGasReady(timeoutMs = 8000) {
  if (isGasReady && gasPort) {
    return Promise.resolve();
  }
  return Promise.race([
    gasReadyPromise,
    new Promise((_, reject) => {
      setTimeout(() => {
        setConnectionStatus("failed");
        reject(
          "GAS 尚未連線。若你看到 403，請先在新分頁開啟 GAS 的 /exec 並完成授權，再回來重試。",
        );
      }, timeoutMs);
    }),
  ]);
}

function callGas(action, payload) {
  return waitForGasReady().then(() => {
    return new Promise((resolve, reject) => {
      const messageId = Date.now().toString() + Math.random().toString();

      // 在地道口監聽專屬的結果回傳
      const handler = function (event) {
        if (event.data && event.data.id === messageId) {
          gasPort.removeEventListener("message", handler); // 收到後移除監聽器
          if (event.data.status === "success") {
            resolve(event.data.result);
          } else {
            reject(event.data.error);
          }
        }
      };

      gasPort.addEventListener("message", handler);
      gasPort.start(); // 啟動監聽 (使用 addEventListener 時必加)

      // 直接從地道把指令塞進去 (完全繞過 window.postMessage，不會被擋)
      gasPort.postMessage({
        id: messageId,
        action: action,
        payload: payload,
      });
    });
  });
}

function showResult(message, type) {
  // Remove previous loading toast if it exists
  if (activeToast && activeToast.dataset.type === "loading") {
    hideToast(activeToast);
  }

  console.log(JSON.stringify(message));

  const toast = document.createElement("div");
  toast.dataset.type = type;

  let iconName = "";
  let bgColor = "";

  if (type === "success") {
    iconName = "bi-check-circle-fill";
    bgColor = "green darken-1";
  } else if (type === "error") {
    iconName = "bi-x-circle-fill";
    bgColor = "red darken-1";
  } else if (type === "loading") {
    iconName = "bi-arrow-repeat";
    bgColor = "blue darken-1";
  }

  toast.className = `custom-toast ${bgColor}`;
  toast.innerHTML = `<i class="bi ${iconName} ${type === "loading" ? "loading-icon" : ""}"></i><span class="toast-message"></span>`;
  toast.querySelector(".toast-message").textContent = message;

  toastContainer.appendChild(toast);
  activeToast = toast;

  // Auto hide for status messages (not loading)
  if (type !== "loading") {
    setTimeout(() => {
      hideToast(toast);
    }, 3500);
  }

  toast.onclick = () => hideToast(toast);
}

function hideToast(toast) {
  if (!toast || !toast.classList) return;
  toast.classList.add("hide");
  setTimeout(() => {
    if (toast && toast.parentNode) {
      toast.parentNode.removeChild(toast);
    }
    if (activeToast === toast) activeToast = null;
  }, 300);
}

function hideResult() {
  if (activeToast) hideToast(activeToast);
}

function addToHistory(message, type) {
  const time = new Date().toLocaleTimeString([], {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
  const entry = { message, type, time };

  checkInHistory.unshift(entry); // 新的在前面
  if (checkInHistory.length > MAX_HISTORY) {
    checkInHistory.pop();
  }

  renderHistory();
}

function renderHistory() {
  if (checkInHistory.length === 0) {
    historyContainer.style.display = "none";
    return;
  }

  historyContainer.style.display = "block";
  historyList.innerHTML = "";

  checkInHistory.forEach((item) => {
    const li = document.createElement("li");
    li.className = "collection-item";
    li.style.display = "flex";
    li.style.alignItems = "center";
    li.style.gap = "12px";
    li.style.padding = "10px 16px";
    li.style.borderBottom = "1px solid #f0f0f0";

    const icon = document.createElement("i");
    icon.className =
      item.type === "success"
        ? "bi bi-check-circle-fill green-text"
        : "bi bi-x-circle-fill red-text";
    icon.style.fontSize = "1.1rem";

    const content = document.createElement("div");
    content.style.flex = "1";
    content.style.whiteSpace = "pre-wrap"; // 支援換行字元
    content.innerHTML = `<span style="font-weight: 500; font-size: 0.95rem;">${item.message}</span>`;

    const timeSpan = document.createElement("span");
    timeSpan.className = "grey-text";
    timeSpan.style.fontSize = "0.75rem";
    timeSpan.textContent = item.time;

    li.appendChild(icon);
    li.appendChild(content);
    li.appendChild(timeSpan);
    historyList.appendChild(li);
  });
}

if (clearHistoryBtn) {
  clearHistoryBtn.onclick = (e) => {
    e.preventDefault();
    checkInHistory = [];
    renderHistory();
  };
}

function updateUI() {
  if (status === "scanning") {
    videoContainer.style.display = "block";
    flipBtn.disabled = false;
    startBtnText.textContent = "停止掃描";
    startBtn.className = "btn waves-effect waves-light red darken-1";
    startBtn.disabled = false;
  } else if (status === "processing") {
    startBtn.disabled = true;
    flipBtn.disabled = true;
  } else {
    // idle
    videoContainer.style.display = "none";
    flipBtn.disabled = true;
    startBtnText.textContent = "開始掃描";
    startBtn.className = "btn waves-effect waves-light blue darken-1";
    startBtn.disabled = false;
  }

  if (currentCamera === "user") {
    videoContainer.classList.add("mirrored");
  } else {
    videoContainer.classList.remove("mirrored");
  }
}

// Result message click handler is now handled inside showResult

function checkInStudent(id) {
  if (!id || id.trim() === "") return;
  status = "processing";
  updateUI();
  showResult(`${id} 報到中...`, "loading");
  qrInput.value = "";

  callGas("checkInStudent", {
    id: id,
    sheetName: activeSheetName,
  })
    .then((result) => {
      const isContinuous = continuousScanToggle.checked;
      status = isContinuous ? "scanning" : "idle";
      updateUI();

      if (result && result.message) {
        showResult(result.message, result.success ? "success" : "error");
        addToHistory(result.message, result.success ? "success" : "error");
      } else {
        showResult("未知錯誤", "error");
        addToHistory(`${id} 發生未知錯誤`, "error");
      }
    })
    .catch((error) => {
      const isContinuous = continuousScanToggle.checked;
      status = isContinuous ? "scanning" : "idle";
      updateUI();
      showResult(`報到失敗: ${error}`, "error");
      addToHistory(`${id} 報到失敗: ${error}`, "error");
    });
}

function handleScanSuccess(decodedText) {
  if (status !== "scanning") return;
  const data = (decodedText || "").trim();
  if (!data) return;

  const now = Date.now();
  if (
    data === lastScannedCode &&
    now - lastScanTimestamp < DUPLICATE_SCAN_COOLDOWN_MS
  ) {
    return;
  }

  lastScannedCode = data;
  lastScanTimestamp = now;
  qrInput.value = data;
  M.updateTextFields();

  if (!continuousScanToggle.checked) {
    stopScanning().finally(() => checkInStudent(data));
    return;
  }

  checkInStudent(data);
}

function handleScanError() {
  // Intentionally no-op: html5-qrcode calls this frequently when no code is in frame.
}

async function startCameraWithFallback(scannerConfig) {
  const cameraCandidates = [
    { facingMode: { exact: currentCamera } },
    { facingMode: currentCamera },
  ];

  let lastError = null;
  for (const candidate of cameraCandidates) {
    try {
      await html5Qrcode.start(
        candidate,
        scannerConfig,
        handleScanSuccess,
        handleScanError,
      );
      isScannerRunning = true;
      return;
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error("No available camera configuration.");
}

async function startScanning() {
  if (status !== "idle") return;

  if (!window.Html5Qrcode || !window.Html5QrcodeSupportedFormats) {
    showResult("掃描器載入失敗，請重新整理頁面。", "error");
    return;
  }

  if (!html5Qrcode) {
    html5Qrcode = new Html5Qrcode(qrReaderElem.id, {
      formatsToSupport: SCAN_FORMATS,
      verbose: false,
    });
  }

  const scannerConfig = {
    fps: SCAN_FPS,
    qrbox: SCAN_QRBOX,
    disableFlip: false,
  };

  try {
    status = "scanning";
    updateUI();
    videoContainer.scrollIntoView({
      behavior: "smooth",
      block: "center",
    });
    await startCameraWithFallback(scannerConfig);
  } catch (error) {
    status = "idle";
    updateUI();
    isScannerRunning = false;
    showResult("無法開啟相機，請檢查權限及連線安全性。", "error");
    console.error("startScanning failed:", error);
  }
}

async function stopScanning() {
  status = "idle";
  lastScannedCode = null;
  lastScanTimestamp = 0;

  if (html5Qrcode && isScannerRunning) {
    try {
      await html5Qrcode.stop();
    } catch (error) {
      console.warn("stop scanner failed:", error);
    }
    try {
      await html5Qrcode.clear();
    } catch (error) {
      console.warn("clear scanner failed:", error);
    }
  }

  isScannerRunning = false;
  updateUI();
}

startBtn.addEventListener("click", async () => {
  if (status === "scanning") await stopScanning();
  else if (status === "idle") await startScanning();
});

flipBtn.addEventListener("click", async () => {
  currentCamera = currentCamera === "environment" ? "user" : "environment";
  if (status === "scanning") {
    await stopScanning();
    await startScanning();
  } else {
    updateUI();
  }
});

checkinBtn.addEventListener("click", () => {
  const id = qrInput.value.trim();
  if (id) checkInStudent(id);
  else showResult("請輸入學員代號或進行掃描", "error");
});

qrInput.addEventListener("keypress", (e) => {
  if (e.key === "Enter") {
    const id = qrInput.value.trim();
    if (id) checkInStudent(id);
    else showResult("請輸入學員代號", "error");
  }
});

// QR Code Logic
let qrCodeInstance = null;
const qrUrlInput = document.getElementById("qr-url-input");
const qrPanelLink = document.getElementById("qr-panel-link");

function updateQrCode(text) {
  if (!qrCodeInstance) {
    qrCodeInstance = new QRCode(qrPanelQrcode, {
      text: text || " ",
      width: 180,
      height: 180,
      colorDark: "#2196f3",
      colorLight: "#ffffff",
      correctLevel: QRCode.CorrectLevel.H,
    });
  } else {
    qrCodeInstance.clear();
    qrCodeInstance.makeCode(text || " ");
  }
  qrPanelLink.href = text || "javascript:void(0)";
}

qrUrlInput.addEventListener("input", (e) => {
  updateQrCode(e.target.value);
});

toggleQrBtn.addEventListener("click", () => {
  const isActive = qrPanel.classList.toggle("active");

  if (isActive) {
    // Initial generation or check for reset if needed
    if (!qrCodeInstance) {
      updateQrCode(qrUrlInput.value || appUrl);
    }
  } else {
    // Reset to appUrl after closing as requested
    qrUrlInput.value = appUrl;
    updateQrCode(appUrl);
  }

  // Toggle icon
  const icon = toggleQrBtn.querySelector("i");
  icon.className = isActive ? "bi bi-x-lg" : "bi bi-qr-code";
});

// Initialize tooltips and modals
document.addEventListener("DOMContentLoaded", function () {
  var tooltips = document.querySelectorAll(".tooltipped");
  M.Tooltip.init(tooltips);
});
