// Office MIS Claim Assistant - Popup Script
document.addEventListener("DOMContentLoaded", async () => {
    const connDot = document.getElementById("conn-dot");
    const connLabel = document.getElementById("conn-label");
    const alertBox = document.getElementById("alert-box");
    const sendBtn = document.getElementById("send-btn");
    const dashboardBtn = document.getElementById("dashboard-btn");

    let extractedData = null;

    function showAlert(msg, type = "success") {
        alertBox.className = `alert-box alert-${type}`;
        alertBox.innerText = msg;
        alertBox.style.display = "block";
    }

    function renderClaimPreview(data) {
        if (!data) return;
        extractedData = data;
        document.getElementById("prev-ins").innerText = data.ins_co || "--";
        document.getElementById("prev-claim").innerText = data.claim_no || "(Not detected)";
        document.getElementById("prev-policy").innerText = data.policy_no || "--";
        document.getElementById("prev-veh").innerText = data.vehicle_no || "--";
        document.getElementById("prev-name").innerText = data.insured_name || "--";

        const modelEl = document.getElementById("prev-model");
        if (modelEl) modelEl.innerText = data.make_model || "--";

        const wsEl = document.getElementById("prev-workshop");
        if (wsEl) wsEl.innerText = data.workshop || "--";
    }

    // 1. Check local server connection
    let activeEndpoint = "http://127.0.0.1:5055";
    try {
        let res;
        try {
            res = await fetch("http://127.0.0.1:5055/api/config");
            activeEndpoint = "http://127.0.0.1:5055";
        } catch (e1) {
            res = await fetch("http://localhost:5055/api/config");
            activeEndpoint = "http://localhost:5055";
        }

        if (res && res.ok) {
            connDot.classList.add("connected");
            connLabel.innerText = "Connected";
            connLabel.style.color = "#15803d";
        } else {
            connLabel.innerText = "Server error";
        }
    } catch (e) {
        connLabel.innerText = "App Offline";
        showAlert("Desktop app offline. Please launch OfficeAutomation.exe.", "error");
    }

    // 2. Read from chrome.storage.local first (handles iframe extractions immediately)
    try {
        if (chrome && chrome.storage && chrome.storage.local) {
            chrome.storage.local.get(["last_extracted_claim", "extracted_at"], (res) => {
                if (res && res.last_extracted_claim) {
                    renderClaimPreview(res.last_extracted_claim);
                }
            });
        }
    } catch (e) {
        console.warn("Storage read failed:", e);
    }

    // 3. Query active tab for latest real-time claim data
    try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (tab && tab.id) {
            chrome.tabs.sendMessage(tab.id, { action: "get_claim_data" }, (response) => {
                if (chrome.runtime.lastError || !response || !response.data) {
                    return;
                }
                renderClaimPreview(response.data);
            });
        }
    } catch (err) {
        console.error("Tab query failed:", err);
    }

    // 4. Send button handler
    sendBtn.addEventListener("click", async () => {
        if (!extractedData || (!extractedData.claim_no && !extractedData.vehicle_no && !extractedData.policy_no)) {
            showAlert("No claim data found on current page.", "error");
            return;
        }

        sendBtn.disabled = true;
        sendBtn.innerText = "Sending to Excel...";

        try {
            const res = await fetch(`${activeEndpoint}/api/web-claim`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ record: extractedData })
            });
            const data = await res.json();
            if (data.success) {
                showAlert(`Appended Ref #${data.ref_no} into Master Excel!`, "success");
                sendBtn.innerText = "✅ Sent to MIS";
                setTimeout(() => window.close(), 2200);
            } else {
                showAlert(data.error || "Failed to save claim.", "error");
                sendBtn.disabled = false;
                sendBtn.innerText = "⚡ Send Claim to MIS Excel";
            }
        } catch (err) {
            showAlert("Could not connect to Desktop App on localhost:5055.", "error");
            sendBtn.disabled = false;
            sendBtn.innerText = "⚡ Send Claim to MIS Excel";
        }
    });

    // 5. Open dashboard button
    dashboardBtn.addEventListener("click", () => {
        chrome.tabs.create({ url: "http://localhost:5055" });
    });
});
