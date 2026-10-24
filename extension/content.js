// Office MIS Claim Assistant - Universal Insurance Portal Content Script
(function () {
    // Avoid double injection in same frame
    if (window.__officeMisExtensionInjected) return;
    window.__officeMisExtensionInjected = true;

    // In-memory cache for extracted data
    window.__cachedClaimData = null;

    // Helper: Normalize clean text
    function clean(str) {
        if (!str) return "";
        return String(str)
            .replace(/[\u200B-\u200D\uFEFF]/g, "")
            .replace(/\s+/g, " ")
            .trim();
    }

    // Helper: Extract clean date portion if string has time attached (e.g. "20/10/2026 00:00:00")
    function cleanDate(str) {
        if (!str) return "";
        const m = str.match(/(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})/);
        return m ? m[1] : clean(str);
    }

    // Helper: Safely get the value of any DOM element (input, select, textarea, div, span)
    function getElementValue(el) {
        if (!el) return "";
        const tag = el.tagName ? el.tagName.toUpperCase() : "";

        if (tag === "SELECT") {
            if (el.selectedIndex >= 0 && el.options && el.options[el.selectedIndex]) {
                const optText = clean(el.options[el.selectedIndex].text);
                const optVal = clean(el.options[el.selectedIndex].value);
                // Avoid returning placeholder selections like "Select", "--Select--", "Choose"
                if (/^(select|--|none|choose|\s*)$/i.test(optText)) return "";
                return optText || optVal;
            }
            return clean(el.value);
        }

        if (tag === "INPUT" || tag === "TEXTAREA") {
            if (el.type === "checkbox" || el.type === "radio") {
                return el.checked ? clean(el.value || "Yes") : "";
            }
            if (el.value !== undefined && el.value !== null && clean(el.value) !== "") {
                return clean(el.value);
            }
            const attrVal = el.getAttribute("value");
            if (attrVal && clean(attrVal) !== "") return clean(attrVal);
            return "";
        }

        // For plain elements (span, div, td, b)
        return clean(el.innerText || el.textContent);
    }

    // Universal DOM Search: Finds value across any portal architecture (Tables, Bootstrap grids, Angular, React, ASP.NET)
    function findValueByAnchors(anchors) {
        for (const anchor of anchors) {
            let re;
            try {
                re = new RegExp(anchor, "i");
            } catch (e) {
                re = new RegExp(anchor.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), "i");
            }

            // --- STRATEGY 1: Search Form Labels & Field Headers ---
            const labelCandidates = document.querySelectorAll(
                "label, .control-label, .form-label, .field-label, [class*='label'], [class*='title'], [class*='header'], th, dt, b, strong, span, div, p"
            );

            for (const el of labelCandidates) {
                // Must be short enough to be a field label, not an entire paragraph or container
                if (el.children.length > 4) continue;
                const txt = clean(el.innerText || el.textContent);
                if (!txt || txt.length > 70 || !re.test(txt)) continue;

                // 1A. Associated by HTML 'for' attribute: <label for="fieldId">
                const forId = el.getAttribute("for") || el.getAttribute("htmlFor");
                if (forId) {
                    const target = document.getElementById(forId);
                    if (target) {
                        const v = getElementValue(target);
                        if (v && !re.test(v)) return v;
                    }
                }

                // 1B. Input nested directly inside label
                const nestedInput = el.querySelector("input:not([type=hidden]):not([type=submit]), select, textarea");
                if (nestedInput) {
                    const v = getElementValue(nestedInput);
                    if (v && !re.test(v)) return v;
                }

                // 1C. Container search: walk up to form-group, col, row, or parent container
                let parent = el.parentElement;
                for (let depth = 0; depth < 3 && parent; depth++) {
                    const inputsInContainer = parent.querySelectorAll("input:not([type=hidden]):not([type=submit]):not([type=button]), select, textarea");
                    for (const inp of inputsInContainer) {
                        const v = getElementValue(inp);
                        if (v && !re.test(v)) return v;
                    }
                    parent = parent.parentElement;
                }

                // 1D. Following sibling element (e.g. input or wrapper div right after label)
                let sib = el.nextElementSibling;
                while (sib) {
                    const sibInput = (sib.matches && sib.matches("input, select, textarea"))
                        ? sib
                        : sib.querySelector ? sib.querySelector("input:not([type=hidden]):not([type=submit]), select, textarea") : null;
                    if (sibInput) {
                        const v = getElementValue(sibInput);
                        if (v && !re.test(v)) return v;
                    }

                    const sibText = clean(sib.innerText || sib.textContent);
                    if (sibText && !re.test(sibText) && sibText.length < 150) {
                        return sibText;
                    }
                    sib = sib.nextElementSibling;
                }

                // 1E. Table cell layout: <tr><td>Label</td><td>Value</td></tr>
                const tr = el.closest("tr");
                if (tr) {
                    const cells = Array.from(tr.querySelectorAll("td, th"));
                    const idx = cells.findIndex(c => re.test(clean(c.innerText || c.textContent)));
                    if (idx !== -1 && idx + 1 < cells.length) {
                        const targetCell = cells[idx + 1];
                        const cellInput = targetCell.querySelector("input:not([type=hidden]):not([type=submit]), select, textarea");
                        if (cellInput) {
                            const v = getElementValue(cellInput);
                            if (v && !re.test(v)) return v;
                        }
                        const v = clean(targetCell.innerText || targetCell.textContent);
                        if (v && !re.test(v)) return v;
                    }
                }
            }

            // --- STRATEGY 2: Direct Input Attribute Inspection ---
            // Checks input id, name, placeholder, aria-label, title, ng-model, formcontrolname
            const allInputs = document.querySelectorAll("input:not([type=hidden]):not([type=submit]), select, textarea");
            for (const inp of allInputs) {
                const attrs = [
                    inp.id,
                    inp.name,
                    inp.placeholder,
                    inp.getAttribute("aria-label"),
                    inp.getAttribute("title"),
                    inp.getAttribute("data-field"),
                    inp.getAttribute("data-id"),
                    inp.getAttribute("ng-model"),
                    inp.getAttribute("formcontrolname")
                ].filter(Boolean).join(" ");

                if (re.test(attrs)) {
                    const v = getElementValue(inp);
                    if (v && !re.test(v)) return v;
                }
            }
        }
        return "";
    }

    // Comprehensive Portal Claim Extractor
    function extractClaimData() {
        const fullText = document.body ? (document.body.innerText || "") : "";
        const host = window.location.hostname.toLowerCase();

        // 1. Auto-detect Insurer across all Indian insurance companies
        let insCo = "SOMPO";
        if (/sompo|universal/i.test(host) || /universal\s*sompo/i.test(fullText)) {
            insCo = "SOMPO";
        } else if (/hdfc|ergo/i.test(host) || /hdfc\s*ergo/i.test(fullText)) {
            insCo = "HDFC";
        } else if (/bajaj|allianz/i.test(host) || /bajaj\s*allianz/i.test(fullText)) {
            insCo = "BAJAJ";
        } else if (/icici|lombard/i.test(host) || /icici\s*lombard/i.test(fullText)) {
            insCo = "ICICI";
        } else if (/tata|aig/i.test(host) || /tata\s*aig/i.test(fullText)) {
            insCo = "TATA AIG";
        } else if (/new\s*india|nia\.co/i.test(host) || /new\s*india/i.test(fullText)) {
            insCo = "NIA";
        } else if (/chola|cholamandalam/i.test(host) || /chola/i.test(fullText)) {
            insCo = "CHOLA";
        } else if (/liberty/i.test(host) || /liberty/i.test(fullText)) {
            insCo = "LIBERTY";
        } else if (/digit|godigit/i.test(host) || /go\s*digit/i.test(fullText)) {
            insCo = "DIGIT";
        } else if (/acko/i.test(host) || /acko/i.test(fullText)) {
            insCo = "ACKO";
        } else if (/national|nic\.co/i.test(host) || /national\s*insurance/i.test(fullText)) {
            insCo = "NIC";
        } else if (/oriental|oic/i.test(host) || /oriental\s*insurance/i.test(fullText)) {
            insCo = "OIC";
        } else if (/united|uiic/i.test(host) || /united\s*india/i.test(fullText)) {
            insCo = "UIIC";
        } else if (/reliance/i.test(host) || /reliance/i.test(fullText)) {
            insCo = "RELIANCE";
        } else if (/iffco|tokio/i.test(host) || /iffco\s*tokio/i.test(fullText)) {
            insCo = "IFFCO";
        } else if (/sbi/i.test(host) || /sbi\s*general/i.test(fullText)) {
            insCo = "SBI";
        }

        // Today's date formatted e.g. "25-OCT-2026"
        const todayDate = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }).replace(/ /g, '-').toUpperCase();

        // 2. Extract Data using Synonym Matrix
        const data = {
            ins_co: insCo,
            policy_no: findValueByAnchors([
                "Policy No",
                "Policy Number",
                "Policy Ref",
                "Policy\\s*#",
                "Policy\\s*No\\.",
                "Certificate Number",
                "Certificate No",
                "Policy / Certificate",
                "Policy/Covernote",
                "Covernote No",
                "PolicyID",
                "PolicyNum"
            ]),
            claim_no: findValueByAnchors([
                "Insurer Claim Number",
                "Claim Number",
                "Claim No",
                "Claim Ref",
                "Claim Reference",
                "Intimation Number",
                "Intimation No",
                "Child Claim Number",
                "Claim\\s*#",
                "Claim\\s*No\\.",
                "ClaimId",
                "ClaimNum"
            ]),
            insured_name: findValueByAnchors([
                "Insured Name",
                "Name of Insured",
                "Customer Name",
                "Policyholder Name",
                "Proposer Name",
                "Client Name",
                "Name of the Insured",
                "Insured's Name",
                "Insured Person",
                "Owner Name"
            ]),
            date_of_loss: cleanDate(findValueByAnchors([
                "Accident Date Time",
                "Accident Date",
                "Date of Loss",
                "Loss Date",
                "Accident Date & Time",
                "Date & Time of Accident",
                "Date & Time of Loss",
                "Date of Accident",
                "Time & Date of Loss",
                "Date/Time of Loss",
                "Loss Date / Time"
            ])),
            policy_start_date: cleanDate(findValueByAnchors([
                "Policy Start Date",
                "Policy Inception Date",
                "Period of Insurance From",
                "From Date",
                "Inception Date",
                "Cover From",
                "Policy From",
                "Start Date"
            ])),
            vehicle_no: findValueByAnchors([
                "Registration Number",
                "Reg No",
                "Reg. No",
                "Regn No",
                "Regn. No",
                "Regn Number",
                "Vehicle Registration",
                "Vehicle No",
                "Vehicle Number",
                "Veh Reg No",
                "Veh No",
                "Registration No",
                "Registration\\s*#",
                "Asset Number"
            ]),
            make_model: "",
            workshop: findValueByAnchors([
                "Garage Name",
                "Workshop Name",
                "Dealer Name",
                "Repairer Name",
                "Inspection Address",
                "Garage / Workshop",
                "Workshop",
                "Dealer",
                "Service Center",
                "Garage Location",
                "Workshop Location"
            ]),
            survey_done_by: findValueByAnchors([
                "Survey Done By",
                "Surveyor Name",
                "Deputed Surveyor",
                "Surveyor",
                "Allocated Surveyor",
                "Assigned Surveyor"
            ]),
            advisor_name: findValueByAnchors([
                "Advisor Name",
                "Service Advisor",
                "Claim Advisor",
                "Workshop Advisor"
            ]),
            idv: findValueByAnchors([
                "^IDV$",
                "Insured Declared Value",
                "Vehicle IDV",
                "Total IDV"
            ]),
            ncb: findValueByAnchors([
                "No Claim Bonus",
                "NCB",
                "NCB %",
                "NCB Discount"
            ]) || "0",
            driver_relation: findValueByAnchors([
                "Driver Relation with Insured",
                "Driver Relation",
                "Relationship with Insured"
            ]) || "Self",
            date_of_survey: cleanDate(findValueByAnchors([
                "Date of Survey",
                "Survey Date",
                "Inspection Date"
            ])) || todayDate,
            source_url: window.location.href,
            extracted_at: new Date().toLocaleTimeString()
        };

        // 3. Make & Model Resolver: Handles combined and separate Make / Model fields
        let combinedMakeModel = findValueByAnchors([
            "Make & Model",
            "Vehicle Make & Model",
            "Make / Model",
            "Make Model",
            "Vehicle Details"
        ]);

        if (combinedMakeModel) {
            data.make_model = combinedMakeModel;
        } else {
            const makeVal = findValueByAnchors(["^Make$", "Vehicle Make", "Manufacturer"]);
            const modelVal = findValueByAnchors(["^Model$", "Vehicle Model", "Model Name", "Variant"]);
            if (makeVal && modelVal) {
                data.make_model = `${makeVal} ${modelVal}`.trim();
            } else {
                data.make_model = makeVal || modelVal || "";
            }
        }

        // 4. Universal Pattern Heuristics: Fallback scan over all inputs for distinct Indian vehicle & policy patterns
        if (!data.vehicle_no) {
            const allInputs = document.querySelectorAll("input:not([type=hidden]):not([type=submit])");
            // Standard Indian Vehicle Reg pattern e.g. PB-10-JV-1202, MH02CD1234, DL-1C-AA-1111
            const vehPattern = /\b([A-Z]{2}[-\s]?[0-9]{1,2}[-\s]?[A-Z]{0,3}[-\s]?[0-9]{4})\b/i;
            for (const inp of allInputs) {
                const val = clean(inp.value);
                const m = val.match(vehPattern);
                if (m && !/^(OD|TP|IDV|PA|MDL)$/i.test(m[1])) {
                    data.vehicle_no = m[1].toUpperCase();
                    break;
                }
            }
        }

        // 5. Cache and Broadcast to Top Frame and Storage
        if (data.claim_no || data.vehicle_no || data.policy_no) {
            window.__cachedClaimData = data;
            try {
                if (chrome && chrome.storage && chrome.storage.local) {
                    chrome.storage.local.set({ last_extracted_claim: data, extracted_at: Date.now() });
                }
            } catch (e) {}

            try {
                window.top.postMessage({ type: "OFFICE_MIS_CLAIM_DATA", data: data }, "*");
            } catch (e) {}
        }

        return data;
    }

    // Listen for cross-frame messages from subframes (iframes)
    window.addEventListener("message", (event) => {
        if (event.data && event.data.type === "OFFICE_MIS_CLAIM_DATA" && event.data.data) {
            window.__cachedClaimData = event.data.data;
            try {
                if (chrome && chrome.storage && chrome.storage.local) {
                    chrome.storage.local.set({ last_extracted_claim: event.data.data, extracted_at: Date.now() });
                }
            } catch (e) {}
        }
    });

    // Floating Button UI
    function createFloatingBadge() {
        if (document.getElementById("office-mis-floating-widget")) return;

        const widget = document.createElement("div");
        widget.id = "office-mis-floating-widget";
        widget.style.cssText = `
            position: fixed;
            bottom: 24px;
            right: 24px;
            z-index: 2147483647;
            background: linear-gradient(135deg, #2563eb, #1e40af);
            color: #ffffff;
            padding: 11px 20px;
            border-radius: 30px;
            box-shadow: 0 10px 25px -5px rgba(37, 99, 235, 0.4), 0 8px 10px -6px rgba(0, 0, 0, 0.1);
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            font-size: 13.5px;
            font-weight: 600;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 8px;
            transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1);
            border: 1px solid rgba(255, 255, 255, 0.3);
            user-select: none;
        `;
        widget.innerHTML = `<span>📋</span> <span>Send Claim to MIS</span>`;

        widget.onmouseenter = () => {
            widget.style.transform = "translateY(-2px) scale(1.02)";
            widget.style.boxShadow = "0 14px 28px -5px rgba(37, 99, 235, 0.5), 0 10px 10px -5px rgba(0, 0, 0, 0.15)";
        };
        widget.onmouseleave = () => {
            widget.style.transform = "translateY(0) scale(1)";
            widget.style.boxShadow = "0 10px 25px -5px rgba(37, 99, 235, 0.4), 0 8px 10px -6px rgba(0, 0, 0, 0.1)";
        };

        widget.onclick = async () => {
            const originalHTML = widget.innerHTML;
            widget.innerHTML = `<span>⏳</span> <span>Extracting...</span>`;
            widget.style.pointerEvents = "none";

            const claimData = window.__cachedClaimData || extractClaimData();

            let resData = null;
            try {
                let res;
                try {
                    res = await fetch("http://127.0.0.1:5055/api/web-claim", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ record: claimData })
                    });
                } catch (e1) {
                    res = await fetch("http://localhost:5055/api/web-claim", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ record: claimData })
                    });
                }
                resData = await res.json();
            } catch (err) {
                alert("Cannot connect to Office MIS Desktop App. Please make sure OfficeAutomation.exe is running on this PC.");
                widget.innerHTML = originalHTML;
                widget.style.pointerEvents = "auto";
                return;
            }

            if (resData && resData.success) {
                widget.style.background = "linear-gradient(135deg, #10b981, #047857)";
                widget.innerHTML = `<span>✅</span> <span>Sent to MIS (#${resData.ref_no || 'Saved'})!</span>`;
                setTimeout(() => {
                    widget.style.background = "linear-gradient(135deg, #2563eb, #1e40af)";
                    widget.innerHTML = originalHTML;
                    widget.style.pointerEvents = "auto";
                }, 3500);
            } else {
                alert("Office MIS App: " + (resData ? resData.error : "Could not save claim."));
                widget.innerHTML = originalHTML;
                widget.style.pointerEvents = "auto";
            }
        };

        document.body.appendChild(widget);
    }

    // Listen for messages from extension popup
    chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
        if (msg.action === "get_claim_data") {
            const data = window.__cachedClaimData || extractClaimData();
            sendResponse({ success: true, data: data });
        }
        return true;
    });

    // Check if page looks like an insurance claim page
    function initScanner() {
        // Run extraction once on load
        extractClaimData();

        // Only show floating badge on top frame or if in an iframe with actual claim content
        const text = (document.body ? document.body.innerText || "" : "").toLowerCase();
        const isInsurancePage = text.includes("claim") || text.includes("policy") || text.includes("sompo") ||
                                text.includes("hdfc") || text.includes("bajaj") || text.includes("icici") ||
                                text.includes("survey") || text.includes("intimation") || window.location.pathname.includes("portal");

        if (isInsurancePage) {
            // If top window, always show. If in iframe, only show if top window doesn't have it
            if (window.self === window.top) {
                createFloatingBadge();
            }
        }

        // Periodic light scan for dynamic/single-page apps (SPAs)
        setInterval(() => {
            extractClaimData();
        }, 3000);
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", initScanner);
    } else {
        initScanner();
    }
})();
