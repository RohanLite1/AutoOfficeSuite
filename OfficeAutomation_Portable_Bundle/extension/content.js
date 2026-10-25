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

    // Helper: Extract clean date portion (returns empty string if string is not a date)
    function cleanDate(str) {
        if (!str) return "";
        const m = String(str).match(/\b(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\b/);
        return m ? m[1] : "";
    }

    // Helper: Validate if a string is a real calendar date (rejects status words like "Partial Loss")
    function isValidDate(val) {
        if (!val) return false;
        const v = clean(val);
        if (v.length < 6) return false;
        if (/^(survey|loss|partial|total|pending|approved|rejected|closed|report|submitted|draft|non rapid)/i.test(v)) return false;
        return /\b\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}\b/.test(v) ||
               /\b\d{1,2}-[A-Za-z]{3}-\d{2,4}\b/i.test(v);
    }

    // Helper: Clean & extract Indian phone / mobile number (10 digits)
    function cleanPhone(raw) {
        if (!raw) return "";
        const str = clean(raw);
        if (/^[\-\.\sNAna\/None]+$/i.test(str)) return "";
        // Match 10 digit Indian mobile number (optionally prefixed with +91 or 0)
        const m10 = str.match(/(?:(?:\+|0{0,2})91[\s\-]*)?([6-9]\d{9})\b/);
        if (m10) return m10[1];
        // General digit extraction
        const digits = str.replace(/[^\d]/g, "");
        if (digits.length === 10) return digits;
        if (digits.length === 11 && digits.startsWith("0")) return digits.substring(1);
        if (digits.length === 12 && digits.startsWith("91")) return digits.substring(2);
        if (digits.length >= 7 && digits.length <= 13) return digits;
        return "";
    }

    // Helper: Detect if extracted text is actually a field header / form label rather than a real value
    function isFormLabelText(txt) {
        if (!txt) return false;
        const t = clean(txt);
        if (t.length < 2) return false;
        if (t.endsWith(":")) return true;
        if (/^(no\s*objection.*|noc.*|cubic\s*capacity.*|color\s*as\s*per\s*rc.*|repairer\s*city.*|repairer\s*name.*|repairer.*|intimation\s*date.*|date\s*of\s*notification.*|time\s*of\s*notification.*|time\s*of\s*loss.*|policy\s*start\s*time.*|od\s*policy.*time.*|loss\s*description.*|place\s*of\s*loss.*|claim\s*status.*|chassis\s*no.*|engine\s*no.*|nature\s*of\s*loss.*|type\s*of\s*loss.*|product\s*code.*|police\s*station.*|fir\/gdr\s*no.*|inspection\s*address.*|garage\s*name.*|trailer\s*opted.*|purpose\s*of\s*using.*|license\s*issuing.*|age\s*of\s*driver.*|relationship\s*of\s*the\s*driver.*|reported\s*branch.*|claim\s*reported\s*by.*|insured\s*interest\s*code.*|type\s*of\s*vehicle\s*insured.*|registered\s*no.*|view\s*claim\s*details|policy\s*details|risk\s*details|estimate\s*details|assessment\s*details)$/i.test(t)) return true;
        return false;
    }

    // Helper: Normalize clean IDV (rejects non-numeric labels like "No Objection Certificate")
    function cleanIdv(val) {
        if (!val) return "";
        const str = clean(val);
        if (/^(nil|na|n\/a|none|nil\s*depreciation|no\s*objection.*|noc.*|0|0\.00)$/i.test(str)) return "";
        if (!/\d/.test(str)) return ""; // Must contain at least one digit
        if (/^(no\s*objection|noc|nil|na|partial|repairer)/i.test(str)) return "";
        const m = str.match(/(?:₹|Rs\.?)?\s*([0-9,]+(?:\.\d{2})?)/);
        if (!m) return "";
        const num = m[1].replace(/,/g, "");
        if (num === "0" || num === "0.00") return "";
        return m[1];
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

    // Helper: Validate extracted Policy Number (must have digits, cannot be claim status like "Survey Done")
    function isValidPolicyNo(val, claimNo) {
        if (!val) return false;
        const v = clean(val);
        if (v.length < 5) return false;
        if (!/\d/.test(v)) return false; // Must contain at least one digit
        // Disallow status strings, driver relation, or non-policy keywords
        if (/^(survey done|fla done|done|pending|approved|rejected|closed|in progress|initial technical|under process|open|draft|completed|cancelled|issued|self|spouse|driver|father|mother|brother|son|daughter|nil|na|n\/a|none|null|undefined)$/i.test(v)) return false;
        // Disallow pure dates
        if (/^\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}$/.test(v)) return false;
        // Disallow claim number
        if (claimNo && v.toLowerCase() === clean(claimNo).toLowerCase()) return false;
        return true;
    }

    // Helper: Normalize clean NCB percentage (only valid percentage between 0 and 65, defaults to "0")
    function cleanNcb(rawVal) {
        if (!rawVal) return "0";
        const str = clean(rawVal);
        // Match percentage like "20%" or "25 %"
        const pctMatch = str.match(/\b([0-9]{1,2})\s*%/);
        if (pctMatch) {
            const num = parseInt(pctMatch[1], 10);
            if (num <= 65) return String(num);
        }
        // Match standard NCB numbers (0, 20, 25, 35, 45, 50, 65)
        const stdMatch = str.match(/\b(0|20|25|35|45|50|65)\b/);
        if (stdMatch) {
            return stdMatch[1];
        }
        // Match any 1-2 digit number up to 65
        const anyMatch = str.match(/^\s*([0-9]{1,2})\s*$/);
        if (anyMatch) {
            const num = parseInt(anyMatch[1], 10);
            if (num <= 65) return String(num);
        }
        // Reject garbage like "321", "311", status strings, or internal IDs
        return "0";
    }

    // Helper: Validate Vehicle Make & Model (must not be claim number like CL26404587)
    function isValidMakeModel(val, claimNo) {
        if (!val) return false;
        const v = clean(val);
        if (v.length < 2) return false;
        if (claimNo && v.toLowerCase() === clean(claimNo).toLowerCase()) return false;
        if (/^CL\d{5,}/i.test(v)) return false;
        if (/^(survey done|fla done|pending|approved|rejected|closed|self|spouse)$/i.test(v)) return false;
        return true;
    }

    // Universal DOM Search: Finds value across any portal architecture (Tables, Bootstrap grids, Angular, React, ASP.NET)
    function findValueByAnchors(anchors) {
        for (const anchor of anchors) {
            let re;
            try {
                if (typeof anchor === "string") {
                    if (/^[a-zA-Z0-9\s\.\/%#-]+$/.test(anchor)) {
                        const escaped = anchor.trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
                        const prefix = /^\w/.test(anchor.trim()) ? "\\b" : "";
                        const suffix = /\w$/.test(anchor.trim()) ? "\\b" : "";
                        re = new RegExp(prefix + escaped + suffix, "i");
                    } else {
                        re = new RegExp(anchor, "i");
                    }
                } else {
                    re = new RegExp(anchor, "i");
                }
            } catch (e) {
                re = new RegExp(anchor.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), "i");
            }

            // --- STRATEGY 1: Search Form Labels & Field Headers ---
            const labelCandidates = document.querySelectorAll(
                "label, .control-label, .form-label, .field-label, [class*='label'], [class*='title'], [class*='header'], th, dt, b, strong, span, div, p"
            );

            for (const el of labelCandidates) {
                // Must not be a dropdown option, button, or script/style
                if (el.closest("option, select, button, script, style, [role='option']")) continue;
                if (/^(select|input|textarea|button|option)$/i.test(el.tagName)) continue;

                // Must be short enough to be a field label, not an entire container
                if (el.children.length > 4) continue;
                const txt = clean(el.innerText || el.textContent);
                if (!txt || txt.length > 70 || !re.test(txt)) continue;

                // 1A. Associated by HTML 'for' attribute: <label for="fieldId">
                const forId = el.getAttribute("for") || el.getAttribute("htmlFor");
                if (forId) {
                    const target = document.getElementById(forId);
                    if (target) {
                        const v = getElementValue(target);
                        if (v && !re.test(v) && !isFormLabelText(v)) return v;
                    }
                }

                // 1B. Input nested directly inside label
                const nestedInput = el.querySelector("input:not([type=hidden]):not([type=submit]), select, textarea");
                if (nestedInput) {
                    const v = getElementValue(nestedInput);
                    if (v && !re.test(v) && !isFormLabelText(v)) return v;
                }

                // 1C. Following sibling element (e.g. input or wrapper div right after label)
                let sib = el.nextElementSibling;
                while (sib) {
                    const sibInput = (sib.matches && sib.matches("input:not([type=hidden]):not([type=submit]), select, textarea"))
                        ? sib
                        : sib.querySelector ? sib.querySelector("input:not([type=hidden]):not([type=submit]), select, textarea") : null;
                    if (sibInput) {
                        const v = getElementValue(sibInput);
                        if (v && !re.test(v) && !isFormLabelText(v)) return v;
                    }

                    const sibText = clean(sib.innerText || sib.textContent);
                    if (sibText && !re.test(sibText) && !isFormLabelText(sibText) && sibText.length < 150) {
                        return sibText;
                    }
                    sib = sib.nextElementSibling;
                }

                // 1D. Immediate container input (only depth 0: immediate form-group or column)
                if (el.parentElement) {
                    const immediateInputs = el.parentElement.querySelectorAll("input:not([type=hidden]):not([type=submit]):not([type=button]), select, textarea");
                    for (const inp of immediateInputs) {
                        const v = getElementValue(inp);
                        if (v && !re.test(v) && !isFormLabelText(v)) return v;
                    }
                    // Check sibling column in grid (e.g. <div class="col-sm-4"><label>...</label></div><div class="col-sm-8"><input></div>)
                    const parentSib = el.parentElement.nextElementSibling;
                    if (parentSib) {
                        const sibInp = parentSib.matches && parentSib.matches("input:not([type=hidden]):not([type=submit]), select, textarea")
                            ? parentSib
                            : parentSib.querySelector ? parentSib.querySelector("input:not([type=hidden]):not([type=submit]), select, textarea") : null;
                        if (sibInp) {
                            const v = getElementValue(sibInp);
                            if (v && !re.test(v) && !isFormLabelText(v)) return v;
                        }
                    }
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
                            if (v && !re.test(v) && !isFormLabelText(v)) return v;
                        }
                        const v = clean(targetCell.innerText || targetCell.textContent);
                        if (v && !re.test(v) && !isFormLabelText(v)) return v;
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
                "\\bPolicy\\s*No\\b",
                "\\bPolicy\\s*Number\\b",
                "\\bPolicy\\s*Ref\\b",
                "\\bPolicy\\s*#\\b",
                "\\bPolicy\\s*No\\.",
                "\\bCertificate\\s*Number\\b",
                "\\bCertificate\\s*No\\b",
                "\\bPolicy\\s*/\\s*Certificate\\b",
                "\\bPolicy\\s*/\\s*Covernote\\b",
                "\\bCovernote\\s*No\\b",
                "\\bPolicyID\\b",
                "\\bPolicyNum\\b"
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
            phone: cleanPhone(findValueByAnchors([
                "Insured Contact No",
                "Insured Contact Number",
                "Contact Number",
                "Contact No",
                "Contact No.",
                "Insured Mobile",
                "Insured Mobile No",
                "Mobile Number",
                "Mobile No",
                "Mobile No.",
                "Customer Mobile",
                "Customer Phone",
                "Caller Mobile",
                "Insured Phone",
                "Phone Number",
                "Phone No",
                "Phone No.",
                "Mobile",
                "Phone"
            ])),
            ncb: "",
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

        // --- SPECIFIC FIELD VALIDATION & PORTAL HEURISTICS ---

        // 2A. Policy Number Validation: Cannot be status keywords like "Survey Done"
        if (!isValidPolicyNo(data.policy_no, data.claim_no)) {
            data.policy_no = "";
        }

        // Dedicated Sompo / Portal Policy No fallback search
        if (!data.policy_no) {
            // First: Look for input directly associated with label "Policy No"
            const labels = document.querySelectorAll("label, th, td, b, strong, span, div");
            for (const l of labels) {
                if (l.closest("option, select, button, script, style")) continue;
                const txt = clean(l.innerText || l.textContent);
                if (/^Policy\s*No\.?$/i.test(txt)) {
                    // Check sibling input or wrapper input
                    const inp = l.parentElement ? l.parentElement.querySelector("input:not([type=hidden]):not([type=submit])") : null;
                    if (inp && isValidPolicyNo(inp.value, data.claim_no)) {
                        data.policy_no = clean(inp.value);
                        break;
                    }
                    const nextInp = l.nextElementSibling && l.nextElementSibling.tagName === "INPUT"
                        ? l.nextElementSibling
                        : (l.nextElementSibling ? l.nextElementSibling.querySelector("input:not([type=hidden]):not([type=submit])") : null);
                    if (nextInp && isValidPolicyNo(nextInp.value, data.claim_no)) {
                        data.policy_no = clean(nextInp.value);
                        break;
                    }
                    const parentSib = l.parentElement && l.parentElement.nextElementSibling
                        ? l.parentElement.nextElementSibling.querySelector("input:not([type=hidden]):not([type=submit])") : null;
                    if (parentSib && isValidPolicyNo(parentSib.value, data.claim_no)) {
                        data.policy_no = clean(parentSib.value);
                        break;
                    }
                }
            }
        }

        // Second: Scan all inputs for Sompo policy number regex: e.g. 2367/80826795/50/000 or 2371/80787797/00/000
        if (!data.policy_no) {
            const sompoPattern = /\b(23\d{2}\/\d{8}\/[A-Z0-9]{2}\/[A-Z0-9]{3})\b/i;
            for (const inp of document.querySelectorAll("input:not([type=hidden]):not([type=submit])")) {
                const val = clean(inp.value);
                const m = val.match(sompoPattern);
                if (m) {
                    data.policy_no = m[1];
                    break;
                }
            }
        }

        // Third: Scan inputs for general policy number pattern
        if (!data.policy_no) {
            for (const inp of document.querySelectorAll("input:not([type=hidden]):not([type=submit])")) {
                const val = clean(inp.value);
                if (isValidPolicyNo(val, data.claim_no) && /^(OG|POL|\d{4}|\d{5}|[A-Z0-9\/-]{10,})/i.test(val)) {
                    // Make sure it's not chassis/engine or vehicle number
                    if (val !== data.vehicle_no && !/\b[A-Z]{2}[-\s]?\d{1,2}[-\s]?[A-Z]{0,3}[-\s]?\d{4}\b/i.test(val)) {
                        data.policy_no = val;
                        break;
                    }
                }
            }
        }

        // 2B. NCB Percentage Resolver: Only clean percentage if given on website, else "0"
        const rawNcb = findValueByAnchors([
            "\\bNo\\s*Claim\\s*Bonus\\b",
            "\\bNCB\\s*%\\b",
            "\\bNCB\\s*Discount\\b",
            "\\bNCB\\b"
        ]);
        data.ncb = cleanNcb(rawNcb);

        // If not found in inputs, check full webpage text for "NCB : 20%" or "No Claim Bonus: 25%"
        if (data.ncb === "0" || !data.ncb) {
            const ncbTextMatch = fullText.match(/\b(?:No\s*Claim\s*Bonus|NCB)\s*[:\-]?\s*([0-9]{1,2})\s*%/i);
            if (ncbTextMatch) {
                data.ncb = cleanNcb(ncbTextMatch[1]);
            }
        }

        // 3. Make & Model Resolver: Handles combined and separate Make / Model fields
        let combinedMakeModel = findValueByAnchors([
            "Vehicle Make & Model",
            "Make & Model",
            "Make / Model",
            "Make Model"
        ]); // NOTE: "Vehicle Details" removed as it captures claim container

        if (combinedMakeModel && isValidMakeModel(combinedMakeModel, data.claim_no)) {
            data.make_model = combinedMakeModel;
        } else {
            const makeVal = findValueByAnchors(["Vehicle Make", "^Make$", "Manufacturer"]);
            const modelVal = findValueByAnchors(["Vehicle Model", "^Model$", "Model Name", "Variant"]);
            if (makeVal && modelVal && isValidMakeModel(`${makeVal} ${modelVal}`, data.claim_no)) {
                data.make_model = `${makeVal} ${modelVal}`.trim();
            } else if (makeVal && isValidMakeModel(makeVal, data.claim_no)) {
                data.make_model = makeVal;
            } else if (modelVal && isValidMakeModel(modelVal, data.claim_no)) {
                data.make_model = modelVal;
            }
        }

        // Targeted Make & Model search for Universal Sompo portal
        if (!data.make_model || !isValidMakeModel(data.make_model, data.claim_no)) {
            let mk = "";
            let mdl = "";
            for (const l of document.querySelectorAll("label, th, td, b, strong, span")) {
                if (l.closest("option, select, button, script, style")) continue;
                const txt = clean(l.innerText || l.textContent);
                if (/^Make\*?$/i.test(txt)) {
                    const inp = l.parentElement ? l.parentElement.querySelector("input:not([type=hidden]), select") : null;
                    if (inp) mk = getElementValue(inp);
                }
                if (/^Model\*?$/i.test(txt)) {
                    const inp = l.parentElement ? l.parentElement.querySelector("input:not([type=hidden]), select") : null;
                    if (inp) mdl = getElementValue(inp);
                }
            }
            if (mk || mdl) {
                const combined = `${mk} ${mdl}`.trim();
                if (isValidMakeModel(combined, data.claim_no)) {
                    data.make_model = combined;
                }
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

        // 5. Dedicated United India Insurance (UIIC) Extractor & Heuristics
        const isUiic = (insCo === "UIIC") || /uiic|united\s*india/i.test(host) || /united\s*india/i.test(fullText);
        if (isUiic) {
            data.ins_co = "UIIC";

            // 5A. UIIC Claim No: e.g. 2006023126C050106001
            const uiicClaimMatch = fullText.match(/\bClaim\s*No\.?\s*:\s*([A-Z0-9]+)\b/i);
            if (uiicClaimMatch && uiicClaimMatch[1].length >= 10) {
                data.claim_no = clean(uiicClaimMatch[1]);
            }

            // 5B. UIIC Policy No: e.g. 2006023126P106023411
            const uiicPolicyMatch = fullText.match(/\bPolicy\s*No\.?\s*:\s*([A-Z0-9]+)\b/i);
            if (uiicPolicyMatch && uiicPolicyMatch[1].length >= 10) {
                data.policy_no = clean(uiicPolicyMatch[1]);
            }

            // 5C. UIIC Insured Name: must NOT capture next label (Intimation Date)
            const uiicNameMatch = fullText.match(/\bInsured\s*Name\s*:\s*([^:\r\n]+?)(?=\s*(?:Intimation|Contact|Email|Loss|Estimated|Assessed|Claim|Vehicle|Engine|Chassis|Workshop|$))/i);
            if (uiicNameMatch && clean(uiicNameMatch[1])) {
                const parsedName = clean(uiicNameMatch[1]);
                if (!/intimation/i.test(parsedName)) {
                    data.insured_name = parsedName;
                }
            }

            // 5D. UIIC Insured Contact / Phone No: e.g. 9872072656
            const uiicContactMatch = fullText.match(/\bInsured\s*Contact\s*No\.?\s*:\s*([0-9\+\-\s]+)/i);
            if (uiicContactMatch) {
                const p = cleanPhone(uiicContactMatch[1]);
                if (p) data.phone = p;
            }

            // 5E. UIIC Loss Date: format 02/10/2026, must NOT capture Estimated Amount
            const uiicLossMatch = fullText.match(/\bLoss\s*Date\s*:\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\b/i);
            if (uiicLossMatch) {
                data.date_of_loss = uiicLossMatch[1];
            }

            // 5F. UIIC Vehicle Registration: e.g. PB13BP9108
            const uiicVehMatch = fullText.match(/\bVehicle\s*Registration\s*Number\s*:\s*([A-Z0-9\-]+)\b/i);
            if (uiicVehMatch) {
                data.vehicle_no = uiicVehMatch[1].toUpperCase();
            }

            // 5G. UIIC Workshop: e.g. DUMMY, must NOT capture "Workshop Phone No : -"
            const uiicWsMatch = fullText.match(/\bWorkshop\s*Name\s*:\s*([^:\r\n]+?)(?=\s*(?:Workshop\s*Phone|Workshop\s*Address|Claim\s*Type|Engine|Chassis|$))/i);
            if (uiicWsMatch && clean(uiicWsMatch[1])) {
                data.workshop = clean(uiicWsMatch[1]);
            }

            // 5H. UIIC IDV: User specified: "in UIIC specifically, leave the IDV empty as the assessed amount given in the web report is zero."
            const uiicAssessedMatch = fullText.match(/\bAssessed\s*Amount(?:\s*\(₹\))?\s*:\s*([0-9,\.]*)/i);
            const assessedVal = uiicAssessedMatch ? clean(uiicAssessedMatch[1]) : "";
            if (!assessedVal || assessedVal === "0" || assessedVal === "0.00") {
                data.idv = ""; // Explicitly leave empty as requested
            } else {
                data.idv = assessedVal;
            }

            // 5I. UIIC Date of Survey: defaults to today, never "Partial Loss"
            data.date_of_survey = todayDate;
        }

        // 6. Dedicated The New India Assurance (NIA) Extractor & Heuristics
        const isNia = (insCo === "NIA") || /newindia|web\.newindia\.co\.in|niabancsportal/i.test(host) || /new\s*india/i.test(fullText);
        if (isNia) {
            data.ins_co = "NIA";

            // 6A. NIA Claim No: 20-digit claim number e.g. 36100331260390000180
            const niaClaimMatch = fullText.match(/\bClaim\s*Number\s*:\s*([0-9]{15,25})\b/i) || fullText.match(/\bClaim\s*No\.?\s*:\s*([0-9]{15,25})\b/i);
            if (niaClaimMatch) {
                data.claim_no = clean(niaClaimMatch[1]);
            }

            // 6B. NIA Policy No: e.g. 36100331260300000844
            const niaPolicyMatch = fullText.match(/\bPolicy\s*Number\s*:\s*([0-9]{15,25})\b/i) || fullText.match(/\bPolicy\s*No\.?\s*:\s*([0-9]{15,25})\b/i);
            if (niaPolicyMatch) {
                data.policy_no = clean(niaPolicyMatch[1]);
            }

            // 6C. NIA Insured Name: e.g. M/S GOGNA TRANSPORT SERVICES PVT LTD
            // Must NOT capture "Intimation Date" or "Date of Notification"
            const niaNameMatch = fullText.match(/\bInsured\s*Name\s*:\s*([^:\r\n]+?)(?=\s*(?:Registered\s*No|Make|Model|Chassis|Engine|Relationship|Driver|Loss|Policy|Claim|$))/i);
            if (niaNameMatch && clean(niaNameMatch[1])) {
                const parsedName = clean(niaNameMatch[1]);
                if (!/intimation|notification/i.test(parsedName) && !isFormLabelText(parsedName)) {
                    data.insured_name = parsedName;
                }
            }

            // 6D. NIA Date of Loss: format 25/07/2026, must NOT capture "Time of Loss/Accident: 00:00:00"
            const niaLossMatch = fullText.match(/\bDate\s*of\s*Loss(?:\/Accident)?\s*:\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\b/i);
            if (niaLossMatch) {
                data.date_of_loss = niaLossMatch[1];
            }

            // 6E. NIA Vehicle Reg No: e.g. PB-09-AQ-8169, must NOT capture "Cubic Capacity"
            const niaVehMatch = fullText.match(/\bRegistered\s*No\.?\s*:\s*([A-Z0-9\-]+)\b/i);
            if (niaVehMatch && !isFormLabelText(niaVehMatch[1])) {
                data.vehicle_no = clean(niaVehMatch[1]).toUpperCase();
            }

            // 6F. NIA Make & Model: e.g. Make: ASHOK LEYL, Model: 1212 -> ASHOK LEYL 1212
            // Must NOT capture "color as per RC" or "Cubic Capacity"
            let niaMake = "";
            let niaModel = "";
            const makeM = fullText.match(/\bMake\s*:\s*([^:\r\n]+?)(?=\s*(?:Model|Chassis|Engine|Registered|Type|$))/i);
            if (makeM && !isFormLabelText(makeM[1]) && !/color|cubic|capacity/i.test(makeM[1])) {
                niaMake = clean(makeM[1]);
            }
            const modelM = fullText.match(/\bModel\s*:\s*([^:\r\n]+?)(?=\s*(?:Chassis|Engine|Type|Make|$))/i);
            if (modelM && !isFormLabelText(modelM[1]) && !/color|cubic|capacity/i.test(modelM[1])) {
                niaModel = clean(modelM[1]);
            }
            if (niaMake || niaModel) {
                const combined = `${niaMake} ${niaModel}`.trim();
                if (isValidMakeModel(combined, data.claim_no)) {
                    data.make_model = combined;
                }
            }

            // 6G. NIA Driver Relation: e.g. Other, Self
            const niaRelMatch = fullText.match(/\bRelationship\s*of\s*the\s*Driver\s*with\s*the\s*Insured\s*:\s*([^:\r\n]+?)(?=\s*(?:Driver\s*Name|Age|Address|$))/i);
            if (niaRelMatch && clean(niaRelMatch[1]) && !isFormLabelText(niaRelMatch[1])) {
                data.driver_relation = clean(niaRelMatch[1]);
            }

            // 6H. NIA IDV: Must be numeric only, never "No Objection Certificate"
            data.idv = cleanIdv(data.idv);

            // 6I. NIA Policy Start Date: If it contains "time", strip or wipe
            if (data.policy_start_date && (/time/i.test(data.policy_start_date) || !isValidDate(data.policy_start_date))) {
                const pDateMatch = data.policy_start_date.match(/(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})/);
                data.policy_start_date = pDateMatch ? pDateMatch[1] : "";
            }
        }

        // 7. Universal Fallback & Defensive Sanitization across ALL Companies
        // 7A. Phone number fallback if not detected via form anchors
        if (!data.phone) {
            const phoneMatch = fullText.match(/\b(?:Insured\s*Contact|Contact\s*No|Mobile\s*No|Customer\s*Mobile|Mobile|Phone\s*No|Phone)\s*[:\-]?\s*(?:(?:\+|0{0,2})91[\s\-]*)?([6-9]\d{9})\b/i);
            if (phoneMatch) {
                data.phone = cleanPhone(phoneMatch[1]);
            }
        }

        // 7B. Clean insured name: strip leaked "Intimation Date" or label suffixes
        if (data.insured_name) {
            if (/intimation\s*date/i.test(data.insured_name)) {
                data.insured_name = clean(data.insured_name.replace(/\bIntimation\s*Date\b.*$/i, ""));
            }
            if (/date\s*of\s*notification/i.test(data.insured_name)) {
                data.insured_name = clean(data.insured_name.replace(/\bDate\s*of\s*Notification\b.*$/i, ""));
            }
            if (/^(intimation|contact|loss|policy|claim|date|time)/i.test(data.insured_name)) {
                data.insured_name = "";
            }
        }

        // 7C. Clean claim number: strip accidental prefix labels like "mber : " or "claim no : "
        if (data.claim_no) {
            data.claim_no = clean(data.claim_no.replace(/^.*?(?:claim\s*(?:number|no\.?)|mber|no\.?)\s*[:\-]?\s*/i, ""));
        }

        // 7D. Clean vehicle number: must never be form labels like "Cubic Capacity"
        if (data.vehicle_no && (/cubic|capacity|color\s*as\s*per|make|model/i.test(data.vehicle_no) || isFormLabelText(data.vehicle_no))) {
            data.vehicle_no = "";
        }

        // 7E. Clean make & model: must never be form labels like "color as per RC"
        if (data.make_model && (/color\s*as\s*per|cubic\s*capacity|chassis|engine/i.test(data.make_model) || isFormLabelText(data.make_model))) {
            data.make_model = "";
        }

        // 7F. Clean date of survey: must be a real calendar date, NEVER status strings like "Partial Loss"
        if (!isValidDate(data.date_of_survey)) {
            data.date_of_survey = todayDate;
        }

        // 7G. Clean date of loss: must be a real date, never contain "Estimated Amount", "Time of Loss"
        if (data.date_of_loss) {
            const dateM = data.date_of_loss.match(/(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})/);
            if (dateM) {
                data.date_of_loss = dateM[1];
            } else if (/estimated|amount|intimation|time|loss/i.test(data.date_of_loss)) {
                data.date_of_loss = "";
            }
        }

        // 7H. Clean policy start date: must be a real date, never "OD policy start time : 00:00:01"
        if (data.policy_start_date) {
            const dateM = data.policy_start_date.match(/(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})/);
            if (dateM) {
                data.policy_start_date = dateM[1];
            } else if (/time|policy|start|od/i.test(data.policy_start_date)) {
                data.policy_start_date = "";
            }
        }

        // 7I. Clean workshop: if it contains "Workshop Phone No" or is a label like "Repairer City:", clean it
        if (data.workshop) {
            if (/workshop\s*phone/i.test(data.workshop)) {
                data.workshop = clean(data.workshop.replace(/\bWorkshop\s*Phone\b.*$/i, ""));
            }
            if (/^repairer\s*city/i.test(data.workshop) || data.workshop.endsWith(":") || isFormLabelText(data.workshop)) {
                data.workshop = "";
            }
        }

        // 7J. Clean IDV: if "0", "Nil Depreciation", "Nil", "NA", "No Objection Certificate", leave empty
        data.idv = cleanIdv(data.idv);

        // 7. Cache and Broadcast to Top Frame and Storage
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
