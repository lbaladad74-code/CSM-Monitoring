import { initializeApp } from "https://www.gstatic.com/firebasejs/12.10.0/firebase-app.js";
import { getFirestore, collection, addDoc, onSnapshot, query, orderBy, doc, deleteDoc, updateDoc } 
from "https://www.gstatic.com/firebasejs/12.10.0/firebase-firestore.js";

// --- FIREBASE CONFIGURATION ---
const firebaseConfig = {
    apiKey: "AIzaSyA9WsRovhzkLjT27DxoK7R4wzh-Ou4FxZY",
    authDomain: "csm-monitoring.firebaseapp.com",
    projectId: "csm-monitoring",
    storageBucket: "csm-monitoring.firebasestorage.app",
    messagingSenderId: "605728862475",
    appId: "1:605728862475:web:31158389ab5b6893aef20e"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// --- GLOBAL VARIABLES ---
let globalCsmData = []; 
let currentEditId = null; 
let divisionChart, sourceChart, clientChart; 

// --- FULL SERVICES DATABASE LIST ---
const servicesData = [
    { group: "1.0 OFFICE OF THE REGIONAL DIRECTOR", items: [{ ref: "1.1", name: "Issuance of Foreign Official Travel Authority" }, { ref: "1.2", name: "Issuance of Foreign Personnal Travel Authority" }, { ref: "1.3", name: "Other Inquiries / Request" }] },
    { group: "2.0 OFFICE OF THE ASSISTANT REGIONAL DIRECTOR", items: [{ ref: "2.1", name: "BAC" }, { ref: "2.2", name: "Other Transactions" }] },
    { group: "3.a Asset Management Section", items: [{ ref: "3.a.1", name: "Requisition of Supplies, Materials and Equipment" }, { ref: "3.a.2", name: "Acceptance of Deliveries" }, { ref: "3.a.3", name: "Other Transactions" }] },
    { group: "3.b Cash unit", items: [{ ref: "3.b.1", name: "Payment of External and Internal Claims" }, { ref: "3.b.2", name: "Payment of Obligation" }, { ref: "3.b.3", name: "Handling of Cash Advances" }, { ref: "3.b.4", name: "Other Transactions" }] },
    { group: "3.c General Services Unit", items: [{ ref: "3.c.1", name: "Services provided to offices" }, { ref: "3.c.2", name: "Other Inquiries/Requests" }] },
    { group: "3.d Personnel Section", items: [{ ref: "3.d.1", name: "Acceptance of Employment (Walk-in)" }, { ref: "3.d.2", name: "Acceptance of Employment (Online)" }, { ref: "3.d.3", name: "Issuance of Certificate of Last Payment" }, { ref: "3.d.4", name: "Application for Leave" }, { ref: "3.d.5", name: "Application for Retirement/ Survivorship/Disability Benefit" }, { ref: "3.d.6", name: "Issuance of Certificate for Remittances" }, { ref: "3.d.7", name: "Issuance of Certificate of Employment and/or Service Record" }, { ref: "3.d.8", name: "Issuance of Foreign Official Travel Authority" }, { ref: "3.d.9", name: "Issuance of Foreign Personal Travel Authority" }, { ref: "3.d.10", name: "Processing of Equivalent Record Form (ERF)" }, { ref: "3.d.11", name: "Processing of Study Leave" }, { ref: "3.d.12", name: "Processing of Terminal Leave Benefits" }, { ref: "3.d.13", name: "Request for Transfer from Another Region" }, { ref: "3.d.14", name: "Other Inquiries/Requests" }, { ref: "3.d.15", name: "Other Transactions" }] },
    { group: "3.e Procurement Unit", items: [{ ref: "3.e.1", name: "BAC/Procurement-related transactions" }] },
    { group: "3.f Records Section", items: [{ ref: "3.f.1", name: "Certification, Authenthication and Verification" }, { ref: "3.f.2", name: "Issuance of Requested Documents (CTC and Photocopy of Documents)" }, { ref: "3.f.3", name: "Issuance of Requested Documents (Non-CTC)" }, { ref: "3.f.4", name: "Receiving of Communication" }, { ref: "3.f.5", name: "Receiving of Complaint" }, { ref: "3.f.6", name: "Document Routing and Tracking using the Documented Management System" }, { ref: "3.f.7", name: "Other Requests/Inquiries" }] },
    { group: "3.g Regional Payroll Services Unit", items: [{ ref: "3.g.1", name: "Stoppage/ Deletion of Deductions in the Payroll (Loans and Insurances)" }, { ref: "3.g.2", name: "Request for Copy of Payslip" }, { ref: "3.g.3", name: "Submission of billing from Private Kending Instituitions" }, { ref: "3.g.4", name: "Other Requests/Inquiries" }] },
    { group: "4.0 CURRICULUM AND LEARNING MANAGEMENT DIVISION", items: [{ ref: "4.1", name: "Joint Delivery Voucher Program (JDVP)" }, { ref: "4.2", name: "Inquiries/ Feedback" }, { ref: "4.3", name: "Other Requests/Inquiries" }, { ref: "4.4", name: "Other Transactions" }] },
    { group: "5.0 CLMD - LEARNING RESOURCE MANAGEMENT SECTION", items: [{ ref: "5.1", name: "Access to LRMDS Portal" }, { ref: "5.2", name: "Procedure for the Use of LRMDS Computers" }, { ref: "5.3", name: "Development and Quality Assurance of Learning Materials" }, { ref: "5.4", name: "Borrowing of Books/Learning Materials" }, { ref: "5.5", name: "Other Transactions" }] },
    { group: "6.0 EDUCATION SUPPORT SERVICES DIVISION", items: [{ ref: "6.1", name: "School Facilities" }, { ref: "6.2", name: "School Health and Nutrition Services" }, { ref: "6.3", name: "Other Services" }] },
    { group: "7.0 FINANCE DIVISION", items: [{ ref: "7.a.1", name: "Certification as to Availability of Funds" }, { ref: "7.a.2", name: "Endorsement of Request for Cash Allocation from SDOs" }, { ref: "7.a.3", name: "Other Requests/Inquiries" }, { ref: "7.a.4", name: "Other Transactions" }, { ref: "7.b.1", name: "Disbursement Updating" }, { ref: "7.b.2", name: "Downloading/Fund Transfers of SAROs received from Central Office to School Division Office and Implementing Units" }, { ref: "7.b.3", name: "Letter of Acceptance for Downloaded Funds" }, { ref: "7.b.4", name: "Obligation of Expenditure (Incurrence of Obligation Charged to Approved Budget Allocation per GAARD and Other Budget Laws / Authority)" }, { ref: "7.b.5", name: "Processing of Budget Utilization Request & Status (BURS)" }, { ref: "7.b.6", name: "Other Inquiries/Requests" }, { ref: "7.b.7", name: "Other Transactions" }] },
    { group: "8.0 FIELD TECHNICAL ASSISTANCE DIVISION", items: [{ ref: "8.1", name: "Provision of Technical Assistance" }, { ref: "8.2", name: "Other Transactions" }] },
    { group: "9.0 Human Resource Development Division", items: [{ ref: "9.1", name: "Rewards and Recognitions" }, { ref: "9.2", name: "Other Transactions" }] },
    { group: "10.0 National Educators Academy of the Philippines (NEAP-R1)", items: [{ ref: "10.1", name: "Recognition of Professional Development Programs / Courses" }, { ref: "10.2", name: "Use of Venue; Trainings Provided/Facilitated" }, { ref: "10.3", name: "Others" }] },
    { group: "11.0 ORD - ICTU", items: [{ ref: "11.1", name: "Troubleshooting of  ICT Equipmnet" }, { ref: "11.2", name: "Other Requests/Inquiries" }] },
    { group: "12.0 ORD-Legal Unit", items: [{ ref: "12.1", name: "Legal Assistance to Walk-in Clients" }, { ref: "12.2", name: "Request for Correction of  Entries in School Record" }, { ref: "12.3", name: "Request for Certification as to the Pendency or Non-Pendency of an Administrative Case" }, { ref: "12.4", name: "Other Transactions" }] },
    { group: "13.0 ORD- Public Affairs Unit", items: [{ ref: "13.1", name: "Public Assistance (Email)" }, { ref: "13.2", name: "Public assistance (Hotline and Walk-in)" }, { ref: "13.3", name: "Standard Freedom of Information Request through Walk-In Facility and Mail" }, { ref: "13.4", name: "Processing of communication received through the Public Assistance Action Center (PAAC)" }, { ref: "13.5", name: "Other Transactions" }] },
    { group: "14.0 Policy, Planning and Research Division", items: [{ ref: "14.1", name: "Generation of School IDs for New Schools and/or Adding or Updating of SHS Program Offering" }, { ref: "14.2", name: "Request for Reversion" }, { ref: "14.3", name: "Request for Basic Education Data" }, { ref: "14.4", name: "Other Transactions" }] },
    { group: "15.0 Quality Assurance Division", items: [{ ref: "15.1", name: "Application for Opening/Additional Offering of SHS Program for Private Schools" }, { ref: "15.2", name: "Application for Tuition and Other School Fees (TOSF), No Increase, and Proposed New Fees of Private Schools" }, { ref: "15.3", name: "Issuance of Special Orders for the Graduation of Private School Learners" }, { ref: "15.4", name: "Application for Establishment, Merging, Conversion, and Naming/Renaming of Public Schools and Separation of Public Schools" }, { ref: "15.5", name: "Application of Government Permit to Operate" }, { ref: "15.6", name: "Application for Renewal of Government Permit to Operate" }, { ref: "15.7", name: "Application for Government Recognition" }, { ref: "15.8", name: "Other Transactions" }] }
];

// --- INITIALIZATION ON LOAD ---
window.onload = function() {
    // 1. Build Dropdowns
    const sectionSelect = document.getElementById('sectionUnit');
    const docSectionSelect = document.getElementById('reportDivisionSelect');
    
    servicesData.forEach(group => {
        const option = document.createElement('option'); option.value = group.group; option.textContent = group.group;
        sectionSelect.appendChild(option);
        
        const docOpt = document.createElement('option'); docOpt.value = group.group; docOpt.textContent = group.group;
        docSectionSelect.appendChild(docOpt);
    });

    // 2. Build the 9 SQD Form Inputs Dynamically
    const sqdContainer = document.getElementById('sqdGridContainer');
    const sqdLabels = ["SQD0", "SQD1 (Responsiveness)", "SQD2 (Reliability)", "SQD3 (Access & Facilities)", "SQD4 (Communication)", "SQD5 (Costs)", "SQD6 (Integrity)", "SQD7 (Assurance)", "SQD8 (Outcome)"];
    let sqdHtml = '';
    sqdLabels.forEach((label, index) => {
        sqdHtml += `<div><label>${label}</label>
            <select id="sqd${index}" required>
                <option value="">Select...</option>
                <option value="STRONGLY AGREE">STRONGLY AGREE</option>
                <option value="AGREE">AGREE</option>
                <option value="NEITHER AGREE OR DISAGREE">NEITHER AGREE OR DISAGREE</option>
                <option value="DISAGREE">DISAGREE</option>
                <option value="STRONGLY DISAGREE">STRONGLY DISAGREE</option>
                <option value="N/A">N/A</option>
            </select></div>`;
    });
    if(sqdContainer) sqdContainer.innerHTML = sqdHtml;

    // 3. Setup dates
    const dateInput = document.getElementById('date');
    if(dateInput) dateInput.valueAsDate = new Date();
};

// --- UI NAVIGATION & FORM LOGIC ---
window.switchTab = function(tabId) {
    document.querySelectorAll('.view-section').forEach(el => el.classList.remove('active-section'));
    document.querySelectorAll('.nav-links li').forEach(el => el.classList.remove('active-nav'));
    document.getElementById(tabId).classList.add('active-section');
    document.getElementById('nav-' + tabId).classList.add('active-nav');
};

window.updateServiceDropdown = function(prefillValue = null) {
    const sectionValue = document.getElementById('sectionUnit').value;
    const serviceSelect = document.getElementById('serviceCC');
    const refInput = document.getElementById('refCode');
    serviceSelect.innerHTML = '<option value="">Select Service...</option>';
    refInput.value = "";
    if (sectionValue) {
        const selectedGroup = servicesData.find(g => g.group === sectionValue);
        if (selectedGroup) {
            selectedGroup.items.forEach(item => {
                const option = document.createElement('option');
                option.value = `${item.ref}|||${item.name}`; 
                option.textContent = item.name;
                serviceSelect.appendChild(option);
            });
        }
        if (prefillValue) serviceSelect.value = prefillValue;
    }
};

window.updateRefCode = function() {
    const select = document.getElementById('serviceCC');
    const refInput = document.getElementById('refCode');
    if (select.value) refInput.value = select.value.split('|||')[0]; 
    else refInput.value = "";
};

// --- TABLE GENERATION ---
function generateMegaTableHTML(tbodyId) {
    let sqdHeaders = ''; let sqdSubHeaders = '';
    const sqdNames = ["SQD0", "SQD1 (RESPONSIVENESS)", "SQD2 (RELIABILITY)", "SQD3 (ACCESS AND FACILITIES)", "SQD4 (COMMUNICATION)", "SQD5 (COSTS)", "SQD6 (INTEGRITY)", "SQD7 (ASSURANCE)", "SQD8 (OUTCOME)"];
    sqdNames.forEach(name => {
        sqdHeaders += `<th colspan="6" class="bg-green">${name}</th>`;
        sqdSubHeaders += `<th class="bg-green">STRONGLY AGREE</th><th class="bg-green">AGREE</th><th class="bg-green">NEITHER</th><th class="bg-green">DISAGREE</th><th class="bg-green">STRONGLY DISAGREE</th><th class="bg-green">N/A</th>`;
    });
    return `<table class="mega-table"><thead><tr><th rowspan="3" class="bg-yellow">Ref Code</th><th rowspan="3" class="bg-yellow">Service availed as per Citizen's Charter</th><th colspan="4" class="bg-green">CLIENT TYPE</th><th colspan="3" class="bg-green">SEX</th><th colspan="6" class="bg-green">AGE</th><th rowspan="3" class="bg-green">REGION OF RESIDENCE</th><th colspan="13" class="bg-green">CITIZEN'S CHARTER</th>${sqdHeaders}<th rowspan="3" class="bg-yellow">Actions</th></tr><tr><th rowspan="2" class="bg-green">CITIZEN</th><th rowspan="2" class="bg-green">GOVERNMENT</th><th rowspan="2" class="bg-green">BUSINESS</th><th rowspan="2" class="bg-green">DID NOT SPECIFY</th><th rowspan="2" class="bg-green">MALE</th><th rowspan="2" class="bg-green">FEMALE</th><th rowspan="2" class="bg-green">DID NOT SPECIFY</th><th rowspan="2" class="bg-green">19 or Lower</th><th rowspan="2" class="bg-green">20-34</th><th rowspan="2" class="bg-green">35-49</th><th rowspan="2" class="bg-green">50-64</th><th rowspan="2" class="bg-green">65 or higher</th><th rowspan="2" class="bg-green">DID NOT SPECIFY</th><th colspan="4" class="bg-green">CC1</th><th colspan="5" class="bg-green">CC2</th><th colspan="4" class="bg-green">CC3</th></tr><tr><th class="bg-green">1</th><th class="bg-green">2</th><th class="bg-green">3</th><th class="bg-green">4</th><th class="bg-green">1</th><th class="bg-green">2</th><th class="bg-green">3</th><th class="bg-green">4</th><th class="bg-green">5</th><th class="bg-green">1</th><th class="bg-green">2</th><th class="bg-green">3</th><th class="bg-green">4</th>${sqdSubHeaders}</tr></thead><tbody id="${tbodyId}"></tbody></table>`;
}
document.getElementById('onsiteTableContainer').innerHTML = generateMegaTableHTML('onsiteTallyBody');
document.getElementById('onlineTableContainer').innerHTML = generateMegaTableHTML('onlineTallyBody');


// --- SUBMIT / UPDATE DATA ---
document.getElementById('csmForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const sqdValues = {};
    for(let i=0; i<=8; i++) sqdValues[`sqd${i}`] = document.getElementById(`sqd${i}`).value;
    const serviceCombined = document.getElementById('serviceCC').value;
    const cleanServiceName = serviceCombined ? serviceCombined.split('|||')[1] : "";

    const payload = {
        source: document.getElementById('dataSource').value, date: document.getElementById('date').value, controlNo: document.getElementById('controlNo').value,
        clientType: document.getElementById('clientType').value, sex: document.getElementById('sex').value, age: document.getElementById('age').value,
        region: document.getElementById('region').value, sectionUnit: document.getElementById('sectionUnit').value, serviceCC: cleanServiceName, 
        refCode: document.getElementById('refCode').value, actualService: document.getElementById('actualService').value,
        cc: { cc1: document.getElementById('cc1').value, cc2: document.getElementById('cc2').value, cc3: document.getElementById('cc3').value }, sqd: sqdValues, remarks: document.getElementById('remarks').value
    };

    try {
        if (currentEditId) { await updateDoc(doc(db, "csm_master", currentEditId), payload); alert("Record Updated Successfully!"); } 
        else { payload.timestamp = new Date(); await addDoc(collection(db, "csm_master"), payload); alert("Record Saved Successfully!"); }
        window.cancelEdit();
    } catch (error) { console.error(error); }
});

// --- EDIT & DELETE LOGIC ---
window.deleteRecord = async function(id) {
    if(confirm("Are you sure you want to permanently delete this record?")) { try { await deleteDoc(doc(db, "csm_master", id)); } catch (e) { console.error("Error deleting:", e); } }
};

window.editRecord = function(id) {
    const record = globalCsmData.find(item => item.id === id); if(!record) return;
    switchTab('master-data'); document.getElementById('mainScrollArea').scrollTop = 0;
    
    document.getElementById('dataSource').value = record.source; document.getElementById('date').value = record.date; document.getElementById('controlNo').value = record.controlNo;
    document.getElementById('clientType').value = record.clientType; document.getElementById('sex').value = record.sex; document.getElementById('age').value = record.age;
    document.getElementById('region').value = record.region; document.getElementById('sectionUnit').value = record.sectionUnit;
    window.updateServiceDropdown(`${record.refCode}|||${record.serviceCC}`);
    document.getElementById('refCode').value = record.refCode; document.getElementById('actualService').value = record.actualService;
    document.getElementById('cc1').value = record.cc.cc1; document.getElementById('cc2').value = record.cc.cc2; document.getElementById('cc3').value = record.cc.cc3;
    
    for(let i=0; i<=8; i++) document.getElementById(`sqd${i}`).value = record.sqd[`sqd${i}`];
    document.getElementById('remarks').value = record.remarks || "";
    
    currentEditId = id; document.getElementById('formTitle').innerText = `Edit Record (ID: ${id.substring(0,5)}...)`;
    document.getElementById('formCard').style.border = "2px solid #f59e0b"; document.getElementById('submitBtn').innerText = "Update Record";
    document.getElementById('submitBtn').style.backgroundColor = "#f59e0b"; document.getElementById('cancelEditBtn').style.display = "block";
};

window.cancelEdit = function() {
    currentEditId = null; document.getElementById('csmForm').reset(); document.getElementById('serviceCC').innerHTML = '<option value="">Select Unit First...</option>';
    document.getElementById('refCode').value = ""; document.getElementById('date').valueAsDate = new Date(); document.getElementById('formTitle').innerText = "Data Input Form";
    document.getElementById('formCard').style.border = "none"; document.getElementById('submitBtn').innerText = "Save Record to Master Data";
    document.getElementById('submitBtn').style.backgroundColor = "var(--primary)"; document.getElementById('cancelEditBtn').style.display = "none";
};

// --- REAL-TIME DATA PROCESSING & REPORTS ---
const sqdScoreMap = { "STRONGLY AGREE": 5, "AGREE": 4, "NEITHER AGREE OR DISAGREE": 3, "DISAGREE": 2, "STRONGLY DISAGREE": 1 };
function getInterpretation(score) {
    if (score >= 95.0) return "Outstanding"; if (score >= 90.0) return "Very Satisfactory";
    if (score >= 80.0) return "Satisfactory"; if (score >= 60.0) return "Fair"; return "Poor";
}

function updateDashboardCharts() {
    const divCounts = {}; let onsite = 0, online = 0; let cit = 0, gov = 0, bus = 0, dns = 0;
    globalCsmData.forEach(d => {
        const divName = d.sectionUnit ? d.sectionUnit.substring(0, 15) + "..." : "Unknown"; divCounts[divName] = (divCounts[divName] || 0) + 1;
        if (d.source === 'Onsite') onsite++; else online++;
        if (d.clientType === 'Citizen') cit++; else if (d.clientType === 'Government') gov++; else if (d.clientType === 'Business') bus++; else dns++;
    });
    const divCtx = document.getElementById('divisionChart'); if(divisionChart) divisionChart.destroy();
    divisionChart = new Chart(divCtx, { type: 'bar', data: { labels: Object.keys(divCounts), datasets: [{ label: 'Total Submissions', data: Object.values(divCounts), backgroundColor: '#3b82f6', borderRadius: 4 }] }, options: { responsive: true, maintainAspectRatio: false } });
    const srcCtx = document.getElementById('sourceChart'); if(sourceChart) sourceChart.destroy();
    sourceChart = new Chart(srcCtx, { type: 'doughnut', data: { labels: ['Onsite', 'Online'], datasets: [{ data: [onsite, online], backgroundColor: ['#f59e0b', '#10b981'] }] }, options: { responsive: true, maintainAspectRatio: false, cutout: '70%' } });
    const cliCtx = document.getElementById('clientChart'); if(clientChart) clientChart.destroy();
    clientChart = new Chart(cliCtx, { type: 'doughnut', data: { labels: ['Citizen', 'Gov', 'Business', 'N/A'], datasets: [{ data: [cit, gov, bus, dns], backgroundColor: ['#3b82f6', '#8b5cf6', '#ec4899', '#cbd5e1'] }] }, options: { responsive: true, maintainAspectRatio: false, cutout: '70%' } });
}

window.updateOfficialReport = function() {
    const divName = document.getElementById('reportDivisionSelect').value;
    let filteredData = divName !== "ALL" ? globalCsmData.filter(d => d.sectionUnit === divName) : globalCsmData;
    document.getElementById('docDivName').innerText = divName !== "ALL" ? divName : "ALL FUNCTIONAL DIVISIONS (REGIONAL SUMMARY)";
    const serviceStats = {}; let totalResp = filteredData.length; document.getElementById('docTotalResp').innerText = totalResp;
    let sqdTotals = { 0:0, 1:0, 2:0, 3:0, 4:0, 5:0, 6:0, 7:0, 8:0 }; let sqdValids = { 0:0, 1:0, 2:0, 3:0, 4:0, 5:0, 6:0, 7:0, 8:0 };
    let grandTotalSqd = 0; let grandValidSqd = 0; const remarksList = [];

    filteredData.forEach(d => {
        const svc = d.serviceCC || "Unspecified Service"; if (!serviceStats[svc]) serviceStats[svc] = { total: 0, valid: 0 };
        let rowTotal = 0; let rowValid = 0;
        for(let i=0; i<=8; i++) {
            const val = d.sqd[`sqd${i}`];
            if (val && val !== "N/A" && sqdScoreMap[val]) {
                const score = sqdScoreMap[val]; sqdTotals[i] += score; sqdValids[i]++;
                if (i > 0) { rowTotal += score; rowValid++; grandTotalSqd += score; grandValidSqd++; }
            }
        }
        if (rowValid > 0) { serviceStats[svc].total += (rowTotal / rowValid); serviceStats[svc].valid++; }
        if (d.remarks && d.remarks.trim() !== "") remarksList.push(d.remarks);
    });

    let svcHtml = "";
    for (const [svcName, stats] of Object.entries(serviceStats)) {
        const avg = stats.valid > 0 ? (stats.total / stats.valid).toFixed(2) : "N/A";
        svcHtml += `<tr><td style="padding: 5px;">${svcName}</td><td style="text-align:center; padding: 5px;">${avg}</td></tr>`;
    }
    document.getElementById('docServicesBody').innerHTML = svcHtml || "<tr><td colspan='2' style='text-align:center; padding: 5px;'>No services recorded</td></tr>";
    const sqdNames = ["SQD0", "SQD1 Responsiveness", "SQD2 Reliability", "SQD3 Access and Facilities", "SQD4 Communication", "SQD5 Costs", "SQD6 Integrity", "SQD7 Assurance", "SQD8 Outcome"];
    let sqdHtml = "";
    for(let i=0; i<=8; i++) {
        const avg = sqdValids[i] > 0 ? (sqdTotals[i] / sqdValids[i]).toFixed(2) : "N/A";
        sqdHtml += `<tr><td style="padding: 5px;">${sqdNames[i]}</td><td style="text-align:center; padding: 5px;">${avg}</td></tr>`;
    }
    sqdHtml += `<tr><td style="font-weight:bold; padding: 5px;">OVERALL RATING</td><td style="text-align:center; font-weight:bold; padding: 5px;">${grandValidSqd > 0 ? (grandTotalSqd / grandValidSqd).toFixed(2) : "N/A"}</td></tr>`;
    document.getElementById('docSqdBody').innerHTML = sqdHtml;
    document.getElementById('docFeedbackList').innerHTML = remarksList.length > 0 ? remarksList.map(r => `<li style="margin-bottom: 5px;">${r}</li>`).join("") : "<li><em>No feedback provided.</em></li>";
};

window.printReport = function() {
    const printWindow = window.open('', '', 'height=800,width=800');
    printWindow.document.write(`<html><head><title>CSM Official Report</title></head><body style="margin: 0; padding: 20px;">${document.getElementById('printableReport').innerHTML}</body></html>`);
    printWindow.document.close(); printWindow.focus(); setTimeout(() => { printWindow.print(); printWindow.close(); }, 500);
};

window.downloadDocx = function() {
    const html = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Export HTML To Doc</title></head><body>" + document.getElementById('printableReport').innerHTML + "</body></html>";
    const link = document.createElement('a'); link.href = URL.createObjectURL(new Blob(['\ufeff', html], { type: 'application/msword' }));
    const divName = document.getElementById('reportDivisionSelect').value;
    link.download = `CSM_Report_${divName === "ALL" ? "Regional_Summary" : divName.replace(/[^a-z0-9]/gi, '_').toLowerCase()}.doc`;
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
};

window.generateReport = function() {
    const totalResp = globalCsmData.length;
    document.getElementById('dashTotalSubs').innerText = totalResp;
    const totalTrans = parseInt(document.getElementById('totalTransactionsInput').value) || totalResp;
    document.getElementById('repResponseRate').innerText = `${totalTrans > 0 ? ((totalResp / totalTrans) * 100).toFixed(2) : 0}%`;

    if (totalResp === 0) { window.updateOfficialReport(); return; }
    
    const sqdNamesArr = ["SQD0", "SQD1 (Responsiveness)", "SQD2 (Reliability)", "SQD3 (Access & Facilities)", "SQD4 (Communication)", "SQD5 (Costs)", "SQD6 (Integrity)", "SQD7 (Assurance)", "SQD8 (Outcome)"];
    
    let ccAwareCount = 0; let ccHelpfulCount = 0;
    const sqdCounts = {}; for(let i=0; i<=8; i++) sqdCounts[i] = { SA:0, A:0, N:0, D:0, SD:0, NA:0 };
    let totalOverallSA_A = 0; let totalOverallValidResp = 0;

    const ctCounts = { "Citizen": 0, "Government": 0, "Business": 0, "Did Not Specify": 0 };
    const sexCounts = { "Male": 0, "Female": 0, "Did not Specify": 0 };
    const ageCounts = { "19 or Lower": 0, "20-34": 0, "35-49": 0, "50-64": 0, "65 or higher": 0, "DID NOT SPECIFY": 0 };
    const ccTally = { cc1: { '1':0, '2':0, '3':0, '4':0, 'NR':0 }, cc2: { '1':0, '2':0, '3':0, '4':0, '5':0, 'NR':0 }, cc3: { '1':0, '2':0, '3':0, '4':0, 'NR':0 } };

    const refCounts = {}; servicesData.forEach(group => { group.items.forEach(item => { refCounts[item.ref] = 0; }); });
    const divStats = {}; servicesData.forEach(g => { divStats[g.group] = { subs: { q1:0, q2:0, q3:0, q4:0, total:0 }, goods: { q1:0, q2:0, q3:0, q4:0, total:0 } }; });
    const goodCommentsList = document.getElementById('goodCommentsList'); goodCommentsList.innerHTML = "";

    globalCsmData.forEach(data => {
        if(ctCounts[data.clientType] !== undefined) ctCounts[data.clientType]++;
        if(sexCounts[data.sex] !== undefined) sexCounts[data.sex]++;
        if(ageCounts[data.age] !== undefined) ageCounts[data.age]++;
        if(data.refCode) { if (refCounts[data.refCode] !== undefined) refCounts[data.refCode]++; else refCounts[data.refCode] = 1; }

        if (['1','2','3'].includes(data.cc.cc1)) ccAwareCount++;
        if (data.cc.cc3 === '1') ccHelpfulCount++;
        
        if(data.cc.cc1 && ccTally.cc1[data.cc.cc1] !== undefined) ccTally.cc1[data.cc.cc1]++; else ccTally.cc1['NR']++;
        if(data.cc.cc2 && ccTally.cc2[data.cc.cc2] !== undefined) ccTally.cc2[data.cc.cc2]++; else ccTally.cc2['NR']++;
        if(data.cc.cc3 && ccTally.cc3[data.cc.cc3] !== undefined) ccTally.cc3[data.cc.cc3]++; else ccTally.cc3['NR']++;
        
        let tempTotalSqd = 0; let tempValidSqd = 0;

        for(let i=0; i<=8; i++) {
            const val = data.sqd[`sqd${i}`];
            if (val === "STRONGLY AGREE") { sqdCounts[i].SA++; if(i>0) totalOverallSA_A++; }
            else if (val === "AGREE") { sqdCounts[i].A++; if(i>0) totalOverallSA_A++; }
            else if (val === "NEITHER AGREE OR DISAGREE") sqdCounts[i].N++;
            else if (val === "DISAGREE") sqdCounts[i].D++;
            else if (val === "STRONGLY DISAGREE") sqdCounts[i].SD++;
            else if (val === "N/A" || !val) sqdCounts[i].NA++; 
            
            if (i>0 && val !== "N/A" && val) {
                totalOverallValidResp++;
                if(sqdScoreMap[val]) { tempTotalSqd += sqdScoreMap[val]; tempValidSqd++; }
            }
        }

        const div = data.sectionUnit;
        if(divStats[div]) {
            const month = new Date(data.date).getMonth(); 
            let q = 'q1'; if (month >= 3 && month <= 5) q = 'q2'; else if (month >= 6 && month <= 8) q = 'q3'; else if (month >= 9 && month <= 11) q = 'q4';
            divStats[div].subs[q]++; divStats[div].subs.total++;
            const avgScore = tempValidSqd > 0 ? (tempTotalSqd / tempValidSqd) : 0;
            if (data.remarks && data.remarks.trim() !== "" && avgScore >= 4) {
                divStats[div].goods[q]++; divStats[div].goods.total++;
                const tr = document.createElement('tr');
                tr.innerHTML = `<td>${data.date}</td><td>${div}</td><td>${data.serviceCC}</td><td style="color:#16a34a;font-weight:bold;">${avgScore.toFixed(2)}</td><td><em>"${data.remarks}"</em></td>`;
                goodCommentsList.appendChild(tr);
            }
        }
    });

    document.getElementById('repCcAwareness').innerText = `${((ccAwareCount / totalResp) * 100).toFixed(2)}%`;
    document.getElementById('repCcHelpfulness').innerText = `${((ccHelpfulCount / totalResp) * 100).toFixed(2)}%`;
    const overallScore = totalOverallValidResp > 0 ? ((totalOverallSA_A / totalOverallValidResp) * 100).toFixed(2) : 0;
    document.getElementById('repOverallScore').innerText = `${overallScore}%`;
    document.getElementById('repInterpretation').innerText = getInterpretation(overallScore);
    document.getElementById('dashOverallScore').innerText = `${overallScore}%`;

    // ARTA CC Analysis Builder
    const calcPct = (c) => totalResp > 0 ? ((c / totalResp) * 100).toFixed(2) : "0.00";
    document.getElementById('ccAnalysisBody').innerHTML = `
        <tr><td colspan="3" style="background:#e2e8f0; font-weight:bold;">CC1. Which of the following describes your awareness of the CC?</td></tr>
        <tr><td>1. I know what a CC is and I saw this office's CC</td><td style="text-align:center;">${ccTally.cc1['1']}</td><td style="text-align:center;">${calcPct(ccTally.cc1['1'])}%</td></tr>
        <tr><td>2. I know what a CC is but I did not see this office's CC</td><td style="text-align:center;">${ccTally.cc1['2']}</td><td style="text-align:center;">${calcPct(ccTally.cc1['2'])}%</td></tr>
        <tr><td>3. I learned of the CC only when I saw this office's CC</td><td style="text-align:center;">${ccTally.cc1['3']}</td><td style="text-align:center;">${calcPct(ccTally.cc1['3'])}%</td></tr>
        <tr><td>4. I do not know what a CC is and I did not see this office's CC</td><td style="text-align:center;">${ccTally.cc1['4']}</td><td style="text-align:center;">${calcPct(ccTally.cc1['4'])}%</td></tr>
        <tr><td>5. No Response</td><td style="text-align:center;">${ccTally.cc1['NR']}</td><td style="text-align:center;">${calcPct(ccTally.cc1['NR'])}%</td></tr>
        <tr><td colspan="3" style="background:#e2e8f0; font-weight:bold;">CC2. If aware of CC, would you say that the CC of this office was ...?</td></tr>
        <tr><td>1. Easy to see</td><td style="text-align:center;">${ccTally.cc2['1']}</td><td style="text-align:center;">${calcPct(ccTally.cc2['1'])}%</td></tr>
        <tr><td>2. Somewhat easy to see</td><td style="text-align:center;">${ccTally.cc2['2']}</td><td style="text-align:center;">${calcPct(ccTally.cc2['2'])}%</td></tr>
        <tr><td>3. Difficult to see</td><td style="text-align:center;">${ccTally.cc2['3']}</td><td style="text-align:center;">${calcPct(ccTally.cc2['3'])}%</td></tr>
        <tr><td>4. Not visible at all</td><td style="text-align:center;">${ccTally.cc2['4']}</td><td style="text-align:center;">${calcPct(ccTally.cc2['4'])}%</td></tr>
        <tr><td>5. No Response</td><td style="text-align:center;">${ccTally.cc2['5'] + ccTally.cc2['NR']}</td><td style="text-align:center;">${calcPct(ccTally.cc2['5'] + ccTally.cc2['NR'])}%</td></tr>
        <tr><td colspan="3" style="background:#e2e8f0; font-weight:bold;">CC3. If aware of CC, how much did the CC help you in your transaction?</td></tr>
        <tr><td>1. Helped very much</td><td style="text-align:center;">${ccTally.cc3['1']}</td><td style="text-align:center;">${calcPct(ccTally.cc3['1'])}%</td></tr>
        <tr><td>2. Somewhat helped</td><td style="text-align:center;">${ccTally.cc3['2']}</td><td style="text-align:center;">${calcPct(ccTally.cc3['2'])}%</td></tr>
        <tr><td>3. Did not help</td><td style="text-align:center;">${ccTally.cc3['3']}</td><td style="text-align:center;">${calcPct(ccTally.cc3['3'])}%</td></tr>
        <tr><td>4. No Response</td><td style="text-align:center;">${ccTally.cc3['4'] + ccTally.cc3['NR']}</td><td style="text-align:center;">${calcPct(ccTally.cc3['4'] + ccTally.cc3['NR'])}%</td></tr>
    `;

    // ARTA SQD Detailed Breakdown
    let sqdDetailedHtml = ""; let grandSA=0, grandA=0, grandN=0, grandD=0, grandSD=0, grandNA=0;
    for(let i=0; i<=8; i++) {
        const c = sqdCounts[i]; const validResp = totalResp - c.NA;
        const score = validResp > 0 ? (((c.SA + c.A) / validResp) * 100).toFixed(2) : 0;
        if (i > 0) { grandSA += c.SA; grandA += c.A; grandN += c.N; grandD += c.D; grandSD += c.SD; grandNA += c.NA; }
        const rowLabel = i === 0 ? "SQD0" : `${i}-${sqdNamesArr[i].replace("SQD" + i + " ", "").replace(/\(|\)/g, '')}`;
        sqdDetailedHtml += `<tr><td><strong>${rowLabel}</strong></td><td style="text-align:center;">${c.SA}</td><td style="text-align:center;">${c.A}</td><td style="text-align:center;">${c.N}</td><td style="text-align:center;">${c.D}</td><td style="text-align:center;">${c.SD}</td><td style="text-align:center;">${c.NA}</td><td style="text-align:center; background:#e2e8f0;">${totalResp}</td><td style="text-align:center; font-weight:bold; color: #16a34a; background:#dbeafe;">${i===0 ? score + "%" : score}</td></tr>`;
    }
    sqdDetailedHtml += `<tr style="border-top: 2px solid var(--border);"><td style="font-weight:bold;">Overall (SQD)</td><td style="text-align:center; font-weight:bold;">${grandSA}</td><td style="text-align:center; font-weight:bold;">${grandA}</td><td style="text-align:center; font-weight:bold;">${grandN}</td><td style="text-align:center; font-weight:bold;">${grandD}</td><td style="text-align:center; font-weight:bold;">${grandSD}</td><td style="text-align:center; font-weight:bold;">${grandNA}</td><td style="text-align:center; background:#e2e8f0; font-weight:bold;">${totalResp * 8}</td><td style="text-align:center; font-weight:bold; color: #16a34a; background:#dbeafe;">${overallScore}</td></tr>`;
    document.getElementById('sqdAnalysisBody').innerHTML = sqdDetailedHtml;
    
    let basicHtml = "";
    for(let i=1; i<=8; i++) {
        const c = sqdCounts[i]; const validResp = totalResp - c.NA;
        const score = validResp > 0 ? (((c.SA + c.A) / validResp) * 100).toFixed(2) : 0;
        basicHtml += `<tr><td><strong>${sqdNamesArr[i]}</strong></td><td style="text-align:center;">${c.SA}</td><td style="text-align:center;">${c.A}</td><td style="text-align:center;">${c.N}</td><td style="text-align:center;">${c.D}</td><td style="text-align:center;">${c.SD}</td><td style="text-align:center;">${c.NA}</td><td style="text-align:center; font-weight:bold; color: #16a34a; background:#f1f5f9;">${score}%</td><td style="text-align:center; background:#f1f5f9;">${getInterpretation(score)}</td></tr>`;
    }
    document.getElementById('sqdReportBody').innerHTML = basicHtml;

    // Consolidations Tab Updates
    document.getElementById('ctTallyBody').innerHTML = Object.entries(ctCounts).map(([k,v]) => `<tr><td>${k}</td><td style="font-weight:bold;text-align:right;">${v}</td></tr>`).join('');
    document.getElementById('sexTallyBody').innerHTML = Object.entries(sexCounts).map(([k,v]) => `<tr><td>${k}</td><td style="font-weight:bold;text-align:right;">${v}</td></tr>`).join('');
    document.getElementById('ageTallyBody').innerHTML = Object.entries(ageCounts).map(([k,v]) => `<tr><td>${k}</td><td style="font-weight:bold;text-align:right;">${v}</td></tr>`).join('');
    
    let svcMasterHtml = "";
    servicesData.forEach(group => {
        svcMasterHtml += `<tr style="background-color: #e2e8f0;"><td colspan="3" style="font-weight: bold; text-align: left;">${group.group}</td></tr>`;
        group.items.forEach(item => {
            let count = refCounts[item.ref] || 0;
            svcMasterHtml += `<tr><td style="text-align: center; font-weight: bold; color: #64748b;">${item.ref}</td><td style="padding-left: 15px;">${item.name}</td><td style="text-align: center; font-weight: bold; color: ${count > 0 ? '#2563eb' : '#cbd5e1'}; font-size: 14px;">${count}</td></tr>`;
        });
    });
    document.getElementById('serviceMasterTallyBody').innerHTML = svcMasterHtml;

    document.getElementById('ccTallyBody').innerHTML = `
        <tr><td style="text-align:left;font-weight:bold;">CC1</td><td>${ccTally.cc1['1']}</td><td>${ccTally.cc1['2']}</td><td>${ccTally.cc1['3']}</td><td>${ccTally.cc1['4']}</td><td>-</td></tr>
        <tr><td style="text-align:left;font-weight:bold;">CC2</td><td>${ccTally.cc2['1']}</td><td>${ccTally.cc2['2']}</td><td>${ccTally.cc2['3']}</td><td>${ccTally.cc2['4']}</td><td>${ccTally.cc2['5']}</td></tr>
        <tr><td style="text-align:left;font-weight:bold;">CC3</td><td>${ccTally.cc3['1']}</td><td>${ccTally.cc3['2']}</td><td>${ccTally.cc3['3']}</td><td>${ccTally.cc3['4']}</td><td>-</td></tr>
    `;

    let masterSqdHtml = ""; let mGrandSA=0, mGrandA=0, mGrandN=0, mGrandD=0, mGrandSD=0, mGrandNA=0;
    for(let i=0; i<=8; i++) {
        const c = sqdCounts[i]; const validResp = totalResp - c.NA;
        const score = validResp > 0 ? (((c.SA + c.A) / validResp) * 100).toFixed(2) : 0;
        let ratingText = i===0 ? "-" : `${score}% (${getInterpretation(score)})`;
        if (i>0){ mGrandSA += c.SA; mGrandA += c.A; mGrandN += c.N; mGrandD += c.D; mGrandSD += c.SD; mGrandNA += c.NA; }
        masterSqdHtml += `<tr><td style="text-align:left; font-weight:bold;">${sqdNamesArr[i]}</td><td>${c.SA}</td><td>${c.A}</td><td>${c.N}</td><td>${c.D}</td><td>${c.SD}</td><td style="color:#64748b;">${c.NA}</td><td style="background:#f1f5f9; font-weight:bold; color: #10b981;">${ratingText}</td></tr>`;
    }
    masterSqdHtml += `<tr style="background:#e2e8f0; border-top: 2px solid var(--border);"><td style="text-align:left; font-weight:bold; font-size: 14px;">TOTAL</td><td style="font-weight:bold;">${mGrandSA}</td><td style="font-weight:bold;">${mGrandA}</td><td style="font-weight:bold;">${mGrandN}</td><td style="font-weight:bold;">${mGrandD}</td><td style="font-weight:bold;">${mGrandSD}</td><td style="font-weight:bold; color:#64748b;">${mGrandNA}</td><td style="font-weight:bold; color: #16a34a; font-size: 14px;">${overallScore}% (${getInterpretation(overallScore)})</td></tr>`;
    document.getElementById('sqdMasterTallyBody').innerHTML = masterSqdHtml;
    
    let subHtml = ""; let goodHtml = "";
    servicesData.forEach(group => {
        const div = group.group; const s = divStats[div].subs; const g = divStats[div].goods;
        subHtml += `<tr><td>${div}</td><td style="text-align:center;">${s.q1}</td><td style="text-align:center;">${s.q2}</td><td style="text-align:center;">${s.q3}</td><td style="text-align:center;">${s.q4}</td><td style="text-align:center; font-weight:bold; background:#f1f5f9;">${s.total}</td></tr>`;
        goodHtml += `<tr><td>${div}</td><td style="text-align:center;">${g.q1}</td><td style="text-align:center;">${g.q2}</td><td style="text-align:center;">${g.q3}</td><td style="text-align:center;">${g.q4}</td><td style="text-align:center; font-weight:bold; background:#f1f5f9; color:#16a34a;">${g.total}</td></tr>`;
    });
    document.getElementById('subQuarterlyBody').innerHTML = subHtml;
    document.getElementById('goodQuarterlyBody').innerHTML = goodHtml;

    updateDashboardCharts();
    window.updateOfficialReport();
};

function buildTallyRow(data, id) {
    let r = `<td>${data.refCode}</td><td>${data.serviceCC}</td>`;
    r += `<td>${data.clientType==='Citizen'?'1':''}</td><td>${data.clientType==='Government'?'1':''}</td><td>${data.clientType==='Business'?'1':''}</td><td>${data.clientType==='Did Not Specify'?'1':''}</td>`;
    r += `<td>${data.sex==='Male'?'1':''}</td><td>${data.sex==='Female'?'1':''}</td><td>${data.sex==='Did not Specify'?'1':''}</td>`;
    r += `<td>${data.age==='19 or Lower'?'1':''}</td><td>${data.age==='20-34'?'1':''}</td><td>${data.age==='35-49'?'1':''}</td><td>${data.age==='50-64'?'1':''}</td><td>${data.age==='65 or higher'?'1':''}</td><td>${data.age==='DID NOT SPECIFY'?'1':''}</td>`;
    r += `<td>${data.region || ''}</td>`;
    let c1 = data.cc.cc1, c2 = data.cc.cc2, c3 = data.cc.cc3;
    r += `<td>${c1==='1'?'1':''}</td><td>${c1==='2'?'1':''}</td><td>${c1==='3'?'1':''}</td><td>${c1==='4'?'1':''}</td>`;
    r += `<td>${c2==='1'?'1':''}</td><td>${c2==='2'?'1':''}</td><td>${c2==='3'?'1':''}</td><td>${c2==='4'?'1':''}</td><td>${c2==='5'?'1':''}</td>`;
    r += `<td>${c3==='1'?'1':''}</td><td>${c3==='2'?'1':''}</td><td>${c3==='3'?'1':''}</td><td>${c3==='4'?'1':''}</td>`;
    for (let i = 0; i <= 8; i++) {
        let s = data.sqd[`sqd${i}`];
        r += `<td>${s==='STRONGLY AGREE'?'1':''}</td><td>${s==='AGREE'?'1':''}</td><td>${s==='NEITHER AGREE OR DISAGREE'?'1':''}</td><td>${s==='DISAGREE'?'1':''}</td><td>${s==='STRONGLY DISAGREE'?'1':''}</td><td>${s==='N/A'?'1':''}</td>`;
    }
    r += `<td><button class="action-btn" onclick="editRecord('${id}')" title="Edit">✏️</button><button class="action-btn" onclick="deleteRecord('${id}')" title="Delete">🗑️</button></td>`;
    return `<tr>${r}</tr>`;
}

onSnapshot(query(collection(db, "csm_master"), orderBy("timestamp", "desc")), (snapshot) => {
    const masterBody = document.getElementById('masterDataBody');
    const onsiteTally = document.getElementById('onsiteTallyBody');
    const onlineTally = document.getElementById('onlineTallyBody');
    masterBody.innerHTML = ""; onsiteTally.innerHTML = ""; onlineTally.innerHTML = ""; 
    globalCsmData = []; 
    snapshot.forEach((doc) => {
        const data = doc.data(); data.id = doc.id; 
        globalCsmData.push(data); 
        let total = 0, valid = 0;
        for(let i=1; i<=8; i++) {
            let val = data.sqd[`sqd${i}`];
            if(val !== "N/A" && sqdScoreMap[val]) { total += sqdScoreMap[val]; valid++; }
        }
        let avg = valid > 0 ? (total / valid).toFixed(2) : "N/A";
        const badge = data.source === 'Online' ? `<span class="badge-online">Online</span>` : `<span class="badge-onsite">Onsite</span>`;
        const basicRow = document.createElement("tr");
        basicRow.innerHTML = `<td>${badge}</td><td>${data.date}</td><td>${data.controlNo}</td><td>${data.clientType}</td><td>${data.sectionUnit}</td><td><strong>${data.refCode}</strong></td><td style="color:#16a34a;font-weight:bold;">${avg}</td><td>${data.remarks || "-"}</td><td><button class="action-btn" onclick="editRecord('${data.id}')" title="Edit">✏️</button><button class="action-btn" onclick="deleteRecord('${data.id}')" title="Delete">🗑️</button></td>`;
        masterBody.appendChild(basicRow);
        const tallyRow = document.createElement("tr");
        tallyRow.innerHTML = buildTallyRow(data, data.id); 
        if (data.source === "Onsite") onsiteTally.appendChild(tallyRow);
        else if (data.source === "Online") onlineTally.appendChild(tallyRow);
    });
    window.generateReport();
});
