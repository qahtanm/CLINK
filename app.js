/**
 * نظام إحصائيات عيادة الرشيدية 2025
 */

let workbookData = null;
let allData = [];
let processedData = [];

// قائمة الأمراض للتقرير الرئيسي (المصفوفة)
const DISEASES_MAPPING = [
    { label: "ارتفاع ضغط الدم", excelKey: "H.T" },
    { label: "مرض القلب الرئوي المزمن", excelKey: "" },
    { label: "امراض القلب الافقازية", excelKey: "" },
    { label: "احتشاء عضلة القلب", excelKey: "" },
    { label: "عجز القلب", excelKey: "قلبية" },
    { label: "امراض قلبية اخرى", excelKey: "" },
    { label: "قرحة الاثني عشر", excelKey: "" },
    { label: "قرحة المعدة", excelKey: "" },
    { label: "فرط الدرقية/انسمام الدرقية(السامة)", excelKey: "" },
    { label: "ضمور الغدة الدرقية", excelKey: "" },
    { label: "اضطرابات اخرى بالدرقية", excelKey: "غدة" },
    { label: "السكر المعتمد على الانسولين", excelKey: "انسولين" },
    { label: "السكر المعتمد على الحبوب", excelKey: "داء السكري" },
    { label: "الصرع", excelKey: "صرع" },
    { label: "الربو", excelKey: "ربو قصبي" },
    { label: "امراض نفسية", excelKey: "ذهان" },
    { label: "امراض اوعية المخ", excelKey: "دماغية" },
    { label: "امراض اخرى", excelKey: "أخرى" }
];

// قائمة الأمراض للتقرير المختصر المضاف حديثاً
const SUMMARY_DISEASES = [
    { label: "انسولين", excelKey: "انسولين" },
    { label: "داء السكري", excelKey: "داء السكري" },
    { label: "قلبية", excelKey: "قلبية" },
    { label: "H.T", excelKey: "H.T" },
    { label: "دماغية", excelKey: "دماغية" },
    { label: "صرع", excelKey: "صرع" },
    { label: "ربو قصبي", excelKey: "ربو قصبي" },
    { label: "غدة", excelKey: "غدة" },
    { label: "ذهان", excelKey: "ذهان" },
    { label: "أخرى", excelKey: "أخرى" }
];

const AGE_GROUPS = [
    "أقل من 15", "15 - 19", "20 - 44", "45 - 64", "65 - 74", "75 فأكثر"
];

function handleFile(input) {
    const file = input.files[0];
    if (!file) return;

    document.getElementById("fileStatus").innerText = `تم اختيار: ${file.name}`;
    document.getElementById("loading").style.display = "block";
    document.getElementById("emptyState").style.display = "none";
    document.getElementById("output").style.display = "none";
    document.getElementById("outputSummary").style.display = "none";
    document.getElementById("filterSection").style.display = "none";

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            workbookData = XLSX.read(data, { type: "array" });

            const sheetFilter = document.getElementById("sheetFilter");
            sheetFilter.innerHTML = "";
            workbookData.SheetNames.forEach((name) => {
                const opt = document.createElement("option");
                opt.value = name;
                opt.innerText = name;
                sheetFilter.appendChild(opt);
            });

            loadSheetData(workbookData.SheetNames[0]);

            document.getElementById("loading").style.display = "none";
            document.getElementById("filterSection").style.display = "flex";
            document.getElementById("output").style.display = "block";
            document.getElementById("outputSummary").style.display = "block";
        } catch (error) {
            console.error(error);
            alert("خطأ في قراءة ملف Excel.");
            resetUI();
        }
    };
    reader.readAsArrayBuffer(file);
}

function changeSheet() {
    const sheetName = document.getElementById("sheetFilter").value;
    if (sheetName) {
        document.getElementById("loading").style.display = "block";
        setTimeout(() => {
            loadSheetData(sheetName);
            document.getElementById("loading").style.display = "none";
        }, 100);
    }
}

function loadSheetData(sheetName) {
    const sheet = workbookData.Sheets[sheetName];
    allData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    if (allData.length === 0) {
        alert("هذه الورقة فارغة!");
        return;
    }
    populateMonthFilter(allData);
    processRawData();
    applyFilters();
}

function processRawData() {
    processedData = [];
    allData.forEach(r => {
        let group = String(r["الفئة العمرية"] || "").trim();
        if (!group) {
            const birthYear = r["سنة"] || r["العمر"];
            const age = calculateAge(birthYear);
            group = getAgeGroup(age);
        }

        const rawGender = String(r["الجنس"] || "").trim();
        const gender = (rawGender.includes("ذكر") || rawGender === "ذ") ? "ذ" : "أ";
        let monthKey = String(r["الشهر"] || "غير محدد").trim();

        // تجميع للتقريرين معاً
        const itemDiseasesMatrix = [];
        DISEASES_MAPPING.forEach(map => {
            if (map.excelKey && (r[map.excelKey] === 1 || r[map.excelKey] === "1")) {
                itemDiseasesMatrix.push(map.label);
            }
        });

        const itemDiseasesSummary = [];
        SUMMARY_DISEASES.forEach(map => {
            if (map.excelKey && (r[map.excelKey] === 1 || r[map.excelKey] === "1")) {
                itemDiseasesSummary.push(map.label);
            }
        });

        processedData.push({
            diseasesMatrix: itemDiseasesMatrix,
            diseasesSummary: itemDiseasesSummary,
            gender: gender,
            ageGroup: group,
            month: monthKey
        });
    });
}

function populateMonthFilter(rows) {
    const monthSelect = document.getElementById("monthFilter");
    monthSelect.innerHTML = '<option value="all">كل الأشهر</option>';
    const months = new Set();
    rows.forEach(row => {
        const m = String(row["الشهر"] || "").trim();
        if (m) months.add(m);
    });
    Array.from(months).forEach(m => {
        const opt = document.createElement("option");
        opt.value = m;
        opt.innerText = m;
        monthSelect.appendChild(opt);
    });
}

function applyFilters() {
    const selectedMonth = document.getElementById("monthFilter").value;
    const filtered = selectedMonth === "all"
        ? processedData
        : processedData.filter(d => d.month === selectedMonth);

    renderSummaryTable(filtered);
    renderStatsTable(filtered);
}

/**
 * بناء التقرير المختصر الجديد (الجدول المضاف أخيراً)
 */
function renderSummaryTable(data) {
    const container = document.getElementById("outputSummary");
    const stats = {};
    SUMMARY_DISEASES.forEach(d => {
        stats[d.label] = { "ذ": 0, "أ": 0 };
    });

    data.forEach(item => {
        item.diseasesSummary.forEach(dLabel => {
            if (stats[dLabel]) {
                stats[dLabel][item.gender]++;
            }
        });
    });

    let html = `
        <h3 style="padding: 1rem; color: var(--primary);">ملخص عام (حسب الجنس)</h3>
        <table class="report-table">
            <thead>
                <tr style="background: #e0f2fe;">
                    <th>المرض</th>
                    <th>انثى</th>
                    <th>ذكر</th>
                    <th>Grand Total</th>
                </tr>
            </thead>
            <tbody>
    `;

    let totalFemales = 0;
    let totalMales = 0;

    SUMMARY_DISEASES.forEach(d => {
        const females = stats[d.label]["أ"];
        const males = stats[d.label]["ذ"];
        const total = females + males;

        totalFemales += females;
        totalMales += males;

        html += `
            <tr>
                <td style="text-align: right; padding-right: 15px;">Sum of ${d.label}</td>
                <td class="num">${females || ''}</td>
                <td class="num">${males || ''}</td>
                <td class="num" style="font-weight: bold;">${total || ''}</td>
            </tr>
        `;
    });

    html += `
        </tbody>
        <tfoot>
            <tr style="background: #f1f5f9; font-weight: 900; border-top: 2px solid #94a3b8;">
                <td style="text-align: right; padding-right: 15px;">المجموع النهائي</td>
                <td class="num">${totalFemales}</td>
                <td class="num">${totalMales}</td>
                <td class="num" style="color: var(--primary); font-size: 1.1rem;">${totalFemales + totalMales}</td>
            </tr>
        </tfoot>
    </table>`;
    container.innerHTML = html;
}

/**
 * بناء التقرير الرئيسي (المصفوفة الكبيرة)
 */
function renderStatsTable(data) {
    const container = document.getElementById("output");
    const stats = {};
    DISEASES_MAPPING.forEach(d => {
        stats[d.label] = {};
        AGE_GROUPS.forEach(ag => { stats[d.label][ag] = { "ذ": 0, "أ": 0 }; });
        stats[d.label]["total"] = { "ذ": 0, "أ": 0 };
    });

    data.forEach(item => {
        item.diseasesMatrix.forEach(dLabel => {
            let matchedGroup = AGE_GROUPS.find(g => item.ageGroup.includes(g) || g.includes(item.ageGroup));
            if (!matchedGroup) matchedGroup = "20 - 44";
            if (stats[dLabel] && stats[dLabel][matchedGroup]) {
                stats[dLabel][matchedGroup][item.gender]++;
                stats[dLabel]["total"][item.gender]++;
            }
        });
    });

    let html = `
        <h3 style="padding: 1rem; color: var(--primary);">التقرير التفصيلي (الفئات العمرية)</h3>
        <table class="report-table">
            <thead>
                <tr>
                    <th rowspan="2">التصنيف</th>
                    <th colspan="2">المجموع</th>
                    ${AGE_GROUPS.map(ag => `<th colspan="2">${ag}</th>`).join('')}
                    <th rowspan="2" style="writing-mode: vertical-rl; text-orientation: mixed; padding: 5px;">العدد الكلي</th>
                </tr>
                <tr>
                    <th>ذ</th><th>أ</th>
                    ${AGE_GROUPS.map(() => `<th>ذ</th><th>أ</th>`).join('')}
                </tr>
            </thead>
            <tbody>
    `;

    let grandTotalMales = 0, grandTotalFemales = 0;
    const ageGroupTotals = {};
    AGE_GROUPS.forEach(ag => ageGroupTotals[ag] = { "ذ": 0, "أ": 0 });

    DISEASES_MAPPING.forEach(d => {
        const rowTotal = stats[d.label]["total"]["ذ"] + stats[d.label]["total"]["أ"];
        grandTotalMales += stats[d.label]["total"]["ذ"];
        grandTotalFemales += stats[d.label]["total"]["أ"];
        html += `<tr>
            <td style="text-align: right; padding-right: 15px; font-weight: 500;">${d.label}</td>
            <td class="num">${stats[d.label]["total"]["ذ"] || '.'}</td>
            <td class="num">${stats[d.label]["total"]["أ"] || '.'}</td>
            ${AGE_GROUPS.map(ag => {
            ageGroupTotals[ag]["ذ"] += stats[d.label][ag]["ذ"];
            ageGroupTotals[ag]["أ"] += stats[d.label][ag]["أ"];
            return `<td class="num">${stats[d.label][ag]["ذ"] || '.'}</td><td class="num">${stats[d.label][ag]["أ"] || '.'}</td>`;
        }).join('')}
            <td class="num" style="background: #f1f5f9; font-weight: bold;">${rowTotal || '.'}</td>
        </tr>`;
    });

    html += `
            <tr style="background: #e2e8f0; font-weight: bold;">
                <td>المجموع</td>
                <td class="num">${grandTotalMales}</td>
                <td class="num">${grandTotalFemales}</td>
                ${AGE_GROUPS.map(ag => `<td class="num">${ageGroupTotals[ag]["ذ"]}</td><td class="num">${ageGroupTotals[ag]["أ"]}</td>`).join('')}
                <td class="num" style="font-size: 1.1rem; color: var(--primary);">${grandTotalMales + grandTotalFemales}</td>
            </tr>
        </tbody>
    </table>`;
    container.innerHTML = html;
}

function calculateAge(val) {
    if (!val) return 0;
    const currentYear = new Date().getFullYear();
    const strVal = String(val).trim();
    if (strVal.length === 4 && !isNaN(strVal)) return currentYear - parseInt(strVal);
    if (!isNaN(strVal)) return parseInt(strVal);
    return 0;
}

function getAgeGroup(age) {
    if (age < 15) return "أقل من 15";
    if (age <= 19) return "15 - 19";
    if (age <= 44) return "20 - 44";
    if (age <= 64) return "45 - 64";
    if (age <= 74) return "65 - 74";
    return "75 فأكثر";
}

function resetUI() {
    document.getElementById("loading").style.display = "none";
    document.getElementById("emptyState").style.display = "block";
    document.getElementById("output").style.display = "none";
    document.getElementById("outputSummary").style.display = "none";
    document.getElementById("filterSection").style.display = "none";
}

function exportToExcel() {
    const tableMatrix = document.querySelector("#output table");
    if (tableMatrix) {
        const wb = XLSX.utils.table_to_book(tableMatrix);
        XLSX.writeFile(wb, "إحصائية_العيادة_التفصيلية.xlsx");
    }
}
