// ======= Editable Dashboard System =======
// Author: You + ChatGPT
// Version: 2.1 (Excel borders + better column sizing)
// =========================================

// --- SETTINGS ---
const EDIT_PASSWORD = "12345"; // Change this password

// Load saved data
window.addEventListener("DOMContentLoaded", () => {
    restoreTableData();
    addEditButtonsToTables();
    createTopToolbar();
    showLastUpdated();
});

// === CREATE GLOBAL CONTROL BAR (Save / Download) ===
function createTopToolbar() {
    const controlBar = document.createElement("div");
    controlBar.style.textAlign = "center";
    controlBar.style.margin = "20px";
    controlBar.style.padding = "10px";
    controlBar.style.borderBottom = "2px solid #ccc";
    controlBar.style.background = "#f9f9f9";

    controlBar.innerHTML = `
        <button id="saveBtn" class="toolbar-btn save">💾 Save All Tables</button>
        <button id="downloadHtmlBtn" class="toolbar-btn html">⬇️ Download HTML</button>
        <button id="downloadExcelBtn" class="toolbar-btn excel">📊 Download Excel</button>
        <p id="lastUpdated" style="margin-top:10px; font-weight:bold;"></p>
    `;

    document.body.prepend(controlBar);

    // Add toolbar button styles
    const style = document.createElement("style");
    style.textContent = `
        .toolbar-btn {
            background: #007bff;
            border: none;
            color: white;
            padding: 8px 14px;
            margin: 5px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 15px;
            transition: background 0.3s;
        }
        .toolbar-btn:hover { background: #0056b3; }
        .toolbar-btn.save { background: #28a745; }
        .toolbar-btn.save:hover { background: #1e7e34; }
        .toolbar-btn.html { background: #17a2b8; }
        .toolbar-btn.html:hover { background: #117a8b; }
        .toolbar-btn.excel { background: #ffc107; color: #000; }
        .toolbar-btn.excel:hover { background: #e0a800; }
        .edit-table-btn {
            background: #6c757d;
            color: white;
            border: none;
            padding: 6px 12px;
            margin: 8px 0;
            border-radius: 5px;
            cursor: pointer;
            font-size: 14px;
            transition: background 0.3s;
        }
        .edit-table-btn:hover { background: #5a6268; }
    `;
    document.head.appendChild(style);

    // Hook up the buttons
    document.getElementById("saveBtn").addEventListener("click", saveTableData);
    document.getElementById("downloadHtmlBtn").addEventListener("click", downloadHTML);
    document.getElementById("downloadExcelBtn").addEventListener("click", downloadExcel);
}

// === Add individual edit buttons to each table ===
function addEditButtonsToTables() {
    const tables = document.querySelectorAll("table");

    tables.forEach((table) => {
        const wrapper = document.createElement("div");
        wrapper.style.position = "relative";
        wrapper.style.marginBottom = "30px";

        const editBtn = document.createElement("button");
        editBtn.textContent = "🔒 Edit Table";
        editBtn.className = "edit-table-btn";

        let isEditing = false;

        editBtn.addEventListener("click", () => {
            if (!isEditing) {
                const pass = prompt("Enter password to edit:");
                if (pass !== EDIT_PASSWORD) {
                    alert("❌ Incorrect password!");
                    return;
                }
                table.querySelectorAll("td, th").forEach(cell => cell.contentEditable = true);
                editBtn.textContent = "✅ Save Table";
                editBtn.style.background = "#28a745";
                isEditing = true;
            } else {
                table.querySelectorAll("td, th").forEach(cell => cell.contentEditable = false);
                editBtn.textContent = "🔒 Edit Table";
                editBtn.style.background = "#6c757d";
                isEditing = false;
                saveTableData();
            }
        });

        // Wrap and insert
        table.parentNode.insertBefore(wrapper, table);
        wrapper.appendChild(editBtn);
        wrapper.appendChild(table);
    });
}

// --- Save Table Data to localStorage ---
function saveTableData() {
    const allTables = document.querySelectorAll("table");
    const tableData = [];

    allTables.forEach((table, index) => {
        tableData[index] = table.outerHTML;
    });

    localStorage.setItem("savedTables", JSON.stringify(tableData));
    localStorage.setItem("lastUpdated", new Date().toLocaleString());
    showLastUpdated();
    alert("✅ All tables saved successfully!");
}

// --- Restore Data from localStorage ---
function restoreTableData() {
    const saved = localStorage.getItem("savedTables");
    if (saved) {
        const tables = JSON.parse(saved);
        const htmlTables = document.querySelectorAll("table");
        tables.forEach((tableHTML, i) => {
            if (htmlTables[i]) htmlTables[i].outerHTML = tableHTML;
        });
    }
}

// --- Show Last Updated Time ---
function showLastUpdated() {
    const time = localStorage.getItem("lastUpdated");
    if (time) {
        const el = document.getElementById("lastUpdated");
        if (el) el.textContent = "🕒 Last Updated: " + time;
    }
}

// --- Download as HTML File ---
function downloadHTML() {
    saveTableData(); // ensure data is up to date
    const blob = new Blob([document.documentElement.outerHTML], { type: "text/html" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "updated_dashboard.html";
    link.click();
}

// --- Download as Excel (with borders + column width) ---
function downloadExcel() {
    const tables = document.querySelectorAll("table");
    let html = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" 
              xmlns:x="urn:schemas-microsoft-com:office:excel" 
              xmlns="http://www.w3.org/TR/REC-html40">
        <head>
            <meta charset="UTF-8">
            <style>
                table {
                    border-collapse: collapse;
                    width: auto;
                    font-family: Arial, sans-serif;
                    font-size: 12pt;
                }
                th, td {
                    border: 1px solid #000;
                    padding: 4px 6px;
                    text-align: left;
                    white-space: nowrap;
                    width: 120px; /* reasonable default column width */
                }
                th {
                    background: #f2f2f2;
                }
            </style>
        </head><body>`;

    tables.forEach(t => html += t.outerHTML);
    html += "</body></html>";

    const blob = new Blob([html], { type: "application/vnd.ms-excel" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "dashboard.xls";
    link.click();
}
