/* =========================================
   DATA DETECTIVE AI
   Frontend Application
========================================= */

const state = {
    dataset: null,
    conversation: [],
    isLoading: false,
    recognition: null,
    isListening: false,
    lastDatasetResponse: null
};

/* =========================================
   DOM
========================================= */

const fileInput =
    document.getElementById("fileInput");

const attachBtn =
    document.getElementById("attachBtn");

const questionInput =
    document.getElementById("questionInput");

const sendBtn =
    document.getElementById("sendBtn");

const voiceBtn =
    document.getElementById("voiceBtn");

const chatContainer =
    document.getElementById("chatContainer");

const welcomeMessage =
    document.getElementById("welcomeMessage");

const typingIndicator =
    document.getElementById("typingIndicator");

const suggestions =
    document.getElementById("suggestions");

const datasetCard =
    document.getElementById("datasetCard");

const datasetName =
    document.getElementById("datasetName");

const datasetRows =
    document.getElementById("datasetRows");

const datasetColumns =
    document.getElementById("datasetColumns");

const clearChatBtn =
    document.getElementById("clearChatBtn");

const statusText =
    document.getElementById("statusText");

const connectionStatus =
    document.getElementById("connectionStatus");

const uploadProgress =
    document.getElementById("uploadProgress");

const uploadProgressBar =
    document.getElementById("uploadProgressBar");

const resultModal =
    document.getElementById("resultModal");

const closeModalBtn =
    document.getElementById("closeModalBtn");

const modalBody =
    document.getElementById("modalBody");

/* =========================================
   INITIALIZATION
========================================= */

document.addEventListener("DOMContentLoaded", () => {
    setupEvents();
    setupVoice();
    checkHealth();
});

/* =========================================
   EVENTS
========================================= */

function setupEvents() {
    fileInput?.addEventListener(
        "change",
        handleFileUpload
    );

    attachBtn?.addEventListener(
        "click",
        () => fileInput?.click()
    );

    sendBtn?.addEventListener(
        "click",
        sendMessage
    );

    voiceBtn?.addEventListener(
        "click",
        toggleVoice
    );

    clearChatBtn?.addEventListener(
        "click",
        clearChat
    );

    closeModalBtn?.addEventListener(
        "click",
        closeModal
    );

    resultModal?.addEventListener(
        "click",
        (event) => {
            if (event.target === resultModal) {
                closeModal();
            }
        }
    );

    questionInput?.addEventListener(
        "keydown",
        handleInputKeydown
    );

    questionInput?.addEventListener(
        "input",
        autoResizeTextarea
    );

    document
        .querySelectorAll(".example-btn")
        .forEach((button) => {
            button.addEventListener(
                "click",
                () => {
                    questionInput.value =
                        button.dataset.question || "";

                    autoResizeTextarea();
                    questionInput.focus();
                }
            );
        });
}

/* =========================================
   HEALTH CHECK
========================================= */

async function checkHealth() {
    try {
        const response =
            await fetch("/api/health");

        if (!response.ok) {
            throw new Error(
                "Server unavailable"
            );
        }

        const data =
            await response.json();

        setOnlineStatus(
            data?.status === "ok"
        );
    } catch (error) {
        setOnlineStatus(false);
    }
}

function setOnlineStatus(isOnline) {
    if (!connectionStatus) {
        return;
    }

    if (isOnline) {
        connectionStatus.innerHTML =
            `<span class="status-dot"></span> Online`;

        if (statusText) {
            statusText.textContent =
                "Ready";
        }
    } else {
        connectionStatus.innerHTML =
            `<span class="status-dot" style="background:#dc2626"></span> Offline`;

        if (statusText) {
            statusText.textContent =
                "Server unavailable";
        }
    }
}

/* =========================================
   FILE UPLOAD
========================================= */

async function handleFileUpload(event) {
    const file =
        event.target.files?.[0];

    if (!file) {
        return;
    }

    const allowedExtensions = [
        ".csv",
        ".xlsx",
        ".xls",
        ".json"
    ];

    const extension =
        "." +
        file.name
            .split(".")
            .pop()
            .toLowerCase();

    if (!allowedExtensions.includes(extension)) {
        showTemporaryMessage(
            "Please upload CSV, XLSX, XLS or JSON file."
        );

        fileInput.value = "";
        return;
    }

    try {
        setUploading(true);

        const formData =
            new FormData();

        formData.append(
            "file",
            file
        );

        const response =
            await fetch("/api/upload", {
                method: "POST",
                body: formData
            });

        const data =
            await parseResponse(response);

        if (!response.ok) {
            throw new Error(
                data?.error ||
                data?.message ||
                "Upload failed."
            );
        }

        state.dataset = data;

        updateDatasetCard(data);

        clearWelcome();

        addAssistantMessage(
            buildUploadMessage(data)
        );

        setStatus(
            "Dataset loaded"
        );
    } catch (error) {
        console.error(
            "UPLOAD ERROR:",
            error
        );

        addAssistantMessage(
            `❌ ${escapeHtml(
                error.message ||
                "Unable to upload dataset."
            )}`
        );

        setStatus(
            "Upload failed"
        );
    } finally {
        setUploading(false);
        fileInput.value = "";
    }
}

function updateDatasetCard(data) {
    if (!datasetCard) {
        return;
    }

    datasetCard.classList.remove(
        "hidden"
    );

    if (datasetName) {
        datasetName.textContent =
            data?.fileName ||
            data?.filename ||
            data?.name ||
            "Dataset";
    }

    if (datasetRows) {
        datasetRows.textContent =
            formatNumber(
                data?.rows ??
                data?.rowCount ??
                0
            );
    }

    if (datasetColumns) {
        datasetColumns.textContent =
            formatNumber(
                data?.columns ??
                data?.columnCount ??
                data?.columns?.length ??
                0
            );
    }
}

function buildUploadMessage(data) {
    const fileName =
        data?.fileName ||
        data?.filename ||
        data?.name ||
        "dataset";

    const rows =
        data?.rows ??
        data?.rowCount ??
        0;

    const columns =
        data?.columns ??
        data?.columnCount ??
        data?.columns?.length ??
        0;

    return `
        <strong>I've loaded ${escapeHtml(fileName)} successfully.</strong>
        <br><br>
        I found <strong>${formatNumber(rows)}</strong> rows
        and <strong>${formatNumber(columns)}</strong> columns.
        <br><br>
        Ask me anything about your data — naturally in
        English, Tamil or Tanglish.
    `;
}

function setUploading(isUploading) {
    if (uploadProgress) {
        uploadProgress.classList.toggle(
            "hidden",
            !isUploading
        );
    }

    if (uploadProgressBar) {
        uploadProgressBar.style.width =
            isUploading ? "100%" : "0%";
    }

    if (attachBtn) {
        attachBtn.disabled =
            isUploading;
    }
}

/* =========================================
   SEND MESSAGE
========================================= */

async function sendMessage() {
    if (state.isLoading) {
        return;
    }

    const question =
        questionInput.value.trim();

    if (!question) {
        return;
    }

    clearSuggestions();

    addUserMessage(question);

    state.conversation.push({
        role: "user",
        content: question
    });

    questionInput.value = "";

    autoResizeTextarea();

    setLoading(true);

    try {
        const response =
            await fetch("/api/ask", {
                method: "POST",
                headers: {
                    "Content-Type":
                        "application/json"
                },
                body: JSON.stringify({
                    question,
                    conversation:
                        state.conversation
                            .slice(-12),
                    dataset:
                        state.dataset
                })
            });

        const data =
            await parseResponse(response);

        if (!response.ok) {
            throw new Error(
                data?.error ||
                data?.message ||
                "Something went wrong."
            );
        }

        handleAIResponse(data);
    } catch (error) {
        console.error(
            "ASK ERROR:",
            error
        );

        addAssistantMessage(
            `❌ ${escapeHtml(
                error.message ||
                "Unable to get a response."
            )}`
        );
    } finally {
        setLoading(false);
    }
}

/* =========================================
   RESPONSE HANDLING
========================================= */

function handleAIResponse(data) {
    if (!data) {
        throw new Error(
            "AI returned an empty response."
        );
    }

    const type =
        data.type || "chat";

    if (type === "dataset") {
        state.lastDatasetResponse = data;

        addDatasetMessage(data);

        if (
            Array.isArray(
                data.suggestions
            )
        ) {
            renderSuggestions(
                data.suggestions
            );
        }
    } else {
        const answer =
            data.answer ||
            data.message ||
            "AI returned an empty response.";

        addAssistantMessage(
            formatAIText(answer)
        );
    }

    state.conversation.push({
        role: "assistant",
        content:
            data.answer ||
            data.message ||
            ""
    });
}

/* =========================================
   USER MESSAGE
========================================= */

function addUserMessage(text) {
    clearWelcome();

    const wrapper =
        document.createElement("div");

    wrapper.className =
        "message user";

    wrapper.innerHTML = `
        <div class="message-avatar">
            👤
        </div>

        <div class="message-content">
            <div class="message-bubble">
                ${escapeHtml(text)}
            </div>

            <div class="message-time">
                ${currentTime()}
            </div>
        </div>
    `;

    chatContainer.appendChild(wrapper);

    scrollToBottom();
}

/* =========================================
   ASSISTANT MESSAGE
========================================= */

function addAssistantMessage(
    content
) {
    clearWelcome();

    const wrapper =
        document.createElement("div");

    wrapper.className =
        "message assistant";

    wrapper.innerHTML = `
        <div class="message-avatar">
            🤖
        </div>

        <div class="message-content">
            <div class="message-bubble">
                ${content}
            </div>

            <div class="message-time">
                ${currentTime()}
            </div>
        </div>
    `;

    chatContainer.appendChild(wrapper);

    scrollToBottom();
}

/* =========================================
   DATASET MESSAGE
========================================= */

function addDatasetMessage(data) {
    clearWelcome();

    const wrapper =
        document.createElement("div");

    wrapper.className =
        "message assistant";

    const answer =
        data.answer ||
        "Here is the verified result.";

    const sql =
        data.sql || "";

    const result =
        data.result;

    const columns =
        Array.isArray(data.columns)
            ? data.columns
            : inferColumns(result);

    const resultHTML =
        buildResultHTML(
            result,
            columns
        );

    const sqlHTML =
        sql
            ? `
                <details class="sql-details">
                    <summary>
                        🔍 View generated SQL
                    </summary>

                    <pre>${escapeHtml(sql)}</pre>
                </details>
              `
            : "";

    wrapper.innerHTML = `
        <div class="message-avatar">
            🤖
        </div>

        <div class="message-content">
            <div class="message-bubble">
                ${formatAIText(answer)}

                ${resultHTML}

                ${sqlHTML}

                ${buildReportToolsHTML()}
            </div>

            <div class="message-time">
                ${currentTime()}
            </div>
        </div>
    `;

    chatContainer.appendChild(wrapper);

    setupReportTools(
        wrapper,
        data
    );

    scrollToBottom();
}

/* =========================================
   REPORT TOOLBAR
========================================= */

function buildReportToolsHTML() {
    return `
        <div class="dd-report-tools">

            <button
                type="button"
                class="dd-report-btn"
                data-report-action="report"
            >
                📄 Generate Report
            </button>

            <button
                type="button"
                class="dd-report-btn"
                data-report-action="chart"
            >
                📊 Generate Chart
            </button>

            <button
                type="button"
                class="dd-report-btn"
                data-report-action="pie"
            >
                🍩 Generate Doughnut
            </button>

            <button
                type="button"
                class="dd-report-btn"
                data-report-action="csv"
            >
                📥 Download CSV
            </button>

        </div>
    `;
}

function setupReportTools(
    wrapper,
    data
) {
    const toolbar =
        wrapper.querySelector(
            ".dd-report-tools"
        );

    if (!toolbar) {
        return;
    }

    toolbar.addEventListener(
        "click",
        (event) => {
            const button =
                event.target.closest(
                    "[data-report-action]"
                );

            if (!button) {
                return;
            }

            const action =
                button.dataset
                    .reportAction;

            if (
                action === "report"
            ) {
                downloadTextReport(
                    data
                );
            }

            if (
                action === "chart"
            ) {
                generateChart(
                    data,
                    false
                );
            }

            if (
                action === "pie"
            ) {
                generateChart(
                    data,
                    true
                );
            }

            if (
                action === "csv"
            ) {
                downloadCSV(
                    data
                );
            }
        }
    );
}

/* =========================================
   REPORT DOWNLOAD
========================================= */

function downloadTextReport(data) {
    const rows =
        normalizeResultRows(
            data?.result
        );

    let report = "";

    report +=
        "DATA DETECTIVE AI - DATA REPORT\n";

    report +=
        "========================================\n\n";

    report +=
        "Question:\n";

    report +=
        `${data?.question || "Dataset Analysis"}\n\n`;

    report +=
        "AI Summary:\n";

    report +=
        `${data?.answer || ""}\n\n`;

    report +=
        "Verified Result:\n";

    report +=
        "========================================\n";

    if (rows.length) {
        const columns =
            Object.keys(
                rows[0]
            );

        report +=
            columns.join(" | ") +
            "\n";

        report +=
            columns
                .map(() => "---------")
                .join(" | ") +
            "\n";

        rows.forEach((row) => {
            report +=
                columns
                    .map(
                        (column) =>
                            String(
                                row?.[column] ??
                                ""
                            )
                    )
                    .join(" | ") +
                "\n";
        });
    } else {
        report +=
            "No matching rows found.\n";
    }

    report += "\n";

    if (data?.sql) {
        report +=
            "Generated SQL:\n";

        report +=
            `${data.sql}\n\n`;
    }

    report +=
        "Generated by Data Detective AI\n";

    downloadBlob(
        "data-detective-report.txt",
        report,
        "text/plain;charset=utf-8"
    );
}

/* =========================================
   CSV DOWNLOAD
========================================= */

function downloadCSV(data) {
    const rows =
        normalizeResultRows(
            data?.result
        );

    if (!rows.length) {
        showTemporaryMessage(
            "No result data available for CSV."
        );

        return;
    }

    const columns =
        Object.keys(
            rows[0]
        );

    const csv = [];

    csv.push(
        columns
            .map(csvEscape)
            .join(",")
    );

    rows.forEach((row) => {
        csv.push(
            columns
                .map(
                    (column) =>
                        csvEscape(
                            row?.[column]
                        )
                )
                .join(",")
        );
    });

    downloadBlob(
        "data-detective-result.csv",
        csv.join("\n"),
        "text/csv;charset=utf-8"
    );
}

function csvEscape(value) {
    const text =
        String(
            value ?? ""
        );

    return `"${text.replace(
        /"/g,
        '""'
    )}"`;
}

/* =========================================
   CHART GENERATION
========================================= */

function generateChart(
    data,
    doughnut = false
) {
    const rows =
        normalizeResultRows(
            data?.result
        );

    if (!rows.length) {
        showTemporaryMessage(
            "No result data available for chart."
        );

        return;
    }

    const labelColumn =
        findChartLabelColumn(
            rows
        );

    const valueColumn =
        findChartValueColumn(
            rows
        );

    if (
        !labelColumn ||
        !valueColumn
    ) {
        showTemporaryMessage(
            "This result does not contain suitable chart data."
        );

        return;
    }

    const values =
        rows
            .map((row) => ({
                label:
                    String(
                        row?.[
                            labelColumn
                        ] ?? ""
                    ),

                value:
                    toNumber(
                        row?.[
                            valueColumn
                        ]
                    )
            }))
            .filter(
                (item) =>
                    item.label !== ""
            )
            .slice(0, 20);

    if (!values.length) {
        showTemporaryMessage(
            "No numeric chart values found."
        );

        return;
    }

    const canvas =
        document.createElement(
            "canvas"
        );

    canvas.width = 1100;
    canvas.height = 650;

    drawChartCanvas(
        canvas,
        values,
        data,
        doughnut
    );

    const url =
        canvas.toDataURL(
            "image/png"
        );

    const fileName =
        doughnut
            ? "data-detective-doughnut-chart.png"
            : "data-detective-chart.png";

    const link =
        document.createElement(
            "a"
        );

    link.href = url;
    link.download = fileName;

    document.body.appendChild(link);

    link.click();

    link.remove();
}

/* =========================================
   CHART CANVAS
========================================= */

function drawChartCanvas(
    canvas,
    values,
    data,
    doughnut
) {
    const ctx =
        canvas.getContext("2d");

    const width =
        canvas.width;

    const height =
        canvas.height;

    ctx.fillStyle =
        "#ffffff";

    ctx.fillRect(
        0,
        0,
        width,
        height
    );

    ctx.fillStyle =
        "#111827";

    ctx.font =
        "bold 28px Arial";

    ctx.fillText(
        doughnut
            ? "Data Detective AI - Doughnut Chart"
            : "Data Detective AI - Chart",
        40,
        45
    );

    if (doughnut) {
        drawDoughnut(
            ctx,
            values,
            width,
            height
        );
    } else {
        drawBars(
            ctx,
            values,
            width,
            height
        );
    }

    ctx.fillStyle =
        "#6b7280";

    ctx.font =
        "14px Arial";

    ctx.fillText(
        data?.question ||
        "Dataset analysis",
        40,
        height - 20
    );
}

function drawBars(
    ctx,
    values,
    width,
    height
) {
    const left = 90;
    const right = 40;
    const top = 90;
    const bottom = 120;

    const chartWidth =
        width -
        left -
        right;

    const chartHeight =
        height -
        top -
        bottom;

    const maxValue =
        Math.max(
            ...values.map(
                (item) =>
                    item.value
            ),
            1
        );

    const slotWidth =
        chartWidth /
        values.length;

    const barWidth =
        Math.max(
            20,
            slotWidth * 0.65
        );

    values.forEach(
        (item, index) => {
            const barHeight =
                (
                    item.value /
                    maxValue
                ) *
                chartHeight;

            const x =
                left +
                index *
                    slotWidth +
                (
                    slotWidth -
                    barWidth
                ) /
                    2;

            const y =
                top +
                chartHeight -
                barHeight;

            ctx.fillStyle =
                "#2563eb";

            ctx.fillRect(
                x,
                y,
                barWidth,
                barHeight
            );

            ctx.save();

            ctx.translate(
                x +
                    barWidth /
                        2,
                height -
                    bottom +
                    25
            );

            ctx.rotate(
                -Math.PI / 4
            );

            ctx.fillStyle =
                "#111827";

            ctx.font =
                "14px Arial";

            ctx.textAlign =
                "right";

            ctx.fillText(
                truncateText(
                    item.label,
                    24
                ),
                0,
                0
            );

            ctx.restore();
        }
    );

    ctx.fillStyle =
        "#111827";

    ctx.font =
        "14px Arial";

    ctx.fillText(
        "Value",
        20,
        top
    );
}

function drawDoughnut(
    ctx,
    values,
    width,
    height
) {
    const centerX =
        width * 0.38;

    const centerY =
        height * 0.52;

    const radius = 190;

    const total =
        values.reduce(
            (sum, item) =>
                sum +
                item.value,
            0
        );

    if (!total) {
        ctx.fillStyle =
            "#111827";

        ctx.font =
            "20px Arial";

        ctx.fillText(
            "No numeric values available",
            350,
            300
        );

        return;
    }

    let startAngle = 0;

    values.forEach(
        (item, index) => {
            const slice =
                (
                    item.value /
                    total
                ) *
                Math.PI *
                2;

            ctx.beginPath();

            ctx.moveTo(
                centerX,
                centerY
            );

            ctx.arc(
                centerX,
                centerY,
                radius,
                startAngle,
                startAngle +
                    slice
            );

            ctx.closePath();

            ctx.fillStyle =
                getChartColor(
                    index
                );

            ctx.fill();

            startAngle +=
                slice;
        }
    );

    ctx.beginPath();

    ctx.arc(
        centerX,
        centerY,
        radius * 0.48,
        0,
        Math.PI * 2
    );

    ctx.fillStyle =
        "#ffffff";

    ctx.fill();

    ctx.fillStyle =
        "#111827";

    ctx.font =
        "bold 18px Arial";

    ctx.textAlign =
        "center";

    ctx.fillText(
        "Total",
        centerX,
        centerY - 5
    );

    ctx.font =
        "14px Arial";

    ctx.fillText(
        formatNumber(
            Math.round(total)
        ),
        centerX,
        centerY + 20
    );

    ctx.textAlign =
        "left";

    let legendY = 100;

    values.forEach(
        (item, index) => {
            const percentage =
                (
                    item.value /
                    total *
                    100
                ).toFixed(1);

            ctx.fillStyle =
                getChartColor(
                    index
                );

            ctx.fillRect(
                700,
                legendY - 12,
                16,
                16
            );

            ctx.fillStyle =
                "#111827";

            ctx.font =
                "14px Arial";

            ctx.fillText(
                `${truncateText(
                    item.label,
                    25
                )} (${percentage}%)`,
                725,
                legendY
            );

            legendY += 28;
        }
    );
}

function getChartColor(index) {
    const colors = [
        "#2563eb",
        "#16a34a",
        "#dc2626",
        "#9333ea",
        "#ea580c",
        "#0891b2",
        "#ca8a04",
        "#db2777",
        "#4f46e5",
        "#059669"
    ];

    return colors[
        index % colors.length
    ];
}

/* =========================================
   CHART HELPERS
========================================= */

function findChartLabelColumn(rows) {
    if (!rows.length) {
        return null;
    }

    const columns =
        Object.keys(
            rows[0]
        );

    const preferred = [
        "Year",
        "District",
        "Product",
        "Month",
        "Order_Status",
        "Payment_Mode",
        "Gender",
        "Sale_Date"
    ];

    for (
        const preferredName of preferred
    ) {
        const found =
            columns.find(
                (column) =>
                    column.toLowerCase() ===
                    preferredName.toLowerCase()
            );

        if (found) {
            return found;
        }
    }

    return columns.find(
        (column) =>
            !isColumnNumeric(
                rows,
                column
            )
    ) || columns[0];
}

function findChartValueColumn(rows) {
    if (!rows.length) {
        return null;
    }

    const columns =
        Object.keys(
            rows[0]
        );

    const preferred = [
        "Total_Sales_INR",
        "Total_Sales",
        "Sales",
        "Revenue",
        "Quantity",
        "Count"
    ];

    for (
        const preferredName of preferred
    ) {
        const found =
            columns.find(
                (column) =>
                    column.toLowerCase() ===
                    preferredName.toLowerCase()
            );

        if (
            found &&
            isColumnNumeric(
                rows,
                found
            )
        ) {
            return found;
        }
    }

    return columns.find(
        (column) =>
            isColumnNumeric(
                rows,
                column
            )
    );
}

function isColumnNumeric(
    rows,
    column
) {
    return rows.some((row) => {
        const value =
            toNumber(
                row?.[column]
            );

        return Number.isFinite(
            value
        );
    });
}

function toNumber(value) {
    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {
        return NaN;
    }

    const cleaned =
        String(value)
            .replace(
                /[^0-9.-]/g,
                ""
            );

    if (!cleaned) {
        return NaN;
    }

    const number =
        Number(cleaned);

    return Number.isFinite(number)
        ? number
        : NaN;
}

function truncateText(
    text,
    length
) {
    const value =
        String(text ?? "");

    if (
        value.length <=
        length
    ) {
        return value;
    }

    return (
        value.slice(
            0,
            length - 3
        ) +
        "..."
    );
}

/* =========================================
   RESULT TABLE
========================================= */

function buildResultHTML(
    result,
    columns
) {
    if (
        result === null ||
        result === undefined
    ) {
        return "";
    }

    let rows = [];

    if (Array.isArray(result)) {
        rows = result;
    } else if (
        typeof result === "object"
    ) {
        rows = [result];
    } else {
        return `
            <div class="result-card">
                <div class="result-header">
                    📊 Verified Result
                </div>

                <div style="padding:14px;">
                    ${escapeHtml(
                        String(result)
                    )}
                </div>
            </div>
        `;
    }

    if (!rows.length) {
        return `
            <div class="result-card">
                <div class="result-header">
                    📊 Verified Result
                </div>

                <div style="padding:14px;color:#6b7280;">
                    No matching rows found.
                </div>
            </div>
        `;
    }

    const safeColumns =
        columns?.length
            ? columns
            : Object.keys(
                rows[0] || {}
            );

    const header =
        safeColumns
            .map(
                (column) =>
                    `<th>${escapeHtml(
                        String(column)
                    )}</th>`
            )
            .join("");

    const body =
        rows
            .slice(0, 100)
            .map((row) => {
                const cells =
                    safeColumns
                        .map((column) => {
                            const value =
                                row?.[column];

                            return `
                                <td>
                                    ${escapeHtml(
                                        formatCell(
                                            value
                                        )
                                    )}
                                </td>
                            `;
                        })
                        .join("");

                return `<tr>${cells}</tr>`;
            })
            .join("");

    return `
        <div class="result-card">

            <div class="result-header">
                <span>
                    📊 Verified Result
                </span>

                <span>
                    ${rows.length}
                    row${rows.length === 1 ? "" : "s"}
                </span>
            </div>

            <div class="result-table-wrapper">

                <table class="result-table">

                    <thead>
                        <tr>
                            ${header}
                        </tr>
                    </thead>

                    <tbody>
                        ${body}
                    </tbody>

                </table>

            </div>

        </div>
    `;
}

/* =========================================
   SUGGESTIONS
========================================= */

function renderSuggestions(items) {
    if (
        !suggestions ||
        !Array.isArray(items) ||
        !items.length
    ) {
        return;
    }

    suggestions.innerHTML = "";

    items
        .filter(Boolean)
        .slice(0, 5)
        .forEach((item) => {
            const button =
                document.createElement(
                    "button"
                );

            button.type =
                "button";

            button.className =
                "suggestion-btn";

            button.textContent =
                String(item);

            button.addEventListener(
                "click",
                () => {
                    questionInput.value =
                        String(item);

                    autoResizeTextarea();

                    questionInput.focus();

                    sendMessage();
                }
            );

            suggestions.appendChild(
                button
            );
        });

    suggestions.classList.remove(
        "hidden"
    );
}

function clearSuggestions() {
    if (!suggestions) {
        return;
    }

    suggestions.innerHTML = "";

    suggestions.classList.add(
        "hidden"
    );
}

/* =========================================
   LOADING
========================================= */

function setLoading(
    loading
) {
    state.isLoading =
        loading;

    if (typingIndicator) {
        typingIndicator.classList.toggle(
            "hidden",
            !loading
        );
    }

    if (sendBtn) {
        sendBtn.disabled =
            loading;
    }

    if (questionInput) {
        questionInput.disabled =
            loading;
    }

    if (loading) {
        scrollToBottom();
    }
}

/* =========================================
   VOICE INPUT
========================================= */

function setupVoice() {
    const SpeechRecognition =
        window.SpeechRecognition ||
        window.webkitSpeechRecognition;

    if (!SpeechRecognition) {
        if (voiceBtn) {
            voiceBtn.title =
                "Voice input is not supported in this browser";

            voiceBtn.style.opacity =
                "0.5";
        }

        return;
    }

    state.recognition =
        new SpeechRecognition();

    state.recognition.continuous =
        false;

    state.recognition.interimResults =
        true;

    state.recognition.lang =
        "en-IN";

    state.recognition.onstart =
        () => {
            state.isListening =
                true;

            voiceBtn?.classList.add(
                "voice-active"
            );
        };

    state.recognition.onresult =
        (event) => {
            let transcript = "";

            for (
                let i =
                    event.resultIndex;
                i <
                event.results.length;
                i++
            ) {
                transcript +=
                    event.results[i][0]
                        .transcript;
            }

            questionInput.value =
                transcript;

            autoResizeTextarea();
        };

    state.recognition.onend =
        () => {
            state.isListening =
                false;

            voiceBtn?.classList.remove(
                "voice-active"
            );
        };

    state.recognition.onerror =
        (event) => {
            console.warn(
                "VOICE ERROR:",
                event.error
            );

            state.isListening =
                false;

            voiceBtn?.classList.remove(
                "voice-active"
            );
        };
}

function toggleVoice() {
    if (!state.recognition) {
        showTemporaryMessage(
            "Voice input is not supported in this browser."
        );

        return;
    }

    if (state.isListening) {
        state.recognition.stop();
    } else {
        try {
            state.recognition.start();
        } catch (error) {
            console.warn(
                "Voice start error:",
                error
            );
        }
    }
}

/* =========================================
   CLEAR CHAT
========================================= */

function clearChat() {
    state.conversation = [];

    state.lastDatasetResponse =
        null;

    if (chatContainer) {
        chatContainer.innerHTML = "";

        if (welcomeMessage) {
            chatContainer.appendChild(
                welcomeMessage
            );

            welcomeMessage.classList.remove(
                "hidden"
            );
        }
    }

    clearSuggestions();

    setStatus(
        state.dataset
            ? "Dataset loaded"
            : "Ready"
    );

    questionInput?.focus();
}

/* =========================================
   MODAL
========================================= */

function closeModal() {
    resultModal?.classList.add(
        "hidden"
    );
}

/* =========================================
   INPUT
========================================= */

function handleInputKeydown(
    event
) {
    if (
        event.key === "Enter" &&
        !event.shiftKey
    ) {
        event.preventDefault();

        sendMessage();
    }
}

function autoResizeTextarea() {
    if (!questionInput) {
        return;
    }

    questionInput.style.height =
        "auto";

    questionInput.style.height =
        Math.min(
            questionInput.scrollHeight,
            150
        ) + "px";
}

/* =========================================
   HELPERS
========================================= */

async function parseResponse(
    response
) {
    const text =
        await response.text();

    if (!text) {
        return {};
    }

    try {
        return JSON.parse(text);
    } catch {
        return {
            message: text
        };
    }
}

function clearWelcome() {
    if (welcomeMessage) {
        welcomeMessage.classList.add(
            "hidden"
        );
    }
}

function scrollToBottom() {
    requestAnimationFrame(() => {
        if (chatContainer) {
            chatContainer.scrollTop =
                chatContainer.scrollHeight;
        }
    });
}

function setStatus(text) {
    if (statusText) {
        statusText.textContent =
            text;
    }
}

function showTemporaryMessage(
    message
) {
    const element =
        document.createElement(
            "div"
        );

    element.className =
        "message assistant";

    element.innerHTML = `
        <div class="message-avatar">
            🤖
        </div>

        <div class="message-content">
            <div class="message-bubble">
                ${escapeHtml(message)}
            </div>
        </div>
    `;

    chatContainer.appendChild(
        element
    );

    scrollToBottom();

    setTimeout(() => {
        element.remove();
    }, 4000);
}

function formatAIText(text) {
    if (
        text === null ||
        text === undefined
    ) {
        return "";
    }

    let value =
        String(text);

    value =
        escapeHtml(value);

    value =
        value.replace(
            /\*\*(.*?)\*\*/g,
            "<strong>$1</strong>"
        );

    value =
        value.replace(
            /`([^`]+)`/g,
            "<code>$1</code>"
        );

    value =
        value.replace(
            /\n/g,
            "<br>"
        );

    return value;
}

function escapeHtml(value) {
    if (
        value === null ||
        value === undefined
    ) {
        return "";
    }

    return String(value)
        .replace(
            /&/g,
            "&amp;"
        )
        .replace(
            /</g,
            "&lt;"
        )
        .replace(
            />/g,
            "&gt;"
        )
        .replace(
            /"/g,
            "&quot;"
        )
        .replace(
            /'/g,
            "&#039;"
        );
}

function formatCell(value) {
    if (
        value === null ||
        value === undefined
    ) {
        return "";
    }

    if (
        typeof value === "object"
    ) {
        try {
            return JSON.stringify(
                value
            );
        } catch {
            return String(value);
        }
    }

    return String(value);
}

function inferColumns(result) {
    if (
        !Array.isArray(result) ||
        !result.length
    ) {
        return [];
    }

    return Object.keys(
        result[0] || {}
    );
}

function formatNumber(value) {
    const number =
        Number(value);

    if (
        Number.isFinite(number)
    ) {
        return number.toLocaleString();
    }

    return String(
        value ?? 0
    );
}

function currentTime() {
    return new Date().toLocaleTimeString(
        [],
        {
            hour: "2-digit",
            minute: "2-digit"
        }
    );
}

/* =========================================
   REPORT HELPERS
========================================= */

function normalizeResultRows(
    result
) {
    if (Array.isArray(result)) {
        return result;
    }

    if (
        result &&
        typeof result === "object"
    ) {
        return [result];
    }

    return [];
}

function downloadBlob(
    filename,
    content,
    type
) {
    const blob =
        new Blob(
            [content],
            {
                type
            }
        );

    const url =
        URL.createObjectURL(
            blob
        );

    const link =
        document.createElement(
            "a"
        );

    link.href = url;

    link.download =
        filename;

    document.body.appendChild(
        link
    );

    link.click();

    link.remove();

    URL.revokeObjectURL(
        url
    );
}