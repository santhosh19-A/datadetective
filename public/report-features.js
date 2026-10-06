/* ============================================================
   DATA DETECTIVE AI - EXTRA REPORT FEATURES
   This file is ADDITIVE ONLY.
   It does not replace existing chat / SQL / upload logic.
============================================================ */

(function () {
    "use strict";

    let lastDatasetResponse = null;

    function escapeHTML(value) {
        return String(value ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    function downloadFile(
        filename,
        content,
        type
    ) {
        const blob = new Blob(
            [content],
            { type }
        );

        const url =
            URL.createObjectURL(blob);

        const a =
            document.createElement("a");

        a.href = url;
        a.download = filename;

        document.body.appendChild(a);
        a.click();
        a.remove();

        URL.revokeObjectURL(url);
    }

    function getNumericColumns(rows) {
        if (!rows || !rows.length) {
            return [];
        }

        const first = rows[0];

        return Object.keys(first).filter(
            (column) => {
                return rows.some((row) => {
                    const value =
                        String(
                            row[column] ?? ""
                        )
                            .replace(
                                /[^0-9.-]/g,
                                ""
                            );

                    return (
                        value !== "" &&
                        !Number.isNaN(
                            Number(value)
                        )
                    );
                });
            }
        );
    }

    function numberValue(value) {
        if (
            value === null ||
            value === undefined
        ) {
            return 0;
        }

        const cleaned =
            String(value)
                .replace(
                    /[^0-9.-]/g,
                    ""
                );

        const number =
            Number(cleaned);

        return Number.isFinite(number)
            ? number
            : 0;
    }

    function findLabelColumn(rows) {
        if (!rows.length) {
            return null;
        }

        const columns =
            Object.keys(rows[0]);

        const preferred = [
            "District",
            "Product",
            "Year",
            "Sale_Date",
            "Order_Status",
            "Payment_Mode",
            "Gender",
            "Month"
        ];

        for (
            const name of preferred
        ) {
            const found =
                columns.find(
                    (column) =>
                        column.toLowerCase() ===
                        name.toLowerCase()
                );

            if (found) {
                return found;
            }
        }

        return columns[0];
    }

    function findValueColumn(rows) {
        const numeric =
            getNumericColumns(rows);

        if (!numeric.length) {
            return null;
        }

        const preferred = [
            "Total_Sales",
            "Total_Sales_INR",
            "Sales",
            "Revenue",
            "Quantity",
            "Count"
        ];

        for (
            const name of preferred
        ) {
            const found =
                numeric.find(
                    (column) =>
                        column.toLowerCase() ===
                        name.toLowerCase()
                );

            if (found) {
                return found;
            }
        }

        return numeric[0];
    }

    function generateReport() {
        if (!lastDatasetResponse) {
            alert(
                "First ask a dataset question."
            );
            return;
        }

        const data =
            lastDatasetResponse;

        const rows =
            Array.isArray(data.result)
                ? data.result
                : [];

        let report = "";

        report +=
            "DATA DETECTIVE AI - SALES REPORT\n";

        report +=
            "========================================\n\n";

        report +=
            "Question:\n";

        report +=
            `${data.question || "Dataset analysis"}\n\n`;

        report +=
            "AI Summary:\n";

        report +=
            `${data.answer || ""}\n\n`;

        report +=
            "Verified Result:\n";

        report +=
            "========================================\n";

        if (rows.length) {
            const columns =
                Object.keys(rows[0]);

            report +=
                columns.join(" | ") +
                "\n";

            report +=
                columns
                    .map(() => "---------")
                    .join(" | ") +
                "\n";

            rows.forEach(
                (row) => {
                    report +=
                        columns
                            .map(
                                (column) =>
                                    String(
                                        row[column] ??
                                            ""
                                    )
                            )
                            .join(" | ") +
                        "\n";
                }
            );
        } else {
            report +=
                "No rows returned.\n";
        }

        report += "\n";

        report +=
            "SQL Used:\n";

        report +=
            `${data.sql || ""}\n\n`;

        report +=
            "Generated by Data Detective AI\n";

        downloadFile(
            "data-detective-report.txt",
            report,
            "text/plain;charset=utf-8"
        );
    }

    function generateCSV() {
        if (!lastDatasetResponse) {
            alert(
                "First ask a dataset question."
            );
            return;
        }

        const rows =
            Array.isArray(
                lastDatasetResponse.result
            )
                ? lastDatasetResponse.result
                : [];

        if (!rows.length) {
            alert(
                "No result data available."
            );
            return;
        }

        const columns =
            Object.keys(rows[0]);

        const escapeCSV = (value) => {
            const text =
                String(
                    value ?? ""
                );

            return `"${text.replace(
                /"/g,
                '""'
            )}"`;
        };

        const csv = [];

        csv.push(
            columns
                .map(escapeCSV)
                .join(",")
        );

        rows.forEach((row) => {
            csv.push(
                columns
                    .map(
                        (column) =>
                            escapeCSV(
                                row[column]
                            )
                    )
                    .join(",")
            );
        });

        downloadFile(
            "data-detective-result.csv",
            csv.join("\n"),
            "text/csv;charset=utf-8"
        );
    }

    function createChartCanvas(
        title,
        pie
    ) {
        const rows =
            Array.isArray(
                lastDatasetResponse?.result
            )
                ? lastDatasetResponse.result
                : [];

        if (!rows.length) {
            alert(
                "No result data available for chart."
            );
            return;
        }

        const labelColumn =
            findLabelColumn(rows);

        const valueColumn =
            findValueColumn(rows);

        if (
            !labelColumn ||
            !valueColumn
        ) {
            alert(
                "This result does not contain suitable chart data."
            );
            return;
        }

        const width = 1000;
        const height = 600;

        const canvas =
            document.createElement(
                "canvas"
            );

        canvas.width = width;
        canvas.height = height;

        const ctx =
            canvas.getContext("2d");

        ctx.fillStyle = "#ffffff";
        ctx.fillRect(
            0,
            0,
            width,
            height
        );

        ctx.fillStyle = "#111827";
        ctx.font =
            "bold 26px Arial";

        ctx.fillText(
            title,
            40,
            45
        );

        const values =
            rows
                .map((row) => ({
                    label:
                        String(
                            row[
                                labelColumn
                            ] ?? ""
                        ),
                    value:
                        numberValue(
                            row[
                                valueColumn
                            ]
                        )
                }))
                .filter(
                    (item) =>
                        item.label
                )
                .slice(0, 20);

        if (pie) {
            drawPieChart(
                ctx,
                values,
                width,
                height
            );
        } else {
            drawBarChart(
                ctx,
                values,
                width,
                height
            );
        }

        const url =
            canvas.toDataURL(
                "image/png"
            );

        const a =
            document.createElement("a");

        a.href = url;

        a.download =
            pie
                ? "data-detective-pie-chart.png"
                : "data-detective-chart.png";

        document.body.appendChild(a);
        a.click();
        a.remove();
    }

    function drawBarChart(
        ctx,
        values,
        width,
        height
    ) {
        const left = 90;
        const bottom = 80;
        const top = 90;

        const chartWidth =
            width - left - 40;

        const chartHeight =
            height -
            top -
            bottom;

        const max =
            Math.max(
                ...values.map(
                    (item) =>
                        item.value
                ),
                1
            );

        const barWidth =
            chartWidth /
            Math.max(
                values.length,
                1
            ) *
            0.65;

        values.forEach(
            (item, index) => {
                const x =
                    left +
                    index *
                        (chartWidth /
                            values.length) +
                    10;

                const barHeight =
                    (item.value /
                        max) *
                    chartHeight;

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
                    height - 25
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
                    item.label,
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

    function drawPieChart(
        ctx,
        values,
        width,
        height
    ) {
        const centerX =
            width / 2;

        const centerY =
            height / 2 + 20;

        const radius = 180;

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
                300,
                300
            );

            return;
        }

        let startAngle = 0;

        values.forEach(
            (item, index) => {
                const slice =
                    (item.value /
                        total) *
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
                    `hsl(${
                        (index * 55) %
                        360
                    }, 70%, 55%)`;

                ctx.fill();

                startAngle += slice;
            }
        );

        ctx.fillStyle =
            "#111827";

        ctx.font =
            "14px Arial";

        let legendY = 90;

        values.forEach(
            (item, index) => {
                const percentage =
                    (
                        (item.value /
                            total) *
                        100
                    ).toFixed(1);

                ctx.fillStyle =
                    `hsl(${
                        (index * 55) %
                        360
                    }, 70%, 55%)`;

                ctx.fillRect(
                    700,
                    legendY - 12,
                    15,
                    15
                );

                ctx.fillStyle =
                    "#111827";

                ctx.fillText(
                    `${item.label} (${percentage}%)`,
                    725,
                    legendY
                );

                legendY += 25;
            }
        );
    }

    function addReportButtons(container) {
        if (
            !container ||
            container.querySelector(
                ".dd-report-tools"
            )
        ) {
            return;
        }

        const box =
            document.createElement(
                "div"
            );

        box.className =
            "dd-report-tools";

        box.innerHTML = `
            <button type="button" data-report-action="report">
                📄 Generate Report
            </button>

            <button type="button" data-report-action="chart">
                📊 Generate Chart
            </button>

            <button type="button" data-report-action="pie">
                🥧 Generate Pie Chart
            </button>

            <button type="button" data-report-action="csv">
                📥 Download CSV
            </button>
        `;

        box.addEventListener(
            "click",
            function (event) {
                const button =
                    event.target.closest(
                        "button"
                    );

                if (!button) {
                    return;
                }

                const action =
                    button.dataset
                        .reportAction;

                if (
                    action ===
                    "report"
                ) {
                    generateReport();
                }

                if (
                    action ===
                    "chart"
                ) {
                    createChartCanvas(
                        "Data Detective AI - Chart",
                        false
                    );
                }

                if (
                    action ===
                    "pie"
                ) {
                    createChartCanvas(
                        "Data Detective AI - Pie Chart",
                        true
                    );
                }

                if (
                    action ===
                    "csv"
                ) {
                    generateCSV();
                }
            }
        );

        container.appendChild(box);
    }

    window.DataDetectiveReports = {
        setResult(data) {
            lastDatasetResponse =
                data;

            setTimeout(
                () => {
                    const datasetMessages =
                        document.querySelectorAll(
                            ".message.assistant"
                        );

                    const last =
                        datasetMessages[
                            datasetMessages.length -
                                1
                        ];

                    if (last) {
                        addReportButtons(
                            last
                        );
                    }
                },
                50
            );
        },

        addButtons(container) {
            addReportButtons(
                container
            );
        }
    };

})();