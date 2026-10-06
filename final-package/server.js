const express = require("express");
const cors = require("cors");
const OpenAI = require("openai");
const multer = require("multer");
const csvParser = require("csv-parser");
const XLSX = require("xlsx");
const duckdb = require("duckdb");
const fs = require("fs");
const path = require("path");

require("dotenv").config();

const app = express();

const PORT = process.env.PORT || 5000;

const MODEL =
    process.env.OPENROUTER_MODEL ||
    "openrouter/free";

const UPLOAD_DIR =
    path.join(__dirname, "uploads");

const OUTPUT_DIR =
    path.join(__dirname, "outputs");


// ============================================================
// DIRECTORIES
// ============================================================

if (!fs.existsSync(UPLOAD_DIR)) {
    fs.mkdirSync(UPLOAD_DIR, {
        recursive: true,
    });
}

if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR, {
        recursive: true,
    });
}


// ============================================================
// OPENROUTER
// ============================================================

const ai = new OpenAI({
    apiKey: process.env.OPENROUTER_API_KEY,
    baseURL: "https://openrouter.ai/api/v1",
});


// ============================================================
// DUCKDB
// ============================================================

const db = new duckdb.Database(":memory:");

let datasetLoaded = false;

let currentDataset = null;

let lastAnalysis = null;


// ============================================================
// MIDDLEWARE
// ============================================================

app.use(cors());

app.use(
    express.json({
        limit: "10mb",
    })
);

app.use(
    express.static(
        path.join(__dirname, "public")
    )
);

app.use(
    "/outputs",
    express.static(OUTPUT_DIR)
);


// ============================================================
// UPLOAD
// ============================================================

const upload = multer({
    dest: UPLOAD_DIR,

    limits: {
        fileSize: 100 * 1024 * 1024,
    },
});


// ============================================================
// SAFE JSON
// ============================================================

function makeJSONSafe(value) {
    if (typeof value === "bigint") {
        return Number(value);
    }

    if (value instanceof Date) {
        return value.toISOString();
    }

    if (Array.isArray(value)) {
        return value.map(makeJSONSafe);
    }

    if (
        value &&
        typeof value === "object"
    ) {
        const result = {};

        for (const [key, val] of Object.entries(value)) {
            result[key] = makeJSONSafe(val);
        }

        return result;
    }

    return value;
}


// ============================================================
// DUCKDB HELPERS
// ============================================================

function runQuery(sql) {
    return new Promise(
        (resolve, reject) => {
            db.all(
                sql,
                (error, rows) => {
                    if (error) {
                        reject(error);
                        return;
                    }

                    resolve(
                        makeJSONSafe(
                            rows || []
                        )
                    );
                }
            );
        }
    );
}


function runStatement(sql) {
    return new Promise(
        (resolve, reject) => {
            db.run(
                sql,
                (error) => {
                    if (error) {
                        reject(error);
                        return;
                    }

                    resolve();
                }
            );
        }
    );
}


// ============================================================
// TEXT HELPERS
// ============================================================

function normalizeQuestion(question) {
    return String(question || "")
        .trim()
        .toLowerCase()
        .replace(/\s+/g, " ");
}


function cleanColumnName(
    name,
    index
) {
    let value =
        String(name ?? "").trim();

    if (!value) {
        value =
            `Column_${index + 1}`;
    }

    return value;
}


function uniqueHeaders(headers) {
    const used = new Map();

    return headers.map(
        (header, index) => {
            const base =
                cleanColumnName(
                    header,
                    index
                );

            if (!used.has(base)) {
                used.set(base, 1);
                return base;
            }

            const count =
                used.get(base) + 1;

            used.set(
                base,
                count
            );

            return `${base}_${count}`;
        }
    );
}


function normalizeCell(value) {
    if (
        value === null ||
        value === undefined
    ) {
        return "";
    }

    if (value instanceof Date) {
        return value.toISOString();
    }

    return String(value).trim();
}


function csvEscape(value) {
    const text =
        normalizeCell(value);

    if (
        text.includes(",") ||
        text.includes('"') ||
        text.includes("\n") ||
        text.includes("\r")
    ) {
        return `"${text.replace(
            /"/g,
            '""'
        )}"`;
    }

    return text;
}


function rowsToCsv(
    headers,
    rows
) {
    const lines = [];

    lines.push(
        headers
            .map(csvEscape)
            .join(",")
    );

    for (const row of rows) {
        lines.push(
            headers
                .map(
                    (header) =>
                        csvEscape(
                            row[header]
                        )
                )
                .join(",")
        );
    }

    return lines.join("\n");
}


function normalizeRows(
    headers,
    rows
) {
    return rows.map(
        (row) => {
            const result = {};

            headers.forEach(
                (header) => {
                    result[header] =
                        normalizeCell(
                            row[header]
                        );
                }
            );

            return result;
        }
    );
}


// ============================================================
// EXCEL HEADER DETECTION
// ============================================================

function scoreExcelHeaderRow(
    row
) {
    const values =
        row
            .map((value) =>
                String(
                    value ?? ""
                )
                    .trim()
                    .toLowerCase()
            )
            .filter(Boolean);

    if (values.length < 2) {
        return 0;
    }

    let score = 0;

    if (
        values.some(
            (v) =>
                /department|dept|branch|course|program|stream/.test(
                    v
                )
        )
    ) {
        score += 10;
    }

    if (
        values.some(
            (v) =>
                /year|batch|academic/.test(
                    v
                )
        )
    ) {
        score += 10;
    }

    if (
        values.some(
            (v) =>
                /placement|placed|eligible|student|percentage|percent/.test(
                    v
                )
        )
    ) {
        score += 10;
    }

    if (values.length >= 4) {
        score += 2;
    }

    return score;
}


function detectExcelHeader(
    rawRows
) {
    let bestIndex = -1;

    let bestScore = 0;

    for (
        let i = 0;
        i < rawRows.length;
        i++
    ) {
        const score =
            scoreExcelHeaderRow(
                rawRows[i]
            );

        if (
            score > bestScore
        ) {
            bestScore = score;

            bestIndex = i;
        }
    }

    if (bestIndex === -1) {
        for (
            let i = 0;
            i < rawRows.length;
            i++
        ) {
            const count =
                rawRows[i].filter(
                    (v) =>
                        String(
                            v ?? ""
                        ).trim() !== ""
                ).length;

            if (count >= 2) {
                bestIndex = i;
                break;
            }
        }
    }

    return bestIndex;
}


// ============================================================
// CSV
// ============================================================

function parseCSVFile(
    filePath
) {
    return new Promise(
        (resolve, reject) => {
            const rows = [];

            fs.createReadStream(
                filePath
            )
                .pipe(
                    csvParser({
                        mapHeaders:
                            ({
                                header,
                                index,
                            }) =>
                                cleanColumnName(
                                    header,
                                    index
                                ),
                    })
                )
                .on(
                    "data",
                    (row) => {
                        rows.push(row);
                    }
                )
                .on(
                    "end",
                    () => {
                        const headers =
                            rows.length
                                ? Object.keys(
                                      rows[0]
                                  )
                                : [];

                        resolve({
                            headers,

                            rows:
                                normalizeRows(
                                    headers,
                                    rows
                                ),
                        });
                    }
                )
                .on(
                    "error",
                    reject
                );
        }
    );
}


// ============================================================
// EXCEL
// ============================================================

function parseExcelFile(
    filePath
) {
    const workbook =
        XLSX.readFile(
            filePath,
            {
                cellDates: true,
            }
        );

    const sheetName =
        workbook.SheetNames[0];

    if (!sheetName) {
        throw new Error(
            "Excel file has no worksheet."
        );
    }

    const sheet =
        workbook.Sheets[
            sheetName
        ];

    const rawRows =
        XLSX.utils.sheet_to_json(
            sheet,
            {
                header: 1,
                defval: "",
                raw: true,
            }
        );

    if (!rawRows.length) {
        return {
            headers: [],
            rows: [],
        };
    }

    const headerIndex =
        detectExcelHeader(
            rawRows
        );

    if (
        headerIndex === -1
    ) {
        throw new Error(
            "Unable to detect table header."
        );
    }

    const headers =
        uniqueHeaders(
            rawRows[
                headerIndex
            ].map(
                (
                    value,
                    index
                ) =>
                    cleanColumnName(
                        value,
                        index
                    )
            )
        );

    const rows = [];

    for (
        let i =
            headerIndex + 1;
        i < rawRows.length;
        i++
    ) {
        const rawRow =
            rawRows[i];

        const hasData =
            rawRow.some(
                (value) =>
                    String(
                        value ?? ""
                    ).trim() !== ""
            );

        if (!hasData) {
            continue;
        }

        const row = {};

        headers.forEach(
            (
                header,
                index
            ) => {
                row[header] =
                    normalizeCell(
                        rawRow[index]
                    );
            }
        );

        rows.push(row);
    }

    return {
        headers,
        rows,
        sheetName,
        headerIndex,
    };
}


// ============================================================
// JSON
// ============================================================

function parseJSONFile(
    filePath
) {
    const content =
        fs.readFileSync(
            filePath,
            "utf8"
        );

    const parsed =
        JSON.parse(content);

    let data = parsed;

    if (!Array.isArray(data)) {
        if (
            Array.isArray(
                parsed.data
            )
        ) {
            data =
                parsed.data;
        } else if (
            Array.isArray(
                parsed.rows
            )
        ) {
            data =
                parsed.rows;
        } else {
            throw new Error(
                "JSON must contain an array of objects."
            );
        }
    }

    if (!data.length) {
        return {
            headers: [],
            rows: [],
        };
    }

    const headerSet =
        new Set();

    data.forEach(
        (item) => {
            if (
                item &&
                typeof item ===
                    "object" &&
                !Array.isArray(
                    item
                )
            ) {
                Object.keys(
                    item
                ).forEach(
                    (key) =>
                        headerSet.add(
                            key
                        )
                );
            }
        }
    );

    const headers =
        uniqueHeaders(
            Array.from(
                headerSet
            )
        );

    const rows =
        data.map(
            (item) => {
                const row = {};

                headers.forEach(
                    (header) => {
                        row[header] =
                            normalizeCell(
                                item?.[
                                    header
                                ]
                            );
                    }
                );

                return row;
            }
        );

    return {
        headers,
        rows,
    };
}


// ============================================================
// PROFILE
// ============================================================

function profileColumn(
    rows,
    column
) {
    const values =
        rows.map(
            (row) =>
                row[column]
        );

    const nonEmpty =
        values.filter(
            (value) =>
                String(
                    value ?? ""
                ).trim() !== ""
        );

    const numericValues =
        nonEmpty
            .map(
                (value) => {
                    const n =
                        Number(
                            String(
                                value
                            )
                                .replace(
                                    /,/g,
                                    ""
                                )
                                .replace(
                                    /₹/g,
                                    ""
                                )
                                .replace(
                                    /%/g,
                                    ""
                                )
                        );

                    return Number.isFinite(
                        n
                    )
                        ? n
                        : null;
                }
            )
            .filter(
                (v) =>
                    v !== null
            );

    const isNumeric =
        nonEmpty.length >
            0 &&
        numericValues.length ===
            nonEmpty.length;

    const unique =
        new Set(
            nonEmpty.map(
                (v) =>
                    String(v)
            )
        );

    const result = {
        name: column,

        type: isNumeric
            ? "numeric"
            : "text",

        total:
            values.length,

        missing:
            values.length -
            nonEmpty.length,

        unique:
            unique.size,

        samples:
            Array.from(
                unique
            ).slice(0, 10),
    };

    if (isNumeric) {
        result.min =
            Math.min(
                ...numericValues
            );

        result.max =
            Math.max(
                ...numericValues
            );

        result.average =
            Number(
                (
                    numericValues.reduce(
                        (a, b) =>
                            a + b,
                        0
                    ) /
                    numericValues.length
                ).toFixed(2)
            );
    }

    return result;
}


function createSchemaProfile(
    headers,
    rows
) {
    return headers.map(
        (column) =>
            profileColumn(
                rows,
                column
            )
    );
}


// ============================================================
// LOAD INTO DUCKDB
// ============================================================

async function loadRowsIntoDuckDB(
    headers,
    rows
) {
    if (!headers.length) {
        throw new Error(
            "No usable columns found."
        );
    }

    const csvPath =
        path.join(
            UPLOAD_DIR,
            `dataset_${Date.now()}.csv`
        );

    fs.writeFileSync(
        csvPath,
        rowsToCsv(
            headers,
            rows
        ),
        "utf8"
    );

    const safePath =
        csvPath
            .replace(
                /\\/g,
                "/"
            )
            .replace(
                /'/g,
                "''"
            );

    await runStatement(
        `DROP TABLE IF EXISTS dataset`
    );

    await runStatement(
        `
        CREATE TABLE dataset AS
        SELECT *
        FROM read_csv(
            '${safePath}',
            HEADER = TRUE,
            ALL_VARCHAR = TRUE,
            IGNORE_ERRORS = TRUE
        )
        `
    );

    return csvPath;
}


// ============================================================
// COLUMN HELPERS
// ============================================================

function quoteIdentifier(
    column
) {
    return `"${String(
        column
    ).replace(
        /"/g,
        '""'
    )}"`;
}


function findColumn(
    columns,
    keywords
) {
    const lowered =
        columns.map(
            (column) => ({
                original:
                    column,

                lower:
                    String(
                        column
                    ).toLowerCase(),
            })
        );

    // Exact
    for (
        const keyword of keywords
    ) {
        const exact =
            lowered.find(
                (item) =>
                    item.lower ===
                    keyword
            );

        if (exact) {
            return exact.original;
        }
    }

    // Partial
    for (
        const keyword of keywords
    ) {
        const partial =
            lowered.find(
                (item) =>
                    item.lower.includes(
                        keyword
                    )
            );

        if (partial) {
            return partial.original;
        }
    }

    return null;
}


// ============================================================
// DETECT SEMANTIC COLUMNS
// ============================================================

function detectColumns(
    schema
) {
    const columns =
        schema.map(
            (c) => c.name
        );

    return {
        columns,

        year:
            findColumn(
                columns,
                [
                    "academic year",
                    "placement year",
                    "passout year",
                    "graduation year",
                    "batch year",
                    "year",
                    "batch",
                ]
            ),

        department:
            findColumn(
                columns,
                [
                    "department",
                    "dept",
                    "branch",
                    "stream",
                    "course",
                    "program",
                ]
            ),

        placementPercentage:
            findColumn(
                columns,
                [
                    "placement percentage",
                    "placement %",
                    "placement percent",
                    "placed percentage",
                    "placement rate",
                    "percentage",
                    "percent",
                ]
            ),

        eligible:
            findColumn(
                columns,
                [
                    "eligible students",
                    "students eligible",
                    "eligible count",
                    "eligible",
                ]
            ),

        placed:
            findColumn(
                columns,
                [
                    "students placed",
                    "student placed",
                    "placed students",
                    "students selected",
                    "selected students",
                    "placed",
                ]
            ),

        totalStudents:
            findColumn(
                columns,
                [
                    "total students",
                    "total student",
                    "student count",
                    "students",
                    "total",
                ]
            ),
    };
}


// ============================================================
// NUMERIC SQL
// ============================================================

function numericExpression(
    column
) {
    return `
        TRY_CAST(
            regexp_replace(
                CAST(
                    ${quoteIdentifier(
                        column
                    )}
                    AS VARCHAR
                ),
                '[^0-9.-]',
                '',
                'g'
            ) AS DOUBLE
        )
    `;
}


// ============================================================
// QUESTION HELPERS
// ============================================================

function extractYear(
    question,
    columns
) {
    const match =
        String(
            question
        ).match(
            /\b(19|20)\d{2}\b/
        );

    if (!match) {
        return null;
    }

    const year =
        match[0];

    const yearColumn =
        findColumn(
            columns,
            [
                "academic year",
                "placement year",
                "graduation year",
                "passout year",
                "batch year",
                "year",
                "batch",
            ]
        );

    if (!yearColumn) {
        return null;
    }

    return {
        year,
        column:
            yearColumn,
    };
}


function buildYearFilter(
    yearInfo
) {
    if (!yearInfo) {
        return "";
    }

    const column =
        quoteIdentifier(
            yearInfo.column
        );

    const year =
        yearInfo.year;

    return `
        AND (
            CAST(
                ${column}
                AS VARCHAR
            ) = '${year}'

            OR

            regexp_matches(
                CAST(
                    ${column}
                    AS VARCHAR
                ),
                '(^|[^0-9])${year}([^0-9]|$)'
            )
        )
    `;
}


// ============================================================
// DEPARTMENT DETECTION
// ============================================================

function extractDepartment(
    question,
    schema
) {
    const detected =
        detectColumns(
            schema
        );

    if (!detected.department) {
        return null;
    }

    const departmentValues =
        [
            ...new Set(
                currentDataset.rows
                    .map(
                        (row) =>
                            String(
                                row[
                                    detected
                                        .department
                                ] ??
                                    ""
                            ).trim()
                    )
                    .filter(
                        Boolean
                    )
            ),
        ];

    const q =
        normalizeQuestion(
            question
        );

    // Exact
    for (
        const value of departmentValues
    ) {
        if (
            q.includes(
                value.toLowerCase()
            )
        ) {
            return value;
        }
    }

    // Normalized compact match
    const compactQuestion =
        q.replace(
            /[^a-z0-9&]+/g,
            ""
        );

    for (
        const value of departmentValues
    ) {
        const compactValue =
            value
                .toLowerCase()
                .replace(
                    /[^a-z0-9&]+/g,
                    ""
                );

        if (
            compactValue &&
            compactQuestion.includes(
                compactValue
            )
        ) {
            return value;
        }
    }

    // Word match
    for (
        const value of departmentValues
    ) {
        const words =
            value
                .toLowerCase()
                .split(/\s+/)
                .filter(
                    (word) =>
                        word.length >=
                        3
                );

        if (
            words.some(
                (word) =>
                    q.includes(
                        word
                    )
            )
        ) {
            return value;
        }
    }

    return null;
}


// ============================================================
// PREVIOUS YEAR
// ============================================================

function findPreviousYear(
    question,
    schema,
    conversation
) {
    const q =
        normalizeQuestion(
            question
        );

    if (
        !q.includes(
            "previous year"
        ) &&
        !q.includes(
            "last year"
        ) &&
        !q.includes(
            "prior year"
        ) &&
        !q.includes(
            "previous"
        )
    ) {
        return null;
    }

    const detected =
        detectColumns(
            schema
        );

    if (!detected.year) {
        return null;
    }

    const allYears =
        [
            ...new Set(
                currentDataset.rows
                    .map(
                        (row) =>
                            String(
                                row[
                                    detected
                                        .year
                                ] ??
                                    ""
                            )
                    )
                    .map(
                        (value) => {
                            const match =
                                value.match(
                                    /(19|20)\d{2}/
                                );

                            return match
                                ? Number(
                                      match[0]
                                  )
                                : null;
                        }
                    )
                    .filter(
                        (year) =>
                            year !==
                            null
                    )
            ),
        ].sort(
            (a, b) => a - b
        );

    if (!allYears.length) {
        return null;
    }

    let mentionedYear =
        null;

    const history =
        Array.isArray(
            conversation
        )
            ? conversation
                  .slice(-10)
                  .map(
                      (m) =>
                          String(
                              m.content ||
                                  ""
                          )
                  )
                  .join(" ")
            : "";

    const matches =
        history.match(
            /\b(19|20)\d{2}\b/g
        );

    if (
        matches &&
        matches.length
    ) {
        mentionedYear =
            Number(
                matches[
                    matches.length -
                        1
                ]
            );
    }

    if (
        mentionedYear !== null
    ) {
        const previous =
            allYears.filter(
                (year) =>
                    year <
                    mentionedYear
            );

        if (previous.length) {
            return previous[
                previous.length - 1
            ];
        }
    }

    return null;
}


// ============================================================
// INTENT DETECTION
// ============================================================

function hasAny(
    q,
    words
) {
    return words.some(
        (word) =>
            q.includes(word)
    );
}


function detectIntent(
    question
) {
    const q =
        normalizeQuestion(
            question
        );

    return {
        eligible:
            hasAny(q, [
                "eligible",
                "eligibility",
            ]),

        placed:
            hasAny(q, [
                "placed",
                "students placed",
                "placed students",
                "selected",
            ]),

        totalStudents:
            hasAny(q, [
                "total students",
                "total student",
                "student strength",
                "student count",
                "how many students",
            ]),

        percentage:
            hasAny(q, [
                "percentage",
                "percent",
                "%",
                "placement rate",
                "placement ratio",
            ]),

        average:
            hasAny(q, [
                "average",
                "avg",
                "mean",
                "overall",
            ]),

        highest:
            hasAny(q, [
                "highest",
                "maximum",
                "max",
                "best",
                "top",
                "greatest",
                "largest",
            ]),

        lowest:
            hasAny(q, [
                "lowest",
                "minimum",
                "min",
                "worst",
                "least",
                "smallest",
            ]),

        compare:
            hasAny(q, [
                "compare",
                "comparison",
                "compared",
                "difference",
                "change",
                "growth",
                "increase",
                "decrease",
                "year over year",
                "yoy",
            ]),

        trend:
            hasAny(q, [
                "trend",
                "over the years",
                "year wise",
                "year-wise",
                "yearwise",
                "history",
            ]),

        why:
            hasAny(q, [
                "why",
                "reason",
                "reasons",
                "cause",
                "caused",
            ]),

        explain:
            hasAny(q, [
                "explain",
                "interpret",
                "meaning",
            ]),
    };
}


// ============================================================
// DATA QUESTION DETECTION
// ============================================================

function isDataQuestion(
    question
) {
    const q =
        normalizeQuestion(
            question
        );

    if (
        !currentDataset
    ) {
        return false;
    }

    const columns =
        currentDataset.headers
            .map(
                (x) =>
                    String(
                        x
                    ).toLowerCase()
            );

    const department =
        extractDepartment(
            q,
            currentDataset.schema
        );

    const year =
        /\b(19|20)\d{2}\b/.test(
            q
        );

    const dataWords = [
        "data",
        "dataset",
        "excel",
        "department",
        "dept",
        "placement",
        "placed",
        "eligible",
        "students",
        "percentage",
        "percent",
        "year",
        "highest",
        "lowest",
        "average",
        "total",
        "count",
        "compare",
        "comparison",
        "trend",
        "record",
        "rate",
        "how many",
        "which department",
        "which year",
    ];

    const columnMention =
        columns.some(
            (column) =>
                q.includes(column)
        );

    return (
        Boolean(department) ||
        year ||
        columnMention ||
        hasAny(
            q,
            dataWords
        )
    );
}


// ============================================================
// DETERMINISTIC SQL ENGINE
// ============================================================

function buildDeterministicSQL(
    question,
    schema,
    conversation = []
) {
    const q =
        normalizeQuestion(
            question
        );

    const detected =
        detectColumns(
            schema
        );

    const {
        department,
        year,
        placementPercentage,
        eligible,
        placed,
        totalStudents,
    } = detected;

    if (
        !detected.columns.length
    ) {
        throw new Error(
            "Dataset schema is empty."
        );
    }

    const intent =
        detectIntent(
            question
        );

    const requestedDepartment =
        extractDepartment(
            question,
            schema
        );

    const yearInfo =
        extractYear(
            q,
            detected.columns
        );

    let yearFilter =
        buildYearFilter(
            yearInfo
        );

    const previousYear =
        findPreviousYear(
            question,
            schema,
            conversation
        );

    if (
        previousYear &&
        year
    ) {
        yearFilter = `
            AND (
                CAST(
                    ${quoteIdentifier(
                        year
                    )}
                    AS VARCHAR
                ) = '${previousYear}'

                OR

                regexp_matches(
                    CAST(
                        ${quoteIdentifier(
                            year
                        )}
                        AS VARCHAR
                    ),
                    '(^|[^0-9])${previousYear}([^0-9]|$)'
                )
            )
        `;
    }


    // ========================================================
    // DEPARTMENT FILTER
    // ========================================================

    const departmentFilter =
        requestedDepartment &&
        department
            ? `
                AND LOWER(
                    CAST(
                        ${quoteIdentifier(
                            department
                        )}
                        AS VARCHAR
                    )
                ) =
                LOWER(
                    '${requestedDepartment.replace(
                        /'/g,
                        "''"
                    )}'
                )
            `
            : "";


    // ========================================================
    // WHY / EXPLAIN / TREND
    // ========================================================

    if (
        requestedDepartment &&
        department &&
        year &&
        placementPercentage &&
        (
            intent.why ||
            intent.explain ||
            intent.trend ||
            intent.compare
        )
    ) {
        return `
            SELECT
                ${quoteIdentifier(
                    department
                )} AS "Department",

                ${quoteIdentifier(
                    year
                )} AS "Year",

                ROUND(
                    AVG(
                        ${numericExpression(
                            placementPercentage
                        )}
                    ),
                    2
                ) AS "Placement Percentage"

            FROM dataset

            WHERE 1 = 1

            ${departmentFilter}

            GROUP BY
                ${quoteIdentifier(
                    department
                )},

                ${quoteIdentifier(
                    year
                )}

            ORDER BY
                TRY_CAST(
                    regexp_replace(
                        CAST(
                            ${quoteIdentifier(
                                year
                            )}
                            AS VARCHAR
                        ),
                        '[^0-9.-]',
                        '',
                        'g'
                    ) AS DOUBLE
                )
        `;
    }


    // ========================================================
    // HIGHEST / LOWEST YEAR FOR DEPARTMENT
    // ========================================================

    if (
        requestedDepartment &&
        department &&
        placementPercentage &&
        year &&
        (
            intent.highest ||
            intent.lowest
        )
    ) {
        const order =
            intent.lowest
                ? "ASC"
                : "DESC";

        return `
            SELECT
                ${quoteIdentifier(
                    department
                )} AS "Department",

                ${quoteIdentifier(
                    year
                )} AS "Year",

                ROUND(
                    AVG(
                        ${numericExpression(
                            placementPercentage
                        )}
                    ),
                    2
                ) AS "Placement Percentage"

            FROM dataset

            WHERE 1 = 1

            ${departmentFilter}

            GROUP BY
                ${quoteIdentifier(
                    department
                )},

                ${quoteIdentifier(
                    year
                )}

            ORDER BY
                "Placement Percentage"
                ${order}

            LIMIT 1
        `;
    }


    // ========================================================
    // BOTH HIGHEST + LOWEST
    // ========================================================

    if (
        requestedDepartment &&
        department &&
        placementPercentage &&
        year &&
        intent.highest &&
        intent.lowest
    ) {
        return `
            WITH yearly AS (
                SELECT
                    ${quoteIdentifier(
                        department
                    )} AS "Department",

                    ${quoteIdentifier(
                        year
                    )} AS "Year",

                    ROUND(
                        AVG(
                            ${numericExpression(
                                placementPercentage
                            )}
                        ),
                        2
                    ) AS "Placement Percentage"

                FROM dataset

                WHERE 1 = 1

                ${departmentFilter}

                GROUP BY
                    ${quoteIdentifier(
                        department
                    )},

                    ${quoteIdentifier(
                        year
                    )}
            ),

            ranked AS (
                SELECT
                    *,

                    ROW_NUMBER() OVER (
                        ORDER BY
                            "Placement Percentage"
                            DESC
                    ) AS highest_rank,

                    ROW_NUMBER() OVER (
                        ORDER BY
                            "Placement Percentage"
                            ASC
                    ) AS lowest_rank

                FROM yearly
            )

            SELECT
                "Department",
                "Year",
                "Placement Percentage",

                CASE
                    WHEN highest_rank = 1
                    THEN 'Highest'

                    WHEN lowest_rank = 1
                    THEN 'Lowest'

                    ELSE ''
                END AS "Result Type"

            FROM ranked

            WHERE
                highest_rank = 1
                OR lowest_rank = 1

            ORDER BY
                "Placement Percentage"
                DESC
        `;
    }


    // ========================================================
    // ELIGIBLE STUDENTS
    // ========================================================

    if (
        intent.eligible &&
        eligible
    ) {
        if (
            requestedDepartment &&
            department &&
            yearInfo &&
            year
        ) {
            return `
                SELECT
                    ${quoteIdentifier(
                        department
                    )} AS "Department",

                    ${quoteIdentifier(
                        year
                    )} AS "Year",

                    SUM(
                        ${numericExpression(
                            eligible
                        )}
                    ) AS "Eligible Students"

                FROM dataset

                WHERE 1 = 1

                ${departmentFilter}

                ${yearFilter}

                GROUP BY
                    ${quoteIdentifier(
                        department
                    )},

                    ${quoteIdentifier(
                        year
                    )}
            `;
        }

        if (
            requestedDepartment &&
            department
        ) {
            return `
                SELECT
                    ${quoteIdentifier(
                        department
                    )} AS "Department",

                    SUM(
                        ${numericExpression(
                            eligible
                        )}
                    ) AS "Eligible Students"

                FROM dataset

                WHERE 1 = 1

                ${departmentFilter}

                ${yearFilter}

                GROUP BY
                    ${quoteIdentifier(
                        department
                    )}
            `;
        }

        if (
            yearInfo &&
            year
        ) {
            return `
                SELECT
                    ${quoteIdentifier(
                        year
                    )} AS "Year",

                    SUM(
                        ${numericExpression(
                            eligible
                        )}
                    ) AS "Eligible Students"

                FROM dataset

                WHERE 1 = 1

                ${yearFilter}

                GROUP BY
                    ${quoteIdentifier(
                        year
                    )}
            `;
        }

        return `
            SELECT
                SUM(
                    ${numericExpression(
                        eligible
                    )}
                ) AS "Eligible Students"

            FROM dataset
        `;
    }


    // ========================================================
    // STUDENTS PLACED
    // ========================================================

    if (
        intent.placed &&
        placed
    ) {
        if (
            requestedDepartment &&
            department &&
            yearInfo &&
            year
        ) {
            return `
                SELECT
                    ${quoteIdentifier(
                        department
                    )} AS "Department",

                    ${quoteIdentifier(
                        year
                    )} AS "Year",

                    SUM(
                        ${numericExpression(
                            placed
                        )}
                    ) AS "Students Placed"

                FROM dataset

                WHERE 1 = 1

                ${departmentFilter}

                ${yearFilter}

                GROUP BY
                    ${quoteIdentifier(
                        department
                    )},

                    ${quoteIdentifier(
                        year
                    )}
            `;
        }

        if (
            requestedDepartment &&
            department
        ) {
            return `
                SELECT
                    ${quoteIdentifier(
                        department
                    )} AS "Department",

                    SUM(
                        ${numericExpression(
                            placed
                        )}
                    ) AS "Students Placed"

                FROM dataset

                WHERE 1 = 1

                ${departmentFilter}

                ${yearFilter}

                GROUP BY
                    ${quoteIdentifier(
                        department
                    )}
            `;
        }

        if (
            yearInfo &&
            year
        ) {
            return `
                SELECT
                    ${quoteIdentifier(
                        year
                    )} AS "Year",

                    SUM(
                        ${numericExpression(
                            placed
                        )}
                    ) AS "Students Placed"

                FROM dataset

                WHERE 1 = 1

                ${yearFilter}

                GROUP BY
                    ${quoteIdentifier(
                        year
                    )}
            `;
        }

        return `
            SELECT
                SUM(
                    ${numericExpression(
                        placed
                    )}
                ) AS "Students Placed"

            FROM dataset
        `;
    }


    // ========================================================
    // TOTAL STUDENTS
    // ========================================================

    if (
        intent.totalStudents &&
        totalStudents
    ) {
        if (
            requestedDepartment &&
            department &&
            yearInfo &&
            year
        ) {
            return `
                SELECT
                    ${quoteIdentifier(
                        department
                    )} AS "Department",

                    ${quoteIdentifier(
                        year
                    )} AS "Year",

                    SUM(
                        ${numericExpression(
                            totalStudents
                        )}
                    ) AS "Total Students"

                FROM dataset

                WHERE 1 = 1

                ${departmentFilter}

                ${yearFilter}

                GROUP BY
                    ${quoteIdentifier(
                        department
                    )},

                    ${quoteIdentifier(
                        year
                    )}
            `;
        }

        if (
            requestedDepartment &&
            department
        ) {
            return `
                SELECT
                    ${quoteIdentifier(
                        department
                    )} AS "Department",

                    SUM(
                        ${numericExpression(
                            totalStudents
                        )}
                    ) AS "Total Students"

                FROM dataset

                WHERE 1 = 1

                ${departmentFilter}

                ${yearFilter}

                GROUP BY
                    ${quoteIdentifier(
                        department
                    )}
            `;
        }

        if (
            yearInfo &&
            year
        ) {
            return `
                SELECT
                    ${quoteIdentifier(
                        year
                    )} AS "Year",

                    SUM(
                        ${numericExpression(
                            totalStudents
                        )}
                    ) AS "Total Students"

                FROM dataset

                WHERE 1 = 1

                ${yearFilter}

                GROUP BY
                    ${quoteIdentifier(
                        year
                    )}
            `;
        }

        return `
            SELECT
                SUM(
                    ${numericExpression(
                        totalStudents
                    )}
                ) AS "Total Students"

            FROM dataset
        `;
    }


    // ========================================================
    // AVERAGE PLACEMENT
    // ========================================================

    if (
        intent.average &&
        placementPercentage
    ) {
        if (
            requestedDepartment &&
            department
        ) {
            return `
                SELECT
                    ${quoteIdentifier(
                        department
                    )} AS "Department",

                    ROUND(
                        AVG(
                            ${numericExpression(
                                placementPercentage
                            )}
                        ),
                        2
                    ) AS "Average Placement Percentage"

                FROM dataset

                WHERE 1 = 1

                ${departmentFilter}

                ${yearFilter}

                GROUP BY
                    ${quoteIdentifier(
                        department
                    )}
            `;
        }

        return `
            SELECT
                ROUND(
                    AVG(
                        ${numericExpression(
                            placementPercentage
                        )}
                    ),
                    2
                ) AS "Average Placement Percentage"

            FROM dataset

            WHERE 1 = 1

            ${yearFilter}
        `;
    }


    // ========================================================
    // HIGHEST / LOWEST DEPARTMENT
    // ========================================================

    if (
        placementPercentage &&
        department &&
        (
            intent.highest ||
            intent.lowest
        )
    ) {
        const order =
            intent.lowest
                ? "ASC"
                : "DESC";

        return `
            SELECT
                ${quoteIdentifier(
                    department
                )} AS "Department",

                ROUND(
                    AVG(
                        ${numericExpression(
                            placementPercentage
                        )}
                    ),
                    2
                ) AS "Placement Percentage"

            FROM dataset

            WHERE 1 = 1

            ${yearFilter}

            GROUP BY
                ${quoteIdentifier(
                    department
                )}

            ORDER BY
                "Placement Percentage"
                ${order}

            LIMIT 1
        `;
    }


    // ========================================================
    // HIGHEST / LOWEST YEAR
    // ========================================================

    if (
        placementPercentage &&
        year &&
        (
            intent.highest ||
            intent.lowest
        )
    ) {
        const order =
            intent.lowest
                ? "ASC"
                : "DESC";

        return `
            SELECT
                ${quoteIdentifier(
                    year
                )} AS "Year",

                ROUND(
                    AVG(
                        ${numericExpression(
                            placementPercentage
                        )}
                    ),
                    2
                ) AS "Placement Percentage"

            FROM dataset

            WHERE 1 = 1

            ${departmentFilter}

            ${yearFilter}

            GROUP BY
                ${quoteIdentifier(
                    year
                )}

            ORDER BY
                "Placement Percentage"
                ${order}

            LIMIT 1
        `;
    }


    // ========================================================
    // DEPARTMENT + YEAR NORMAL RECORD
    // ========================================================

    if (
        requestedDepartment &&
        department &&
        yearInfo
    ) {
        return `
            SELECT
                *

            FROM dataset

            WHERE 1 = 1

            ${departmentFilter}

            ${yearFilter}
        `;
    }


    // ========================================================
    // DEPARTMENT NORMAL
    // ========================================================

    if (
        requestedDepartment &&
        department
    ) {
        return `
            SELECT
                *

            FROM dataset

            WHERE 1 = 1

            ${departmentFilter}

            ${yearFilter}

            ORDER BY
                TRY_CAST(
                    regexp_replace(
                        CAST(
                            ${quoteIdentifier(
                                year
                            )}
                            AS VARCHAR
                        ),
                        '[^0-9.-]',
                        '',
                        'g'
                    ) AS DOUBLE
                )
        `;
    }


    // ========================================================
    // YEAR NORMAL
    // ========================================================

    if (
        yearInfo
    ) {
        return `
            SELECT
                *

            FROM dataset

            WHERE 1 = 1

            ${yearFilter}

            LIMIT 100
        `;
    }


    // ========================================================
    // UNSUPPORTED / AMBIGUOUS
    // ========================================================

    throw new Error(
        "No deterministic SQL plan is available for this question."
    );
}


// ============================================================
// AI CORE - SINGLE SEMANTIC BRAIN
// ============================================================

function historyText(conversation = []) {
    return Array.isArray(conversation)
        ? conversation.slice(-12).map(m => `${m.role}: ${String(m.content || "")}`).join("\n")
        : "";
}

function fullDatasetContext() {
    if (!currentDataset) return "No dataset is loaded.";
    return [
        `File: ${currentDataset.fileName}`,
        `Rows: ${currentDataset.rows.length}`,
        `Columns: ${currentDataset.headers.join(", ")}`,
        "Schema:",
        ...currentDataset.schema.map(c =>
            `- "${c.name}" type=${c.type} samples=${JSON.stringify(c.samples)}`
        )
    ].join("\n");
}

function extractAIContent(response) {
    const choice = response?.choices?.[0];
    if (!choice) throw new Error("AI returned no choices.");

    let content = choice?.message?.content;

    if (Array.isArray(content)) {
        content = content.map(part => {
            if (typeof part === "string") return part;
            return part?.text || part?.content || "";
        }).join("");
    }

    if (!content && typeof choice?.text === "string") {
        content = choice.text;
    }

    if (!content && typeof choice?.message?.reasoning === "string") {
        content = choice.message.reasoning;
    }

    if (!content || !String(content).trim()) {
        console.error("OPENROUTER EMPTY RESPONSE:", JSON.stringify(response, null, 2));
        throw new Error("AI returned an empty response.");
    }

    return String(content).trim();
}

async function callAI(messages, options = {}) {
    if (!process.env.OPENROUTER_API_KEY) {
        throw new Error("OPENROUTER_API_KEY is missing.");
    }

    let lastError = null;

    for (let attempt = 1; attempt <= 2; attempt++) {
        try {
            const response = await ai.chat.completions.create({
                model: MODEL,
                messages,
                temperature: options.temperature ?? 0.2,
                max_tokens: options.max_tokens ?? 1200,
            });

            return extractAIContent(response);
        } catch (error) {
            lastError = error;
            console.error(`AI attempt ${attempt} failed:`, error.message);
            if (attempt === 1) {
                await new Promise(resolve => setTimeout(resolve, 400));
            }
        }
    }

    throw lastError || new Error("AI request failed.");
}

function parseAIJson(text) {
    const cleaned = String(text)
        .replace(/^```json\s*/i, "")
        .replace(/```$/i, "")
        .trim();

    try {
        return JSON.parse(cleaned);
    } catch (_) {
        const start = cleaned.indexOf("{");
        const end = cleaned.lastIndexOf("}");
        if (start >= 0 && end > start) {
            return JSON.parse(cleaned.slice(start, end + 1));
        }
        throw new Error("AI returned invalid JSON.");
    }
}

async function classifyQuestion(question, conversation = []) {
    const text = await callAI([
        {
            role: "system",
            content: `You are the semantic router for Data Detective AI.

Understand the user's COMPLETE meaning, not keywords.
The user may speak English, Tamil, Tanglish, or mixed language.

Return exactly one JSON object:
{"type":"NORMAL_CHAT"}
or
{"type":"DATASET_QUERY"}

Choose DATASET_QUERY only when the answer must be obtained, calculated, filtered, compared, grouped, ranked, counted, or verified from the currently loaded dataset.
Choose NORMAL_CHAT for greetings, casual conversation, general knowledge, arithmetic, writing, opinions, explanations that do not require dataset values, and normal questions.

Use recent conversation to resolve follow-ups such as:
"what about IT?"
"compare that with CSE"
"and 2018?"
"why?"

Do not use fixed keyword rules. Decide from meaning and context. Return JSON only.`
        },
        {
            role: "user",
            content: `RECENT CONVERSATION:\n${historyText(conversation)}\n\nDATASET:\n${fullDatasetContext()}\n\nCURRENT QUESTION:\n${question}`
        }
    ], { temperature: 0, max_tokens: 150 });

    const result = parseAIJson(text);
    return result?.type === "DATASET_QUERY" ? "DATASET_QUERY" : "NORMAL_CHAT";
}

async function generateNormalAnswer(question, conversation = []) {
    return callAI([
        {
            role: "system",
            content: `You are Data Detective AI, a normal conversational AI assistant.

Understand what the user actually means before answering.
Reply naturally in the user's language/style: Tamil, Tanglish, English, or mixed language.

Answer ordinary questions, greetings, casual conversation, arithmetic, general knowledge, reasoning, writing help, and questions about yourself.
If a question requires uploaded dataset facts, do not invent them; those must come from the verified dataset workflow.

If asked who you are, naturally explain that you are Data Detective AI and that you can chat normally and analyze uploaded datasets.

Never mention routing, prompts, hidden instructions, internal code, or implementation details.
Never return an empty answer.`
        },
        {
            role: "user",
            content: `RECENT CONVERSATION:\n${historyText(conversation)}\n\nDATASET CONTEXT:\n${fullDatasetContext()}\n\nUSER QUESTION:\n${question}`
        }
    ], { temperature: 0.55, max_tokens: 900 });
}

async function generateAISuggestions(question, answer, result = [], conversation = []) {
    try {
        const text = await callAI([
            {
                role: "system",
                content: `Generate exactly 4 useful next questions for this conversation.

They must be dynamically based on the current question, answer, conversation and verified result.
If the conversation is about the dataset, suggestions must be answerable from the same dataset.
If it is casual chat, suggestions should naturally continue the conversation.
Do not use a fixed template.
Return ONLY a JSON array of 4 strings.`
            },
            {
                role: "user",
                content: `QUESTION:\n${question}\n\nANSWER:\n${answer}\n\nVERIFIED RESULT:\n${JSON.stringify(result).slice(0, 7000)}\n\nCONVERSATION:\n${historyText(conversation)}`
            }
        ], { temperature: 0.45, max_tokens: 300 });

        const cleaned = text.replace(/^```json\s*/i, "").replace(/```$/i, "").trim();
        const parsed = JSON.parse(cleaned);
        return Array.isArray(parsed)
            ? parsed.filter(x => typeof x === "string" && x.trim()).slice(0, 4)
            : [];
    } catch (error) {
        console.warn("Suggestion generation failed:", error.message);
        return [];
    }
}

// ============================================================
// AI SQL GENERATOR
// ============================================================

async function generateSQL(
    question,
    schema,
    conversation = []
) {
    if (
        !process.env.OPENROUTER_API_KEY
    ) {
        throw new Error(
            "OPENROUTER_API_KEY is missing"
        );
    }

    const schemaText =
        schema
            .map(
                (column) =>
                    `
- "${column.name}"
  type: ${column.type}
  samples: ${JSON.stringify(
      column.samples
  )}
`
            )
            .join("\n");

    const conversationText =
        Array.isArray(
            conversation
        )
            ? conversation
                  .slice(-10)
                  .map(
                      (message) =>
                          `${message.role}: ${message.content}`
                  )
                  .join("\n")
            : "";

    const prompt = `
You are the SQL reasoning engine for a real data analysis application.

DATABASE TABLE:
dataset

SCHEMA:
${schemaText}

RECENT CONVERSATION:
${conversationText || "(none)"}

USER QUESTION:
${question}

Your task is ONLY to generate SQL.

RULES:

1. Use ONLY columns present in the schema.

2. Never invent columns.

3. All columns are VARCHAR.

4. Numeric calculations must use:

TRY_CAST(
    regexp_replace(
        CAST("COLUMN" AS VARCHAR),
        '[^0-9.-]',
        '',
        'g'
    ) AS DOUBLE
)

5. Column names with spaces must use double quotes.

6. Understand the question semantically.

7. If user asks for eligible students,
   use the eligible students column.

8. If user asks for students placed,
   use the students placed column.

9. If user asks total students,
   use total students column.

10. If user asks placement percentage,
    use placement percentage column if present.

11. If user asks average placement percentage,
    use AVG().

12. If user asks highest/lowest,
    aggregate at the correct level and ORDER BY the calculated value.

13. If department is mentioned,
    identify the actual department value from samples.

14. If a year is mentioned,
    filter the actual year column.

15. If user asks year-wise,
    GROUP BY year.

16. If user asks comparison,
    return the actual values needed for comparison.

17. If user asks why/explain a result,
    return the relevant actual data required to explain it.
    Never invent causal reasons.

18. Never use SELECT * LIMIT 100 when an aggregate answer is required.

19. Return only SELECT or WITH queries.

20. Never modify data.

21. Return ONLY SQL.
`;

    const content = await callAI(
        [
            {
                role: "system",
                content: "Return only valid DuckDB SELECT/WITH SQL. No markdown."
            },
            {
                role: "user",
                content: prompt
            }
        ],
        {
            temperature: 0,
            max_tokens: 1400
        }
    );

    return cleanSQL(content);
}


// ============================================================
// SQL CLEANING
// ============================================================

function cleanSQL(
    text
) {
    let sql =
        String(
            text || ""
        ).trim();

    sql = sql
        .replace(
            /^```sql/i,
            ""
        )
        .replace(
            /^```/i,
            ""
        )
        .replace(
            /```$/i,
            ""
        )
        .trim();

    const lower =
        sql.toLowerCase();

    const selectIndex =
        lower.indexOf(
            "select"
        );

    const withIndex =
        lower.indexOf(
            "with"
        );

    let startIndex = -1;

    if (
        selectIndex !== -1 &&
        withIndex !== -1
    ) {
        startIndex =
            Math.min(
                selectIndex,
                withIndex
            );
    } else if (
        selectIndex !== -1
    ) {
        startIndex =
            selectIndex;
    } else if (
        withIndex !== -1
    ) {
        startIndex =
            withIndex;
    }

    if (
        startIndex > 0
    ) {
        sql =
            sql.substring(
                startIndex
            );
    }

    sql =
        sql.replace(
            /```[\s\S]*$/g,
            ""
        );

    const semicolon =
        sql.indexOf(";");

    if (
        semicolon !== -1
    ) {
        sql =
            sql.substring(
                0,
                semicolon
            );
    }

    return sql.trim();
}


// ============================================================
// SQL VALIDATION
// ============================================================

function validateSQL(
    sql
) {
    const normalized =
        String(
            sql || ""
        )
            .trim()
            .toLowerCase();

    if (!normalized) {
        throw new Error(
            "AI did not generate SQL."
        );
    }

    if (
        !normalized.startsWith(
            "select"
        ) &&
        !normalized.startsWith(
            "with"
        )
    ) {
        throw new Error(
            "Only SELECT queries are allowed."
        );
    }

    const blocked = [
        "drop ",
        "delete ",
        "update ",
        "insert ",
        "alter ",
        "create ",
        "truncate ",
        "attach ",
        "detach ",
        "copy ",
        "install ",
        "load ",
        "pragma ",
    ];

    for (
        const keyword of blocked
    ) {
        if (
            normalized.includes(
                keyword
            )
        ) {
            throw new Error(
                "Unsafe SQL query blocked."
            );
        }
    }

    return true;
}


// ============================================================
// EVIDENCE ANSWER
// ============================================================

function formatEvidenceAnswer(
    question,
    rows,
    sql
) {
    if (
        !rows ||
        rows.length === 0
    ) {
        return {
            answer:
                "Dataset-la indha question-ku matching data kidaikala.",

            result: [],

            sql,
        };
    }

    const first =
        rows[0];

    const keys =
        Object.keys(
            first
        );


    // --------------------------------------------------------
    // SINGLE VALUE
    // --------------------------------------------------------

    if (
        rows.length === 1 &&
        keys.length === 1
    ) {
        const key =
            keys[0];

        const value =
            first[key];

        return {
            answer:
                `${key}: ${value}`,

            result: rows,

            sql,
        };
    }


    // --------------------------------------------------------
    // HIGHEST / LOWEST TWO RESULTS
    // --------------------------------------------------------

    if (
        rows.length === 2 &&
        rows.every(
            (row) =>
                Object.prototype.hasOwnProperty.call(
                    row,
                    "Result Type"
                )
        )
    ) {
        const highest =
            rows.find(
                (row) =>
                    row[
                        "Result Type"
                    ] ===
                    "Highest"
            );

        const lowest =
            rows.find(
                (row) =>
                    row[
                        "Result Type"
                    ] ===
                    "Lowest"
            );

        let text = "";

        if (highest) {
            text +=
                `Highest: ${formatRow(
                    highest
                )}`;
        }

        if (lowest) {
            text +=
                `\nLowest: ${formatRow(
                    lowest
                )}`;
        }

        return {
            answer:
                text.trim(),

            result: rows,

            sql,
        };
    }


    // --------------------------------------------------------
    // SINGLE RECORD
    // --------------------------------------------------------

    if (
        rows.length === 1
    ) {
        return {
            answer:
                formatRow(
                    first
                ),

            result: rows,

            sql,
        };
    }


    // --------------------------------------------------------
    // MULTIPLE RECORDS
    // --------------------------------------------------------

    const maxRows =
        Math.min(
            rows.length,
            20
        );

    const lines = [];

    for (
        let i = 0;
        i < maxRows;
        i++
    ) {
        lines.push(
            `${i + 1}. ${formatRow(
                rows[i]
            )}`
        );
    }

    return {
        answer:
            `${rows.length} matching records found.\n\n${lines.join(
                "\n"
            )}`,

        result: rows,

        sql,
    };
}


function formatRow(
    row
) {
    return Object.entries(
        row
    )
        .map(
            ([key, value]) =>
                `${key}: ${value}`
        )
        .join(" | ");
}


// ============================================================
// LOCAL EXPLANATION
// ============================================================

function explainVerifiedResult(
    previous
) {
    if (
        !previous ||
        !previous.result ||
        !previous.result.length
    ) {
        return "Previous result available illa.";
    }

    const rows =
        previous.result;

    const first =
        rows[0];

    const keys =
        Object.keys(
            first
        );

    // Percentage explanation
    const percentageKey =
        keys.find(
            (key) =>
                /percentage|percent|rate/i.test(
                    key
                )
        );

    const departmentKey =
        keys.find(
            (key) =>
                /department|dept|branch/i.test(
                    key
                )
        );

    const yearKey =
        keys.find(
            (key) =>
                /^year$/i.test(
                    key
                )
        );

    if (
        percentageKey
    ) {
        let text =
            "Sure nanba. Previous result-la verify aana actual data idhu:\n\n";

        for (
            const row of rows.slice(
                0,
                20
            )
        ) {
            const dept =
                departmentKey
                    ? row[
                          departmentKey
                      ]
                    : "";

            const year =
                yearKey
                    ? row[
                          yearKey
                      ]
                    : "";

            const percentage =
                row[
                    percentageKey
                ];

            text +=
                `${dept ? dept + " " : ""}` +
                `${year ? "(" + year + ") " : ""}` +
                `placement percentage = ${percentage}%.\n`;
        }

        text +=
            "\nIndha explanation uploaded dataset-la irundhu verify aana values based on pannadhu.";

        return text;
    }

    return (
        "Sure nanba. Previous verified result:\n\n" +
        rows
            .slice(0, 20)
            .map(
                (row) =>
                    formatRow(
                        row
                    )
            )
            .join("\n")
    );
}


// ============================================================
// LOCAL SUGGESTIONS
// ============================================================

function generateLocalSuggestions(
    question,
    result
) {
    return [
        "How did this compare with the previous year?",
        "Which department had the lowest placement percentage?",
        "What was the overall placement percentage?",
        "Can you explain this result?",
    ];
}


// ============================================================
// DATASET CONTEXT
// ============================================================

function buildDatasetContext() {
    if (
        !currentDataset
    ) {
        return "No dataset loaded.";
    }

    return `
File:
${currentDataset.fileName}

Rows:
${currentDataset.rows.length}

Columns:
${currentDataset.headers.join(
    ", "
)}
`;
}


// ============================================================
// NORMAL CHAT COMPATIBILITY WRAPPER
// ============================================================

async function generateCasualReply(question, conversation = []) {
    return generateNormalAnswer(question, conversation);
}

// ============================================================
// RESULT EXCEL
// ============================================================

function createResultExcel({
    question,
    result,
}) {
    const safeResult =
        Array.isArray(
            result
        )
            ? result
            : [];

    const workbook =
        XLSX.utils.book_new();

    let resultSheet;

    if (
        safeResult.length
    ) {
        resultSheet =
            XLSX.utils.json_to_sheet(
                safeResult
            );
    } else {
        resultSheet =
            XLSX.utils.aoa_to_sheet(
                [
                    ["Result"],
                    [
                        "No matching records found",
                    ],
                ]
            );
    }

    XLSX.utils.book_append_sheet(
        workbook,
        resultSheet,
        "Answer"
    );

    const questionSheet =
        XLSX.utils.aoa_to_sheet(
            [
                [
                    "Question",
                    question,
                ],
                [
                    "Generated At",
                    new Date().toISOString(),
                ],
            ]
        );

    XLSX.utils.book_append_sheet(
        workbook,
        questionSheet,
        "Query Info"
    );

    const fileName =
        `answer_${Date.now()}.xlsx`;

    const filePath =
        path.join(
            OUTPUT_DIR,
            fileName
        );

    XLSX.writeFile(
        workbook,
        filePath
    );

    return {
        fileName,

        filePath,

        url:
            `/outputs/${fileName}`,
    };
}


// ============================================================
// HEALTH
// ============================================================

app.get(
    "/api/health",
    (req, res) => {
        res.json({
            success: true,

            message:
                "Data Detective AI backend is running",

            datasetLoaded,

            model: MODEL,

            lastAnalysisAvailable:
                Boolean(
                    lastAnalysis
                ),
        });
    }
);


// ============================================================
// AI TEST
// ============================================================

app.get(
    "/api/ai-test",
    async (req, res) => {
        try {
            const message = await callAI([
                { role: "user", content: "Say hello to Data Detective AI in one short sentence." }
            ], { temperature: 0, max_tokens: 60 });

            res.json({ success: true, model: MODEL, message });
        } catch (error) {
            res.status(500).json({ success: false, error: error.message });
        }
    }
);


// ============================================================
// UPLOAD
// ============================================================

app.post(
    "/api/upload",
    upload.single("file"),
    async (
        req,
        res
    ) => {
        let filePath =
            null;

        try {
            if (!req.file) {
                return res
                    .status(400)
                    .json({
                        success: false,

                        error:
                            "No file uploaded",
                    });
            }

            filePath =
                req.file.path;

            const originalName =
                req.file
                    .originalname;

            const extension =
                path
                    .extname(
                        originalName
                    )
                    .toLowerCase();

            let parsed;

            if (
                extension ===
                ".csv"
            ) {
                parsed =
                    await parseCSVFile(
                        filePath
                    );
            } else if (
                extension ===
                    ".xlsx" ||
                extension ===
                    ".xls"
            ) {
                parsed =
                    parseExcelFile(
                        filePath
                    );
            } else if (
                extension ===
                ".json"
            ) {
                parsed =
                    parseJSONFile(
                        filePath
                    );
            } else {
                return res
                    .status(400)
                    .json({
                        success: false,

                        error:
                            "Supported files: CSV, XLSX, XLS, JSON",
                    });
            }

            const {
                headers,
                rows,
            } = parsed;

            if (
                !headers ||
                !headers.length
            ) {
                throw new Error(
                    "No columns detected."
                );
            }

            if (
                !rows ||
                !rows.length
            ) {
                throw new Error(
                    "No data rows detected."
                );
            }

            const schema =
                createSchemaProfile(
                    headers,
                    rows
                );

            const normalizedPath =
                await loadRowsIntoDuckDB(
                    headers,
                    rows
                );

            datasetLoaded =
                true;

            currentDataset = {
                fileName:
                    originalName,

                extension,

                rows,

                headers,

                schema,

                normalizedPath,
            };

            // New dataset = old analysis invalid
            lastAnalysis =
                null;

            console.log(
                "\n========================================"
            );

            console.log(
                "DATASET LOADED"
            );

            console.log(
                "File:",
                originalName
            );

            console.log(
                "Rows:",
                rows.length
            );

            console.log(
                "Columns:",
                headers
            );

            console.log(
                "========================================\n"
            );

            res.json({
                success: true,

                fileName:
                    originalName,

                rowCount:
                    rows.length,

                columnCount:
                    headers.length,

                columns:
                    headers,

                schema,

                preview:
                    rows.slice(
                        0,
                        10
                    ),

                rows,
            });
        } catch (error) {
            console.error(
                "UPLOAD ERROR:",
                error
            );

            res.status(500).json({
                success: false,

                error:
                    error.message ||
                    "Unable to analyse file",
            });
        } finally {
            if (
                filePath &&
                fs.existsSync(
                    filePath
                )
            ) {
                try {
                    fs.unlinkSync(
                        filePath
                    );
                } catch (
                    cleanupError
                ) {
                    console.warn(
                        "Temporary cleanup failed:",
                        cleanupError.message
                    );
                }
            }
        }
    }
);


// ============================================================
// DATASET INFO
// ============================================================

app.get(
    "/api/dataset",
    (
        req,
        res
    ) => {
        if (
            !datasetLoaded ||
            !currentDataset
        ) {
            return res
                .status(404)
                .json({
                    success: false,

                    error:
                        "No dataset is currently loaded",
                });
        }

        res.json({
            success: true,

            fileName:
                currentDataset.fileName,

            rowCount:
                currentDataset.rows
                    .length,

            columnCount:
                currentDataset
                    .headers.length,

            columns:
                currentDataset
                    .headers,

            schema:
                currentDataset
                    .schema,

            preview:
                currentDataset.rows.slice(
                    0,
                    10
                ),
        });
    }
);


// ============================================================
// ASK DATASET
// ============================================================

app.post(
    "/api/ask",
    async (req, res) => {
        const startedAt = Date.now();

        try {
            const question = String(req.body?.question || "").trim();
            const conversation = Array.isArray(req.body?.conversation)
                ? req.body.conversation
                : [];

            if (!question) {
                return res.status(400).json({
                    success: false,
                    error: "Question is required."
                });
            }

            // One semantic AI router decides whether this is normal chat or dataset work.
            const route = await classifyQuestion(question, conversation);

            if (route === "NORMAL_CHAT" || !currentDataset) {
                const answer = await generateNormalAnswer(question, conversation);
                return res.json({
                    success: true,
                    type: "chat",
                    question,
                    answer,
                    verified: false,
                    source: "AI",
                    suggestions: await generateAISuggestions(
                        question,
                        answer,
                        [],
                        conversation
                    ),
                    responseTime: Date.now() - startedAt
                });
            }

            // Dataset question: AI understands the question -> generates SQL -> DuckDB verifies it.
            const sql = cleanSQL(
                await generateSQL(
                    question,
                    currentDataset.schema,
                    conversation
                )
            );

            validateSQL(sql);

            const result = await runQuery(sql);

            const answer = await explainVerifiedResultWithAI(
                question,
                sql,
                result,
                conversation
            );

            lastAnalysis = {
                question,
                sql,
                result,
                answer,
                createdAt: new Date().toISOString()
            };

            const excel = createResultExcel({
                question,
                result
            });

            return res.json({
                success: true,
                type: "dataset",
                question,
                answer,
                sql,
                result,
                columns: result.length ? Object.keys(result[0]) : [],
                verified: true,
                confidence: result.length ? "verified" : "no-match",
                evidence: result,
                sqlSource: "AI semantic SQL",
                source: "DuckDB verified dataset analysis",
                responseTime: Date.now() - startedAt,
                suggestions: await generateAISuggestions(
                    question,
                    answer,
                    result,
                    conversation
                ),
                excel: {
                    fileName: excel.fileName,
                    url: excel.url
                },
                excelUrl: excel.url
            });
        } catch (error) {
            console.error("CHAT ERROR:", error);

            return res.status(500).json({
                success: false,
                error: error?.message || "AI analysis failed."
            });
        }
    }
);

// AI explanation is always based on the actual DuckDB result.
async function explainVerifiedResultWithAI(question, sql, result, conversation = []) {
    if (!result.length) {
        return "Dataset-la indha question-ku matching data kidaikala.";
    }

    return callAI([
        {
            role: "system",
            content: `You explain verified database results.

Use ONLY the supplied DuckDB result. Never invent facts, causes, values, departments, years, or interpretations unsupported by the result.
Reply naturally in Tamil, Tanglish, English, or mixed language according to the user.
Be concise but clear.
If the result is a count, average, percentage, ranking, comparison, or filtered record, state the actual value clearly.
If evidence is insufficient for a causal "why", say that the dataset does not establish the cause.`
        },
        {
            role: "user",
            content: `QUESTION:\n${question}\n\nGENERATED SQL:\n${sql}\n\nVERIFIED DUCKDB RESULT:\n${JSON.stringify(result)}\n\nRECENT CONVERSATION:\n${historyText(conversation)}`
        }
    ], { temperature: 0.2, max_tokens: 800 });
}

// ============================================================
// DOWNLOAD EXCEL
// ============================================================

app.get(
    "/api/download/:fileName",
    (
        req,
        res
    ) => {
        try {
            const fileName =
                path.basename(
                    req.params
                        .fileName
                );

            const filePath =
                path.join(
                    OUTPUT_DIR,
                    fileName
                );

            if (
                !fs.existsSync(
                    filePath
                )
            ) {
                return res
                    .status(404)
                    .json({
                        success: false,

                        error:
                            "Excel file not found",
                    });
            }

            res.download(
                filePath,
                fileName
            );
        } catch (error) {
            res.status(500).json({
                success: false,

                error:
                    error.message,
            });
        }
    }
);


// ============================================================
// SERVER
// ============================================================

const server =
    app.listen(
        PORT,
        () => {
            console.log(
                "\n========================================"
            );

            console.log(
                "       DATA DETECTIVE AI"
            );

            console.log(
                "========================================"
            );

            console.log(
                `Backend: http://localhost:${PORT}`
            );

            console.log(
                "OpenRouter:",
                MODEL
            );

            console.log(
                "DuckDB: READY"
            );

            console.log(
                "CSV: READY"
            );

            console.log(
                "Excel: READY"
            );

            console.log(
                "JSON: READY"
            );

            console.log(
                "Real SQL Verification: READY"
            );

            console.log(
                "Evidence Answers: READY"
            );

            console.log(
                "Eligible Analysis: READY"
            );

            console.log(
                "Placed Analysis: READY"
            );

            console.log(
                "Placement Percentage: READY"
            );

            console.log(
                "Average Analysis: READY"
            );

            console.log(
                "Highest / Lowest: READY"
            );

            console.log(
                "Year Analysis: READY"
            );

            console.log(
                "Comparison: READY"
            );

            console.log(
                "Follow-up Explanation: READY"
            );

            console.log(
                "AI Rate-Limit Fallback: READY"
            );

            console.log(
                "========================================\n"
            );
        }
    );


server.on(
    "error",
    (error) => {
        console.error(
            "SERVER ERROR:",
            error
        );
    }
);