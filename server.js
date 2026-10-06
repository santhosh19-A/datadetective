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
const MODEL = process.env.OPENROUTER_MODEL;
const os = require("os");

const UPLOAD_DIR = path.join(
    os.tmpdir(),
    "data-detective-uploads"
);

const OUTPUT_DIR = path.join(
    os.tmpdir(),
    "data-detective-outputs"
);

if (!fs.existsSync(UPLOAD_DIR)) {
    fs.mkdirSync(UPLOAD_DIR, { recursive: true });
}

if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

if (!process.env.OPENROUTER_API_KEY) {
    console.warn(
        "WARNING: OPENROUTER_API_KEY is missing."
    );
}

if (!MODEL) {
    console.warn(
        "WARNING: OPENROUTER_MODEL is missing."
    );
}

const ai = new OpenAI({
    apiKey: process.env.OPENROUTER_API_KEY,
    baseURL: "https://openrouter.ai/api/v1",
});

const db = new duckdb.Database(":memory:");

let currentDataset = null;
let conversation = [];
let lastAnalysis = null;


/* ============================================================
   EXPRESS
============================================================ */

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


/* ============================================================
   MULTER
============================================================ */

const upload = multer({
    dest: UPLOAD_DIR,
    limits: {
        fileSize: 100 * 1024 * 1024,
    },
});


/* ============================================================
   BASIC HELPERS
============================================================ */

function safeJSON(value) {
    if (typeof value === "bigint") {
        return Number(value);
    }

    if (value instanceof Date) {
        return value.toISOString();
    }

    if (Array.isArray(value)) {
        return value.map(safeJSON);
    }

    if (value && typeof value === "object") {
        const result = {};

        for (const [key, val] of Object.entries(value)) {
            result[key] = safeJSON(val);
        }

        return result;
    }

    return value;
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


function cleanHeader(value, index) {
    const text = String(
        value ?? ""
    ).trim();

    return (
        text ||
        `Column_${index + 1}`
    );
}


function uniqueHeaders(headers) {
    const used = new Map();

    return headers.map(
        (header, index) => {
            const base =
                cleanHeader(
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


function quoteIdentifier(name) {
    return `"${String(name).replace(/"/g, '""')}"`;
}


function sqlString(value) {
    return `'${String(value).replace(/'/g, "''")}'`;
}


function isEmpty(value) {
    return (
        value === null ||
        value === undefined ||
        String(value).trim() === ""
    );
}


/* ============================================================
   DUCKDB
============================================================ */

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
                        safeJSON(
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


/* ============================================================
   CSV
============================================================ */

function parseCSVFile(filePath) {
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
                                cleanHeader(
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

                        const normalized =
                            rows.map(
                                (row) => {
                                    const result =
                                        {};

                                    for (
                                        const header of headers
                                    ) {
                                        result[
                                            header
                                        ] =
                                            normalizeCell(
                                                row[
                                                    header
                                                ]
                                            );
                                    }

                                    return result;
                                }
                            );

                        resolve({
                            headers,
                            rows:
                                normalized,
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


/* ============================================================
   EXCEL
============================================================ */

function detectHeaderRow(rows) {
    let bestIndex = -1;
    let bestScore = -1;

    const keywords = [
        "customer",
        "customer id",
        "name",
        "city",
        "location",
        "date",
        "year",
        "sales",
        "revenue",
        "amount",
        "quantity",
        "product",
        "department",
        "category",
        "price",
        "order",
    ];

    for (
        let i = 0;
        i < Math.min(
            rows.length,
            30
        );
        i++
    ) {
        const values =
            rows[i]
                .map((v) =>
                    String(
                        v ?? ""
                    )
                        .trim()
                        .toLowerCase()
                )
                .filter(Boolean);

        if (values.length < 2) {
            continue;
        }

        let score =
            values.length;

        for (
            const value of values
        ) {
            for (
                const keyword of keywords
            ) {
                if (
                    value ===
                    keyword
                ) {
                    score += 10;
                } else if (
                    value.includes(
                        keyword
                    )
                ) {
                    score += 3;
                }
            }
        }

        if (
            score >
            bestScore
        ) {
            bestScore =
                score;

            bestIndex =
                i;
        }
    }

    return bestIndex === -1
        ? 0
        : bestIndex;
}


function parseExcelFile(filePath) {
    const workbook =
        XLSX.readFile(
            filePath,
            {
                cellDates: true,
            }
        );

    const sheetName =
        workbook
            .SheetNames[0];

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
        detectHeaderRow(
            rawRows
        );

    const headers =
        uniqueHeaders(
            rawRows[
                headerIndex
            ].map(
                cleanHeader
            )
        );

    const rows = [];

    for (
        let i =
            headerIndex + 1;
        i < rawRows.length;
        i++
    ) {
        const raw =
            rawRows[i];

        const hasData =
            raw.some(
                (value) =>
                    !isEmpty(value)
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
                row[
                    header
                ] =
                    normalizeCell(
                        raw[index]
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


/* ============================================================
   JSON
============================================================ */

function parseJSONFile(filePath) {
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

    for (
        const item of data
    ) {
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

    const originalHeaders =
        Array.from(
            headerSet
        );

    const headers =
        uniqueHeaders(
            originalHeaders
        );

    const rows =
        data.map(
            (item) => {
                const row = {};

                headers.forEach(
                    (
                        header,
                        index
                    ) => {
                        const original =
                            originalHeaders[
                                index
                            ];

                        row[
                            header
                        ] =
                            normalizeCell(
                                item?.[
                                    original
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


/* ============================================================
   CSV CONVERSION
============================================================ */

function csvEscape(value) {
    const text =
        normalizeCell(
            value
        );

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


function rowsToCSV(
    headers,
    rows
) {
    const output = [];

    output.push(
        headers
            .map(csvEscape)
            .join(",")
    );

    for (
        const row of rows
    ) {
        output.push(
            headers
                .map(
                    (header) =>
                        csvEscape(
                            row[
                                header
                            ]
                        )
                )
                .join(",")
        );
    }

    return output.join(
        "\n"
    );
}
/* ============================================================
   CREATE DUCKDB DATASET TABLE
============================================================ */

async function createDatasetTable(headers, rows) {
    if (!headers || !headers.length) {
        throw new Error(
            "Dataset does not contain any columns."
        );
    }

    await runStatement(
        `DROP TABLE IF EXISTS dataset`
    );

    const columnDefinitions =
        headers
            .map(
                (header) =>
                    `${quoteIdentifier(
                        header
                    )} VARCHAR`
            )
            .join(", ");

    await runStatement(
        `CREATE TABLE dataset (${columnDefinitions})`
    );

    if (
        !rows ||
        !rows.length
    ) {
        return;
    }

    const columnList =
        headers
            .map(quoteIdentifier)
            .join(", ");

    /*
     * Insert in batches.
     * This avoids DuckDB prepared-statement parameter
     * errors and avoids creating one enormous INSERT.
     */

    const BATCH_SIZE = 500;

    for (
        let start = 0;
        start < rows.length;
        start += BATCH_SIZE
    ) {
        const batch =
            rows.slice(
                start,
                start + BATCH_SIZE
            );

        const valueRows =
            batch.map(
                (row) => {
                    const values =
                        headers.map(
                            (header) =>
                                sqlString(
                                    normalizeCell(
                                        row[
                                            header
                                        ]
                                    )
                                )
                        );

                    return `(${values.join(
                        ", "
                    )})`;
                }
            );

        const insertSQL = `
            INSERT INTO dataset (${columnList})
            VALUES ${valueRows.join(
                ",\n"
            )}
        `;

        await runStatement(
            insertSQL
        );
    }
}


/* ============================================================
   DATASET SCHEMA
============================================================ */

async function getDatasetSchema() {
    return runQuery(`
        SELECT
            column_name,
            data_type
        FROM information_schema.columns
        WHERE table_name = 'dataset'
        ORDER BY ordinal_position
    `);
}


async function getDatasetPreview(
    limit = 10
) {
    const safeLimit =
        Math.max(
            1,
            Math.min(
                Number(limit) || 10,
                100
            )
        );

    return runQuery(
        `SELECT * FROM dataset LIMIT ${safeLimit}`
    );
}


/* ============================================================
   AI HELPERS
============================================================ */

function requireAIKey() {
    if (
        !process.env
            .OPENROUTER_API_KEY
    ) {
        throw new Error(
            "OPENROUTER_API_KEY is not configured."
        );
    }

    if (!MODEL) {
        throw new Error(
            "OPENROUTER_MODEL is not configured."
        );
    }
}


async function callAI(
    messages,
    options = {}
) {
    requireAIKey();

    const response =
        await ai.chat.completions.create(
            {
                model: MODEL,
                messages,
                temperature:
                    options.temperature ??
                    0.2,
                max_tokens:
                    options.max_tokens ??
                    1200,
            }
        );

    console.log(
        "OPENROUTER RESPONSE:",
        JSON.stringify(
            response,
            null,
            2
        )
    );

    const choice =
        response?.choices?.[0];

    if (!choice) {
        throw new Error(
            "AI returned no choices."
        );
    }

    let content =
        choice?.message?.content;

    if (
        Array.isArray(
            content
        )
    ) {
        content =
            content
                .map(
                    (part) => {
                        if (
                            typeof part ===
                            "string"
                        ) {
                            return part;
                        }

                        return (
                            part?.text ||
                            part?.content ||
                            ""
                        );
                    }
                )
                .join("");
    }

    if (
        !content &&
        typeof choice?.text ===
            "string"
    ) {
        content =
            choice.text;
    }

    if (
        !content &&
        typeof choice
            ?.message
            ?.reasoning ===
            "string"
    ) {
        content =
            choice.message
                .reasoning;
    }

    if (
        !content ||
        !String(
            content
        ).trim()
    ) {
        throw new Error(
            "AI returned an empty response."
        );
    }

    return String(
        content
    ).trim();
}


/* ============================================================
   CONVERSATION CONTEXT
============================================================ */

function addConversation(
    role,
    content
) {
    conversation.push({
        role,
        content,
    });

    if (
        conversation.length >
        20
    ) {
        conversation =
            conversation.slice(
                -20
            );
    }
}


function getConversationContext() {
    if (
        !conversation.length
    ) {
        return "";
    }

    return conversation
        .map(
            (message) =>
                `${message.role}: ${message.content}`
        )
        .join("\n");
}


/* ============================================================
   SEMANTIC QUESTION CLASSIFIER
============================================================ */

async function classifyQuestion(question) {
    const context = getConversationContext();

    let schemaText = "";
    let previewText = "";

    try {
        const schema = await getDatasetSchema();

        schemaText = schema
            .map(
                (column) =>
                    `${column.column_name} (${column.data_type})`
            )
            .join("\n");

        const preview = await getDatasetPreview(5);

        previewText = JSON.stringify(
            preview,
            null,
            2
        );
    } catch (error) {
        console.warn(
            "Classifier dataset context error:",
            error.message
        );
    }

    const answer = await callAI(
        [
            {
                role: "system",
                content: `
You are the semantic routing brain of a data-analysis chatbot.

A real dataset is already uploaded.

Your ONLY job is to decide whether the user's message
needs the uploaded dataset.

Return exactly one word:

DATASET

or

CHAT

DATASET means the user wants information from,
about, or related to the uploaded dataset.

This includes:
- asking what the dataset is about
- asking what the dataset is for
- asking what fields or columns exist
- asking what data is stored
- asking what the records represent
- asking about customers
- asking about sales
- asking about locations
- asking about years
- asking about departments
- asking about products
- asking for counts
- asking for totals
- asking for averages
- asking for comparisons
- asking for rankings
- asking for filtering
- asking for trends
- asking follow-up questions about previous dataset results

Understand:
- English
- Tamil
- Tanglish
- mixed language

Examples:

"what is this dataset about?" = DATASET
"what is this dataset for?" = DATASET
"what fields are there?" = DATASET
"what data is stored here?" = DATASET
"இந்த dataset எதுக்காக?" = DATASET
"இந்த data எதுக்காக?" = DATASET
"intha dataset ethukaga iruku" = DATASET
"intha dataset pathi sollu" = DATASET
"ithula entha field oda data iruku" = DATASET
"ithula enna data store pannirukku" = DATASET
"Chennai la evlo customers?" = DATASET
"what about IT?" = DATASET
"2018 la enna nadanthuchu?" = DATASET

CHAT means the message does NOT require the uploaded dataset.

Examples:

"hi" = CHAT
"hello" = CHAT
"nanba" = CHAT
"enna pandra?" = CHAT
"nee yaru?" = CHAT
"tell me a joke" = CHAT
"what is photosynthesis?" = CHAT
"what is 2+2?" = CHAT

The dataset is already uploaded.
Never ask the user to upload it again.

Dataset filename:
${currentDataset?.filename || "uploaded dataset"}

Dataset schema:
${schemaText || "(schema unavailable)"}

Sample rows:
${previewText || "(sample unavailable)"}

Previous conversation:
${context || "(none)"}

Return ONLY DATASET or CHAT.
`,
            },
            {
                role: "user",
                content: question,
            },
        ],
        {
            temperature: 0,
            max_tokens: 30,
        }
    );

    const raw = String(answer || "")
        .trim()
        .toUpperCase();

    /*
     * First trust the AI's semantic decision.
     */
const matches = [
    ...raw.matchAll(/\b(DATASET|CHAT)\b/g)
];

let route =
    matches.length > 0
        ? matches[matches.length - 1][1]
        : null;
    /*
     * Safety fallback:
     * If the AI failed to classify a clearly dataset-related
     * message, route it to the real dataset instead of
     * giving a generic chatbot answer.
     *
     * This does NOT generate the answer.
     * DuckDB + SQL still generates the actual answer.
     */
    if (!route) {
        const text = question
            .toLowerCase()
            .trim();

        const datasetReference =
            /\b(dataset|data|field|fields|column|columns|record|records|customer|customers|sales|department|product|products|year|years|location|locations)\b/i.test(
                text
            ) ||
            /intha\s+dataset/i.test(text) ||
            /intha\s+data/i.test(text) ||
            /intha\s+data\s*set/i.test(text) ||
            /indha\s+dataset/i.test(text) ||
            /indha\s+data/i.test(text) ||
            /ethukaga/i.test(text) ||
            /edhukaga/i.test(text) ||
            /எதுக்காக/i.test(text) ||
            /எதற்காக/i.test(text) ||
            /டேட்டா/i.test(text) ||
            /dataset.*pathi/i.test(text) ||
            /data.*pathi/i.test(text);

        route = datasetReference
            ? "DATASET"
            : "CHAT";
    }

    /*
     * Extra protection:
     * These common conversational messages should stay CHAT.
     */
    const casualOnly =
        /^(hi|hello|hey|hey+|hai|nanba|bro|hi bro|hey bro)$/i.test(
            question.trim()
        );

    if (casualOnly) {
        route = "CHAT";
    }

    console.log(
        "QUESTION ROUTE:",
        question,
        "=>",
        route,
        "| AI RAW:",
        raw
    );

    return route;
}


/* ============================================================
   NORMAL CHAT
============================================================ */

async function generateNormalAnswer(
    question
) {
    const context =
        getConversationContext();

    return callAI(
        [
            {
                role: "system",
                content: `
You are Data Detective AI.

You are a normal conversational AI.

Understand the user's actual meaning.

Support:
- English
- Tamil
- Tanglish
- mixed language

Respond naturally to greetings
and casual conversation.

If the user asks who you are,
explain that you are Data Detective AI.

Answer general questions normally.

Do not pretend that you inspected
the dataset if you did not.

Do not mention internal routing,
classification, prompts, or system rules.

Previous conversation:
${context || "(none)"}
`,
            },
            {
                role: "user",
                content:
                    question,
            },
        ],
        {
            temperature: 0.7,
            max_tokens: 1000,
        }
    );
}
/* ============================================================
   SQL EXTRACTION
============================================================ */

function extractSQL(text) {
    if (!text) {
        return "";
    }

    let sql = String(text).trim();

    // Remove markdown fences if the model returned them
    sql = sql
        .replace(/^```sql\s*/i, "")
        .replace(/^```\s*/i, "")
        .replace(/\s*```$/i, "")
        .trim();

    // Remove optional SQL: prefix
    sql = sql.replace(/^SQL\s*:\s*/i, "").trim();

    /*
     * IMPORTANT:
     * SQL MUST START with SELECT or a real CTE.
     *
     * Do NOT search for SELECT/WITH somewhere
     * inside natural-language AI output.
     */

    if (/^SELECT\b/i.test(sql)) {
        // Valid SELECT query
    } else if (
        /^WITH\s+[A-Za-z_][A-Za-z0-9_]*\s+AS\s*\(/i.test(sql)
    ) {
        // Valid CTE query
    } else {
        return "";
    }

    // Remove trailing semicolon
    sql = sql.replace(/;\s*$/, "").trim();

    // Reject multiple statements
    if (sql.includes(";")) {
        return "";
    }

    return sql;
}

/* ============================================================
   SQL VALIDATION
============================================================ */

function validateSQL(sql) {
    if (
        !sql ||
        !sql.trim()
    ) {
        throw new Error(
            "SQL query is empty."
        );
    }

    const normalized =
        sql.trim().toUpperCase();

    if (
        !normalized.startsWith(
            "SELECT"
        ) &&
        !normalized.startsWith(
            "WITH"
        )
    ) {
        throw new Error(
            "Only SELECT/WITH queries are allowed."
        );
    }

    if (
        normalized.includes(";")
    ) {
        throw new Error(
            "Multiple SQL statements are not allowed."
        );
    }
const dangerousPatterns = [
    /\bINSERT\b/i,
    /\bUPDATE\b/i,
    /\bDELETE\b/i,
    /\bDROP\b/i,
    /\bALTER\b/i,
    /\bTRUNCATE\b/i,
    /\bATTACH\b/i,
    /\bDETACH\b/i,
    /\bCOPY\b/i,
    /\bEXPORT\b/i,
    /\bIMPORT\b/i,
    /\bPRAGMA\b/i,
    /\bCREATE\b/i,
    /\bVACUUM\b/i,
    /--/,
    /\/\*/,
    /\*\//
];

    for (
        const pattern of
            dangerousPatterns
    ) {
        if (
            pattern.test(sql)
        ) {
            throw new Error(
                "Unsafe SQL detected."
            );
        }
    }

    return true;
}


/* ============================================================
   SQL GENERATION
============================================================ */

async function generateSQL(
    question,
    schema
) {
    const schemaText =
        schema
            .map(
                (column) =>
                    `${column.column_name} (${column.data_type})`
            )
            .join("\n");

    const context =
        getConversationContext();

    const answer =
        await callAI(
            [
                {
                    role: "system",
                    content: `
You are the SQL generation engine
for Data Detective AI.

The uploaded dataset is available as:

dataset

Available columns:
${schemaText}

Generate ONE DuckDB SQL query that
answers the user's question.

Rules:
IMPORTANT SQL RULES:
- Return ONLY executable DuckDB SQL.
- Never return explanations, comments, reasoning, or natural language.
- The available table is exactly: dataset
- If using alias d, always write: FROM dataset AS d
- Never use an undefined table alias.
- Never invent column names.
- Do not use markdown code fences.
- The output must start with SELECT or WITH.
1. Return ONLY SQL.
2. Query must start with SELECT or WITH.
3. Only read data.
4. Never use INSERT, UPDATE, DELETE,
   DROP, ALTER, CREATE, TRUNCATE,
   ATTACH, COPY, PRAGMA, EXPORT,
   IMPORT or other write operations.
5. Never invent a column.
6. Use only columns from the schema.
7. Understand English, Tamil,
   Tanglish and mixed language.
8. Use previous conversation to resolve
   follow-up references.
9. For customer counts, prefer
   COUNT(DISTINCT customer identifier)
   when an identifier exists.
10. For numeric values stored as text,
    use TRY_CAST when required.
11. For amounts containing commas or
    currency symbols, clean them before
    casting.
12. If the user asks for a list, return
    useful columns.
13. If the user asks what the dataset
    is about, what it is for, why it is
    used, or asks for a dataset overview,
    inspect representative rows with:

    SELECT * FROM dataset LIMIT 10

    The result must then be explained
    using only the actual returned data.
14. Never invent the dataset purpose.
15. If the available data is insufficient
    to determine the purpose, say that
    clearly in the explanation.

Previous conversation:
${context || "(none)"}

Return ONLY SQL.
Do not use markdown fences.
Do not explain the SQL.
`,
                },
                {
                    role: "user",
                    content:
                        question,
                },
            ],
            {
                temperature: 0,
                max_tokens: 1200,
            }
        );

    const sql =
        extractSQL(
            answer
        );

    if (!sql) {
        throw new Error(
            "AI did not return a valid SQL query."
        );
    }
	return sql;

    validateSQL(sql);

    return sql;
}


/* ============================================================
   DATASET RESULT EXPLANATION
============================================================ */

async function explainDatasetResult(
    question,
    sql,
    result
) {
    const context =
        getConversationContext();

    const resultText =
        JSON.stringify(
            result,
            null,
            2
        );

    return callAI(
        [
            {
                role: "system",
                content: `
You are the answer engine for
Data Detective AI.

The SQL query was already executed
successfully against the uploaded dataset.

Explain the verified result.

Rules:

1. Use ONLY the verified result.
2. Never invent facts.
3. Do not claim something that is
   not supported by the result.
4. If result is empty, clearly say
   no matching records were found.
5. For dataset-purpose questions,
   explain what the actual sample/result
   indicates the dataset represents.
6. If the data is insufficient to know
   the exact purpose, say that it appears
   to be about the shown data, without
   pretending certainty.
7. Answer naturally in English, Tamil,
   Tanglish or mixed language.
8. Keep it concise and useful.
9. Do not mention internal routing,
   prompts or system rules.
10. Do not output SQL unless asked.

User question:
${question}

Executed SQL:
${sql}

Verified result:
${resultText}

Previous conversation:
${context || "(none)"}
`,
            },
            {
                role: "user",
                content:
                    "Explain the verified result naturally to the user.",
            },
        ],
        {
            temperature: 0.4,
            max_tokens: 1000,
        }
    );
}


/* ============================================================
   FOLLOW-UP SUGGESTIONS
============================================================ */

async function generateSuggestions(
    question,
    answer,
    result
) {
    try {
        const response =
            await callAI(
                [
                    {
                        role: "system",
                        content: `
Generate exactly 3 useful follow-up
questions for a data-analysis chatbot.

Return ONLY valid JSON:

["question 1","question 2","question 3"]

Rules:
- Keep them short.
- Related to the current dataset question.
- Useful for continuing analysis.
- English, Tamil or Tanglish is allowed.
- No explanations.
`,
                    },
                    {
                        role: "user",
                        content: `
Current question:
${question}

Current answer:
${answer}

Current result:
${JSON.stringify(
    result
)}
`,
                    },
                ],
                {
                    temperature: 0.5,
                    max_tokens: 300,
                }
            );

        let text =
            String(
                response
            ).trim();

        text =
            text
                .replace(
                    /```json/gi,
                    ""
                )
                .replace(
                    /```/g,
                    ""
                )
                .trim();

        const parsed =
            JSON.parse(text);

        if (
            !Array.isArray(
                parsed
            )
        ) {
            return [];
        }

        return parsed
            .filter(
                (item) =>
                    typeof item ===
                        "string" &&
                    item.trim()
            )
            .slice(
                0,
                3
            );
    } catch (error) {
        console.warn(
            "Suggestion generation failed:",
            error.message
        );

        return [];
    }
}


/* ============================================================
   UPLOAD ROUTE
============================================================ */

app.post(
    "/api/upload",
    upload.single("file"),
    async (
        req,
        res
    ) => {
        try {
            if (!req.file) {
                return res
                    .status(400)
                    .json({
                        error:
                            "No file was uploaded.",
                    });
            }

            const originalName =
                req.file
                    .originalname;

            const extension =
                path.extname(
                    originalName
                ).toLowerCase();

            let parsed;

            if (
                extension ===
                ".csv"
            ) {
                parsed =
                    await parseCSVFile(
                        req.file.path
                    );
            } else if (
                extension ===
                    ".xlsx" ||
                extension ===
                    ".xls"
            ) {
                parsed =
                    parseExcelFile(
                        req.file.path
                    );
            } else if (
                extension ===
                ".json"
            ) {
                parsed =
                    parseJSONFile(
                        req.file.path
                    );
            } else {
                return res
                    .status(400)
                    .json({
                        error:
                            "Unsupported file type. Use CSV, XLSX, XLS, or JSON.",
                    });
            }

            if (
                !parsed.headers ||
                !parsed.headers.length
            ) {
                return res
                    .status(400)
                    .json({
                        error:
                            "Could not detect columns in the uploaded file.",
                    });
            }

            await createDatasetTable(
                parsed.headers,
                parsed.rows
            );

            const schema =
                await getDatasetSchema();

            const preview =
                await getDatasetPreview(
                    10
                );

            currentDataset = {
                filename:
                    originalName,

                fileName:
                    originalName,

                extension,

                rows:
                    parsed.rows.length,

                rowCount:
                    parsed.rows.length,

                columns:
                    parsed.headers.length,

                columnCount:
                    parsed.headers.length,

                headers:
                    parsed.headers,

                schema,

                preview,

                uploadedAt:
                    new Date().toISOString(),
            };

            conversation = [];

            lastAnalysis = null;

            try {
                fs.unlinkSync(
                    req.file.path
                );
            } catch {
                // Ignore cleanup errors.
            }

            return res.json({
                success: true,

                type: "upload",

                message:
                    `I've loaded ${originalName} successfully.`,

                filename:
                    originalName,

                fileName:
                    originalName,

                rows:
                    parsed.rows.length,

                rowCount:
                    parsed.rows.length,

                columns:
                    parsed.headers.length,

                columnCount:
                    parsed.headers.length,

                headers:
                    parsed.headers,

                schema,

                preview,

                dataset:
                    currentDataset,
            });
        } catch (error) {
            console.error(
                "UPLOAD ERROR:",
                error
            );

            try {
                if (
                    req.file?.path &&
                    fs.existsSync(
                        req.file.path
                    )
                ) {
                    fs.unlinkSync(
                        req.file.path
                    );
                }
            } catch {
                // Ignore cleanup errors.
            }

            return res
                .status(500)
                .json({
                    error:
                        error.message ||
                        "Failed to process uploaded file.",
                });
        }
    }
);


/* ============================================================
   ASK ROUTE
============================================================ */

app.post(
    "/api/ask",
    async (
        req,
        res
    ) => {
        const startedAt =
            Date.now();

        try {
            const question =
                String(
                    req.body?.question ||
                        ""
                ).trim();

            if (!question) {
                return res
                    .status(400)
                    .json({
                        error:
                            "Please enter a question.",
                    });
            }

            const hasDataset =
                Boolean(
                    currentDataset
                );

            let route =
                "CHAT";

            if (
                hasDataset
            ) {
                route =
                    await classifyQuestion(
                        question
                    );
            }

            /*
             * Only DATASET enters SQL workflow.
             * CHAT always uses normal conversation.
             */

            if (
                route !==
                    "DATASET" ||
                !hasDataset
            ) {
                const answer =
                    await generateNormalAnswer(
                        question
                    );

                addConversation(
                    "user",
                    question
                );

                addConversation(
                    "assistant",
                    answer
                );

                return res.json({
                    type: "chat",

                    answer,

                    executionTime:
                        Date.now() -
                        startedAt,
                });
            }

            /* ================================================
               DATASET WORKFLOW
            ================================================= */

            const schema =
                await getDatasetSchema();

            const sql =
                await generateSQL(
                    question,
                    schema
                );

            validateSQL(
                sql
            );

            const result =
                await runQuery(
                    sql
                );

            const answer =
                await explainDatasetResult(
                    question,
                    sql,
                    result
                );

            const suggestions =
                await generateSuggestions(
                    question,
                    answer,
                    result
                );

            addConversation(
                "user",
                question
            );

            addConversation(
                "assistant",
                answer
            );

            lastAnalysis = {
                question,

                sql,

                result,

                answer,

                createdAt:
                    new Date().toISOString(),
            };

            const columns =
                result.length
                    ? Object.keys(
                          result[0]
                      )
                    : [];

            return res.json({
                type: "dataset",

                answer,

                sql,

                result,

                columns,

                rowCount:
                    result.length,

                suggestions,

                executionTime:
                    Date.now() -
                    startedAt,
            });
        } catch (error) {
            console.error(
                "ASK ERROR:",
                error
            );

            return res
                .status(500)
                .json({
                    error:
                        error.message ||
                        "Failed to process your question.",

                    executionTime:
                        Date.now() -
                        startedAt,
                });
        }
    }
);
/* ============================================================
   DATASET INFO
============================================================ */

app.get(
    "/api/dataset",
    async (
        req,
        res
    ) => {
        try {
            if (!currentDataset) {
                return res.json({
                    success: true,
                    dataset: null,
                });
            }

            return res.json({
                success: true,
                dataset:
                    currentDataset,
            });
        } catch (error) {
            return res
                .status(500)
                .json({
                    error:
                        error.message,
                });
        }
    }
);


/* ============================================================
   LAST ANALYSIS
============================================================ */

app.get(
    "/api/last-analysis",
    (
        req,
        res
    ) => {
        return res.json({
            success: true,
            analysis:
                lastAnalysis,
        });
    }
);


/* ============================================================
   CLEAR DATASET
============================================================ */

app.post(
    "/api/clear",
    async (
        req,
        res
    ) => {
        try {
            await runStatement(
                "DROP TABLE IF EXISTS dataset"
            );

            currentDataset = null;
            conversation = [];
            lastAnalysis = null;

            return res.json({
                success: true,
                message:
                    "Dataset cleared successfully.",
            });
        } catch (error) {
            console.error(
                "CLEAR ERROR:",
                error
            );

            return res
                .status(500)
                .json({
                    error:
                        error.message ||
                        "Failed to clear dataset.",
                });
        }
    }
);


/* ============================================================
   HEALTH
============================================================ */

app.get(
    "/api/health",
    (
        req,
        res
    ) => {
        return res.json({
            success: true,

            status: "ok",

            server:
                "Data Detective AI",

            model:
                MODEL || null,

            datasetLoaded:
                Boolean(
                    currentDataset
                ),

            dataset:
                currentDataset
                    ? {
                          filename:
                              currentDataset.filename,

                          rows:
                              currentDataset.rows,

                          columns:
                              currentDataset.columns,
                      }
                    : null,

            uptime:
                process.uptime(),
        });
    }
);


/* ============================================================
   AI TEST
============================================================ */

app.get(
    "/api/ai-test",
    async (
        req,
        res
    ) => {
        try {
            const answer =
                await callAI(
                    [
                        {
                            role: "system",
                            content:
                                "Reply with exactly: AI OK",
                        },
                        {
                            role: "user",
                            content:
                                "Test connection.",
                        },
                    ],
                    {
                        temperature: 0,
                        max_tokens: 20,
                    }
                );

            return res.json({
                success: true,
                answer,
                model:
                    MODEL,
            });
        } catch (error) {
            console.error(
                "AI TEST ERROR:",
                error
            );

            return res
                .status(500)
                .json({
                    success: false,

                    error:
                        error.message,

                    model:
                        MODEL,
                });
        }
    }
);


/* ============================================================
   FRONTEND FALLBACK
============================================================ */

app.use(
    (
        req,
        res
    ) => {
        const indexPath =
            path.join(
                __dirname,
                "public",
                "index.html"
            );

        if (
            fs.existsSync(
                indexPath
            )
        ) {
            return res.sendFile(
                indexPath
            );
        }

        return res
            .status(404)
            .send(
                "Data Detective AI server is running."
            );
    }
);


/* ============================================================
   GLOBAL ERROR HANDLER
============================================================ */

app.use(
    (
        error,
        req,
        res,
        next
    ) => {
        console.error(
            "GLOBAL ERROR:",
            error
        );

        if (
            res.headersSent
        ) {
            return next(
                error
            );
        }

        return res
            .status(
                error.status || 500
            )
            .json({
                error:
                    error.message ||
                    "Internal server error.",
            });
    }
);


/* ============================================================
   SERVER EXPORT
============================================================ */

module.exports = app;


/* ============================================================
   LOCAL DEVELOPMENT
============================================================ */

if (require.main === module) {
    app.listen(PORT, () => {
        console.log("");
        console.log("==========================================");
        console.log("       DATA DETECTIVE AI SERVER");
        console.log("==========================================");
        console.log(`Server running on http://localhost:${PORT}`);
        console.log(`Model: ${MODEL || "NOT CONFIGURED"}`);
        console.log("Dataset: ready for upload");
        console.log("==========================================");
        console.log("");
    });
}