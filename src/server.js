// ============================================================
// NATIONAL FOOD COMPANY - AMERICANA CAKE
// PREVENTIVE MAINTENANCE SYSTEM
//
// SERVER
// ============================================================

require("dotenv").config();


const express =
    require("express");


const cors =
    require("cors");


const fs =
    require("fs");


const path =
    require("path");


const crypto =
    require("crypto");


const multer =
    require("multer");


const XLSX =
    require("xlsx");


const pool =
    require("./database");


// ============================================================
// AUTHENTICATION
// ============================================================

const {

    authRouter,

    sessionMiddleware,

    requireAuth,

    requirePageAuth,

    requireRole,

    ensureDefaultAdmin

} = require("./auth");


// ============================================================
// CORRECTED PM SCHEDULING ENGINE
// ============================================================

const {

    getISOWeekNumber,

    getISOWeeksInYear,

    getISOWeekMonday,

    getMonthForWeek,

    getWeeksForMonth,

    shouldScheduleTask,

    generateWeekSchedule,

    generateFullYear,

    rebuildYear,

    checkDistribution,

    validateDistribution

} = require("./generateSchedule");


// ============================================================
// EXPRESS
// ============================================================

const app =
    express();


// ============================================================
// GENERAL MIDDLEWARE
// ============================================================

app.use(
    cors({
        origin:
            true,

        credentials:
            true
    })
);


app.use(
    express.json({
        limit:
            "10mb"
    })
);


// ============================================================
// SESSION
// ============================================================

app.use(
    sessionMiddleware
);


// ============================================================
// AUTHENTICATION API
//
// /api/auth/login
// /api/auth/logout
// /api/auth/me
// /api/auth/change-password
// /api/auth/users
// ============================================================

app.use(
    "/api/auth",
    authRouter
);


// ============================================================
// LOGIN PAGE
// ============================================================

app.get(
    "/login",
    (
        req,
        res
    ) => {

        if (
            req.session &&
            req.session.user
        ) {

            return res.redirect(
                "/"
            );
        }


        res.sendFile(
            path.join(
                __dirname,
                "../public/login.html"
            )
        );
    }
);


// ============================================================
// PROTECT DIRECT INDEX.HTML
// ============================================================

app.get(
    "/index.html",

    requirePageAuth,

    (
        req,
        res
    ) => {

        res.sendFile(
            path.join(
                __dirname,
                "../public/index.html"
            )
        );
    }
);


// ============================================================
// STATIC FILES
//
// index:false stops Express opening index.html before login.
// ============================================================

app.use(
    express.static(
        path.join(
            __dirname,
            "../public"
        ),
        {
            index:
                false
        }
    )
);


// ============================================================
// MAIN PM APPLICATION
// ============================================================

app.get(
    "/",

    requirePageAuth,

    (
        req,
        res
    ) => {

        res.sendFile(
            path.join(
                __dirname,
                "../public/index.html"
            )
        );
    }
);


// ============================================================
// PROTECT ALL PM API ROUTES BELOW THIS POINT
// ============================================================

app.use(
    "/api",
    requireAuth
);


// ============================================================
// EXCEL UPLOAD CONFIGURATION
// ============================================================

const upload =
    multer({

        storage:
            multer.memoryStorage(),

        limits: {

            fileSize:
                20 *
                1024 *
                1024
        },

        fileFilter:
            (
                req,
                file,
                callback
            ) => {

                const extension =
                    path
                        .extname(
                            file.originalname
                        )
                        .toLowerCase();


                if (
                    extension !==
                    ".xlsx" &&
                    extension !==
                    ".xls"
                ) {

                    return callback(
                        new Error(
                            "Only Excel .xlsx and .xls files are allowed."
                        )
                    );
                }


                callback(
                    null,
                    true
                );
            }
    });


// ============================================================
// TEMPORARY EXCEL PREVIEW STORAGE
// ============================================================

const importPreviews =
    new Map();


const IMPORT_PREVIEW_LIFETIME =
    30 *
    60 *
    1000;


// ============================================================
// NORMALIZE TEXT
// ============================================================

function normalizeText(
    value
) {

    return String(
        value ??
        ""
    )
        .trim()
        .replace(
            /\s+/g,
            " "
        )
        .toLowerCase();
}


// ============================================================
// NORMALIZE EXCEL HEADER
// ============================================================

function normalizeHeader(
    value
) {

    return normalizeText(
        value
    )
        .replace(
            /[^a-z0-9]/g,
            ""
        );
}


// ============================================================
// EXCEL COLUMN NAMES
// ============================================================

const COLUMN_ALIASES =
{

    line_name:
    [
        "line",
        "line name",
        "linename",
        "production line",
        "productionline"
    ],

    section:
    [
        "section",
        "area",
        "department"
    ],

    machine_code:
    [
        "machine code",
        "machinecode",
        "machine no",
        "machineno",
        "machine number",
        "machinenumber",
        "asset code",
        "assetcode"
    ],

    machine_name:
    [
        "machine",
        "machine name",
        "machinename",
        "equipment",
        "equipment name",
        "equipmentname"
    ],

    part_name:
    [
        "part",
        "part name",
        "partname",
        "component",
        "component name",
        "componentname"
    ],

    maintenance_task:
    [
        "pm",
        "pm task",
        "pmtask",
        "task",
        "maintenance",
        "maintenance task",
        "maintenancetask",
        "maintenance action",
        "action"
    ],

    frequency_text:
    [
        "freq",
        "frequency",
        "frequency text",
        "frequencytext",
        "pm frequency",
        "pmfrequency",
        "interval"
    ],

    active:
    [
        "active",
        "enabled",
        "status"
    ]
};


// ============================================================
// IDENTIFY EXCEL HEADER
// ============================================================

function identifyHeaderField(
    header
) {

    const normalized =
        normalizeHeader(
            header
        );


    for (
        const [
            field,
            aliases
        ]
        of Object.entries(
            COLUMN_ALIASES
        )
    ) {

        for (
            const alias
            of aliases
        ) {

            if (
                normalized ===
                normalizeHeader(
                    alias
                )
            ) {

                return field;
            }
        }
    }


    return null;
}


// ============================================================
// FIND HEADER ROW IN EXCEL
// ============================================================

function findExcelHeaderRow(
    rows
) {

    const maxRows =
        Math.min(
            rows.length,
            30
        );


    for (
        let rowIndex = 0;
        rowIndex < maxRows;
        rowIndex++
    ) {

        const identified =
            rows[
                rowIndex
            ]
                .map(
                    identifyHeaderField
                )
                .filter(
                    Boolean
                );


        const hasMachine =
            identified.includes(
                "machine_code"
            ) ||
            identified.includes(
                "machine_name"
            );


        const hasTask =
            identified.includes(
                "maintenance_task"
            );


        const hasFrequency =
            identified.includes(
                "frequency_text"
            );


        if (
            hasMachine &&
            hasTask &&
            hasFrequency
        ) {

            return rowIndex;
        }
    }


    return -1;
}


// ============================================================
// PRODUCTION LINE FROM EXCEL SHEET NAME
// ============================================================

function getLineFromSheetName(
    sheetName
) {

    const name =
        String(
            sheetName ||
            ""
        ).trim();


    const normalized =
        normalizeText(
            name
        );


    if (
        normalized.includes(
            "mini"
        ) &&
        normalized.includes(
            "cookie"
        )
    ) {

        return "Mini Cookies Line";
    }


    if (
        normalized.includes(
            "cookie"
        )
    ) {

        return "Cookies Line";
    }


    if (
        normalized.includes(
            "roll"
        )
    ) {

        return "Roll Cake Line";
    }


    if (
        normalized.includes(
            "layer"
        )
    ) {

        return "Layer Cake Line";
    }


    if (
        normalized.includes(
            "pound"
        )
    ) {

        return "Pound Cake Line";
    }


    if (
        normalized.includes(
            "pizza"
        )
    ) {

        return "Pizza Line";
    }


    if (
        normalized.includes(
            "cup"
        )
    ) {

        return "Cup Cake Line";
    }


    return name;
}


// ============================================================
// PARSE ACTIVE
// ============================================================

function parseActiveValue(
    value
) {

    if (
        value ===
        undefined ||
        value ===
        null ||
        String(
            value
        ).trim() ===
        ""
    ) {

        return true;
    }


    const normalized =
        normalizeText(
            value
        );


    if (
        [
            "false",
            "no",
            "n",
            "inactive",
            "disabled",
            "0"
        ]
            .includes(
                normalized
            )
    ) {

        return false;
    }


    return true;
}


// ============================================================
// PARSE FREQUENCY
// Supports spelling mistakes and common Excel variations
// ============================================================

function parseFrequency(
    input
) {

    const original =
        String(
            input ??
            ""
        ).trim();


    let text =
        normalizeText(
            original
        );


    // ========================================================
    // EMPTY VALUE
    // ========================================================

    if (
        !text
    ) {

        return {

            valid:
                false,

            error:
                "Frequency is empty"
        };
    }


    // ========================================================
    // CLEAN COMMON TYPING / SPELLING VARIATIONS
    // ========================================================

    text =
        text
            .replace(
                /\s+/g,
                " "
            )
            .trim();


    // ========================================================
    // WEEKLY
    //
    // Supports:
    // weekly
    // weeekly
    // wekly
    // weekely
    // every week
    // ========================================================

    if (
        text ===
            "weekly" ||

        text ===
            "weeekly" ||

        text ===
            "wekly" ||

        text ===
            "weekely" ||

        text ===
            "weekley" ||

        text ===
            "every week" ||

        text ===
            "1 week"
    ) {

        return {

            valid:
                true,

            frequency_text:
                "Weekly",

            frequency_type:
                "weekly",

            frequency_value:
                1
        };
    }


    // ========================================================
    // MONTHLY
    //
    // Supports:
    // monthly
    // montly
    // monthy
    // mounthly
    // every month
    // 1 month
    // ========================================================

    if (
        text ===
            "monthly" ||

        text ===
            "montly" ||

        text ===
            "monthy" ||

        text ===
            "mounthly" ||

        text ===
            "monthley" ||

        text ===
            "every month" ||

        text ===
            "1 month"
    ) {

        return {

            valid:
                true,

            frequency_text:
                "Monthly",

            frequency_type:
                "monthly",

            frequency_value:
                1
        };
    }


    // ========================================================
    // QUARTERLY
    //
    // Supports:
    // quarterly
    // quaterly
    // quartely
    // every 3 months
    // 3 months
    // ========================================================

    if (
        text ===
            "quarterly" ||

        text ===
            "quaterly" ||

        text ===
            "quartely" ||

        text ===
            "quarterley" ||

        text ===
            "every 3 months" ||

        text ===
            "3 months" ||

        text ===
            "3 month"
    ) {

        return {

            valid:
                true,

            frequency_text:
                "Quarterly",

            frequency_type:
                "monthly",

            frequency_value:
                3
        };
    }


    // ========================================================
    // SEMIANNUAL
    //
    // Supports:
    // semiannual
    // semi annual
    // semiannually
    // semi annually
    // half yearly
    // every 6 months
    // 6 months
    // ========================================================

    if (
        text ===
            "semiannual" ||

        text ===
            "semi annual" ||

        text ===
            "semiannually" ||

        text ===
            "semi annually" ||

        text ===
            "semi anual" ||

        text ===
            "semi anualy" ||

        text ===
            "half yearly" ||

        text ===
            "half-yearly" ||

        text ===
            "every 6 months" ||

        text ===
            "6 months" ||

        text ===
            "6 month"
    ) {

        return {

            valid:
                true,

            frequency_text:
                "Semiannual",

            frequency_type:
                "monthly",

            frequency_value:
                6
        };
    }


    // ========================================================
    // ANNUAL
    //
    // Supports:
    // annual
    // annually
    // anually
    // anual
    // yearly
    // every year
    // every 12 months
    // ========================================================

    if (
        text ===
            "annual" ||

        text ===
            "annually" ||

        text ===
            "anually" ||

        text ===
            "anual" ||

        text ===
            "anualy" ||

        text ===
            "yearly" ||

        text ===
            "every year" ||

        text ===
            "every 12 months" ||

        text ===
            "12 months" ||

        text ===
            "12 month"
    ) {

        return {

            valid:
                true,

            frequency_text:
                "Annual",

            frequency_type:
                "monthly",

            frequency_value:
                12
        };
    }

// ========================================================
// ADDITIONAL ANNUAL SPELLING VARIATIONS
// ========================================================

if (
    text === "annualy" ||
    text === "annuelly" ||
    text === "annuel" ||
    text === "year"
) {
    return {
        valid: true,
        frequency_text: "Annual",
        frequency_type: "monthly",
        frequency_value: 12
    };
}


// ========================================================
// YEAR-BASED FREQUENCY
//
// Examples:
// 2 years  = 24 months
// 3 years  = 36 months
// 4 years  = 48 months
// 5 years  = 60 months
// every 3 years
// ========================================================

const yearMatch =
    text.match(
        /(?:every\s*)?(\d+)\s*years?/
    );

if (yearMatch) {

    const years =
        Number(yearMatch[1]);

    if (
        Number.isFinite(years) &&
        years > 0
    ) {

        const months =
            years * 12;

        return {
            valid: true,
            frequency_text:
                `${years} ${years === 1 ? "year" : "years"}`,
            frequency_type: "monthly",
            frequency_value: months
        };
    }
}
    // ========================================================
    // MULTI-YEAR CALENDAR FREQUENCY
    //
    // Examples:
    // 24 months
    // 36 months
    // every 24 months
    // ========================================================

    const monthMatch =
        text.match(
            /(?:every\s*)?(\d+)\s*months?/
        );


    if (
        monthMatch
    ) {

        const months =
            Number(
                monthMatch[
                    1
                ]
            );


        if (
            Number.isFinite(
                months
            ) &&
            months >
                0
        ) {

            return {

                valid:
                    true,

                frequency_text:
                    `${months} months`,

                frequency_type:
                    "monthly",

                frequency_value:
                    months
            };
        }
    }


    // ========================================================
    // HOUR-BASED FREQUENCY
    //
    // Supports:
    // 200 hr
    // 300 hrs
    // 500 hour
    // 1000 hours
    // 2,000 hr
    // 6000hrs
    // ========================================================

    const hourMatch =
        text.match(
            /(\d[\d,]*)\s*(hr|hrs|hour|hours)\b/
        );


    if (
        hourMatch
    ) {

        const hours =
            Number(
                hourMatch[
                    1
                ]
                    .replace(
                        /,/g,
                        ""
                    )
            );


        if (
            Number.isFinite(
                hours
            ) &&
            hours >
                0
        ) {

            return {

                valid:
                    true,

                frequency_text:
                    `${hours} hr`,

                frequency_type:
                    "hours",

                frequency_value:
                    hours
            };
        }
    }


    // ========================================================
    // NUMBER ONLY
    //
    // If Excel contains only a number such as:
    // 500
    // 1000
    // 2000
    //
    // Treat as hour-based frequency.
    // ========================================================

    if (
        /^\d[\d,]*$/
            .test(
                text
            )
    ) {

        const hours =
            Number(
                text
                    .replace(
                        /,/g,
                        ""
                    )
            );


        if (
            Number.isFinite(
                hours
            ) &&
            hours >
                0
        ) {

            return {

                valid:
                    true,

                frequency_text:
                    `${hours} hr`,

                frequency_type:
                    "hours",

                frequency_value:
                    hours
            };
        }
    }


    // ========================================================
    // UNKNOWN
    // ========================================================

    return {

        valid:
            false,

        error:
            `Unknown frequency: ${original}`
    };
}
// ============================================================
// READ EXCEL
// ============================================================

function parseWorkbook(
    buffer
) {

    const workbook =
        XLSX.read(
            buffer,
            {

                type:
                    "buffer",

                cellDates:
                    false
            }
        );


    const parsedRows =
        [];


    const errors =
        [];


    for (
        const sheetName
        of workbook.SheetNames
    ) {

        const sheet =
            workbook.Sheets[
                sheetName
            ];


        const rawRows =
            XLSX.utils
                .sheet_to_json(
                    sheet,
                    {

                        header:
                            1,

                        defval:
                            "",

                        raw:
                            false
                    }
                );


        if (
            rawRows.length ===
            0
        ) {

            continue;
        }


        const headerRowIndex =
            findExcelHeaderRow(
                rawRows
            );


        if (
            headerRowIndex <
            0
        ) {

            continue;
        }


        const headers =
            rawRows[
                headerRowIndex
            ];


        const columnMap =
            {};


        headers.forEach(
            (
                header,
                index
            ) => {

                const field =
                    identifyHeaderField(
                        header
                    );


                if (
                    field &&
                    columnMap[
                        field
                    ] ===
                    undefined
                ) {

                    columnMap[
                        field
                    ] =
                        index;
                }
            }
        );


        const defaultLine =
            getLineFromSheetName(
                sheetName
            );


        for (
            let rowIndex =
                headerRowIndex +
                1;

            rowIndex <
            rawRows.length;

            rowIndex++
        ) {

            const row =
                rawRows[
                    rowIndex
                ];


            const getValue =
                field => {

                    const index =
                        columnMap[
                            field
                        ];


                    if (
                        index ===
                        undefined
                    ) {

                        return "";
                    }


                    return String(
                        row[
                            index
                        ] ??
                        ""
                    ).trim();
                };


            const lineName =
                getValue(
                    "line_name"
                ) ||
                defaultLine;


            const machineCode =
                getValue(
                    "machine_code"
                );


            const machineName =
                getValue(
                    "machine_name"
                );


            const section =
                getValue(
                    "section"
                );


            const partName =
                getValue(
                    "part_name"
                );


            const maintenanceTask =
                getValue(
                    "maintenance_task"
                );


            const frequencyText =
                getValue(
                    "frequency_text"
                );


            const active =
                parseActiveValue(
                    getValue(
                        "active"
                    )
                );


            // =================================================
            // EMPTY ROW
            // =================================================

            if (
                !machineCode &&
                !machineName &&
                !maintenanceTask &&
                !frequencyText
            ) {

                continue;
            }


            const excelRow =
                rowIndex +
                1;


            if (
                !lineName
            ) {

                errors.push({

                    sheet:
                        sheetName,

                    row:
                        excelRow,

                    error:
                        "Production line is missing."
                });


                continue;
            }


            if (
                !machineCode &&
                !machineName
            ) {

                errors.push({

                    sheet:
                        sheetName,

                    row:
                        excelRow,

                    error:
                        "Machine is missing."
                });


                continue;
            }


            if (
                !maintenanceTask
            ) {

                errors.push({

                    sheet:
                        sheetName,

                    row:
                        excelRow,

                    error:
                        "PM task is missing."
                });


                continue;
            }


            const frequency =
                parseFrequency(
                    frequencyText
                );


            if (
                !frequency.valid
            ) {

                errors.push({

                    sheet:
                        sheetName,

                    row:
                        excelRow,

                    error:
                        frequency.error
                });


                continue;
            }


            parsedRows.push({

                sheet_name:
                    sheetName,

                excel_row:
                    excelRow,

                line_name:
                    lineName,

                section,

                machine_code:
                    machineCode,

                machine_name:
                    machineName,

                part_name:
                    partName,

                maintenance_task:
                    maintenanceTask,

                frequency_text:
                    frequency.frequency_text,

                frequency_type:
                    frequency.frequency_type,

                frequency_value:
                    frequency.frequency_value,

                active
            });
        }
    }


    return {

        rows:
            parsedRows,

        errors
    };
}


// ============================================================
// LOAD PM MASTER DATA
// ============================================================

async function getImportMasterData(
    client
) {

    const linesResult =
        await client.query(`
            SELECT
                id,
                line_name

            FROM production_lines

            ORDER BY
                id
        `);


    const machinesResult =
        await client.query(`
            SELECT
                m.id,
                m.production_line_id,
                m.section,
                m.machine_no,
                m.machine_code,
                m.machine_name,
                m.machine_type,
                l.line_name

            FROM machines m

            JOIN production_lines l
                ON
                    m.production_line_id =
                    l.id

            ORDER BY
                m.id
        `);


    const tasksResult =
        await client.query(`
            SELECT
                t.id,
                t.machine_id,
                t.part_name,
                t.maintenance_task,
                t.frequency_text,
                t.frequency_type,
                t.frequency_value,
                t.active,

                m.machine_code,
                m.machine_name,

                l.line_name

            FROM pm_tasks t

            JOIN machines m
                ON
                    t.machine_id =
                    m.id

            JOIN production_lines l
                ON
                    m.production_line_id =
                    l.id

            ORDER BY
                t.id
        `);


    return {

        lines:
            linesResult.rows,

        machines:
            machinesResult.rows,

        tasks:
            tasksResult.rows
    };
}


// ============================================================
// FIND PRODUCTION LINE
// ============================================================

function findLine(
    lineName,
    master
) {

    const normalized =
        normalizeText(
            lineName
        );


    return master.lines.find(
        line =>
            normalizeText(
                line.line_name
            ) ===
            normalized
    ) ||
    null;
}


// ============================================================
// FIND MACHINE
// ============================================================

function findMachine(
    excelRow,
    line,
    master
) {

    const candidates =
        master.machines.filter(
            machine =>
                Number(
                    machine.production_line_id
                ) ===
                Number(
                    line.id
                )
        );


    if (
        excelRow.machine_code
    ) {

        const code =
            normalizeText(
                excelRow.machine_code
            );


        const byCode =
            candidates.find(
                machine =>
                    normalizeText(
                        machine.machine_code
                    ) ===
                    code ||
                    normalizeText(
                        machine.machine_no
                    ) ===
                    code
            );


        if (
            byCode
        ) {

            return byCode;
        }
    }


    if (
        excelRow.machine_name
    ) {

        const name =
            normalizeText(
                excelRow.machine_name
            );


        const byName =
            candidates.find(
                machine =>
                    normalizeText(
                        machine.machine_name
                    ) ===
                    name
            );


        if (
            byName
        ) {

            return byName;
        }
    }


    return null;
}


// ============================================================
// FIND EXISTING PM TASK
// ============================================================

function findExistingTask(
    machineId,
    excelRow,
    master
) {

    const part =
        normalizeText(
            excelRow.part_name
        );


    const maintenance =
        normalizeText(
            excelRow.maintenance_task
        );


    return master.tasks.find(
        task =>

            Number(
                task.machine_id
            ) ===
            Number(
                machineId
            ) &&

            normalizeText(
                task.part_name
            ) ===
            part &&

            normalizeText(
                task.maintenance_task
            ) ===
            maintenance
    ) ||
    null;
}


// ============================================================
// CHECK FREQUENCY CHANGE
// ============================================================

function frequencyChanged(
    existing,
    imported
) {

    return (

        normalizeText(
            existing.frequency_type
        ) !==
        normalizeText(
            imported.frequency_type
        ) ||

        Number(
            existing.frequency_value
        ) !==
        Number(
            imported.frequency_value
        )
    );
}


// ============================================================
// BUILD EXCEL PREVIEW
// ============================================================

async function buildImportPreview(
    parsedWorkbook
) {

    const client =
        await pool.connect();


    try {

        const master =
            await getImportMasterData(
                client
            );


        const preview =
            [];


        const operations =
            [];


        let newTasks =
            0;


        let updatedTasks =
            0;


        let unchangedTasks =
            0;


        let inactiveTasks =
            0;


        let errorCount =
            parsedWorkbook.errors.length;


        // ====================================================
        // EXCEL ERRORS
        // ====================================================

        for (
            const error
            of parsedWorkbook.errors
        ) {

            preview.push({

                action:
                    "Error",

                line_name:
                    "",

                machine_code:
                    "",

                machine_name:
                    "",

                section:
                    "",

                part_name:
                    "",

                maintenance_task:
                    `Sheet ${error.sheet} / Row ${error.row}`,

                frequency_text:
                    "",

                validation:
                    error.error
            });
        }


        // ====================================================
        // EXCEL RECORDS
        // ====================================================

        for (
            const row
            of parsedWorkbook.rows
        ) {

            const line =
                findLine(
                    row.line_name,
                    master
                );


            if (
                !line
            ) {

                errorCount++;


                preview.push({

                    action:
                        "Error",

                    ...row,

                    validation:
                        `Production line not found: ${row.line_name}`
                });


                continue;
            }


            const machine =
                findMachine(
                    row,
                    line,
                    master
                );


          if (
    !machine
) {

    newTasks++;

    operations.push({
        type: "insert_new_machine_task",
        line_id: line.id,
        row: row
    });

    preview.push({
        action: "New",
        ...row,
        validation:
            `Machine ${row.machine_code || row.machine_no || row.machine_name} will be created in ${line.line_name} and PM task will be added.`
    });

    continue;
}


            const existing =
                findExistingTask(
                    machine.id,
                    row,
                    master
                );


            // =================================================
            // INACTIVE
            // =================================================

            if (
                row.active ===
                false
            ) {

                if (
                    existing &&
                    existing.active
                ) {

                    inactiveTasks++;


                    operations.push({

                        type:
                            "deactivate",

                        task_id:
                            existing.id
                    });


                    preview.push({

                        action:
                            "Inactive",

                        ...row,

                        machine_code:
                            machine.machine_code,

                        machine_name:
                            machine.machine_name,

                        validation:
                            "PM task will be made inactive."
                    });


                } else {

                    unchangedTasks++;


                    preview.push({

                        action:
                            "Unchanged",

                        ...row,

                        machine_code:
                            machine.machine_code,

                        machine_name:
                            machine.machine_name,

                        validation:
                            "Already inactive or does not exist."
                    });
                }


                continue;
            }


            // =================================================
            // NEW PM
            // =================================================

            if (
                !existing
            ) {

                newTasks++;


                operations.push({

                    type:
                        "insert",

                    machine_id:
                        machine.id,

                    row
                });


                preview.push({

                    action:
                        "New",

                    ...row,

                    machine_code:
                        machine.machine_code,

                    machine_name:
                        machine.machine_name,

                    validation:
                        "New PM task will be added."
                });


                continue;
            }


            // =================================================
            // REACTIVATE
            // =================================================

            if (
                !existing.active &&
                !frequencyChanged(
                    existing,
                    row
                )
            ) {

                updatedTasks++;


                operations.push({

                    type:
                        "reactivate",

                    task_id:
                        existing.id,

                    row
                });


                preview.push({

                    action:
                        "Updated",

                    ...row,

                    machine_code:
                        machine.machine_code,

                    machine_name:
                        machine.machine_name,

                    validation:
                        "Inactive PM task will be reactivated."
                });


                continue;
            }


            // =================================================
            // FREQUENCY CHANGE
            //
            // OLD TASK REMAINS FOR HISTORY.
            // NEW TASK VERSION IS CREATED.
            // =================================================

            if (
                frequencyChanged(
                    existing,
                    row
                )
            ) {

                updatedTasks++;


                operations.push({

                    type:
                        "version",

                    old_task_id:
                        existing.id,

                    machine_id:
                        machine.id,

                    row
                });


                preview.push({

                    action:
                        "Updated",

                    ...row,

                    machine_code:
                        machine.machine_code,

                    machine_name:
                        machine.machine_name,

                    validation:
                        `Frequency changes from ${existing.frequency_text} to ${row.frequency_text}. Historical PM records will remain.`
                });


                continue;
            }


            // =================================================
            // UNCHANGED
            // =================================================

            unchangedTasks++;


            preview.push({

                action:
                    "Unchanged",

                ...row,

                machine_code:
                    machine.machine_code,

                machine_name:
                    machine.machine_name,

                validation:
                    "No change."
            });
        }


        return {

            preview,

            operations,

            summary: {

                total_rows:
                    parsedWorkbook.rows.length +
                    parsedWorkbook.errors.length,

                new_tasks:
                    newTasks,

                updated_tasks:
                    updatedTasks,

                unchanged_tasks:
                    unchangedTasks,

                inactive_tasks:
                    inactiveTasks,

                errors:
                    errorCount
            }
        };


    } finally {

        client.release();
    }
}


// ============================================================
// DELETE FUTURE PENDING SCHEDULES ONLY
//
// COMPLETED AND DEFERRED HISTORY REMAINS.
// ============================================================

async function removeFuturePendingSchedules(
    client,
    taskId
) {

    const now =
        new Date();


    const currentYear =
        now.getFullYear();


    const currentWeek =
        getISOWeekNumber(
            now
        );


    await client.query(
        `
        DELETE FROM pm_schedule

        WHERE
            pm_task_id = $1

        AND
            status = 'Pending'

        AND
        (
            planned_year > $2

            OR

            (
                planned_year = $2

                AND
                    planned_week >= $3
            )
        )
        `,
        [
            taskId,
            currentYear,
            currentWeek
        ]
    );
}


// ============================================================
// APPLY EXCEL IMPORT
// ============================================================

async function applyImportOperations(
    operations
) {

    const client =
        await pool.connect();


    let inserted =
        0;


    let updated =
        0;


    let inactive =
        0;


    try {

        await client.query(
            "BEGIN"
        );


        for (
            const operation
            of operations
        ) {
// =================================================
// CREATE MISSING MACHINE + PM TASK
// =================================================

if (
    operation.type ===
    "insert_new_machine_task"
) {

    const row =
        operation.row;


    const machineCode =
        String(
            row.machine_code ||
            row.machine_no ||
            row.machine_name ||
            ""
        ).trim();


    const machineName =
        String(
            row.machine_name ||
            row.machine_code ||
            row.machine_no ||
            ""
        ).trim();


    if (
        !machineCode ||
        !machineName
    ) {

        throw new Error(
            `Cannot create machine for Excel row ${row.excel_row || ""}`
        );
    }


    const machineResult =
        await client.query(
            `
            INSERT INTO machines
            (
                production_line_id,
                section,
                machine_no,
                machine_name,
                machine_code,
                machine_type
            )

            VALUES
            (
                $1,
                $2,
                $3,
                $4,
                $5,
                $6
            )

            ON CONFLICT
            (
                machine_code
            )

            DO UPDATE SET
                production_line_id =
                    EXCLUDED.production_line_id,

                section =
                    EXCLUDED.section,

                machine_name =
                    EXCLUDED.machine_name,

                updated_at =
                    CURRENT_TIMESTAMP

            RETURNING id
            `,
            [
                operation.line_id,
                row.section || null,
                machineCode,
                machineName,
                machineCode,
                null
            ]
        );


    const machineId =
        machineResult.rows[0].id;


    await client.query(
        `
        INSERT INTO pm_tasks
        (
            machine_id,
            part_name,
            maintenance_task,
            frequency_text,
            frequency_type,
            frequency_value,
            active
        )

        VALUES
        (
            $1,
            $2,
            $3,
            $4,
            $5,
            $6,
            TRUE
        )
        `,
        [
            machineId,
            row.part_name || null,
            row.maintenance_task,
            row.frequency_text,
            row.frequency_type,
            row.frequency_value
        ]
    );


    inserted++;

    continue;
}
            // =================================================
            // INSERT
            // =================================================

            if (
                operation.type ===
                "insert"
            ) {

                await client.query(
                    `
                    INSERT INTO pm_tasks
                    (
                        machine_id,
                        part_name,
                        maintenance_task,
                        frequency_text,
                        frequency_type,
                        frequency_value,
                        active
                    )

                    VALUES
                    (
                        $1,
                        $2,
                        $3,
                        $4,
                        $5,
                        $6,
                        TRUE
                    )
                    `,
                    [
                        operation.machine_id,

                        operation.row.part_name ||
                        null,

                        operation.row.maintenance_task,

                        operation.row.frequency_text,

                        operation.row.frequency_type,

                        operation.row.frequency_value
                    ]
                );


                inserted++;

                continue;
            }


            // =================================================
            // REACTIVATE
            // =================================================

            if (
                operation.type ===
                "reactivate"
            ) {

                await client.query(
                    `
                    UPDATE pm_tasks

                    SET
                        active =
                            TRUE,

                        frequency_text =
                            $1,

                        frequency_type =
                            $2,

                        frequency_value =
                            $3

                    WHERE
                        id =
                            $4
                    `,
                    [
                        operation.row.frequency_text,

                        operation.row.frequency_type,

                        operation.row.frequency_value,

                        operation.task_id
                    ]
                );


                updated++;

                continue;
            }


            // =================================================
            // DEACTIVATE
            // =================================================

            if (
                operation.type ===
                "deactivate"
            ) {

                await client.query(
                    `
                    UPDATE pm_tasks

                    SET
                        active =
                            FALSE

                    WHERE
                        id =
                            $1
                    `,
                    [
                        operation.task_id
                    ]
                );


                await removeFuturePendingSchedules(
                    client,
                    operation.task_id
                );


                inactive++;

                continue;
            }


            // =================================================
            // CREATE NEW VERSION
            // =================================================

            if (
                operation.type ===
                "version"
            ) {

                await client.query(
                    `
                    UPDATE pm_tasks

                    SET
                        active =
                            FALSE

                    WHERE
                        id =
                            $1
                    `,
                    [
                        operation.old_task_id
                    ]
                );


                await removeFuturePendingSchedules(
                    client,
                    operation.old_task_id
                );


                await client.query(
                    `
                    INSERT INTO pm_tasks
                    (
                        machine_id,
                        part_name,
                        maintenance_task,
                        frequency_text,
                        frequency_type,
                        frequency_value,
                        active
                    )

                    VALUES
                    (
                        $1,
                        $2,
                        $3,
                        $4,
                        $5,
                        $6,
                        TRUE
                    )
                    `,
                    [
                        operation.machine_id,

                        operation.row.part_name ||
                        null,

                        operation.row.maintenance_task,

                        operation.row.frequency_text,

                        operation.row.frequency_type,

                        operation.row.frequency_value
                    ]
                );


                updated++;

                continue;
            }
        }


        await client.query(
            "COMMIT"
        );


        return {

            inserted,

            updated,

            inactive
        };


    } catch (
        error
    ) {

        await client.query(
            "ROLLBACK"
        );


        throw error;


    } finally {

        client.release();
    }
}


// ============================================================
// CLEAN OLD EXCEL PREVIEWS
// ============================================================

function cleanupImportPreviews() {

    const now =
        Date.now();


    for (
        const [
            token,
            preview
        ]
        of importPreviews.entries()
    ) {

        if (
            now -
            preview.created_at >
            IMPORT_PREVIEW_LIFETIME
        ) {

            importPreviews.delete(
                token
            );
        }
    }
}


// ============================================================
// DATABASE TEST
// ============================================================

app.get(
    "/api/test-db",
    async (
        req,
        res
    ) => {

        try {

            const result =
                await pool.query(
                    "SELECT NOW()"
                );


            res.json({

                connected:
                    true,

                databaseTime:
                    result.rows[
                        0
                    ].now
            });


        } catch (
            error
        ) {

            console.error(
                error
            );


            res
                .status(
                    500
                )
                .json({

                    connected:
                        false,

                    error:
                        error.message
                });
        }
    }
);


// ============================================================
// STATISTICS
// ============================================================

app.get(
    "/api/stats",
    async (
        req,
        res
    ) => {

        try {

            const lines =
                await pool.query(`
                    SELECT
                        COUNT(*)::int AS count
                    FROM
                        production_lines
                `);


            const machines =
                await pool.query(`
                    SELECT
                        COUNT(*)::int AS count
                    FROM
                        machines
                `);


            const tasks =
                await pool.query(`
                    SELECT
                        COUNT(*)::int AS count
                    FROM
                        pm_tasks
                    WHERE
                        active = TRUE
                `);


            res.json({

                productionLines:
                    lines.rows[
                        0
                    ].count,

                machines:
                    machines.rows[
                        0
                    ].count,

                pmTasks:
                    tasks.rows[
                        0
                    ].count
            });


        } catch (
            error
        ) {

            console.error(
                error
            );


            res
                .status(
                    500
                )
                .json({

                    error:
                        error.message
                });
        }
    }
);


// ============================================================
// PRODUCTION LINES
// ============================================================

app.get(
    "/api/lines",
    async (
        req,
        res
    ) => {

        try {

            const result =
                await pool.query(`
                    SELECT
                        id,
                        line_name

                    FROM
                        production_lines

                    ORDER BY
                        id
                `);


            res.json(
                result.rows
            );


        } catch (
            error
        ) {

            console.error(
                error
            );


            res
                .status(
                    500
                )
                .json({

                    error:
                        error.message
                });
        }
    }
);


// ============================================================
// MACHINES BY LINE
// ============================================================

app.get(
    "/api/lines/:lineId/machines",
    async (
        req,
        res
    ) => {

        try {

            const lineId =
                Number(
                    req.params.lineId
                );


            const result =
                await pool.query(
                    `
                    SELECT
                        id,
                        section,
                        machine_no,
                        machine_name,
                        machine_code,
                        machine_type

                    FROM
                        machines

                    WHERE
                        production_line_id =
                            $1

                    ORDER BY
                        section,
                        machine_no,
                        machine_name
                    `,
                    [
                        lineId
                    ]
                );


            res.json(
                result.rows
            );


        } catch (
            error
        ) {

            console.error(
                error
            );


            res
                .status(
                    500
                )
                .json({

                    error:
                        error.message
                });
        }
    }
);


// ============================================================
// MACHINE PM MASTER TASKS
// ============================================================

app.get(
    "/api/machines/:machineId/tasks",
    async (
        req,
        res
    ) => {

        try {

            const machineId =
                Number(
                    req.params.machineId
                );


            const result =
                await pool.query(
                    `
                    SELECT
                        id,
                        part_name,
                        maintenance_task,
                        frequency_text,
                        frequency_type,
                        frequency_value,
                        active

                    FROM
                        pm_tasks

                    WHERE
                        machine_id =
                            $1

                    AND
                        active =
                            TRUE

                    ORDER BY
                        id
                    `,
                    [
                        machineId
                    ]
                );


            res.json(
                result.rows
            );


        } catch (
            error
        ) {

            console.error(
                error
            );


            res
                .status(
                    500
                )
                .json({

                    error:
                        error.message
                });
        }
    }
);


// ============================================================
// WEEKLY PM SCHEDULE
// ============================================================

app.get(
    "/api/schedule",
    async (
        req,
        res
    ) => {

        try {

            const year =
                Number(
                    req.query.year
                );


            const week =
                Number(
                    req.query.week
                );


            if (
                !year ||
                !week
            ) {

                return res
                    .status(
                        400
                    )
                    .json({

                        error:
                            "year and week are required."
                    });
            }


            const result =
                await pool.query(
                    `
                    SELECT

                        s.id,

                        l.id AS line_id,
                        l.line_name,

                        m.id AS machine_id,
                        m.machine_code,
                        m.machine_name,
                        m.section,
                        m.machine_type,

                        t.id AS pm_task_id,
                        t.part_name,
                        t.maintenance_task,
                        t.frequency_text,
                        t.frequency_type,
                        t.frequency_value,

                        s.planned_year,
                        s.planned_week,
                        s.status,

                        s.completed_at,
                        s.technician_name,
                        s.deferred_reason,
                        s.notes

                    FROM
                        pm_schedule s

                    JOIN pm_tasks t
                        ON
                            s.pm_task_id =
                            t.id

                    JOIN machines m
                        ON
                            t.machine_id =
                            m.id

                    JOIN production_lines l
                        ON
                            m.production_line_id =
                            l.id

                    WHERE
                        s.planned_year =
                            $1

                    AND
                        s.planned_week =
                            $2

                    ORDER BY
                        l.id,
                        m.section,
                        m.machine_code,
                        t.id
                    `,
                    [
                        year,
                        week
                    ]
                );


            res.json(
                result.rows
            );


        } catch (
            error
        ) {

            console.error(
                error
            );


            res
                .status(
                    500
                )
                .json({

                    error:
                        error.message
                });
        }
    }
);


// ============================================================
// GENERATE WEEK
//
// USER MUST BE LOGGED IN.
// GENERATE SCHEDULE USES CORRECTED ENGINE.
// ============================================================

app.post(
    "/api/schedule/generate",
    async (
        req,
        res
    ) => {

        const year =
            Number(
                req.body.year
            );


        const week =
            Number(
                req.body.week
            );


        const maxWeek =
            year
                ?
                getISOWeeksInYear(
                    year
                )
                :
                0;


        if (
            !year ||
            !week ||
            week <
            1 ||
            week >
            maxWeek
        ) {

            return res
                .status(
                    400
                )
                .json({

                    error:
                        "Valid year and ISO week are required."
                });
        }


        const client =
            await pool.connect();


        try {

            await client.query(
                "BEGIN"
            );


            const result =
                await generateWeekSchedule(
                    client,
                    year,
                    week
                );


            await client.query(
                "COMMIT"
            );


            res.json({

                message:
                    "Schedule generated successfully",

                year,

                week,

                month:
                    getMonthForWeek(
                        year,
                        week
                    ),

                created:
                    result.created,

                already_exists:
                    result.alreadyExists,

                not_due:
                    result.notDue,

                total_pm_master:
                    result.totalMaster
            });


        } catch (
            error
        ) {

            try {

                await client.query(
                    "ROLLBACK"
                );

            } catch (
                rollbackError
            ) {

                console.error(
                    rollbackError
                );
            }


            console.error(
                "Schedule generation failed:",
                error
            );


            res
                .status(
                    500
                )
                .json({

                    error:
                        error.message
                });


        } finally {

            client.release();
        }
    }
);


// ============================================================
// SAFE YEAR REBUILD
//
// ADMIN ONLY
//
// COMPLETED + DEFERRED ARE PRESERVED.
// ONLY PENDING IS REBUILT.
// ============================================================

app.post(
    "/api/schedule/rebuild-year",

    requireRole(
        "admin"
    ),

    async (
        req,
        res
    ) => {

        const year =
            Number(
                req.body.year
            );


        if (
            !year ||
            year <
            2000 ||
            year >
            2100
        ) {

            return res
                .status(
                    400
                )
                .json({

                    error:
                        "Valid year is required."
                });
        }


        const client =
            await pool.connect();


        try {

            await client.query(
                "BEGIN"
            );


            const result =
                await rebuildYear(
                    client,
                    year
                );


            await client.query(
                "COMMIT"
            );


            const distribution =
                await checkDistribution(
                    client,
                    year
                );


            const problems =
                validateDistribution(
                    distribution,
                    year
                );


            res.json({

                message:
                    "PM schedule rebuilt successfully",

                ...result,

                validation: {

                    passed:
                        problems.length ===
                        0,

                    problem_count:
                        problems.length,

                    problems:
                        problems.slice(
                            0,
                            100
                        )
                }
            });


        } catch (
            error
        ) {

            try {

                await client.query(
                    "ROLLBACK"
                );

            } catch (
                rollbackError
            ) {

                console.error(
                    rollbackError
                );
            }


            console.error(
                "Schedule rebuild failed:",
                error
            );


            res
                .status(
                    500
                )
                .json({

                    error:
                        error.message
                });


        } finally {

            client.release();
        }
    }
);


// ============================================================
// CHECK PM DISTRIBUTION
// ============================================================

app.get(
    "/api/schedule/check-distribution",
    async (
        req,
        res
    ) => {

        const year =
            Number(
                req.query.year
            );


        if (
            !year
        ) {

            return res
                .status(
                    400
                )
                .json({

                    error:
                        "year is required."
                });
        }


        const client =
            await pool.connect();


        try {

            const distribution =
                await checkDistribution(
                    client,
                    year
                );


            const problems =
                validateDistribution(
                    distribution,
                    year
                );


            res.json({

                year,

                passed:
                    problems.length ===
                    0,

                problem_count:
                    problems.length,

                problems
            });


        } catch (
            error
        ) {

            console.error(
                error
            );


            res
                .status(
                    500
                )
                .json({

                    error:
                        error.message
                });


        } finally {

            client.release();
        }
    }
);


// ============================================================
// ANNUAL PLAN
// ============================================================

app.get(
    "/api/annual-plan",
    async (
        req,
        res
    ) => {

        const year =
            Number(
                req.query.year
            );


        if (
            !year
        ) {

            return res
                .status(
                    400
                )
                .json({

                    error:
                        "year is required."
                });
        }


        const client =
            await pool.connect();


        try {

            // =================================================
            // GENERATE MISSING PM RECORDS
            // =================================================

            await client.query(
                "BEGIN"
            );


            const generation =
                await generateFullYear(
                    client,
                    year
                );


            await client.query(
                "COMMIT"
            );


            // =================================================
            // LOAD COMPLETE ANNUAL PLAN
            // =================================================

            const annualResult =
                await client.query(
                    `
                    SELECT

                        s.id,

                        s.planned_year,
                        s.planned_week,

                        l.id AS line_id,
                        l.line_name,

                        m.id AS machine_id,
                        m.machine_code,
                        m.machine_name,
                        m.section,
                        m.machine_type,

                        t.id AS pm_task_id,
                        t.part_name,
                        t.maintenance_task,
                        t.frequency_text,
                        t.frequency_type,
                        t.frequency_value,

                        s.status,
                        s.completed_at,
                        s.technician_name,
                        s.deferred_reason,
                        s.notes

                    FROM
                        pm_schedule s

                    JOIN pm_tasks t
                        ON
                            s.pm_task_id =
                            t.id

                    JOIN machines m
                        ON
                            t.machine_id =
                            m.id

                    JOIN production_lines l
                        ON
                            m.production_line_id =
                            l.id

                    WHERE
                        s.planned_year =
                            $1

                    ORDER BY
                        s.planned_week,
                        l.id,
                        m.section,
                        m.machine_code,
                        t.id
                    `,
                    [
                        year
                    ]
                );


            const rows =
                annualResult.rows;


            res.json({

                year,

                summary: {

                    total:
                        rows.length,

                    completed:
                        rows.filter(
                            item =>
                                item.status ===
                                "Completed"
                        ).length,

                    pending:
                        rows.filter(
                            item =>
                                item.status ===
                                "Pending"
                        ).length,

                    deferred:
                        rows.filter(
                            item =>
                                item.status ===
                                "Deferred"
                        ).length
                },

                generation,

                schedule:
                    rows
            });


        } catch (
            error
        ) {

            try {

                await client.query(
                    "ROLLBACK"
                );

            } catch (
                rollbackError
            ) {

                console.error(
                    rollbackError
                );
            }


            console.error(
                "Annual plan failed:",
                error
            );


            res
                .status(
                    500
                )
                .json({

                    error:
                        error.message
                });


        } finally {

            client.release();
        }
    }
);


// ============================================================
// COMPLETE PM
//
// ADMIN / SUPERVISOR / TECHNICIAN
// ============================================================

app.put(
    "/api/schedule/:id/complete",

    requireRole(
        "admin",
        "supervisor",
        "technician"
    ),

    async (
        req,
        res
    ) => {

        try {

            const scheduleId =
                Number(
                    req.params.id
                );


            const technicianName =
                String(
                    req.body.technician_name ||
                    ""
                )
                    .trim();


            const notes =
                String(
                    req.body.notes ||
                    ""
                )
                    .trim();


            if (
                !technicianName
            ) {

                return res
                    .status(
                        400
                    )
                    .json({

                        error:
                            "Technician name is required."
                    });
            }


            const result =
                await pool.query(
                    `
                    UPDATE
                        pm_schedule

                    SET
                        status =
                            'Completed',

                        completed_at =
                            CURRENT_TIMESTAMP,

                        technician_name =
                            $1,

                        notes =
                            $2,

                        deferred_reason =
                            NULL,

                        updated_at =
                            CURRENT_TIMESTAMP

                    WHERE
                        id =
                            $3

                    RETURNING
                        *
                    `,
                    [
                        technicianName,

                        notes ||
                        null,

                        scheduleId
                    ]
                );


            if (
                result.rows.length ===
                0
            ) {

                return res
                    .status(
                        404
                    )
                    .json({

                        error:
                            "PM schedule not found."
                    });
            }


            res.json({

                message:
                    "PM completed successfully",

                schedule:
                    result.rows[
                        0
                    ]
            });


        } catch (
            error
        ) {

            console.error(
                "Complete PM error:",
                error
            );


            res
                .status(
                    500
                )
                .json({

                    error:
                        error.message
                });
        }
    }
);


// ============================================================
// DEFER PM
//
// ADMIN / SUPERVISOR / TECHNICIAN
// ============================================================

app.put(
    "/api/schedule/:id/defer",

    requireRole(
        "admin",
        "supervisor",
        "technician"
    ),

    async (
        req,
        res
    ) => {

        try {

            const scheduleId =
                Number(
                    req.params.id
                );


            const technicianName =
                String(
                    req.body.technician_name ||
                    ""
                )
                    .trim();


            const deferredReason =
                String(
                    req.body.deferred_reason ||
                    ""
                )
                    .trim();


            const notes =
                String(
                    req.body.notes ||
                    ""
                )
                    .trim();


            if (
                !technicianName
            ) {

                return res
                    .status(
                        400
                    )
                    .json({

                        error:
                            "Technician name is required."
                    });
            }


            if (
                !deferredReason
            ) {

                return res
                    .status(
                        400
                    )
                    .json({

                        error:
                            "Deferred reason is required."
                    });
            }


            const result =
                await pool.query(
                    `
                    UPDATE
                        pm_schedule

                    SET
                        status =
                            'Deferred',

                        deferred_reason =
                            $1,

                        technician_name =
                            $2,

                        notes =
                            $3,

                        completed_at =
                            NULL,

                        updated_at =
                            CURRENT_TIMESTAMP

                    WHERE
                        id =
                            $4

                    RETURNING
                        *
                    `,
                    [
                        deferredReason,

                        technicianName,

                        notes ||
                        null,

                        scheduleId
                    ]
                );


            if (
                result.rows.length ===
                0
            ) {

                return res
                    .status(
                        404
                    )
                    .json({

                        error:
                            "PM schedule not found."
                    });
            }


            res.json({

                message:
                    "PM deferred successfully",

                schedule:
                    result.rows[
                        0
                    ]
            });


        } catch (
            error
        ) {

            console.error(
                "Defer PM error:",
                error
            );


            res
                .status(
                    500
                )
                .json({

                    error:
                        error.message
                });
        }
    }
);


// ============================================================
// EXCEL IMPORT PREVIEW
//
// ADMIN ONLY
// ============================================================

app.post(
    "/api/import/preview",

    requireRole(
        "admin"
    ),

    upload.single(
        "file"
    ),

    async (
        req,
        res
    ) => {

        try {

            cleanupImportPreviews();


            if (
                !req.file
            ) {

                return res
                    .status(
                        400
                    )
                    .json({

                        error:
                            "Excel file is required."
                    });
            }


            const parsed =
                parseWorkbook(
                    req.file.buffer
                );


            if (
                parsed.rows.length ===
                0 &&
                parsed.errors.length ===
                0
            ) {

                return res
                    .status(
                        400
                    )
                    .json({

                        error:
                            "No recognizable PM data was found. Required columns include Machine, PM Task and Frequency."
                    });
            }


            const previewResult =
                await buildImportPreview(
                    parsed
                );


            const token =
                crypto
                    .randomBytes(
                        24
                    )
                    .toString(
                        "hex"
                    );


            importPreviews.set(
                token,
                {

                    created_at:
                        Date.now(),

                    filename:
                        req.file.originalname,

                    operations:
                        previewResult.operations,

                    summary:
                        previewResult.summary
                }
            );


            res.json({

                message:
                    "Excel validation completed",

                filename:
                    req.file.originalname,

                preview_token:
                    token,

                summary:
                    previewResult.summary,

                preview:
                    previewResult.preview.slice(
                        0,
                        1000
                    ),

                preview_limited:
                    previewResult.preview.length >
                    1000
            });


        } catch (
            error
        ) {

            console.error(
                "Excel preview error:",
                error
            );


            res
                .status(
                    500
                )
                .json({

                    error:
                        error.message
                });
        }
    }
);


// ============================================================
// EXCEL IMPORT CONFIRM
//
// ADMIN ONLY
// ============================================================

app.post(
    "/api/import/confirm",

    requireRole(
        "admin"
    ),

    async (
        req,
        res
    ) => {

        try {

            cleanupImportPreviews();


            const token =
                String(
                    req.body.preview_token ||
                    ""
                );


            if (
                !token
            ) {

                return res
                    .status(
                        400
                    )
                    .json({

                        error:
                            "preview_token is required."
                    });
            }


            const preview =
                importPreviews.get(
                    token
                );


            if (
                !preview
            ) {

                return res
                    .status(
                        400
                    )
                    .json({

                        error:
                            "Import preview expired. Please preview the Excel file again."
                    });
            }

// ============================================================
// ALLOW IMPORT OF VALID ROWS
// INVALID ROWS ARE SKIPPED
// ============================================================

if (
    Number(
        preview.summary.errors
    ) >
    0
) {

    console.warn(
        `Excel import contains ${preview.summary.errors} validation error(s). Invalid rows will be skipped.`
    );
}


            const result =
                await applyImportOperations(
                    preview.operations
                );


            importPreviews.delete(
                token
            );


            res.json({

                message:
                    "PM Master updated successfully",

                historical_records_preserved:
                    true,

                summary: {

                    total_rows:
                        preview.summary.total_rows,

                    new_tasks:
                        result.inserted,

                    updated_tasks:
                        result.updated,

                    unchanged_tasks:
                        preview.summary.unchanged_tasks,

                    inactive_tasks:
                        result.inactive,

                    errors:
                        0
                }
            });


        } catch (
            error
        ) {

            console.error(
                "Excel import error:",
                error
            );


            res
                .status(
                    500
                )
                .json({

                    error:
                        error.message
                });
        }
    }
);


// ============================================================
// MULTER / GENERAL ERROR HANDLER
// ============================================================

app.use(
    (
        error,
        req,
        res,
        next
    ) => {

        if (
            error instanceof
            multer.MulterError
        ) {

            return res
                .status(
                    400
                )
                .json({

                    error:
                        error.message
                });
        }


        if (
            error
        ) {

            return res
                .status(
                    400
                )
                .json({

                    error:
                        error.message
                });
        }


        next();
    }
);


// ============================================================
// INITIALIZE DATABASE
// ============================================================

async function initializeDatabase() {

    try {

        const schemaPath =
            path.join(
                __dirname,
                "../database/schema.sql"
            );


        const schema =
            fs.readFileSync(
                schemaPath,
                "utf8"
            );


        await pool.query(
            schema
        );
// ============================================================
// ENSURE PRODUCTION LINES EXIST
// ============================================================

const productionLines = [
    "Cookies Line",
    "Roll Cake Line",
    "Layer Cake Line",
    "Mini Cookies Line",
    "Pound Cake Line",
    "Pizza Line"
];

for (
    const lineName
    of productionLines
) {

    await pool.query(
        `
        INSERT INTO production_lines
        (
            line_name
        )
        VALUES
        (
            $1
        )
        ON CONFLICT
        (
            line_name
        )
        DO NOTHING
        `,
        [
            lineName
        ]
    );
}

console.log(
    "Production lines initialized successfully"
);

        console.log(
            "Database tables initialized successfully"
        );


    } catch (
        error
    ) {

        console.error(
            "Database initialization failed:",
            error
        );


        throw error;
    }
}


// ============================================================
// START SERVER
// ============================================================

const PORT =
    process.env.PORT ||
    3000;

// ============================================================
// TEMPORARY PM DUPLICATE DIAGNOSTIC
// ============================================================

app.get("/api/debug/pm-duplicates", async (req, res) => {
    try {
        const result = await pool.query(`
            SELECT
                planned_year,
                planned_week,
                pm_task_id,
                COUNT(*) AS copies
            FROM pm_schedule
            GROUP BY
                planned_year,
                planned_week,
                pm_task_id
            HAVING COUNT(*) > 1
            ORDER BY copies DESC, planned_week
            LIMIT 100
        `);

        res.json({
            duplicateGroups: result.rows.length,
            duplicates: result.rows
        });
    } catch (error) {
        console.error("Duplicate diagnostic error:", error);

        res.status(500).json({
            error: error.message
        });
    }
});

// ============================================================
// ADMIN - CLEAN DUPLICATE PM SCHEDULE RECORDS
// Keeps one record for each PM task + year + week
// Priority: Completed > Deferred > Pending
// ============================================================

app.post(
    "/api/admin/cleanup-duplicate-schedules",
    requireRole("admin"),
    async (req, res) => {

        try {

            const result = await pool.query(`
                WITH ranked AS
                (
                    SELECT
                        id,

                        ROW_NUMBER() OVER
                        (
                            PARTITION BY
                                pm_task_id,
                                planned_year,
                                planned_week

                            ORDER BY
                                CASE status
                                    WHEN 'Completed' THEN 1
                                    WHEN 'Deferred' THEN 2
                                    ELSE 3
                                END,
                                id
                        ) AS rn

                    FROM pm_schedule
                )

                DELETE FROM pm_schedule

                WHERE id IN
                (
                    SELECT id
                    FROM ranked
                    WHERE rn > 1
                )

                RETURNING id
            `);

            res.json({
                success: true,
                deleted_duplicates: result.rowCount
            });

        } catch (error) {

            console.error(
                "Duplicate schedule cleanup failed:",
                error
            );

            res.status(500).json({
                success: false,
                error: error.message
            });
        }
    }
);
app.listen(
    PORT,
    async () => {

        console.log(
            ""
        );


        console.log(
            "=============================================="
        );


        console.log(
            " NATIONAL FOOD COMPANY - AMERICANA CAKE"
        );


        console.log(
            " PREVENTIVE MAINTENANCE SYSTEM"
        );


        console.log(
            "=============================================="
        );


        console.log(
            `PM Server running on port ${PORT}`
        );


        console.log(
            "Using corrected PM distribution engine"
        );


        // ====================================================
        // DATABASE
        // ====================================================

        try {

            await initializeDatabase();

        } catch (
            error
        ) {

            console.error(
                "Server database initialization error:",
                error
            );
        }


        // ====================================================
        // ADMIN ACCOUNT
        // ====================================================

        try {

            await ensureDefaultAdmin();

        } catch (
            error
        ) {

            console.error(
                "Default admin creation failed:",
                error
            );
        }


        console.log(
            ""
        );


        console.log(
            `Open: http://localhost:${PORT}`
        );


        console.log(
            "=============================================="
        );


        console.log(
            ""
        );
    }
);
