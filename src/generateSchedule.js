require("dotenv").config();

const pool = require("./database");


// ============================================================
// CONFIGURATION
// ============================================================

const BASE_YEAR = 2026;


// ============================================================
// ISO WEEK NUMBER
// ============================================================

function getISOWeekNumber(date) {

    const workingDate =
        new Date(
            Date.UTC(
                date.getFullYear(),
                date.getMonth(),
                date.getDate()
            )
        );


    const day =
        workingDate.getUTCDay() || 7;


    workingDate.setUTCDate(
        workingDate.getUTCDate() +
        4 -
        day
    );


    const yearStart =
        new Date(
            Date.UTC(
                workingDate.getUTCFullYear(),
                0,
                1
            )
        );


    return Math.ceil(
        (
            (
                workingDate -
                yearStart
            ) /
            86400000 +
            1
        ) /
        7
    );
}


// ============================================================
// ISO WEEKS IN YEAR
// ============================================================

function getISOWeeksInYear(year) {

    const december28 =
        new Date(
            year,
            11,
            28
        );


    return getISOWeekNumber(
        december28
    );
}


// ============================================================
// GET MONDAY OF ISO WEEK
// ============================================================

function getISOWeekMonday(
    year,
    week
) {

    const january4 =
        new Date(
            year,
            0,
            4
        );


    const january4Day =
        january4.getDay() || 7;


    const firstMonday =
        new Date(
            january4
        );


    firstMonday.setDate(
        january4.getDate() -
        january4Day +
        1
    );


    const result =
        new Date(
            firstMonday
        );


    result.setDate(
        firstMonday.getDate() +
        (
            week -
            1
        ) *
        7
    );


    return result;
}


// ============================================================
// GET MONTH OF ISO WEEK
// ============================================================

function getMonthForWeek(
    year,
    week
) {

    const monday =
        getISOWeekMonday(
            year,
            week
        );


    return (
        monday.getMonth() +
        1
    );
}


// ============================================================
// GET WEEKS BELONGING TO MONTH
//
// We use Monday of each ISO week.
// ============================================================

function getWeeksForMonth(
    year,
    month
) {

    const result = [];


    const maxWeek =
        Math.min(
            getISOWeeksInYear(
                year
            ),
            52
        );


    for (
        let week = 1;
        week <= maxWeek;
        week++
    ) {

        const monday =
            getISOWeekMonday(
                year,
                week
            );


        if (
            monday.getFullYear() === year &&
            monday.getMonth() + 1 === month
        ) {

            result.push(
                week
            );
        }
    }


    return result;
}


// ============================================================
// FIXED DISTRIBUTION SLOT
//
// IMPORTANT:
//
// Every task gets one permanent slot:
//
// Slot 0 = first PM week of month
// Slot 1 = second PM week
// Slot 2 = third PM week
// Slot 3 = fourth PM week
//
// We always use 4 slots.
//
// This prevents:
//
// January = 4th week
// February = 1st week
//
// which could create consecutive W9 / W10 type schedules.
// ============================================================

function getTaskFixedSlot(task) {

    const taskId =
        Number(
            task.id
        ) || 1;


    const machineId =
        Number(
            task.machine_id
        ) || 1;


    // Distribute tasks between 4 slots.
    // Same task always receives same slot.

    const distributionNumber =
        (
            taskId * 31
        ) +
        (
            machineId * 17
        );


    return (
        distributionNumber %
        4
    );
}


// ============================================================
// ASSIGN EXACTLY ONE WEEK INSIDE MONTH
// ============================================================

function getAssignedWeek(
    task,
    year,
    month
) {

    const weeks =
        getWeeksForMonth(
            year,
            month
        );


    if (
        weeks.length === 0
    ) {

        return null;
    }


    const fixedSlot =
        getTaskFixedSlot(
            task
        );


    // Protect against unusual month arrays.

    const safeIndex =
        Math.min(
            fixedSlot,
            weeks.length - 1
        );


    return weeks[
        safeIndex
    ];
}


// ============================================================
// IS MONTH DUE
//
// Example:
//
// 1 month:
// Jan Feb Mar Apr...
//
// 2 months:
// Jan Mar May Jul...
//
// 3 months:
// Jan Apr Jul Oct
//
// 6 months:
// Jan Jul
//
// 12 months:
// Jan
// ============================================================

function isMonthIntervalDue(
    month,
    intervalMonths
) {

    const interval =
        Number(
            intervalMonths
        );


    if (
        !interval ||
        interval < 1
    ) {

        return false;
    }


    return (
        (
            month -
            1
        ) %
        interval ===
        0
    );
}


// ============================================================
// MULTI-YEAR FREQUENCY
// ============================================================

function isMultiYearDue(
    year,
    intervalMonths
) {

    const months =
        Number(
            intervalMonths
        );


    if (
        months <= 12
    ) {

        return true;
    }


    const intervalYears =
        Math.round(
            months /
            12
        );


    if (
        intervalYears < 1
    ) {

        return false;
    }


    return (
        (
            year -
            BASE_YEAR
        ) %
        intervalYears ===
        0
    );
}


// ============================================================
// NORMALIZE FREQUENCY TYPE
// ============================================================

function normalizeFrequencyType(task) {

    return String(
        task.frequency_type ||
        ""
    )
        .trim()
        .toLowerCase();
}


// ============================================================
// SHOULD PM RUN THIS WEEK?
// ============================================================

function shouldScheduleTask(
    task,
    year,
    week
) {

    const type =
        normalizeFrequencyType(
            task
        );


    const value =
        Number(
            task.frequency_value
        );


    const month =
        getMonthForWeek(
            year,
            week
        );


    // ========================================================
    // WEEKLY
    // ========================================================

    if (
        type === "weekly"
    ) {

        return true;
    }


    // ========================================================
    // HOURS
    // ========================================================

    if (
        type === "hours"
    ) {

        // ----------------------------------------------------
        // BELOW 500 HOURS
        // WEEKLY
        // ----------------------------------------------------

        if (
            value < 500
        ) {

            return true;
        }


        // ----------------------------------------------------
        // 500 - 2000 HOURS
        // MONTHLY
        // ONE WEEK ONLY PER MONTH
        // ----------------------------------------------------

        if (
            value >= 500 &&
            value <= 2000
        ) {

            return (
                week ===
                getAssignedWeek(
                    task,
                    year,
                    month
                )
            );
        }


        // ----------------------------------------------------
        // >2000 - 3000 HOURS
        // QUARTERLY
        // JAN / APR / JUL / OCT
        // ----------------------------------------------------

        if (
            value > 2000 &&
            value <= 3000
        ) {

            if (
                !isMonthIntervalDue(
                    month,
                    3
                )
            ) {

                return false;
            }


            return (
                week ===
                getAssignedWeek(
                    task,
                    year,
                    month
                )
            );
        }


        // ----------------------------------------------------
        // >3000 - 6000 HOURS
        // SEMIANNUAL
        // JAN / JUL
        // ----------------------------------------------------

        if (
            value > 3000 &&
            value <= 6000
        ) {

            if (
                !isMonthIntervalDue(
                    month,
                    6
                )
            ) {

                return false;
            }


            return (
                week ===
                getAssignedWeek(
                    task,
                    year,
                    month
                )
            );
        }


        // ----------------------------------------------------
        // >6000 HOURS
        // ANNUAL
        // JANUARY
        // ----------------------------------------------------

        if (
            value > 6000
        ) {

            if (
                month !== 1
            ) {

                return false;
            }


            return (
                week ===
                getAssignedWeek(
                    task,
                    year,
                    1
                )
            );
        }


        return false;
    }


    // ========================================================
    // MONTH-BASED FREQUENCY
    //
    // frequency_value examples:
    //
    // 1  = Monthly
    // 2  = Every 2 months
    // 3  = Quarterly
    // 6  = Semiannual
    // 12 = Annual
    // 24 = Every 2 years
    // 36 = Every 3 years
    // ========================================================

    if (
        type === "monthly"
    ) {

        if (
            !value ||
            value < 1
        ) {

            return false;
        }


        // ----------------------------------------------------
        // MULTI YEAR
        // ----------------------------------------------------

        if (
            value > 12
        ) {

            if (
                !isMultiYearDue(
                    year,
                    value
                )
            ) {

                return false;
            }


            if (
                month !== 1
            ) {

                return false;
            }


            return (
                week ===
                getAssignedWeek(
                    task,
                    year,
                    1
                )
            );
        }


        // ----------------------------------------------------
        // MONTHLY / QUARTERLY / SEMIANNUAL / ANNUAL
        // ----------------------------------------------------

        if (
            !isMonthIntervalDue(
                month,
                value
            )
        ) {

            return false;
        }


        return (
            week ===
            getAssignedWeek(
                task,
                year,
                month
            )
        );
    }


    return false;
}


// ============================================================
// LOAD ACTIVE PM TASKS
// ============================================================

async function getActiveTasks(
    client
) {

    const result =
        await client.query(`
            SELECT
                t.id,
                t.machine_id,
                t.part_name,
                t.maintenance_task,
                t.frequency_text,
                t.frequency_type,
                t.frequency_value,

                m.machine_code,
                m.machine_name,

                l.line_name

            FROM pm_tasks t

            JOIN machines m
                ON
                    m.id =
                    t.machine_id

            JOIN production_lines l
                ON
                    l.id =
                    m.production_line_id

            WHERE
                t.active = TRUE

            ORDER BY
                l.id,
                m.id,
                t.id
        `);


    return result.rows;
}


// ============================================================
// GENERATE ONE WEEK
// ============================================================

async function generateWeekSchedule(
    client,
    year,
    week,
    tasks = null
) {

    const masterTasks =
        tasks ||
        await getActiveTasks(
            client
        );


    let created =
        0;


    let alreadyExists =
        0;


    let notDue =
        0;


    for (
        const task
        of masterTasks
    ) {

        const due =
            shouldScheduleTask(
                task,
                year,
                week
            );


        if (
            !due
        ) {

            notDue++;

            continue;
        }


        const existing =
            await client.query(
                `
                SELECT
                    id

                FROM pm_schedule

                WHERE
                    pm_task_id = $1

                AND
                    planned_year = $2

                AND
                    planned_week = $3

                LIMIT 1
                `,
                [
                    task.id,
                    year,
                    week
                ]
            );


        if (
            existing.rows.length >
            0
        ) {

            alreadyExists++;

            continue;
        }


        await client.query(
            `
            INSERT INTO pm_schedule
            (
                pm_task_id,
                planned_year,
                planned_week,
                status
            )

            VALUES
            (
                $1,
                $2,
                $3,
                'Pending'
            )
            `,
            [
                task.id,
                year,
                week
            ]
        );


        created++;
    }


    return {

        created,

        alreadyExists,

        notDue,

        totalMaster:
            masterTasks.length
    };
}


// ============================================================
// GENERATE FULL YEAR
// ============================================================

async function generateFullYear(
    client,
    year
) {

    const tasks =
        await getActiveTasks(
            client
        );


    const maxWeek =
        Math.min(
            getISOWeeksInYear(
                year
            ),
            52
        );


    let created =
        0;


    let alreadyExists =
        0;


    let notDue =
        0;


    const weeks =
        [];


    for (
        let week = 1;
        week <= maxWeek;
        week++
    ) {

        const result =
            await generateWeekSchedule(
                client,
                year,
                week,
                tasks
            );


        created +=
            result.created;


        alreadyExists +=
            result.alreadyExists;


        notDue +=
            result.notDue;


        weeks.push({

            week,

            created:
                result.created,

            already_exists:
                result.alreadyExists,

            not_due:
                result.notDue
        });
    }


    return {

        year,

        total_pm_master:
            tasks.length,

        created,

        already_exists:
            alreadyExists,

        not_due:
            notDue,

        weeks
    };
}


// ============================================================
// SAFE REBUILD YEAR
//
// ONLY PENDING RECORDS ARE DELETED.
//
// COMPLETED = PRESERVED
// DEFERRED  = PRESERVED
// ============================================================

async function rebuildYear(
    client,
    year
) {

    const pendingBefore =
        await client.query(
            `
            SELECT
                COUNT(*)::int AS count

            FROM pm_schedule

            WHERE
                planned_year = $1

            AND
                status = 'Pending'
            `,
            [
                year
            ]
        );


    const completedBefore =
        await client.query(
            `
            SELECT
                COUNT(*)::int AS count

            FROM pm_schedule

            WHERE
                planned_year = $1

            AND
                status = 'Completed'
            `,
            [
                year
            ]
        );


    const deferredBefore =
        await client.query(
            `
            SELECT
                COUNT(*)::int AS count

            FROM pm_schedule

            WHERE
                planned_year = $1

            AND
                status = 'Deferred'
            `,
            [
                year
            ]
        );


    // ========================================================
    // DELETE ONLY PENDING
    // ========================================================

    const deleted =
        await client.query(
            `
            DELETE FROM pm_schedule

            WHERE
                planned_year = $1

            AND
                status = 'Pending'
            `,
            [
                year
            ]
        );


    // ========================================================
    // REBUILD CORRECT PLAN
    // ========================================================

    const generation =
        await generateFullYear(
            client,
            year
        );


    return {

        year,

        pending_before:
            pendingBefore.rows[
                0
            ].count,

        pending_deleted:
            deleted.rowCount,

        completed_preserved:
            completedBefore.rows[
                0
            ].count,

        deferred_preserved:
            deferredBefore.rows[
                0
            ].count,

        ...generation
    };
}


// ============================================================
// CHECK DISTRIBUTION
// ============================================================

async function checkDistribution(
    client,
    year
) {

    const result =
        await client.query(
            `
            SELECT

                t.id AS pm_task_id,

                l.line_name,

                m.machine_code,

                m.machine_name,

                t.part_name,

                t.maintenance_task,

                t.frequency_text,

                t.frequency_type,

                t.frequency_value,

                COUNT(
                    s.id
                )::int AS planned_count,

                ARRAY_AGG(
                    s.planned_week
                    ORDER BY
                        s.planned_week
                )
                FILTER
                (
                    WHERE
                        s.id IS NOT NULL
                ) AS planned_weeks

            FROM pm_tasks t

            JOIN machines m
                ON
                    m.id =
                    t.machine_id

            JOIN production_lines l
                ON
                    l.id =
                    m.production_line_id

            LEFT JOIN pm_schedule s
                ON
                    s.pm_task_id =
                    t.id

                AND
                    s.planned_year =
                    $1

            WHERE
                t.active = TRUE

            GROUP BY
                t.id,
                l.line_name,
                m.machine_code,
                m.machine_name,
                t.part_name,
                t.maintenance_task,
                t.frequency_text,
                t.frequency_type,
                t.frequency_value

            ORDER BY
                l.line_name,
                m.machine_code,
                t.id
            `,
            [
                year
            ]
        );


    return result.rows;
}


// ============================================================
// VALIDATE DISTRIBUTION
// ============================================================

function validateDistribution(
    rows,
    year
) {

    const problems =
        [];


    for (
        const row
        of rows
    ) {

        const weeks =
            row.planned_weeks ||
            [];


        const type =
            String(
                row.frequency_type ||
                ""
            )
                .trim()
                .toLowerCase();


        const value =
            Number(
                row.frequency_value
            );


        // ====================================================
        // MONTHLY RULE
        // ====================================================

        const monthlyRule =
            (
                type === "monthly" &&
                value === 1
            ) ||
            (
                type === "hours" &&
                value >= 500 &&
                value <= 2000
            );


        if (
            monthlyRule
        ) {

            const monthCounts =
                {};


            for (
                const week
                of weeks
            ) {

                const month =
                    getMonthForWeek(
                        year,
                        Number(
                            week
                        )
                    );


                monthCounts[
                    month
                ] =
                    (
                        monthCounts[
                            month
                        ] ||
                        0
                    ) +
                    1;
            }


            for (
                const [
                    month,
                    count
                ]
                of Object.entries(
                    monthCounts
                )
            ) {

                if (
                    count >
                    1
                ) {

                    problems.push({

                        pm_task_id:
                            row.pm_task_id,

                        machine_code:
                            row.machine_code,

                        machine_name:
                            row.machine_name,

                        part_name:
                            row.part_name,

                        maintenance_task:
                            row.maintenance_task,

                        frequency_text:
                            row.frequency_text,

                        month:
                            Number(
                                month
                            ),

                        count,

                        problem:
                            `Monthly PM appears ${count} times in same month`
                    });
                }
            }


            // =================================================
            // ALSO CHECK SLOT CONSISTENCY
            //
            // Monthly tasks should keep approximately the same
            // relative week position each month.
            // =================================================

            let previousWeek =
                null;


            for (
                const week
                of weeks
            ) {

                if (
                    previousWeek !==
                    null
                ) {

                    const gap =
                        Number(
                            week
                        ) -
                        Number(
                            previousWeek
                        );


                    // Monthly PM normally should be around
                    // 4 or 5 weeks apart.
                    //
                    // Gap <= 2 indicates suspicious consecutive
                    // scheduling.

                    if (
                        gap <= 2
                    ) {

                        problems.push({

                            pm_task_id:
                                row.pm_task_id,

                            machine_code:
                                row.machine_code,

                            machine_name:
                                row.machine_name,

                            part_name:
                                row.part_name,

                            maintenance_task:
                                row.maintenance_task,

                            frequency_text:
                                row.frequency_text,

                            previous_week:
                                previousWeek,

                            current_week:
                                week,

                            problem:
                                `Monthly PM weeks are too close: W${previousWeek} and W${week}`
                        });
                    }
                }


                previousWeek =
                    week;
            }
        }
    }


    return problems;
}


// ============================================================
// COMMAND LINE
//
// NORMAL:
// node src/generateSchedule.js 2026
//
// SAFE REBUILD:
// node src/generateSchedule.js 2026 --rebuild
//
// CHECK:
// node src/generateSchedule.js 2026 --check
// ============================================================

async function runFromCommandLine() {

    const year =
        Number(
            process.argv[
                2
            ] ||
            new Date()
                .getFullYear()
        );


    const command =
        String(
            process.argv[
                3
            ] ||
            ""
        )
            .trim()
            .toLowerCase();


    if (
        !year ||
        year < 2000 ||
        year > 2100
    ) {

        console.error(
            "Invalid year."
        );


        process.exit(
            1
        );
    }


    const client =
        await pool.connect();


    try {

        // ====================================================
        // SAFE REBUILD
        // ====================================================

        if (
            command ===
            "--rebuild"
        ) {

            console.log("");
            console.log(
                "=============================================="
            );

            console.log(
                ` SAFE PM SCHEDULE REBUILD - ${year}`
            );

            console.log(
                "=============================================="
            );

            console.log("");

            console.log(
                "Completed records will be preserved."
            );

            console.log(
                "Deferred records will be preserved."
            );

            console.log(
                "Only Pending records will be rebuilt."
            );

            console.log("");


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


            console.log(
                `Pending before rebuild: ${result.pending_before}`
            );


            console.log(
                `Pending deleted: ${result.pending_deleted}`
            );


            console.log(
                `Completed preserved: ${result.completed_preserved}`
            );


            console.log(
                `Deferred preserved: ${result.deferred_preserved}`
            );


            console.log(
                `Active PM master tasks: ${result.total_pm_master}`
            );


            console.log(
                `New Pending schedule created: ${result.created}`
            );


            console.log(
                `Existing historical records found: ${result.already_exists}`
            );


            console.log("");
            console.log(
                "Checking distribution..."
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


            console.log("");


            if (
                problems.length ===
                0
            ) {

                console.log(
                    "✓ Distribution validation passed."
                );


                console.log(
                    "✓ Monthly PM appears only once per month."
                );


                console.log(
                    "✓ Monthly PM week position is consistent."
                );

            } else {

                console.log(
                    `WARNING: ${problems.length} distribution problem(s) found.`
                );


                console.table(
                    problems.slice(
                        0,
                        100
                    )
                );
            }


            console.log("");

            console.log(
                "=============================================="
            );

            console.log(
                " REBUILD COMPLETED"
            );

            console.log(
                "=============================================="
            );


            return;
        }


        // ====================================================
        // CHECK ONLY
        // ====================================================

        if (
            command ===
            "--check"
        ) {

            console.log(
                `Checking PM distribution for ${year}...`
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


            console.log("");


            if (
                problems.length ===
                0
            ) {

                console.log(
                    "✓ Distribution is valid."
                );


                console.log(
                    "✓ Monthly PM occurs only once per month."
                );


                console.log(
                    "✓ Monthly PM weeks are consistently distributed."
                );

            } else {

                console.log(
                    `Found ${problems.length} problem(s):`
                );


                console.table(
                    problems.slice(
                        0,
                        100
                    )
                );
            }


            return;
        }


        // ====================================================
        // NORMAL GENERATION
        // ====================================================

        await client.query(
            "BEGIN"
        );


        const result =
            await generateFullYear(
                client,
                year
            );


        await client.query(
            "COMMIT"
        );


        console.log("");

        console.log(
            `PM schedule generated for ${year}`
        );


        console.log(
            `Active PM master tasks: ${result.total_pm_master}`
        );


        console.log(
            `New records created: ${result.created}`
        );


        console.log(
            `Already existing: ${result.already_exists}`
        );


        console.log("");


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

            // Ignore
        }


        console.error("");

        console.error(
            "Schedule generation failed:"
        );


        console.error(
            error
        );


        process.exitCode =
            1;


    } finally {

        client.release();


        await pool.end();
    }
}


// ============================================================
// EXPORT FOR SERVER.JS
// ============================================================

module.exports = {

    getISOWeekNumber,

    getISOWeeksInYear,

    getISOWeekMonday,

    getMonthForWeek,

    getWeeksForMonth,

    getAssignedWeek,

    shouldScheduleTask,

    generateWeekSchedule,

    generateFullYear,

    rebuildYear,

    checkDistribution,

    validateDistribution
};


// ============================================================
// RUN DIRECTLY
// ============================================================

if (
    require.main ===
    module
) {

    runFromCommandLine();
}