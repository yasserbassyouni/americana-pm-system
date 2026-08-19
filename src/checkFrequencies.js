require("dotenv").config();

const pool = require("./database");

async function checkFrequencies() {

    try {

        console.log("");
        console.log("Checking PM frequency distribution...");
        console.log("");

        const result = await pool.query(`
            SELECT
                frequency_text,
                frequency_type,
                frequency_value,
                COUNT(*)::int AS task_count
            FROM pm_tasks
            WHERE active = TRUE
            GROUP BY
                frequency_text,
                frequency_type,
                frequency_value
            ORDER BY
                frequency_type,
                frequency_value,
                frequency_text
        `);

        console.table(result.rows);

        console.log("");
        console.log(
            "Total active PM tasks:",
            result.rows.reduce(
                (sum, row) =>
                    sum + Number(row.task_count),
                0
            )
        );

    } catch (error) {

        console.error(
            "Frequency check failed:"
        );

        console.error(
            error
        );

    } finally {

        await pool.end();
    }
}

checkFrequencies();