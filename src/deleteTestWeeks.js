require("dotenv").config();

const pool = require("./database");

async function deleteTestWeeks() {

    const year = 2026;
    const weeks = [36, 37, 38, 39, 40];

    try {

        console.log("");
        console.log("======================================");
        console.log(" CLEANING TEST PM WEEKS");
        console.log("======================================");
        console.log("");

        for (const week of weeks) {

            const result = await pool.query(
                `
                DELETE FROM pm_schedule
                WHERE planned_year = $1
                  AND planned_week = $2
                RETURNING id
                `,
                [year, week]
            );

            console.log(
                `Week ${week} deleted: ${result.rowCount} records`
            );
        }

        console.log("");
        console.log("======================================");
        console.log(" Test weeks deleted successfully.");
        console.log("======================================");
        console.log("");
        console.log("Week 34 and Week 35 were NOT changed.");
        console.log("");

    } catch (error) {

        console.error("");
        console.error("Delete failed:");
        console.error(error);

    } finally {

        await pool.end();
    }
}

deleteTestWeeks();