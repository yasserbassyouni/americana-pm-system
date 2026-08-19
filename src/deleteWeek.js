require("dotenv").config();

const pool = require("./database");

async function deleteWeek() {

    // ==========================================
    // WEEK TO DELETE
    // ==========================================

    const year = 2026;
    const week = 35;


    try {

        console.log("");
        console.log(
            `Deleting PM schedule for Week ${week}, ${year}...`
        );
        console.log("");


        // ==========================================
        // DELETE ONLY THE SELECTED WEEK
        // ==========================================

        const result = await pool.query(
            `
            DELETE FROM pm_schedule

            WHERE planned_year = $1
              AND planned_week = $2

            RETURNING id
            `,
            [
                year,
                week
            ]
        );


        // ==========================================
        // RESULT
        // ==========================================

        console.log(
            "Week deleted successfully."
        );

        console.log(
            `Records deleted: ${result.rowCount}`
        );

        console.log("");

        console.log(
            `Only Week ${week} / ${year} was deleted.`
        );

        console.log(
            "Other weeks were NOT changed."
        );


    } catch (error) {

        console.error("");
        console.error(
            "Delete failed:"
        );

        console.error(
            error
        );


    } finally {

        await pool.end();

    }

}


deleteWeek();