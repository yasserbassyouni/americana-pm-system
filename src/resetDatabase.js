const fs = require("fs");
const path = require("path");
const pool = require("./database");

async function resetDatabase() {
    const client = await pool.connect();

    try {
        console.log("Resetting PM database...");

        await client.query("BEGIN");

        // Remove the old PM tables
        await client.query(`
            DROP TABLE IF EXISTS pm_schedule CASCADE;
            DROP TABLE IF EXISTS pm_tasks CASCADE;
            DROP TABLE IF EXISTS machines CASCADE;
            DROP TABLE IF EXISTS production_lines CASCADE;
        `);

        console.log("Old tables removed.");

        // Read the new schema
        const schemaPath = path.join(
            __dirname,
            "../database/schema.sql"
        );

        const schema = fs.readFileSync(schemaPath, "utf8");

        // Create the new tables
        await client.query(schema);

        await client.query("COMMIT");

        console.log("Database reset successfully.");
        console.log("New PM tables created successfully.");

    } catch (error) {
        await client.query("ROLLBACK");

        console.error("Database reset failed:");
        console.error(error);

    } finally {
        client.release();
        await pool.end();
    }
}

resetDatabase();