const XLSX = require("xlsx");
const path = require("path");
const pool = require("./database");

async function importExcel() {
    const client = await pool.connect();

    try {
        console.log("Starting Excel import...");

        const excelPath = path.join(
            __dirname,
            "../Cookies PM 2026.xlsx"
        );

        const workbook = XLSX.readFile(excelPath);

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const rows = XLSX.utils.sheet_to_json(worksheet, {
            defval: ""
        });

        console.log(`Rows found: ${rows.length}`);

        await client.query("BEGIN");

        let importedTasks = 0;

        for (const row of rows) {
            const line = String(row["Line"] || "").trim();
            const section = String(row["Section"] || "").trim();
            const machineNo = String(row["Machine no"] || "").trim();
            const machineName = String(row["Machine name"] || "").trim();
            const machineCode = String(row["Machine code"] || "").trim();
            const machineType = String(row["Type"] || "").trim();
            const partName = String(row["Parts"] || "").trim();
            const maintenanceTask = String(row["Maintenance"] || "").trim();
            const frequency = String(row["Freq."] || "").trim();

            // Skip incomplete rows
            if (!line || !machineName || !machineCode || !maintenanceTask) {
                continue;
            }

            // 1. Production Line
            const lineResult = await client.query(
                `
                INSERT INTO production_lines (line_name)
                VALUES ($1)
                ON CONFLICT (line_name)
                DO UPDATE SET line_name = EXCLUDED.line_name
                RETURNING id
                `,
                [line]
            );

            const productionLineId = lineResult.rows[0].id;

            // 2. Machine
            const machineResult = await client.query(
                `
                INSERT INTO machines (
                    production_line_id,
                    section,
                    machine_no,
                    machine_name,
                    machine_code,
                    machine_type
                )
                VALUES ($1, $2, $3, $4, $5, $6)

                ON CONFLICT (machine_code)
                DO UPDATE SET
                    production_line_id = EXCLUDED.production_line_id,
                    section = EXCLUDED.section,
                    machine_no = EXCLUDED.machine_no,
                    machine_name = EXCLUDED.machine_name,
                    machine_type = EXCLUDED.machine_type,
                    updated_at = CURRENT_TIMESTAMP

                RETURNING id
                `,
                [
                    productionLineId,
                    section,
                    machineNo,
                    machineName,
                    machineCode,
                    machineType
                ]
            );

            const machineId = machineResult.rows[0].id;

            // 3. Avoid duplicate PM task
            const existingTask = await client.query(
                `
                SELECT id
                FROM pm_tasks
                WHERE machine_id = $1
                  AND part_name = $2
                  AND maintenance_task = $3
                  AND frequency_text = $4
                `,
                [
                    machineId,
                    partName,
                    maintenanceTask,
                    frequency
                ]
            );

            if (existingTask.rows.length === 0) {
                await client.query(
                    `
                    INSERT INTO pm_tasks (
                        machine_id,
                        part_name,
                        maintenance_task,
                        frequency_text
                    )
                    VALUES ($1, $2, $3, $4)
                    `,
                    [
                        machineId,
                        partName,
                        maintenanceTask,
                        frequency
                    ]
                );

                importedTasks++;
            }
        }

        await client.query("COMMIT");

        console.log("Excel import completed successfully.");
        console.log(`New PM tasks imported: ${importedTasks}`);

    } catch (error) {
        await client.query("ROLLBACK");

        console.error("Excel import failed:");
        console.error(error);

    } finally {
        client.release();
        await pool.end();
    }
}

importExcel();