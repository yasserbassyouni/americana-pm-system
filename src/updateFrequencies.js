const pool = require("./database");

function parseFrequency(text) {

    let value = String(text || "")
        .trim()
        .toLowerCase()
        .replace(/,/g, "")
        .replace(/\s+/g, " ");

    if (!value) {
        return {
            frequencyType: "other",
            frequencyValue: null
        };
    }

    // WEEKLY
    if (
        value === "weekly" ||
        value === "week" ||
        value === "weeekly"
    ) {
        return {
            frequencyType: "weekly",
            frequencyValue: 1
        };
    }

    // MONTHLY
    if (
        value === "monthly" ||
        value === "month" ||
        value === "montly"
    ) {
        return {
            frequencyType: "monthly",
            frequencyValue: 1
        };
    }

    // QUARTERLY = EVERY 3 MONTHS
    if (
        value === "quarterly" ||
        value === "quarter"
    ) {
        return {
            frequencyType: "monthly",
            frequencyValue: 3
        };
    }

    // ANNUALLY = EVERY 12 MONTHS
    if (
        value === "annually" ||
        value === "annual" ||
        value === "anually" ||
        value === "annualy" ||
        value === "yearly"
    ) {
        return {
            frequencyType: "monthly",
            frequencyValue: 12
        };
    }

    // NUMBER OF MONTHS
    // Examples:
    // 2month
    // 2 month
    // 3 months
    // 6 months
    const monthMatch = value.match(/^(\d+)\s*months?$/);

    if (monthMatch) {
        return {
            frequencyType: "monthly",
            frequencyValue: Number(monthMatch[1])
        };
    }

    // NUMBER OF YEARS
    // Examples:
    // 2 years
    // 3 years
    // 5 years
    const yearMatch = value.match(/^(\d+)\s*years?$/);

    if (yearMatch) {
        return {
            frequencyType: "monthly",
            frequencyValue: Number(yearMatch[1]) * 12
        };
    }

    // OPERATING HOURS
    // Examples:
    // 40 hr
    // 120 hr
    // 500 hr
    // 1000 hr
    // 2000 hr
    // 6000 hr
    // 10,000 hr
    const hourMatch = value.match(
        /^(\d+)\s*(hr|hrs|hour|hours)$/
    );

    if (hourMatch) {
        return {
            frequencyType: "hours",
            frequencyValue: Number(hourMatch[1])
        };
    }

    // Anything we still cannot recognize
    return {
        frequencyType: "other",
        frequencyValue: null
    };
}


async function updateFrequencies() {

    const client = await pool.connect();

    try {

        console.log("Updating PM frequencies...");

        const result = await client.query(`
            SELECT
                id,
                frequency_text
            FROM pm_tasks
            ORDER BY id
        `);

        let updated = 0;
        let unrecognized = 0;

        for (const row of result.rows) {

            const parsed = parseFrequency(
                row.frequency_text
            );

            await client.query(
                `
                UPDATE pm_tasks
                SET
                    frequency_type = $1,
                    frequency_value = $2,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = $3
                `,
                [
                    parsed.frequencyType,
                    parsed.frequencyValue,
                    row.id
                ]
            );

            updated++;

            if (parsed.frequencyType === "other") {

                unrecognized++;

                console.log(
                    `Unrecognized frequency: "${row.frequency_text}"`
                );
            }
        }

        console.log("");
        console.log("Frequency update completed.");
        console.log(`PM tasks updated: ${updated}`);
        console.log(
            `Unrecognized frequencies: ${unrecognized}`
        );

    } catch (error) {

        console.error("Frequency update failed:");
        console.error(error);

    } finally {

        client.release();
        await pool.end();
    }
}

updateFrequencies();