// ============================================================
// DATABASE CONNECTION
//
// LOCAL PC:
// Uses DB_HOST / DB_PORT / DB_NAME / DB_USER / DB_PASSWORD
//
// RENDER CLOUD:
// Uses DATABASE_URL
// ============================================================

require(
    "dotenv"
).config();


const {
    Pool
} =
    require(
        "pg"
    );


// ============================================================
// DETECT CLOUD DATABASE URL
// ============================================================

const isCloudDatabase =
    Boolean(
        process.env.DATABASE_URL
    );


// ============================================================
// DATABASE CONFIGURATION
// ============================================================

let poolConfig;


// ============================================================
// RENDER / CLOUD
// ============================================================

if (
    isCloudDatabase
) {

    poolConfig = {

        connectionString:
            process.env.DATABASE_URL,

        // Render PostgreSQL connections may require SSL
        // depending on how the connection URL is being used.

        ssl: {

            rejectUnauthorized:
                false
        },

        max:
            10,

        idleTimeoutMillis:
            30000,

        connectionTimeoutMillis:
            10000
    };


    console.log(
        "Database mode: CLOUD"
    );

}


// ============================================================
// LOCAL POSTGRESQL
// ============================================================

else {

    poolConfig = {

        host:
            process.env.DB_HOST ||
            "localhost",

        port:
            Number(
                process.env.DB_PORT ||
                5432
            ),

        database:
            process.env.DB_NAME,

        user:
            process.env.DB_USER,

        password:
            process.env.DB_PASSWORD,

        max:
            10,

        idleTimeoutMillis:
            30000,

        connectionTimeoutMillis:
            10000
    };


    console.log(
        "Database mode: LOCAL"
    );

}


// ============================================================
// CREATE POSTGRESQL POOL
// ============================================================

const pool =
    new Pool(
        poolConfig
    );


// ============================================================
// CONNECTION ERROR
// ============================================================

pool.on(
    "error",
    error => {

        console.error(
            "Unexpected PostgreSQL pool error:",
            error
        );

    }
);


// ============================================================
// TEST DATABASE CONNECTION
// ============================================================

async function testDatabaseConnection() {

    let client;


    try {

        client =
            await pool.connect();


        const result =
            await client.query(
                `
                SELECT
                    NOW() AS database_time,
                    current_database() AS database_name,
                    current_user AS database_user
                `
            );


        console.log(
            ""
        );


        console.log(
            "=============================================="
        );


        console.log(
            " PostgreSQL Connected"
        );


        console.log(
            "=============================================="
        );


        console.log(
            `Mode: ${
                isCloudDatabase
                    ?
                    "CLOUD"
                    :
                    "LOCAL"
            }`
        );


        console.log(
            `Database: ${result.rows[0].database_name}`
        );


        console.log(
            `User: ${result.rows[0].database_user}`
        );


        console.log(
            `Database Time: ${result.rows[0].database_time}`
        );


        console.log(
            "=============================================="
        );


        console.log(
            ""
        );


        return true;


    } catch (
        error
    ) {

        console.error(
            ""
        );


        console.error(
            "=============================================="
        );


        console.error(
            " PostgreSQL Connection Failed"
        );


        console.error(
            "=============================================="
        );


        console.error(
            error.message
        );


        console.error(
            "=============================================="
        );


        console.error(
            ""
        );


        return false;


    } finally {

        if (
            client
        ) {

            client.release();
        }
    }
}


// ============================================================
// EXPORTS
// ============================================================

module.exports =
    pool;


module.exports.testDatabaseConnection =
    testDatabaseConnection;


module.exports.isCloudDatabase =
    isCloudDatabase;