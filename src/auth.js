require(
    "dotenv"
).config();


const express =
    require(
        "express"
    );


const session =
    require(
        "express-session"
    );


const bcrypt =
    require(
        "bcryptjs"
    );


const pool =
    require(
        "./database"
    );


const router =
    express.Router();


// ============================================================
// SESSION CONFIGURATION
// ============================================================

const sessionMiddleware =
    session(
        {

            name:
                "pm_session",

            secret:
                process.env.SESSION_SECRET ||
                "change-this-pm-session-secret",

            resave:
                false,

            saveUninitialized:
                false,

            cookie: {

                httpOnly:
                    true,

                secure:
                    false,

                sameSite:
                    "lax",

                maxAge:
                    1000 *
                    60 *
                    60 *
                    12

            }

        }
    );


// ============================================================
// CREATE USERS TABLE
// ============================================================

async function ensureUsersTable() {

    await pool.query(
        `
        CREATE TABLE IF NOT EXISTS app_users
        (
            id SERIAL PRIMARY KEY,

            username VARCHAR(80)
                NOT NULL
                UNIQUE,

            full_name VARCHAR(150)
                NOT NULL,

            password_hash TEXT
                NOT NULL,

            role VARCHAR(30)
                NOT NULL
                DEFAULT 'technician',

            active BOOLEAN
                NOT NULL
                DEFAULT TRUE,

            last_login_at TIMESTAMPTZ,

            password_changed_at TIMESTAMPTZ,

            created_at TIMESTAMPTZ
                NOT NULL
                DEFAULT CURRENT_TIMESTAMP,

            updated_at TIMESTAMPTZ
                NOT NULL
                DEFAULT CURRENT_TIMESTAMP,

            CONSTRAINT app_users_role_check
            CHECK
            (
                role IN
                (
                    'admin',
                    'supervisor',
                    'technician',
                    'viewer'
                )
            )
        )
        `
    );


    await pool.query(
        `
        CREATE UNIQUE INDEX IF NOT EXISTS
            idx_app_users_username_lower

        ON app_users
        (
            LOWER(username)
        )
        `
    );

}


// ============================================================
// DEFAULT ADMIN
// ============================================================

async function ensureDefaultAdmin() {

    await ensureUsersTable();


    const username =
        String(
            process.env.ADMIN_USERNAME ||
            "admin"
        )
            .trim()
            .toLowerCase();


    const password =
        String(
            process.env.ADMIN_PASSWORD ||
            "Admin@123"
        );


    const fullName =
        String(
            process.env.ADMIN_FULL_NAME ||
            "PM Administrator"
        )
            .trim();


    const existing =
        await pool.query(
            `
            SELECT
                id,
                username

            FROM app_users

            WHERE
                LOWER(username) =
                LOWER($1)

            LIMIT 1
            `,
            [
                username
            ]
        );


    if (
        existing.rows.length >
        0
    ) {

        console.log(
            `Admin user already exists: ${username}`
        );

        return;
    }


    const passwordHash =
        await bcrypt.hash(
            password,
            12
        );


    await pool.query(
        `
        INSERT INTO app_users
        (
            username,
            full_name,
            password_hash,
            role,
            active
        )

        VALUES
        (
            $1,
            $2,
            $3,
            'admin',
            TRUE
        )
        `,
        [
            username,
            fullName,
            passwordHash
        ]
    );


    console.log(
        ""
    );


    console.log(
        "=============================================="
    );


    console.log(
        " DEFAULT ADMIN USER CREATED"
    );


    console.log(
        "=============================================="
    );


    console.log(
        `Username: ${username}`
    );


    console.log(
        "Change the default password after first login."
    );


    console.log(
        "=============================================="
    );

}


// ============================================================
// REQUIRE LOGIN - API
// ============================================================

function requireAuth(
    req,
    res,
    next
) {

    if (
        req.session &&
        req.session.user
    ) {

        return next();
    }


    return res
        .status(
            401
        )
        .json(
            {
                error:
                    "Authentication required"
            }
        );

}


// ============================================================
// REQUIRE LOGIN - PAGE
// ============================================================

function requirePageAuth(
    req,
    res,
    next
) {

    if (
        req.session &&
        req.session.user
    ) {

        return next();
    }


    return res.redirect(
        "/login"
    );

}


// ============================================================
// REQUIRE ROLE
// ============================================================

function requireRole(
    ...allowedRoles
) {

    return function (
        req,
        res,
        next
    ) {

        if (
            !req.session ||
            !req.session.user
        ) {

            return res
                .status(
                    401
                )
                .json(
                    {
                        error:
                            "Authentication required"
                    }
                );
        }


        const userRole =
            String(
                req.session.user.role ||
                ""
            )
                .toLowerCase();


        const allowed =
            allowedRoles.map(
                role =>
                    String(
                        role
                    )
                        .toLowerCase()
            );


        if (
            !allowed.includes(
                userRole
            )
        ) {

            return res
                .status(
                    403
                )
                .json(
                    {
                        error:
                            "You do not have permission to perform this action."
                    }
                );
        }


        next();

    };

}


// ============================================================
// LOGIN
// ============================================================

router.post(
    "/login",
    async (
        req,
        res
    ) => {

        try {

            const username =
                String(
                    req.body.username ||
                    ""
                )
                    .trim()
                    .toLowerCase();


            const password =
                String(
                    req.body.password ||
                    ""
                );


            if (
                !username ||
                !password
            ) {

                return res
                    .status(
                        400
                    )
                    .json(
                        {
                            error:
                                "Username and password are required."
                        }
                    );
            }


            const result =
                await pool.query(
                    `
                    SELECT
                        id,
                        username,
                        full_name,
                        password_hash,
                        role,
                        active

                    FROM app_users

                    WHERE
                        LOWER(username) =
                        LOWER($1)

                    LIMIT 1
                    `,
                    [
                        username
                    ]
                );


            if (
                result.rows.length ===
                0
            ) {

                return res
                    .status(
                        401
                    )
                    .json(
                        {
                            error:
                                "Invalid username or password."
                        }
                    );
            }


            const user =
                result.rows[0];


            if (
                user.active !==
                true
            ) {

                return res
                    .status(
                        403
                    )
                    .json(
                        {
                            error:
                                "This user account is inactive."
                        }
                    );
            }


            const validPassword =
                await bcrypt.compare(
                    password,
                    user.password_hash
                );


            if (
                !validPassword
            ) {

                return res
                    .status(
                        401
                    )
                    .json(
                        {
                            error:
                                "Invalid username or password."
                        }
                    );
            }


            req.session.user =
                {

                    id:
                        user.id,

                    username:
                        user.username,

                    fullName:
                        user.full_name,

                    role:
                        user.role

                };


            await pool.query(
                `
                UPDATE app_users

                SET
                    last_login_at =
                        CURRENT_TIMESTAMP,

                    updated_at =
                        CURRENT_TIMESTAMP

                WHERE
                    id = $1
                `,
                [
                    user.id
                ]
            );


            req.session.save(
                error => {

                    if (
                        error
                    ) {

                        console.error(
                            "Session save error:",
                            error
                        );


                        return res
                            .status(
                                500
                            )
                            .json(
                                {
                                    error:
                                        "Unable to create login session."
                                }
                            );
                    }


                    res.json(
                        {

                            success:
                                true,

                            message:
                                "Login successful",

                            user:
                                req.session.user

                        }
                    );

                }
            );

        } catch (
            error
        ) {

            console.error(
                "Login error:",
                error
            );


            res
                .status(
                    500
                )
                .json(
                    {
                        error:
                            error.message
                    }
                );

        }

    }
);


// ============================================================
// CURRENT USER
// ============================================================

router.get(
    "/me",
    (
        req,
        res
    ) => {

        if (
            !req.session ||
            !req.session.user
        ) {

            return res
                .status(
                    401
                )
                .json(
                    {
                        authenticated:
                            false
                    }
                );
        }


        res.json(
            {

                authenticated:
                    true,

                user:
                    req.session.user

            }
        );

    }
);


// ============================================================
// LOGOUT
// ============================================================

router.post(
    "/logout",
    (
        req,
        res
    ) => {

        if (
            !req.session
        ) {

            return res.json(
                {
                    success:
                        true
                }
            );
        }


        req.session.destroy(
            error => {

                if (
                    error
                ) {

                    console.error(
                        "Logout error:",
                        error
                    );


                    return res
                        .status(
                            500
                        )
                        .json(
                            {
                                error:
                                    "Unable to logout."
                            }
                        );
                }


                res.clearCookie(
                    "pm_session"
                );


                res.json(
                    {

                        success:
                            true,

                        message:
                            "Logged out successfully"

                    }
                );

            }
        );

    }
);


// ============================================================
// CHANGE PASSWORD
// ============================================================

router.post(
    "/change-password",

    requireAuth,

    async (
        req,
        res
    ) => {

        try {

            const currentPassword =
                String(
                    req.body.currentPassword ||
                    ""
                );


            const newPassword =
                String(
                    req.body.newPassword ||
                    ""
                );


            if (
                !currentPassword ||
                !newPassword
            ) {

                return res
                    .status(
                        400
                    )
                    .json(
                        {
                            error:
                                "Current password and new password are required."
                        }
                    );
            }


            if (
                newPassword.length <
                8
            ) {

                return res
                    .status(
                        400
                    )
                    .json(
                        {
                            error:
                                "New password must contain at least 8 characters."
                        }
                    );
            }


            const result =
                await pool.query(
                    `
                    SELECT
                        id,
                        password_hash

                    FROM app_users

                    WHERE
                        id = $1

                    LIMIT 1
                    `,
                    [
                        req.session.user.id
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
                    .json(
                        {
                            error:
                                "User not found."
                        }
                    );
            }


            const valid =
                await bcrypt.compare(
                    currentPassword,
                    result.rows[0].password_hash
                );


            if (
                !valid
            ) {

                return res
                    .status(
                        401
                    )
                    .json(
                        {
                            error:
                                "Current password is incorrect."
                        }
                    );
            }


            const newHash =
                await bcrypt.hash(
                    newPassword,
                    12
                );


            await pool.query(
                `
                UPDATE app_users

                SET
                    password_hash = $1,

                    password_changed_at =
                        CURRENT_TIMESTAMP,

                    updated_at =
                        CURRENT_TIMESTAMP

                WHERE
                    id = $2
                `,
                [
                    newHash,
                    req.session.user.id
                ]
            );


            res.json(
                {

                    success:
                        true,

                    message:
                        "Password changed successfully."

                }
            );

        } catch (
            error
        ) {

            console.error(
                "Change password error:",
                error
            );


            res
                .status(
                    500
                )
                .json(
                    {
                        error:
                            error.message
                    }
                );

        }

    }
);


// ============================================================
// LIST USERS - ADMIN ONLY
// ============================================================

router.get(
    "/users",

    requireRole(
        "admin"
    ),

    async (
        req,
        res
    ) => {

        try {

            const result =
                await pool.query(
                    `
                    SELECT
                        id,
                        username,
                        full_name,
                        role,
                        active,
                        last_login_at,
                        created_at

                    FROM app_users

                    ORDER BY
                        full_name,
                        username
                    `
                );


            res.json(
                result.rows
            );

        } catch (
            error
        ) {

            console.error(
                "List users error:",
                error
            );


            res
                .status(
                    500
                )
                .json(
                    {
                        error:
                            error.message
                    }
                );

        }

    }
);


// ============================================================
// CREATE USER - ADMIN ONLY
// ============================================================

router.post(
    "/users",

    requireRole(
        "admin"
    ),

    async (
        req,
        res
    ) => {

        try {

            const username =
                String(
                    req.body.username ||
                    ""
                )
                    .trim()
                    .toLowerCase();


            const fullName =
                String(
                    req.body.fullName ||
                    ""
                )
                    .trim();


            const password =
                String(
                    req.body.password ||
                    ""
                );


            const role =
                String(
                    req.body.role ||
                    "technician"
                )
                    .trim()
                    .toLowerCase();


            const allowedRoles =
                [
                    "admin",
                    "supervisor",
                    "technician",
                    "viewer"
                ];


            if (
                !username ||
                !fullName ||
                !password
            ) {

                return res
                    .status(
                        400
                    )
                    .json(
                        {
                            error:
                                "Username, full name and password are required."
                        }
                    );
            }


            if (
                password.length <
                8
            ) {

                return res
                    .status(
                        400
                    )
                    .json(
                        {
                            error:
                                "Password must contain at least 8 characters."
                        }
                    );
            }


            if (
                !allowedRoles.includes(
                    role
                )
            ) {

                return res
                    .status(
                        400
                    )
                    .json(
                        {
                            error:
                                "Invalid user role."
                        }
                    );
            }


            const existing =
                await pool.query(
                    `
                    SELECT
                        id

                    FROM app_users

                    WHERE
                        LOWER(username) =
                        LOWER($1)

                    LIMIT 1
                    `,
                    [
                        username
                    ]
                );


            if (
                existing.rows.length >
                0
            ) {

                return res
                    .status(
                        409
                    )
                    .json(
                        {
                            error:
                                "Username already exists."
                        }
                    );
            }


            const passwordHash =
                await bcrypt.hash(
                    password,
                    12
                );


            const result =
                await pool.query(
                    `
                    INSERT INTO app_users
                    (
                        username,
                        full_name,
                        password_hash,
                        role,
                        active
                    )

                    VALUES
                    (
                        $1,
                        $2,
                        $3,
                        $4,
                        TRUE
                    )

                    RETURNING
                        id,
                        username,
                        full_name,
                        role,
                        active,
                        created_at
                    `,
                    [
                        username,
                        fullName,
                        passwordHash,
                        role
                    ]
                );


            res
                .status(
                    201
                )
                .json(
                    {

                        success:
                            true,

                        user:
                            result.rows[0]

                    }
                );

        } catch (
            error
        ) {

            console.error(
                "Create user error:",
                error
            );


            res
                .status(
                    500
                )
                .json(
                    {
                        error:
                            error.message
                    }
                );

        }

    }
);


// ============================================================
// ACTIVATE / DEACTIVATE USER
// ============================================================

router.put(
    "/users/:id/status",

    requireRole(
        "admin"
    ),

    async (
        req,
        res
    ) => {

        try {

            const userId =
                Number(
                    req.params.id
                );


            const active =
                req.body.active ===
                true;


            if (
                !userId
            ) {

                return res
                    .status(
                        400
                    )
                    .json(
                        {
                            error:
                                "Invalid user ID."
                        }
                    );
            }


            if (
                userId ===
                Number(
                    req.session.user.id
                ) &&
                active ===
                false
            ) {

                return res
                    .status(
                        400
                    )
                    .json(
                        {
                            error:
                                "You cannot deactivate your own account."
                        }
                    );
            }


            const result =
                await pool.query(
                    `
                    UPDATE app_users

                    SET
                        active = $1,

                        updated_at =
                            CURRENT_TIMESTAMP

                    WHERE
                        id = $2

                    RETURNING
                        id,
                        username,
                        full_name,
                        role,
                        active
                    `,
                    [
                        active,
                        userId
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
                    .json(
                        {
                            error:
                                "User not found."
                        }
                    );
            }


            res.json(
                {

                    success:
                        true,

                    user:
                        result.rows[0]

                }
            );

        } catch (
            error
        ) {

            console.error(
                "Update user status error:",
                error
            );


            res
                .status(
                    500
                )
                .json(
                    {
                        error:
                            error.message
                    }
                );

        }

    }
);


// ============================================================
// EXPORTS
// ============================================================

module.exports =
    {

        authRouter:
            router,

        sessionMiddleware,

        requireAuth,

        requirePageAuth,

        requireRole,

        ensureDefaultAdmin

    };