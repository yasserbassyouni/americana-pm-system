CREATE TABLE IF NOT EXISTS production_lines (
    id SERIAL PRIMARY KEY,
    line_name VARCHAR(150) UNIQUE NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);


CREATE TABLE IF NOT EXISTS machines (
    id SERIAL PRIMARY KEY,
    production_line_id INTEGER REFERENCES production_lines(id) ON DELETE CASCADE,
    section VARCHAR(150),
    machine_no VARCHAR(100),
    machine_name VARCHAR(200) NOT NULL,
    machine_code VARCHAR(200) UNIQUE NOT NULL,
    machine_type VARCHAR(150),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);


CREATE TABLE IF NOT EXISTS pm_tasks (
    id SERIAL PRIMARY KEY,
    machine_id INTEGER REFERENCES machines(id) ON DELETE CASCADE,
    part_name VARCHAR(200),
    maintenance_task TEXT NOT NULL,
    frequency_text VARCHAR(100) NOT NULL,

    frequency_type VARCHAR(30),
    frequency_value INTEGER,
    last_completed_hours INTEGER DEFAULT 0,

    active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);


CREATE TABLE IF NOT EXISTS pm_schedule (
    id SERIAL PRIMARY KEY,
    pm_task_id INTEGER REFERENCES pm_tasks(id) ON DELETE CASCADE,

    planned_year INTEGER NOT NULL,
    planned_week INTEGER NOT NULL,

    status VARCHAR(50) DEFAULT 'Pending',

    deferred_reason TEXT,
    completed_at TIMESTAMP,
    technician_name VARCHAR(150),
    notes TEXT,

    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);


-- Add new frequency columns to an existing pm_tasks table
-- without deleting previously imported PM data.

ALTER TABLE pm_tasks
ADD COLUMN IF NOT EXISTS frequency_type VARCHAR(30);

ALTER TABLE pm_tasks
ADD COLUMN IF NOT EXISTS frequency_value INTEGER;

ALTER TABLE pm_tasks
ADD COLUMN IF NOT EXISTS last_completed_hours INTEGER DEFAULT 0;-- ============================================================
-- APPLICATION USERS
-- ============================================================

CREATE TABLE IF NOT EXISTS app_users
(
    id SERIAL PRIMARY KEY,

    username VARCHAR(80) NOT NULL UNIQUE,

    full_name VARCHAR(150) NOT NULL,

    password_hash TEXT NOT NULL,

    role VARCHAR(30) NOT NULL DEFAULT 'technician',

    active BOOLEAN NOT NULL DEFAULT TRUE,

    last_login_at TIMESTAMPTZ,

    password_changed_at TIMESTAMPTZ,

    created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,

    updated_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,

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
);


CREATE UNIQUE INDEX IF NOT EXISTS idx_app_users_username_lower
ON app_users
(
    LOWER(username)
);-- ============================================================
-- APPLICATION USERS
-- ============================================================

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
);


CREATE UNIQUE INDEX IF NOT EXISTS
    idx_app_users_username_lower

ON app_users
(
    LOWER(username)
);