let allUsers = [];

let currentUser = null;


// ============================================================
// START
// ============================================================

document.addEventListener(
    "DOMContentLoaded",
    async () => {

        setupEvents();

        await checkAdmin();

        await loadUsers();
    }
);


// ============================================================
// EVENTS
// ============================================================

function setupEvents() {

    document
        .getElementById(
            "backToPM"
        )
        .addEventListener(
            "click",
            () => {

                window.location.href =
                    "/";
            }
        );


    document
        .getElementById(
            "logoutButton"
        )
        .addEventListener(
            "click",
            logout
        );


    document
        .getElementById(
            "openAddUser"
        )
        .addEventListener(
            "click",
            openUserModal
        );


    document
        .getElementById(
            "closeUserModal"
        )
        .addEventListener(
            "click",
            closeUserModal
        );


    document
        .getElementById(
            "cancelUserButton"
        )
        .addEventListener(
            "click",
            closeUserModal
        );


    document
        .getElementById(
            "newUserForm"
        )
        .addEventListener(
            "submit",
            createUser
        );


    document
        .getElementById(
            "userSearch"
        )
        .addEventListener(
            "input",
            renderUsers
        );


    document
        .getElementById(
            "newUserRole"
        )
        .addEventListener(
            "change",
            updateRoleDescription
        );


    document
        .getElementById(
            "showNewPassword"
        )
        .addEventListener(
            "click",
            togglePassword
        );
}


// ============================================================
// VERIFY ADMIN
// ============================================================

async function checkAdmin() {

    try {

        const response =
            await fetch(
                "/api/auth/me",
                {
                    credentials:
                        "include"
                }
            );


        if (
            response.status ===
            401
        ) {

            window.location.href =
                "/login";

            return;
        }


        const data =
            await response.json();


        currentUser =
            data.user;


        if (
            !currentUser ||
            currentUser.role !==
            "admin"
        ) {

            alert(
                "Administrator access is required."
            );


            window.location.href =
                "/";

            return;
        }


        document
            .getElementById(
                "currentAdminName"
            )
            .textContent =
            currentUser.fullName ||
            currentUser.username ||
            "Administrator";


    } catch (
        error
    ) {

        console.error(
            error
        );


        window.location.href =
            "/login";
    }
}


// ============================================================
// LOAD USERS
// ============================================================

async function loadUsers() {

    const tbody =
        document
            .getElementById(
                "usersTableBody"
            );


    tbody.innerHTML = `

        <tr>

            <td colspan="7">
                Loading users...
            </td>

        </tr>
    `;


    try {

        const response =
            await fetch(
                "/api/auth/users",
                {
                    credentials:
                        "include"
                }
            );


        const data =
            await response.json();


        if (
            response.status ===
            401
        ) {

            window.location.href =
                "/login";

            return;
        }


        if (
            response.status ===
            403
        ) {

            alert(
                "Administrator access is required."
            );


            window.location.href =
                "/";

            return;
        }


        if (
            !response.ok
        ) {

            throw new Error(
                data.error ||
                "Unable to load users."
            );
        }


        allUsers =
            Array.isArray(
                data
            )
                ?
                data
                :
                [];


        updateSummary();

        renderUsers();


    } catch (
        error
    ) {

        showPageMessage(
            error.message,
            "error"
        );


        tbody.innerHTML = `

            <tr>

                <td colspan="7">
                    Unable to load users.
                </td>

            </tr>
        `;
    }
}


// ============================================================
// SUMMARY
// ============================================================

function updateSummary() {

    document
        .getElementById(
            "totalUsers"
        )
        .textContent =
        allUsers.length;


    document
        .getElementById(
            "activeUsers"
        )
        .textContent =
        allUsers.filter(
            user =>
                user.active ===
                true
        ).length;


    document
        .getElementById(
            "technicianUsers"
        )
        .textContent =
        allUsers.filter(
            user =>
                user.role ===
                "technician"
        ).length;


    document
        .getElementById(
            "supervisorUsers"
        )
        .textContent =
        allUsers.filter(
            user =>
                user.role ===
                "supervisor"
        ).length;
}


// ============================================================
// RENDER USERS
// ============================================================

function renderUsers() {

    const tbody =
        document
            .getElementById(
                "usersTableBody"
            );


    const search =
        document
            .getElementById(
                "userSearch"
            )
            .value
            .trim()
            .toLowerCase();


    let rows =
        [
            ...allUsers
        ];


    if (
        search
    ) {

        rows =
            rows.filter(
                user => {

                    const text =
                        [
                            user.full_name,
                            user.username,
                            user.role
                        ]
                            .join(
                                " "
                            )
                            .toLowerCase();


                    return text.includes(
                        search
                    );
                }
            );
    }


    tbody.innerHTML =
        "";


    if (
        rows.length ===
        0
    ) {

        tbody.innerHTML = `

            <tr>

                <td colspan="7">
                    No users found.
                </td>

            </tr>
        `;

        return;
    }


    for (
        const user
        of rows
    ) {

        const tr =
            document
                .createElement(
                    "tr"
                );


        const isCurrent =
            currentUser &&
            Number(
                currentUser.id
            ) ===
            Number(
                user.id
            );


        let actionButton;


        if (
            isCurrent
        ) {

            actionButton = `

                <span>
                    Current User
                </span>
            `;

        } else if (
            user.active
        ) {

            actionButton = `

                <button
                    class="deactivate-button"
                    type="button"
                    onclick="
                        changeUserStatus(
                            ${user.id},
                            false
                        )
                    "
                >
                    Deactivate
                </button>
            `;

        } else {

            actionButton = `

                <button
                    class="activate-button"
                    type="button"
                    onclick="
                        changeUserStatus(
                            ${user.id},
                            true
                        )
                    "
                >
                    Activate
                </button>
            `;
        }


        tr.innerHTML = `

            <td>

                <strong>
                    ${escapeHtml(
                        user.full_name ||
                        ""
                    )}
                </strong>

            </td>


            <td>
                ${escapeHtml(
                    user.username ||
                    ""
                )}
            </td>


            <td>

                <span class="
                    role-badge
                    ${getRoleClass(
                        user.role
                    )}
                ">
                    ${formatRole(
                        user.role
                    )}
                </span>

            </td>


            <td>

                <span class="
                    status-badge
                    ${
                        user.active
                            ?
                            "status-active"
                            :
                            "status-inactive"
                    }
                ">
                    ${
                        user.active
                            ?
                            "Active"
                            :
                            "Inactive"
                    }
                </span>

            </td>


            <td>
                ${
                    user.last_login_at
                        ?
                        formatDate(
                            user.last_login_at
                        )
                        :
                        "Never"
                }
            </td>


            <td>
                ${
                    user.created_at
                        ?
                        formatDate(
                            user.created_at
                        )
                        :
                        "-"
                }
            </td>


            <td>
                ${actionButton}
            </td>
        `;


        tbody.appendChild(
            tr
        );
    }
}


// ============================================================
// OPEN ADD USER
// ============================================================

function openUserModal() {

    document
        .getElementById(
            "newUserForm"
        )
        .reset();


    document
        .getElementById(
            "newUserRole"
        )
        .value =
        "technician";


    document
        .getElementById(
            "modalMessage"
        )
        .textContent =
        "";


    updateRoleDescription();


    document
        .getElementById(
            "userModal"
        )
        .classList
        .remove(
            "hidden"
        );


    setTimeout(
        () => {

            document
                .getElementById(
                    "newFullName"
                )
                .focus();

        },
        100
    );
}


// ============================================================
// CLOSE ADD USER
// ============================================================

function closeUserModal() {

    document
        .getElementById(
            "userModal"
        )
        .classList
        .add(
            "hidden"
        );
}


// ============================================================
// CREATE USER
// ============================================================

async function createUser(
    event
) {

    event.preventDefault();


    const fullName =
        document
            .getElementById(
                "newFullName"
            )
            .value
            .trim();


    const username =
        document
            .getElementById(
                "newUsername"
            )
            .value
            .trim()
            .toLowerCase();


    const password =
        document
            .getElementById(
                "newPassword"
            )
            .value;


    const confirmPassword =
        document
            .getElementById(
                "confirmNewPassword"
            )
            .value;


    const role =
        document
            .getElementById(
                "newUserRole"
            )
            .value;


    const message =
        document
            .getElementById(
                "modalMessage"
            );


    if (
        !fullName ||
        !username ||
        !password
    ) {

        showModalMessage(
            "Please complete all required fields.",
            "error"
        );

        return;
    }


    if (
        password.length <
        8
    ) {

        showModalMessage(
            "Password must contain at least 8 characters.",
            "error"
        );

        return;
    }


    if (
        password !==
        confirmPassword
    ) {

        showModalMessage(
            "Passwords do not match.",
            "error"
        );

        return;
    }


    const button =
        document
            .getElementById(
                "createUserButton"
            );


    const oldText =
        button.textContent;


    button.disabled =
        true;


    button.textContent =
        "Creating...";


    try {

        const response =
            await fetch(
                "/api/auth/users",
                {

                    method:
                        "POST",

                    credentials:
                        "include",

                    headers: {

                        "Content-Type":
                            "application/json"
                    },

                    body:
                        JSON.stringify({

                            fullName,

                            username,

                            password,

                            role
                        })
                }
            );


        const data =
            await response.json();


        if (
            !response.ok
        ) {

            throw new Error(
                data.error ||
                "Unable to create user."
            );
        }


        showModalMessage(
            "User created successfully.",
            "success"
        );


        await loadUsers();


        setTimeout(
            () => {

                closeUserModal();

                showPageMessage(
                    `User "${username}" created successfully.`,
                    "success"
                );

            },
            600
        );


    } catch (
        error
    ) {

        showModalMessage(
            error.message,
            "error"
        );


    } finally {

        button.disabled =
            false;


        button.textContent =
            oldText;
    }
}


// ============================================================
// ACTIVATE / DEACTIVATE
// ============================================================

async function changeUserStatus(
    userId,
    active
) {

    const action =
        active
            ?
            "activate"
            :
            "deactivate";


    if (
        !window.confirm(
            `Are you sure you want to ${action} this user?`
        )
    ) {

        return;
    }


    try {

        const response =
            await fetch(
                `/api/auth/users/${userId}/status`,
                {

                    method:
                        "PUT",

                    credentials:
                        "include",

                    headers: {

                        "Content-Type":
                            "application/json"
                    },

                    body:
                        JSON.stringify({

                            active
                        })
                }
            );


        const data =
            await response.json();


        if (
            !response.ok
        ) {

            throw new Error(
                data.error ||
                "Unable to update user."
            );
        }


        showPageMessage(
            `User ${active ? "activated" : "deactivated"} successfully.`,
            "success"
        );


        await loadUsers();


    } catch (
        error
    ) {

        showPageMessage(
            error.message,
            "error"
        );
    }
}


// ============================================================
// ROLE DESCRIPTION
// ============================================================

function updateRoleDescription() {

    const role =
        document
            .getElementById(
                "newUserRole"
            )
            .value;


    const descriptions =
    {

        admin:
            "Administrator has full access including Excel Import, users, reports and maintenance actions.",

        supervisor:
            "Supervisor can access schedules, reports and Complete / Defer PM tasks.",

        technician:
            "Technician can view PM schedules and Complete / Defer maintenance tasks.",

        viewer:
            "Viewer can view the system, calendar and reports but cannot Complete or Defer PM."
    };


    document
        .getElementById(
            "roleDescription"
        )
        .textContent =
        descriptions[
            role
        ] ||
        "";
}


// ============================================================
// PASSWORD VISIBILITY
// ============================================================

function togglePassword() {

    const input =
        document
            .getElementById(
                "newPassword"
            );


    const button =
        document
            .getElementById(
                "showNewPassword"
            );


    if (
        input.type ===
        "password"
    ) {

        input.type =
            "text";


        button.textContent =
            "Hide";

    } else {

        input.type =
            "password";


        button.textContent =
            "Show";
    }
}


// ============================================================
// LOGOUT
// ============================================================

async function logout() {

    try {

        await fetch(
            "/api/auth/logout",
            {

                method:
                    "POST",

                credentials:
                    "include"
            }
        );

    } catch (
        error
    ) {

        console.error(
            error
        );
    }


    window.location.href =
        "/login";
}


// ============================================================
// CLASSES
// ============================================================

function getRoleClass(
    role
) {

    return `role-${role}`;
}


function formatRole(
    role
) {

    const names =
    {

        admin:
            "Administrator",

        supervisor:
            "Supervisor",

        technician:
            "Technician",

        viewer:
            "Viewer"
    };


    return names[
        role
    ] ||
    role;
}


// ============================================================
// MESSAGES
// ============================================================

function showPageMessage(
    message,
    type
) {

    const element =
        document
            .getElementById(
                "pageMessage"
            );


    element.textContent =
        message;


    element.className =
        "page-message";


    element.classList.add(
        type ===
        "success"
            ?
            "message-success"
            :
            "message-error"
    );
}


function showModalMessage(
    message,
    type
) {

    const element =
        document
            .getElementById(
                "modalMessage"
            );


    element.textContent =
        message;


    element.className =
        "modal-message";


    element.classList.add(
        type ===
        "success"
            ?
            "message-success"
            :
            "message-error"
    );
}


// ============================================================
// DATE
// ============================================================

function formatDate(
    value
) {

    return new Date(
        value
    )
        .toLocaleString();
}


// ============================================================
// SAFE HTML
// ============================================================

function escapeHtml(
    value
) {

    return String(
        value ??
        ""
    )
        .replace(
            /&/g,
            "&amp;"
        )
        .replace(
            /</g,
            "&lt;"
        )
        .replace(
            />/g,
            "&gt;"
        )
        .replace(
            /"/g,
            "&quot;"
        )
        .replace(
            /'/g,
            "&#039;"
        );
}