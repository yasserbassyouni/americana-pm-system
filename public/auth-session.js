// ============================================================
// PM SYSTEM AUTHENTICATION
// ============================================================

let currentPMUser = null;


// ============================================================
// ORIGINAL FETCH
// ============================================================

const originalPMFetch =
    window.fetch.bind(
        window
    );


// ============================================================
// AUTOMATIC SESSION HANDLING
// ============================================================

window.fetch =
    async function (
        resource,
        options = {}
    ) {

        const finalOptions = {

            ...options,

            credentials:
                options.credentials ||
                "include"
        };


        const response =
            await originalPMFetch(
                resource,
                finalOptions
            );


        const resourceText =
            typeof resource ===
            "string"
                ?
                resource
                :
                resource.url;


        const isAuthRequest =
            resourceText.includes(
                "/api/auth/"
            );


        if (
            response.status ===
            401 &&
            !isAuthRequest
        ) {

            window.location.href =
                "/login";
        }


        return response;
    };


// ============================================================
// START
// ============================================================

document.addEventListener(
    "DOMContentLoaded",
    async () => {

        await loadCurrentPMUser();
    }
);


// ============================================================
// LOAD CURRENT USER
// ============================================================

async function loadCurrentPMUser() {

    try {

        const response =
            await originalPMFetch(
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


        if (
            !response.ok
        ) {

            throw new Error(
                "Unable to verify login session."
            );
        }


        const data =
            await response.json();


        currentPMUser =
            data.user;


        window.PM_CURRENT_USER =
            currentPMUser;


        renderUserHeader();


        applyRolePermissions();


    } catch (
        error
    ) {

        console.error(
            "Authentication error:",
            error
        );
    }
}


// ============================================================
// FIND EXISTING BLUE HEADER
// ============================================================

function findPMHeader() {

    return (

        document.querySelector(
            ".top-header"
        )

        ||

        document.querySelector(
            "header"
        )

        ||

        document.querySelector(
            ".header"
        )

        ||

        document.querySelector(
            ".app-header"
        )

        ||

        document.querySelector(
            ".main-header"
        )
    );
}


// ============================================================
// USER HEADER
// ============================================================

function renderUserHeader() {

    if (
        !currentPMUser
    ) {

        return;
    }


    const header =
        findPMHeader();


    if (
        !header
    ) {

        console.error(
            "PM header was not found."
        );

        return;
    }


    // Prevent duplicate controls

    const existingArea =
        document.getElementById(
            "pmUserArea"
        );


    if (
        existingArea
    ) {

        existingArea.remove();
    }


    const userArea =
        document.createElement(
            "div"
        );


    userArea.id =
        "pmUserArea";


    const isAdmin =
        currentPMUser.role ===
        "admin";


    userArea.innerHTML = `

        <div class="pm-user-info">

            <strong>

                ${escapeAuthHtml(
                    currentPMUser.fullName ||
                    currentPMUser.full_name ||
                    currentPMUser.username ||
                    "User"
                )}

            </strong>


            <span>

                ${escapeAuthHtml(
                    formatAuthRole(
                        currentPMUser.role
                    )
                )}

            </span>

        </div>


        ${
            isAdmin
                ?
                `

                <button
                    id="pmUsersButton"
                    class="pm-header-button"
                    type="button"
                >
                    Users
                </button>

                `
                :
                ""
        }


        <button
            id="pmPasswordButton"
            class="pm-header-button"
            type="button"
        >
            Password
        </button>


        <button
            id="pmLogoutButton"
            class="pm-logout-button"
            type="button"
        >
            Logout
        </button>
    `;


    injectAuthStyles();


    // Put controls at far right of header

    header.appendChild(
        userArea
    );


    // ========================================================
    // ADMIN USERS BUTTON
    // ========================================================

    if (
        isAdmin
    ) {

        const usersButton =
            document.getElementById(
                "pmUsersButton"
            );


        if (
            usersButton
        ) {

            usersButton.addEventListener(
                "click",
                () => {

                    window.location.href =
                        "/users.html";
                }
            );
        }
    }


    // ========================================================
    // PASSWORD BUTTON
    // ========================================================

    const passwordButton =
        document.getElementById(
            "pmPasswordButton"
        );


    if (
        passwordButton
    ) {

        passwordButton.addEventListener(
            "click",
            changePassword
        );
    }


    // ========================================================
    // LOGOUT BUTTON
    // ========================================================

    const logoutButton =
        document.getElementById(
            "pmLogoutButton"
        );


    if (
        logoutButton
    ) {

        logoutButton.addEventListener(
            "click",
            logoutPMUser
        );
    }
}


// ============================================================
// ROLE PERMISSIONS
// ============================================================

function applyRolePermissions() {

    if (
        !currentPMUser
    ) {

        return;
    }


    const role =
        String(
            currentPMUser.role ||
            ""
        )
            .toLowerCase();


    // ========================================================
    // EXCEL IMPORT = ADMIN ONLY
    // ========================================================

    if (
        role !==
        "admin"
    ) {

        const importButton =

            document.getElementById(
                "importNav"
            )

            ||

            document.getElementById(
                "importExcelNav"
            );


        if (
            importButton
        ) {

            importButton.style.display =
                "none";
        }
    }


    // ========================================================
    // VIEWER = READ ONLY
    // ========================================================

    if (
        role ===
        "viewer"
    ) {

        disablePMActionButtons();


        const observer =
            new MutationObserver(
                disablePMActionButtons
            );


        observer.observe(
            document.body,
            {

                childList:
                    true,

                subtree:
                    true
            }
        );
    }
}


// ============================================================
// VIEWER ACTIONS
// ============================================================

function disablePMActionButtons() {

    const selectors = [

        ".complete-button",

        ".defer-button",

        "#confirmAction",

        "#confirmExcelImport",

        "#previewExcelImport"
    ];


    for (
        const selector
        of selectors
    ) {

        document
            .querySelectorAll(
                selector
            )
            .forEach(
                element => {

                    element.style.display =
                        "none";
                }
            );
    }
}


// ============================================================
// CHANGE PASSWORD
// ============================================================

async function changePassword() {

    const currentPassword =
        window.prompt(
            "Enter your current password:"
        );


    if (
        currentPassword ===
        null
    ) {

        return;
    }


    const newPassword =
        window.prompt(
            "Enter new password (minimum 8 characters):"
        );


    if (
        newPassword ===
        null
    ) {

        return;
    }


    if (
        newPassword.length <
        8
    ) {

        alert(
            "New password must contain at least 8 characters."
        );

        return;
    }


    const confirmPassword =
        window.prompt(
            "Enter new password again:"
        );


    if (
        confirmPassword ===
        null
    ) {

        return;
    }


    if (
        newPassword !==
        confirmPassword
    ) {

        alert(
            "Passwords do not match."
        );

        return;
    }


    try {

        const response =
            await fetch(
                "/api/auth/change-password",
                {

                    method:
                        "POST",

                    headers: {

                        "Content-Type":
                            "application/json"
                    },

                    body:
                        JSON.stringify({

                            currentPassword,

                            newPassword
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
                "Password change failed."
            );
        }


        alert(
            "Password changed successfully."
        );


    } catch (
        error
    ) {

        alert(
            error.message
        );
    }
}


// ============================================================
// LOGOUT
// ============================================================

async function logoutPMUser() {

    const confirmed =
        window.confirm(
            "Do you want to sign out?"
        );


    if (
        !confirmed
    ) {

        return;
    }


    try {

        await originalPMFetch(
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
            "Logout error:",
            error
        );
    }


    window.location.href =
        "/login";
}


// ============================================================
// ROLE NAME
// ============================================================

function formatAuthRole(
    role
) {

    const roles = {

        admin:
            "Administrator",

        supervisor:
            "Supervisor",

        technician:
            "Technician",

        viewer:
            "Viewer"
    };


    return roles[
        role
    ] ||
    role;
}


// ============================================================
// ESCAPE HTML
// ============================================================

function escapeAuthHtml(
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


// ============================================================
// AUTH HEADER STYLES
// ============================================================

function injectAuthStyles() {

    if (
        document.getElementById(
            "pmAuthStyles"
        )
    ) {

        return;
    }


    const style =
        document.createElement(
            "style"
        );


    style.id =
        "pmAuthStyles";


    style.textContent = `

        #pmUserArea {

            display:
                flex;

            align-items:
                center;

            gap:
                8px;

            margin-left:
                auto;

            margin-right:
                12px;

            padding:
                6px 8px;

            border-radius:
                10px;

            background:
                rgba(
                    255,
                    255,
                    255,
                    0.10
                );

            flex-shrink:
                0;
        }


        .pm-user-info {

            display:
                flex;

            flex-direction:
                column;

            min-width:
                105px;

            line-height:
                1.2;
        }


        .pm-user-info strong {

            color:
                white;

            font-size:
                12px;

            white-space:
                nowrap;
        }


        .pm-user-info span {

            color:
                rgba(
                    255,
                    255,
                    255,
                    0.75
                );

            font-size:
                10px;

            margin-top:
                2px;
        }


        .pm-header-button,
        .pm-logout-button {

            height:
                32px;

            border:
                0;

            border-radius:
                7px;

            padding:
                0 10px;

            color:
                white;

            font-size:
                11px;

            font-weight:
                700;

            cursor:
                pointer;

            white-space:
                nowrap;
        }


        .pm-header-button {

            background:
                rgba(
                    255,
                    255,
                    255,
                    0.16
                );
        }


        .pm-header-button:hover {

            background:
                rgba(
                    255,
                    255,
                    255,
                    0.27
                );
        }


        .pm-logout-button {

            background:
                #dc2626;
        }


        .pm-logout-button:hover {

            background:
                #b91c1c;
        }


        @media
        (
            max-width:
                1100px
        ) {

            #pmUserArea {

                gap:
                    5px;
            }


            .pm-user-info {

                min-width:
                    80px;
            }


            .pm-header-button,
            .pm-logout-button {

                padding:
                    0 7px;

                font-size:
                    10px;
            }
        }
    `;


    document.head.appendChild(
        style
    );
}