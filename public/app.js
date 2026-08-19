// ============================================================
// GLOBAL DATA
// ============================================================

let allSchedule = [];

let annualPlanData = [];

let loadedAnnualYear = null;

let currentLineId = null;
let currentLineName = "";

let currentMachineId = null;
let currentMachineName = "";
let currentMachineCode = "";

let currentScheduleId = null;
let currentAction = null;

let importPreviewToken = null;


// ============================================================
// CURRENT WEEK
// ============================================================

let selectedYear = 2026;
let selectedWeek = 34;


// ============================================================
// START APPLICATION
// ============================================================

document.addEventListener(
    "DOMContentLoaded",
    async () => {

        updateWeekDisplay();

        setupEventListeners();

        populateWeekSelectors();

        await loadSchedule();

        showMainView(
            "dashboard"
        );
    }
);


// ============================================================
// EVENT LISTENERS
// ============================================================

function setupEventListeners() {

    // ========================================================
    // WEEK NAVIGATION
    // ========================================================

    byId("previousWeek")
        .addEventListener(
            "click",
            previousWeek
        );


    byId("nextWeek")
        .addEventListener(
            "click",
            nextWeek
        );


    // ========================================================
    // MAIN NAVIGATION
    // ========================================================

    byId("dashboardNav")
        .addEventListener(
            "click",
            () =>
                showMainView(
                    "dashboard"
                )
        );


    byId("scheduleNav")
        .addEventListener(
            "click",
            () =>
                showMainView(
                    "schedule"
                )
        );


    byId("annualPlanNav")
        .addEventListener(
            "click",
            () =>
                showMainView(
                    "annual"
                )
        );


    byId("reportsNav")
        .addEventListener(
            "click",
            () =>
                showMainView(
                    "reports"
                )
        );


    byId("calendarNav")
        .addEventListener(
            "click",
            () =>
                showMainView(
                    "calendar"
                )
        );


    byId("importNav")
        .addEventListener(
            "click",
            () =>
                showMainView(
                    "import"
                )
        );


    // ========================================================
    // DASHBOARD
    // ========================================================

    byId("backToLines")
        .addEventListener(
            "click",
            showLines
        );


    byId("backToMachines")
        .addEventListener(
            "click",
            showMachines
        );


    // ========================================================
    // PM ACTION MODAL
    // ========================================================

    byId("closeModal")
        .addEventListener(
            "click",
            closeModal
        );


    byId("cancelAction")
        .addEventListener(
            "click",
            closeModal
        );


    byId("confirmAction")
        .addEventListener(
            "click",
            confirmPMAction
        );


    // ========================================================
    // WEEKLY SCHEDULE FILTERS
    // ========================================================

    byId("scheduleLineFilter")
        .addEventListener(
            "change",
            renderScheduleTable
        );


    byId("scheduleStatusFilter")
        .addEventListener(
            "change",
            renderScheduleTable
        );


    byId("scheduleFrequencyFilter")
        .addEventListener(
            "change",
            renderScheduleTable
        );


    byId("scheduleSearch")
        .addEventListener(
            "input",
            renderScheduleTable
        );


    // ========================================================
    // ANNUAL PLAN
    // ========================================================

    byId("annualYearSelect")
        .addEventListener(
            "change",
            async event => {

                await loadAnnualPlan(
                    Number(
                        event.target.value
                    )
                );
            }
        );


    byId("annualLineFilter")
        .addEventListener(
            "change",
            renderAnnualPlanTable
        );


    byId("annualWeekFilter")
        .addEventListener(
            "change",
            renderAnnualPlanTable
        );


    byId("annualStatusFilter")
        .addEventListener(
            "change",
            renderAnnualPlanTable
        );


    byId("annualFrequencyFilter")
        .addEventListener(
            "change",
            renderAnnualPlanTable
        );


    byId("annualSearch")
        .addEventListener(
            "input",
            renderAnnualPlanTable
        );


    byId("exportAnnualExcel")
        .addEventListener(
            "click",
            exportAnnualExcel
        );


    byId("exportAnnualPdf")
        .addEventListener(
            "click",
            exportAnnualPdf
        );


    // ========================================================
    // REPORTS
    // ========================================================

    byId("reportYearSelect")
        .addEventListener(
            "change",
            async event => {

                await loadAnnualPlan(
                    Number(
                        event.target.value
                    )
                );

                renderReports();
            }
        );


    byId("reportLineFilter")
        .addEventListener(
            "change",
            () => {

                populateReportMachineFilter();

                renderReports();
            }
        );


    byId("reportMachineFilter")
        .addEventListener(
            "change",
            renderReports
        );


    byId("reportStatusFilter")
        .addEventListener(
            "change",
            renderReportActions
        );


    byId("reportWeekFilter")
        .addEventListener(
            "change",
            renderReportActions
        );


    byId("reportSearch")
        .addEventListener(
            "input",
            renderReportActions
        );


    byId("exportReportExcel")
        .addEventListener(
            "click",
            exportReportExcel
        );


    byId("exportReportPdf")
        .addEventListener(
            "click",
            exportReportPdf
        );


    // ========================================================
    // PM CALENDAR
    // ========================================================

    byId("calendarYearFilter")
        .addEventListener(
            "change",
            async event => {

                await loadAnnualPlan(
                    Number(
                        event.target.value
                    )
                );

                resetCalendarFilters(
                    false
                );

                renderCalendarMatrix();
            }
        );


    byId("calendarLineFilter")
        .addEventListener(
            "change",
            () => {

                populateCalendarMachineFilter();
            }
        );


    byId("calendarMonthFilter")
        .addEventListener(
            "change",
            applyCalendarMonthToWeeks
        );


    byId("applyCalendarFilters")
        .addEventListener(
            "click",
            renderCalendarMatrix
        );


    byId("resetCalendarFilters")
        .addEventListener(
            "click",
            () =>
                resetCalendarFilters(
                    true
                )
        );


    byId("exportCalendarExcel")
        .addEventListener(
            "click",
            exportCalendarExcel
        );


    byId("exportCalendarPdf")
        .addEventListener(
            "click",
            exportCalendarPdf
        );


    // ========================================================
    // EXCEL IMPORT
    // ========================================================

    byId("pmExcelFile")
        .addEventListener(
            "change",
            handleExcelFileSelected
        );


    byId("previewExcelImport")
        .addEventListener(
            "click",
            previewExcelImport
        );


    byId("confirmExcelImport")
        .addEventListener(
            "click",
            confirmExcelImport
        );
}


// ============================================================
// MAIN VIEW
// ============================================================

async function showMainView(
    view
) {

    const views = {

        dashboard:
            byId(
                "dashboardView"
            ),

        schedule:
            byId(
                "scheduleView"
            ),

        annual:
            byId(
                "annualPlanView"
            ),

        reports:
            byId(
                "reportsView"
            ),

        calendar:
            byId(
                "calendarView"
            ),

        import:
            byId(
                "importView"
            )
    };


    const buttons = {

        dashboard:
            byId(
                "dashboardNav"
            ),

        schedule:
            byId(
                "scheduleNav"
            ),

        annual:
            byId(
                "annualPlanNav"
            ),

        reports:
            byId(
                "reportsNav"
            ),

        calendar:
            byId(
                "calendarNav"
            ),

        import:
            byId(
                "importNav"
            )
    };


    Object
        .values(
            views
        )
        .forEach(
            element => {

                element
                    .classList
                    .add(
                        "hidden"
                    );
            }
        );


    Object
        .values(
            buttons
        )
        .forEach(
            element => {

                element
                    .classList
                    .remove(
                        "active"
                    );
            }
        );


    if (
        !views[
            view
        ]
    ) {

        view =
            "dashboard";
    }


    views[
        view
    ]
        .classList
        .remove(
            "hidden"
        );


    buttons[
        view
    ]
        .classList
        .add(
            "active"
        );


    // ========================================================
    // DASHBOARD
    // ========================================================

    if (
        view ===
        "dashboard"
    ) {

        showLines();

        return;
    }


    // ========================================================
    // WEEKLY SCHEDULE
    // ========================================================

    if (
        view ===
        "schedule"
    ) {

        populateScheduleFilters();

        renderScheduleTable();

        return;
    }


    // ========================================================
    // ANNUAL PLAN
    // ========================================================

    if (
        view ===
        "annual"
    ) {

        const year =
            Number(
                byId(
                    "annualYearSelect"
                ).value
            );


        await loadAnnualPlan(
            year
        );

        return;
    }


    // ========================================================
    // REPORTS
    // ========================================================

    if (
        view ===
        "reports"
    ) {

        const year =
            Number(
                byId(
                    "reportYearSelect"
                ).value
            );


        await loadAnnualPlan(
            year
        );


        renderReports();

        return;
    }


    // ========================================================
    // CALENDAR
    // ========================================================

    if (
        view ===
        "calendar"
    ) {

        const year =
            Number(
                byId(
                    "calendarYearFilter"
                ).value
            );


        await loadAnnualPlan(
            year
        );


        populateCalendarFilters();

        renderCalendarMatrix();

        return;
    }
}


// ============================================================
// WEEK NAVIGATION
// ============================================================

// ============================================================
// UPDATE WEEK DISPLAY
// ============================================================

function updateWeekDisplay() {

    // ========================================================
    // WEEK NUMBER
    // ========================================================

    const weekElement =
        document.getElementById(
            "currentWeek"
        );


    if (
        weekElement
    ) {

        weekElement.textContent =
            selectedWeek;
    }


    // ========================================================
    // YEAR
    // ========================================================

    const yearElement =
        document.getElementById(
            "currentYear"
        );


    if (
        yearElement
    ) {

        yearElement.textContent =
            selectedYear;
    }


    // ========================================================
    // WEEK START / END DATE
    // ========================================================

    const dateRangeElement =
        document.getElementById(
            "currentWeekDates"
        );


    if (
        dateRangeElement
    ) {

        const range =
            getISOWeekDateRange(
                selectedYear,
                selectedWeek
            );


        dateRangeElement.textContent =
            range.text;
    }
}

// ============================================================
// LOAD WEEKLY SCHEDULE
// ============================================================

async function loadSchedule() {

    try {

        let response =
            await fetch(
                `/api/schedule?year=${selectedYear}&week=${selectedWeek}`
            );


        let data =
            await response.json();


        // ----------------------------------------------------
        // Generate missing week
        // ----------------------------------------------------

        if (
            response.ok &&
            Array.isArray(
                data
            ) &&
            data.length ===
            0
        ) {

            const generateResponse =
                await fetch(
                    "/api/schedule/generate",
                    {

                        method:
                            "POST",

                        headers: {

                            "Content-Type":
                                "application/json"
                        },

                        body:
                            JSON.stringify({

                                year:
                                    selectedYear,

                                week:
                                    selectedWeek
                            })
                    }
                );


            if (
                generateResponse.ok
            ) {

                response =
                    await fetch(
                        `/api/schedule?year=${selectedYear}&week=${selectedWeek}`
                    );


                data =
                    await response.json();
            }
        }


        if (
            response.ok &&
            Array.isArray(
                data
            )
        ) {

            allSchedule =
                data;

        } else {

            allSchedule =
                [];
        }


        updateSummary();


        await loadLines();


        populateScheduleFilters();


        renderScheduleTable();


    } catch (
        error
    ) {

        console.error(
            "Weekly schedule error:",
            error
        );


        allSchedule =
            [];


        updateSummary();
    }
}


// ============================================================
// DASHBOARD SUMMARY
// ============================================================

function updateSummary() {

    setText(
        "pendingCount",
        countStatus(
            allSchedule,
            "Pending"
        )
    );


    setText(
        "completedCount",
        countStatus(
            allSchedule,
            "Completed"
        )
    );


    setText(
        "deferredCount",
        countStatus(
            allSchedule,
            "Deferred"
        )
    );


    setText(
        "totalCount",
        allSchedule.length
    );
}


// ============================================================
// PRODUCTION LINES
// ============================================================

async function loadLines() {

    const container =
        byId(
            "lineButtons"
        );


    container.innerHTML = `

        <div class="loading">
            Loading production lines...
        </div>
    `;


    try {

        const response =
            await fetch(
                "/api/lines"
            );


        const lines =
            await response.json();


        container.innerHTML =
            "";


        for (
            const line
            of lines
        ) {

            const rows =
                allSchedule.filter(
                    item =>
                        item.line_name ===
                        line.line_name
                );


            const card =
                document
                    .createElement(
                        "div"
                    );


            card.className =
                "line-card";


            card.innerHTML = `

                <h3>
                    ${escapeHtml(
                        line.line_name
                    )}
                </h3>

                <p>
                    View machines and PM tasks
                </p>


                <div class="pm-meta">

                    <span class="badge">
                        Total:
                        ${rows.length}
                    </span>


                    <span class="
                        badge
                        status-pending
                    ">
                        Pending:
                        ${countStatus(
                            rows,
                            "Pending"
                        )}
                    </span>


                    <span class="
                        badge
                        status-completed
                    ">
                        Completed:
                        ${countStatus(
                            rows,
                            "Completed"
                        )}
                    </span>


                    <span class="
                        badge
                        status-deferred
                    ">
                        Deferred:
                        ${countStatus(
                            rows,
                            "Deferred"
                        )}
                    </span>

                </div>
            `;


            card
                .addEventListener(
                    "click",
                    () =>
                        selectLine(
                            line.id,
                            line.line_name
                        )
                );


            container
                .appendChild(
                    card
                );
        }


    } catch (
        error
    ) {

        console.error(
            error
        );


        container.innerHTML = `

            <div class="loading">
                Failed to load production lines.
            </div>
        `;
    }
}


// ============================================================
// SELECT LINE
// ============================================================

async function selectLine(
    lineId,
    lineName
) {

    currentLineId =
        lineId;

    currentLineName =
        lineName;


    const linePanel =
        document
            .querySelector(
                "#dashboardView > .panel"
            );


    if (
        linePanel
    ) {

        linePanel
            .classList
            .add(
                "hidden"
            );
    }


    byId(
        "machineSection"
    )
        .classList
        .remove(
            "hidden"
        );


    byId(
        "pmSection"
    )
        .classList
        .add(
            "hidden"
        );


    setText(
        "selectedLineTitle",
        lineName
    );


    await loadMachines(
        lineId
    );
}


// ============================================================
// MACHINES
// ============================================================

async function loadMachines(
    lineId
) {

    const container =
        byId(
            "machineGrid"
        );


    container.innerHTML = `

        <div class="loading">
            Loading machines...
        </div>
    `;


    try {

        const response =
            await fetch(
                `/api/lines/${lineId}/machines`
            );


        const machines =
            await response.json();


        container.innerHTML =
            "";


        for (
            const machine
            of machines
        ) {

            const rows =
                allSchedule.filter(
                    item =>
                        Number(
                            item.machine_id
                        ) ===
                        Number(
                            machine.id
                        )
                );


            const card =
                document
                    .createElement(
                        "div"
                    );


            card.className =
                "machine-card";


            card.innerHTML = `

                <span class="machine-code">
                    ${escapeHtml(
                        machine.machine_code ||
                        ""
                    )}
                </span>


                <h3>
                    ${escapeHtml(
                        machine.machine_name ||
                        ""
                    )}
                </h3>


                <p>

                    <strong>
                        Section:
                    </strong>

                    ${escapeHtml(
                        machine.section ||
                        "-"
                    )}

                </p>


                <p>

                    <strong>
                        Type:
                    </strong>

                    ${escapeHtml(
                        machine.machine_type ||
                        "-"
                    )}

                </p>


                <div class="pm-meta">

                    <span class="badge">
                        Total:
                        ${rows.length}
                    </span>


                    <span class="
                        badge
                        status-pending
                    ">
                        Pending:
                        ${countStatus(
                            rows,
                            "Pending"
                        )}
                    </span>


                    <span class="
                        badge
                        status-completed
                    ">
                        Completed:
                        ${countStatus(
                            rows,
                            "Completed"
                        )}
                    </span>


                    <span class="
                        badge
                        status-deferred
                    ">
                        Deferred:
                        ${countStatus(
                            rows,
                            "Deferred"
                        )}
                    </span>

                </div>
            `;


            card
                .addEventListener(
                    "click",
                    () =>
                        selectMachine(
                            machine
                        )
                );


            container
                .appendChild(
                    card
                );
        }


    } catch (
        error
    ) {

        console.error(
            error
        );


        container.innerHTML = `

            <div class="loading">
                Failed to load machines.
            </div>
        `;
    }
}


// ============================================================
// SELECT MACHINE
// ============================================================

function selectMachine(
    machine
) {

    currentMachineId =
        machine.id;

    currentMachineName =
        machine.machine_name;

    currentMachineCode =
        machine.machine_code;


    byId(
        "machineSection"
    )
        .classList
        .add(
            "hidden"
        );


    byId(
        "pmSection"
    )
        .classList
        .remove(
            "hidden"
        );


    setText(
        "selectedMachineTitle",
        `${machine.machine_code || ""} - ${machine.machine_name || ""}`
    );


    setText(
        "selectedMachineInfo",
        `${currentLineName} | ${machine.section || ""}`
    );


    showMachinePMs(
        machine.id
    );
}


// ============================================================
// MACHINE PM TASKS
// ============================================================

function showMachinePMs(
    machineId
) {

    const container =
        byId(
            "pmList"
        );


    const tasks =
        allSchedule
            .filter(
                item =>
                    Number(
                        item.machine_id
                    ) ===
                    Number(
                        machineId
                    )
            )
            .sort(
                (
                    a,
                    b
                ) =>
                    getStatusPriority(
                        a.status
                    ) -
                    getStatusPriority(
                        b.status
                    )
            );


    container.innerHTML =
        "";


    if (
        tasks.length ===
        0
    ) {

        container.innerHTML = `

            <div class="pm-card">

                <div class="pm-task">
                    No PM tasks scheduled.
                </div>

                <div class="pm-part">
                    Week ${selectedWeek} / ${selectedYear}
                </div>

            </div>
        `;

        return;
    }


    for (
        const task
        of tasks
    ) {

        const card =
            document
                .createElement(
                    "div"
                );


        card.className =
            `pm-card ${getCardStatusClass(
                task.status
            )}`;


        let details =
            "";


        if (
            task.completed_at
        ) {

            details += `

                <div class="pm-meta">

                    <span class="
                        badge
                        status-completed
                    ">
                        Completed:
                        ${formatDate(
                            task.completed_at
                        )}
                    </span>

                </div>
            `;
        }


        if (
            task.deferred_reason
        ) {

            details += `

                <div class="pm-meta">

                    <span class="
                        badge
                        status-deferred
                    ">
                        Reason:
                        ${escapeHtml(
                            task.deferred_reason
                        )}
                    </span>

                </div>
            `;
        }


        if (
            task.technician_name
        ) {

            details += `

                <div class="pm-meta">

                    <span class="badge">
                        Technician:
                        ${escapeHtml(
                            task.technician_name
                        )}
                    </span>

                </div>
            `;
        }


        if (
            task.notes
        ) {

            details += `

                <div class="pm-meta">

                    <span class="badge">
                        Notes:
                        ${escapeHtml(
                            task.notes
                        )}
                    </span>

                </div>
            `;
        }


        let actions =
            "";


        if (
            task.status !==
            "Completed"
        ) {

            actions = `

                <div class="pm-actions">

                    <button
                        class="
                            action-button
                            complete-button
                        "
                        type="button"
                        onclick="
                            openPMModal(
                                ${task.id},
                                'complete'
                            )
                        "
                    >
                        ✓ Complete
                    </button>


                    <button
                        class="
                            action-button
                            defer-button
                        "
                        type="button"
                        onclick="
                            openPMModal(
                                ${task.id},
                                'defer'
                            )
                        "
                    >
                        Deferred
                    </button>

                </div>
            `;
        }


        card.innerHTML = `

            <div class="pm-card-top">

                <div>

                    <div class="pm-part">
                        ${escapeHtml(
                            task.part_name ||
                            "General"
                        )}
                    </div>


                    <div class="pm-task">
                        ${escapeHtml(
                            task.maintenance_task ||
                            ""
                        )}
                    </div>

                </div>


                <span class="
                    badge
                    ${getStatusClass(
                        task.status
                    )}
                ">
                    ${escapeHtml(
                        task.status
                    )}
                </span>

            </div>


            <div class="pm-meta">

                <span class="
                    badge
                    badge-frequency
                ">
                    ${escapeHtml(
                        task.frequency_text ||
                        ""
                    )}
                </span>


                <span class="badge">
                    Week
                    ${task.planned_week}
                </span>


                <span class="badge">
                    ${task.planned_year}
                </span>

            </div>


            ${details}

            ${actions}
        `;


        container
            .appendChild(
                card
            );
    }
}


// ============================================================
// WEEKLY SCHEDULE FILTERS
// ============================================================

function populateScheduleFilters() {

    populateSelectFromData(
        "scheduleLineFilter",
        allSchedule,
        "line_name",
        "All Production Lines"
    );


    populateSelectFromData(
        "scheduleFrequencyFilter",
        allSchedule,
        "frequency_text",
        "All Frequencies"
    );
}


// ============================================================
// WEEKLY SCHEDULE TABLE
// ============================================================

function renderScheduleTable() {

    const tbody =
        byId(
            "scheduleTableBody"
        );


    const line =
        byId(
            "scheduleLineFilter"
        ).value;


    const status =
        byId(
            "scheduleStatusFilter"
        ).value;


    const frequency =
        byId(
            "scheduleFrequencyFilter"
        ).value;


    const search =
        byId(
            "scheduleSearch"
        )
            .value
            .trim()
            .toLowerCase();


    let rows =
        [
            ...allSchedule
        ];


    if (
        line
    ) {

        rows =
            rows.filter(
                item =>
                    item.line_name ===
                    line
            );
    }


    if (
        status
    ) {

        rows =
            rows.filter(
                item =>
                    item.status ===
                    status
            );
    }


    if (
        frequency
    ) {

        rows =
            rows.filter(
                item =>
                    item.frequency_text ===
                    frequency
            );
    }


    if (
        search
    ) {

        rows =
            rows.filter(
                item =>
                    rowMatchesSearch(
                        item,
                        search
                    )
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

                <td
                    colspan="9"
                    style="text-align:center;"
                >
                    No PM records found.
                </td>

            </tr>
        `;

        return;
    }


    for (
        const item
        of rows
    ) {

        const tr =
            document
                .createElement(
                    "tr"
                );


        tr.innerHTML = `

            <td>
                ${escapeHtml(
                    item.line_name ||
                    ""
                )}
            </td>


            <td>

                <strong>
                    ${escapeHtml(
                        item.machine_code ||
                        ""
                    )}
                </strong>

                <br>

                ${escapeHtml(
                    item.machine_name ||
                    ""
                )}

            </td>


            <td>
                ${escapeHtml(
                    item.section ||
                    ""
                )}
            </td>


            <td>
                ${escapeHtml(
                    item.part_name ||
                    ""
                )}
            </td>


            <td>
                ${escapeHtml(
                    item.maintenance_task ||
                    ""
                )}
            </td>


            <td>
                ${escapeHtml(
                    item.frequency_text ||
                    ""
                )}
            </td>


            <td>

                <span class="
                    table-status
                    ${getStatusClass(
                        item.status
                    )}
                ">
                    ${escapeHtml(
                        item.status
                    )}
                </span>

            </td>


            <td>
                ${escapeHtml(
                    item.technician_name ||
                    "-"
                )}
            </td>


            <td>
                ${
                    item.completed_at
                        ?
                        formatDate(
                            item.completed_at
                        )
                        :
                        "-"
                }
            </td>
        `;


        tbody
            .appendChild(
                tr
            );
    }
}


// ============================================================
// LOAD ANNUAL PLAN
// ============================================================

async function loadAnnualPlan(
    year
) {

    if (
        loadedAnnualYear ===
        year &&
        annualPlanData.length >
        0
    ) {

        updateAnnualInterface();

        return;
    }


    try {

        const response =
            await fetch(
                `/api/annual-plan?year=${year}`
            );


        const data =
            await response.json();


        if (
            !response.ok
        ) {

            throw new Error(
                data.error ||
                "Annual plan loading failed."
            );
        }


        annualPlanData =
            Array.isArray(
                data.schedule
            )
                ?
                data.schedule
                :
                [];


        loadedAnnualYear =
            year;


        syncYearSelectors(
            year
        );


        updateAnnualInterface();


    } catch (
        error
    ) {

        console.error(
            "Annual plan error:",
            error
        );


        annualPlanData =
            [];


        setImportMessage(
            `Annual plan error: ${error.message}`,
            "error"
        );
    }
}


// ============================================================
// UPDATE ALL ANNUAL-BASED SCREENS
// ============================================================

function updateAnnualInterface() {

    updateAnnualSummary();

    populateAnnualFilters();

    renderAnnualPlanTable();

    populateReportFilters();

    renderReports();

    populateCalendarFilters();
}


// ============================================================
// ANNUAL SUMMARY
// ============================================================

function updateAnnualSummary() {

    setText(
        "annualTotal",
        annualPlanData.length
    );


    setText(
        "annualCompleted",
        countStatus(
            annualPlanData,
            "Completed"
        )
    );


    setText(
        "annualPending",
        countStatus(
            annualPlanData,
            "Pending"
        )
    );


    setText(
        "annualDeferred",
        countStatus(
            annualPlanData,
            "Deferred"
        )
    );
}


// ============================================================
// ANNUAL FILTERS
// ============================================================

function populateAnnualFilters() {

    populateSelectFromData(
        "annualLineFilter",
        annualPlanData,
        "line_name",
        "All Production Lines"
    );


    populateSelectFromData(
        "annualFrequencyFilter",
        annualPlanData,
        "frequency_text",
        "All Frequencies"
    );


    populateWeekSelect(
        "annualWeekFilter",
        true
    );
}


// ============================================================
// ANNUAL FILTERED DATA
// ============================================================

function getFilteredAnnualRows() {

    const line =
        byId(
            "annualLineFilter"
        ).value;


    const week =
        byId(
            "annualWeekFilter"
        ).value;


    const status =
        byId(
            "annualStatusFilter"
        ).value;


    const frequency =
        byId(
            "annualFrequencyFilter"
        ).value;


    const search =
        byId(
            "annualSearch"
        )
            .value
            .trim()
            .toLowerCase();


    return annualPlanData
        .filter(
            item => {

                if (
                    line &&
                    item.line_name !==
                    line
                ) {

                    return false;
                }


                if (
                    week &&
                    Number(
                        item.planned_week
                    ) !==
                    Number(
                        week
                    )
                ) {

                    return false;
                }


                if (
                    status &&
                    item.status !==
                    status
                ) {

                    return false;
                }


                if (
                    frequency &&
                    item.frequency_text !==
                    frequency
                ) {

                    return false;
                }


                if (
                    search &&
                    !rowMatchesSearch(
                        item,
                        search
                    )
                ) {

                    return false;
                }


                return true;
            }
        );
}


// ============================================================
// ANNUAL PLAN TABLE
// ============================================================

function renderAnnualPlanTable() {

    const tbody =
        byId(
            "annualPlanTableBody"
        );


    const rows =
        getFilteredAnnualRows();


    tbody.innerHTML =
        "";


    if (
        rows.length ===
        0
    ) {

        tbody.innerHTML = `

            <tr>

                <td
                    colspan="10"
                    style="text-align:center;"
                >
                    No annual PM records found.
                </td>

            </tr>
        `;

        return;
    }


    for (
        const item
        of rows
    ) {

        const tr =
            document
                .createElement(
                    "tr"
                );


        tr.innerHTML = `

            <td>
                ${item.planned_week}
            </td>


            <td>
                ${escapeHtml(
                    item.line_name ||
                    ""
                )}
            </td>


            <td>

                <strong>
                    ${escapeHtml(
                        item.machine_code ||
                        ""
                    )}
                </strong>

                <br>

                ${escapeHtml(
                    item.machine_name ||
                    ""
                )}

            </td>


            <td>
                ${escapeHtml(
                    item.section ||
                    ""
                )}
            </td>


            <td>
                ${escapeHtml(
                    item.part_name ||
                    ""
                )}
            </td>


            <td>
                ${escapeHtml(
                    item.maintenance_task ||
                    ""
                )}
            </td>


            <td>
                ${escapeHtml(
                    item.frequency_text ||
                    ""
                )}
            </td>


            <td>

                <span class="
                    table-status
                    ${getStatusClass(
                        item.status
                    )}
                ">
                    ${escapeHtml(
                        item.status
                    )}
                </span>

            </td>


            <td>
                ${escapeHtml(
                    item.technician_name ||
                    "-"
                )}
            </td>


            <td>
                ${getActionDate(
                    item
                )}
            </td>
        `;


        tbody
            .appendChild(
                tr
            );
    }
}


// ============================================================
// REPORT FILTERS
// ============================================================

function populateReportFilters() {

    populateSelectFromData(
        "reportLineFilter",
        annualPlanData,
        "line_name",
        "All Production Lines"
    );


    populateReportMachineFilter();


    populateWeekSelect(
        "reportWeekFilter",
        true
    );
}


function populateReportMachineFilter() {

    const line =
        byId(
            "reportLineFilter"
        ).value;


    let rows =
        annualPlanData;


    if (
        line
    ) {

        rows =
            rows.filter(
                item =>
                    item.line_name ===
                    line
            );
    }


    populateMachineSelect(
        "reportMachineFilter",
        rows,
        "All Machines"
    );
}


// ============================================================
// REPORT SCOPE
//
// Line + Machine filters affect ALL adherence calculations.
// Status/Week/Search affect only detailed action table.
// ============================================================

function getReportScopeRows() {

    const line =
        byId(
            "reportLineFilter"
        ).value;


    const machine =
        byId(
            "reportMachineFilter"
        ).value;


    return annualPlanData
        .filter(
            item => {

                if (
                    line &&
                    item.line_name !==
                    line
                ) {

                    return false;
                }


                if (
                    machine &&
                    String(
                        item.machine_id
                    ) !==
                    String(
                        machine
                    )
                ) {

                    return false;
                }


                return true;
            }
        );
}


// ============================================================
// REPORTS
// ============================================================

function renderReports() {

    renderYtdSummary();

    renderFullYearSummary();

    renderWeeklyReport();

    renderMonthlyReport();

    renderLineReport();

    renderReportActions();
}


// ============================================================
// YTD SUMMARY
// ============================================================

function getYtdWeek(
    reportYear
) {

    const now =
        new Date();


    const currentYear =
        now.getFullYear();


    if (
        reportYear <
        currentYear
    ) {

        return getISOWeeksInYear(
            reportYear
        );
    }


    if (
        reportYear >
        currentYear
    ) {

        return 0;
    }


    return getISOWeekNumber(
        now
    );
}


function getYtdRows() {

    const reportYear =
        loadedAnnualYear ||
        selectedYear;


    const ytdWeek =
        getYtdWeek(
            reportYear
        );


    const scope =
        getReportScopeRows();


    return scope
        .filter(
            item =>
                Number(
                    item.planned_week
                ) <=
                ytdWeek
        );
}


function renderYtdSummary() {

    const rows =
        getYtdRows();


    const total =
        rows.length;


    const completed =
        countStatus(
            rows,
            "Completed"
        );


    const pending =
        countStatus(
            rows,
            "Pending"
        );


    const deferred =
        countStatus(
            rows,
            "Deferred"
        );


    const adherence =
        calculatePercent(
            completed,
            total
        );


    setText(
        "reportYtdPlanned",
        total
    );


    setText(
        "reportYtdCompleted",
        completed
    );


    setText(
        "reportYtdPending",
        pending
    );


    setText(
        "reportYtdDeferred",
        deferred
    );


    setAdherenceElement(
        "reportYtdAdherence",
        adherence
    );
}


// ============================================================
// FULL YEAR SUMMARY
// ============================================================

function renderFullYearSummary() {

    const rows =
        getReportScopeRows();


    const total =
        rows.length;


    const completed =
        countStatus(
            rows,
            "Completed"
        );


    const pending =
        countStatus(
            rows,
            "Pending"
        );


    const deferred =
        countStatus(
            rows,
            "Deferred"
        );


    const adherence =
        calculatePercent(
            completed,
            total
        );


    setText(
        "reportAnnualPlanned",
        total
    );


    setText(
        "reportAnnualCompleted",
        completed
    );


    setText(
        "reportAnnualPending",
        pending
    );


    setText(
        "reportAnnualDeferred",
        deferred
    );


    setAdherenceElement(
        "reportAnnualAdherence",
        adherence
    );
}


// ============================================================
// WEEKLY SUMMARY
// ============================================================

function getWeeklySummary(
    rows =
        getReportScopeRows()
) {

    const reportYear =
        loadedAnnualYear ||
        selectedYear;


    const maxWeek =
        getISOWeeksInYear(
            reportYear
        );


    const result =
        [];


    for (
        let week = 1;
        week <= maxWeek;
        week++
    ) {

        const weekRows =
            rows.filter(
                item =>
                    Number(
                        item.planned_week
                    ) ===
                    week
            );


        const planned =
            weekRows.length;


        const completed =
            countStatus(
                weekRows,
                "Completed"
            );


        const pending =
            countStatus(
                weekRows,
                "Pending"
            );


        const deferred =
            countStatus(
                weekRows,
                "Deferred"
            );


        result.push({

            week,

            planned,

            completed,

            pending,

            deferred,

            adherence:
                calculatePercent(
                    completed,
                    planned
                )
        });
    }


    return result;
}


// ============================================================
// WEEKLY REPORT
// ============================================================

function renderWeeklyReport() {

    const tbody =
        byId(
            "weeklyReportTableBody"
        );


    tbody.innerHTML =
        "";


    for (
        const row
        of getWeeklySummary()
    ) {

        const tr =
            document
                .createElement(
                    "tr"
                );


        tr.innerHTML = `

            <td>
                Week ${row.week}
            </td>

            <td>
                ${row.planned}
            </td>

            <td>
                ${row.completed}
            </td>

            <td>
                ${row.pending}
            </td>

            <td>
                ${row.deferred}
            </td>

            <td>

                <span class="
                    ${getAdherenceClass(
                        row.adherence
                    )}
                ">
                    ${row.adherence}%
                </span>

            </td>
        `;


        tbody
            .appendChild(
                tr
            );
    }
}


// ============================================================
// MONTHLY SUMMARY
// ============================================================

function getMonthlySummary(
    rows =
        getReportScopeRows()
) {

    const reportYear =
        loadedAnnualYear ||
        selectedYear;


    const result =
        [];


    for (
        let month = 1;
        month <= 12;
        month++
    ) {

        const monthRows =
            rows.filter(
                item =>
                    getPlannedMonth(
                        reportYear,
                        Number(
                            item.planned_week
                        )
                    ) ===
                    month
            );


        const planned =
            monthRows.length;


        const completed =
            countStatus(
                monthRows,
                "Completed"
            );


        const pending =
            countStatus(
                monthRows,
                "Pending"
            );


        const deferred =
            countStatus(
                monthRows,
                "Deferred"
            );


        result.push({

            month,

            monthName:
                getMonthName(
                    month
                ),

            planned,

            completed,

            pending,

            deferred,

            adherence:
                calculatePercent(
                    completed,
                    planned
                )
        });
    }


    return result;
}


// ============================================================
// MONTHLY REPORT
// ============================================================

function renderMonthlyReport() {

    const tbody =
        byId(
            "monthlyReportTableBody"
        );


    tbody.innerHTML =
        "";


    for (
        const row
        of getMonthlySummary()
    ) {

        const tr =
            document
                .createElement(
                    "tr"
                );


        tr.innerHTML = `

            <td>
                ${row.monthName}
            </td>

            <td>
                ${row.planned}
            </td>

            <td>
                ${row.completed}
            </td>

            <td>
                ${row.pending}
            </td>

            <td>
                ${row.deferred}
            </td>

            <td>

                <span class="
                    ${getAdherenceClass(
                        row.adherence
                    )}
                ">
                    ${row.adherence}%
                </span>

            </td>
        `;


        tbody
            .appendChild(
                tr
            );
    }
}


// ============================================================
// LINE SUMMARY
// ============================================================

function getLineSummary() {

    const scope =
        getReportScopeRows();


    const lines =
        [
            ...new Set(
                scope
                    .map(
                        item =>
                            item.line_name
                    )
                    .filter(
                        Boolean
                    )
            )
        ]
            .sort();


    return lines.map(
        lineName => {

            const rows =
                scope.filter(
                    item =>
                        item.line_name ===
                        lineName
                );


            const planned =
                rows.length;


            const completed =
                countStatus(
                    rows,
                    "Completed"
                );


            const pending =
                countStatus(
                    rows,
                    "Pending"
                );


            const deferred =
                countStatus(
                    rows,
                    "Deferred"
                );


            return {

                lineName,

                planned,

                completed,

                pending,

                deferred,

                adherence:
                    calculatePercent(
                        completed,
                        planned
                    )
            };
        }
    );
}


// ============================================================
// LINE REPORT
// ============================================================

function renderLineReport() {

    const tbody =
        byId(
            "reportLineTableBody"
        );


    tbody.innerHTML =
        "";


    const rows =
        getLineSummary();


    if (
        rows.length ===
        0
    ) {

        tbody.innerHTML = `

            <tr>

                <td
                    colspan="6"
                    style="text-align:center;"
                >
                    No production line data.
                </td>

            </tr>
        `;

        return;
    }


    for (
        const row
        of rows
    ) {

        const tr =
            document
                .createElement(
                    "tr"
                );


        tr.innerHTML = `

            <td>
                <strong>
                    ${escapeHtml(
                        row.lineName
                    )}
                </strong>
            </td>

            <td>
                ${row.planned}
            </td>

            <td>
                ${row.completed}
            </td>

            <td>
                ${row.pending}
            </td>

            <td>
                ${row.deferred}
            </td>

            <td>

                <span class="
                    ${getAdherenceClass(
                        row.adherence
                    )}
                ">
                    ${row.adherence}%
                </span>

            </td>
        `;


        tbody
            .appendChild(
                tr
            );
    }
}


// ============================================================
// DETAILED ACTION FILTER
// ============================================================

function getFilteredReportActions() {

    const status =
        byId(
            "reportStatusFilter"
        ).value;


    const week =
        byId(
            "reportWeekFilter"
        ).value;


    const search =
        byId(
            "reportSearch"
        )
            .value
            .trim()
            .toLowerCase();


    return getReportScopeRows()
        .filter(
            item => {

                if (
                    status &&
                    item.status !==
                    status
                ) {

                    return false;
                }


                if (
                    week &&
                    Number(
                        item.planned_week
                    ) !==
                    Number(
                        week
                    )
                ) {

                    return false;
                }


                if (
                    search &&
                    !rowMatchesSearch(
                        item,
                        search
                    )
                ) {

                    return false;
                }


                return true;
            }
        );
}


// ============================================================
// DETAILED ACTION TABLE
// ============================================================

function renderReportActions() {

    const tbody =
        byId(
            "reportActionsTableBody"
        );


    const rows =
        getFilteredReportActions();


    tbody.innerHTML =
        "";


    if (
        rows.length ===
        0
    ) {

        tbody.innerHTML = `

            <tr>

                <td
                    colspan="11"
                    style="text-align:center;"
                >
                    No PM action records found.
                </td>

            </tr>
        `;

        return;
    }


    for (
        const item
        of rows
    ) {

        const tr =
            document
                .createElement(
                    "tr"
                );


        tr.innerHTML = `

            <td>
                ${item.planned_week}
            </td>


            <td>
                ${escapeHtml(
                    item.line_name ||
                    ""
                )}
            </td>


            <td>

                <strong>
                    ${escapeHtml(
                        item.machine_code ||
                        ""
                    )}
                </strong>

                <br>

                ${escapeHtml(
                    item.machine_name ||
                    ""
                )}

            </td>


            <td>
                ${escapeHtml(
                    item.part_name ||
                    ""
                )}
            </td>


            <td>
                ${escapeHtml(
                    item.maintenance_task ||
                    ""
                )}
            </td>


            <td>
                ${escapeHtml(
                    item.frequency_text ||
                    ""
                )}
            </td>


            <td>

                <span class="
                    table-status
                    ${getStatusClass(
                        item.status
                    )}
                ">
                    ${escapeHtml(
                        item.status ||
                        ""
                    )}
                </span>

            </td>


            <td>
                ${escapeHtml(
                    item.technician_name ||
                    "-"
                )}
            </td>


            <td>
                ${getActionDate(
                    item
                )}
            </td>


            <td>
                ${escapeHtml(
                    item.deferred_reason ||
                    "-"
                )}
            </td>


            <td>
                ${escapeHtml(
                    item.notes ||
                    "-"
                )}
            </td>
        `;


        tbody
            .appendChild(
                tr
            );
    }
}


// ============================================================
// CALENDAR FILTERS
// ============================================================

function populateCalendarFilters() {

    populateSelectFromData(
        "calendarLineFilter",
        annualPlanData,
        "line_name",
        "All Production Lines"
    );


    populateSelectFromData(
        "calendarFrequencyFilter",
        annualPlanData,
        "frequency_text",
        "All Frequencies"
    );


    populateCalendarMachineFilter();
}


function populateCalendarMachineFilter() {

    const line =
        byId(
            "calendarLineFilter"
        ).value;


    let rows =
        annualPlanData;


    if (
        line
    ) {

        rows =
            rows.filter(
                item =>
                    item.line_name ===
                    line
            );
    }


    populateMachineSelect(
        "calendarMachineFilter",
        rows,
        "All Machines"
    );
}


// ============================================================
// WEEK SELECTORS
// ============================================================

function populateWeekSelectors() {

    populateWeekSelect(
        "calendarWeekFrom",
        false,
        1
    );


    populateWeekSelect(
        "calendarWeekTo",
        false,
        52
    );
}


function populateWeekSelect(
    elementId,
    includeAll = true,
    defaultValue = null
) {

    const select =
        byId(
            elementId
        );


    if (
        !select
    ) {

        return;
    }


    const oldValue =
        select.value;


    select.innerHTML =
        "";


    if (
        includeAll
    ) {

        const allOption =
            document
                .createElement(
                    "option"
                );


        allOption.value =
            "";


        allOption.textContent =
            "All Weeks";


        select
            .appendChild(
                allOption
            );
    }


    for (
        let week = 1;
        week <= 53;
        week++
    ) {

        const option =
            document
                .createElement(
                    "option"
                );


        option.value =
            String(
                week
            );


        option.textContent =
            `Week ${week}`;


        select
            .appendChild(
                option
            );
    }


    if (
        oldValue &&
        [
            ...select.options
        ]
            .some(
                option =>
                    option.value ===
                    oldValue
            )
    ) {

        select.value =
            oldValue;

    } else if (
        defaultValue !==
        null
    ) {

        select.value =
            String(
                defaultValue
            );
    }
}


// ============================================================
// MONTH → WEEK RANGE
// ============================================================

function applyCalendarMonthToWeeks() {

    const month =
        Number(
            byId(
                "calendarMonthFilter"
            ).value
        );


    if (
        !month
    ) {

        return;
    }


    const year =
        Number(
            byId(
                "calendarYearFilter"
            ).value
        );


    const weeks =
        getWeeksForMonth(
            year,
            month
        );


    if (
        weeks.length ===
        0
    ) {

        return;
    }


    byId(
        "calendarWeekFrom"
    ).value =
        String(
            Math.min(
                ...weeks
            )
        );


    byId(
        "calendarWeekTo"
    ).value =
        String(
            Math.max(
                ...weeks
            )
        );
}


// ============================================================
// RESET CALENDAR
// ============================================================

function resetCalendarFilters(
    render = true
) {

    byId(
        "calendarLineFilter"
    ).value =
        "";


    byId(
        "calendarMachineFilter"
    ).value =
        "";


    byId(
        "calendarFrequencyFilter"
    ).value =
        "";


    byId(
        "calendarMonthFilter"
    ).value =
        "";


    byId(
        "calendarWeekFrom"
    ).value =
        "1";


    byId(
        "calendarWeekTo"
    ).value =
        "52";


    byId(
        "calendarSearch"
    ).value =
        "";


    populateCalendarMachineFilter();


    if (
        render
    ) {

        renderCalendarMatrix();
    }
}


// ============================================================
// CALENDAR FILTERED DATA
// ============================================================

function getCalendarFilterState() {

    let from =
        Number(
            byId(
                "calendarWeekFrom"
            ).value
        ) ||
        1;


    let to =
        Number(
            byId(
                "calendarWeekTo"
            ).value
        ) ||
        52;


    if (
        from >
        to
    ) {

        const temporary =
            from;

        from =
            to;

        to =
            temporary;
    }


    return {

        line:
            byId(
                "calendarLineFilter"
            ).value,

        machine:
            byId(
                "calendarMachineFilter"
            ).value,

        frequency:
            byId(
                "calendarFrequencyFilter"
            ).value,

        month:
            Number(
                byId(
                    "calendarMonthFilter"
                ).value
            ) ||
            null,

        from,

        to,

        search:
            byId(
                "calendarSearch"
            )
                .value
                .trim()
                .toLowerCase()
    };
}


function getFilteredCalendarRows() {

    const filter =
        getCalendarFilterState();


    const year =
        loadedAnnualYear ||
        Number(
            byId(
                "calendarYearFilter"
            ).value
        );


    return annualPlanData
        .filter(
            item => {

                const week =
                    Number(
                        item.planned_week
                    );


                if (
                    week <
                    filter.from ||
                    week >
                    filter.to
                ) {

                    return false;
                }


                if (
                    filter.month &&
                    getPlannedMonth(
                        year,
                        week
                    ) !==
                    filter.month
                ) {

                    return false;
                }


                if (
                    filter.line &&
                    item.line_name !==
                    filter.line
                ) {

                    return false;
                }


                if (
                    filter.machine &&
                    String(
                        item.machine_id
                    ) !==
                    String(
                        filter.machine
                    )
                ) {

                    return false;
                }


                if (
                    filter.frequency &&
                    item.frequency_text !==
                    filter.frequency
                ) {

                    return false;
                }


                if (
                    filter.search &&
                    !rowMatchesSearch(
                        item,
                        filter.search
                    )
                ) {

                    return false;
                }


                return true;
            }
        );
}


// ============================================================
// CALENDAR WEEKS TO DISPLAY
// ============================================================

function getCalendarWeeks() {

    const filter =
        getCalendarFilterState();


    const year =
        loadedAnnualYear ||
        Number(
            byId(
                "calendarYearFilter"
            ).value
        );


    const weeks =
        [];


    for (
        let week =
            filter.from;

        week <=
        filter.to;

        week++
    ) {

        if (
            filter.month &&
            getPlannedMonth(
                year,
                week
            ) !==
            filter.month
        ) {

            continue;
        }


        weeks.push(
            week
        );
    }


    return weeks;
}


// ============================================================
// CALENDAR GROUPING
// ============================================================

function groupCalendarTasks(
    rows
) {

    const map =
        new Map();


    for (
        const item
        of rows
    ) {

        const key =
            String(
                item.pm_task_id
            );


        if (
            !map.has(
                key
            )
        ) {

            map.set(
                key,
                {

                    pmTaskId:
                        item.pm_task_id,

                    lineName:
                        item.line_name,

                    machineId:
                        item.machine_id,

                    machineCode:
                        item.machine_code,

                    machineName:
                        item.machine_name,

                    partName:
                        item.part_name,

                    maintenanceTask:
                        item.maintenance_task,

                    frequency:
                        item.frequency_text,

                    weeks:
                        new Map()
                }
            );
        }


        const group =
            map.get(
                key
            );


        group.weeks.set(
            Number(
                item.planned_week
            ),
            item
        );
    }


    return [
        ...map.values()
    ]
        .sort(
            (
                a,
                b
            ) => {

                return (

                    String(
                        a.lineName
                    )
                        .localeCompare(
                            String(
                                b.lineName
                            )
                        ) ||

                    String(
                        a.machineCode
                    )
                        .localeCompare(
                            String(
                                b.machineCode
                            )
                        ) ||

                    String(
                        a.partName
                    )
                        .localeCompare(
                            String(
                                b.partName
                            )
                        ) ||

                    String(
                        a.maintenanceTask
                    )
                        .localeCompare(
                            String(
                                b.maintenanceTask
                            )
                        )
                );
            }
        );
}


// ============================================================
// RENDER CALENDAR MATRIX
// ============================================================

function renderCalendarMatrix() {

    const rows =
        getFilteredCalendarRows();


    const weeks =
        getCalendarWeeks();


    renderCalendarKpis(
        rows,
        weeks
    );


    renderCalendarHeader(
        weeks
    );


    renderCalendarBody(
        rows,
        weeks
    );
}


// ============================================================
// CALENDAR KPI
// ============================================================

function renderCalendarKpis(
    rows,
    weeks
) {

    const planned =
        rows.length;


    const completed =
        countStatus(
            rows,
            "Completed"
        );


    const pending =
        countStatus(
            rows,
            "Pending"
        );


    const deferred =
        countStatus(
            rows,
            "Deferred"
        );


    const adherence =
        calculatePercent(
            completed,
            planned
        );


    setText(
        "calendarPlanned",
        planned
    );


    setText(
        "calendarCompleted",
        completed
    );


    setText(
        "calendarPending",
        pending
    );


    setText(
        "calendarDeferred",
        deferred
    );


    setAdherenceElement(
        "calendarAdherence",
        adherence
    );


    let rangeText =
        "No weeks selected";


    if (
        weeks.length >
        0
    ) {

        rangeText =
            weeks.length ===
            1
                ?
                `Week ${weeks[0]}`
                :
                `Week ${weeks[0]} - Week ${weeks[weeks.length - 1]}`;


        const month =
            byId(
                "calendarMonthFilter"
            ).value;


        if (
            month
        ) {

            rangeText =
                `${getMonthName(
                    Number(
                        month
                    )
                )} | ${rangeText}`;
        }
    }


    setText(
        "calendarRangeLabel",
        rangeText
    );
}


// ============================================================
// CALENDAR HEADER
// ============================================================

function renderCalendarHeader(
    weeks
) {

    const head =
        byId(
            "calendarMatrixHead"
        );


    let html = `

        <tr>

            <th>
                Line
            </th>

            <th>
                Machine
            </th>

            <th>
                Part
            </th>

            <th>
                PM Task
            </th>

            <th>
                Frequency
            </th>
    `;


    for (
        const week
        of weeks
    ) {

        html += `

            <th class="calendar-week-header">
                W${week}
            </th>
        `;
    }


    html += `
        </tr>
    `;


    head.innerHTML =
        html;
}


// ============================================================
// CALENDAR BODY
// ============================================================

function renderCalendarBody(
    rows,
    weeks
) {

    const body =
        byId(
            "calendarMatrixBody"
        );


    body.innerHTML =
        "";


    const groups =
        groupCalendarTasks(
            rows
        );


    if (
        groups.length ===
        0
    ) {

        body.innerHTML = `

            <tr>

                <td
                    colspan="${5 + weeks.length}"
                    style="text-align:center;"
                >
                    No PM records found for selected filters.
                </td>

            </tr>
        `;

        return;
    }


    for (
        const group
        of groups
    ) {

        const tr =
            document
                .createElement(
                    "tr"
                );


        let html = `

            <td>

                <span class="calendar-line-name">
                    ${escapeHtml(
                        group.lineName ||
                        ""
                    )}
                </span>

            </td>


            <td>

                <div class="calendar-machine-name">
                    ${escapeHtml(
                        group.machineName ||
                        ""
                    )}
                </div>

                <div class="calendar-machine-code">
                    ${escapeHtml(
                        group.machineCode ||
                        ""
                    )}
                </div>

            </td>


            <td>

                <span class="calendar-part-name">
                    ${escapeHtml(
                        group.partName ||
                        ""
                    )}
                </span>

            </td>


            <td>

                <span class="calendar-task-text">
                    ${escapeHtml(
                        group.maintenanceTask ||
                        ""
                    )}
                </span>

            </td>


            <td class="calendar-frequency-text">
                ${escapeHtml(
                    group.frequency ||
                    ""
                )}
            </td>
        `;


        for (
            const week
            of weeks
        ) {

            const item =
                group.weeks.get(
                    week
                );


            if (
                !item
            ) {

                html += `

                    <td class="calendar-week-cell">

                        <span class="calendar-status-notplanned">
                            -
                        </span>

                    </td>
                `;

                continue;
            }


            const cell =
                getCalendarStatusDisplay(
                    item
                );


            html += `

                <td class="calendar-week-cell">

                    <span
                        class="${cell.className}"
                        title="${escapeHtml(
                            cell.title
                        )}"
                    >
                        ${cell.label}
                    </span>

                </td>
            `;
        }


        tr.innerHTML =
            html;


        body
            .appendChild(
                tr
            );
    }
}


// ============================================================
// CALENDAR CELL STATUS
// ============================================================

function getCalendarStatusDisplay(
    item
) {

    if (
        item.status ===
        "Completed"
    ) {

        return {

            className:
                "calendar-status-completed",

            label:
                "✓",

            title:
                `Completed${item.technician_name ? " - " + item.technician_name : ""}${item.completed_at ? " - " + formatDate(item.completed_at) : ""}`
        };
    }


    if (
        item.status ===
        "Deferred"
    ) {

        return {

            className:
                "calendar-status-deferred",

            label:
                "D",

            title:
                `Deferred${item.deferred_reason ? " - " + item.deferred_reason : ""}`
        };
    }


    if (
        item.status ===
        "Overdue"
    ) {

        return {

            className:
                "calendar-status-overdue",

            label:
                "O",

            title:
                "Overdue"
        };
    }


    return {

        className:
            "calendar-status-pending",

        label:
            "P",

        title:
            "Pending"
    };
}


// ============================================================
// CALENDAR EXCEL EXPORT
// ============================================================

function exportCalendarExcel() {

    const rows =
        getFilteredCalendarRows();


    const weeks =
        getCalendarWeeks();


    const groups =
        groupCalendarTasks(
            rows
        );


    if (
        groups.length ===
        0
    ) {

        alert(
            "No PM Calendar data to export."
        );

        return;
    }


    const headers = [

        "Production Line",

        "Machine Code",

        "Machine Name",

        "Part",

        "PM Task",

        "Frequency",

        ...weeks.map(
            week =>
                `W${week}`
        )
    ];


    const matrixRows =
        groups.map(
            group => {

                const row = [

                    group.lineName,

                    group.machineCode,

                    group.machineName,

                    group.partName,

                    group.maintenanceTask,

                    group.frequency
                ];


                for (
                    const week
                    of weeks
                ) {

                    const item =
                        group.weeks.get(
                            week
                        );


                    row.push(
                        item
                            ?
                            item.status
                            :
                            ""
                    );
                }


                return row;
            }
        );


    const summary = [

        [
            "Metric",
            "Value"
        ],

        [
            "Year",
            loadedAnnualYear
        ],

        [
            "Week From",
            weeks[0] ||
            ""
        ],

        [
            "Week To",
            weeks[
                weeks.length -
                1
            ] ||
            ""
        ],

        [
            "Planned",
            rows.length
        ],

        [
            "Completed",
            countStatus(
                rows,
                "Completed"
            )
        ],

        [
            "Pending",
            countStatus(
                rows,
                "Pending"
            )
        ],

        [
            "Deferred",
            countStatus(
                rows,
                "Deferred"
            )
        ],

        [
            "Adherence %",
            calculatePercent(
                countStatus(
                    rows,
                    "Completed"
                ),
                rows.length
            )
        ]
    ];


    downloadExcelWorkbook(
        `PM_Calendar_${loadedAnnualYear}.xls`,
        [

            {

                name:
                    "Calendar Summary",

                headers:
                    summary[0],

                rows:
                    summary.slice(
                        1
                    )
            },

            {

                name:
                    "PM Calendar",

                headers,

                rows:
                    matrixRows
            }
        ]
    );
}


// ============================================================
// CALENDAR PDF EXPORT
// ============================================================

function exportCalendarPdf() {

    const rows =
        getFilteredCalendarRows();


    const weeks =
        getCalendarWeeks();


    const groups =
        groupCalendarTasks(
            rows
        );


    if (
        groups.length ===
        0
    ) {

        alert(
            "No PM Calendar data to export."
        );

        return;
    }


    let table = `

        <table>

            <thead>

                <tr>

                    <th>
                        Line
                    </th>

                    <th>
                        Machine
                    </th>

                    <th>
                        Part
                    </th>

                    <th>
                        PM Task
                    </th>

                    <th>
                        Frequency
                    </th>
    `;


    for (
        const week
        of weeks
    ) {

        table += `

            <th>
                W${week}
            </th>
        `;
    }


    table += `

                </tr>

            </thead>

            <tbody>
    `;


    const maxRows =
        250;


    for (
        const group
        of groups.slice(
            0,
            maxRows
        )
    ) {

        table += `

            <tr>

                <td>
                    ${escapeHtml(
                        group.lineName
                    )}
                </td>

                <td>
                    ${escapeHtml(
                        group.machineCode
                    )}
                </td>

                <td>
                    ${escapeHtml(
                        group.partName
                    )}
                </td>

                <td>
                    ${escapeHtml(
                        group.maintenanceTask
                    )}
                </td>

                <td>
                    ${escapeHtml(
                        group.frequency
                    )}
                </td>
        `;


        for (
            const week
            of weeks
        ) {

            const item =
                group.weeks.get(
                    week
                );


            table += `

                <td>
                    ${
                        item
                            ?
                            escapeHtml(
                                item.status
                            )
                            :
                            ""
                    }
                </td>
            `;
        }


        table += `
            </tr>
        `;
    }


    table += `

            </tbody>

        </table>
    `;


    if (
        groups.length >
        maxRows
    ) {

        table += `

            <p>
                PDF limited to first ${maxRows} PM tasks.
                Use Excel for complete matrix.
            </p>
        `;
    }


    const summary =
        createPrintSummary(
            rows
        );


    openPrintReport(
        `PM Calendar Matrix - ${loadedAnnualYear}`,
        summary +
        table,
        "landscape"
    );
}


// ============================================================
// EXCEL IMPORT - FILE SELECTED
// ============================================================

function handleExcelFileSelected() {

    const file =
        byId(
            "pmExcelFile"
        ).files[
            0
        ];


    importPreviewToken =
        null;


    byId(
        "confirmExcelImport"
    ).disabled =
        true;


    byId(
        "importSummarySection"
    )
        .classList
        .add(
            "hidden"
        );


    byId(
        "importPreviewSection"
    )
        .classList
        .add(
            "hidden"
        );


    if (
        !file
    ) {

        setText(
            "selectedExcelFileName",
            "No file selected"
        );

        return;
    }


    setText(
        "selectedExcelFileName",
        file.name
    );


    setImportMessage(
        "File selected. Click Preview & Validate.",
        "warning"
    );
}


// ============================================================
// EXCEL IMPORT PREVIEW
//
// Requires server endpoint:
// POST /api/import/preview
//
// Next server.js step will add this endpoint.
// ============================================================

async function previewExcelImport() {

    const file =
        byId(
            "pmExcelFile"
        ).files[
            0
        ];


    if (
        !file
    ) {

        setImportMessage(
            "Please select an Excel file first.",
            "error"
        );

        return;
    }


    const button =
        byId(
            "previewExcelImport"
        );


    const oldText =
        button.textContent;


    button.disabled =
        true;


    button.textContent =
        "Reading Excel...";


    try {

        const formData =
            new FormData();


        formData.append(
            "file",
            file
        );


        const response =
            await fetch(
                "/api/import/preview",
                {

                    method:
                        "POST",

                    body:
                        formData
                }
            );


        const data =
            await response.json();


        if (
            !response.ok
        ) {

            throw new Error(
                data.error ||
                "Excel preview failed."
            );
        }


        importPreviewToken =
            data.preview_token ||
            null;


        renderImportSummary(
            data.summary ||
            {}
        );


        renderImportPreview(
            data.preview ||
            []
        );


        byId(
            "confirmExcelImport"
        ).disabled =
            !importPreviewToken;


        setImportMessage(
            "Validation completed. Review the changes before importing.",
            "success"
        );


    } catch (
        error
    ) {

        console.error(
            error
        );


        setImportMessage(
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
// IMPORT SUMMARY
// ============================================================

function renderImportSummary(
    summary
) {

    byId(
        "importSummarySection"
    )
        .classList
        .remove(
            "hidden"
        );


    setText(
        "importTotalRows",
        summary.total_rows ||
        0
    );


    setText(
        "importNewTasks",
        summary.new_tasks ||
        0
    );


    setText(
        "importUpdatedTasks",
        summary.updated_tasks ||
        0
    );


    setText(
        "importUnchangedTasks",
        summary.unchanged_tasks ||
        0
    );


    setText(
        "importInactiveTasks",
        summary.inactive_tasks ||
        0
    );


    setText(
        "importErrorCount",
        summary.errors ||
        0
    );
}


// ============================================================
// IMPORT PREVIEW TABLE
// ============================================================

function renderImportPreview(
    rows
) {

    const section =
        byId(
            "importPreviewSection"
        );


    const body =
        byId(
            "importPreviewTableBody"
        );


    section
        .classList
        .remove(
            "hidden"
        );


    body.innerHTML =
        "";


    if (
        rows.length ===
        0
    ) {

        body.innerHTML = `

            <tr>

                <td colspan="8">
                    No changes detected.
                </td>

            </tr>
        `;

        return;
    }


    for (
        const item
        of rows
    ) {

        const tr =
            document
                .createElement(
                    "tr"
                );


        tr.innerHTML = `

            <td>

                <span class="${getImportActionClass(
                    item.action
                )}">
                    ${escapeHtml(
                        item.action ||
                        ""
                    )}
                </span>

            </td>


            <td>
                ${escapeHtml(
                    item.line_name ||
                    ""
                )}
            </td>


            <td>
                ${escapeHtml(
                    item.machine_code ||
                    item.machine_name ||
                    ""
                )}
            </td>


            <td>
                ${escapeHtml(
                    item.section ||
                    ""
                )}
            </td>


            <td>
                ${escapeHtml(
                    item.part_name ||
                    ""
                )}
            </td>


            <td>
                ${escapeHtml(
                    item.maintenance_task ||
                    ""
                )}
            </td>


            <td>
                ${escapeHtml(
                    item.frequency_text ||
                    ""
                )}
            </td>


            <td>
                ${escapeHtml(
                    item.validation ||
                    "OK"
                )}
            </td>
        `;


        body
            .appendChild(
                tr
            );
    }
}


// ============================================================
// CONFIRM EXCEL IMPORT
//
// Requires server endpoint:
// POST /api/import/confirm
// ============================================================

async function confirmExcelImport() {

    if (
        !importPreviewToken
    ) {

        setImportMessage(
            "Preview the Excel file before importing.",
            "error"
        );

        return;
    }


    const confirmed =
        window.confirm(
            "Import this PM master update?\n\nHistorical Completed and Deferred PM records will be preserved."
        );


    if (
        !confirmed
    ) {

        return;
    }


    const button =
        byId(
            "confirmExcelImport"
        );


    const oldText =
        button.textContent;


    button.disabled =
        true;


    button.textContent =
        "Updating PM Master...";


    try {

        const response =
            await fetch(
                "/api/import/confirm",
                {

                    method:
                        "POST",

                    headers: {

                        "Content-Type":
                            "application/json"
                    },

                    body:
                        JSON.stringify({

                            preview_token:
                                importPreviewToken
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
                "PM master update failed."
            );
        }


        renderImportSummary(
            data.summary ||
            {}
        );


        setImportMessage(
            "PM Master updated successfully. Historical PM records were preserved.",
            "success"
        );


        importPreviewToken =
            null;


        loadedAnnualYear =
            null;


        await loadSchedule();


    } catch (
        error
    ) {

        console.error(
            error
        );


        setImportMessage(
            error.message,
            "error"
        );


    } finally {

        button.textContent =
            oldText;


        button.disabled =
            true;
    }
}


// ============================================================
// IMPORT ACTION CLASS
// ============================================================

function getImportActionClass(
    action
) {

    const value =
        String(
            action ||
            ""
        )
            .toLowerCase();


    if (
        value ===
        "new"
    ) {

        return "import-action-new";
    }


    if (
        value ===
        "update" ||
        value ===
        "updated"
    ) {

        return "import-action-update";
    }


    if (
        value ===
        "inactive"
    ) {

        return "import-action-inactive";
    }


    if (
        value ===
        "error"
    ) {

        return "import-action-error";
    }


    return "import-action-unchanged";
}


// ============================================================
// IMPORT MESSAGE
// ============================================================

function setImportMessage(
    message,
    type
) {

    const element =
        byId(
            "importMessage"
        );


    if (
        !element
    ) {

        return;
    }


    element.textContent =
        message ||
        "";


    element
        .classList
        .remove(
            "success",
            "error",
            "warning"
        );


    if (
        type
    ) {

        element
            .classList
            .add(
                type
            );
    }
}


// ============================================================
// EXPORT ANNUAL EXCEL
// ============================================================

function exportAnnualExcel() {

    if (
        annualPlanData.length ===
        0
    ) {

        alert(
            "No annual plan data."
        );

        return;
    }


    const year =
        loadedAnnualYear ||
        selectedYear;


    const fullYear =
        annualPlanData;


    const weekly =
        getWeeklySummary(
            fullYear
        );


    const monthly =
        getMonthlySummary(
            fullYear
        );


    const sheets = [

        {

            name:
                "Annual Summary",

            headers: [
                "Metric",
                "Value"
            ],

            rows: [

                [
                    "Year",
                    year
                ],

                [
                    "Planned",
                    fullYear.length
                ],

                [
                    "Completed",
                    countStatus(
                        fullYear,
                        "Completed"
                    )
                ],

                [
                    "Pending",
                    countStatus(
                        fullYear,
                        "Pending"
                    )
                ],

                [
                    "Deferred",
                    countStatus(
                        fullYear,
                        "Deferred"
                    )
                ],

                [
                    "Adherence %",
                    calculatePercent(
                        countStatus(
                            fullYear,
                            "Completed"
                        ),
                        fullYear.length
                    )
                ]
            ]
        },


        {

            name:
                "Weekly Summary",

            headers: [
                "Week",
                "Planned",
                "Completed",
                "Pending",
                "Deferred",
                "Adherence %"
            ],

            rows:
                weekly.map(
                    row => [

                        row.week,

                        row.planned,

                        row.completed,

                        row.pending,

                        row.deferred,

                        row.adherence
                    ]
                )
        },


        {

            name:
                "Monthly Summary",

            headers: [
                "Month",
                "Planned",
                "Completed",
                "Pending",
                "Deferred",
                "Adherence %"
            ],

            rows:
                monthly.map(
                    row => [

                        row.monthName,

                        row.planned,

                        row.completed,

                        row.pending,

                        row.deferred,

                        row.adherence
                    ]
                )
        },


        {

            name:
                "Annual Plan",

            headers:
                getActionHeaders(),

            rows:
                getFilteredAnnualRows()
                    .map(
                        actionRowToArray
                    )
        }
    ];


    downloadExcelWorkbook(
        `PM_Annual_Plan_${year}.xls`,
        sheets
    );
}


// ============================================================
// EXPORT REPORT EXCEL
// ============================================================

function exportReportExcel() {

    const scope =
        getReportScopeRows();


    if (
        scope.length ===
        0
    ) {

        alert(
            "No report data."
        );

        return;
    }


    const year =
        loadedAnnualYear ||
        selectedYear;


    const ytd =
        getYtdRows();


    const sheets = [

        {

            name:
                "YTD Summary",

            headers: [
                "Metric",
                "Value"
            ],

            rows: [

                [
                    "YTD Planned",
                    ytd.length
                ],

                [
                    "YTD Completed",
                    countStatus(
                        ytd,
                        "Completed"
                    )
                ],

                [
                    "YTD Pending",
                    countStatus(
                        ytd,
                        "Pending"
                    )
                ],

                [
                    "YTD Deferred",
                    countStatus(
                        ytd,
                        "Deferred"
                    )
                ],

                [
                    "YTD Adherence %",
                    calculatePercent(
                        countStatus(
                            ytd,
                            "Completed"
                        ),
                        ytd.length
                    )
                ]
            ]
        },


        {

            name:
                "Weekly Adherence",

            headers: [
                "Week",
                "Planned",
                "Completed",
                "Pending",
                "Deferred",
                "Adherence %"
            ],

            rows:
                getWeeklySummary()
                    .map(
                        row => [

                            row.week,

                            row.planned,

                            row.completed,

                            row.pending,

                            row.deferred,

                            row.adherence
                        ]
                    )
        },


        {

            name:
                "Monthly Adherence",

            headers: [
                "Month",
                "Planned",
                "Completed",
                "Pending",
                "Deferred",
                "Adherence %"
            ],

            rows:
                getMonthlySummary()
                    .map(
                        row => [

                            row.monthName,

                            row.planned,

                            row.completed,

                            row.pending,

                            row.deferred,

                            row.adherence
                        ]
                    )
        },


        {

            name:
                "Line Performance",

            headers: [
                "Production Line",
                "Planned",
                "Completed",
                "Pending",
                "Deferred",
                "Adherence %"
            ],

            rows:
                getLineSummary()
                    .map(
                        row => [

                            row.lineName,

                            row.planned,

                            row.completed,

                            row.pending,

                            row.deferred,

                            row.adherence
                        ]
                    )
        },


        {

            name:
                "PM Actions",

            headers:
                getActionHeaders(),

            rows:
                getFilteredReportActions()
                    .map(
                        actionRowToArray
                    )
        }
    ];


    downloadExcelWorkbook(
        `PM_Performance_Report_${year}.xls`,
        sheets
    );
}


// ============================================================
// ANNUAL PDF
// ============================================================

function exportAnnualPdf() {

    const rows =
        getFilteredAnnualRows();


    if (
        rows.length ===
        0
    ) {

        alert(
            "No annual plan rows to export."
        );

        return;
    }


    openPrintReport(
        `Annual Preventive Maintenance Plan - ${loadedAnnualYear}`,
        createPrintSummary(
            rows
        ) +
        createPrintableActionTable(
            rows
        ),
        "landscape"
    );
}


// ============================================================
// REPORT PDF
// ============================================================

function exportReportPdf() {

    const scope =
        getReportScopeRows();


    if (
        scope.length ===
        0
    ) {

        alert(
            "No report data."
        );

        return;
    }


    const ytd =
        getYtdRows();


    const html = `

        <h2>
            YTD Performance
        </h2>

        ${createPrintSummary(
            ytd
        )}


        <h2>
            Full-Year Performance
        </h2>

        ${createPrintSummary(
            scope
        )}


        <h2>
            Weekly PM Adherence
        </h2>

        ${createSummaryPrintTable(
            getWeeklySummary(),
            "week"
        )}


        <h2>
            Monthly PM Adherence
        </h2>

        ${createSummaryPrintTable(
            getMonthlySummary(),
            "month"
        )}


        <h2>
            Detailed PM Actions
        </h2>

        ${createPrintableActionTable(
            getFilteredReportActions()
        )}
    `;


    openPrintReport(
        `PM Performance Report - ${loadedAnnualYear}`,
        html,
        "landscape"
    );
}


// ============================================================
// PRINT SUMMARY
// ============================================================

function createPrintSummary(
    rows
) {

    const total =
        rows.length;


    const completed =
        countStatus(
            rows,
            "Completed"
        );


    const pending =
        countStatus(
            rows,
            "Pending"
        );


    const deferred =
        countStatus(
            rows,
            "Deferred"
        );


    const adherence =
        calculatePercent(
            completed,
            total
        );


    return `

        <div class="summary">

            <div>

                <b>
                    Planned
                </b>

                <span>
                    ${total}
                </span>

            </div>


            <div>

                <b>
                    Completed
                </b>

                <span>
                    ${completed}
                </span>

            </div>


            <div>

                <b>
                    Pending
                </b>

                <span>
                    ${pending}
                </span>

            </div>


            <div>

                <b>
                    Deferred
                </b>

                <span>
                    ${deferred}
                </span>

            </div>


            <div>

                <b>
                    Adherence
                </b>

                <span>
                    ${adherence}%
                </span>

            </div>

        </div>
    `;
}


// ============================================================
// PRINT ACTION TABLE
// ============================================================

function createPrintableActionTable(
    rows
) {

    const maxRows =
        1000;


    const displayRows =
        rows.slice(
            0,
            maxRows
        );


    let html = `

        <table>

            <thead>

                <tr>

                    <th>Week</th>
                    <th>Line</th>
                    <th>Machine</th>
                    <th>Part</th>
                    <th>PM Task</th>
                    <th>Frequency</th>
                    <th>Status</th>
                    <th>Technician</th>
                    <th>Action Date</th>
                    <th>Deferred Reason</th>
                    <th>Notes</th>

                </tr>

            </thead>

            <tbody>
    `;


    for (
        const item
        of displayRows
    ) {

        html += `

            <tr>

                <td>
                    ${item.planned_week}
                </td>

                <td>
                    ${escapeHtml(
                        item.line_name ||
                        ""
                    )}
                </td>

                <td>
                    ${escapeHtml(
                        item.machine_code ||
                        ""
                    )}
                </td>

                <td>
                    ${escapeHtml(
                        item.part_name ||
                        ""
                    )}
                </td>

                <td>
                    ${escapeHtml(
                        item.maintenance_task ||
                        ""
                    )}
                </td>

                <td>
                    ${escapeHtml(
                        item.frequency_text ||
                        ""
                    )}
                </td>

                <td>
                    ${escapeHtml(
                        item.status ||
                        ""
                    )}
                </td>

                <td>
                    ${escapeHtml(
                        item.technician_name ||
                        "-"
                    )}
                </td>

                <td>
                    ${getActionDate(
                        item
                    )}
                </td>

                <td>
                    ${escapeHtml(
                        item.deferred_reason ||
                        "-"
                    )}
                </td>

                <td>
                    ${escapeHtml(
                        item.notes ||
                        "-"
                    )}
                </td>

            </tr>
        `;
    }


    html += `

            </tbody>

        </table>
    `;


    if (
        rows.length >
        maxRows
    ) {

        html += `

            <p>
                PDF limited to first ${maxRows} rows.
                Export Excel for all ${rows.length} records.
            </p>
        `;
    }


    return html;
}


// ============================================================
// PRINT SUMMARY TABLE
// ============================================================

function createSummaryPrintTable(
    rows,
    type
) {

    let html = `

        <table>

            <thead>

                <tr>

                    <th>
                        ${
                            type ===
                            "week"
                                ?
                                "Week"
                                :
                                "Month"
                        }
                    </th>

                    <th>Planned</th>
                    <th>Completed</th>
                    <th>Pending</th>
                    <th>Deferred</th>
                    <th>Adherence %</th>

                </tr>

            </thead>

            <tbody>
    `;


    for (
        const row
        of rows
    ) {

        html += `

            <tr>

                <td>

                    ${
                        type ===
                        "week"
                            ?
                            `Week ${row.week}`
                            :
                            escapeHtml(
                                row.monthName
                            )
                    }

                </td>

                <td>
                    ${row.planned}
                </td>

                <td>
                    ${row.completed}
                </td>

                <td>
                    ${row.pending}
                </td>

                <td>
                    ${row.deferred}
                </td>

                <td>
                    ${row.adherence}%
                </td>

            </tr>
        `;
    }


    html += `

            </tbody>

        </table>
    `;


    return html;
}


// ============================================================
// PRINT WINDOW
// ============================================================

function openPrintReport(
    title,
    bodyHtml,
    orientation =
        "landscape"
) {

    const printWindow =
        window.open(
            "",
            "_blank"
        );


    if (
        !printWindow
    ) {

        alert(
            "Please allow pop-ups to export PDF."
        );

        return;
    }


    printWindow.document.write(`

        <!DOCTYPE html>

        <html>

        <head>

            <title>
                ${escapeHtml(
                    title
                )}
            </title>

            <style>

                @page {

                    size:
                        A4 ${orientation};

                    margin:
                        10mm;
                }

                body {

                    font-family:
                        Arial,
                        sans-serif;

                    color:
                        #111827;

                    font-size:
                        9px;
                }

                h1 {

                    color:
                        #173c8f;

                    font-size:
                        20px;

                    margin:
                        0 0 4px;
                }

                h2 {

                    color:
                        #173c8f;

                    margin-top:
                        22px;

                    font-size:
                        14px;
                }

                .company {

                    color:
                        #173c8f;

                    font-size:
                        13px;

                    font-weight:
                        bold;
                }

                .generated {

                    color:
                        #64748b;

                    margin:
                        5px 0 15px;
                }

                .summary {

                    display:
                        flex;

                    flex-wrap:
                        wrap;

                    gap:
                        8px;

                    margin:
                        12px 0;
                }

                .summary div {

                    border:
                        1px solid #d1d5db;

                    border-radius:
                        6px;

                    padding:
                        7px 10px;

                    min-width:
                        100px;
                }

                .summary b {

                    display:
                        block;

                    color:
                        #475569;
                }

                .summary span {

                    display:
                        block;

                    margin-top:
                        4px;

                    font-size:
                        14px;

                    font-weight:
                        bold;
                }

                table {

                    width:
                        100%;

                    border-collapse:
                        collapse;

                    margin-bottom:
                        14px;
                }

                th {

                    background:
                        #173c8f;

                    color:
                        white;

                    padding:
                        4px;

                    border:
                        1px solid #cbd5e1;

                    text-align:
                        left;
                }

                td {

                    padding:
                        4px;

                    border:
                        1px solid #d1d5db;

                    vertical-align:
                        top;
                }

                tr {

                    page-break-inside:
                        avoid;
                }

            </style>

        </head>


        <body>

            <div class="company">
                National Food Company - Americana Cake
            </div>


            <h1>
                ${escapeHtml(
                    title
                )}
            </h1>


            <div class="generated">
                Generated:
                ${escapeHtml(
                    new Date()
                        .toLocaleString()
                )}
            </div>


            ${bodyHtml}


            <script>

                window.onload =
                    function () {

                        window.print();
                    };

            <\/script>

        </body>

        </html>
    `);


    printWindow
        .document
        .close();
}


// ============================================================
// EXCEL WORKBOOK
// ============================================================

function downloadExcelWorkbook(
    filename,
    sheets
) {

    let worksheets =
        "";


    for (
        const sheet
        of sheets
    ) {

        let tableRows =
            createExcelRow(
                sheet.headers,
                true
            );


        for (
            const row
            of sheet.rows
        ) {

            tableRows +=
                createExcelRow(
                    row,
                    false
                );
        }


        worksheets += `

            <Worksheet
                ss:Name="${escapeXml(
                    safeSheetName(
                        sheet.name
                    )
                )}"
            >

                <Table>
                    ${tableRows}
                </Table>

            </Worksheet>
        `;
    }


    const xml = `

        <?xml version="1.0"?>

        <?mso-application progid="Excel.Sheet"?>

        <Workbook

            xmlns="urn:schemas-microsoft-com:office:spreadsheet"

            xmlns:o="urn:schemas-microsoft-com:office:office"

            xmlns:x="urn:schemas-microsoft-com:office:excel"

            xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
        >

            <Styles>

                <Style ss:ID="Header">

                    <Font ss:Bold="1"/>

                    <Interior
                        ss:Color="#D9EAF7"
                        ss:Pattern="Solid"
                    />

                </Style>

            </Styles>


            ${worksheets}

        </Workbook>
    `;


    downloadBlob(
        filename,
        xml,
        "application/vnd.ms-excel"
    );
}


function createExcelRow(
    cells,
    header =
        false
) {

    const cellXml =
        cells
            .map(
                value => {

                    const numeric =
                        typeof value ===
                        "number";


                    return `

                        <Cell ${
                            header
                                ?
                                'ss:StyleID="Header"'
                                :
                                ""
                        }>

                            <Data
                                ss:Type="${
                                    numeric
                                        ?
                                        "Number"
                                        :
                                        "String"
                                }"
                            >
                                ${escapeXml(
                                    value ??
                                    ""
                                )}
                            </Data>

                        </Cell>
                    `;
                }
            )
            .join(
                ""
            );


    return `

        <Row>
            ${cellXml}
        </Row>
    `;
}


// ============================================================
// PM COMPLETE / DEFER MODAL
// ============================================================

function openPMModal(
    scheduleId,
    action
) {

    currentScheduleId =
        scheduleId;

    currentAction =
        action;


    const task =
        allSchedule.find(
            item =>
                Number(
                    item.id
                ) ===
                Number(
                    scheduleId
                )
        );


    if (
        !task
    ) {

        return;
    }


    byId(
        "pmModal"
    )
        .classList
        .remove(
            "hidden"
        );


    setText(
        "modalMessage",
        ""
    );


    byId(
        "technicianName"
    ).value =
        task.technician_name ||
        "";


    byId(
        "pmNotes"
    ).value =
        task.notes ||
        "";


    byId(
        "deferredReason"
    ).value =
        task.deferred_reason ||
        "";


    byId(
        "modalTask"
    ).innerHTML = `

        <strong>

            ${escapeHtml(
                task.machine_code ||
                ""
            )}

            -

            ${escapeHtml(
                task.machine_name ||
                ""
            )}

        </strong>

        <br>

        ${escapeHtml(
            task.part_name ||
            "General"
        )}

        <br>

        ${escapeHtml(
            task.maintenance_task ||
            ""
        )}

        <br>

        Frequency:

        ${escapeHtml(
            task.frequency_text ||
            ""
        )}
    `;


    if (
        action ===
        "complete"
    ) {

        setText(
            "modalTitle",
            "Complete Preventive Maintenance"
        );


        byId(
            "deferredReasonGroup"
        )
            .classList
            .add(
                "hidden"
            );


        setText(
            "confirmAction",
            "✓ Confirm Complete"
        );


        byId(
            "confirmAction"
        )
            .style
            .background =
            "#16a34a";


    } else {

        setText(
            "modalTitle",
            "Defer Preventive Maintenance"
        );


        byId(
            "deferredReasonGroup"
        )
            .classList
            .remove(
                "hidden"
            );


        setText(
            "confirmAction",
            "Confirm Deferred"
        );


        byId(
            "confirmAction"
        )
            .style
            .background =
            "#f59e0b";
    }
}


// ============================================================
// CLOSE MODAL
// ============================================================

function closeModal() {

    byId(
        "pmModal"
    )
        .classList
        .add(
            "hidden"
        );


    currentScheduleId =
        null;

    currentAction =
        null;
}


// ============================================================
// CONFIRM PM ACTION
// ============================================================

async function confirmPMAction() {

    const technician =
        byId(
            "technicianName"
        )
            .value
            .trim();


    const notes =
        byId(
            "pmNotes"
        )
            .value
            .trim();


    const reason =
        byId(
            "deferredReason"
        ).value;


    const message =
        byId(
            "modalMessage"
        );


    if (
        !technician
    ) {

        message.textContent =
            "Please enter technician name.";

        return;
    }


    if (
        currentAction ===
        "defer" &&
        !reason
    ) {

        message.textContent =
            "Please select a deferred reason.";

        return;
    }


    const button =
        byId(
            "confirmAction"
        );


    const originalText =
        button.textContent;


    button.disabled =
        true;


    button.textContent =
        "Saving...";


    try {

        let endpoint;


        const body = {

            technician_name:
                technician,

            notes
        };


        if (
            currentAction ===
            "complete"
        ) {

            endpoint =
                `/api/schedule/${currentScheduleId}/complete`;

        } else {

            endpoint =
                `/api/schedule/${currentScheduleId}/defer`;


            body.deferred_reason =
                reason;
        }


        const response =
            await fetch(
                endpoint,
                {

                    method:
                        "PUT",

                    headers: {

                        "Content-Type":
                            "application/json"
                    },

                    body:
                        JSON.stringify(
                            body
                        )
                }
            );


        const data =
            await response.json();


        if (
            !response.ok
        ) {

            throw new Error(
                data.error ||
                "PM update failed."
            );
        }


        message.style.color =
            "#15803d";


        message.textContent =
            currentAction ===
            "complete"
                ?
                "PM completed successfully ✓"
                :
                "PM deferred successfully";


        loadedAnnualYear =
            null;


        await reloadCurrentWeek();


        setTimeout(
            closeModal,
            500
        );


    } catch (
        error
    ) {

        console.error(
            error
        );


        message.style.color =
            "#b91c1c";


        message.textContent =
            error.message;


    } finally {

        button.disabled =
            false;


        button.textContent =
            originalText;
    }
}


// ============================================================
// RELOAD CURRENT WEEK
// ============================================================

async function reloadCurrentWeek() {

    const response =
        await fetch(
            `/api/schedule?year=${selectedYear}&week=${selectedWeek}`
        );


    const data =
        await response.json();


    if (
        response.ok &&
        Array.isArray(
            data
        )
    ) {

        allSchedule =
            data;


        updateSummary();


        await loadLines();


        populateScheduleFilters();


        renderScheduleTable();


        if (
            currentLineId
        ) {

            await loadMachines(
                currentLineId
            );
        }


        if (
            currentMachineId
        ) {

            showMachinePMs(
                currentMachineId
            );
        }
    }
}


// ============================================================
// SHOW LINES
// ============================================================

function showLines() {

    const panel =
        document
            .querySelector(
                "#dashboardView > .panel"
            );


    if (
        panel
    ) {

        panel
            .classList
            .remove(
                "hidden"
            );
    }


    byId(
        "machineSection"
    )
        .classList
        .add(
            "hidden"
        );


    byId(
        "pmSection"
    )
        .classList
        .add(
            "hidden"
        );
}


// ============================================================
// SHOW MACHINES
// ============================================================

function showMachines() {

    byId(
        "pmSection"
    )
        .classList
        .add(
            "hidden"
        );


    byId(
        "machineSection"
    )
        .classList
        .remove(
            "hidden"
        );
}


// ============================================================
// ISO WEEK HELPERS
// ============================================================

function getISOWeekMonday(
    year,
    week
) {

    const jan4 =
        new Date(
            year,
            0,
            4
        );


    const jan4Day =
        jan4.getDay() ||
        7;


    const monday =
        new Date(
            jan4
        );


    monday.setDate(

        jan4.getDate() -

        jan4Day +

        1 +

        (
            week -
            1
        ) *
        7
    );


    return monday;
}


function getISOWeekNumber(
    date
) {

    const working =
        new Date(
            Date.UTC(

                date.getFullYear(),

                date.getMonth(),

                date.getDate()
            )
        );


    const day =
        working.getUTCDay() ||
        7;


    working.setUTCDate(

        working.getUTCDate() +

        4 -

        day
    );


    const yearStart =
        new Date(
            Date.UTC(
                working.getUTCFullYear(),
                0,
                1
            )
        );


    return Math.ceil(

        (

            (

                working -

                yearStart

            ) /

            86400000 +

            1

        ) /

        7
    );
}


function getISOWeeksInYear(
    year
) {

    const date =
        new Date(
            year,
            11,
            28
        );


    return getISOWeekNumber(
        date
    );
}


function getPlannedMonth(
    year,
    week
) {

    return (
        getISOWeekMonday(
            year,
            week
        )
            .getMonth() +
        1
    );
}


function getWeeksForMonth(
    year,
    month
) {

    const weeks =
        [];


    const maxWeek =
        getISOWeeksInYear(
            year
        );


    for (
        let week = 1;
        week <= maxWeek;
        week++
    ) {

        if (
            getPlannedMonth(
                year,
                week
            ) ===
            month
        ) {

            weeks.push(
                week
            );
        }
    }


    return weeks;
}


// ============================================================
// MONTH NAME
// ============================================================

function getMonthName(
    month
) {

    const months = [

        "",

        "January",

        "February",

        "March",

        "April",

        "May",

        "June",

        "July",

        "August",

        "September",

        "October",

        "November",

        "December"
    ];


    return months[
        month
    ] ||
    "";
}


// ============================================================
// STATUS COUNTER
// ============================================================

function countStatus(
    rows,
    status
) {

    return rows
        .filter(
            item =>
                item.status ===
                status
        )
        .length;
}


// ============================================================
// STATUS PRIORITY
// ============================================================

function getStatusPriority(
    status
) {

    if (
        status ===
        "Pending"
    ) {

        return 1;
    }


    if (
        status ===
        "Deferred"
    ) {

        return 2;
    }


    if (
        status ===
        "Completed"
    ) {

        return 3;
    }


    return 4;
}


// ============================================================
// STATUS CSS
// ============================================================

function getStatusClass(
    status
) {

    if (
        status ===
        "Completed"
    ) {

        return "status-completed";
    }


    if (
        status ===
        "Deferred"
    ) {

        return "status-deferred";
    }


    if (
        status ===
        "Overdue"
    ) {

        return "status-overdue";
    }


    return "status-pending";
}


function getCardStatusClass(
    status
) {

    if (
        status ===
        "Completed"
    ) {

        return "pm-card-completed";
    }


    if (
        status ===
        "Deferred"
    ) {

        return "pm-card-deferred";
    }


    return "pm-card-pending";
}


// ============================================================
// ADHERENCE
// ============================================================

function calculatePercent(
    value,
    total
) {

    if (
        !total
    ) {

        return "0.0";
    }


    return (

        (
            Number(
                value
            ) /

            Number(
                total
            )
        ) *

        100

    )
        .toFixed(
            1
        );
}


function getAdherenceClass(
    percentage
) {

    const value =
        Number(
            percentage
        );


    if (
        value >=
        95
    ) {

        return "adherence-good";
    }


    if (
        value >=
        85
    ) {

        return "adherence-warning";
    }


    return "adherence-poor";
}


function setAdherenceElement(
    elementId,
    percentage
) {

    const element =
        byId(
            elementId
        );


    if (
        !element
    ) {

        return;
    }


    element.textContent =
        `${percentage}%`;


    element
        .classList
        .remove(

            "adherence-good",

            "adherence-warning",

            "adherence-poor"
        );


    element
        .classList
        .add(
            getAdherenceClass(
                percentage
            )
        );
}


// ============================================================
// ACTION DATE
// ============================================================

function getActionDate(
    item
) {

    if (
        item.completed_at
    ) {

        return formatDate(
            item.completed_at
        );
    }


    return "-";
}


// ============================================================
// SEARCH
// ============================================================

function rowMatchesSearch(
    item,
    search
) {

    const text =
        [

            item.line_name,

            item.machine_code,

            item.machine_name,

            item.section,

            item.part_name,

            item.maintenance_task,

            item.frequency_text,

            item.status,

            item.technician_name,

            item.deferred_reason,

            item.notes

        ]
            .join(
                " "
            )
            .toLowerCase();


    return text
        .includes(
            search
        );
}


// ============================================================
// GENERIC SELECT
// ============================================================

function populateSelectFromData(
    elementId,
    data,
    field,
    defaultText
) {

    const select =
        byId(
            elementId
        );


    if (
        !select
    ) {

        return;
    }


    const oldValue =
        select.value;


    const values =
        [
            ...new Set(
                data
                    .map(
                        item =>
                            item[
                                field
                            ]
                    )
                    .filter(
                        Boolean
                    )
            )
        ]
            .sort(
                (
                    a,
                    b
                ) =>
                    String(
                        a
                    )
                        .localeCompare(
                            String(
                                b
                            )
                        )
            );


    select.innerHTML = `

        <option value="">
            ${escapeHtml(
                defaultText
            )}
        </option>
    `;


    for (
        const value
        of values
    ) {

        const option =
            document
                .createElement(
                    "option"
                );


        option.value =
            value;


        option.textContent =
            value;


        select
            .appendChild(
                option
            );
    }


    if (
        values
            .map(
                String
            )
            .includes(
                String(
                    oldValue
                )
            )
    ) {

        select.value =
            oldValue;
    }
}


// ============================================================
// MACHINE SELECT
// ============================================================

function populateMachineSelect(
    elementId,
    rows,
    defaultText
) {

    const select =
        byId(
            elementId
        );


    if (
        !select
    ) {

        return;
    }


    const oldValue =
        select.value;


    const machines =
        new Map();


    for (
        const item
        of rows
    ) {

        machines.set(
            String(
                item.machine_id
            ),
            {

                id:
                    item.machine_id,

                code:
                    item.machine_code,

                name:
                    item.machine_name
            }
        );
    }


    const sorted =
        [
            ...machines.values()
        ]
            .sort(
                (
                    a,
                    b
                ) =>
                    String(
                        a.code ||
                        a.name
                    )
                        .localeCompare(
                            String(
                                b.code ||
                                b.name
                            )
                        )
            );


    select.innerHTML = `

        <option value="">
            ${escapeHtml(
                defaultText
            )}
        </option>
    `;


    for (
        const machine
        of sorted
    ) {

        const option =
            document
                .createElement(
                    "option"
                );


        option.value =
            String(
                machine.id
            );


        option.textContent =
            `${machine.code || ""} - ${machine.name || ""}`;


        select
            .appendChild(
                option
            );
    }


    if (
        sorted
            .map(
                item =>
                    String(
                        item.id
                    )
            )
            .includes(
                String(
                    oldValue
                )
            )
    ) {

        select.value =
            oldValue;
    }
}


// ============================================================
// YEAR SELECTORS
// ============================================================

function syncYearSelectors(
    year
) {

    const ids = [

        "annualYearSelect",

        "reportYearSelect",

        "calendarYearFilter"
    ];


    for (
        const id
        of ids
    ) {

        const element =
            byId(
                id
            );


        if (
            element
        ) {

            element.value =
                String(
                    year
                );
        }
    }
}


// ============================================================
// EXCEL ACTION DATA
// ============================================================

function getActionHeaders() {

    return [

        "Year",

        "Week",

        "Production Line",

        "Machine Code",

        "Machine Name",

        "Section",

        "Part",

        "PM Task",

        "Frequency",

        "Status",

        "Technician",

        "Action Date",

        "Deferred Reason",

        "Notes"
    ];
}


function actionRowToArray(
    item
) {

    return [

        item.planned_year,

        item.planned_week,

        item.line_name ||
        "",

        item.machine_code ||
        "",

        item.machine_name ||
        "",

        item.section ||
        "",

        item.part_name ||
        "",

        item.maintenance_task ||
        "",

        item.frequency_text ||
        "",

        item.status ||
        "",

        item.technician_name ||
        "",

        item.completed_at
            ?
            formatDate(
                item.completed_at
            )
            :
            "",

        item.deferred_reason ||
        "",

        item.notes ||
        ""
    ];
}


// ============================================================
// FORMAT DATE
// ============================================================

function formatDate(
    value
) {

    if (
        !value
    ) {

        return "-";
    }


    const date =
        new Date(
            value
        );


    return date
        .toLocaleString();
}


// ============================================================
// SAFE SHEET NAME
// ============================================================

function safeSheetName(
    name
) {

    return String(
        name ||
        "Sheet"
    )
        .replace(
            /[\\\/\?\*\[\]\:]/g,
            " "
        )
        .substring(
            0,
            31
        );
}


// ============================================================
// DOWNLOAD FILE
// ============================================================

function downloadBlob(
    filename,
    content,
    type
) {

    const blob =
        new Blob(
            [
                content
            ],
            {

                type:
                    `${type};charset=utf-8`
            }
        );


    const url =
        URL.createObjectURL(
            blob
        );


    const link =
        document
            .createElement(
                "a"
            );


    link.href =
        url;


    link.download =
        filename;


    document.body
        .appendChild(
            link
        );


    link.click();


    link.remove();


    URL.revokeObjectURL(
        url
    );
}


// ============================================================
// ELEMENT HELPER
// ============================================================

function byId(
    id
) {

    return document
        .getElementById(
            id
        );
}


// ============================================================
// SET TEXT
// ============================================================

function setText(
    id,
    value
) {

    const element =
        byId(
            id
        );


    if (
        element
    ) {

        element.textContent =
            value;
    }
}


// ============================================================
// ESCAPE HTML
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


// ============================================================
// ESCAPE XML
// ============================================================

function escapeXml(
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
            "&apos;"
        );
}// ============================================================
// WEEK START / END DATE FIX
// ============================================================

function getISOWeekDateRange(
    year,
    week
) {

    const startDate =
        getISOWeekMonday(
            year,
            week
        );


    const endDate =
        new Date(
            startDate
        );


    endDate.setDate(
        startDate.getDate() + 6
    );


    return {

        start:
            startDate,

        end:
            endDate,

        text:
            `${formatWeekDate(startDate)} - ${formatWeekDate(endDate)}`
    };
}


// ============================================================
// FORMAT WEEK DATE
// ============================================================

function formatWeekDate(
    date
) {

    if (
        !date ||
        Number.isNaN(
            date.getTime()
        )
    ) {

        return "";
    }


    return date.toLocaleDateString(
        "en-GB",
        {
            day:
                "2-digit",

            month:
                "short",

            year:
                "numeric"
        }
    );
}// ============================================================
// WEEK NAVIGATION FIX
// ============================================================


// ============================================================
// PREVIOUS WEEK
// ============================================================

async function previousWeek() {

    selectedWeek--;


    // If we go before Week 1,
    // move to the last ISO week of previous year.

    if (
        selectedWeek <
        1
    ) {

        selectedYear--;


        selectedWeek =
            getISOWeeksInYear(
                selectedYear
            );
    }


    // Update header week / year / dates.

    updateWeekDisplay();


    // Clear annual cache only if needed.

    loadedAnnualYear =
        null;


    // Load selected week's PM schedule.

    await loadSchedule();


    // Return dashboard to production lines.

    showMainView(
        "dashboard"
    );
}


// ============================================================
// NEXT WEEK
// ============================================================

async function nextWeek() {

    const maxWeek =
        getISOWeeksInYear(
            selectedYear
        );


    selectedWeek++;


    // If we go beyond last ISO week,
    // move to Week 1 of next year.

    if (
        selectedWeek >
        maxWeek
    ) {

        selectedYear++;


        selectedWeek =
            1;
    }


    // Update header week / year / dates.

    updateWeekDisplay();


    // Clear annual cache only if needed.

    loadedAnnualYear =
        null;


    // Load selected week's PM schedule.

    await loadSchedule();


    // Return dashboard to production lines.

    showMainView(
        "dashboard"
    );
}