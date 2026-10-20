/**
 * TimelyAI Interactive Timetable Manager
 * Handles:
 * - Switching between Class View and Teacher View
 * - Dynamic rendering with color-coded subjects and free period indicators
 * - Interactive Edit Mode (add, edit, move, delete sessions)
 * - Real-time teacher & class conflict detection
 * - Client-server state synchronization for exports
 */

(function () {
    'use strict';

    // State Variables
    let currentSessions = [];
    let originalSessions = [];
    let configDays = [];
    let configRows = [];
    let configPeriods = [];
    let allClasses = [];
    let allTeachers = [];
    let allSubjects = [];

    let currentView = 'class'; // 'class' or 'teacher'
    let selectedClass = '';
    let selectedTeacher = '';
    let editMode = false;
    let currentEditingId = null;

    // Palette for distinct subject badges
    const SUBJECT_COLORS = [
        { bg: '#EEF2FF', text: '#3730A3', border: '#C7D2FE' }, // Indigo
        { bg: '#F0FDF4', text: '#166534', border: '#BBF7D0' }, // Green
        { bg: '#EFF6FF', text: '#1E40AF', border: '#BFDBFE' }, // Blue
        { bg: '#FAF5FF', text: '#6B21A8', border: '#E9D5FF' }, // Purple
        { bg: '#FFF7ED', text: '#9A3412', border: '#FED7AA' }, // Orange
        { bg: '#ECFDF5', text: '#065F46', border: '#A7F3D0' }, // Emerald
        { bg: '#FDF2F8', text: '#9D174D', border: '#FBCFE8' }, // Pink
        { bg: '#F0FDFA', text: '#115E59', border: '#99F6E4' }, // Teal
        { bg: '#FEF2F2', text: '#991B1B', border: '#FECACA' }, // Rose
    ];

    function getSubjectColor(subjectName) {
        if (!subjectName) return SUBJECT_COLORS[0];
        let hash = 0;
        for (let i = 0; i < subjectName.length; i++) {
            hash = subjectName.charCodeAt(i) + ((hash << 5) - hash);
        }
        const index = Math.abs(hash) % SUBJECT_COLORS.length;
        return SUBJECT_COLORS[index];
    }

    // Initialize from embedded data
    window.initTimetableApp = function (data) {
        currentSessions = JSON.parse(JSON.stringify(data.timetable || []));
        originalSessions = JSON.parse(JSON.stringify(data.timetable || []));
        configDays = data.days || [];
        configRows = data.rows || [];
        configPeriods = data.periods || [];
        allClasses = data.classes || [];
        allTeachers = data.teachers || [];
        allSubjects = data.all_subjects || [];

        // Ensure every session has a unique id
        currentSessions.forEach((s, i) => {
            if (!s.id) s.id = 's_' + (i + 1);
        });

        if (allClasses.length > 0) selectedClass = allClasses[0];
        if (allTeachers.length > 0) selectedTeacher = allTeachers[0];

        setupEventHandlers();
        populateDropdowns();
        renderActiveView();
        updateStats();
    };

    function setupEventHandlers() {
        // Tab switching
        const btnViewClass = document.getElementById('btnViewClass');
        const btnViewTeacher = document.getElementById('btnViewTeacher');

        if (btnViewClass) {
            btnViewClass.addEventListener('click', () => switchView('class'));
        }
        if (btnViewTeacher) {
            btnViewTeacher.addEventListener('click', () => switchView('teacher'));
        }

        // Dropdown changes
        const classSelect = document.getElementById('classSelect');
        if (classSelect) {
            classSelect.addEventListener('change', (e) => {
                selectedClass = e.target.value;
                renderClassTimetable();
            });
        }

        const teacherSelect = document.getElementById('teacherSelect');
        if (teacherSelect) {
            teacherSelect.addEventListener('change', (e) => {
                selectedTeacher = e.target.value;
                renderTeacherTimetable();
            });
        }

        // Edit Mode Toggle
        const btnToggleEdit = document.getElementById('btnToggleEdit');
        if (btnToggleEdit) {
            btnToggleEdit.addEventListener('click', toggleEditMode);
        }

        // Reset Timetable Button
        const btnResetTimetable = document.getElementById('btnResetTimetable');
        if (btnResetTimetable) {
            btnResetTimetable.addEventListener('click', resetTimetable);
        }

        // Modal Form Inputs Conflict Detection
        ['modalTeacher', 'modalDay', 'modalPeriod', 'modalLength', 'modalClass'].forEach(id => {
            const el = document.getElementById(id);
            if (el) {
                el.addEventListener('change', checkModalConflict);
            }
        });

        // Modal Save Button
        const btnSaveModal = document.getElementById('btnSaveModal');
        if (btnSaveModal) {
            btnSaveModal.addEventListener('click', saveModalSession);
        }

        // Modal Delete Button
        const btnDeleteModal = document.getElementById('btnDeleteModal');
        if (btnDeleteModal) {
            btnDeleteModal.addEventListener('click', deleteModalSession);
        }

        // Modal Cancel / Close
        const btnCancelModal = document.getElementById('btnCancelModal');
        const dialog = document.getElementById('sessionModal');
        if (btnCancelModal && dialog) {
            btnCancelModal.addEventListener('click', () => dialog.close());
        }

        // Export dropdown buttons
        const btnDownloadMain = document.getElementById('btnDownloadMain');
        if (btnDownloadMain) {
            btnDownloadMain.addEventListener('click', downloadMainExcel);
        }
        const btnDownloadAllTeachers = document.getElementById('btnDownloadAllTeachers');
        if (btnDownloadAllTeachers) {
            btnDownloadAllTeachers.addEventListener('click', downloadAllTeachersExcel);
        }
        const btnDownloadCurrent = document.getElementById('btnDownloadCurrent');
        if (btnDownloadCurrent) {
            btnDownloadCurrent.addEventListener('click', downloadCurrentViewExcel);
        }
    }

    function switchView(viewName) {
        currentView = viewName;
        const btnClass = document.getElementById('btnViewClass');
        const btnTeacher = document.getElementById('btnViewTeacher');
        const panelClass = document.getElementById('panelClassView');
        const panelTeacher = document.getElementById('panelTeacherView');

        if (viewName === 'class') {
            btnClass.classList.add('active');
            btnTeacher.classList.remove('active');
            panelClass.style.display = 'block';
            panelTeacher.style.display = 'none';
            renderClassTimetable();
        } else {
            btnClass.classList.remove('active');
            btnTeacher.classList.add('active');
            panelClass.style.display = 'none';
            panelTeacher.style.display = 'block';
            renderTeacherTimetable();
        }
    }

    function toggleEditMode() {
        editMode = !editMode;
        const btn = document.getElementById('btnToggleEdit');
        const badge = document.getElementById('editStatusBadge');

        if (editMode) {
            btn.classList.add('active');
            btn.innerHTML = '<span>✏️ Exit Edit Mode</span>';
            if (badge) {
                badge.style.display = 'inline-flex';
                badge.innerText = '● Edit Mode Active';
            }
            showToast('Edit Mode Enabled: Click any session to edit or click empty cells to add sessions.', 'info');
        } else {
            btn.classList.remove('active');
            btn.innerHTML = '<span>✏️ Edit Timetable</span>';
            if (badge) badge.style.display = 'none';
        }

        renderActiveView();
    }

    function resetTimetable() {
        if (!confirm('Are you sure you want to reset all modifications back to the original AI-generated timetable?')) {
            return;
        }
        currentSessions = JSON.parse(JSON.stringify(originalSessions));
        syncWithServer();
        renderActiveView();
        updateStats();
        showToast('Timetable reset to original generated schedule.', 'success');
    }

    function populateDropdowns() {
        // Class Selector
        const classSelect = document.getElementById('classSelect');
        if (classSelect) {
            classSelect.innerHTML = '';
            allClasses.forEach(cls => {
                const opt = document.createElement('option');
                opt.value = cls;
                opt.textContent = cls;
                if (cls === selectedClass) opt.selected = true;
                classSelect.appendChild(opt);
            });
        }

        // Teacher Selector
        const teacherSelect = document.getElementById('teacherSelect');
        if (teacherSelect) {
            teacherSelect.innerHTML = '';
            allTeachers.forEach(tch => {
                const opt = document.createElement('option');
                opt.value = tch;
                opt.textContent = '👨‍🏫 ' + tch;
                if (tch === selectedTeacher) opt.selected = true;
                teacherSelect.appendChild(opt);
            });
        }

        // Modal Selectors
        populateModalDropdowns();
    }

    function populateModalDropdowns() {
        const modalClass = document.getElementById('modalClass');
        const modalTeacher = document.getElementById('modalTeacher');
        const modalSubject = document.getElementById('modalSubject');
        const modalDay = document.getElementById('modalDay');
        const modalPeriod = document.getElementById('modalPeriod');

        if (modalClass) {
            modalClass.innerHTML = '';
            allClasses.forEach(cls => {
                const opt = document.createElement('option');
                opt.value = cls;
                opt.textContent = cls;
                modalClass.appendChild(opt);
            });
        }

        if (modalTeacher) {
            modalTeacher.innerHTML = '';
            allTeachers.forEach(tch => {
                const opt = document.createElement('option');
                opt.value = tch;
                opt.textContent = tch;
                modalTeacher.appendChild(opt);
            });
        }

        if (modalSubject) {
            modalSubject.innerHTML = '';
            allSubjects.forEach(sbj => {
                const opt = document.createElement('option');
                opt.value = sbj;
                opt.textContent = sbj;
                modalSubject.appendChild(opt);
            });
        }

        if (modalDay) {
            modalDay.innerHTML = '';
            configDays.forEach(day => {
                const opt = document.createElement('option');
                opt.value = day;
                opt.textContent = day;
                modalDay.appendChild(opt);
            });
        }

        if (modalPeriod) {
            modalPeriod.innerHTML = '';
            configPeriods.forEach(p => {
                const opt = document.createElement('option');
                opt.value = p;
                opt.textContent = 'Period ' + p;
                modalPeriod.appendChild(opt);
            });
        }
    }

    function renderActiveView() {
        if (currentView === 'class') {
            renderClassTimetable();
        } else {
            renderTeacherTimetable();
        }
    }

    // ==========================================
    // Class Timetable Rendering
    // ==========================================
    function renderClassTimetable() {
        const container = document.getElementById('classTimetableContent');
        if (!container) return;

        if (!selectedClass) {
            container.innerHTML = '<div class="empty-state">No class selected.</div>';
            return;
        }

        // Build grid for selected class: grid[day][period] = { session, skip }
        const grid = {};
        configDays.forEach(d => { grid[d] = {}; });

        const classSessions = currentSessions.filter(s => s.class_id === selectedClass);

        classSessions.forEach(item => {
            const len = item.periods_per_session || 1;
            if (grid[item.day]) {
                grid[item.day][item.period] = { session: item, skip: false };
                for (let extra = 1; extra < len; extra++) {
                    grid[item.day][item.period + extra] = { skip: true };
                }
            }
        });

        // Compute class stats
        const classLoad = classSessions.reduce((acc, s) => acc + (s.periods_per_session || 1), 0);
        const classSubjs = [...new Set(classSessions.map(s => s.subject))];

        let html = `
            <div class="schedule-summary-card">
                <div class="summary-info">
                    <h4>Class: <span class="highlight-text">${selectedClass}</span></h4>
                    <p class="summary-meta">
                        <span>⏱️ <strong>${classLoad}</strong> Periods / Week</span> •
                        <span>📚 <strong>${classSubjs.length}</strong> Subjects</span>
                    </p>
                </div>
                <div class="summary-actions">
                    <button class="btn-sm btn-outline-action" onclick="window.downloadSingleClassExcel('${selectedClass}')">
                        📥 Download ${selectedClass} Excel
                    </button>
                    <button class="btn-sm btn-outline-action" onclick="window.print()">
                        🖨️ Print PDF
                    </button>
                </div>
            </div>

            <div class="timetable-wrapper">
                <table class="timetable-table ${editMode ? 'edit-mode-active' : ''}">
                    <thead>
                        <tr>
                            <th class="period-header">Period</th>
                            ${configDays.map(day => `<th>${day}</th>`).join('')}
                        </tr>
                    </thead>
                    <tbody>
        `;

        configRows.forEach(row => {
            if (row.type === 'break') {
                html += `
                    <tr class="break-row">
                        <td class="period-header">
                            ${row.label}
                            <span class="period-time">${row.start} – ${row.end}</span>
                        </td>
                        <td colspan="${configDays.length}">
                            <div class="break-content">🍽️ ${row.label} Break (${row.start} - ${row.end})</div>
                        </td>
                    </tr>
                `;
            } else {
                const p = row.period;
                html += `
                    <tr>
                        <td class="period-header">
                            Period ${p}
                            <span class="period-time">${row.start} – ${row.end}</span>
                        </td>
                `;

                configDays.forEach(day => {
                    const cell = grid[day] ? grid[day][p] : null;

                    if (cell && cell.skip) {
                        // Covered by rowspan
                    } else if (cell && cell.session) {
                        const s = cell.session;
                        const len = s.periods_per_session || 1;
                        const color = getSubjectColor(s.subject);
                        const shared = s.shared_with && s.shared_with.length > 0
                            ? `<div class="cell-shared">🔗 Common: ${s.shared_with.join(', ')}</div>`
                            : '';

                        html += `
                            <td rowspan="${len}" class="table-session-cell">
                                <div class="cell-session-card" style="background:${color.bg}; border-left-color:${color.text};">
                                    <div class="cell-subject" style="color:${color.text};">${s.subject}</div>
                                    <div class="cell-teacher">👤 ${s.teacher}</div>
                                    ${shared}
                                    <div class="cell-time">⏰ ${s.start_time || ''} - ${s.end_time || ''} (${len > 1 ? len + ' periods' : '1 period'})</div>
                                    ${editMode ? `
                                        <div class="cell-edit-overlay">
                                            <button type="button" class="btn-cell-action edit-btn" onclick="window.openEditModal('${s.id}')" title="Edit Session">✏️ Edit</button>
                                            <button type="button" class="btn-cell-action delete-btn" onclick="window.quickDeleteSession('${s.id}')" title="Delete Session">🗑️</button>
                                        </div>
                                    ` : ''}
                                </div>
                            </td>
                        `;
                    } else {
                        // Empty slot
                        html += `
                            <td class="table-empty-cell">
                                ${editMode ? `
                                    <button type="button" class="btn-add-slot" onclick="window.openAddModal('${selectedClass}', '', '${day}', ${p})" title="Add Class Session here">
                                        + Add
                                    </button>
                                ` : `<span class="cell-empty">—</span>`}
                            </td>
                        `;
                    }
                });

                html += `</tr>`;
            }
        });

        html += `
                    </tbody>
                </table>
            </div>
        `;

        container.innerHTML = html;
    }

    // ==========================================
    // Teacher Timetable Rendering
    // ==========================================
    function renderTeacherTimetable() {
        const container = document.getElementById('teacherTimetableContent');
        if (!container) return;

        if (!selectedTeacher) {
            container.innerHTML = '<div class="empty-state">No faculty member selected.</div>';
            return;
        }

        // Build teacher grid: grid[day][period] = { session, skip }
        const grid = {};
        configDays.forEach(d => { grid[d] = {}; });

        const teacherSessions = currentSessions.filter(s => s.teacher === selectedTeacher);

        teacherSessions.forEach(item => {
            const len = item.periods_per_session || 1;
            if (grid[item.day]) {
                // If parallel session already exists, merge class labels
                if (grid[item.day][item.period] && grid[item.day][item.period].session) {
                    const ex = grid[item.day][item.period].session;
                    const setCls = new Set((ex.class_ids || [ex.class_id]).concat(item.class_id));
                    ex.class_ids = Array.from(setCls);
                    ex.display_class = ex.class_ids.join(', ');
                } else {
                    const copy = JSON.parse(JSON.stringify(item));
                    copy.class_ids = [item.class_id].concat(item.shared_with || []);
                    copy.display_class = Array.from(new Set(copy.class_ids)).join(', ');

                    grid[item.day][item.period] = { session: copy, skip: false };
                    for (let extra = 1; extra < len; extra++) {
                        grid[item.day][item.period + extra] = { skip: true };
                    }
                }
            }
        });

        // Compute teacher stats
        const totalPeriods = teacherSessions.reduce((acc, s) => acc + (s.periods_per_session || 1), 0);
        const subjects = [...new Set(teacherSessions.map(s => s.subject))];
        const classes = [...new Set(teacherSessions.map(s => s.class_id))];
        const activeDays = [...new Set(teacherSessions.map(s => s.day))];

        let html = `
            <div class="faculty-profile-banner">
                <div class="faculty-avatar">👨‍🏫</div>
                <div class="faculty-details">
                    <div class="faculty-name-row">
                        <h3>Prof. / Dr. <span class="highlight-text">${selectedTeacher}</span></h3>
                        <span class="badge-role">Faculty Schedule</span>
                    </div>
                    <div class="faculty-stats-grid">
                        <div class="stat-pill">
                            <span class="stat-icon">⏱️</span>
                            <span><strong>${totalPeriods}</strong> Periods / Week</span>
                        </div>
                        <div class="stat-pill">
                            <span class="stat-icon">📚</span>
                            <span><strong>${subjects.length}</strong> Subject(s): ${subjects.join(', ')}</span>
                        </div>
                        <div class="stat-pill">
                            <span class="stat-icon">👥</span>
                            <span>Classes: ${classes.join(', ')}</span>
                        </div>
                        <div class="stat-pill">
                            <span class="stat-icon">📅</span>
                            <span>Active on <strong>${activeDays.length}</strong> Days</span>
                        </div>
                    </div>
                </div>
                <div class="faculty-actions">
                    <a href="/download/teacher-excel/${encodeURIComponent(selectedTeacher)}" class="btn-sm btn-primary-action">
                        📥 Download ${selectedTeacher}'s Timetable (Excel)
                    </a>
                    <a href="/download/all-teachers-excel" class="btn-sm btn-outline-action">
                        📥 Download All Teachers (Excel)
                    </a>
                    <button class="btn-sm btn-outline-action" onclick="window.print()">
                        🖨️ Print Faculty PDF
                    </button>
                </div>
            </div>

            <div class="timetable-wrapper">
                <table class="timetable-table ${editMode ? 'edit-mode-active' : ''}">
                    <thead>
                        <tr>
                            <th class="period-header">Period</th>
                            ${configDays.map(day => `<th>${day}</th>`).join('')}
                        </tr>
                    </thead>
                    <tbody>
        `;

        configRows.forEach(row => {
            if (row.type === 'break') {
                html += `
                    <tr class="break-row">
                        <td class="period-header">
                            ${row.label}
                            <span class="period-time">${row.start} – ${row.end}</span>
                        </td>
                        <td colspan="${configDays.length}">
                            <div class="break-content">🍽️ ${row.label} Break (${row.start} - ${row.end})</div>
                        </td>
                    </tr>
                `;
            } else {
                const p = row.period;
                html += `
                    <tr>
                        <td class="period-header">
                            Period ${p}
                            <span class="period-time">${row.start} – ${row.end}</span>
                        </td>
                `;

                configDays.forEach(day => {
                    const cell = grid[day] ? grid[day][p] : null;

                    if (cell && cell.skip) {
                        // Covered by rowspan
                    } else if (cell && cell.session) {
                        const s = cell.session;
                        const len = s.periods_per_session || 1;
                        const color = getSubjectColor(s.subject);
                        const classText = s.display_class || s.class_id;

                        html += `
                            <td rowspan="${len}" class="table-session-cell">
                                <div class="cell-session-card teacher-session-card" style="background:${color.bg}; border-left-color:${color.text};">
                                    <div class="cell-subject" style="color:${color.text};">${s.subject}</div>
                                    <div class="cell-class-badge">🏫 Class: <strong>${classText}</strong></div>
                                    <div class="cell-time">⏰ ${s.start_time || ''} - ${s.end_time || ''}</div>
                                    ${editMode ? `
                                        <div class="cell-edit-overlay">
                                            <button type="button" class="btn-cell-action edit-btn" onclick="window.openEditModal('${s.id}')" title="Edit Session">✏️ Edit</button>
                                            <button type="button" class="btn-cell-action delete-btn" onclick="window.quickDeleteSession('${s.id}')" title="Delete Session">🗑️</button>
                                        </div>
                                    ` : ''}
                                </div>
                            </td>
                        `;
                    } else {
                        // Free period for this teacher
                        html += `
                            <td class="table-empty-cell cell-free-period">
                                ${editMode ? `
                                    <button type="button" class="btn-add-slot" onclick="window.openAddModal('', '${selectedTeacher}', '${day}', ${p})" title="Assign a class session to ${selectedTeacher}">
                                        + Assign
                                    </button>
                                ` : `
                                    <span class="badge-free-slot">✨ Free Period</span>
                                `}
                            </td>
                        `;
                    }
                });

                html += `</tr>`;
            }
        });

        html += `
                    </tbody>
                </table>
            </div>
        `;

        container.innerHTML = html;
    }

    // ==========================================
    // Interactive Edit & Modal Management
    // ==========================================
    window.openEditModal = function (sessionId) {
        currentEditingId = sessionId;
        const session = currentSessions.find(s => s.id === sessionId);
        if (!session) return;

        const dialog = document.getElementById('sessionModal');
        const modalTitle = document.getElementById('modalTitle');
        const btnDelete = document.getElementById('btnDeleteModal');

        if (modalTitle) modalTitle.innerText = '✏️ Edit Scheduled Session';
        if (btnDelete) btnDelete.style.display = 'inline-block';

        document.getElementById('modalClass').value = session.class_id;
        document.getElementById('modalTeacher').value = session.teacher;
        document.getElementById('modalSubject').value = session.subject;
        document.getElementById('modalDay').value = session.day;
        document.getElementById('modalPeriod').value = session.period;
        document.getElementById('modalLength').value = session.periods_per_session || 1;

        checkModalConflict();
        if (dialog) dialog.showModal();
    };

    window.openAddModal = function (cls, tch, day, period) {
        currentEditingId = null;
        const dialog = document.getElementById('sessionModal');
        const modalTitle = document.getElementById('modalTitle');
        const btnDelete = document.getElementById('btnDeleteModal');

        if (modalTitle) modalTitle.innerText = '➕ Add New Class Session';
        if (btnDelete) btnDelete.style.display = 'none';

        if (cls) document.getElementById('modalClass').value = cls;
        else if (selectedClass) document.getElementById('modalClass').value = selectedClass;

        if (tch) document.getElementById('modalTeacher').value = tch;
        else if (selectedTeacher) document.getElementById('modalTeacher').value = selectedTeacher;

        if (day) document.getElementById('modalDay').value = day;
        if (period) document.getElementById('modalPeriod').value = period;
        document.getElementById('modalLength').value = 1;

        checkModalConflict();
        if (dialog) dialog.showModal();
    };

    function checkModalConflict() {
        const conflictBox = document.getElementById('modalConflictWarning');
        if (!conflictBox) return;

        const tch = document.getElementById('modalTeacher').value;
        const cls = document.getElementById('modalClass').value;
        const day = document.getElementById('modalDay').value;
        const period = parseInt(document.getElementById('modalPeriod').value, 10);
        const len = parseInt(document.getElementById('modalLength').value, 10) || 1;

        const conflicts = [];

        // Check if teacher has another class at this time
        for (let p = period; p < period + len; p++) {
            const conflictTeacher = currentSessions.find(s =>
                s.id !== currentEditingId &&
                s.teacher === tch &&
                s.day === day &&
                (s.period <= p && p < s.period + (s.periods_per_session || 1))
            );
            if (conflictTeacher) {
                conflicts.push(`⚠️ Faculty conflict: ${tch} is already teaching "${conflictTeacher.subject}" to ${conflictTeacher.class_id} on ${day} Period ${p}.`);
                break;
            }

            const conflictClass = currentSessions.find(s =>
                s.id !== currentEditingId &&
                s.class_id === cls &&
                s.day === day &&
                (s.period <= p && p < s.period + (s.periods_per_session || 1))
            );
            if (conflictClass) {
                conflicts.push(`⚠️ Class conflict: ${cls} already has "${conflictClass.subject}" with ${conflictClass.teacher} on ${day} Period ${p}.`);
                break;
            }
        }

        if (conflicts.length > 0) {
            conflictBox.innerHTML = conflicts.join('<br>');
            conflictBox.style.display = 'block';
        } else {
            conflictBox.innerHTML = '';
            conflictBox.style.display = 'none';
        }
    }

    function saveModalSession() {
        const cls = document.getElementById('modalClass').value.trim();
        const tch = document.getElementById('modalTeacher').value.trim();
        const sbj = document.getElementById('modalSubject').value.trim();
        const day = document.getElementById('modalDay').value;
        const period = parseInt(document.getElementById('modalPeriod').value, 10);
        const len = parseInt(document.getElementById('modalLength').value, 10) || 1;

        if (!cls || !tch || !sbj || !day || isNaN(period)) {
            alert('Please fill in all session fields (Class, Teacher, Subject, Day, Period).');
            return;
        }

        // Calculate slot time strings if possible from configRows
        const startRow = configRows.find(r => r.type === 'period' && r.period === period);
        const endRow = configRows.find(r => r.type === 'period' && r.period === (period + len - 1)) || startRow;
        const startTime = startRow ? startRow.start : '';
        const endTime = endRow ? endRow.end : '';

        if (currentEditingId) {
            // Update existing
            const session = currentSessions.find(s => s.id === currentEditingId);
            if (session) {
                session.class_id = cls;
                session.teacher = tch;
                session.subject = sbj;
                session.day = day;
                session.period = period;
                session.periods_per_session = len;
                session.start_time = startTime;
                session.end_time = endTime;
            }
            showToast('Session updated successfully!', 'success');
        } else {
            // Create new
            const newSession = {
                id: 's_' + Date.now(),
                class_id: cls,
                teacher: tch,
                subject: sbj,
                day: day,
                period: period,
                periods_per_session: len,
                start_time: startTime,
                end_time: endTime,
                shared_with: []
            };
            currentSessions.push(newSession);

            // Track any new class or teacher or subject
            if (!allClasses.includes(cls)) { allClasses.push(cls); allClasses.sort(); }
            if (!allTeachers.includes(tch)) { allTeachers.push(tch); allTeachers.sort(); }
            if (!allSubjects.includes(sbj)) { allSubjects.push(sbj); allSubjects.sort(); }
            populateDropdowns();

            showToast('New session added successfully!', 'success');
        }

        const dialog = document.getElementById('sessionModal');
        if (dialog) dialog.close();

        syncWithServer();
        renderActiveView();
        updateStats();
    }

    window.quickDeleteSession = function (sessionId) {
        if (!confirm('Are you sure you want to remove this session from the schedule?')) {
            return;
        }
        currentSessions = currentSessions.filter(s => s.id !== sessionId);
        syncWithServer();
        renderActiveView();
        updateStats();
        showToast('Session deleted.', 'info');
    };

    function deleteModalSession() {
        if (currentEditingId) {
            window.quickDeleteSession(currentEditingId);
            const dialog = document.getElementById('sessionModal');
            if (dialog) dialog.close();
        }
    }

    function syncWithServer() {
        fetch('/api/save-timetable', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ sessions: currentSessions })
        })
        .then(res => res.json())
        .then(data => {
            console.log('Saved to server:', data);
        })
        .catch(err => {
            console.warn('Auto-save error:', err);
        });
    }

    function updateStats() {
        const elTotal = document.getElementById('statTotalSessions');
        const elClasses = document.getElementById('statTotalClasses');
        const elTeachers = document.getElementById('statTotalTeachers');

        if (elTotal) elTotal.innerText = currentSessions.length;
        if (elClasses) elClasses.innerText = new Set(currentSessions.map(s => s.class_id)).size;
        if (elTeachers) elTeachers.innerText = new Set(currentSessions.map(s => s.teacher)).size;
    }

    // Resilient serverless-safe Excel export helpers
    function triggerExcelDownload(url, filename) {
        if (!currentSessions || currentSessions.length === 0) {
            window.location.href = url;
            return;
        }

        showToast('Preparing Excel workbook download...', 'info');

        const payload = {
            sessions: currentSessions,
            config: {
                days: configDays,
                rows: configRows,
                periods: configPeriods
            }
        };

        fetch(url, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            },
            body: JSON.stringify(payload)
        })
        .then(response => {
            if (!response.ok) {
                throw new Error('Download request failed with status ' + response.status);
            }
            return response.blob();
        })
        .then(blob => {
            const blobUrl = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = blobUrl;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            setTimeout(() => window.URL.revokeObjectURL(blobUrl), 1500);
            showToast('Excel workbook downloaded successfully!', 'success');
        })
        .catch(err => {
            console.warn('POST download failed, falling back to direct GET download:', err);
            window.location.href = url;
        });
    }

    function downloadMainExcel() {
        triggerExcelDownload('/download/main-excel', 'College_Timetable_Master.xlsx');
    }

    function downloadAllTeachersExcel() {
        triggerExcelDownload('/download/all-teachers-excel', 'Timetable_All_Faculty.xlsx');
    }

    function downloadCurrentViewExcel() {
        if (currentView === 'teacher' && selectedTeacher) {
            const safeName = selectedTeacher.replace(/[^a-zA-Z0-9_\-]/g, '_');
            triggerExcelDownload(`/download/teacher-excel/${encodeURIComponent(selectedTeacher)}`, `Timetable_Faculty_${safeName}.xlsx`);
        } else {
            triggerExcelDownload('/download/main-excel', 'College_Timetable_Master.xlsx');
        }
    }

    // Single class download helper
    window.downloadSingleClassExcel = function (className) {
        const safeName = (className || 'Class').replace(/[^a-zA-Z0-9_\-]/g, '_');
        triggerExcelDownload('/download/main-excel', `Timetable_${safeName}.xlsx`);
    };

    function showToast(msg, type = 'info') {
        let toast = document.getElementById('timetableToast');
        if (!toast) {
            toast = document.createElement('div');
            toast.id = 'timetableToast';
            document.body.appendChild(toast);
        }
        toast.className = `timetable-toast toast-${type} show`;
        toast.innerText = msg;

        setTimeout(() => {
            toast.classList.remove('show');
        }, 3500);
    }

})();
