// Service Worker Registration
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('./service-worker.js')
            .then(reg => console.log('Service Worker registered!'))
            .catch(err => console.log('Service Worker failed:', err));
    });
}

import { initializeApp } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-app.js";
import { getAuth, signInAnonymously, onAuthStateChanged } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-auth.js";
import { getFirestore, doc, addDoc, onSnapshot, collection, query, updateDoc, deleteDoc, serverTimestamp, orderBy, getDocs, limit, writeBatch } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-firestore.js";

const firebaseConfig = { apiKey: "AIzaSyAgDFeTSfSfYaUPoDLwmBuPUwmdtd9QuxM", authDomain: "checklist-c4255.firebaseapp.com", projectId: "checklist-c4255", storageBucket: "checklist-c4255.firebasestorage.app", messagingSenderId: "64837874160", appId: "1:64837874160:web:9fc4a131f16a33e9f3ed35" };
const appId = "checklist-app-main";
let db, auth, userId = null;
let allTasks = [], allSchools = [], allConsultants = [], editingTaskId = null, editingSchoolId = null, activePopoverTaskId = null, currentView = 'kanban', schoolSort = { key: 'school', dir: 'asc' };
let currentUserDisplay = sessionStorage.getItem('checklist_username') || '';
let pendingConfirmCallback = null;

// Prevent iOS Zoom on Gesture
document.addEventListener('gesturestart', function(e) {
    e.preventDefault();
});

// View State for Independent Filtering
const viewFilters = {
    kanban: { search: '', type: 'all', ec: 'all', status: 'all' },
    table: { search: '', type: 'all', ec: 'all', status: 'all' },
    zones: { search: '' }
};

// Legacy Color Map for migration
const LEGACY_COLORS = { "Sarah": "#A855F7", "Martin": "#3B82F6", "Krystle": "#EC4899", "Charlynn": "#10B981", "Sapnaa": "#6B7280", "Yuniza": "#EAB308" };

const KANBAN_STATUSES = ["In Progress", "For Submission", "Pending Award", "Awarded", "Not Awarded", "No Award", "Skipped"];
const PROGRESS_ITEMS = [{ id: 'proposal', label: 'Programme Proposal' }, { id: 'annex_b', label: 'Annex B - Price Proposal' }, { id: 'annex_f', label: 'Annex F - Instructor Deployment' }, { id: 'outline', label: 'Programme Outline' }, { id: 'track_record', label: 'Programme Track Record' }, { id: 'trainers_files', label: 'Trainers\' Files' }, { id: 'gebiz', label: 'GeBIZ Tracker Update' }, { id: 'email', label: 'Email Follow-up' }];

const formatDate = (dateStr) => { if (!dateStr) return '-'; try { return new Date(dateStr).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }); } catch (e) { return dateStr; } };
const getProgressState = (task) => { const p = task.progress || {}; const annexB = p.annex_b || (task.cost_val && task.cost_val.trim().length > 0); const firstSix = [p.proposal, annexB, p.annex_f, p.outline, p.track_record, p.trainers_files].every(Boolean); const bothLastTwo = [p.gebiz, p.email].every(Boolean); return { proposal: p.proposal, annex_b: annexB, annex_f: p.annex_f, outline: p.outline, track_record: p.track_record, trainers_files: p.trainers_files, gebiz: p.gebiz, email: p.email, allFirstSix: firstSix, bothLastTwo: bothLastTwo }; };

// Helper to get formatted assignment label with Brand
const getTaskLabel = (t) => {
    const brand = t.brand === 'Vivarch Enrichment' ? 'Vivarch' : (t.brand || '');
    if (!t.assignment && !brand) return 'Unassigned';
    if (!t.assignment) return brand;
    if (!brand) return t.assignment;
    return `${t.assignment}, ${brand}`;
};

// --- NEW COLOR LOGIC ---
function getConsultantStyles(name) {
    const consultant = allConsultants.find(c => c.name === name);
    if (!consultant || !consultant.active) {
        // Inactive or Unknown = Gray
        return {
            headerBg: '#9CA3AF', // Gray 400
            rowBg: '#F3F4F6',    // Gray 100
            hoverBg: 'hover:bg-gray-100',
            dotColor: '#D1D5DB'
        };
    }
    // Active = Custom Color
    return {
        headerBg: consultant.color,
        rowBg: `${consultant.color}40`, // 25% opacity hex (increased from 1A/10%)
        hoverBg: '', // Handled via inline for custom
        dotColor: consultant.color
    };
}

async function initApp() {
    const app = initializeApp(firebaseConfig); db = getFirestore(app); auth = getAuth(app);
    await signInAnonymously(auth);
    onAuthStateChanged(auth, (user) => {
        if (user) { 
            userId = user.uid; 
            document.getElementById('loading-overlay').classList.add('opacity-0', 'pointer-events-none'); 
            
            // CHECK USER NAME
            if(!currentUserDisplay) {
                document.getElementById('name-modal').classList.remove('hidden');
            } else {
                document.getElementById('user-badge').textContent = currentUserDisplay;
                initListeners(); 
            }
            
            // Force UI update to show filters for initial view (Kanban)
            toggleView(currentView);
        }
    });
}

function showToast(msg, type='success') {
    const toast = document.createElement('div'); const color = type === 'error' ? 'bg-red-500' : 'bg-dark';
    toast.className = `${color} text-white px-6 py-3 rounded-full shadow-2xl flex items-center gap-3 text-sm font-semibold pointer-events-auto toast-enter`;
    toast.innerHTML = type==='error' ? `<svg class="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>${msg}` : `<svg class="w-5 h-5 text-green-400" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7" /></svg>${msg}`;
    document.getElementById('toast-container').appendChild(toast);
    setTimeout(() => { toast.style.opacity = '0'; toast.style.transform = 'translate(-50%, 20px)'; setTimeout(()=>toast.remove(), 300); }, 3000);
}

// --- ACTIVITY LOGGING ---
async function logAction(message) {
    if(!currentUserDisplay) return;
    try {
        await addDoc(collection(db, `artifacts/${appId}/public/data/notifications`), {
            user: currentUserDisplay,
            message: message,
            timestamp: serverTimestamp()
        });
    } catch(e) { console.error("Log failed", e); }
}

function initListeners() {
    onSnapshot(query(collection(db, `artifacts/${appId}/public/data/tasks`)), (s) => { allTasks = s.docs.map(d => ({ id: d.id, ...d.data() })); refreshView(); });
    onSnapshot(query(collection(db, `artifacts/${appId}/public/data/schools`), orderBy('school')), (s) => { allSchools = s.docs.map(d => ({ id: d.id, ...d.data() })); populateSchoolDatalist(); if(currentView === 'zones') renderZonesTable(); });
    onSnapshot(query(collection(db, `artifacts/${appId}/public/data/consultants`), orderBy('name')), (s) => { 
        allConsultants = s.docs.map(d => ({ id: d.id, ...d.data() }));
        if(allConsultants.length === 0) seedConsultants();
        else { updateConsultantUI(); refreshView(); }
    });
    
    // Notifications Listener
    onSnapshot(query(collection(db, `artifacts/${appId}/public/data/notifications`), orderBy('timestamp', 'desc'), limit(50)), (s) => {
        const list = document.getElementById('notif-list');
        const badge = document.getElementById('notif-badge');
        list.innerHTML = '';
        if(s.empty) {
            list.innerHTML = '<div class="p-8 text-center text-gray-400 text-xs">No recent activity.</div>';
            badge.classList.add('hidden');
        } else {
            badge.classList.remove('hidden'); 
            
            s.docs.forEach(d => {
                const n = d.data();
                const time = n.timestamp ? new Date(n.timestamp.toDate()).toLocaleString([], { month: 'short', day: 'numeric', hour: '2-digit', minute:'2-digit'}) : 'Just now';
                const div = document.createElement('div');
                div.className = "px-4 py-3 border-b border-gray-50 hover:bg-gray-50 transition";
                div.innerHTML = `
                    <div class="flex justify-between items-start mb-0.5">
                        <span class="text-xs font-bold text-dark">${n.user}</span>
                        <span class="text-[10px] text-gray-400">${time}</span>
                    </div>
                    <p class="text-xs text-gray-600 leading-snug">${n.message}</p>
                `;
                list.appendChild(div);
            });
        }
    });
}

// --- NAME INPUT LOGIC ---
window.saveUserName = (e) => {
    e.preventDefault();
    const input = document.getElementById('user-name-input');
    const val = input.value.trim().toUpperCase();
    if(val) {
        currentUserDisplay = val;
        sessionStorage.setItem('checklist_username', val);
        document.getElementById('name-modal').classList.add('hidden');
        document.getElementById('user-badge').textContent = val;
        initListeners();
    }
};

window.toggleNotifications = () => {
    const pop = document.getElementById('notif-popover');
    if(pop.classList.contains('hidden')) {
        pop.classList.remove('hidden');
        // Hide badge when opened
        document.getElementById('notif-badge').classList.add('hidden');
    } else {
        pop.classList.add('hidden');
    }
};

async function seedConsultants() {
    const batch = writeBatch(db);
    const defaults = ["Sarah", "Martin", "Krystle", "Charlynn", "Sapnaa", "Yuniza"];
    defaults.forEach(name => {
        const ref = doc(collection(db, `artifacts/${appId}/public/data/consultants`));
        // Default active, map legacy colors
        batch.set(ref, { name, color: LEGACY_COLORS[name] || '#6B7280', active: true, createdAt: serverTimestamp() });
    });
    await batch.commit();
}

function updateConsultantUI() {
    // Populate Filters using Native Selects
    const populateSelect = (targetId, showAll) => {
        const select = document.getElementById(targetId);
        if(!select) return;
        
        const currentValue = select.value;
        select.innerHTML = '';
        
        // Add Default Option
        const defaultOpt = document.createElement('option');
        if(targetId === 'filter-ec' || targetId === 'mob-filter-ec') {
            defaultOpt.value = 'all'; defaultOpt.textContent = 'All Consultants';
        } else {
            defaultOpt.value = ''; defaultOpt.textContent = 'Unassigned';
        }
        select.appendChild(defaultOpt);

        // Sort: Active first, then by name
        const sorted = [...allConsultants].sort((a,b) => (a.active === b.active) ? a.name.localeCompare(b.name) : (a.active ? -1 : 1));

        sorted.forEach(c => {
            if (!showAll && !c.active) return;
            const opt = document.createElement('option');
            opt.value = c.name;
            opt.textContent = c.name;
            select.appendChild(opt);
        });
        
        if(currentValue) select.value = currentValue;
    };

    populateSelect('filter-ec', true); 
    populateSelect('mob-filter-ec', true); 
    populateSelect('sf_consultant', false); 

    // Manage Modal List - Split Active/Inactive
    const activeList = document.getElementById('consultant-list-active');
    const inactiveList = document.getElementById('consultant-list-inactive');
    activeList.innerHTML = '';
    inactiveList.innerHTML = '';

    allConsultants.forEach(c => {
        const row = document.createElement('tr');
        const isActive = c.active;
        row.innerHTML = `
            <td class="px-4 py-3 whitespace-nowrap text-sm font-medium text-gray-900">${c.name}</td>
            <td class="px-4 py-3 whitespace-nowrap">
                <div class="color-input-wrapper">
                    <input type="color" value="${c.color}" onchange="window.updateConsultantColor('${c.id}', this.value)" title="Change Color">
                </div>
            </td>
            <td class="px-4 py-3 whitespace-nowrap text-right">
                    <label class="relative inline-flex items-center cursor-pointer">
                    <input type="checkbox" value="" class="sr-only peer" ${isActive ? 'checked' : ''} onchange="window.toggleConsultantStatus('${c.id}', this.checked)">
                    <div class="w-9 h-5 bg-gray-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-ios_blue"></div>
                </label>
            </td>
        `;
        if(isActive) activeList.appendChild(row); else inactiveList.appendChild(row);
    });
}

// --- CONSULTANT ACTIONS ---
window.addConsultant = async () => {
    const inputName = document.getElementById('new-consultant-name');
    const inputColor = document.getElementById('new-consultant-color');
    const name = inputName.value.trim();
    if(!name) return;
    await addDoc(collection(db, `artifacts/${appId}/public/data/consultants`), { 
        name, 
        color: inputColor.value, 
        active: true, 
        createdAt: serverTimestamp() 
    });
    logAction(`Added new consultant: ${name}`);
    inputName.value = '';
    showToast("Consultant added");
}

window.toggleConsultantStatus = async (id, isActive) => {
    await updateDoc(doc(db, `artifacts/${appId}/public/data/consultants`, id), { active: isActive });
}

window.updateConsultantColor = async (id, color) => {
    await updateDoc(doc(db, `artifacts/${appId}/public/data/consultants`, id), { color: color });
}

window.openConsultantModal = () => document.getElementById('consultant-modal-overlay').classList.remove('hidden');


function populateSchoolDatalist() { const dl = document.getElementById('school-list'); dl.innerHTML = ''; allSchools.forEach(s => { const o = document.createElement('option'); o.value = s.school; dl.appendChild(o); }); }

// --- MOBILE FILTERS ---
window.toggleMobileFilters = () => {
    const sheet = document.getElementById('mobile-filter-sheet');
    const sheetContent = sheet.querySelector('div.transform');
    if (sheet.classList.contains('hidden')) {
        document.getElementById('mob-filter-type').value = document.getElementById('filter-type').value;
        document.getElementById('mob-filter-ec').value = document.getElementById('filter-ec').value;
        document.getElementById('mob-filter-status').value = document.getElementById('filter-status').value;
        sheet.classList.remove('hidden'); setTimeout(() => { sheetContent.classList.remove('translate-y-full'); }, 10);
    } else { sheetContent.classList.add('translate-y-full'); setTimeout(() => { sheet.classList.add('hidden'); }, 300); }
};

window.setMobileFilter = (key, value) => { 
    document.getElementById(`filter-${key}`).value = value;
    if(currentView === 'kanban' || currentView === 'table') {
        viewFilters[currentView][key] = value;
    }
    refreshView(); 
}

// --- VIEWS ---
window.refreshView = function() {
    if(currentView === 'kanban' || currentView === 'table') {
            viewFilters[currentView].search = document.getElementById('table-search').value;
            viewFilters[currentView].type = document.getElementById('filter-type').value;
            viewFilters[currentView].ec = document.getElementById('filter-ec').value;
            viewFilters[currentView].status = document.getElementById('filter-status').value;
    } else if (currentView === 'zones') {
            viewFilters.zones.search = document.getElementById('zones-search').value;
    }

    if(currentView === 'kanban') renderKanban();
    else if (currentView === 'table') renderTable();
    else if (currentView === 'zones') renderZonesTable();
}

function renderKanban() {
    const itqC = document.getElementById('kanban-area-itq'); const svpC = document.getElementById('kanban-area-svp');
    itqC.innerHTML = ''; svpC.innerHTML = '';
    const filtered = filterTasks(allTasks);
    KANBAN_STATUSES.forEach(status => {
        const tasks = filtered.filter(t => t.status === status).sort((a,b) => new Date(a.closing_date||'9999-12-31') - new Date(b.closing_date||'9999-12-31'));
        itqC.appendChild(createKanbanColumn(status, tasks.filter(t => t.type !== 'SVP')));
        svpC.appendChild(createKanbanColumn(status, tasks.filter(t => t.type === 'SVP')));
    });
}

function createKanbanColumn(status, tasks) {
    const col = document.createElement('div');
    col.className = 'kanban-column bg-white rounded-2xl shadow-ios border border-gray-100 flex flex-col h-[600px] flex-shrink-0 snap-center';
    col.innerHTML = `<div class="p-4 border-b border-gray-100 font-semibold text-gray-700 flex justify-between items-center ${status==='Skipped'?'bg-gray-200':'bg-gray-50'} rounded-t-2xl"><span class="text-sm uppercase tracking-wide">${status}</span><span class="bg-white border border-gray-200 text-gray-600 text-xs px-2 py-0.5 rounded-full font-bold">${tasks.length}</span></div><div class="kanban-column-content p-3 flex-grow overflow-y-auto space-y-3 bg-gray-50/30 custom-scrollbar" ondragover="window.allowDrop(event)" ondrop="window.drop(event, '${status}')"></div>`;
    const list = col.querySelector('.kanban-column-content');
    tasks.forEach(task => list.appendChild(createKanbanCard(task)));
    return col;
}

function createKanbanCard(task) {
    const ec = task.assignment ? task.assignment.split(',')[0].trim() : '';
    const styles = getConsultantStyles(ec);
    const ps = getProgressState(task);
    const segs = PROGRESS_ITEMS.map((item, i) => `<div class="progress-segment ${i<6?(ps.allFirstSix && ps[item.id]?'active-solid':(ps[item.id]?'active-light':'')):(ps.bothLastTwo && ps[item.id]?'active-solid':(ps[item.id]?'active-light':''))}"></div>`).join('');
    const moeSuffix = task.moe_code && task.moe_code.length > 4 ? task.moe_code.slice(-4) : task.moe_code;
    const assignmentDisplay = getTaskLabel(task);

    const card = document.createElement('div');
    card.className = `kanban-card bg-white rounded-xl shadow-sm border border-gray-100 cursor-grab hover:shadow-ios-hover transition-all transform hover:-translate-y-1 overflow-hidden relative`;
    card.setAttribute('draggable', 'true'); card.ondragstart = (e) => e.dataTransfer.setData("text", task.id);
    
    card.innerHTML = `
        <div class="kanban-card-header text-white px-3 py-2 flex justify-between items-center" style="background-color: ${styles.headerBg}" onclick="window.openModalById('${task.id}')">
            <div class="flex items-center gap-2">
                <span class="text-[10px] font-bold uppercase tracking-wider bg-black/20 px-1.5 py-0.5 rounded">${task.type}</span>
                ${moeSuffix ? `<span class="text-[10px] font-medium opacity-90 tracking-wide border-l border-white/30 pl-2">${moeSuffix}</span>` : ''}
            </div>
            <span class="text-xs font-medium">${formatDate(task.closing_date)}</span>
        </div>
        <div class="p-3">
            <h4 class="text-sm font-bold text-gray-800 leading-snug mb-1 line-clamp-2 cursor-pointer hover:text-primary" onclick="window.openModalById('${task.id}')">${task.school}</h4>
            <div class="text-xs text-gray-500 mb-3 line-clamp-1">${task.programme}</div>
            <div class="flex justify-between items-center pt-2 border-t border-gray-50">
                    <div class="flex gap-1">
                        <a href="${task.specs}" target="_blank" class="text-gray-400 hover:text-ios_blue p-1" onclick="event.stopPropagation()"><svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path></svg></a>
                        <a href="${task.folder}" target="_blank" class="text-gray-400 hover:text-accent p-1" onclick="event.stopPropagation()"><svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 7v10a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-6l-2-2H5a2 2 0 00-2 2z"></path></svg></a>
                        <button class="text-gray-400 hover:text-green-600 p-1" onclick="event.stopPropagation(); window.toggleProgress(event, '${task.id}')"><svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01"></path></svg></button>
                    </div>
                <span class="text-[10px] font-medium px-2 py-0.5 rounded bg-blue-50 text-blue-700 whitespace-nowrap">${assignmentDisplay}</span>
            </div>
            <div class="flex gap-0.5 mt-2 h-1 w-full opacity-80">${segs}</div>
        </div>`;
    return card;
}

function renderTable() {
    const tbody = document.getElementById('table-body'); tbody.innerHTML = '';
    
    let tasks = filterTasks(allTasks);
    
    tasks.sort((a, b) => {
        const statusOrder = KANBAN_STATUSES.indexOf(a.status) - KANBAN_STATUSES.indexOf(b.status);
        if (statusOrder !== 0) return statusOrder;
        
        const dateA = a.closing_date ? new Date(a.closing_date) : new Date('9999-12-31');
        const dateB = b.closing_date ? new Date(b.closing_date) : new Date('9999-12-31');
        return dateA - dateB;
    });

    tasks.forEach(task => {
        const ec = task.assignment ? task.assignment.split(',')[0].trim() : '';
        const styles = getConsultantStyles(ec);
        const ps = getProgressState(task);
        const segs = PROGRESS_ITEMS.map((item, i) => `<div class="progress-segment ${i<6?(ps.allFirstSix && ps[item.id]?'active-solid':(ps[item.id]?'active-light':'')):(ps.bothLastTwo && ps[item.id]?'active-solid':(ps[item.id]?'active-light':''))}"></div>`).join('');
        const assignmentDisplay = getTaskLabel(task);
        const tr = document.createElement('tr');
        tr.className = `cursor-pointer border-b border-gray-100 transition-all duration-150`;
        tr.style.backgroundColor = styles.rowBg;
        
        tr.onmouseenter = () => { 
            tr.style.filter = "brightness(0.92) saturate(1.05)"; 
            tr.style.zIndex = "10";
            tr.style.boxShadow = "0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03)";
            tr.style.transform = "scale(1.002)";
        }
        tr.onmouseleave = () => { 
            tr.style.filter = "none";
            tr.style.zIndex = "auto";
            tr.style.boxShadow = "none";
            tr.style.transform = "none";
        }

        tr.onclick = (e) => { if(!e.target.closest('button') && !e.target.closest('a')) window.openModalById(task.id); };
        tr.innerHTML = `
            <td class="px-4 sm:px-6 py-4 text-sm font-bold text-gray-700 align-middle">${task.type}</td>
            <td class="px-4 sm:px-6 py-4 align-middle"><div class="text-sm font-bold text-gray-900">${task.school}</div><div class="text-xs text-gray-600">${task.programme}</div><div class="text-xs text-gray-400 font-mono">${task.moe_code||''}</div></td>
            <td class="px-4 sm:px-6 py-4 text-sm text-gray-500 font-mono align-middle">${formatDate(task.closing_date)}</td>
            <td class="px-4 sm:px-6 py-4 text-sm text-gray-600 align-middle whitespace-nowrap"><span class="bg-white border border-gray-200 px-2 py-1 rounded-md text-xs shadow-sm">${assignmentDisplay}</span></td>
            <td class="px-4 sm:px-6 py-4 align-middle whitespace-nowrap"><div class="flex flex-col gap-1.5"><span class="px-2.5 py-0.5 text-xs font-bold rounded-full w-fit ${task.status==='In Progress'?'bg-green-100 text-green-800':'bg-gray-100 text-gray-600'}">${task.status}</span><div class="flex gap-0.5 h-1 w-full max-w-[120px] opacity-80">${segs}</div></div></td>
            <td class="px-4 sm:px-6 py-4 text-sm text-gray-500 align-middle">${task.costing_type||'-'}</td>
            <td class="px-4 sm:px-6 py-4 text-right align-middle"><div class="flex gap-2 justify-end"><a href="${task.specs}" target="_blank" class="text-gray-400 hover:text-ios_blue p-1" onclick="event.stopPropagation()"><svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path></svg></a><a href="${task.folder}" target="_blank" class="text-gray-400 hover:text-accent p-1" onclick="event.stopPropagation()"><svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 7v10a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-6l-2-2H5a2 2 0 00-2 2z"></path></svg></a><button class="text-gray-400 hover:text-green-600 p-1" onclick="event.stopPropagation(); window.toggleProgress(event, '${task.id}')"><svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01"></path></svg></button></div></td>`;
        tbody.appendChild(tr);
    });
}

window.sortSchools = (key) => {
        ['school', 'zone', 'consultant'].forEach(k => document.getElementById(`sort-icon-${k}`).classList.replace('text-ios_blue', 'text-gray-300'));
        if (schoolSort.key === key) schoolSort.dir = schoolSort.dir === 'asc' ? 'desc' : 'asc'; else { schoolSort.key = key; schoolSort.dir = 'asc'; }
        document.getElementById(`sort-icon-${key}`).classList.replace('text-gray-300', 'text-ios_blue'); renderZonesTable();
}

function renderZonesTable() {
    const tbody = document.getElementById('zones-body'); tbody.innerHTML = '';
    document.getElementById('zones-count').textContent = `${allSchools.length} Schools`;
    const searchInput = document.getElementById('zones-search');
    const search = searchInput.value.toLowerCase();
    
    // Toggle Clear Button
    const clearBtn = document.getElementById('zones-search-clear');
    if(search.length > 0) clearBtn.classList.remove('hidden'); else clearBtn.classList.add('hidden');

    let filtered = allSchools.filter(s => !search || s.school.toLowerCase().includes(search) || s.zone.toLowerCase().includes(search) || s.consultant.toLowerCase().includes(search));
    filtered.sort((a, b) => { const valA = (a[schoolSort.key]||'').toLowerCase(); const valB = (b[schoolSort.key]||'').toLowerCase(); return valA < valB ? (schoolSort.dir==='asc'?-1:1) : (valA > valB ? (schoolSort.dir==='asc'?1:-1) : 0); });
    filtered.forEach((s, idx) => {
        const styles = getConsultantStyles(s.consultant);
        const tr = document.createElement('tr'); tr.className = "border-b border-gray-50 transition-colors group zone-row cursor-pointer";
        tr.onclick = () => window.openSchoolModal(s.id); 
        tr.innerHTML = `<td class="px-6 py-3 text-xs text-gray-400 font-mono align-middle">${idx+1}</td><td class="px-6 py-3 text-sm font-bold text-dark align-middle">${s.school}</td><td class="px-6 py-3 align-middle"><span class="px-2 py-1 rounded bg-gray-100 text-xs font-medium text-gray-600 border border-gray-200">${s.zone}</span></td><td class="px-6 py-3 text-sm text-gray-600 align-middle flex items-center gap-2"><span class="w-2 h-2 rounded-full" style="background-color:${styles.dotColor}"></span>${s.consultant || 'Unassigned'}</td>`;
        tbody.appendChild(tr);
    });
}
document.getElementById('zones-search').addEventListener('input', renderZonesTable);
window.clearZonesSearch = () => {
    document.getElementById('zones-search').value = '';
    renderZonesTable();
}

function filterTasks(tasks) {
    const search = document.getElementById('table-search').value.toLowerCase();
    const type = document.getElementById('filter-type').value; 
    const ec = document.getElementById('filter-ec').value; 
    const status = document.getElementById('filter-status').value; 
    
    return tasks.filter(t => {
        const searchableItems = [
            t.school, t.programme, t.type, t.moe_code,
            t.contact1?.name, t.contact1?.email, t.contact1?.cont,
            t.contact2?.name, t.contact2?.email, t.contact2?.cont,
            t.contact3?.name, t.contact3?.email, t.contact3?.cont
        ];
        
        const fullText = searchableItems.filter(Boolean).join(' ').toLowerCase();

        if (search && !fullText.includes(search)) return false;
        if (type !== 'all' && t.type !== type) return false;
        if (status !== 'all' && t.status !== status) return false;
        if (ec !== 'all' && (!t.assignment || !t.assignment.includes(ec))) return false;
        return true;
    });
}
document.getElementById('table-search').addEventListener('input', refreshView);
document.getElementById('filter-type').addEventListener('change', refreshView);
document.getElementById('filter-ec').addEventListener('change', refreshView);
document.getElementById('filter-status').addEventListener('change', refreshView);

// --- NEW MOE FIELD LOGIC ---
const updateMoeFieldState = (type) => {
    const group = document.getElementById('moe-field-group');
    const input = document.getElementById('f_moe_code');
    const star = document.getElementById('moe-star');

    group.classList.remove('hidden');
    input.required = false;
    input.disabled = false;
    star.classList.add('hidden');
    input.placeholder = "e.g. ITQ-2025-001";

    if (type === 'ITQ') {
        input.required = true;
        star.classList.remove('hidden');
    } else if (type === 'SVP') {
        group.classList.add('hidden');
        input.disabled = true; 
        input.value = ''; 
    } else if (type === 'Tender') {
        input.placeholder = "Optional";
    }
};
document.getElementById('f_type').addEventListener('change', (e) => updateMoeFieldState(e.target.value));


window.openModal = (task=null) => {
    editingTaskId = task ? task.id : null; document.getElementById('task-form').reset(); document.getElementById('modal-title').textContent = task ? 'Edit Task' : 'New Task'; document.getElementById('delete-btn').classList.toggle('hidden', !task);
    
    const typeVal = task?.type || 'ITQ';
    document.getElementById('f_type').value = typeVal;
    document.getElementById('f_brand').value = task?.brand || '';
    document.getElementById('f_costing_type').value = task?.costing_type || '';
    document.getElementById('f_status').value = task?.status || 'In Progress';

    updateMoeFieldState(typeVal);

    // --- POPULATE CONTACTS SPREADSHEET ---
    const tbody = document.getElementById('contacts-body');
    tbody.innerHTML = '';
    // Create 3 empty rows by default
    for(let i=1; i<=3; i++) {
        const c = task ? (task[`contact${i}`] || {}) : {};
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td contenteditable="true" class="spreadsheet-cell p-2 border-b border-r border-gray-100 outline-none">${c.name||''}</td>
            <td contenteditable="true" class="spreadsheet-cell p-2 border-b border-r border-gray-100 outline-none">${c.des||''}</td>
            <td contenteditable="true" class="spreadsheet-cell p-2 border-b border-r border-gray-100 outline-none">${c.dept||''}</td>
            <td contenteditable="true" class="spreadsheet-cell p-2 border-b border-r border-gray-100 outline-none">${c.cont||''}</td>
            <td contenteditable="true" class="spreadsheet-cell p-2 border-b border-gray-100 outline-none">${c.email||''}</td>
        `;
        tbody.appendChild(tr);
    }

    if(task) {
        ['f_moe_code','f_closing_date','f_school','f_assignment','f_programme','f_specs','f_folder','f_trainers','f_cost_specs','f_cost_val','f_notes','f_appointment'].forEach(id => { const el = document.getElementById(id); if(el) el.value = task[id.replace('f_','')]||''; });
        document.getElementById('costing-details').classList.toggle('hidden', !task.costing_type);
    } 
    document.getElementById('modal-overlay').classList.remove('hidden');
}

// --- SPREADSHEET LOGIC ---
const table = document.getElementById('contacts-table');

// Arrow Navigation
table.addEventListener('keydown', (e) => {
    if (!e.target.classList.contains('spreadsheet-cell')) return;
    
    // Allow default behavior for text editing if not an arrow key
    if (!['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) return;

    const cell = e.target;
    const row = cell.parentElement;
    const cellIndex = Array.from(row.children).indexOf(cell);
    const rowIndex = Array.from(row.parentElement.children).indexOf(row);
    
    let nextCell;

    // Logic: If caret is at start/end, or generic navigation
    // Simplified: Arrows always navigate cells to mimic Excel navigation feel
    e.preventDefault(); 

    if (e.key === 'ArrowRight') {
        nextCell = row.children[cellIndex + 1];
    } else if (e.key === 'ArrowLeft') {
        nextCell = row.children[cellIndex - 1];
    } else if (e.key === 'ArrowDown') {
        const nextRow = row.parentElement.children[rowIndex + 1];
        if (nextRow) nextCell = nextRow.children[cellIndex];
    } else if (e.key === 'ArrowUp') {
        const prevRow = row.parentElement.children[rowIndex - 1];
        if (prevRow) nextCell = prevRow.children[cellIndex];
    }

    if (nextCell) {
        nextCell.focus();
        // Optional: Select text when focusing via keyboard like Excel F2? No, just cursor placement.
        // For better UX, we could select all text so user can overwrite.
        // document.execCommand('selectAll', false, null); 
    }
});

// Paste Support (Paste rows from Excel)
table.addEventListener('paste', (e) => {
    if (!e.target.classList.contains('spreadsheet-cell')) return;
    e.preventDefault();
    const clipboardData = (e.clipboardData || window.clipboardData).getData('text');
    const rows = clipboardData.split(/\r\n|\n|\r/);
    
    let currentRow = e.target.parentElement;
    const startCellIndex = Array.from(currentRow.children).indexOf(e.target);

    rows.forEach((rowData) => {
        if (!rowData.trim() && rows.length > 1) return; // Skip empty trailing lines
        if (!currentRow) return;

        const cells = rowData.split('\t');
        let currentCell = currentRow.children[startCellIndex];

        cells.forEach((cellData) => {
            if (currentCell) {
                currentCell.textContent = cellData.trim();
                currentCell = currentCell.nextElementSibling;
            }
        });
        
        currentRow = currentRow.nextElementSibling;
    });
});

// Copy Support (Copy to Excel)
table.addEventListener('copy', (e) => {
    if (!e.target.classList.contains('spreadsheet-cell') && !table.contains(e.target)) return;
    const selection = window.getSelection();
    if (selection.isCollapsed) return;

    e.preventDefault();
    
    let rowsData = [];
    const rows = table.querySelectorAll('tbody tr');
    
    rows.forEach(tr => {
        let rowCells = [];
        let hasSelected = false;
        tr.querySelectorAll('td').forEach(td => {
            if (selection.containsNode(td, true)) {
                hasSelected = true;
                rowCells.push(td.innerText);
            }
        });
        if (hasSelected) rowsData.push(rowCells.join('\t'));
    });
    
    const clipboardText = rowsData.length > 0 ? rowsData.join('\n') : selection.toString();
    e.clipboardData.setData('text/plain', clipboardText);
});


window.openModalById = (id) => { if(!window.getSelection().toString()) window.openModal(allTasks.find(t=>t.id===id)); };
document.getElementById('close-modal-btn').addEventListener('click', () => document.getElementById('modal-overlay').classList.add('hidden'));
document.getElementById('cancel-btn').addEventListener('click', () => document.getElementById('modal-overlay').classList.add('hidden'));
document.getElementById('f_school').addEventListener('change', (e) => { const found = allSchools.find(s => s.school === e.target.value); document.getElementById('f_assignment').value = found ? `${found.consultant}, ${found.zone}` : ''; });
window.toggleCostingDetails = (val) => document.getElementById('costing-details').classList.toggle('hidden', !val);

// CONFIRMATION DIALOG LOGIC
window.showConfirm = (title, msg, callback) => {
    const modal = document.getElementById('confirm-modal');
    const content = document.getElementById('confirm-modal-content');
    document.getElementById('confirm-title').textContent = title;
    document.getElementById('confirm-msg').textContent = msg;
    pendingConfirmCallback = callback;
    
    modal.classList.remove('hidden');
    // Small delay to allow transition
    requestAnimationFrame(() => {
        modal.classList.remove('opacity-0');
        content.classList.remove('scale-95');
        content.classList.add('scale-100');
    });
};

window.closeConfirm = () => {
    const modal = document.getElementById('confirm-modal');
    const content = document.getElementById('confirm-modal-content');
    modal.classList.add('opacity-0');
    content.classList.remove('scale-100');
    content.classList.add('scale-95');
    setTimeout(() => {
        modal.classList.add('hidden');
        pendingConfirmCallback = null;
    }, 200);
};

document.getElementById('confirm-yes-btn').addEventListener('click', () => {
    if (pendingConfirmCallback) pendingConfirmCallback();
    window.closeConfirm();
});

document.getElementById('save-btn').addEventListener('click', async () => {
    const form = document.getElementById('task-form'); if(!form.checkValidity()) { form.reportValidity(); return; }
    const getData = (id) => document.getElementById(id).value;
    
    // Scrape Contacts Table
    const contactData = {};
    const tbody = document.getElementById('contacts-body');
    Array.from(tbody.children).forEach((tr, i) => {
        const cells = tr.children;
        // Mapping: 0:Name, 1:Des, 2:Dept, 3:Cont, 4:Email
        contactData[`contact${i+1}`] = {
            name: cells[0].textContent.trim(),
            des: cells[1].textContent.trim(),
            dept: cells[2].textContent.trim(),
            cont: cells[3].textContent.trim(),
            email: cells[4].textContent.trim()
        };
    });

    const data = {
        type: getData('f_type'), moe_code: getData('f_moe_code'), closing_date: getData('f_closing_date'), school: getData('f_school'), assignment: getData('f_assignment'), programme: getData('f_programme'), brand: getData('f_brand'), specs: getData('f_specs'), folder: getData('f_folder'), costing_type: getData('f_costing_type'), cost_specs: getData('f_cost_specs'), cost_val: getData('f_cost_val'), trainers: getData('f_trainers'), notes: getData('f_notes'), appointment: getData('f_appointment'), status: getData('f_status'), updatedAt: serverTimestamp(),
        ...contactData
    };
    
    const schoolName = getData('f_school');
    
    try { 
        if(editingTaskId) { 
            await updateDoc(doc(collection(db, `artifacts/${appId}/public/data/tasks`), editingTaskId), data); 
            logAction(`Updated task for ${schoolName}`);
            showToast("Task updated"); 
        } else { 
            data.createdAt = serverTimestamp(); 
            await addDoc(collection(db, `artifacts/${appId}/public/data/tasks`), data); 
            logAction(`Created new task for ${schoolName}`);
            showToast("Task created"); 
        } 
        document.getElementById('modal-overlay').classList.add('hidden'); 
    } catch(e) { console.error(e); showToast("Error saving", 'error'); }
});

// UPDATED TASK DELETE
document.getElementById('delete-btn').addEventListener('click', () => {
    const task = allTasks.find(t => t.id === editingTaskId); // Get task for logging
    const schoolName = task ? task.school : 'Unknown';
    
    window.showConfirm("Delete Task?", "This action cannot be undone.", async () => {
        if(editingTaskId) {
            await deleteDoc(doc(collection(db, `artifacts/${appId}/public/data/tasks`), editingTaskId));
            logAction(`Deleted task: ${schoolName}`);
            showToast("Task deleted");
            document.getElementById('modal-overlay').classList.add('hidden');
        }
    });
});

window.openSchoolModal = (id) => {
    editingSchoolId = id; document.getElementById('school-form').reset(); document.getElementById('school-delete-btn').classList.toggle('hidden', !id); document.getElementById('school-modal-title').textContent = id ? 'Edit School' : 'Add School';
    if(id) { 
        const s = allSchools.find(x => x.id === id); 
        if(s) { 
            document.getElementById('sf_school').value = s.school; 
            document.getElementById('sf_zone').value = s.zone || 'North'; 
            document.getElementById('sf_consultant').value = s.consultant || ''; 
        } 
    } else {
            document.getElementById('sf_zone').value = 'North';
            document.getElementById('sf_consultant').value = '';
    }
    document.getElementById('school-modal-overlay').classList.remove('hidden');
}
document.getElementById('school-save-btn').addEventListener('click', async () => {
    const form = document.getElementById('school-form'); if(!form.checkValidity()) { form.reportValidity(); return; }
    const sName = document.getElementById('sf_school').value;
    const data = { school: sName, zone: document.getElementById('sf_zone').value, consultant: document.getElementById('sf_consultant').value };
    try { 
        if(editingSchoolId) { 
            await updateDoc(doc(collection(db, `artifacts/${appId}/public/data/schools`), editingSchoolId), data); 
            logAction(`Updated school details: ${sName}`);
            showToast("School updated"); 
        } else { 
            await addDoc(collection(db, `artifacts/${appId}/public/data/schools`), data); 
            logAction(`Added new school: ${sName}`);
            showToast("School added"); 
        } 
        document.getElementById('school-modal-overlay').classList.add('hidden'); 
    } catch(e) { console.error(e); showToast("Error saving school", 'error'); }
});

// UPDATED SCHOOL DELETE
document.getElementById('school-delete-btn').addEventListener('click', () => {
    window.showConfirm("Delete School?", "This will remove the school entry.", async () => {
        if(editingSchoolId) {
            await deleteDoc(doc(collection(db, `artifacts/${appId}/public/data/schools`), editingSchoolId));
            logAction("Deleted a school entry");
            showToast("School deleted");
            document.getElementById('school-modal-overlay').classList.add('hidden');
        }
    });
});

function toggleView(view) {
    if(currentView === 'kanban' || currentView === 'table') {
            viewFilters[currentView].search = document.getElementById('table-search').value;
            viewFilters[currentView].type = document.getElementById('filter-type').value;
            viewFilters[currentView].ec = document.getElementById('filter-ec').value;
            viewFilters[currentView].status = document.getElementById('filter-status').value;
    }

    currentView = view;
    
    if(view === 'kanban' || view === 'table') {
        document.getElementById('table-search').value = viewFilters[view].search;
        document.getElementById('filter-type').value = viewFilters[view].type;
        document.getElementById('filter-ec').value = viewFilters[view].ec;
        document.getElementById('filter-status').value = viewFilters[view].status;
    }

    const kb = document.getElementById('kanban-view'); const tb = document.getElementById('table-view'); const zv = document.getElementById('zones-view');
    const tl = document.getElementById('toolbar'); const kt = document.getElementById('kanban-toolbar'); const zt = document.getElementById('zones-toolbar');
    ['view-kanban','view-table','view-zones'].forEach(id => { const btn = document.getElementById(id); if (view === id.replace('view-','')) { btn.classList.add('bg-white', 'shadow-sm', 'text-dark', 'font-semibold'); btn.classList.remove('text-gray-500'); } else { btn.classList.remove('bg-white', 'shadow-sm', 'text-dark', 'font-semibold'); btn.classList.add('text-gray-500'); } });
    kb.classList.add('hidden'); tb.classList.add('hidden'); zv.classList.add('hidden'); tl.classList.add('hidden'); kt.classList.add('hidden'); zt.classList.add('hidden');
    
    if(view === 'kanban') { kb.classList.remove('hidden'); kt.classList.remove('hidden'); tl.classList.remove('hidden'); } 
    else if(view === 'table') { tb.classList.remove('hidden'); tl.classList.remove('hidden'); } 
    else if(view === 'zones') { zv.classList.remove('hidden'); zt.classList.remove('hidden'); }
    
    refreshView();
}
document.getElementById('view-kanban').addEventListener('click', () => toggleView('kanban')); document.getElementById('view-table').addEventListener('click', () => toggleView('table')); document.getElementById('view-zones').addEventListener('click', () => toggleView('zones'));

window.toggleProgress = (e, taskId) => {
    const p = document.getElementById('progress-popover'); if(!p.classList.contains('hidden') && activePopoverTaskId === taskId) { window.closeProgressPopover(); return; }
    activePopoverTaskId = taskId; const task = allTasks.find(t=>t.id===taskId); const list = document.getElementById('progress-list'); list.innerHTML = '';
    PROGRESS_ITEMS.forEach(item => { const checked = task.progress?.[item.id] || (item.id==='annex_b' && task.cost_val?.length>0); list.innerHTML += `<label class="flex items-center space-x-3 cursor-pointer"><input type="checkbox" class="form-checkbox h-4 w-4 text-ios_blue rounded" ${checked?'checked':''} onchange="window.updateProgress('${taskId}','${item.id}',this.checked)"><span class="text-xs font-medium text-gray-700">${item.label}</span></label>`; });
    p.classList.remove('hidden'); 
    
    const rect = e.target.getBoundingClientRect(); 
    const pWidth = 288;
    const pHeight = 350;
    
    let top = rect.bottom + window.scrollY + 5;
    let left = rect.left;

    if (left + pWidth > window.innerWidth) {
        left = window.innerWidth - pWidth - 20; 
    }
    if (top + pHeight > window.innerHeight + window.scrollY) {
        top = rect.top + window.scrollY - pHeight + 50;
    }

    p.style.top = `${top}px`; 
    p.style.left = `${left}px`;
};
window.closeProgressPopover = () => document.getElementById('progress-popover').classList.add('hidden');
window.updateProgress = async (id, key, val) => { 
    const t = allTasks.find(x=>x.id===id); 
    const p = {...(t.progress||{})}; 
    p[key]=val; 
    await updateDoc(doc(collection(db, `artifacts/${appId}/public/data/tasks`), id), {progress:p}); 
    
    const itemName = PROGRESS_ITEMS.find(i=>i.id===key)?.label || key;
    logAction(`Updated progress for ${t.school}: ${val ? 'Completed' : 'Unchecked'} ${itemName}`);
    
    showToast("Progress updated"); 
};
window.allowDrop = (e) => e.preventDefault(); 
window.drop = async (e, status) => { 
    e.preventDefault(); 
    const id = e.dataTransfer.getData("text"); 
    if(id) {
        const t = allTasks.find(x=>x.id===id);
        await updateDoc(doc(collection(db, `artifacts/${appId}/public/data/tasks`), id), { status }); 
        logAction(`Moved ${t.school} to ${status}`);
        showToast(`Moved to ${status}`); 
    }
};

// --- FAB & MENU LOGIC ---
window.toggleFabMenu = () => {
    const menu = document.getElementById('fab-menu');
    const icon = document.getElementById('fab-icon');
    
    if (menu.classList.contains('opacity-0')) {
        // Open
        menu.classList.remove('opacity-0', 'translate-y-4', 'pointer-events-none');
        icon.classList.add('rotate-45');
    } else {
        // Close
        menu.classList.add('opacity-0', 'translate-y-4', 'pointer-events-none');
        icon.classList.remove('rotate-45');
    }
};

// --- EXPORT CSV LOGIC ---
window.exportCSV = () => {
    const escapeCsv = (val) => {
        if (val === null || val === undefined) return '';
        const str = String(val);
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
            return `"${str.replace(/"/g, '""')}"`;
        }
        return str;
    };

    const headers = [
        "Export Date", "Item Code", "Closing Date", "Zone", "Consultant", "School", 
        "Contact Person 1 Name", "Contact Person 1 Designation", "Contact Person 1 Department", "Contact Person 1 Contact Number", "Contact Person 1 Email", 
        "Contact Person 2 Name", "Contact Person 2 Designation", "Contact Person 2 Department", "Contact Person 2 Contact Number", "Contact Person 2 Email", 
        "Programme Name", "Brand", "Status"
    ];

    const today = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });

    // Map tasks to rows
    const rows = allTasks.map(t => {
        // Find detailed school info for Zone and Consultant separation
        const schoolData = allSchools.find(s => s.school === t.school) || {};
        
        return [
            today,
            t.moe_code,
            t.closing_date,
            schoolData.zone || '',
            schoolData.consultant || '', // Use the consultant from Schools DB for consistency, or parse t.assignment
            t.school,
            // Contact 1
            t.contact1?.name, t.contact1?.des, t.contact1?.dept, t.contact1?.cont, t.contact1?.email,
            // Contact 2
            t.contact2?.name, t.contact2?.des, t.contact2?.dept, t.contact2?.cont, t.contact2?.email,
            t.programme,
            t.brand,
            t.status
        ].map(escapeCsv).join(',');
    });

    const csvContent = [headers.join(','), ...rows].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `checklist_export_${new Date().toISOString().slice(0,10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    // Close FAB menu after export
    window.toggleFabMenu();
};

// AUTO HIDE LOGIC FOR UI ELEMENTS
document.addEventListener('click', (e) => {
    // FAB Menu
    const fabMenu = document.getElementById('fab-menu');
    const fabBtn = document.getElementById('main-fab');
    if (!fabBtn.contains(e.target) && !fabMenu.contains(e.target) && !fabMenu.classList.contains('opacity-0')) {
        window.toggleFabMenu();
    }

    // Notification Popover
    const notifPop = document.getElementById('notif-popover');
    const notifBtn = document.getElementById('notif-btn');
    if (!notifBtn.contains(e.target) && !notifPop.contains(e.target) && !notifPop.classList.contains('hidden')) {
        window.toggleNotifications();
    }

    // Progress Popover
    const progPop = document.getElementById('progress-popover');
    // Ensure we don't close when clicking the trigger button itself (dynamically created)
    if (!progPop.classList.contains('hidden') && !progPop.contains(e.target) && !e.target.closest('button[onclick^="window.toggleProgress"]')) {
         window.closeProgressPopover();
    }
});

initApp();