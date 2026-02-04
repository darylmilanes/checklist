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
let allTasks = [], allSchools = [], allConsultants = [], editingTaskId = null, editingSchoolId = null, activePopoverTaskId = null, currentView = 'kanban';
// UPDATED: Default sort by Zone first
let schoolSort = { key: 'zone', dir: 'asc' };
let currentUserDisplay = sessionStorage.getItem('checklist_username') || '';
let pendingConfirmCallback = null;
let isDragging = false;
let selectionStart = null; // {r, c}
let selectionEnd = null;   // {r, c}

// Prevent iOS Zoom on Gesture
document.addEventListener('gesturestart', function(e) { e.preventDefault(); });

// View State for Independent Filtering
const viewFilters = {
    kanban: { search: '', type: 'all', ec: 'all', status: 'all' },
    table: { search: '', type: 'all', ec: 'all', status: 'all' },
    zones: { search: '' } 
};

// Legacy Color Map & Zone Colors
const LEGACY_COLORS = { "Sarah": "#A855F7", "Martin": "#3B82F6", "Krystle": "#EC4899", "Charlynn": "#10B981", "Sapnaa": "#6B7280", "Yuniza": "#EAB308" };
// CSS classes mapped to Zone names
const ZONE_COLORS = { "East": "bg-zone-east", "North": "bg-zone-north", "South": "bg-zone-south", "West": "bg-zone-west" };

const KANBAN_STATUSES = ["In Progress", "For Submission", "Pending Award", "Awarded", "Not Awarded", "No Award", "Skipped"];
const PROGRESS_ITEMS = [{ id: 'proposal', label: 'Programme Proposal' }, { id: 'annex_b', label: 'Annex B - Price Proposal' }, { id: 'annex_f', label: 'Annex F - Instructor Deployment' }, { id: 'outline', label: 'Programme Outline' }, { id: 'track_record', label: 'Programme Track Record' }, { id: 'trainers_files', label: 'Trainers\' Files' }, { id: 'gebiz', label: 'GeBIZ Tracker Update' }, { id: 'email', label: 'Email Follow-up' }];

// UPDATED: Added weekday 'short' (e.g. Mon, 25 Dec 2025)
const formatDate = (dateStr) => { if (!dateStr) return '-'; try { return new Date(dateStr).toLocaleDateString('en-GB', { weekday: 'short', day: '2-digit', month: 'short', year: 'numeric' }); } catch (e) { return dateStr; } };
const getProgressState = (task) => { const p = task.progress || {}; const annexB = p.annex_b || (task.cost_val && task.cost_val.trim().length > 0); const firstSix = [p.proposal, annexB, p.annex_f, p.outline, p.track_record, p.trainers_files].every(Boolean); const bothLastTwo = [p.gebiz, p.email].every(Boolean); return { proposal: p.proposal, annex_b: annexB, annex_f: p.annex_f, outline: p.outline, track_record: p.track_record, trainers_files: p.trainers_files, gebiz: p.gebiz, email: p.email, allFirstSix: firstSix, bothLastTwo: bothLastTwo }; };
const getTaskLabel = (t) => { const brand = t.brand === 'Vivarch Enrichment' ? 'Vivarch' : (t.brand || ''); if (!t.assignment && !brand) return 'Unassigned'; if (!t.assignment) return brand; if (!brand) return t.assignment; return `${t.assignment}, ${brand}`; };

// Helper to make links clickable pills
const linkify = (text) => {
    if (!text) return '';
    // Simple Regex for URLs (http/https/ftp)
    const urlRegex = /(\b(https?|ftp|file):\/\/[-A-Z0-9+&@#\/%?=~_|!:,.;]*[-A-Z0-9+&@#\/%=~_|])/ig;
    return text.replace(urlRegex, (url) => {
        return `<a href="${url}" target="_blank" class="inline-block px-2 py-0.5 rounded-full bg-blue-100 text-blue-600 text-xs hover:bg-blue-200 transition-colors break-all align-middle mt-1 mb-1" onclick="event.stopPropagation()">${url}</a>`;
    });
}

// --- HELPER: ROBUST NAME NORMALIZATION ---
// Removes extra spaces, invisible chars, and trims.
const cleanStr = (str) => str ? str.toString().trim().replace(/\s+/g, ' ').replace(/[\u200B-\u200D\uFEFF]/g, '') : '';

// Generates a strict comparison key (lowercase, alpha-numeric only to prevent duplicates like "Krystle" vs "Krystle.")
const normalizeKey = (str) => {
    const s = cleanStr(str).toLowerCase();
    // Replace anything that is not a letter or number (e.g. remove spaces, dots, hyphens for key comparison)
    return s.replace(/[^a-z0-9]/g, '');
};

function getConsultantStyles(name) {
    const key = normalizeKey(name);
    // Find consultant by matching normalized key
    const consultant = allConsultants.find(c => normalizeKey(c.name) === key);
    if (!consultant || !consultant.active) return { headerBg: '#9CA3AF', rowBg: '#F3F4F6', hoverBg: 'hover:bg-gray-100', dotColor: '#D1D5DB' };
    return { headerBg: consultant.color, rowBg: `${consultant.color}40`, hoverBg: '', dotColor: consultant.color };
}

async function initApp() {
    const app = initializeApp(firebaseConfig); db = getFirestore(app); auth = getAuth(app);
    await signInAnonymously(auth);
    onAuthStateChanged(auth, (user) => {
        if (user) { 
            userId = user.uid; document.getElementById('loading-overlay').classList.add('opacity-0', 'pointer-events-none'); 
            if(!currentUserDisplay) document.getElementById('name-modal').classList.remove('hidden'); 
            else { 
                document.getElementById('user-badge').textContent = currentUserDisplay; 
                initListeners(); 
                // Initialize Header Time
                updateHeaderTime();
                setInterval(updateHeaderTime, 1000);
            }
            toggleView(currentView);
        }
    });
}

function updateHeaderTime() {
    const now = new Date();
    // Format: Wed, 28 Jan 08:05 AM
    document.getElementById('header-time').textContent = now.toLocaleDateString('en-GB', { weekday: 'short', day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' }).replace(',', '');
    
    // Greeting
    const hr = now.getHours();
    let g = 'Morning';
    if (hr >= 12) g = 'Afternoon';
    if (hr >= 18) g = 'Evening';
    
    // Ensure name is present
    const name = currentUserDisplay || 'User';
    document.getElementById('header-greeting').textContent = `Good ${g}, ${name}!`;
}

function showToast(msg, type='success') {
    const toast = document.createElement('div'); const color = type === 'error' ? 'bg-red-500' : 'bg-dark';
    toast.className = `${color} text-white px-6 py-3 rounded-full shadow-2xl flex items-center gap-3 text-sm font-semibold pointer-events-auto toast-enter`;
    toast.innerHTML = type==='error' ? `<svg class="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>${msg}` : `<svg class="w-5 h-5 text-green-400" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7" /></svg>${msg}`;
    document.getElementById('toast-container').appendChild(toast);
    setTimeout(() => { toast.style.opacity = '0'; toast.style.transform = 'translate(-50%, 20px)'; setTimeout(()=>toast.remove(), 300); }, 3000);
}

async function logAction(message) {
    if(!currentUserDisplay) return;
    try { await addDoc(collection(db, `artifacts/${appId}/public/data/notifications`), { user: currentUserDisplay, message: message, timestamp: serverTimestamp() }); } catch(e) { console.error("Log failed", e); }
}

function initListeners() {
    onSnapshot(query(collection(db, `artifacts/${appId}/public/data/tasks`)), (s) => { allTasks = s.docs.map(d => ({ id: d.id, ...d.data() })); refreshView(); });
    onSnapshot(query(collection(db, `artifacts/${appId}/public/data/schools`), orderBy('school')), (s) => { 
        allSchools = s.docs.map(d => ({ id: d.id, ...d.data() })); 
        populateSchoolDatalist(); 
        updateConsultantUI(); // Ensure dropdowns update when schools import
        if(currentView === 'zones') window.renderZonesTable(); 
    });
    onSnapshot(query(collection(db, `artifacts/${appId}/public/data/consultants`), orderBy('name')), (s) => { 
        allConsultants = s.docs.map(d => ({ id: d.id, ...d.data() })); 
        if(allConsultants.length === 0) seedConsultants(); else { updateConsultantUI(); refreshView(); } 
    });
    onSnapshot(query(collection(db, `artifacts/${appId}/public/data/notifications`), orderBy('timestamp', 'desc'), limit(50)), (s) => {
        const list = document.getElementById('notif-list'); const badge = document.getElementById('notif-badge'); list.innerHTML = '';
        if(s.empty) { list.innerHTML = '<div class="p-8 text-center text-gray-400 text-xs">No recent activity.</div>'; badge.classList.add('hidden'); } 
        else { badge.classList.remove('hidden'); s.docs.forEach(d => { const n = d.data(); const time = n.timestamp ? new Date(n.timestamp.toDate()).toLocaleString([], { month: 'short', day: 'numeric', hour: '2-digit', minute:'2-digit'}) : 'Just now'; list.innerHTML += `<div class="px-4 py-3 border-b border-gray-50 hover:bg-gray-50 transition"><div class="flex justify-between items-start mb-0.5"><span class="text-xs font-bold text-dark">${n.user}</span><span class="text-[10px] text-gray-400">${time}</span></div><p class="text-xs text-gray-600 leading-snug">${n.message}</p></div>`; }); }
    });
}

window.saveUserName = (e) => { 
    e.preventDefault(); 
    const val = document.getElementById('user-name-input').value.trim().toUpperCase(); 
    if(val) { 
        currentUserDisplay = val; 
        sessionStorage.setItem('checklist_username', val); 
        document.getElementById('name-modal').classList.add('hidden'); 
        document.getElementById('user-badge').textContent = val; 
        initListeners();
        // Start header timer
        updateHeaderTime();
        setInterval(updateHeaderTime, 1000);
    } 
};
window.toggleNotifications = () => { const pop = document.getElementById('notif-popover'); pop.classList.toggle('hidden'); if(!pop.classList.contains('hidden')) document.getElementById('notif-badge').classList.add('hidden'); };

async function seedConsultants() { const batch = writeBatch(db); const defaults = ["Sarah", "Martin", "Krystle", "Charlynn", "Sapnaa", "Yuniza"]; defaults.forEach(name => { batch.set(doc(collection(db, `artifacts/${appId}/public/data/consultants`)), { name, color: LEGACY_COLORS[name] || '#6B7280', active: true, createdAt: serverTimestamp() }); }); await batch.commit(); }

function updateConsultantUI() {
    // Deduplicate names using normalized keys (fixes duplicate "Krystle" vs "Krystle " vs "Krystle.")
    const uniqueNamesMap = new Map();

    const addName = (rawName) => {
        if(!rawName) return;
        const clean = cleanStr(rawName);
        if(!clean) return;
        
        const key = normalizeKey(clean);
        
        // Skip if the name is effectively empty after key normalization
        if (key.length === 0) return;

        // Store the "cleanest" version of the name.
        // Priority: If we already have a name, keep it (assuming first source is better/canonical)
        // Note: allConsultants (defined colors) are processed first below, so they take precedence.
        if(!uniqueNamesMap.has(key)) {
            uniqueNamesMap.set(key, clean);
        }
    };

    allConsultants.forEach(c => { if(c.active) addName(c.name); });
    allSchools.forEach(s => { if(s.consultant) addName(s.consultant); });
    
    const sortedNames = Array.from(uniqueNamesMap.values()).sort();

    const populateSelect = (targetId) => {
        const select = document.getElementById(targetId); if(!select) return;
        const currentValue = select.value; 
        select.innerHTML = '';
        
        const defaultOpt = document.createElement('option'); 
        const isMainFilter = targetId === 'filter-ec' || targetId === 'mob-filter-ec';
        defaultOpt.value = isMainFilter ? 'all' : ''; 
        defaultOpt.textContent = isMainFilter ? 'All Consultants' : (targetId === 'zone-filter-ec' ? 'All ECs' : 'Unassigned'); 
        select.appendChild(defaultOpt);
        
        sortedNames.forEach(name => { 
            const opt = document.createElement('option'); 
            opt.value = name; 
            opt.textContent = name; 
            select.appendChild(opt); 
        });
        
        // Try to restore value, normalized
        if(currentValue && currentValue !== 'all') {
             // Find the matching normalized name in the new options
             const currentKey = normalizeKey(currentValue);
             const matchingName = uniqueNamesMap.get(currentKey);
             if(matchingName) select.value = matchingName;
        } else if (currentValue === 'all') {
            select.value = 'all';
        }
    };
    
    populateSelect('filter-ec'); 
    populateSelect('mob-filter-ec'); 
    populateSelect('sf_consultant'); 
    populateSelect('zone-filter-ec');
}

window.addConsultant = async () => { const name = document.getElementById('new-consultant-name').value.trim(); if(!name) return; await addDoc(collection(db, `artifacts/${appId}/public/data/consultants`), { name, color: document.getElementById('new-consultant-color').value, active: true, createdAt: serverTimestamp() }); logAction(`Added consultant: ${name}`); document.getElementById('new-consultant-name').value = ''; showToast("Consultant added"); }
window.toggleConsultantStatus = async (id, isActive) => await updateDoc(doc(db, `artifacts/${appId}/public/data/consultants`, id), { active: isActive });
window.updateConsultantColor = async (id, color) => await updateDoc(doc(db, `artifacts/${appId}/public/data/consultants`, id), { color: color });
window.openConsultantModal = () => document.getElementById('consultant-modal-overlay').classList.remove('hidden');

function populateSchoolDatalist() { const dl = document.getElementById('school-list'); dl.innerHTML = ''; allSchools.forEach(s => { const o = document.createElement('option'); o.value = s.school; dl.appendChild(o); }); }

// --- MOBILE FILTERS ---
window.toggleMobileFilters = () => { const sheet = document.getElementById('mobile-filter-sheet'); const content = sheet.querySelector('div.transform'); if (sheet.classList.contains('hidden')) { ['type','ec','status'].forEach(k => document.getElementById(`mob-filter-${k}`).value = document.getElementById(`filter-${k}`).value); sheet.classList.remove('hidden'); setTimeout(() => content.classList.remove('translate-y-full'), 10); } else { content.classList.add('translate-y-full'); setTimeout(() => sheet.classList.add('hidden'), 300); } };
window.setMobileFilter = (key, value) => { document.getElementById(`filter-${key}`).value = value; if(currentView === 'kanban' || currentView === 'table') viewFilters[currentView][key] = value; refreshView(); }

// --- VIEWS ---
window.refreshView = function() {
    if(currentView === 'kanban' || currentView === 'table') { viewFilters[currentView].search = document.getElementById('table-search').value; viewFilters[currentView].type = document.getElementById('filter-type').value; viewFilters[currentView].ec = document.getElementById('filter-ec').value; viewFilters[currentView].status = document.getElementById('filter-status').value; }
    else if (currentView === 'zones') { viewFilters.zones.search = document.getElementById('zones-search').value; }
    if(currentView === 'kanban') renderKanban(); else if (currentView === 'table') renderTable(); else if (currentView === 'zones') window.renderZonesTable();
}

// --- NEW: COPY TASK DETAILS FUNCTION ---
window.copyTaskDetails = (e, id) => {
    e.stopPropagation(); // Stop the row/card click event (opening modal)
    
    const t = allTasks.find(x => x.id === id);
    if (!t) return;

    // Use existing formatDate helper
    const dateStr = formatDate(t.closing_date);
    const code = t.moe_code || '-';

    // Construct the text format
    const textToCopy = `${t.school}\n${t.programme}\n${code}\nClosing on ${dateStr}`;

    // Write to clipboard
    navigator.clipboard.writeText(textToCopy).then(() => {
        showToast("Details copied to clipboard!");
    }).catch(err => {
        console.error('Copy failed', err);
        showToast("Failed to copy", "error");
    });
};

function renderKanban() {
    const itqC = document.getElementById('kanban-area-itq'); const svpC = document.getElementById('kanban-area-svp'); itqC.innerHTML = ''; svpC.innerHTML = '';
    const filtered = filterTasks(allTasks);
    KANBAN_STATUSES.forEach(status => {
        const tasks = filtered.filter(t => t.status === status).sort((a,b) => new Date(a.closing_date||'9999-12-31') - new Date(b.closing_date||'9999-12-31'));
        itqC.appendChild(createKanbanColumn(status, tasks.filter(t => t.type !== 'SVP'))); svpC.appendChild(createKanbanColumn(status, tasks.filter(t => t.type === 'SVP')));
    });
}

function createKanbanColumn(status, tasks) {
    const col = document.createElement('div'); col.className = 'kanban-column bg-white rounded-2xl shadow-ios border border-gray-100 flex flex-col h-[600px] flex-shrink-0 snap-center';
    col.innerHTML = `<div class="p-4 border-b border-gray-100 font-semibold text-gray-700 flex justify-between items-center ${status==='Skipped'?'bg-gray-200':'bg-gray-50'} rounded-t-2xl"><span class="text-sm uppercase tracking-wide">${status}</span><span class="bg-white border border-gray-200 text-gray-600 text-xs px-2 py-0.5 rounded-full font-bold">${tasks.length}</span></div><div class="kanban-column-content p-3 flex-grow overflow-y-auto space-y-3 bg-gray-50/30 custom-scrollbar" ondragover="window.allowDrop(event)" ondrop="window.drop(event, '${status}')"></div>`;
    const list = col.querySelector('.kanban-column-content'); tasks.forEach(task => list.appendChild(createKanbanCard(task))); return col;
}

function createKanbanCard(task) {
    const ec = task.assignment ? task.assignment.split(',')[0].trim() : '';
    const styles = getConsultantStyles(ec);
    const ps = getProgressState(task);
    const segs = PROGRESS_ITEMS.map((item, i) => `<div class="progress-segment ${i<6?(ps.allFirstSix && ps[item.id]?'active-solid':(ps[item.id]?'active-light':'')):(ps.bothLastTwo && ps[item.id]?'active-solid':(ps[item.id]?'active-light':''))}"></div>`).join('');
    
    // --- APPOINTMENT DOT LOGIC ---
    let apptBadge = '';
    if (task.appointment) {
        const apptDate = new Date(task.appointment);
        const now = new Date();
        if (apptDate > now) {
            apptBadge = `<span class="appt-dot appt-dot-future" title="Upcoming Appointment: ${new Date(task.appointment).toLocaleString()}"></span>`;
        } else {
            apptBadge = `<span class="appt-dot appt-dot-past" title="Past Appointment: ${new Date(task.appointment).toLocaleString()}"></span>`;
        }
    }

    const card = document.createElement('div');
    card.className = `kanban-card bg-white rounded-xl shadow-sm border border-gray-100 cursor-grab hover:shadow-ios-hover transition-all transform hover:-translate-y-1 overflow-hidden relative`;
    card.setAttribute('draggable', 'true'); card.ondragstart = (e) => e.dataTransfer.setData("text", task.id);
    
    // UPDATED: Injected apptBadge into header
    card.innerHTML = `
        <div class="kanban-card-header text-white px-3 py-2 flex justify-between items-center" style="background-color: ${styles.headerBg}" onclick="window.openViewModalById('${task.id}')">
            <div class="flex items-center gap-2">${apptBadge}<span class="text-[10px] font-bold uppercase tracking-wider bg-black/20 px-1.5 py-0.5 rounded">${task.type}</span>${(task.moe_code && task.moe_code.length > 4) ? `<span class="text-[10px] font-medium opacity-90 tracking-wide border-l border-white/30 pl-2">${task.moe_code.slice(-4)}</span>` : ''}</div><span class="text-xs font-medium">${formatDate(task.closing_date)}</span>
        </div>
        <div class="p-3">
            <h4 class="text-sm font-bold text-gray-800 leading-snug mb-1 line-clamp-2 cursor-pointer hover:text-primary" onclick="window.openViewModalById('${task.id}')">${task.school}</h4>
            <div class="text-xs text-gray-500 mb-3 line-clamp-1">${task.programme}</div>
            <div class="flex justify-between items-center pt-2 border-t border-gray-50">
                    <div class="flex gap-1">
                        <a href="${task.specs}" target="_blank" class="text-gray-400 hover:text-ios_blue p-1" onclick="event.stopPropagation()"><svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path></svg></a>
                        <a href="${task.folder}" target="_blank" class="text-gray-400 hover:text-accent p-1" onclick="event.stopPropagation()"><svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 7v10a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-6l-2-2H5a2 2 0 00-2 2z"></path></svg></a>
                        <button class="text-gray-400 hover:text-green-600 p-1" onclick="event.stopPropagation(); window.toggleProgress(event, '${task.id}')"><svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01"></path></svg></button>
                        <button class="text-gray-400 hover:text-purple-600 p-1" onclick="window.copyTaskDetails(event, '${task.id}')" title="Copy Details"><svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z"></path></svg></button>
                    </div>
                <span class="text-[10px] font-medium px-2 py-0.5 rounded bg-blue-50 text-blue-700 whitespace-nowrap">${getTaskLabel(task)}</span>
            </div>
            <div class="flex gap-0.5 mt-2 h-1 w-full opacity-80">${segs}</div>
        </div>`;
    return card;
}

function renderTable() {
    const tbody = document.getElementById('table-body'); tbody.innerHTML = '';
    const tasks = filterTasks(allTasks).sort((a, b) => (KANBAN_STATUSES.indexOf(a.status) - KANBAN_STATUSES.indexOf(b.status)) || (new Date(a.closing_date||'9999-12-31') - new Date(b.closing_date||'9999-12-31')));
    tasks.forEach(task => {
        const ec = task.assignment ? task.assignment.split(',')[0].trim() : '';
        const styles = getConsultantStyles(ec);
        const ps = getProgressState(task);
        const segs = PROGRESS_ITEMS.map((item, i) => `<div class="progress-segment ${i<6?(ps.allFirstSix && ps[item.id]?'active-solid':(ps[item.id]?'active-light':'')):(ps.bothLastTwo && ps[item.id]?'active-solid':(ps[item.id]?'active-light':''))}"></div>`).join('');
        
        // --- APPOINTMENT DOT LOGIC ---
        let apptBadge = '';
        if (task.appointment) {
            const apptDate = new Date(task.appointment);
            const now = new Date();
            if (apptDate > now) {
                apptBadge = `<span class="appt-dot appt-dot-future" title="Upcoming Appointment: ${new Date(task.appointment).toLocaleString()}"></span>`;
            } else {
                apptBadge = `<span class="appt-dot appt-dot-past" title="Past Appointment: ${new Date(task.appointment).toLocaleString()}"></span>`;
            }
        }

        const tr = document.createElement('tr');
        tr.className = `cursor-pointer border-b border-gray-100 transition-all duration-150`; tr.style.backgroundColor = styles.rowBg;
        tr.onmouseenter = () => { tr.style.filter = "brightness(0.92) saturate(1.05)"; tr.style.zIndex = "10"; tr.style.boxShadow = "0 4px 6px -1px rgba(0, 0, 0, 0.05)"; tr.style.transform = "scale(1.002)"; }
        tr.onmouseleave = () => { tr.style.filter = "none"; tr.style.zIndex = "auto"; tr.style.boxShadow = "none"; tr.style.transform = "none"; }
        tr.onclick = (e) => { if(!e.target.closest('button') && !e.target.closest('a')) window.openViewModalById(task.id); };
        
        // UPDATED: Injected apptBadge into Actions column (last td)
        tr.innerHTML = `<td class="px-4 sm:px-6 py-4 text-sm font-bold text-gray-700 align-middle">${task.type}</td><td class="px-4 sm:px-6 py-4 align-middle"><div class="text-sm font-bold text-gray-900">${task.school}</div><div class="text-xs text-gray-600">${task.programme}</div><div class="text-xs text-gray-400 font-mono">${task.moe_code||''}</div></td><td class="px-4 sm:px-6 py-4 text-sm text-gray-500 font-mono align-middle">${formatDate(task.closing_date)}</td><td class="px-4 sm:px-6 py-4 text-sm text-gray-600 align-middle whitespace-nowrap"><span class="bg-white border border-gray-200 px-2 py-1 rounded-md text-xs shadow-sm">${getTaskLabel(task)}</span></td><td class="px-4 sm:px-6 py-4 align-middle whitespace-nowrap"><div class="flex flex-col gap-1.5"><span class="px-2.5 py-0.5 text-xs font-bold rounded-full w-fit ${task.status==='In Progress'?'bg-green-100 text-green-800':'bg-gray-100 text-gray-600'}">${task.status}</span><div class="flex gap-0.5 h-1 w-full max-w-[120px] opacity-80">${segs}</div></div></td><td class="px-4 sm:px-6 py-4 text-sm text-gray-500 align-middle">${task.costing_type||'-'}</td><td class="px-4 sm:px-6 py-4 text-right align-middle"><div class="flex gap-2 justify-end items-center">${apptBadge}<a href="${task.specs}" target="_blank" class="text-gray-400 hover:text-ios_blue p-1" onclick="event.stopPropagation()"><svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path></svg></a><a href="${task.folder}" target="_blank" class="text-gray-400 hover:text-accent p-1" onclick="event.stopPropagation()"><svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 7v10a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-6l-2-2H5a2 2 0 00-2 2z"></path></svg></a><button class="text-gray-400 hover:text-green-600 p-1" onclick="event.stopPropagation(); window.toggleProgress(event, '${task.id}')"><svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01"></path></svg></button><button class="text-gray-400 hover:text-purple-600 p-1" onclick="window.copyTaskDetails(event, '${task.id}')" title="Copy Details"><svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z"></path></svg></button></div></td>`;
        tbody.appendChild(tr);
    });
}

window.renderZonesTable = () => {
    const tbody = document.getElementById('zones-body'); tbody.innerHTML = '';
    const search = document.getElementById('zones-search').value.toLowerCase();
    const zoneFilter = document.getElementById('zone-filter-zone').value;
    const ecFilter = document.getElementById('zone-filter-ec').value;
    const ecFilterKey = normalizeKey(ecFilter); // Normalize the selected filter
    
    const clearBtn = document.getElementById('zones-search-clear'); if(search.length > 0) clearBtn.classList.remove('hidden'); else clearBtn.classList.add('hidden');

    let filtered = allSchools.filter(s => {
        // Use cleanStr to normalize comparisons
        const sConsultantClean = cleanStr(s.consultant);
        const matchesSearch = !search || s.school.toLowerCase().includes(search) || (s.zone||'').toLowerCase().includes(search) || sConsultantClean.toLowerCase().includes(search);
        const matchesZone = !zoneFilter || s.zone === zoneFilter;
        // Normalize comparison for EC filter
        const matchesEC = !ecFilter || normalizeKey(sConsultantClean) === ecFilterKey; 
        return matchesSearch && matchesZone && matchesEC;
    });
    
    document.getElementById('zones-count').textContent = `${filtered.length} Schools`;
    
    // UPDATED SORT LOGIC (School A-Z secondary sort)
    filtered.sort((a, b) => {
        const key = schoolSort.key;
        const dir = schoolSort.dir === 'asc' ? 1 : -1;
        const valA = cleanStr(a[key]).toLowerCase();
        const valB = cleanStr(b[key]).toLowerCase();
        
        if (valA < valB) return -1 * dir;
        if (valA > valB) return 1 * dir;
        
        // Secondary Sort: School Name Ascending
        if (key !== 'school') {
            const schoolA = cleanStr(a.school).toLowerCase();
            const schoolB = cleanStr(b.school).toLowerCase();
            if (schoolA < schoolB) return -1;
            if (schoolA > schoolB) return 1;
        }
        
        return 0;
    });
    
    filtered.forEach((s, idx) => {
        const zoneClass = ZONE_COLORS[s.zone] || '';
        
        const tr = document.createElement('tr'); 
        tr.className = `border-b border-gray-50 transition-colors group zone-row cursor-pointer ${zoneClass}`;
        tr.onclick = () => window.openSchoolModal(s.id); 
        tr.innerHTML = `
            <td class="px-6 py-3 text-xs text-gray-500 font-mono align-middle">${idx+1}</td>
            <td class="px-6 py-3 text-sm font-bold text-dark align-middle">${s.school}</td>
            <td class="px-6 py-3 align-middle"><span class="px-2 py-1 rounded bg-white/50 text-xs font-medium text-gray-700 border border-black/10">${s.zone}</span></td>
            <td class="px-6 py-3 text-sm text-gray-600 align-middle">${s.consultant || 'Unassigned'}</td>
            <td class="px-6 py-3 text-sm text-gray-500 align-middle">${s.aec || '-'}</td>
        `;
        tbody.appendChild(tr);
    });
    
    // Update sort icons visual state
    ['school', 'zone', 'consultant'].forEach(k => {
        const icon = document.getElementById(`sort-icon-${k}`);
        if(schoolSort.key === k) {
            icon.classList.remove('text-gray-300');
            icon.classList.add('text-ios_blue');
        } else {
            icon.classList.remove('text-ios_blue');
            icon.classList.add('text-gray-300');
        }
    });
};
document.getElementById('zones-search').addEventListener('input', window.renderZonesTable);
window.clearZonesSearch = () => { document.getElementById('zones-search').value = ''; window.renderZonesTable(); }
window.sortSchools = (key) => { 
    if (schoolSort.key === key) schoolSort.dir = schoolSort.dir === 'asc' ? 'desc' : 'asc'; 
    else { schoolSort.key = key; schoolSort.dir = 'asc'; } 
    window.renderZonesTable(); 
}

// --- DATA SUMMARY LOGIC ---
window.showAllocationSummary = () => {
    const summary = {};
    const zoneCounts = { "East": 0, "North": 0, "South": 0, "West": 0 };
    let unassignedEC = 0;
    let unassignedAEC = 0;

    allSchools.forEach(s => {
        const ec = cleanStr(s.consultant);
        const aec = cleanStr(s.aec);
        const zone = s.zone || 'Unknown';

        // EC Counts
        if (ec) {
            // Use normalized key for counting to avoid duplicates in stats
            const ecKey = normalizeKey(ec); 
            // Store display name from first occurrence if not set
            const displayName = ec; 
            // We need a map for stats to be accurate
            if(!summary[ecKey]) summary[ecKey] = { name: displayName, count: 0 };
            summary[ecKey].count++;
        } else {
            unassignedEC++;
        }

        // AEC Counts - Updated Logic for Hyphen
        if (!aec || aec === '-') unassignedAEC++;

        // Zone Counts
        if (zoneCounts.hasOwnProperty(zone)) zoneCounts[zone]++;
    });

    let html = `<div class="space-y-6">`;
    
    // 1. Schools per EC
    html += `<div><h4 class="text-xs font-bold text-gray-500 uppercase tracking-wider mb-2">Schools per Consultant</h4><div class="grid grid-cols-2 sm:grid-cols-3 gap-3">`;
    
    // Sort summary by display name
    const sortedSummary = Object.values(summary).sort((a,b) => a.name.localeCompare(b.name));
    
    sortedSummary.forEach(item => {
         html += `<div class="bg-gray-50 p-3 rounded-lg border border-gray-100 flex justify-between items-center"><span class="text-sm font-semibold text-gray-700">${item.name}</span><span class="text-sm font-bold text-ios_blue">${item.count}</span></div>`;
    });
    html += `</div></div>`;

    // 2. Schools per Zone
    html += `<div><h4 class="text-xs font-bold text-gray-500 uppercase tracking-wider mb-2">Schools per Zone</h4><div class="grid grid-cols-2 sm:grid-cols-4 gap-3">`;
    Object.keys(zoneCounts).forEach(z => {
        const colorClass = ZONE_COLORS[z] || 'bg-gray-50';
         html += `<div class="${colorClass} p-3 rounded-lg border border-gray-100 flex justify-between items-center"><span class="text-sm font-semibold text-gray-700">${z}</span><span class="text-sm font-bold text-dark">${zoneCounts[z]}</span></div>`;
    });
    html += `</div></div>`;

    // 3. Unassigned
    html += `<div><h4 class="text-xs font-bold text-gray-500 uppercase tracking-wider mb-2">Unassigned</h4>
             <div class="flex gap-4">
                <div class="flex-1 bg-red-50 p-3 rounded-lg border border-red-100 flex flex-col items-center justify-center">
                    <span class="text-2xl font-bold text-red-600">${unassignedEC}</span>
                    <span class="text-xs text-red-400 font-medium">No EC Assigned</span>
                </div>
                <div class="flex-1 bg-orange-50 p-3 rounded-lg border border-orange-100 flex flex-col items-center justify-center">
                    <span class="text-2xl font-bold text-orange-600">${unassignedAEC}</span>
                    <span class="text-xs text-orange-400 font-medium">No AEC Assigned</span>
                </div>
             </div></div>`;
    
    html += `</div>`;
    document.getElementById('summary-content').innerHTML = html;
    document.getElementById('summary-modal').classList.remove('hidden');
};


// --- IMPORT LOGIC ---
window.verifyImportPassword = () => {
    const p = document.getElementById('import-password').value;
    if(p === "DARYL") {
        document.getElementById('password-modal').classList.add('hidden');
        window.triggerImport();
    } else {
        showToast("Incorrect Password", "error");
    }
    document.getElementById('import-password').value = '';
};

window.triggerImport = () => {
    let input = document.getElementById('csv-upload-input');
    if (!input) {
        input = document.createElement('input');
        input.type = 'file';
        input.id = 'csv-upload-input';
        input.accept = '.csv';
        input.style.display = 'none';
        input.onchange = (e) => window.handleCSVUpload(e.target);
        document.body.appendChild(input);
    }
    input.click();
};

window.handleCSVUpload = (input) => {
    const file = input.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = async (e) => {
        const text = e.target.result;
        // Basic CSV Parsing (assuming Schools, Zone, EC, AEC)
        const rows = text.split(/\r\n|\n/).map(r => r.split(',')).filter(r => r.length > 1);
        
        // Remove header if it exists
        if(rows[0][0].toLowerCase().includes('school')) rows.shift();
        
        try {
            document.getElementById('loading-overlay').classList.remove('opacity-0', 'pointer-events-none');
            
            // 1. Delete all existing schools (Batch)
            const snap = await getDocs(collection(db, `artifacts/${appId}/public/data/schools`));
            const batchSize = 400; 
            let batch = writeBatch(db);
            let count = 0;
            
            snap.docs.forEach((d) => {
                batch.delete(d.ref);
                count++;
                if (count >= batchSize) { batch.commit(); batch = writeBatch(db); count = 0; }
            });
            await batch.commit();

            // 2. Add new schools
            batch = writeBatch(db);
            count = 0;
            rows.forEach(row => {
                // row[0]=School, row[1]=Zone, row[2]=EC, row[3]=AEC 
                const clean = (val) => val ? val.replace(/"/g, '').trim() : '';
                const schoolName = clean(row[0]);
                if(schoolName) {
                    const ref = doc(collection(db, `artifacts/${appId}/public/data/schools`));
                    batch.set(ref, { 
                        school: schoolName, 
                        zone: clean(row[1]), 
                        consultant: clean(row[2]), 
                        aec: clean(row[3])
                    });
                    count++;
                    if (count >= batchSize) { batch.commit(); batch = writeBatch(db); count = 0; }
                }
            });
            await batch.commit();
            
            logAction(`Imported new Zone Allocation DB`);
            showToast("Database Updated Successfully");
        } catch(err) {
            console.error(err);
            showToast("Import Failed", "error");
        } finally {
             document.getElementById('loading-overlay').classList.add('opacity-0', 'pointer-events-none');
             input.value = ''; // Reset
        }
    };
    reader.readAsText(file);
};

function filterTasks(tasks) {
    const search = document.getElementById('table-search').value.toLowerCase();
    const type = document.getElementById('filter-type').value; 
    const ec = document.getElementById('filter-ec').value; 
    const ecKey = normalizeKey(ec); // Normalize filter
    const status = document.getElementById('filter-status').value; 
    
    return tasks.filter(t => {
        const fullText = [t.school, t.programme, t.type, t.moe_code, t.contact1?.name, t.contact1?.email, t.contact1?.cont, t.contact2?.name, t.contact2?.email, t.contact2?.cont, t.contact3?.name, t.contact3?.email, t.contact3?.cont].filter(Boolean).join(' ').toLowerCase();
        
        // Normalize task assignment check
        let assignmentMatch = false;
        if(ec === 'all') assignmentMatch = true;
        else if (t.assignment) {
             const assignmentParts = t.assignment.split(',').map(p => normalizeKey(p));
             if(assignmentParts.includes(ecKey)) assignmentMatch = true;
             // Also check raw inclusion for legacy
             if(t.assignment.toLowerCase().includes(ec.toLowerCase())) assignmentMatch = true;
        }

        if (search && !fullText.includes(search)) return false;
        if (type !== 'all' && t.type !== type) return false;
        if (status !== 'all' && t.status !== status) return false;
        if (!assignmentMatch) return false;
        return true;
    });
}
['table-search','filter-type','filter-ec','filter-status'].forEach(id => document.getElementById(id).addEventListener(id==='table-search'?'input':'change', refreshView));

// --- VIEW MODAL LOGIC (New v2.1) ---
window.openViewModalById = (id) => { const task = allTasks.find(t => t.id === id); if(task) window.openViewModal(task); };
window.openViewModal = (task) => {
    // Populate simple fields
    document.getElementById('view-modal-title').textContent = task.school;
    document.getElementById('view-modal-status').textContent = task.status;
    document.getElementById('view-modal-status').className = `px-2 py-1 rounded text-xs font-bold ${task.status==='In Progress'?'bg-green-100 text-green-800':'bg-gray-100 text-gray-600'}`;
    
    // Overview Column
    document.getElementById('view-school').textContent = task.school;
    const schoolObj = allSchools.find(s => s.school === task.school);
    document.getElementById('view-zone').textContent = schoolObj ? `${schoolObj.zone} Zone` : '';
    
    // UPDATED: Populate Separate Rows
    document.getElementById('view-type').textContent = task.type;
    document.getElementById('view-code').textContent = task.moe_code || '-';
    document.getElementById('view-date').textContent = formatDate(task.closing_date);

    // NEW APPOINTMENT LOGIC
    const apptEl = document.getElementById('view-appointment');
    if (task.appointment) {
        const d = new Date(task.appointment);
        const now = new Date();
        const isFuture = d > now;
        const formatted = d.toLocaleDateString('en-GB', { weekday: 'short', day: '2-digit', month: 'short', year: 'numeric' }) + ', ' + d.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: true });
        
        if (isFuture) {
            apptEl.innerHTML = `<span class="appt-dot appt-dot-future mr-2"></span>${formatted}`;
            apptEl.className = "font-bold text-dark mt-0.5 flex items-center";
        } else {
             apptEl.innerHTML = formatted;
             apptEl.className = "font-semibold text-gray-700 mt-0.5 flex items-center";
        }
    } else {
        apptEl.textContent = "No scheduled appointment";
        apptEl.className = "text-sm text-gray-400 mt-0.5 italic";
    }

    document.getElementById('view-assignment').textContent = getTaskLabel(task);
    document.getElementById('view-brand').textContent = task.brand || '';
    
    // Progress - UPDATED TO INTERACTIVE CHECKBOXES
    const ps = getProgressState(task);
    const progCont = document.getElementById('view-progress-container');
    progCont.innerHTML = PROGRESS_ITEMS.map(item => {
        const checked = ps[item.id];
        return `<label class="flex items-center gap-2 mb-1 cursor-pointer transition-opacity hover:opacity-80">
            <input type="checkbox" class="form-checkbox h-4 w-4 text-ios_blue rounded border-gray-300 focus:ring-ios_blue transition duration-150 ease-in-out" 
                   ${checked ? 'checked' : ''} 
                   onchange="window.updateProgress('${task.id}', '${item.id}', this.checked)">
            <span class="text-xs text-gray-600 ${checked ? 'line-through opacity-70' : ''}">${item.label}</span>
        </label>`;
    }).join('');

    // Programme Column
    document.getElementById('view-programme').textContent = task.programme;
    document.getElementById('view-specs-btn').href = task.specs;
    document.getElementById('view-folder-btn').href = task.folder;
    document.getElementById('view-trainers').textContent = task.trainers || 'No trainer details specified.';
    
    // UPDATED: Use Linkify for Notes
    document.getElementById('view-notes').innerHTML = linkify(task.notes || 'No notes.');

    // Contacts Column
    const contactsList = document.getElementById('view-contacts-list');
    contactsList.innerHTML = '';
    // Check all possible contacts (up to 4 now)
    for(let i=1; i<=4; i++) {
        const c = task[`contact${i}`];
        if(c && (c.name || c.email)) {
            contactsList.innerHTML += `
                <div class="p-3 bg-gray-50 rounded-xl border border-gray-100 text-sm">
                    <div class="font-bold text-dark">${c.name || '-'}</div>
                    <div class="text-xs text-gray-500 mb-1">${c.des || ''} ${c.dept ? `(${c.dept})` : ''}</div>
                    ${c.cont ? `<div class="flex items-center gap-1.5 text-gray-600 text-xs"><svg class="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 5a2 2 0 012-2h3.28a1 1 0 01.948.684l1.498 4.493a1 1 0 01-.502 1.21l-2.257 1.13a11.042 11.042 0 005.516 5.516l1.13-2.257a1 1 0 011.21-.502l4.493 1.498a1 1 0 01.684.949V19a2 2 0 01-2 2h-1C9.716 21 3 14.284 3 6V5z"></path></svg>${c.cont}</div>` : ''}
                    ${c.email ? `<div class="flex items-center gap-1.5 text-ios_blue text-xs truncate mt-0.5"><svg class="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z"></path></svg><a href="mailto:${c.email}">${c.email}</a></div>` : ''}
                </div>
            `;
        }
    }
    if(contactsList.innerHTML === '') contactsList.innerHTML = '<div class="text-xs text-gray-400 italic">No contacts listed.</div>';

    // Costing
    document.getElementById('view-cost-type').textContent = task.costing_type || '-';
    // Optionally use Linkify here too if needed, though notes is primary
    document.getElementById('view-cost-details').innerHTML = linkify((task.cost_specs || '') + '\n' + (task.cost_val || ''));

    // Setup Edit Button
    document.getElementById('view-edit-btn').onclick = () => {
        document.getElementById('view-modal-overlay').classList.add('hidden');
        window.openModal(task);
    };

    document.getElementById('view-modal-overlay').classList.remove('hidden');
}


// --- EDIT MODAL LOGIC ---
document.getElementById('f_type').addEventListener('change', (e) => updateMoeFieldState(e.target.value));
const updateMoeFieldState = (type) => {
    const group = document.getElementById('moe-field-group'); const input = document.getElementById('f_moe_code'); const star = document.getElementById('moe-star');
    group.classList.remove('hidden'); input.required = false; input.disabled = false; star.classList.add('hidden'); input.placeholder = "e.g. ITQ-2025-001";
    if (type === 'ITQ') { input.required = true; star.classList.remove('hidden'); }
    else if (type === 'SVP') { group.classList.add('hidden'); input.disabled = true; input.value = ''; }
    else if (type === 'Tender') { input.placeholder = "Optional"; }
};

window.openModal = (task=null) => {
    editingTaskId = task ? task.id : null; 
    document.getElementById('task-form').reset(); 
    document.getElementById('modal-title').textContent = task ? 'Edit Task' : 'New Task'; 
    document.getElementById('delete-btn').classList.toggle('hidden', !task);
    
    const typeVal = task?.type || 'ITQ';
    document.getElementById('f_type').value = typeVal;
    document.getElementById('f_brand').value = task?.brand || '';
    document.getElementById('f_costing_type').value = task?.costing_type || '';
    document.getElementById('f_status').value = task?.status || 'In Progress';

    updateMoeFieldState(typeVal);

    // --- POPULATE CONTACTS SPREADSHEET (4 Rows) ---
    const tbody = document.getElementById('contacts-body');
    tbody.innerHTML = '';
    for(let i=1; i<=4; i++) {
        const c = task ? (task[`contact${i}`] || {}) : {};
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="spreadsheet-cell p-2 border-b border-r border-gray-100 outline-none select-none text-gray-700" data-row="${i}" data-col="0">${c.name||''}</td>
            <td class="spreadsheet-cell p-2 border-b border-r border-gray-100 outline-none select-none text-gray-700" data-row="${i}" data-col="1">${c.des||''}</td>
            <td class="spreadsheet-cell p-2 border-b border-r border-gray-100 outline-none select-none text-gray-700" data-row="${i}" data-col="2">${c.dept||''}</td>
            <td class="spreadsheet-cell p-2 border-b border-r border-gray-100 outline-none select-none text-gray-700" data-row="${i}" data-col="3">${c.cont||''}</td>
            <td class="spreadsheet-cell p-2 border-b border-gray-100 outline-none select-none text-gray-700" data-row="${i}" data-col="4">${c.email||''}</td>
        `;
        tbody.appendChild(tr);
    }
    
    // Attach Spreadsheet Listeners for the new rows
    setupSpreadsheetListeners(tbody);

    if(task) {
        ['f_moe_code','f_closing_date','f_school','f_assignment','f_programme','f_specs','f_folder','f_trainers','f_cost_specs','f_cost_val','f_notes','f_appointment'].forEach(id => { const el = document.getElementById(id); if(el) el.value = task[id.replace('f_','')]||''; });
        document.getElementById('costing-details').classList.toggle('hidden', !task.costing_type);
    } 
    document.getElementById('modal-overlay').classList.remove('hidden');
}

// --- ADVANCED SPREADSHEET LOGIC ---
function setupSpreadsheetListeners(tbody) {
    const cells = tbody.querySelectorAll('.spreadsheet-cell');
    
    cells.forEach(cell => {
        // 1. Selection Start (MouseDown)
        cell.addEventListener('mousedown', (e) => {
            if(cell.isContentEditable) return; // Allow interaction if editing
            isDragging = true;
            clearSelection();
            selectCell(cell);
            selectionStart = getCellCoords(cell);
            selectionEnd = selectionStart;
            cell.classList.add('active-anchor');
        });

        // 2. Drag Selection (MouseOver)
        cell.addEventListener('mouseover', (e) => {
            if (isDragging) {
                selectionEnd = getCellCoords(cell);
                updateSelectionRange(tbody);
            }
        });

        // 3. Edit Mode (Double Click)
        cell.addEventListener('dblclick', (e) => {
            makeEditable(cell);
        });
        
        // Handle Blur (Stop Editing)
        cell.addEventListener('blur', () => {
            cell.contentEditable = "false";
            cell.classList.remove('editing');
        });
    });

    // 4. End Selection (MouseUp on Table)
    tbody.addEventListener('mouseup', () => {
        isDragging = false;
    });
}

function getCellCoords(cell) {
    return { r: parseInt(cell.dataset.row), c: parseInt(cell.dataset.col) };
}

function clearSelection() {
    document.querySelectorAll('.spreadsheet-cell.selected, .spreadsheet-cell.active-anchor').forEach(c => {
        c.classList.remove('selected', 'active-anchor');
    });
}

function selectCell(cell) {
    cell.classList.add('selected');
}

function updateSelectionRange(tbody) {
    clearSelection();
    if (!selectionStart || !selectionEnd) return;

    const rMin = Math.min(selectionStart.r, selectionEnd.r);
    const rMax = Math.max(selectionStart.r, selectionEnd.r);
    const cMin = Math.min(selectionStart.c, selectionEnd.c);
    const cMax = Math.max(selectionStart.c, selectionEnd.c);

    for (let r = rMin; r <= rMax; r++) {
        for (let c = cMin; c <= cMax; c++) {
            const cell = tbody.querySelector(`.spreadsheet-cell[data-row="${r}"][data-col="${c}"]`);
            if (cell) cell.classList.add('selected');
        }
    }
}

function makeEditable(cell) {
    cell.contentEditable = "true";
    cell.focus();
    cell.classList.add('editing');
    cell.classList.remove('selected');
    // Select all text
    const range = document.createRange();
    range.selectNodeContents(cell);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
}

// Keyboard Navigation & Shortcuts
document.addEventListener('keydown', (e) => {
    // Only capture if we are in the spreadsheet modal and not editing
    const activeModal = !document.getElementById('modal-overlay').classList.contains('hidden');
    if (!activeModal) return;

    const selected = document.querySelector('.spreadsheet-cell.selected');
    if (!selected && !document.querySelector('.spreadsheet-cell.active-anchor')) return;
    
    // If editing, allow default behavior (except maybe Enter/Tab)
    if (document.activeElement.classList.contains('spreadsheet-cell') && document.activeElement.isContentEditable) {
        if (e.key === 'Enter') {
            e.preventDefault();
            document.activeElement.blur();
            // Move down?
        }
        return; 
    }

    const currentCell = document.querySelector('.spreadsheet-cell.active-anchor') || selected;
    if (!currentCell) return;
    
    const { r, c } = getCellCoords(currentCell);
    let nextR = r, nextC = c;

    if (e.key === 'ArrowUp') nextR--;
    else if (e.key === 'ArrowDown') nextR++;
    else if (e.key === 'ArrowLeft') nextC--;
    else if (e.key === 'ArrowRight') nextC++;
    else if (e.key === 'Enter') { 
        e.preventDefault(); 
        makeEditable(currentCell); 
        return; 
    }
    else { return; } // Allow other keys

    // Find next cell
    const nextCell = document.querySelector(`.spreadsheet-cell[data-row="${nextR}"][data-col="${nextC}"]`);
    if (nextCell) {
        e.preventDefault();
        clearSelection();
        selectCell(nextCell);
        nextCell.classList.add('active-anchor');
        selectionStart = { r: nextR, c: nextC };
        selectionEnd = selectionStart;
        nextCell.scrollIntoView({ block: 'nearest' });
    }
});

// Advanced Copy/Paste Logic
document.addEventListener('copy', (e) => {
    const activeModal = !document.getElementById('modal-overlay').classList.contains('hidden');
    if (!activeModal) return;
    
    const selectedCells = document.querySelectorAll('.spreadsheet-cell.selected');
    if (selectedCells.length === 0) return;

    e.preventDefault();

    // Determine bounds
    let rows = new Map();
    selectedCells.forEach(cell => {
        const r = parseInt(cell.dataset.row);
        if(!rows.has(r)) rows.set(r, []);
        rows.get(r).push({ c: parseInt(cell.dataset.col), text: cell.innerText });
    });

    const sortedRowKeys = [...rows.keys()].sort((a,b) => a - b);
    let textOutput = [];

    sortedRowKeys.forEach(r => {
        const cells = rows.get(r).sort((a,b) => a.c - b.c);
        textOutput.push(cells.map(c => c.text).join('\t'));
    });

    e.clipboardData.setData('text/plain', textOutput.join('\n'));
});

document.addEventListener('paste', (e) => {
    const activeModal = !document.getElementById('modal-overlay').classList.contains('hidden');
    if (!activeModal) return;

    const anchor = document.querySelector('.spreadsheet-cell.active-anchor') || document.querySelector('.spreadsheet-cell.selected');
    if (!anchor && !document.activeElement.classList.contains('spreadsheet-cell')) return;

    // If currently editing a cell, let default paste happen
    if (document.activeElement.isContentEditable) return;

    e.preventDefault();
    const text = (e.clipboardData || window.clipboardData).getData('text');
    const rows = text.split(/\r\n|\n|\r/).filter(r => r.trim()); // Split rows

    let startCoords = anchor ? getCellCoords(anchor) : { r: 1, c: 0 };
    
    rows.forEach((rowText, rOffset) => {
        const cols = rowText.split('\t');
        cols.forEach((cellText, cOffset) => {
            const targetR = startCoords.r + rOffset;
            const targetC = startCoords.c + cOffset;
            const targetCell = document.querySelector(`.spreadsheet-cell[data-row="${targetR}"][data-col="${targetC}"]`);
            if (targetCell) {
                targetCell.innerText = cellText.trim();
            }
        });
    });
});

document.getElementById('close-modal-btn').addEventListener('click', () => document.getElementById('modal-overlay').classList.add('hidden'));
document.getElementById('cancel-btn').addEventListener('click', () => document.getElementById('modal-overlay').classList.add('hidden'));
document.getElementById('f_school').addEventListener('change', (e) => { const found = allSchools.find(s => s.school === e.target.value); document.getElementById('f_assignment').value = found ? `${found.consultant}, ${found.zone}` : ''; });
window.toggleCostingDetails = (val) => document.getElementById('costing-details').classList.toggle('hidden', !val);

// CONFIRMATION DIALOG LOGIC
// 1. Reset/Setup Confirm Modal Logic
const resetModalState = () => {
    const cancelBtn = document.getElementById('confirm-cancel-btn');
    const yesBtn = document.getElementById('confirm-yes-btn');
    const iconContainer = document.getElementById('confirm-icon-container');

    // Default "Delete" Style
    cancelBtn.classList.remove('hidden');
    yesBtn.textContent = "Delete";
    yesBtn.className = "flex-1 px-4 py-2.5 rounded-xl text-white bg-red-500 hover:bg-red-600 font-bold shadow-lg transition transform active:scale-95";
    iconContainer.className = "w-12 h-12 rounded-full bg-red-100 text-red-500 flex items-center justify-center mb-4 mx-auto transition-colors";
    iconContainer.innerHTML = '<svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"></path></svg>';
};

window.showConfirm = (title, msg, callback) => {
    resetModalState(); // Ensure it's in Delete mode
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

// 2. New Alert Modal Logic (Reuses Confirm Modal structure)
window.showAlert = (title, msg) => {
    const modal = document.getElementById('confirm-modal');
    const content = document.getElementById('confirm-modal-content');
    const cancelBtn = document.getElementById('confirm-cancel-btn');
    const yesBtn = document.getElementById('confirm-yes-btn');
    const iconContainer = document.getElementById('confirm-icon-container');

    // "Alert" Style
    cancelBtn.classList.add('hidden');
    yesBtn.textContent = "OK";
    yesBtn.className = "flex-1 px-4 py-2.5 rounded-xl text-white bg-ios_blue hover:bg-blue-600 font-bold shadow-lg transition transform active:scale-95";
    iconContainer.className = "w-12 h-12 rounded-full bg-blue-100 text-ios_blue flex items-center justify-center mb-4 mx-auto transition-colors";
    iconContainer.innerHTML = '<svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>';

    document.getElementById('confirm-title').textContent = title;
    document.getElementById('confirm-msg').textContent = msg;
    pendingConfirmCallback = null; // No callback for alerts
    
    modal.classList.remove('hidden');
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
            // GRANULAR LOGGING LOGIC
            const old = allTasks.find(t => t.id === editingTaskId);
            const changes = [];
            
            if (old) {
                // Normalization helper (handle null/undefined vs "")
                const norm = (val) => (val || '').trim();

                // Direct Field Checks
                if (norm(old.status) !== norm(data.status)) changes.push(`Status to ${data.status}`);
                if (norm(old.closing_date) !== norm(data.closing_date)) changes.push(`Closing Date to ${formatDate(data.closing_date)}`);
                if (norm(old.programme) !== norm(data.programme)) changes.push('Programme Name');
                if (norm(old.assignment) !== norm(data.assignment)) changes.push('Assignment');
                if (norm(old.brand) !== norm(data.brand)) changes.push(`Brand to ${data.brand}`);
                if (norm(old.notes) !== norm(data.notes)) changes.push('Notes');
                if (norm(old.cost_val) !== norm(data.cost_val)) changes.push('Cost Value');
                if (norm(old.cost_specs) !== norm(data.cost_specs)) changes.push('Cost Specs');
                if (norm(old.costing_type) !== norm(data.costing_type)) changes.push('Costing Type');
                if (norm(old.trainers) !== norm(data.trainers)) changes.push('Trainers Info'); 
                if (norm(old.specs) !== norm(data.specs)) changes.push('Specs Link');
                if (norm(old.folder) !== norm(data.folder)) changes.push('Folder Link');
                if (norm(old.moe_code) !== norm(data.moe_code)) changes.push('MOE Code');
                if (norm(old.appointment) !== norm(data.appointment)) changes.push('Appointment');
                
                // Robust Contact Check
                // We compare field-by-field to avoid false positives from JSON structure mismatches
                let contactsChanged = false;
                for(let i=1; i<=4; i++) {
                    const k = `contact${i}`;
                    const oldC = old[k] || {};
                    const newC = data[k]; // Guaranteed to exist from form logic above
                    
                    const hasChange = 
                        norm(oldC.name) !== norm(newC.name) ||
                        norm(oldC.des) !== norm(newC.des) ||
                        norm(oldC.dept) !== norm(newC.dept) ||
                        norm(oldC.cont) !== norm(newC.cont) ||
                        norm(oldC.email) !== norm(newC.email);
                        
                    if(hasChange) {
                        contactsChanged = true;
                        break; 
                    }
                }
                if (contactsChanged) changes.push('Contact Details');
            }

            let logMsg = changes.length > 0 ? `Updated ${schoolName}: Changed ${changes.join(', ')}` : `Updated ${schoolName} (No major changes detected)`;

            await updateDoc(doc(collection(db, `artifacts/${appId}/public/data/tasks`), editingTaskId), data); 
            logAction(logMsg);
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
    const sName = document.getElementById('sf_school').value.trim();
    
    // --- DUPLICATE SCHOOL CHECK ---
    // Check if school name exists (case-insensitive) and we are NOT editing that same school
    const exists = allSchools.some(s => s.school.toLowerCase().trim() === sName.toLowerCase() && s.id !== editingSchoolId);
    
    if (exists) {
        window.showAlert("Duplicate School", `The school "${sName}" already exists. Please update the existing entry instead.`);
        return;
    }

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
    
    // UPDATE FAB BUTTON DYNAMICALLY BASED ON VIEW
    const fabAddBtn = document.getElementById('fab-add-btn');
    const fabAddText = document.getElementById('fab-add-text');
    const fabImportBtn = document.getElementById('fab-import-btn'); // NEW

    if (view === 'zones') {
        fabAddText.textContent = "Add School";
        fabAddBtn.setAttribute('onclick', "window.openSchoolModal(null)");
        if(fabImportBtn) fabImportBtn.classList.remove('hidden'); // Show Import
        if(fabImportBtn) fabImportBtn.classList.add('flex');
    } else {
        fabAddText.textContent = "Add Task";
        fabAddBtn.setAttribute('onclick', "window.openModal(null)");
        if(fabImportBtn) fabImportBtn.classList.add('hidden'); // Hide Import
        if(fabImportBtn) fabImportBtn.classList.remove('flex');
    }

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
    // Updated Logging for Progress
    logAction(`Updated progress for ${t.school}: ${val ? 'Completed' : 'Unchecked'} ${itemName}`);
    
    // If View Modal is open, update checkboxes visually without full reload?
    // Not strictly necessary as Firebase snapshot listener calls refreshView/openViewModal logic,
    // but checkbox state will update automatically via user interaction.
    
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

    if (currentView === 'zones') {
        // --- ZONE EXPORT LOGIC (Count - School - Zone - Consultant) ---
        const headers = ["No.", "School Name", "Zone", "Consultant"];
        
        // Replicate the current sort logic from renderZonesTable for consistency
        let schoolsToExport = [...allSchools];
        schoolsToExport.sort((a, b) => { 
            const valA = (a[schoolSort.key]||'').toLowerCase(); 
            const valB = (b[schoolSort.key]||'').toLowerCase(); 
            return valA < valB ? (schoolSort.dir==='asc'?-1:1) : (valA > valB ? (schoolSort.dir==='asc'?1:-1) : 0); 
        });

        const rows = schoolsToExport.map((s, index) => {
            return [
                index + 1,
                s.school,
                s.zone,
                s.consultant
            ].map(escapeCsv).join(',');
        });

        const csvContent = [headers.join(','), ...rows].join('\n');
        downloadCSV(csvContent, 'schools_zone_export');

    } else {
        // --- EXISTING TASK EXPORT LOGIC ---
        const headers = [
            "Export Date", "Item Code", "Closing Date", "Zone", "Consultant", "School", 
            "Contact Person 1 Name", "Contact Person 1 Designation", "Contact Person 1 Department", "Contact Person 1 Contact Number", "Contact Person 1 Email", 
            "Contact Person 2 Name", "Contact Person 2 Designation", "Contact Person 2 Department", "Contact Person 2 Contact Number", "Contact Person 2 Email", 
            "Programme Name", "Brand", "Status", "GeBIZ Updated", "Email Follow-up"
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
                t.status, 
                t.progress?.gebiz ? 'Yes' : 'No', // New Col
                t.progress?.email ? 'Yes' : 'No'  // New Col
            ].map(escapeCsv).join(',');
        });
    
        const csvContent = [headers.join(','), ...rows].join('\n');
        downloadCSV(csvContent, 'checklist_export');
    }
    
    // Close FAB menu after export
    window.toggleFabMenu();
};

const downloadCSV = (content, filenamePrefix) => {
    const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `${filenamePrefix}_${new Date().toISOString().slice(0,10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

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
