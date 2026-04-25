let currentAssembly = null;
let isNewMode = false;
let isEditMode = false;
let counterData = null;
let nextNumber = null;
let selectedClientName = null;
let searchDebounce = null;
let clientDebounce = null;

// ── Init ─────────────────────────────────────────────────────

async function init() {
    // If this is the MSAL popup window, just process the redirect and stop
    if (window.opener && window.opener !== window) {
        await msalInstance.handleRedirectPromise();
        return;
    }
    await msalInstance.handleRedirectPromise();
    const user = await getCurrentUser();
    user ? showApp(user) : showLogin();
}

function showLogin() {
    document.getElementById('login-screen').classList.remove('hidden');
    document.getElementById('app').classList.add('hidden');

    document.getElementById('login-btn').addEventListener('click', async () => {
        try {
            const account = await login();
            showApp(account);
        } catch (e) {
            const err = document.getElementById('login-error');
            err.textContent = 'Login failed. Please try again.';
            err.classList.remove('hidden');
        }
    });
}

function showApp(user) {
    document.getElementById('login-screen').classList.add('hidden');
    document.getElementById('app').classList.remove('hidden');
    document.getElementById('user-name').textContent = user.name || user.username;
    bindEvents();
}

// ── Events ───────────────────────────────────────────────────

function bindEvents() {
    // Search
    const searchInput = document.getElementById('search-input');
    searchInput.addEventListener('input', e => {
        clearTimeout(searchDebounce);
        searchDebounce = setTimeout(() => runSearch(e.target.value), 350);
    });
    searchInput.addEventListener('keydown', e => {
        if (e.key === 'Enter') { clearTimeout(searchDebounce); runSearch(e.target.value); }
    });
    document.getElementById('search-btn').addEventListener('click', () => runSearch(searchInput.value));

    // Toolbar
    document.getElementById('new-btn').addEventListener('click', openNewForm);
    document.getElementById('logout-btn').addEventListener('click', logout);

    // Detail actions
    document.getElementById('edit-btn').addEventListener('click', openEditForm);
    document.getElementById('delete-btn').addEventListener('click', showDeleteModal);

    // Form actions
    document.getElementById('save-btn').addEventListener('click', saveAssembly);
    document.getElementById('cancel-btn').addEventListener('click', cancelForm);

    // Assembly type → preview number
    document.getElementById('f-assembly-type').addEventListener('change', previewAssemblyNumber);

    // Client autocomplete
    document.getElementById('f-client-search').addEventListener('input', e => {
        selectedClientName = null;
        clearTimeout(clientDebounce);
        clientDebounce = setTimeout(() => runClientSearch(e.target.value), 300);
    });

    // Delete modal
    document.getElementById('confirm-delete-btn').addEventListener('click', confirmDelete);
    document.getElementById('cancel-delete-btn').addEventListener('click', hideDeleteModal);
    document.getElementById('modal-overlay').addEventListener('click', hideDeleteModal);

    // Close client suggestions on outside click
    document.addEventListener('click', e => {
        if (!e.target.closest('.client-group')) {
            document.getElementById('client-suggestions').classList.add('hidden');
        }
    });
}

// ── Search ───────────────────────────────────────────────────

async function runSearch(query) {
    query = (query || '').trim();
    const list = document.getElementById('assembly-list');

    if (query.length < 2) {
        list.innerHTML = '<div class="list-empty">Enter at least 2 characters to search</div>';
        return;
    }

    list.innerHTML = '<div class="loading">Searching...</div>';

    try {
        const items = await searchAssemblies(query);
        renderList(items);
    } catch (e) {
        list.innerHTML = `<div class="list-empty">Error: ${e.message}</div>`;
    }
}

function renderList(items) {
    const list = document.getElementById('assembly-list');

    if (!items.length) {
        list.innerHTML = '<div class="list-empty">No assemblies found</div>';
        return;
    }

    list.innerHTML = items.map(item => {
        const f = item.fields;
        return `<div class="list-item" data-id="${item.id}" onclick="selectAssembly('${item.id}')">
            <div class="list-item-number">${esc(f.AssemblyNumber)}</div>
            <div class="list-item-title">${esc(f.Title)}</div>
            <span class="list-item-badge">${esc(f.AssemblyType)}</span>
        </div>`;
    }).join('');

    // Re-highlight if current assembly is in results
    if (currentAssembly) highlightListItem(currentAssembly.id);
}

// ── Detail view ──────────────────────────────────────────────

async function selectAssembly(itemId) {
    highlightListItem(itemId);
    try {
        const item = await getAssembly(itemId);
        currentAssembly = item;
        renderDetail(item);
    } catch (e) {
        showToast('Error loading assembly: ' + e.message, 'error');
    }
}

function renderDetail(item) {
    const f = item.fields;
    showPanel('detail');

    document.getElementById('detail-title').textContent = f.AssemblyNumber || '';
    document.getElementById('detail-content').innerHTML = `
        <div class="detail-field">
            <label>Assembly Type</label>
            <p>${esc(f.AssemblyType)}</p>
        </div>
        <div class="detail-field">
            <label>Revision</label>
            <p>${esc(f.Revision)}</p>
        </div>
        <div class="detail-field" style="grid-column:1/-1">
            <label>Description</label>
            <p>${esc(f.Title)}</p>
        </div>
        <div class="detail-field">
            <label>Client P/N</label>
            <p>${esc(f.ClientP_x002f_N)}</p>
        </div>
        <div class="detail-field">
            <label>Client</label>
            <p>${esc(f.ClientName)}</p>
        </div>
    `;
}

// ── New / Edit form ───────────────────────────────────────────

async function openNewForm() {
    isNewMode = true;
    isEditMode = false;
    currentAssembly = null;
    selectedClientName = null;

    document.getElementById('form-title').textContent = 'New Assembly';
    document.getElementById('assembly-form').reset();
    document.getElementById('f-assembly-number').value = '';
    document.getElementById('client-suggestions').classList.add('hidden');

    showPanel('form');

    // Pre-fetch counter in background
    try {
        counterData = await getCounter();
        nextNumber = counterData.fields.CurrentNumber + 1;
        previewAssemblyNumber();
    } catch (e) {
        showToast('Could not load counter: ' + e.message, 'error');
    }
}

function openEditForm() {
    if (!currentAssembly) return;
    const f = currentAssembly.fields;

    isEditMode = true;
    isNewMode = false;
    selectedClientName = f.ClientName || null;

    document.getElementById('form-title').textContent = 'Edit Assembly';
    document.getElementById('f-assembly-number').value = f.AssemblyNumber || '';
    document.getElementById('f-assembly-type').value   = f.AssemblyType   || '';
    document.getElementById('f-title').value           = f.Title          || '';
    document.getElementById('f-revision').value        = f.Revision       || '';
    document.getElementById('f-client-pn').value       = f.ClientP_x002f_N || '';
    document.getElementById('f-client-search').value   = f.ClientName     || '';
    document.getElementById('client-suggestions').classList.add('hidden');

    showPanel('form');
}

function previewAssemblyNumber() {
    if (!isNewMode) return;
    const type = document.getElementById('f-assembly-type').value;
    document.getElementById('f-assembly-number').value =
        (type && nextNumber) ? `${type}-${nextNumber}` : '';
}

function cancelForm() {
    isNewMode = false;
    isEditMode = false;
    currentAssembly ? renderDetail(currentAssembly) : showPanel('empty');
}

async function saveAssembly() {
    const type    = document.getElementById('f-assembly-type').value;
    const title   = document.getElementById('f-title').value.trim();
    const rev     = document.getElementById('f-revision').value.trim();
    const clientPN = document.getElementById('f-client-pn').value.trim();
    const clientName = selectedClientName || document.getElementById('f-client-search').value.trim();

    if (!title) { showToast('Description is required', 'error'); return; }
    if (isNewMode && !type) { showToast('Assembly Type is required', 'error'); return; }

    const btn = document.getElementById('save-btn');
    btn.textContent = 'Saving...';
    btn.disabled = true;

    try {
        if (isNewMode) {
            const newNum = counterData.fields.CurrentNumber + 1;
            const assemblyNumber = `${type}-${newNum}`;

            await updateCounter(counterData.id, newNum);
            await createAssembly({
                Title: title,
                AssemblyNumber: assemblyNumber,
                AssemblyType: type,
                Revision: rev,
                ClientP_x002f_N: clientPN,
                ClientName: clientName
            });

            showToast(`Created ${assemblyNumber}`, 'success');

        } else if (isEditMode && currentAssembly) {
            await updateAssembly(currentAssembly.id, {
                Title: title,
                AssemblyType: type,
                Revision: rev,
                ClientP_x002f_N: clientPN,
                ClientName: clientName
            });

            // Refresh detail
            currentAssembly = await getAssembly(currentAssembly.id);
            showToast('Assembly updated', 'success');
        }

        // Refresh search list
        const q = document.getElementById('search-input').value;
        if (q.trim().length >= 2) runSearch(q);

        cancelForm();

    } catch (e) {
        showToast('Error saving: ' + e.message, 'error');
    } finally {
        btn.textContent = 'Save';
        btn.disabled = false;
    }
}

// ── Client autocomplete ───────────────────────────────────────

async function runClientSearch(query) {
    const box = document.getElementById('client-suggestions');
    if (!query || query.length < 2) { box.classList.add('hidden'); return; }

    try {
        const results = await searchClients(query);
        if (!results.length) { box.classList.add('hidden'); return; }

        box.innerHTML = results.map(r =>
            `<div class="suggestion-item" onmousedown="pickClient(event,'${r.id}','${escAttr(r.fields.Title)}')">${esc(r.fields.Title)}</div>`
        ).join('');
        box.classList.remove('hidden');
    } catch (e) {
        box.classList.add('hidden');
    }
}

function pickClient(e, id, name) {
    e.preventDefault();
    selectedClientName = name;
    document.getElementById('f-client-search').value = name;
    document.getElementById('client-suggestions').classList.add('hidden');
}

// ── Delete ────────────────────────────────────────────────────

function showDeleteModal() {
    if (!currentAssembly) return;
    document.getElementById('delete-assembly-name').textContent =
        currentAssembly.fields.AssemblyNumber || 'this assembly';
    document.getElementById('delete-modal').classList.remove('hidden');
}

function hideDeleteModal() {
    document.getElementById('delete-modal').classList.add('hidden');
}

async function confirmDelete() {
    if (!currentAssembly) return;
    try {
        await deleteAssembly(currentAssembly.id);
        hideDeleteModal();

        const q = document.getElementById('search-input').value;
        if (q.trim().length >= 2) runSearch(q);

        currentAssembly = null;
        showPanel('empty');
        showToast('Assembly deleted', 'success');
    } catch (e) {
        hideDeleteModal();
        showToast('Error deleting: ' + e.message, 'error');
    }
}

// ── Helpers ───────────────────────────────────────────────────

function showPanel(name) {
    document.getElementById('empty-state').classList.add('hidden');
    document.getElementById('detail-view').classList.add('hidden');
    document.getElementById('form-view').classList.add('hidden');

    if (name === 'detail') document.getElementById('detail-view').classList.remove('hidden');
    else if (name === 'form')  document.getElementById('form-view').classList.remove('hidden');
    else                       document.getElementById('empty-state').classList.remove('hidden');
}

function highlightListItem(itemId) {
    document.querySelectorAll('.list-item').forEach(el => el.classList.remove('selected'));
    document.querySelector(`.list-item[data-id="${itemId}"]`)?.classList.add('selected');
}

function esc(val) {
    if (val == null) return '—';
    return String(val)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function escAttr(val) {
    return String(val || '').replace(/'/g, "\\'").replace(/"/g, '&quot;');
}

let toastTimer = null;
function showToast(msg, type = '') {
    const t = document.getElementById('toast');
    t.textContent = msg;
    t.className = 'toast' + (type ? ` ${type}` : '');
    t.classList.remove('hidden');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.add('hidden'), 3500);
}

// ── Start ─────────────────────────────────────────────────────
init();
