const GRAPH = 'https://graph.microsoft.com/v1.0';
const SHAREPOINT_HOST = 'fpeautomation0.sharepoint.com';
const SITE_PATH = '/sites/FPEAutomation';

let _siteId = null;

async function graphFetch(url, options = {}) {
    const token = await getToken();
    const res = await fetch(url, {
        ...options,
        headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json',
            ...options.headers
        }
    });

    if (res.status === 204) return null;

    const data = await res.json();

    if (!res.ok) {
        throw new Error(data?.error?.message || `Request failed (${res.status})`);
    }

    return data;
}

async function getSiteId() {
    if (_siteId) return _siteId;
    const data = await graphFetch(`${GRAPH}/sites/${SHAREPOINT_HOST}:${SITE_PATH}`);
    _siteId = data.id;
    return _siteId;
}

// ── Assemblies ──────────────────────────────────────────────

async function searchAssemblies(query) {
    const siteId = await getSiteId();
    const encoded = encodeURIComponent(query);
    const url = `${GRAPH}/sites/${siteId}/lists/Assemblies/items` +
        `?$search="${encoded}"` +
        `&$expand=fields($select=Title,AssemblyNumber,AssemblyType,Revision,ClientP_x002f_N,ClientName)` +
        `&$top=100`;
    const data = await graphFetch(url);
    return data.value || [];
}

async function getAssembly(itemId) {
    const siteId = await getSiteId();
    const url = `${GRAPH}/sites/${siteId}/lists/Assemblies/items/${itemId}?$expand=fields`;
    return graphFetch(url);
}

async function createAssembly(fields) {
    const siteId = await getSiteId();
    const url = `${GRAPH}/sites/${siteId}/lists/Assemblies/items`;
    return graphFetch(url, { method: 'POST', body: JSON.stringify({ fields }) });
}

async function updateAssembly(itemId, fields) {
    const siteId = await getSiteId();
    const url = `${GRAPH}/sites/${siteId}/lists/Assemblies/items/${itemId}`;
    return graphFetch(url, { method: 'PATCH', body: JSON.stringify({ fields }) });
}

async function deleteAssembly(itemId) {
    const siteId = await getSiteId();
    const url = `${GRAPH}/sites/${siteId}/lists/Assemblies/items/${itemId}`;
    return graphFetch(url, { method: 'DELETE' });
}

// ── Counter ──────────────────────────────────────────────────

async function getCounter() {
    const siteId = await getSiteId();
    const url = `${GRAPH}/sites/${siteId}/lists/Assembly Counter/items?$expand=fields`;
    const data = await graphFetch(url);
    return data.value[0];
}

async function updateCounter(itemId, newNumber) {
    const siteId = await getSiteId();
    const url = `${GRAPH}/sites/${siteId}/lists/Assembly Counter/items/${itemId}`;
    return graphFetch(url, {
        method: 'PATCH',
        body: JSON.stringify({ fields: { CurrentNumber: newNumber } })
    });
}

// ── Clients ──────────────────────────────────────────────────

async function searchClients(query) {
    if (!query || query.length < 2) return [];
    const siteId = await getSiteId();
    const encoded = encodeURIComponent(query);
    const url = `${GRAPH}/sites/${siteId}/lists/Assembly Drawing Client List/items` +
        `?$search="${encoded}"` +
        `&$expand=fields($select=Title,FPEAccount)` +
        `&$top=20`;
    const data = await graphFetch(url);
    return data.value || [];
}
