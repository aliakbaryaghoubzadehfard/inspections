/* Rincovitch Site Inspections — Microsoft 365 edition.
   Sign-in with company Microsoft accounts (OAuth 2.0 auth code + PKCE, no secrets in the app),
   data kept in a shared SharePoint / OneDrive folder through Microsoft Graph,
   a local copy on the phone so it keeps working with no signal, and an outbox that syncs when signal returns.
   Exposes the same window.claude.use('db' | 'assets' | 'downloads') interface the app is written against. */
(function () {
  'use strict';
  const CFG = Object.assign({
    authority: 'https://login.microsoftonline.com', graph: 'https://graph.microsoft.com/v1.0',
    tenantId: 'organizations', clientId: '', dataFolderLink: '', dataFolderPath: '', registerLink: '', registerPath: '', registerInFolder: '', registerTable: 'InspectionRecords',
    registerSheet: 'Inspection Register', reportsFolder: 'Inspections', pollSeconds: 30,
  }, window.APP_CONFIG || {});
  const SCOPES = 'openid profile offline_access User.Read Files.ReadWrite.All';
  const MAIL_SCOPE = 'Mail.Send'; // asked for only when someone first emails a report
  const unset = (v) => !v || /PASTE-/.test(v);
  if (unset(CFG.clientId) || unset(CFG.tenantId)) CFG.clientId = '';
  if (unset(CFG.dataFolderLink)) CFG.dataFolderLink = '';
  if (unset(CFG.registerLink)) CFG.registerLink = '';
  if (unset(CFG.dataFolderPath)) CFG.dataFolderPath = '';
  if (unset(CFG.registerPath)) CFG.registerPath = '';
  if (unset(CFG.registerInFolder)) CFG.registerInFolder = '';
  const hasRegister = () => !!(CFG.registerLink || CFG.registerPath || CFG.registerInFolder);
  const mePath = (p) => '/me/drive/root:/' + p.replace(/^\/+|\/+$/g, '').split('/').map(encodeURIComponent).join('/') + ':';
  /* ================= projects: each project is its own data folder, with its own copy on this phone ================= */
  const PKEY = 'm365.projects', AKEY = 'm365.project';
  const lsJ = (k, d) => { try { const v = JSON.parse(localStorage.getItem(k) || 'null'); return v == null ? d : v; } catch (e) { return d; } };
  const fromConfig = (CFG.dataFolderLink || CFG.dataFolderPath) ? { id: 'default', name: CFG.projectName || '', link: CFG.dataFolderLink, path: CFG.dataFolderPath, register: CFG.registerInFolder || '', registerLink: CFG.registerLink || '', registerPath: CFG.registerPath || '', fromConfig: true } : null;
  let projects = lsJ(PKEY, []).filter((p) => p && p.id && (p.id !== 'default' || fromConfig));
  if (fromConfig) { const i = projects.findIndex((p) => p.id === 'default'); if (i < 0) projects.unshift(fromConfig); else projects[i] = { ...fromConfig, name: projects[i].name || fromConfig.name }; }
  let active = projects.find((p) => p.id === lsJ(AKEY, null)) || projects[0] || null;
  const b64uJ = (o) => btoa(unescape(encodeURIComponent(JSON.stringify(o)))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  const unb64uJ = (s) => JSON.parse(decodeURIComponent(escape(atob(s.replace(/-/g, '+').replace(/_/g, '/')))));
  const joinParam = new URLSearchParams(location.search).get('join');
  if (joinParam) {   // opened from an invitation: add that project and open it
    try {
      const j = unb64uJ(joinParam);
      if (j && j.drive && j.item) { let p = projects.find((x) => x.drive === j.drive && x.item === j.item); if (!p) { p = { id: 'p' + Date.now().toString(36), name: j.name || 'Project', drive: j.drive, item: j.item, register: j.register || '' }; projects.push(p); } active = p; }
    } catch (e) { console.warn('bad join link', e); }
    history.replaceState(null, '', location.pathname);
  }
  const saveProjects = () => { try { localStorage.setItem(PKEY, JSON.stringify(projects)); localStorage.setItem(AKEY, JSON.stringify(active ? active.id : null)); } catch (e) {} };
  saveProjects();
  CFG.dataFolderItem = null;
  if (active) { CFG.dataFolderLink = active.link || ''; CFG.dataFolderPath = active.path || ''; CFG.dataFolderItem = active.drive && active.item ? { drive: active.drive, id: active.item } : null; CFG.registerInFolder = active.register || ''; CFG.registerLink = active.registerLink || ''; CFG.registerPath = active.registerPath || ''; }
  else { CFG.dataFolderLink = ''; CFG.dataFolderPath = ''; }
  const DB_NAME = 'rincovitch-inspections-m365' + (active && active.id !== 'default' ? '-' + active.id : ''), VER = 1;
  const COLLS = ['settings', 'types', 'inspections'];
  const $ = (s) => document.querySelector(s);
  const escH = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  /* ================= local storage (IndexedDB) ================= */
  let idb = null, memMode = false, storageErr = '';
  const open = () => idb || (idb = new Promise((res, rej) => {
    // some browsers block on-device storage (site data/cookies blocked, some private modes) or never answer
    const t = setTimeout(() => rej(new Error('On-device storage did not respond')), 5000);
    let r;
    try { r = indexedDB.open(DB_NAME, VER); } catch (e) { clearTimeout(t); rej(e); return; }
    r.onblocked = () => { clearTimeout(t); rej(new Error('On-device storage is locked by another copy of the app — close other tabs of it and reload')); };
    r.onupgradeneeded = () => {
      const d = r.result;
      d.createObjectStore('docs', { keyPath: 'path' });
      d.createObjectStore('blobs', { keyPath: 'id' });
      d.createObjectStore('outbox', { keyPath: 'seq', autoIncrement: true });
      d.createObjectStore('meta', { keyPath: 'k' });
    };
    r.onsuccess = () => { clearTimeout(t); res(r.result); };
    r.onerror = () => { clearTimeout(t); rej(r.error || new Error('Storage is not available on this device')); };
  }).catch((e) => { memMode = true; storageErr = (e && (e.message || e.name)) || 'blocked'; console.warn('Using memory storage:', e); return null; }));
  /* memory fallback with the same small API as IndexedDB object stores (online-only, nothing kept after closing) */
  const MEM = { docs: new Map(), blobs: new Map(), outbox: new Map(), meta: new Map() }; let memSeq = 0;
  const KEYS = { docs: 'path', blobs: 'id', outbox: 'seq', meta: 'k' };
  const fakeReq = (v) => { const r = { result: v }; setTimeout(() => r.onsuccess && r.onsuccess()); return r; };
  const memStore = (name) => { const m = MEM[name], k = KEYS[name]; return {
    get: (id) => fakeReq(m.get(id)), getAll: () => fakeReq([...m.values()]), count: () => fakeReq(m.size),
    put: (v) => { m.set(v[k], v); return fakeReq(v[k]); }, add: (v) => { const o = { ...v, [k]: ++memSeq }; m.set(o[k], o); return fakeReq(o[k]); },
    delete: (id) => { m.delete(id); return fakeReq(undefined); }, clear: () => { m.clear(); return fakeReq(undefined); } }; };
  const tx = async (stores, mode, fn) => {
    const d = await open();
    if (!d) { const out = await fn(Array.isArray(stores) ? stores.map(memStore) : memStore(stores)); return out; }
    return new Promise((res, rej) => {
      const t = d.transaction(stores, mode); let out;
      Promise.resolve(fn(Array.isArray(stores) ? stores.map((s) => t.objectStore(s)) : t.objectStore(stores))).then((v) => { out = v; });
      t.oncomplete = () => res(out); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error || new Error('Storage write failed (device full?)'));
    });
  };
  const reqP = (r) => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
  const metaGet = async (k) => { const r = await tx('meta', 'readonly', (st) => reqP(st.get(k))); return r ? r.v : undefined; };
  const metaSet = (k, v) => tx('meta', 'readwrite', (st) => st.put({ k, v }));

  /* ================= auth ================= */
  const TOK = 'm365.tokens';
  const b64url = (buf) => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  const rand = (n) => b64url(crypto.getRandomValues(new Uint8Array(n)));
  const redirectUri = () => location.origin + location.pathname;
  const lsGet = (k) => { try { return JSON.parse(localStorage.getItem(k) || 'null'); } catch (e) { return null; } };
  const lsSet = (k, v) => { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} };
  let tokens = lsGet(TOK);
  const claims = (idt) => { try { return JSON.parse(decodeURIComponent(escape(atob(idt.split('.')[1].replace(/-/g, '+').replace(/_/g, '/'))))); } catch (e) { return {}; } };
  let account = lsGet('m365.account');
  const scopes = () => SCOPES + (lsGet('m365.mail') ? ' ' + MAIL_SCOPE : '');
  const mailOK = () => !!(tokens && tokens.scope && /(^|\s)(https:\/\/graph\.microsoft\.com\/)?Mail\.Send(\s|$)/i.test(tokens.scope));

  async function signIn(opts) {
    if (!CFG.clientId) { showGate('setup'); return; }
    if (opts && opts.mail) { lsSet('m365.mail', true); lsSet('m365.mailask', true); }
    const verifier = rand(48);
    const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
    const state = rand(16);
    try { sessionStorage.setItem('m365.pkce', JSON.stringify({ verifier, state })); } catch (e) { lsSet('m365.pkce', { verifier, state }); }
    const q = new URLSearchParams({ client_id: CFG.clientId, response_type: 'code', redirect_uri: redirectUri(), response_mode: 'query',
      scope: scopes(), code_challenge: challenge, code_challenge_method: 'S256', state, ...(account && account.username ? { login_hint: account.username } : {}) });
    location.assign(`${CFG.authority}/${CFG.tenantId}/oauth2/v2.0/authorize?${q}`);
  }
  async function tokenRequest(body) {
    const r = await fetch(`${CFG.authority}/${CFG.tenantId}/oauth2/v2.0/token`, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams({ client_id: CFG.clientId, scope: scopes(), ...body }) });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) { const e = new Error(j.error_description || j.error || 'Sign-in failed'); e.code = j.error || 'auth'; throw e; }
    tokens = { access: j.access_token, refresh: j.refresh_token || (tokens && tokens.refresh), exp: Date.now() + (j.expires_in || 3600) * 1000, scope: j.scope || (tokens && tokens.scope) || '' };
    lsSet(TOK, tokens);
    if (j.id_token) { const c = claims(j.id_token); account = { name: c.name || '', username: (c.preferred_username || c.email || c.upn || '').toLowerCase(), oid: c.oid || '' }; lsSet('m365.account', account); }
    return tokens.access;
  }
  async function handleRedirect() {
    const p = new URLSearchParams(location.search);
    if (!p.has('code') && !p.has('error')) return;
    let saved = null; try { saved = JSON.parse(sessionStorage.getItem('m365.pkce') || 'null'); sessionStorage.removeItem('m365.pkce'); } catch (e) {}
    if (!saved) { saved = lsGet('m365.pkce'); lsSet('m365.pkce', null); }
    history.replaceState(null, '', location.pathname);
    const askedMail = lsGet('m365.mailask'); lsSet('m365.mailask', null);
    if (p.has('error')) {
      if (askedMail && tokens) { lsSet('m365.mail', null); setTimeout(() => notice('Email permission was not given, so queued emails were not sent. You can still use your phone\'s mail app.'), 800); return; }
      gateError = p.get('error_description') || p.get('error'); return;
    }
    if (!saved || saved.state !== p.get('state')) { gateError = 'Sign-in was interrupted. Please try again.'; return; }
    try { await tokenRequest({ grant_type: 'authorization_code', code: p.get('code'), redirect_uri: redirectUri(), code_verifier: saved.verifier }); }
    catch (e) { if (askedMail) lsSet('m365.mail', null); gateError = e.message; }
    if (askedMail && !mailOK()) lsSet('m365.mail', null);
  }
  let refreshing = null;
  async function accessToken(force) {
    if (!tokens) throw Object.assign(new Error('Not signed in'), { code: 'signin' });
    if (!force && tokens.access && tokens.exp - Date.now() > 90000) return tokens.access;
    if (!tokens.refresh) throw Object.assign(new Error('Sign in again'), { code: 'signin' });
    if (!refreshing) refreshing = tokenRequest({ grant_type: 'refresh_token', refresh_token: tokens.refresh })
      .catch((e) => { if (/invalid_grant|interaction_required/.test(e.code)) { tokens = null; lsSet(TOK, null); e.code = 'signin'; } throw e; })
      .finally(() => { refreshing = null; });
    return refreshing;
  }
  function signOut() {
    tokens = null; lsSet(TOK, null); const a = account; account = null; lsSet('m365.account', null);
    indexedDB.deleteDatabase(DB_NAME);
    location.assign(`${CFG.authority}/${CFG.tenantId}/oauth2/v2.0/logout?` + new URLSearchParams({ post_logout_redirect_uri: redirectUri(), ...(a && a.username ? { logout_hint: a.username } : {}) }));
  }

  /* ================= Microsoft Graph ================= */
  async function graph(path, opts = {}, retry = true) {
    const tok = await accessToken();
    const url = /^https?:/.test(path) ? path : CFG.graph + path;
    const r = await fetch(url, { ...opts, headers: { Authorization: 'Bearer ' + tok, ...(opts.headers || {}) } });
    if (r.status === 401 && retry) { await accessToken(true); return graph(path, opts, false); }
    if (r.status === 429 || r.status === 503) { const wait = (+r.headers.get('Retry-After') || 5) * 1000; await new Promise((res) => setTimeout(res, Math.min(wait, 30000))); if (retry) return graph(path, opts, false); }
    return r;
  }
  async function gjson(path, opts) {
    const r = await graph(path, opts);
    if (r.status === 404) return null;
    const j = r.status === 204 ? {} : await r.json().catch(() => ({}));
    if (!r.ok) { const e = new Error((j.error && j.error.message) || ('Microsoft 365 error ' + r.status)); e.status = r.status; e.code = j.error && j.error.code; throw e; }
    return j;
  }
  const shareId = (link) => 'u!' + btoa(unescape(encodeURIComponent(link.trim()))).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  let ROOT = null, REG = null;
  async function resolveRoot() {
    if (ROOT) return ROOT;
    const key = 'root:' + (CFG.dataFolderItem ? 'item:' + CFG.dataFolderItem.id : CFG.dataFolderLink || 'path:' + CFG.dataFolderPath);
    const cached = await metaGet(key);
    if (cached) return (ROOT = cached);
    if (!CFG.dataFolderLink && !CFG.dataFolderPath && !CFG.dataFolderItem) throw Object.assign(new Error('Choose or create a project first'), { code: 'noproject' });
    if (CFG.dataFolderItem) {
      const it = await gjson(`/drives/${CFG.dataFolderItem.drive}/items/${CFG.dataFolderItem.id}?$select=id,name,parentReference,folder`).catch((e) => { if (e.status === 403) return null; throw e; });
      if (!it || !it.folder) throw Object.assign(new Error(`You don't have access to ${active && active.name ? active.name : "this project's folder"} yet — ask the project owner to share it with you (Projects › Share).`), { code: 'noaccess' });
      ROOT = { drive: CFG.dataFolderItem.drive, id: it.id, name: it.name }; await metaSet(key, ROOT); return ROOT;
    }
    const it = CFG.dataFolderLink ? await gjson(`/shares/${shareId(CFG.dataFolderLink)}/driveItem?$select=id,name,parentReference,folder`)
      : await gjson(mePath(CFG.dataFolderPath) + '?$select=id,name,parentReference,folder');
    if (!it || !it.folder) throw Object.assign(new Error(CFG.dataFolderLink ? 'The data folder link in config.js does not point to a folder you can open.' : `Folder "${CFG.dataFolderPath}" was not found in your OneDrive — open a project shared with you, or start a new one.`), { code: 'noaccess' });
    ROOT = { drive: it.parentReference.driveId, id: it.id, name: it.name };
    await metaSet(key, ROOT);
    return ROOT;
  }
  async function resolveRegister() {
    if (!hasRegister()) return null;
    if (REG) return REG;
    const rkey = 'reg:' + (CFG.registerLink || (CFG.registerInFolder ? 'in:' + (CFG.dataFolderItem ? CFG.dataFolderItem.id : CFG.dataFolderLink + CFG.dataFolderPath) + '/' + CFG.registerInFolder : 'path:' + CFG.registerPath));
    const cached = await metaGet(rkey);
    if (cached) return (REG = cached);
    let it;
    if (CFG.registerLink) it = await gjson(`/shares/${shareId(CFG.registerLink)}/driveItem?$select=id,name,parentReference,file`);
    else if (CFG.registerInFolder) { await resolveRoot(); it = await gjson(itemPath(CFG.registerInFolder) + '?$select=id,name,parentReference,file'); if (it && !it.parentReference) it.parentReference = { driveId: ROOT.drive }; }
    else it = await gjson(mePath(CFG.registerPath) + '?$select=id,name,parentReference,file');
    if (!it || !it.file) throw new Error(CFG.registerLink ? 'The register link in config.js does not point to the workbook.' : `Workbook "${CFG.registerInFolder || CFG.registerPath}" was not found in the data folder.`);
    REG = { drive: it.parentReference.driveId, id: it.id, name: it.name };
    await metaSet(rkey, REG);
    return REG;
  }
  const encPath = (p) => p.split('/').map(encodeURIComponent).join('/');
  const itemPath = (rel) => `/drives/${ROOT.drive}/items/${ROOT.id}:/${encPath(rel)}:`;
  async function ensureFolder(rel) {
    const parts = rel.split('/').filter(Boolean); let parent = `/drives/${ROOT.drive}/items/${ROOT.id}`; let sofar = '';
    for (const name of parts) {
      sofar = sofar ? sofar + '/' + name : name;
      const got = await gjson(itemPath(sofar) + '?$select=id');
      if (got) { parent = `/drives/${ROOT.drive}/items/${got.id}`; continue; }
      const made = await gjson(parent + '/children', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name, folder: {}, '@microsoft.graph.conflictBehavior': 'replace' }) });
      parent = `/drives/${ROOT.drive}/items/${made.id}`;
    }
  }
  async function putFile(rel, data, type) {
    await resolveRoot();
    const blob = data instanceof Blob ? data : new Blob([data], { type: type || 'application/octet-stream' });
    const attempt = async () => {
      if (blob.size <= 4 * 1024 * 1024) {
        const r = await graph(itemPath(rel) + '/content', { method: 'PUT', headers: { 'Content-Type': type || blob.type || 'application/octet-stream' }, body: blob });
        const j = await r.json().catch(() => ({}));
        if (!r.ok) { const e = new Error((j.error && j.error.message) || 'Upload failed ' + r.status); e.status = r.status; throw e; }
        return j;
      }
      const s = await gjson(itemPath(rel) + '/createUploadSession', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }) });
      if (!s) { const e = new Error('Folder missing'); e.status = 404; throw e; }
      const CH = 320 * 1024 * 10; let off = 0, last = null;
      while (off < blob.size) {
        const end = Math.min(off + CH, blob.size);
        const r = await fetch(s.uploadUrl, { method: 'PUT', headers: { 'Content-Range': `bytes ${off}-${end - 1}/${blob.size}` }, body: blob.slice(off, end) });
        if (!r.ok && r.status !== 202) throw Object.assign(new Error('Upload failed ' + r.status), { status: r.status });
        last = await r.json().catch(() => ({})); off = end;
      }
      return last;
    };
    try { return await attempt(); }
    catch (e) {
      if (e.status !== 404 && e.status !== 409) throw e;
      const dir = rel.split('/').slice(0, -1).join('/'); if (!dir) throw e;
      await ensureFolder(dir); return attempt();
    }
  }
  async function getFile(rel) {
    await resolveRoot();
    const it = await gjson(itemPath(rel) + '?$select=id,eTag,@microsoft.graph.downloadUrl');
    if (!it) return null;
    const url = it['@microsoft.graph.downloadUrl'];
    const r = url ? await fetch(url) : await graph(`/drives/${ROOT.drive}/items/${it.id}/content`);
    if (!r.ok) throw new Error('Download failed ' + r.status);
    return { blob: await r.blob(), eTag: it.eTag };
  }
  async function listFolder(rel) {
    await resolveRoot();
    let url = itemPath(rel) + '/children?$select=id,name,eTag,lastModifiedDateTime,size&$top=999'; const out = [];
    while (url) { const j = await gjson(url); if (!j) return out; out.push(...(j.value || [])); url = j['@odata.nextLink'] || null; }
    return out;
  }
  /* ---- shop drawings: "Shop Drawings" folder (and its subfolders) inside the project folder.
     Files named [MARK]-[MARK]-[REV].pdf, e.g. [ML3-A0-L1-PB014.1]-[ML3-A0-L1-PB014.1]-[C2].pdf; the latest revision wins. ---- */
  const SHOP_DIR = 'Shop Drawings';
  let shopIdx = null, shopAt = 0;
  const shopNorm = (s) => String(s || '').toUpperCase().replace(/\s+/g, '');
  const revKey = (r) => { const m = /^([A-Z]*)(\d*)/.exec(shopNorm(r)) || []; return [m[1] || '', +(m[2] || 0)]; };
  const revNewer = (a, b) => { const x = revKey(a), y = revKey(b); return x[0] !== y[0] ? x[0].length - y[0].length || (x[0] > y[0] ? 1 : -1) : x[1] - y[1]; };
  async function shopIndex(force) {
    if (shopIdx && !force && Date.now() - shopAt < 5 * 60e3) return shopIdx;
    await resolveRoot();
    const out = [];
    const walk = async (rel, depth) => {
      let url = itemPath(rel) + '/children?$select=id,name,folder,file,size,lastModifiedDateTime&$top=999';
      while (url) {
        const r = await graph(url); if (r.status === 404) return; if (!r.ok) throw new Error('Shop drawings folder could not be read (' + r.status + ')');
        const j = await r.json();
        for (const it of j.value || []) {
          if (it.folder) { if (depth < 3) await walk(rel + '/' + it.name, depth + 1); continue; }
          if (!/\.pdf$/i.test(it.name)) continue;
          const m = /^\[([^\]]+)\]-\[([^\]]+)\]-\[([^\]]+)\]\.pdf$/i.exec(it.name);
          out.push({ mark: shopNorm(m ? m[1] : it.name.replace(/\.pdf$/i, '').replace(/[\s_-]*rev(ision)?[\s_-]*[A-Z0-9]+$/i, '')), rev: m ? shopNorm(m[3]) : '', name: it.name, rel: rel + '/' + it.name, size: it.size });
        }
        url = j['@odata.nextLink'] || null;
      }
    };
    await walk(SHOP_DIR, 0);
    shopIdx = out; shopAt = Date.now();
    return out;
  }
  /* ids: element IDs as typed or tapped ("ML3-A0-L1-PB014.1", "PB014.1", "P548.2, PL051") → the latest drawing for each */
  async function shopFind(ids, force) {
    const idx = await shopIndex(force), found = [], missing = [];
    const list = [...new Set(String(Array.isArray(ids) ? ids.join(',') : ids || '').split(/[,;\n]+/).map(shopNorm).filter(Boolean))];
    for (const id of list) {
      let hits = idx.filter((d) => d.mark === id);
      if (!hits.length && !id.includes('-')) hits = idx.filter((d) => d.mark.endsWith('-' + id));   // short form: PB014.1 → ML3-A0-L1-PB014.1
      if (!hits.length) { missing.push(id); continue; }
      const marks = [...new Set(hits.map((d) => d.mark))];
      for (const mk of marks) {
        const best = hits.filter((d) => d.mark === mk).sort((a, b) => revNewer(b.rev, a.rev))[0];
        if (!found.some((f) => f.rel === best.rel)) found.push({ id, ...best });
      }
    }
    return { found, missing, folder: SHOP_DIR, files: idx.length };
  }
  async function shopGet(item) { const f = await getFile(item.rel); if (!f) throw new Error(item.name + ' not found'); return new Uint8Array(await f.blob.arrayBuffer()); }
  async function delFile(rel) { await resolveRoot(); const r = await graph(itemPath(rel), { method: 'DELETE' }); if (!r.ok && r.status !== 404) throw new Error('Delete failed ' + r.status); }

  /* ================= documents: local mirror ================= */
  const docs = new Map(), etags = new Map();
  const subs = new Set();
  const notify = () => setTimeout(() => subs.forEach((f) => { try { f(); } catch (e) { console.error(e); } }), 0);
  const clone = (o) => JSON.parse(JSON.stringify(o));
  let loaded = null;
  const load = () => loaded || (loaded = (async () => {
    const all = await tx('docs', 'readonly', (st) => reqP(st.getAll()));
    all.forEach((r) => { docs.set(r.path, r.data); if (r.etag) etags.set(r.path, r.etag); });
  })());
  const snap = (path) => { const d = docs.get(path); return { id: path.split('/').pop(), exists: d !== undefined, data: () => (d === undefined ? undefined : clone(d)), metadata: { fromCache: !state.synced, hasPendingWrites: false } }; };
  const randId = (n = 20) => Array.from(crypto.getRandomValues(new Uint8Array(n)), (b) => 'abcdefghijklmnopqrstuvwxyz0123456789'[b % 36]).join('');
  const docFile = (path) => `_app/data/${path}.json`;

  async function localPut(path, data, etag) {
    if (data === null) { await tx('docs', 'readwrite', (st) => st.delete(path)); docs.delete(path); etags.delete(path); }
    else { const d = clone(data); await tx('docs', 'readwrite', (st) => st.put({ path, data: d, etag: etag || etags.get(path) || null })); docs.set(path, d); if (etag) etags.set(path, etag); }
  }
  async function writeDoc(path, data) {
    await load();
    const before = docs.get(path);
    await localPut(path, data);
    await enqueue({ op: data === null ? 'del' : 'put', path });
    if (path.startsWith('inspections/')) await enqueue({ op: 'reg', path, prevKey: before ? regKey(before) : null });
    notify(); kick();
  }
  function docRef(path) {
    return {
      id: path.split('/').pop(), path,
      async get() { await load(); return snap(path); },
      set: (d) => writeDoc(path, d),
      async update(d) { await load(); if (!docs.has(path)) throw { code: 'invalid_argument', message: 'Document does not exist' }; return writeDoc(path, { ...docs.get(path), ...d }); },
      delete: () => writeDoc(path, null),
      onSnapshot(next) { const f = () => next(snap(path)); subs.add(f); load().then(f); return () => subs.delete(f); },
      collection: (c) => coll(path + '/' + c),
    };
  }
  function coll(path) {
    const depth = path.split('/').length + 1;
    const list = () => [...docs.keys()].filter((k) => k.startsWith(path + '/') && k.split('/').length === depth).sort().map(snap);
    return {
      path, doc: (id) => docRef(path + '/' + (id || randId())),
      async add(d) { const r = docRef(path + '/' + randId()); await r.set(d); return r; },
      async get() { await load(); const ds = list(); return { docs: ds, size: ds.length, empty: !ds.length, docChanges: () => [], metadata: {} }; },
      onSnapshot(next) { const f = () => { const ds = list(); next({ docs: ds, size: ds.length, empty: !ds.length, docChanges: () => [], metadata: {} }); }; subs.add(f); load().then(f); return () => subs.delete(f); },
    };
  }
  const db = { doc: docRef, collection: coll };

  /* ================= photos & signatures ================= */
  const urls = new Map();
  const hexId = () => Array.from(crypto.getRandomValues(new Uint8Array(16)), (b) => b.toString(16).padStart(2, '0')).join('');
  const extOf = (type) => (/png/.test(type) ? 'png' : /pdf/.test(type) ? 'pdf' : 'jpg');
  const typeOfExt = { jpg: 'image/jpeg', png: 'image/png', pdf: 'application/pdf' };
  const assets = {
    async upload(blob, opts) {
      const type = (opts && opts.type) || blob.type || 'image/jpeg';
      const id = hexId(); const b = blob.type === type ? blob : new Blob([blob], { type });
      await tx('blobs', 'readwrite', (st) => st.put({ id, type, blob: b }));
      urls.set(id, URL.createObjectURL(b));
      await enqueue({ op: 'blob', id, type }); kick();
      return { id, url: urls.get(id), sizeBytes: b.size, contentType: type };
    },
  };
  async function getBlob(id) {
    const r = await tx('blobs', 'readonly', (st) => reqP(st.get(id)));
    if (r) return r.blob;
    // not on this phone yet: fetch from the shared folder
    for (const ext of ['jpg', 'png', 'pdf']) {
      const f = await getFile(`_app/files/${id}.${ext}`).catch(() => null);
      if (f) { const type = typeOfExt[ext]; const b = new Blob([f.blob], { type }); await tx('blobs', 'readwrite', (st) => st.put({ id, type, blob: b })); return b; }
    }
    throw new Error('This file is not on this phone yet — open it once with signal.');
  }
  const pendingImg = new Set();
  function blobURL(id) { if (urls.has(id)) return urls.get(id); hydrateSoon(); return 'data:image/gif;base64,R0lGODlhAQABAAAAACw=#' + id; }
  let hydrateT = null;
  function hydrateSoon() { clearTimeout(hydrateT); hydrateT = setTimeout(hydrate, 30); }
  function hydrate() {
    document.querySelectorAll('img[src^="data:image/gif;base64,R0lGODlhAQABAAAAACw=#"]').forEach(async (img) => {
      const id = img.getAttribute('src').split('#')[1]; if (pendingImg.has(img)) return; pendingImg.add(img);
      try { const b = await getBlob(id); if (!urls.has(id)) urls.set(id, URL.createObjectURL(b)); img.src = urls.get(id); } catch (e) { img.alt = 'Photo not downloaded yet'; }
      pendingImg.delete(img);
    });
  }
  new MutationObserver(hydrateSoon).observe(document.documentElement, { childList: true, subtree: true });
  async function blobBytes(id) { return new Uint8Array(await (await getBlob(id)).arrayBuffer()); }

  /* ================= outbox & sync ================= */
  const state = { synced: false, syncing: false, pending: 0, error: '', last: null, online: navigator.onLine };
  async function enqueue(item) { await tx('outbox', 'readwrite', (st) => st.add({ ...item, at: Date.now() })); state.pending++; paint(); }
  async function outboxAll() { return tx('outbox', 'readonly', (st) => reqP(st.getAll())); }
  const regKey = (r) => [String(r.reportNo || '').trim().toLowerCase(), String(r.assetIds || '').replace(/\s+/g, '').toLowerCase()].join('|');

  /* ---- report numbers: one shared counter file, updated with If-Match so two phones can never take the same number ---- */
  const numOf = (no) => { const m = /(\d+)\s*$/.exec(no || ''); return m ? +m[1] : 0; };
  const COUNTER = '_app/data/counters/report-number.json';
  function notice(msg) { const t = document.getElementById('toast'); if (!t) return; t.textContent = msg; t.hidden = false; clearTimeout(notice.t); notice.t = setTimeout(() => { t.hidden = true; }, 7000); }
  async function writeCounter(body, eTag) {
    const url = itemPath(COUNTER) + '/content' + (eTag ? '' : '?@microsoft.graph.conflictBehavior=fail');
    const go = () => graph(url, { method: 'PUT', headers: { 'Content-Type': 'application/json', ...(eTag ? { 'If-Match': eTag } : {}) }, body: JSON.stringify(body) });
    let r = await go();
    if (r.status === 404) { await ensureFolder('_app/data/counters'); r = await go(); }
    return r.status; // 200/201 ok, 409/412 someone else got there first
  }
  async function reserveNumber(path, d) {
    const p = docs.get('settings/project') || {}; const prefix = p.prefix || 'MEL03-Insp-', digits = p.digits || 4;
    const others = () => Math.max(0, ...[...docs.entries()].filter(([k, v]) => k.startsWith('inspections/') && k !== path && !(v.numberAuto && !v.numberConfirmed)).map(([, v]) => numOf(v.reportNo)));
    for (let attempt = 0; attempt < 8; attempt++) {
      const meta = await gjson(itemPath(COUNTER) + '?$select=id,eTag');
      let last = others();
      if (meta) { const f = await getFile(COUNTER); try { last = Math.max(last, +(JSON.parse(await f.blob.text()).last) || 0); } catch (e) {} }
      const want = Math.max(numOf(d.reportNo), last + 1);
      const st = await writeCounter({ last: want, by: account ? account.username : '', at: new Date().toISOString() }, meta ? meta.eTag : null);
      if (st === 409 || st === 412) { await new Promise((r) => setTimeout(r, 150 + Math.random() * 400)); continue; }
      if (st >= 300) throw Object.assign(new Error('Could not reserve a report number (' + st + ')'), { status: st });
      const newNo = prefix + String(want).padStart(digits, '0');
      const oldNo = d.reportNo;
      const upd = { ...d, reportNo: newNo, numberConfirmed: true };
      await localPut(path, upd);
      if (newNo !== oldNo) {
        // re-point any queued report files at the new number
        for (const q of await outboxAll()) if (q.op === 'file' && q.target && q.target.includes(oldNo)) await tx('outbox', 'readwrite', (o) => o.put({ ...q, target: q.target.split(oldNo).join(newNo) }));
        notify();
        notice(`${oldNo} was just used by another inspector — this inspection is now ${newNo}.`);
      }
      return upd;
    }
    throw new Error('Could not reserve a report number — will retry');
  }

  async function runItem(it) {
    if (it.op === 'put') {
      let d = docs.get(it.path); if (d === undefined) return;
      if (it.path.startsWith('inspections/') && d.numberAuto && !d.numberConfirmed) d = await reserveNumber(it.path, d);
      const res = await putFile(docFile(it.path), new Blob([JSON.stringify(d)], { type: 'application/json' }), 'application/json');
      if (res && res.eTag) { etags.set(it.path, res.eTag); await localPut(it.path, d, res.eTag); }
    } else if (it.op === 'del') {
      await delFile(docFile(it.path));
    } else if (it.op === 'blob') {
      const r = await tx('blobs', 'readonly', (st) => reqP(st.get(it.id))); if (!r) return;
      await putFile(`_app/files/${it.id}.${extOf(it.type)}`, r.blob, it.type);
    } else if (it.op === 'file') {
      const r = await tx('blobs', 'readonly', (st) => reqP(st.get(it.blobId))); if (!r) return;
      await putFile(it.target, r.blob, r.type);
      await tx('blobs', 'readwrite', (st) => st.delete(it.blobId));
    } else if (it.op === 'reg') {
      await syncRegisterRow(it);
    } else if (it.op === 'mail') {
      if (!mailOK()) return 'skip'; // waits until the user allows the app to send email
      await sendMailItem(it);
    }
  }
  let flushing = null, kickT = null;
  function kick() { clearTimeout(kickT); kickT = setTimeout(() => { flush().catch(() => {}); }, 300); }
  async function flush() {
    if (flushing) return flushing;
    flushing = (async () => {
      if (!navigator.onLine || !tokens) return;
      state.syncing = true; state.error = ''; paint();
      try {
        await resolveRoot();
        for (const it of await outboxAll()) {
          try { if ((await runItem(it)) === 'skip') continue; }
          catch (e) {
            if (e.code === 'signin') throw e;
            if (e instanceof TypeError || !navigator.onLine) throw e; // network: stop, retry later
            if (it.op === 'mail' && e.status >= 400 && e.status < 500 && e.status !== 429) { await mailFailed(it, e); await tx('outbox', 'readwrite', (st) => st.delete(it.seq)); continue; }
            if (e.status === 403) { state.error = 'No permission to save in the shared folder — ask IT to give you edit access.'; throw e; }
            if ((it.tries || 0) >= 4) { console.warn('Dropping failed sync item', it, e); state.error = 'One change could not be synced: ' + e.message; }
            else { await tx('outbox', 'readwrite', (st) => st.put({ ...it, tries: (it.tries || 0) + 1 })); throw e; }
          }
          await tx('outbox', 'readwrite', (st) => st.delete(it.seq));
        }
        state.last = new Date();
      } catch (e) {
        if (e.code === 'signin') state.error = 'signin';
        else if (!state.error) state.error = (e instanceof TypeError || !navigator.onLine) ? '' : e.message;
      } finally {
        state.pending = (await outboxAll()).length; state.syncing = false; paint();
      }
    })().finally(() => { flushing = null; });
    return flushing;
  }

  /* pull: refresh the local copy from the shared folder (other inspectors' changes) */
  let pulling = null;
  async function pull() {
    if (pulling) return pulling;
    pulling = (async () => {
      if (!navigator.onLine || !tokens) return;
      try {
        await resolveRoot(); await load();
        const pendingPaths = new Set((await outboxAll()).filter((i) => i.path).map((i) => i.path));
        let changed = false;
        for (const c of COLLS) {
          const items = await listFolder('_app/data/' + c);
          const seen = new Set();
          for (const it of items) {
            if (!/\.json$/.test(it.name)) continue;
            const path = c + '/' + it.name.replace(/\.json$/, ''); seen.add(path);
            if (pendingPaths.has(path) || etags.get(path) === it.eTag) continue;
            const f = await getFile(docFile(path)); if (!f) continue;
            try { await localPut(path, JSON.parse(await f.blob.text()), it.eTag); changed = true; } catch (e) { console.warn('Bad record file', path); }
          }
          for (const path of [...docs.keys()].filter((k) => k.startsWith(c + '/'))) {
            if (!seen.has(path) && !pendingPaths.has(path) && etags.has(path)) { await localPut(path, null); changed = true; }
          }
        }
        state.synced = true; state.last = new Date();
        if (changed) notify();
        reconcileRegister().catch((e) => console.warn('register check', e));
      } catch (e) {
        if (e.code === 'signin') state.error = 'signin';
        else if (!(e instanceof TypeError)) state.error = e.message;
      }
      paint();
    })().finally(() => { pulling = null; });
    return pulling;
  }

  /* ================= live register workbook rows ================= */
  async function syncRegisterRow(it) {
    if (!hasRegister()) return;
    const reg = await resolveRegister(); if (!reg) return;
    const base = `/drives/${reg.drive}/items/${reg.id}/workbook`;
    const rec = docs.get(it.path);
    const tbl = await gjson(`${base}/tables/${encodeURIComponent(CFG.registerTable)}/range?$select=address,values`);
    if (!tbl) throw new Error(`Table ${CFG.registerTable} not found in the register workbook`);
    const m = /!([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(tbl.address); const top = +m[2];
    const rows = tbl.values.slice(1); // skip header
    const keyOf = (v) => [String(v[0] || '').trim().toLowerCase(), String(v[4] || '').replace(/\s+/g, '').toLowerCase()].join('|');
    const find = (k) => rows.findIndex((v) => keyOf(v) === k);
    let idx = -1;
    if (it.prevKey) idx = find(it.prevKey);
    if (idx < 0 && rec) idx = find(regKey(rec));
    if (!rec) { // deleted: clear its row
      if (idx < 0) return;
      return gjson(`${base}/worksheets/${encodeURIComponent(CFG.registerSheet)}/range(address='A${top + 1 + idx}:P${top + 1 + idx}')`, { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ values: [Array(16).fill('')] }) });
    }
    const vals = InspEngine.FIELDS.map((f) => {
      const v = rec[f]; if (v == null || v === '') return '';
      if ((f === 'date' || f === 'weekEnding' || f === 'closeOut') && /^\d{4}-\d{2}-\d{2}$/.test(v)) return InspEngine.excelSerial(v);
      return String(v);
    });
    const myKey = regKey(rec);
    if (idx >= 0) { // our own row already: update in place
      const r = top + 1 + idx;
      return gjson(`${base}/worksheets/${encodeURIComponent(CFG.registerSheet)}/range(address='A${r}:P${r}')`, { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ values: [vals] }) });
    }
    // new row: take the first empty row, then read it back — if another inspector wrote the same row at the same moment, try the next one
    for (let attempt = 0; attempt < 6; attempt++) {
      const t2 = attempt ? await gjson(`${base}/tables/${encodeURIComponent(CFG.registerTable)}/range?$select=address,values`) : tbl;
      const rows2 = t2.values.slice(1);
      if (rows2.some((v) => keyOf(v) === myKey)) return; // already there
      const free = rows2.findIndex((v) => v.every((x) => x === '' || x == null));
      if (free < 0) return gjson(`${base}/tables/${encodeURIComponent(CFG.registerTable)}/rows`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ index: null, values: [vals] }) });
      const r = top + 1 + free, addr = `range(address='A${r}:P${r}')`;
      await gjson(`${base}/worksheets/${encodeURIComponent(CFG.registerSheet)}/${addr}`, { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ values: [vals] }) });
      await new Promise((res) => setTimeout(res, 400 + Math.random() * 400));
      const back = await gjson(`${base}/worksheets/${encodeURIComponent(CFG.registerSheet)}/${addr}?$select=values`);
      if (back && back.values && keyOf(back.values[0]) === myKey) return;
    }
    throw new Error('The register is busy — will retry');
  }

  /* every few minutes: make sure every inspection in the app has its row in the register workbook */
  let lastReconcile = 0;
  async function reconcileRegister() {
    if (!hasRegister() || Date.now() - lastReconcile < 180000) return;
    lastReconcile = Date.now();
    const reg = await resolveRegister(); if (!reg) return;
    const tbl = await gjson(`/drives/${reg.drive}/items/${reg.id}/workbook/tables/${encodeURIComponent(CFG.registerTable)}/range?$select=address,values`);
    if (!tbl) return;
    const have = new Set(tbl.values.slice(1).map((v) => [String(v[0] || '').trim().toLowerCase(), String(v[4] || '').replace(/\s+/g, '').toLowerCase()].join('|')));
    const pend = new Set((await outboxAll()).filter((i) => i.path).map((i) => i.path));
    let added = 0;
    for (const [path, rec] of docs) {
      if (!path.startsWith('inspections/') || pend.has(path) || (rec.numberAuto && !rec.numberConfirmed)) continue;
      if (!have.has(regKey(rec))) { await enqueue({ op: 'reg', path, prevKey: null }); added++; }
    }
    if (added) kick();
  }

  /* ================= reports into SharePoint ================= */
  async function filePackage(r, files) {
    const folder = `${CFG.reportsFolder}/${String(r.reportNo || 'Inspection').replace(/[\\/:*?"<>|]/g, '')}`;
    for (const f of files) {
      const id = 'pkg-' + hexId(); const type = /\.pdf$/i.test(f.path) ? 'application/pdf' : /\.docx$/i.test(f.path) ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' : 'image/jpeg';
      await tx('blobs', 'readwrite', (st) => st.put({ id, type, blob: new Blob([f.bytes], { type }) }));
      await enqueue({ op: 'file', blobId: id, target: `${folder}/${f.path}` });
    }
    kick();
    return navigator.onLine ? `Filing ${r.reportNo} in SharePoint › ${folder}` : `${r.reportNo} will be filed in SharePoint when you're back online`;
  }

  /* ================= email (Microsoft Graph, sent from the signed-in user's own mailbox) ================= */
  async function gmust(path, opts) { // like gjson, but "not found" is an error too
    const r = await graph(path, opts);
    const j = r.status === 202 || r.status === 204 ? {} : await r.json().catch(() => ({}));
    if (!r.ok) { const e = new Error((j.error && j.error.message) || ('Microsoft 365 error ' + r.status)); e.status = r.status; e.code = j.error && j.error.code; throw e; }
    return j;
  }
  const blobB64 = (blob) => new Promise((res, rej) => { const fr = new FileReader(); fr.onload = () => res(String(fr.result).split(',')[1]); fr.onerror = () => rej(fr.error); fr.readAsDataURL(blob); });
  async function mailLogSet(mid, patch) { const log = (await metaGet('mailLog')) || []; const i = log.findIndex((x) => x.mid === mid); if (i < 0) log.push({ mid, ...patch }); else log[i] = { ...log[i], ...patch }; await metaSet('mailLog', log.slice(-200)); window.dispatchEvent(new Event('insp-mail')); }
  async function queueMail(m) { // { rid, to:[], cc:[], subject, html, files:[{ name, type, bytes }] }
    const files = [];
    for (const f of m.files) { const id = 'mail-' + hexId(); await tx('blobs', 'readwrite', (st) => st.put({ id, type: f.type, blob: new Blob([f.bytes], { type: f.type }) })); files.push({ id, name: f.name, type: f.type, size: f.bytes.length }); }
    const mid = hexId().slice(0, 12);
    await enqueue({ op: 'mail', mid, rid: m.rid || '', to: m.to, cc: m.cc || [], subject: m.subject, html: m.html, files });
    await mailLogSet(mid, { rid: m.rid || '', to: m.to, subject: m.subject, status: 'queued', at: new Date().toISOString() });
    if (!mailOK()) return { needConsent: true, mid };
    kick(); return { queued: true, online: navigator.onLine, mid };
  }
  async function sendMailItem(it) {
    const json = { 'Content-Type': 'application/json' };
    if (!it.msgId) { // 1. draft in the user's mailbox
      const draft = await gmust('/me/messages', { method: 'POST', headers: json, body: JSON.stringify({ subject: it.subject, body: { contentType: 'HTML', content: it.html },
        toRecipients: it.to.map((a) => ({ emailAddress: { address: a } })), ccRecipients: (it.cc || []).map((a) => ({ emailAddress: { address: a } })) }) });
      it = { ...it, msgId: draft.id, done: [] }; await tx('outbox', 'readwrite', (st) => st.put(it));
    }
    for (const f of it.files) { // 2. attachments: small ones in one request, large ones through an upload session
      if ((it.done || []).includes(f.id)) continue;
      const r = await tx('blobs', 'readonly', (st) => reqP(st.get(f.id))); if (!r) continue;
      const base = `/me/messages/${encodeURIComponent(it.msgId)}/attachments`;
      if (r.blob.size < (CFG.mailInlineMax || 3 * 1024 * 1024)) {
        await gmust(base, { method: 'POST', headers: json, body: JSON.stringify({ '@odata.type': '#microsoft.graph.fileAttachment', name: f.name, contentType: f.type, contentBytes: await blobB64(r.blob) }) });
      } else {
        const ses = await gmust(base + '/createUploadSession', { method: 'POST', headers: json, body: JSON.stringify({ AttachmentItem: { attachmentType: 'file', name: f.name, size: r.blob.size, contentType: f.type } }) });
        const CH = 3 * 1024 * 1024, size = r.blob.size;
        for (let a = 0; a < size; a += CH) {
          const part = r.blob.slice(a, Math.min(size, a + CH));
          const up = await fetch(ses.uploadUrl, { method: 'PUT', headers: { 'Content-Type': 'application/octet-stream', 'Content-Range': `bytes ${a}-${a + part.size - 1}/${size}` }, body: part });
          if (!up.ok) { const e = new Error(`Attachment ${f.name} did not upload (${up.status})`); e.status = up.status; throw e; }
        }
      }
      it = { ...it, done: [...(it.done || []), f.id] }; await tx('outbox', 'readwrite', (st) => st.put(it));
    }
    await gmust(`/me/messages/${encodeURIComponent(it.msgId)}/send`, { method: 'POST' }); // 3. send (a copy stays in Sent Items)
    for (const f of it.files) await tx('blobs', 'readwrite', (st) => st.delete(f.id));
    const sentAt = new Date().toISOString();
    await mailLogSet(it.mid, { status: 'sent', sentAt });
    const path = 'inspections/' + it.rid, d = it.rid && docs.get(path);
    if (d) await docRef(path).set({ ...d, emails: [...(d.emails || []), { to: it.to, cc: it.cc || [], subject: it.subject, files: it.files.map((f) => f.name), sentAt, by: account ? account.name : '' }] });
    notice(`Emailed ${it.subject} to ${it.to.join(', ')}`);
  }
  async function mailFailed(it, e) {
    const why = /MailboxNotEnabledForRESTAPI|MailboxNotFound/i.test(e.code || e.message) ? 'this Microsoft account has no Outlook mailbox the app can send from'
      : e.status === 403 ? 'your organisation does not allow the app to send email' : (e.message || 'error ' + e.status);
    for (const f of it.files) await tx('blobs', 'readwrite', (st) => st.delete(f.id));
    if (it.msgId) gjson(`/me/messages/${encodeURIComponent(it.msgId)}`, { method: 'DELETE' }).catch(() => {});
    await mailLogSet(it.mid, { status: 'failed', error: why, failedAt: new Date().toISOString() });
    notice(`Email "${it.subject}" was not sent: ${why}. Use your phone's mail app instead.`);
  }

  /* any other file the app keeps up to date in the folder (e.g. Trackers/<drawing> - Inspection Tracker.pdf) */
  async function fileAt(target, bytes, type) {
    const id = 'pkg-' + hexId();
    await tx('blobs', 'readwrite', (st) => st.put({ id, type, blob: new Blob([bytes], { type }) }));
    await enqueue({ op: 'file', blobId: id, target }); kick();
  }

  /* ================= file saving (share sheet / download) ================= */
  function fileSheet(file) {
    return new Promise((resolve, reject) => {
      const wrap = document.createElement('div'); wrap.className = 'sheet';
      const canShare = !!(navigator.canShare && navigator.share && (() => { try { return navigator.canShare({ files: [file] }); } catch (e) { return false; } })());
      const url = URL.createObjectURL(file);
      wrap.innerHTML = `<div class="sheet-card" role="dialog" aria-modal="true" aria-label="File ready"><b>File ready</b><div class="sheet-name"></div>
        <div class="sheet-btns">${canShare ? '<button class="btn primary" data-a="share">Save or share…</button>' : ''}<a class="btn ${canShare ? '' : 'primary'}" data-a="dl">Download</a><button class="btn ghost" data-a="close">Close</button></div>
        <p class="hint">${canShare ? 'Save or share lets you save to Files, email it, or send it to Teams/OneDrive.' : 'The file goes to your Downloads folder.'}</p></div>`;
      wrap.querySelector('.sheet-name').textContent = file.name + ' · ' + (file.size > 1048576 ? (file.size / 1048576).toFixed(1) + ' MB' : Math.ceil(file.size / 1024) + ' KB');
      const a = wrap.querySelector('[data-a="dl"]'); a.href = url; a.download = file.name;
      const done = (ok) => { wrap.remove(); setTimeout(() => URL.revokeObjectURL(url), 60000); ok ? resolve({ status: 'saved' }) : reject({ code: 'declined' }); };
      a.addEventListener('click', () => setTimeout(() => done(true), 300));
      wrap.querySelector('[data-a="close"]').onclick = () => done(false);
      const sh = wrap.querySelector('[data-a="share"]');
      if (sh) sh.onclick = async () => { try { await navigator.share({ files: [file], title: file.name }); done(true); } catch (e) { if (e && e.name === 'AbortError') return; a.click(); } };
      document.body.appendChild(wrap); (sh || a).focus();
    });
  }
  const MIME = { docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', zip: 'application/zip', pdf: 'application/pdf' };
  const downloads = { save({ filename, data }) { const ext = (filename.split('.').pop() || '').toLowerCase(); return fileSheet(new File([data], filename, { type: MIME[ext] || 'application/octet-stream' })); } };

  /* ================= status UI ================= */
  function syncLabel() {
    if (!navigator.onLine) return state.pending ? `Offline · ${state.pending} to sync` : 'Offline';
    if (state.error === 'signin') return 'Sign in needed';
    if (state.error) return 'Sync problem';
    if (state.syncing) return 'Syncing…';
    if (state.pending) return `${state.pending} to sync`;
    return state.synced ? 'Synced' : 'Connecting…';
  }
  function paint() {
    const t = $('#synctext'), d = $('#syncdot'); if (!t) return;
    t.textContent = syncLabel();
    d.className = 'dot' + (navigator.onLine && !state.error && !state.pending && state.synced ? ' live' : '');
    const el = $('#acctstatus'); if (el) el.innerHTML = statusHtml();
    const bar = $('#signinbar');
    if (state.error === 'signin' && navigator.onLine) {
      if (!bar) { const b = document.createElement('div'); b.id = 'signinbar'; b.className = 'banner signinbar'; b.innerHTML = 'Your sign-in has expired. Your work is safe on this phone. <button class="btn small primary">Sign in again</button>'; b.querySelector('button').onclick = signIn; document.body.appendChild(b); }
    } else if (bar) bar.remove();
  }
  function statusHtml() {
    const parts = [];
    parts.push(navigator.onLine ? (state.error && state.error !== 'signin' ? `<span style="color:var(--fail)">${escH(state.error)}</span>` : state.pending ? `${state.pending} change${state.pending > 1 ? 's' : ''} waiting to upload` : 'Everything is uploaded') : `Offline — ${state.pending} change${state.pending === 1 ? '' : 's'} will upload when signal returns`);
    if (state.last) parts.push('Last sync ' + state.last.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }));
    return parts.join(' · ');
  }
  function settingsHtml() {
    return `<h2 style="margin-top:12px">Account &amp; sync</h2>
    <section class="sec">
      <div class="row"><div class="grow"><b>${escH(account ? account.name : '')}</b><div class="hint">${escH(account ? account.username : '')}</div></div><button class="btn small ghost" id="signout">Sign out</button></div>
      <div class="hint" id="acctstatus">${statusHtml()}</div>
      <div class="row"><div class="grow"><b>Project: ${escH((active && active.name) || '—')}</b><div class="hint">Data folder: ${escH(ROOT ? ROOT.name : 'shared folder')} · Reports filed in “${escH(CFG.reportsFolder)}” · Register workbook: ${hasRegister() ? 'updated live' : 'not connected'}</div></div><button class="btn small" id="projbtn">Projects…</button></div>
      ${window.__licence ? `<div class="row"><div class="grow hint">Company: <b>${escH(window.__licence.company)}</b>${window.__licence.exp ? ' · licence until ' + escH(window.__licence.exp) : ''}</div><button class="btn small ghost" id="licout">Sign out of company</button></div>` : ''}
      <div class="row"><button class="btn" id="syncnow">Sync now</button></div>
      <div id="signoutconfirm"></div>
    </section>`;
  }
  function bindSettings() {
    $('#syncnow').onclick = async () => { await flush(); await pull(); paint(); };
    $('#projbtn').onclick = () => showProjects(false);
    const lo = $('#licout'); if (lo) lo.onclick = () => { if (confirm('Sign this device out of the company account? You will need the company username and password to open the app again.')) window.__licence.signOut(); };
    $('#signout').onclick = async () => {
      const n = (await outboxAll()).length;
      $('#signoutconfirm').innerHTML = `<div class="confirm">${n ? `${n} change${n > 1 ? 's have' : ' has'} not uploaded yet and will be lost. ` : ''}Sign out and remove the data from this phone? <button class="btn small" id="soyes">Sign out</button><button class="btn small ghost" id="sono">Cancel</button></div>`;
      $('#sono').onclick = () => { $('#signoutconfirm').innerHTML = ''; };
      $('#soyes').onclick = signOut;
    };
  }

  /* ================= projects: list, create, open shared, invite ================= */
  const jsonH = { 'Content-Type': 'application/json' };
  async function gmustP(path, opts) { const r = await graph(path, opts); const j = r.status === 202 || r.status === 204 ? {} : await r.json().catch(() => ({})); if (!r.ok) { const e = new Error((j.error && j.error.message) || ('Microsoft 365 error ' + r.status)); e.status = r.status; e.code = j.error && j.error.code; throw e; } return j; }
  const encSeg = (n) => encodeURIComponent(n);
  async function childFolder(drive, parentId, name) {
    const got = await gjson(`/drives/${drive}/items/${parentId}:/${encSeg(name)}:?$select=id,folder`);
    if (got) return got.id;
    return (await gmustP(`/drives/${drive}/items/${parentId}/children`, { method: 'POST', headers: jsonH, body: JSON.stringify({ name, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }) })).id;
  }
  async function putUnder(drive, rootId, rel, data, type) {
    const parts = rel.split('/'), file = parts.pop(); let id = rootId;
    for (const n of parts) id = await childFolder(drive, id, n);
    return gmustP(`/drives/${drive}/items/${id}:/${encSeg(file)}:/content`, { method: 'PUT', headers: { 'Content-Type': type }, body: data instanceof Blob ? data : new Blob([data], { type }) });
  }
  async function myRoot() { const r = await gmustP('/me/drive/root?$select=id,parentReference'); return { drive: r.parentReference.driveId, id: r.id }; }
  const LIST_DIR = 'Site Inspections';
  async function syncProjectList() { // keep the list in the user's own OneDrive so it follows them to other devices
    try {
      const me = await myRoot(), dir = await childFolder(me.drive, me.id, LIST_DIR);
      const meta = await gjson(`/drives/${me.drive}/items/${dir}:/projects.json:?$select=id,@microsoft.graph.downloadUrl`);
      let remote = [];
      if (meta) { const r = await fetch(meta['@microsoft.graph.downloadUrl']); if (r.ok) remote = (await r.json().catch(() => ({}))).projects || []; }
      const same = (a, b) => (a.drive && a.drive === b.drive && a.item === b.item) || (a.link && a.link === b.link);
      let changed = false;
      for (const p of remote) if (!projects.some((x) => same(x, p))) { projects.push(p); changed = true; }
      const mine = projects.filter((p) => !p.fromConfig);
      if (changed) saveProjects();
      if (mine.length !== remote.length || mine.some((p) => !remote.some((x) => same(x, p) && x.name === p.name))) await putUnder(me.drive, dir, 'projects.json', JSON.stringify({ projects: mine }), 'application/json');
    } catch (e) { console.warn('project list not synced', e); }
  }
  function switchTo(p) { active = p; saveProjects(); location.replace(location.pathname); }
  async function findRegister(drive, id) {
    const j = await gjson(`/drives/${drive}/items/${id}/children?$select=name,file&$top=200`);
    const x = ((j && j.value) || []).find((c) => c.file && /\.xlsx$/i.test(c.name) && /register/i.test(c.name));
    return x ? x.name : '';
  }
  async function createProject(f) {
    const seed = (window.__projectSeed && window.__projectSeed()) || {};
    const me = await myRoot(), base = await childFolder(me.drive, me.id, LIST_DIR);
    const folder = await childFolder(me.drive, base, f.name.replace(/[\\/:*?"<>|#%]/g, '-').trim());
    const src = f.copy && seed.settings ? JSON.parse(JSON.stringify(seed.settings)) : {};
    const settings = { job: f.job || f.name, jobShort: f.short || f.name, prefix: f.prefix, digits: src.digits || 4,
      workAreas: src.workAreas, inspTypes: src.inspTypes, failureCats: src.failureCats, ncrStatuses: src.ncrStatuses, responsible: src.responsible,
      recipients: f.copy ? src.recipients || [] : [], inspectors: f.copy ? src.inspectors || [] : [], emailContacts: f.copy ? src.emailContacts || [] : [] };
    Object.keys(settings).forEach((k) => settings[k] === undefined && delete settings[k]);
    // signature images belong to the old folder: copy them so reports keep working
    for (const i of settings.inspectors || []) if (i.sigAsset) { try { const b = await getBlob(i.sigAsset); await putUnder(me.drive, folder, `_app/files/${i.sigAsset}.${extOf(b.type || 'image/png')}`, b, b.type || 'image/png'); } catch (e) { delete i.sigAsset; } }
    await putUnder(me.drive, folder, '_app/data/settings/project.json', JSON.stringify(settings), 'application/json');
    if (f.copy) for (const [id, t] of Object.entries(seed.types || {})) { const { id: _x, trackerId, ...doc } = t; await putUnder(me.drive, folder, `_app/data/types/${id}.json`, JSON.stringify(doc), 'application/json'); }
    let register = '';
    if (seed.registerBlank) {
      register = `${(f.prefix || 'Project').replace(/-?Insp-?$/i, '') || 'Project'}-Inspection-Register.xlsx`;
      const it = await putUnder(me.drive, folder, register, new Blob([seed.registerBlank], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      for (const [sh, cell] of [['Inspection Register', 'K1'], ['Lists', 'K1'], ['Dashboard', 'K2']]) // job number in the headings (best effort)
        try { await gmustP(`/drives/${me.drive}/items/${it.id}/workbook/worksheets/${encodeURIComponent(sh)}/range(address='${cell}')`, { method: 'PATCH', headers: jsonH, body: JSON.stringify({ values: [['Job No. ' + (f.job || f.name)]] }) }); } catch (e) {}
    }
    const p = { id: 'p' + Date.now().toString(36), name: f.short || f.name, drive: me.drive, item: folder, register, owner: account ? account.username : '' };
    projects.push(p); saveProjects(); await syncProjectList();
    return p;
  }
  const joinLink = (p) => location.origin + location.pathname + '?join=' + b64uJ({ drive: p.drive, item: p.item, name: p.name, register: p.register });
  async function shareProject(email) {
    await resolveRoot();
    const p = active, link = joinLink({ ...p, drive: ROOT.drive, item: ROOT.id, register: CFG.registerInFolder || (REG && REG.name) || '' });
    await gmustP(`/drives/${ROOT.drive}/items/${ROOT.id}/invite`, { method: 'POST', headers: jsonH, body: JSON.stringify({ requireSignIn: true, sendInvitation: true, roles: ['write'], recipients: [{ email }],
      message: `${account ? account.name : 'A colleague'} has added you to the project "${p.name}" in Site Inspections. Open the app with this link to add the project: ${link}` }) });
    return link;
  }
  async function sharedFolders() {
    const j = await gmustP('/me/drive/sharedWithMe');
    return (j.value || []).map((v) => v.remoteItem || v).filter((r) => r && r.folder && r.parentReference).map((r) => ({ name: r.name, drive: r.parentReference.driveId, item: r.id }));
  }
  async function addProject(x) {
    let p = projects.find((q) => q.drive === x.drive && q.item === x.item);
    if (!p) { p = { id: 'p' + Date.now().toString(36), name: x.name, drive: x.drive, item: x.item, register: x.register || await findRegister(x.drive, x.item).catch(() => '') }; projects.push(p); saveProjects(); syncProjectList(); }
    return p;
  }
  async function showProjects(first) {
    let m = $('#projgate'); if (!m) { m = document.createElement('div'); m.id = 'projgate'; document.body.appendChild(m); }
    const n = (await outboxAll()).length, cur = (docs.get('settings/project') || {});
    if (active && !active.name && cur.jobShort) { active.name = cur.jobShort; saveProjects(); }
    const sugg = (cur.prefix || '').replace(/\d+(-?Insp-?)$/i, '');
    m.innerHTML = `<div class="gate-card proj-card" role="dialog" aria-modal="true" aria-labelledby="projtitle">
      <div class="row" style="width:100%"><h1 id="projtitle" class="grow">Projects</h1>${first ? '' : '<button class="btn small ghost" id="projclose">Close</button>'}</div>
      ${first ? '<p>Start a new project, or open one a colleague has shared with you.</p>' : ''}
      ${projects.length ? `<div class="projlist">${projects.map((p, i) => `<div class="row projrow"><div class="grow"><b>${escH(p.name || (p.fromConfig ? 'Main project' : 'Project'))}</b>${p === active ? ' <span class="pill closed">open</span>' : ''}<div class="hint">${escH(p.fromConfig ? 'set up for this app' : p.owner ? 'created by ' + p.owner : 'shared with you')}</div></div>
        ${p === active ? '' : `<button class="btn small primary" data-popen="${i}">Open</button>${p.fromConfig ? '' : `<button class="btn small ghost" data-prm="${i}" aria-label="Remove from list">✕</button>`}`}</div>`).join('')}</div>` : ''}
      ${n && active ? `<p class="hint">${n} change${n > 1 ? 's' : ''} in ${escH(active.name || 'this project')} still waiting to upload — they stay on this phone and upload next time you open it.</p>` : ''}
      <details ${first ? 'open' : ''} class="projsec"><summary><b>Start a new project</b></summary>
        <label class="f"><span>Job number and name (shown on reports)</span><input class="inp" id="np-job" placeholder="e.g. 20140-MEL04 Data Centre"></label>
        <label class="f"><span>Short name (header and project list)</span><input class="inp" id="np-short" placeholder="e.g. 20140-MEL04"></label>
        <label class="f"><span>Report number prefix</span><input class="inp mono" id="np-prefix" placeholder="e.g. MEL04-Insp-" value=""></label>
        ${active && cur.prefix ? `<label style="display:flex;gap:10px;align-items:flex-start"><input type="checkbox" id="np-copy" checked style="margin-top:4px;flex:none"> Copy templates, inspectors, recipients and email contacts from ${escH(active.name || 'this project')}</label>` : ''}
        <p class="hint" style="margin:0">Creates the project folder in your OneDrive › ${LIST_DIR}, with an empty register. Report numbers start at 0001.</p>
        <button class="btn primary" id="np-go">Create project</button></details>
      <details class="projsec"><summary><b>Open a project shared with you</b></summary>
        <button class="btn small" id="sp-list">Show folders shared with me</button><div id="sp-out"></div>
        <label class="f"><span>…or paste the folder's sharing link</span><input class="inp" id="sp-link" placeholder="https://…sharepoint.com/:f:/…"></label><button class="btn small" id="sp-add">Add</button></details>
      ${active && !first ? `<details class="projsec"><summary><b>Share ${escH(active.name || 'this project')} with an engineer</b></summary>
        <label class="f"><span>Their work email (the Microsoft account they sign in with)</span><input class="inp" id="sh-email" type="email" placeholder="name@company.com"></label>
        <div class="row"><button class="btn small primary" id="sh-go">Share and send invitation</button><button class="btn small" id="sh-copy">Copy invite link</button></div>
        <p class="hint" style="margin:0">They get edit access to the project folder and an email with a link that adds the project to their app.</p></details>` : ''}
      <p class="hint" id="projmsg" aria-live="polite"></p>
    </div>`;
    const msg = (t, bad) => { const e = $('#projmsg'); e.textContent = t; e.className = 'hint' + (bad ? ' gate-err' : ''); };
    const busy = async (btn, fn) => { const l = btn.textContent; btn.disabled = true; btn.textContent = 'Working…'; try { await fn(); } catch (e) { console.error(e); msg(e.message || String(e), true); } finally { if (document.body.contains(btn)) { btn.disabled = false; btn.textContent = l; } } };
    const c = $('#projclose'); if (c) c.onclick = () => m.remove();
    m.querySelectorAll('[data-popen]').forEach((b) => { b.onclick = () => busy(b, async () => { await flush().catch(() => {}); switchTo(projects[+b.dataset.popen]); }); });
    m.querySelectorAll('[data-prm]').forEach((b) => { b.onclick = () => { projects.splice(+b.dataset.prm, 1); saveProjects(); syncProjectList(); showProjects(first); }; });
    $('#np-job').oninput = () => { const v = $('#np-job').value.trim(); const sh = $('#np-short'); if (!sh.dataset.t) sh.value = v.split(/\s+/)[0] || ''; const pf = $('#np-prefix'); if (!pf.dataset.t) { const code = (v.match(/[A-Z]{2,}\d+[A-Z]?/i) || [''])[0].toUpperCase(); pf.value = code ? code + '-Insp-' : ''; } };
    $('#np-short').oninput = (e) => { e.target.dataset.t = 1; }; $('#np-prefix').oninput = (e) => { e.target.dataset.t = 1; };
    $('#np-go').onclick = (e) => busy(e.target, async () => {
      const job = $('#np-job').value.trim(), short = $('#np-short').value.trim() || job, prefix = $('#np-prefix').value.trim();
      if (!job || !prefix) { msg('Enter the job number and name, and the report number prefix.', true); return; }
      if (projects.some((p) => (p.name || '').toLowerCase() === short.toLowerCase())) { msg('There is already a project called ' + short, true); return; }
      msg('Creating the project folder…');
      const p = await createProject({ name: short, job, short, prefix, copy: !!($('#np-copy') && $('#np-copy').checked) });
      msg('Project created — opening it…'); switchTo(p);
    });
    $('#sp-list').onclick = (e) => busy(e.target, async () => {
      const list = await sharedFolders();
      $('#sp-out').innerHTML = list.length ? list.map((x, i) => `<div class="row projrow"><span class="grow">${escH(x.name)}</span><button class="btn small" data-spadd="${i}">Add and open</button></div>`).join('') : '<p class="hint">Nothing shared with you yet. Ask the project owner to share it (Projects › Share).</p>';
      m.querySelectorAll('[data-spadd]').forEach((b) => { b.onclick = () => busy(b, async () => switchTo(await addProject(list[+b.dataset.spadd]))); });
    });
    $('#sp-add').onclick = (e) => busy(e.target, async () => {
      const link = $('#sp-link').value.trim(); if (!link) return;
      const it = await gjson(`/shares/${shareId(link)}/driveItem?$select=id,name,parentReference,folder`);
      if (!it || !it.folder) throw new Error('That link does not open a folder you have access to.');
      switchTo(await addProject({ name: it.name, drive: it.parentReference.driveId, item: it.id }));
    });
    const sg = $('#sh-go'); if (sg) sg.onclick = (e) => busy(e.target, async () => {
      const email = $('#sh-email').value.trim(); if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) { msg('Enter a valid email address.', true); return; }
      await shareProject(email); msg(`Shared with ${email}. They'll get an email with the link.`);
    });
    const sc = $('#sh-copy'); if (sc) sc.onclick = (e) => busy(e.target, async () => {
      await resolveRoot(); const link = joinLink({ ...active, drive: ROOT.drive, item: ROOT.id, register: CFG.registerInFolder || '' });
      try { await navigator.clipboard.writeText(link); msg('Invite link copied. They also need access to the folder — use "Share and send invitation" or share the folder in OneDrive.'); } catch (x) { msg(link); }
    });
  }

  /* ================= sign-in screen ================= */
  let gateError = '';
  function showGate(kind) {
    let g = $('#gate');
    if (!g) { g = document.createElement('div'); g.id = 'gate'; document.body.appendChild(g); }
    const logo = (document.querySelector('.logo img') || {}).src || '';
    g.innerHTML = `<div class="gate-card">
      ${logo ? `<img src="${logo}" alt="Rincovitch" class="gate-logo">` : ''}
      <h1>Site Inspections</h1>
      ${window.__licence ? `<p style="margin-top:-10px">for <b style="color:var(--charcoal)">${escH(window.__licence.company)}</b></p>` : ''}
      ${kind === 'setup' ? '<p>This app has not been connected to Microsoft 365 yet. Fill in <b>config.js</b> (see the setup guide).</p>'
        : kind === 'offline' ? '<p>Connect to the internet once to sign in. After that the app also works offline on site.</p>'
        : kind === 'error' ? '<p>The app could not start.</p><button class="btn primary" id="gateretry">Try again</button>'
        : `<p>${window.__licence ? `Company: <b>${escH(window.__licence.company)}</b> ✓<br>Now sign in with your own Microsoft account.` : 'Sign in with your Microsoft account.'}</p><button class="btn primary" id="gatebtn">Sign in with Microsoft</button>`}
      ${gateError ? `<p class="gate-err">${escH(gateError)}</p>` : ''}
      ${memMode ? `<p class="gate-err">This browser is not letting the app store data on this device (${escH(storageErr)}). It will work online, but not offline. Allow site data / cookies for this address, or use a normal (not private) window.</p>` : ''}
      <button class="btn ghost small" id="gatereset">Reset the app on this device</button>
      ${window.QORILO ? `<div class="gate-q"><span>Powered by</span><picture><source srcset="${window.QORILO.logoDark}" media="(prefers-color-scheme: dark)"><img src="${window.QORILO.logo}" alt="Qorilo"></picture></div>` : ''}
      <p class="gate-foot" style="${window.QORILO ? 'border:0;padding:0' : ''}">${window.QORILO ? 'Qorilo ' : ''}Rev ${escH((window.APP_INFO || {}).rev || '1.2')} · Developed by <b>${escH((window.APP_INFO || {}).credit || 'Ali YZ')}</b> · © ${escH((window.APP_INFO || {}).year || '2026')}${window.__licence ? ' · <a href="#" id="gatecompany" style="color:inherit">Change company</a>' : ''}</p>
    </div>`;
    const b = $('#gatebtn'); if (b) b.onclick = signIn;
    const rt = $('#gateretry'); if (rt) rt.onclick = () => location.reload();
    $('#gatereset').onclick = resetDevice;
    const gc = $('#gatecompany'); if (gc) gc.onclick = (e) => { e.preventDefault(); window.__licence.signOut(); };
  }
  async function resetDevice() {
    try { lsSet(TOK, null); lsSet('m365.account', null); lsSet('m365.pkce', null); } catch (e) {}
    try { if (window.caches) for (const k of await caches.keys()) await caches.delete(k); } catch (e) {}
    try { if (navigator.serviceWorker) for (const r of await navigator.serviceWorker.getRegistrations()) await r.unregister(); } catch (e) {}
    try { if (idb) { const d = await idb; if (d) d.close(); } indexedDB.deleteDatabase(DB_NAME); } catch (e) {}
    location.replace(location.pathname);
  }
  const hideGate = () => { const g = $('#gate'); if (g) g.remove(); };

  /* ================= start ================= */
  function matchInspector() {
    const p = docs.get('settings/project'); if (!p || !account) return;
    const me = (p.inspectors || []).find((i) => i.email && i.email.toLowerCase() === account.username);
    if (me) lsSet('insp.me', me.initials);
  }
  const ready = (async () => {
    await handleRedirect();
    await load();
    if (memMode) console.warn('Storage fallback active:', storageErr);
    if (!tokens) {
      if (!CFG.clientId) showGate('setup'); else if (!navigator.onLine && !docs.size) showGate('offline');
      else if (docs.size && account) { /* has data on the phone: let them work, ask to sign in via banner */ state.error = 'signin'; }
      else showGate('signin');
    }
    if (tokens) { hideGate(); if (!active) { syncProjectList().then(() => { if (!active && projects.length) switchTo(projects[0]); }); showProjects(true); return; } const denied = await resolveRoot().then(() => null, (e) => (e.code === 'noaccess' ? e : null));
      if (denied) { await showProjects(false); const pm = $('#projmsg'); if (pm) { pm.textContent = denied.message; pm.className = 'hint gate-err'; } syncProjectList(); return; }
      await pull(); matchInspector(); kick(); syncProjectList(); const sp = docs.get('settings/project'); if (active && sp && sp.jobShort && active.name !== sp.jobShort) { active.name = sp.jobShort; projects = projects.map((x) => (x.id === active.id ? active : x)); saveProjects(); } }
  })().catch((e) => { console.error(e); gateError = (e && e.message) || String(e); showGate('error'); throw e; });

  window.addEventListener('online', () => { paint(); flush().then(pull); });
  window.addEventListener('offline', paint);
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') { flush().then(pull); } });
  setInterval(() => { if (document.visibilityState === 'visible') flush().then(pull); }, CFG.pollSeconds * 1000);
  outboxAll().then((a) => { state.pending = a.length; paint(); });
  setTimeout(paint, 500);

  window.__standalone = { shopFind, shopGet, blobURL, blobBytes, syncLabel, settingsHtml, bindSettings, filePackage, fileAt, openProjects: () => showProjects(false), projectName: () => (active && active.name) || '', queueMail, mailOK, allowMail: () => signIn({ mail: true }), mailLog: async () => (await metaGet('mailLog')) || [], userEmail: () => (account ? account.username : ''), userName: () => (account ? account.name : ''), _state: state, _flush: flush, _pull: pull, _put: putFile, _get: getFile, _reconcile: () => { lastReconcile = 0; return reconcileRegister(); } };
  const caps = { db, assets, downloads };
  window.claude = { use: (name) => (name === 'db' ? ready.then(() => (docs.size || tokens ? db : null), () => null) : Promise.resolve(caps[name] || null)) };

  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
})();
