/**
 * CF-M365-Admin optimized single-file worker
 * Features:
 * - strict Turnstile validation
 * - centered register button
 * - admin toast notifications
 * - grouped invite codes
 * - admin-side create user
 */

const KV = {
  CONFIG: 'config',
  INSTALL_LOCK: 'install_lock',
  SESS_PREFIX: 'sess:',
  INVITES: 'invites',
  COMPAT_CARDS: 'cards',
};

// 订阅（SKU）友好名称映射：skuPartNumber → 中文可读名称
const SUBSCRIPTION_FRIENDLY_NAMES = {
  'SPE_E3': 'Microsoft 365 E3',
  'SPE_E5': 'Microsoft 365 E5',
  'SPE_F1': 'Microsoft 365 F1',
  'SPE_F3': 'Microsoft 365 F3',
  'ENTERPRISEPREMIUM': 'Office 365 E3',
  'ENTERPRISEPACK': 'Office 365 E3',
  'STANDARDPACK': 'Office 365 E1',
  'DESKLESSPACK': 'Office 365 F3',
  'O365_BUSINESS_ESSENTIALS': 'Microsoft 365 商业基础版',
  'O365_BUSINESS_PREMIUM': 'Microsoft 365 商业标准版',
  'O365_BUSINESS': 'Microsoft 365 商业应用版',
  'DEVELOPERPACK': 'Office 365 E3 开发者版',
  'DEVELOPERPACK_E5': 'Microsoft 365 E5 开发者版',
  'STANDARDWOFFPACK_FACULTY': 'Office 365 A1 教职工版',
  'STANDARDWOFFPACK_STUDENT': 'Office 365 A1 学生版',
  'M365EDU_A1': 'Microsoft 365 A1',
  'M365EDU_A3_FACULTY': 'Microsoft 365 A3 教职工版',
  'M365EDU_A3_STUDENT': 'Microsoft 365 A3 学生版',
  'M365EDU_A5_FACULTY': 'Microsoft 365 A5 教职工版',
  'M365EDU_A5_STUDENT': 'Microsoft 365 A5 学生版',
  'EXCHANGESTANDARD': 'Exchange Online 计划 1',
  'EXCHANGEENTERPRISE': 'Exchange Online 计划 2',
  'SHAREPOINTSTANDARD': 'SharePoint Online 计划 1',
  'SHAREPOINTENTERPRISE': 'SharePoint Online 计划 2',
  'TEAMS_EXPLORATORY': 'Microsoft Teams 探索版',
  'TEAMS_FREE': 'Microsoft Teams 免费版',
  'POWER_BI_STANDARD': 'Power BI 免费版',
  'POWER_BI_PRO': 'Power BI Pro',
  'AAD_PREMIUM': 'Microsoft Entra ID P1',
  'AAD_PREMIUM_P2': 'Microsoft Entra ID P2',
  'INTUNE_A': 'Microsoft Intune',
  'EMS': 'Enterprise Mobility + Security E3',
  'EMSPREMIUM': 'Enterprise Mobility + Security E5',
  'WIN10_PRO_ENT_SUB': 'Windows 10/11 企业版 E3',
};

// Service Plan 友好名称映射：planName → 中文可读名称
const PLAN_FRIENDLY_NAMES = {
  'EXCHANGE_S_ENTERPRISE': 'Outlook 邮件',
  'EXCHANGE_S_STANDARD': 'Outlook 邮件',
  'EXCHANGE_L_STANDARD': 'Outlook 邮件',
  'EXCHANGE_S_FOUNDATION': 'Outlook 邮件',
  'EXCHANGE_S_DESKLESS': 'Outlook 邮件',
  'SHAREPOINTWAC': 'OneDrive',
  'SHAREPOINTWAC_DEVELOPER': 'OneDrive',
  'SHAREPOINTENTERPRISE': 'OneDrive',
  'SHAREPOINTSTANDARD': 'OneDrive',
  'SHAREPOINT_PROJECT_EDU': 'OneDrive',
  'SHAREPOINT_L_STANDARD': 'OneDrive',
  'TEAMS1': 'Microsoft Teams',
  'TEAMS_FREE': 'Microsoft Teams 免费版',
  'TEAMS_EXPLORATORY': 'Microsoft Teams 探索版',
  'TEAMS_COMM': 'Microsoft Teams',
  'TEAMSPRO': 'Microsoft Teams Pro',
  'TEAMS_PRO_EDU': 'Microsoft Teams Pro',
  'TEAMS_SMB': 'Microsoft Teams',
  'OFFICESUBSCRIPTION': 'Office 桌面应用',
  'O365_BUSINESS': 'Microsoft 365 商业应用',
  'O365_BUSINESS_PREMIUM': 'Microsoft 365 商业高级版',
  'PROJECT_CLIENT_SUBSCRIPTION': 'Project 桌面客户端',
  'VISIO_CLIENT_SUBSCRIPTION': 'Visio 桌面客户端',
  'WAC_EDU_A1': 'Office 网页版',
  'WAC_EDU_A3': 'Office 网页版',
  'WAC_EDU_A5': 'Office 网页版',
  'ONEDRIVE_BASIC': 'OneDrive',
  'ONEDRIVE_LITE': 'OneDrive',
  'ONEDRIVE_PREMIUM': 'OneDrive',
  'ONEDRIVE_ENT_SUB': 'OneDrive',
  'ONEDRIVE_ENT_SUB_DEVELOPER': 'OneDrive',
  'SWAY': 'Sway',
  'FORMS_PLAN_E1': 'Forms',
  'FORMS_PLAN_E3': 'Forms',
  'FORMS_PLAN_E5': 'Forms',
  'STREAM_O365_E1': 'Stream',
  'STREAM_O365_E3': 'Stream',
  'STREAM_O365_E5': 'Stream',
  'FLOW_O365_P1': 'Power Automate',
  'FLOW_O365_P2': 'Power Automate',
  'FLOW_O365_P3': 'Power Automate',
  'POWERAPPS_O365_P1': 'Power Apps',
  'POWERAPPS_O365_P2': 'Power Apps',
  'POWERAPPS_O365_P3': 'Power Apps',
  'YAMMER_ENTERPRISE': 'Yammer',
  'YAMMER_EDU': 'Yammer',
  'PROJECTWORKMANAGEMENT': 'Planner'
};

// 常见需要被限制的核心应用字典（planName → 应用分组）
const CORE_SERVICE_PLANS = {
  Exchange: ['EXCHANGE_S_ENTERPRISE', 'EXCHANGE_S_STANDARD', 'EXCHANGE_L_STANDARD', 'EXCHANGE_S_FOUNDATION', 'EXCHANGE_S_DESKLESS'],
  Teams: ['TEAMS1', 'TEAMS_FREE', 'TEAMS_EXPLORATORY', 'TEAMS_COMM', 'TEAMSPRO', 'TEAMS_PRO_EDU', 'TEAMS_SMB'],
  Office_Apps: ['OFFICESUBSCRIPTION', 'O365_BUSINESS', 'O365_BUSINESS_PREMIUM', 'PROJECT_CLIENT_SUBSCRIPTION', 'VISIO_CLIENT_SUBSCRIPTION'],
  Office_Web: ['WAC_EDU_A1', 'WAC_EDU_A3', 'WAC_EDU_A5'],
  OneDrive: ['ONEDRIVE_BASIC', 'ONEDRIVE_LITE', 'ONEDRIVE_PREMIUM', 'ONEDRIVE_ENT_SUB', 'ONEDRIVE_ENT_SUB_DEVELOPER', 'SHAREPOINTWAC', 'SHAREPOINTWAC_DEVELOPER', 'SHAREPOINTENTERPRISE', 'SHAREPOINTSTANDARD', 'SHAREPOINT_PROJECT_EDU', 'SHAREPOINT_L_STANDARD'],
  Sway: ['SWAY'],
  Forms: ['FORMS_PLAN_E1', 'FORMS_PLAN_E3', 'FORMS_PLAN_E5'],
  Stream: ['STREAM_O365_E1', 'STREAM_O365_E3', 'STREAM_O365_E5'],
  PowerAutomate: ['FLOW_O365_P1', 'FLOW_O365_P2', 'FLOW_O365_P3'],
  PowerApps: ['POWERAPPS_O365_P1', 'POWERAPPS_O365_P2', 'POWERAPPS_O365_P3'],
  Yammer: ['YAMMER_ENTERPRISE', 'YAMMER_EDU'],
  Planner: ['PROJECTWORKMANAGEMENT']
};

const DEFAULT_CONFIG = {
  adminPath: '/admin',
  adminUsername: 'admin',
  adminPasswordHash: '',
  turnstile: { siteKey: '', secretKey: '' },
  globals: [],
  protectedUsers: [],
  protectedPrefixes: ['admin', 'superadmin', 'root', 'administrator', 'sysadmin', 'owner', 'support', 'helpdesk'],
  invite: { enabled: false },
};

const GITHUB_LINK = 'https://github.com/muzi-cheng/CF-M365-Admin';
const enc = new TextEncoder();

/* -------------------- Utility -------------------- */
async function sha256(txt) {
  const buf = await crypto.subtle.digest('SHA-256', enc.encode(txt));
  return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, '0')).join('');
}

function jsonResponse(obj, status = 200, headers = {}) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { 'Content-Type': 'application/json', ...headers },
  });
}

function redirect(location, status = 302) {
  return new Response(null, { status, headers: { Location: location } });
}

function parseCookies(req) {
  const raw = req.headers.get('Cookie') || '';
  return Object.fromEntries(
    raw
      .split(';')
      .filter(Boolean)
      .map((c) => {
        const [k, ...v] = c.trim().split('=');
        return [k, v.join('=')];
      }),
  );
}

function mergeConfig(raw) {
  const base = structuredClone(DEFAULT_CONFIG);
  if (!raw || typeof raw !== 'object') return base;

  const cfg = { ...base, ...raw };
  cfg.turnstile = { ...base.turnstile, ...(raw.turnstile || {}) };
  cfg.invite = { ...base.invite, ...(raw.invite || {}) };
  cfg.globals = Array.isArray(raw.globals) ? raw.globals : base.globals;
  cfg.protectedUsers = Array.isArray(raw.protectedUsers) ? raw.protectedUsers : base.protectedUsers;
  cfg.protectedPrefixes = Array.isArray(raw.protectedPrefixes) ? raw.protectedPrefixes : base.protectedPrefixes;
  cfg.adminUsername = (raw.adminUsername || base.adminUsername || 'admin').toString().trim() || 'admin';
  cfg.adminPath = (raw.adminPath || base.adminPath || '/admin').toString().trim() || '/admin';
  cfg.adminPasswordHash = (raw.adminPasswordHash || base.adminPasswordHash || '').toString();
  return cfg;
}

async function getConfig(env) {
  const cfg = await env.CONFIG_KV.get(KV.CONFIG, 'json');
  return mergeConfig(cfg);
}

async function setConfig(env, cfg) {
  await env.CONFIG_KV.put(KV.CONFIG, JSON.stringify(cfg));
}

async function ensureInvites(env) {
  let data = await env.CONFIG_KV.get(KV.INVITES, 'json');
  if (!data) {
    const compat = await env.CONFIG_KV.get(KV.COMPAT_CARDS, 'json');
    if (compat) {
      await env.CONFIG_KV.put(KV.INVITES, JSON.stringify(compat));
      data = compat;
    } else {
      await env.CONFIG_KV.put(KV.INVITES, JSON.stringify([]));
      data = [];
    }
  }
  return data;
}

async function getInvites(env) {
  const data = await env.CONFIG_KV.get(KV.INVITES, 'json');
  if (data) return data;
  return ensureInvites(env);
}

async function saveInvites(env, list) {
  await env.CONFIG_KV.put(KV.INVITES, JSON.stringify(list));
}

async function createSession(env) {
  const token = crypto.randomUUID();
  await env.CONFIG_KV.put(KV.SESS_PREFIX + token, Date.now().toString(), {
    expirationTtl: 60 * 60 * 24 * 7,
  });
  return token;
}

async function verifySession(env, req) {
  const cookies = parseCookies(req);
  const token = cookies.ADMIN_SESSION;
  if (!token) return false;
  const val = await env.CONFIG_KV.get(KV.SESS_PREFIX + token);
  return !!val;
}

function htmlResponse(html, status = 200) {
  return new Response(html, {
    status,
    headers: { 'Content-Type': 'text/html;charset=UTF-8' },
  });
}

function sanitizeSkuMap(str) {
  try {
    const obj = typeof str === 'string' ? JSON.parse(str || '{}') : str || {};
    if (typeof obj !== 'object' || Array.isArray(obj)) return {};
    return obj;
  } catch {
    return {};
  }
}

function disableSelectIfSingle(arr) {
  return arr.length <= 1;
}

function checkPasswordComplexity(pwd) {
  if (!pwd || pwd.length < 8) return false;
  let s = 0;
  if (/[a-z]/.test(pwd)) s++;
  if (/[A-Z]/.test(pwd)) s++;
  if (/\d/.test(pwd)) s++;
  if (/[^a-zA-Z0-9]/.test(pwd)) s++;
  return s >= 3;
}

function randomFromPool(pool, len) {
  let s = '';
  for (let i = 0; i < len; i++) s += pool[Math.floor(Math.random() * pool.length)];
  return s;
}

function generateInviteCodeGrouped(prefix = 'ORZ', groups = 3, groupLen = 4) {
  const pool = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const parts = [];
  for (let i = 0; i < groups; i++) parts.push(randomFromPool(pool, groupLen));
  return prefix ? `${prefix}-${parts.join('-')}` : parts.join('-');
}

/**
 * 生成固定总长度的邀请码（纯大写字母+数字，不含易混淆字符）
 * @param {number} len 16 或 32
 */
function generateFixedLengthCode(len = 16) {
  const pool = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  return randomFromPool(pool, len);
}

function inviteCodeExists(invites, code) {
  return invites.some((x) => x.code === code);
}

function getEnvHiddenList(env) {
  if (!env.HIDDEN_USER) return [];
  return env.HIDDEN_USER.split(/[;,]/).map((s) => s.trim()).filter(Boolean);
}

function normalizeLower(v) {
  return (v || '').toString().trim().toLowerCase();
}

function getLocalPartFromUpn(upn) {
  const v = normalizeLower(upn);
  const at = v.indexOf('@');
  return at >= 0 ? v.slice(0, at) : v;
}

function buildProtectionSets(env, cfg) {
  const emailSet = new Set();
  (cfg.protectedUsers || []).forEach((u) => {
    const x = normalizeLower(u);
    if (x) emailSet.add(x);
  });
  getEnvHiddenList(env).forEach((u) => {
    const x = normalizeLower(u);
    if (x) emailSet.add(x);
  });

  const prefixSet = new Set();
  (cfg.protectedPrefixes || []).forEach((p) => {
    const x = normalizeLower(p);
    if (x) prefixSet.add(x);
  });

  return { emailSet, prefixSet };
}

function isProtectedUpn(upn, env, cfg) {
  const { emailSet, prefixSet } = buildProtectionSets(env, cfg);
  const v = normalizeLower(upn);
  if (!v) return false;
  if (emailSet.has(v)) return true;
  const local = getLocalPartFromUpn(v);
  return prefixSet.has(local);
}

function filterProtectedUsers(list, env, cfg) {
  const { emailSet, prefixSet } = buildProtectionSets(env, cfg);
  return list.filter((u) => {
    const upn = normalizeLower(u.userPrincipalName || '');
    if (emailSet.has(upn)) return false;
    const local = getLocalPartFromUpn(upn);
    return !prefixSet.has(local);
  });
}

/* -------------------- Microsoft Graph -------------------- */

// 内存级别的 Token 缓存（Worker 实例存活期间有效），避免每次请求都重新获取
const tokenCache = new Map(); // key: tenantId+clientId, value: { token, expiresAt }

async function getAccessTokenForGlobal(global, fetcher) {
  const cacheKey = `${global.tenantId}::${global.clientId}`;
  const now = Date.now();
  const cached = tokenCache.get(cacheKey);
  // 提前 60 秒过期，避免边界竞争
  if (cached && cached.expiresAt > now + 60_000) {
    return cached.token;
  }

  const params = new URLSearchParams();
  params.append('client_id', global.clientId);
  params.append('scope', 'https://graph.microsoft.com/.default');
  params.append('client_secret', global.clientSecret);
  params.append('grant_type', 'client_credentials');

  const res = await fetcher(`https://login.microsoftonline.com/${global.tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    body: params,
  });
  const data = await res.json().catch(() => ({}));
  if (!data.access_token) {
    throw new Error(data.error_description || data.error || '获取令牌失败');
  }
  // expires_in 单位为秒，默认 3600（1小时）
  const expiresIn = (Number(data.expires_in) || 3600) * 1000;
  tokenCache.set(cacheKey, { token: data.access_token, expiresAt: now + expiresIn });
  return data.access_token;
}

async function fetchServicePlans(global, skuId, fetcher) {
  const token = await getAccessTokenForGlobal(global, fetcher);
  const resp = await fetcher('https://graph.microsoft.com/v1.0/subscribedSkus', {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!resp.ok) {
    throw new Error('获取 Service Plans 失败');
  }
  const data = await resp.json();
  const skus = Array.isArray(data.value) ? data.value : [];
  const targetSku = skus.find(s => s.skuId.toLowerCase() === String(skuId).toLowerCase());
  
  if (!targetSku || !targetSku.servicePlans) return [];
  
  // 过滤并匹配核心应用
  const plans = targetSku.servicePlans.map(p => ({
    servicePlanId: p.servicePlanId,
    servicePlanName: p.servicePlanName,
    provisioningStatus: p.provisioningStatus
  }));

  const corePlans = [];
  for (const [appName, planNames] of Object.entries(CORE_SERVICE_PLANS)) {
    // 在当前 sku 的 servicePlans 中找匹配的
    const matched = plans.find(p => planNames.includes(p.servicePlanName));
    if (matched) {
      corePlans.push({
        appName,
        servicePlanId: matched.servicePlanId,
        servicePlanName: matched.servicePlanName
      });
    }
  }
  
  return corePlans;
}

// 带分页的 Graph API 用户列表拉取，自动跟随 @odata.nextLink
async function fetchAllGraphUsers(token, fetcher) {
  let url = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,createdDateTime,assignedLicenses&$top=999&$orderby=createdDateTime desc&$count=true';
  const headers = { Authorization: `Bearer ${token}`, ConsistencyLevel: 'eventual' };
  const allUsers = [];
  while (url) {
    const resp = await fetcher(url, { headers });
    const data = await resp.json().catch(() => ({}));
    const items = Array.isArray(data.value) ? data.value : [];
    allUsers.push(...items);
    url = data['@odata.nextLink'] || null;
  }
  return allUsers;
}

async function fetchSubscribedSkus(global, fetcher) {
  const token = await getAccessTokenForGlobal(global, fetcher);
  const resp = await fetcher('https://graph.microsoft.com/v1.0/subscribedSkus', {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error(err?.error?.message || '获取订阅 SKU 失败');
  }
  const data = await resp.json();
  return Array.isArray(data.value) ? data.value : [];
}

function remainingFromSubscribedSku(sku) {
  const enabled = Number(sku?.prepaidUnits?.enabled ?? 0);
  const consumed = Number(sku?.consumedUnits ?? 0);
  const remaining = enabled - consumed;
  return Number.isFinite(remaining) ? Math.max(0, remaining) : 0;
}

/* -------------------- UI -------------------- */
const baseStyles = `
:root {
  --primary: #111827;
  --primary-hover: #374151;
  --accent: #111827;
  --bg: #f9fafb;
  --surface: #ffffff;
  --border: #e5e7eb;
  --text-main: #111827;
  --text-sub: #6b7280;
  --text-muted: #9ca3af;
  --radius: 10px;
  --shadow-sm: 0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04);
  --shadow-md: 0 4px 12px rgba(0,0,0,0.08);
  --shadow-lg: 0 8px 24px rgba(0,0,0,0.10);
}
* { box-sizing: border-box; }
body {
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Inter', Roboto, sans-serif;
  background: var(--bg);
  margin: 0;
  padding: 0;
  color: var(--text-main);
  font-size: 14px;
  line-height: 1.6;
}
@keyframes fadeInUp { from {opacity:0; transform: translateY(12px);} to {opacity:1; transform: translateY(0);} }
a { color: var(--text-main); text-decoration: none; }
a:hover { color: var(--primary-hover); }
.card {
  background: var(--surface);
  padding: 32px;
  border-radius: 16px;
  border: 1px solid var(--border);
  box-shadow: var(--shadow-md);
  animation: fadeInUp 0.35s ease;
}
button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 6px;
  padding: 9px 16px;
  background: var(--primary);
  color: #fff;
  border: none;
  border-radius: var(--radius);
  font-weight: 600;
  font-size: 13px;
  cursor: pointer;
  transition: background .15s, opacity .15s;
  white-space: nowrap;
}
button svg { flex-shrink: 0; width: 15px; height: 15px; }
button:hover { background: var(--primary-hover); }
button:disabled { opacity: 0.45; cursor: not-allowed; }
input, select, textarea {
  width: 100%;
  padding: 10px 12px;
  border: 1px solid var(--border);
  border-radius: var(--radius);
  background: var(--surface);
  font-size: 14px;
  color: var(--text-main);
  transition: border-color .15s, box-shadow .15s;
  outline: none;
}
select {
  appearance: none;
  -webkit-appearance: none;
  background-image: url("data:image/svg+xml;charset=utf-8,%3Csvg xmlns='http://www.w3.org/2000/svg' fill='none' viewBox='0 0 20 20' stroke='%236b7280'%3E%3Cpath stroke-linecap='round' stroke-linejoin='round' stroke-width='1.5' d='M6 8l4 4 4-4'/%3E%3C/svg%3E");
  background-position: right 12px center;
  background-repeat: no-repeat;
  background-size: 16px 16px;
  padding-right: 36px;
  cursor: pointer;
}
select::-ms-expand { display: none; }
select option {
  padding: 10px;
  background: var(--surface);
  color: var(--text-main);
}
select:focus > option:checked {
  background: #f3f4f6 !important;
  color: var(--text-main);
}
input:focus, select:focus, textarea:focus {
  border-color: #6b7280;
  box-shadow: 0 0 0 3px rgba(107,114,128,0.12);
}
input::placeholder, textarea::placeholder { color: var(--text-muted); }
.tag { padding: 2px 8px; border-radius: 6px; background: #f3f4f6; color: #374151; font-size: 12px; display:inline-block; margin: 2px 3px 2px 0; border: 1px solid #e5e7eb; font-family: monospace; }
.table { width: 100%; border-collapse: collapse; }
.table thead tr th { padding: 10px 14px; text-align: left; color: var(--text-sub); font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .6px; cursor: pointer; user-select: none; background: #f9fafb; border-bottom: 1px solid var(--border); }
.table th .arrow { margin-left:4px; color: var(--text-muted); }
.table th.active { color: var(--text-main); }
.table th.active .arrow { color: var(--text-main); }
.table tbody tr { background: var(--surface); transition: background .1s; }
.table tbody tr:hover { background: #f9fafb; }
.table tbody tr td { padding: 12px 14px; border-bottom: 1px solid #f3f4f6; vertical-align: middle; color: var(--text-main); }
.table tbody tr:last-child td { border-bottom: none; }
.table-wrap { border-radius: var(--radius); overflow: hidden; border: 1px solid var(--border); }
.toolbar { display:flex; gap:8px; flex-wrap: wrap; margin-bottom: 12px; align-items:center; }
.pill { padding: 5px 12px; border: 1px solid var(--border); border-radius: 999px; font-size: 12px; background: var(--surface); cursor:pointer; color: var(--text-sub); transition: all .15s; }
.pill:hover { border-color: #9ca3af; color: var(--text-main); }
.pill.active { border-color: var(--text-main); color: var(--text-main); background: #f3f4f6; font-weight: 600; }
.input-compact { max-width:200px; }
.row { margin-bottom: 14px; }
.label { font-size: 12px; font-weight: 600; color: var(--text-sub); margin-bottom: 5px; display: block; text-transform: uppercase; letter-spacing: .4px; }
.custom-select{position:relative;}
.select-trigger{border:1px solid var(--border);border-radius:var(--radius);padding:10px 12px;display:flex;justify-content:space-between;align-items:center;background:var(--surface);cursor:pointer;gap:10px;transition:border-color .15s;}
.select-trigger:hover{border-color:#9ca3af;}
.select-trigger span{display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-size:14px;color:var(--text-main);user-select:none;}
.select-trigger.disabled{cursor:not-allowed;opacity:0.5;}
.select-arrow{flex:0 0 auto;width:8px;height:8px;border-right:1.5px solid var(--text-sub);border-bottom:1.5px solid var(--text-sub);transform:rotate(45deg) translateY(-2px);transition:transform .2s;}
.options-container{position:absolute;top:calc(100% + 4px);left:0;right:0;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);box-shadow:var(--shadow-md);opacity:0;visibility:hidden;transform:translateY(-4px);transition:opacity .2s,transform .2s,visibility .2s;z-index:50;overflow:hidden;max-height:48vh;overflow-y:auto;}
.options-container.open{opacity:1;visibility:visible;transform:translateY(0);}
.custom-select.open .select-arrow{transform:rotate(-135deg) translateY(-2px);}
.option{padding:10px 12px;font-size:14px;cursor:pointer;word-break:break-word;color:var(--text-main);user-select:none;}
.option:hover{background:#f3f4f6;}
.option.selected{background:#f3f4f6;font-weight:700;}
@media (max-width: 480px) {
  .card { padding: 20px; border-radius: 12px; }
  button { width: 100%; }
}
`;

const GITHUB_ICON = `<svg viewBox="0 0 16 16" width="18" height="18" aria-hidden="true" fill="currentColor" style="vertical-align:middle;"><path d="M8 0C3.58 0 0 3.58 0 8a8 8 0 0 0 5.47 7.59c.4.07.55-.17.55-.38l-.01-1.49C3.99 14.91 3.48 13.5 3.48 13.5c-.36-.92-.88-1.17-.88-1.17-.72-.5.06-.49.06-.49.79.06 1.2.82 1.2.82.71 1.21 1.86.86 2.31.66.07-.52.28-.86.5-1.06-2-.22-4.1-1-4.1-4.43 0-.98.35-1.78.92-2.41-.09-.22-.4-1.11.09-2.31 0 0 .76-.24 2.49.92a8.64 8.64 0 0 1 4.53 0c1.72-1.16 2.48-.92 2.48-.92.5 1.2.19 2.09.1 2.31.57.63.92 1.43.92 2.41 0 3.44-2.1 4.2-4.11 4.42.29.25.54.73.54 1.48l-.01 2.2c0 .21.15.46.55.38A8 8 0 0 0 16 8c0-4.42-3.58-8-8-8z"></path></svg>`;

const ICONS = {
  plus: `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M10 4v12M4 10h12"/></svg>`,
  refresh: `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M16 4v4h-4"/><path d="M4 16a6 6 0 0 0 10.5-3"/><path d="M4 12v-4h4"/><path d="M16 8a6 6 0 0 0-10.5 3"/></svg>`,
  chart: `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 16V8"/><path d="M10 16V4"/><path d="M16 16v-6"/></svg>`,
  key: `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="7" cy="10" r="3"/><path d="M10 10h6v3"/><path d="M16 13v3"/></svg>`,
  trash: `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 6h12"/><path d="M7 6v-2h6v2"/><path d="M8 9v5"/><path d="M12 9v5"/><path d="M6 6l1 10h6l1-10"/></svg>`,
  download: `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M10 3v9"/><path d="M6.5 9.5L10 13l3.5-3.5"/><path d="M4 16h12"/></svg>`,
  save: `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M5 4h9l2 2v10H5z"/><path d="M7 4v5h6V4"/><path d="M7 16v-4h6v4"/></svg>`,
  edit: `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 13.5V16h2.5L15 7.5l-2.5-2.5L4 13.5z"/><path d="M12.5 5l2.5 2.5"/></svg>`,
  search: `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="9" r="5"/><path d="M14 14l3 3"/></svg>`,
  spark: `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M10 3l1.6 4.4L16 9l-4.4 1.6L10 15l-1.6-4.4L4 9l4.4-1.6L10 3z"/></svg>`
};

function renderRegisterPage({ globals, selectedGlobalId, skuDisplayList, protectedPrefixes, turnstileSiteKey, inviteMode, adminPath }) {
  const disableGlobal = disableSelectIfSingle(globals);
  const selectedGlobal = globals.find((g) => g.id === selectedGlobalId) || globals[0] || null;
  const disableSku = disableSelectIfSingle(skuDisplayList);
  const globalOptions = globals.map((g) => {
    const sel = selectedGlobal && g.id === selectedGlobal.id ? 'selected' : '';
    return `<div class="option ${sel}" data-id="${g.id}">${g.label}</div>`;
  }).join('');
  const skuOptions = (list) => (list || []).map((x) => `<div class="option" data-value="${x.name}">${x.label}</div>`).join('');
  const siteKeyScript = turnstileSiteKey ? `<script src="https://challenges.cloudflare.com/turnstile/v0/api.js" async defer></script>` : '';
  const initialSkuName = skuDisplayList?.[0]?.name || '';
  const initialSkuLabel = skuDisplayList?.[0]?.label || '暂无 SKU';

  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>${inviteMode ? 'Office365 邀请码自助注册' : 'Office 365 自助开通'}</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>${baseStyles}
html,body{max-width:100%;overflow-x:hidden;}
body{display:flex;justify-content:center;align-items:center;min-height:100vh;padding:20px;background:var(--bg);}
.card{max-width:480px;width:100%;position:relative;}
.header-row{display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;gap:10px;flex-wrap:wrap;}
h2{margin:0;font-weight:800;color:var(--text-main);font-size:18px;letter-spacing:-.3px;}
.input-group{margin-bottom:14px;}
.hint{margin-top:5px;font-size:12px;line-height:1.5;color:var(--text-sub);}
.hint.error{color:#dc2626;font-weight:600;}
.message{margin-top:12px;padding:10px 12px;border-radius:8px;font-size:13px;display:none;line-height:1.6;}
.error{background:#fef2f2;color:#dc2626;border:1px solid #fecaca;}
.success{background:#f0fdf4;color:#16a34a;border:1px solid #bbf7d0;}
.form-actions{display:flex;justify-content:center;margin-top:16px;}
.form-actions button{min-width:200px;}
.cf-turnstile{display:flex;justify-content:center;margin:14px 0;}
.footer{margin-top:14px;font-size:12px;color:var(--text-muted);display:flex;gap:6px;align-items:center;justify-content:center;flex-wrap:wrap;text-align:center;}
.icon-link{display:flex;gap:5px;align-items:center;color:var(--text-muted);}
.icon-link:hover{color:var(--text-sub);}
.danger-modal{position:fixed;top:0;left:0;width:100%;height:100%;display:none;align-items:center;justify-content:center;background:rgba(0,0,0,0.4);backdrop-filter:blur(4px);z-index:2000;padding:16px;}
.danger-modal .dlg{width:92vw;max-width:440px;background:var(--surface);border-radius:12px;border:1px solid var(--border);box-shadow:var(--shadow-lg);overflow:hidden;}
.danger-modal .bar{background:#111827;color:#fff;padding:12px 16px;font-weight:700;display:flex;align-items:center;justify-content:space-between;font-size:14px;}
.danger-modal .bar .x{width:28px;height:28px;border-radius:6px;background:rgba(255,255,255,0.1);display:flex;align-items:center;justify-content:center;font-weight:700;cursor:pointer;font-size:14px;}
.danger-modal .content{padding:16px;line-height:1.7;color:var(--text-main);font-size:14px;}
.danger-modal .actions{padding:0 16px 16px;display:flex;gap:8px;}
.danger-modal .actions button{width:100%;background:#111827;}
.danger-modal .actions button:hover{background:#374151;}
@media (max-width: 480px) {
  body{padding:12px;}
  .form-actions button{width:100%;min-width:0;}
  button{width:100%;}
}
</style>
${siteKeyScript}
</head>
<body>
<div class="danger-modal" id="banModal" role="dialog" aria-modal="true">
  <div class="dlg">
    <div class="bar"><span>⚠️ 安全拦截</span><span class="x" onclick="hideBan()">✕</span></div>
    <div class="content">
      <div style="font-size:16px;font-weight:900;margin-bottom:8px;">该用户名被<strong>禁止注册</strong>！</div>
      <div>请勿尝试注册<strong>非法/敏感</strong>用户名，否则系统将持续拦截并记录行为。</div>
      <div style="margin-top:10px;color:#6b7280;font-size:12px;">建议更换一个普通用户名（仅字母/数字）。</div>
    </div>
    <div class="actions"><button type="button" onclick="hideBan()">我已知晓</button></div>
  </div>
</div>

<div class="card">
  <div class="header-row">
    <h2>${inviteMode ? 'Office365 邀请码自助注册' : 'Office 365 自助开通'}</h2>
    <a class="icon-link" href="${GITHUB_LINK}" target="_blank" title="View Source">${GITHUB_ICON}</a>
  </div>

  <form id="regForm">
    <input type="hidden" name="globalId" id="globalId" value="${selectedGlobal ? selectedGlobal.id : ''}">
    <input type="hidden" name="skuName" id="skuName" value="${initialSkuName}">

    <div class="input-group">
      <span class="label">选择全局</span>
      <div class="custom-select">
        <div class="select-trigger ${disableGlobal ? 'disabled' : ''}" id="globalTrigger">
          <span>${selectedGlobal ? selectedGlobal.label : '无可用全局'}</span>
          <div class="select-arrow"></div>
        </div>
        <div class="options-container" id="globalOptions">${globalOptions}</div>
      </div>
    </div>

    <div class="input-group">
      <span class="label">选择订阅</span>
      <div class="custom-select">
        <div class="select-trigger ${disableSku ? 'disabled' : ''}" id="skuTrigger">
          <span>${initialSkuLabel}</span>
          <div class="select-arrow"></div>
        </div>
        <div class="options-container" id="skuOptions">${skuOptions(skuDisplayList)}</div>
      </div>
    </div>

    <div class="input-group">
      <span class="label">用户名</span>
      <input type="text" id="username" required pattern="[a-zA-Z0-9]+" placeholder="例如 user123" autocomplete="off">
      <div class="hint" id="userHint"></div>
    </div>

    <div class="input-group">
      <span class="label">密码</span>
      <input type="password" id="password" required placeholder="至少8位，包含大/小写/数字/符号中的3类" autocomplete="new-password">
      <div class="hint" id="pwdHint"></div>
    </div>

    ${inviteMode ? `<div class="input-group"><span class="label">邀请码</span><input type="text" id="inviteCode" required placeholder="请输入有效邀请码"></div>` : ''}

    ${turnstileSiteKey ? `<div class="cf-turnstile" data-sitekey="${turnstileSiteKey}"></div>` : ''}

    <div class="form-actions"><button type="submit" id="btn">立即创建账号</button></div>
    <div id="msg" class="message"></div>
  </form>

  <div class="footer">
    <span>Powered by Cloudflare Workers</span>
    <a class="icon-link" href="${GITHUB_LINK}" target="_blank">${GITHUB_ICON} CF-M365-Admin</a>
  </div>
</div>

<script>
const selectedGlobalId = ${JSON.stringify(selectedGlobal ? selectedGlobal.id : '')};
const protectedPrefixes = ${JSON.stringify((protectedPrefixes || []).map((s) => String(s).toLowerCase()))};
const inviteMode = ${inviteMode ? 'true' : 'false'};
const turnstileOn = ${turnstileSiteKey ? 'true' : 'false'};

function openSelect(triggerId, containerId, disabled) {
  const trigger = document.getElementById(triggerId);
  const container = document.getElementById(containerId);
  const wrapper = trigger.parentElement;
  
  const optCount = container.querySelectorAll('.option').length;
  if (disabled || optCount <= 1) {
    trigger.classList.add('disabled');
    trigger.style.cursor = 'default';
    trigger.addEventListener('click', (e) => e.stopPropagation());
    return;
  } else {
    trigger.classList.remove('disabled');
    trigger.style.cursor = 'pointer';
  }
  
  trigger.addEventListener('click', (e) => { e.stopPropagation(); container.classList.toggle('open'); wrapper.classList.toggle('open'); });
  document.addEventListener('click', (e) => { if(!wrapper.contains(e.target)) {container.classList.remove('open'); wrapper.classList.remove('open');} });
  if(wrapper) {
    let leaveTimer;
    wrapper.addEventListener('mouseleave', () => { leaveTimer = setTimeout(() => { container.classList.remove('open'); wrapper.classList.remove('open'); }, 200); });
    wrapper.addEventListener('mouseenter', () => clearTimeout(leaveTimer));
  }
}
openSelect('globalTrigger', 'globalOptions', ${disableGlobal ? 'true' : 'false'});
openSelect('skuTrigger', 'skuOptions', ${disableSku ? 'true' : 'false'});

document.querySelectorAll('#globalOptions .option').forEach((opt) => {
  opt.addEventListener('click', () => {
    const gid = opt.getAttribute('data-id');
    if (!gid || gid === selectedGlobalId) return;
    const u = new URL(location.href);
    u.searchParams.set('g', gid);
    location.href = u.toString();
  });
});

document.querySelectorAll('#skuOptions .option').forEach((opt) => {
  opt.addEventListener('click', () => {
    const v = opt.getAttribute('data-value');
    document.getElementById('skuName').value = v || '';
    document.getElementById('skuTrigger').querySelector('span').innerText = opt.innerText;
    document.getElementById('skuOptions').classList.remove('open');
  });
});

function showBan(){ document.getElementById('banModal').style.display='flex'; }
function hideBan(){ document.getElementById('banModal').style.display='none'; }

function checkComplexity(pwd) {
  if(!pwd || pwd.length < 8) return false;
  let s=0;
  if(/[a-z]/.test(pwd)) s++;
  if(/[A-Z]/.test(pwd)) s++;
  if(/\\d/.test(pwd)) s++;
  if(/[^a-zA-Z0-9]/.test(pwd)) s++;
  return s >= 3;
}

function isBannedUsername(name){
  const u = (name||'').trim().toLowerCase();
  return !!u && protectedPrefixes.includes(u);
}

const btn = document.getElementById('btn');
const userEl = document.getElementById('username');
const pwdEl = document.getElementById('password');
const userHint = document.getElementById('userHint');
const pwdHint = document.getElementById('pwdHint');

function validateForm(){
  const username = userEl.value.trim();
  const password = pwdEl.value || '';
  let ok = true;

  if(username && !/^[a-zA-Z0-9]+$/.test(username)){
    userHint.className='hint error';
    userHint.innerText='仅限字母和数字';
    ok=false;
  } else if(isBannedUsername(username)){
    userHint.className='hint error';
    userHint.innerText='敏感用户名，禁止注册';
    ok=false;
  } else {
    userHint.className='hint';
    userHint.innerText='';
  }

  if(password && !checkComplexity(password)){
    pwdHint.className='hint error';
    pwdHint.innerText='密码复杂度不足';
    ok=false;
  } else {
    pwdHint.className='hint';
    pwdHint.innerText='';
  }

  if(username && password && password.toLowerCase().includes(username.toLowerCase())){
    pwdHint.className='hint error';
    pwdHint.innerText='密码不能包含用户名';
    ok=false;
  }

  const globalId = document.getElementById('globalId').value;
  const skuName = document.getElementById('skuName').value;
  if(!globalId || !skuName) ok=false;

  btn.disabled = !ok;
  return ok;
}

userEl.addEventListener('input', validateForm);
pwdEl.addEventListener('input', validateForm);
validateForm();

document.getElementById('regForm').addEventListener('submit', async (e)=>{
  e.preventDefault();
  const msg = document.getElementById('msg');
  const username = userEl.value.trim();

  if(isBannedUsername(username)){
    showBan();
    msg.className='message error';
    msg.style.display='block';
    msg.innerText='❌ 该用户名被禁止注册';
    return;
  }

  if(!validateForm()){
    msg.className='message error';
    msg.style.display='block';
    msg.innerText='❌ 请修正表单错误';
    return;
  }

  const password = pwdEl.value;
  const skuName = document.getElementById('skuName').value;
  const globalId = document.getElementById('globalId').value;
  const inviteCode = inviteMode ? document.getElementById('inviteCode').value.trim() : '';

  if(inviteMode && !inviteCode){
    msg.className='message error';
    msg.style.display='block';
    msg.innerText='请填写邀请码';
    return;
  }

  if(turnstileOn){
    const v = document.querySelector('[name="cf-turnstile-response"]');
    if(!v || !v.value){
      msg.className='message error';
      msg.style.display='block';
      msg.innerText='请先完成人机验证';
      return;
    }
  }

  btn.disabled = true;
  btn.innerText = '正在创建...';
  msg.style.display='none';

  const form = new FormData();
  form.append('username', username);
  form.append('password', password);
  form.append('skuName', skuName);
  form.append('globalId', globalId);
  if(inviteMode) form.append('inviteCode', inviteCode);
  if(turnstileOn){
    const v = document.querySelector('[name="cf-turnstile-response"]');
    form.append('cf-turnstile-response', v ? v.value : '');
  }

  try {
    const res = await fetch('/', { method:'POST', body: form });
    const data = await res.json();
    msg.style.display='block';

    if(data.success){
      msg.className='message success';
      msg.innerHTML = '🎉 开通成功！<br>账号: ' + data.email + '<br>密码: (您刚才设置的)<br><a href="https://portal.office.com" target="_blank" style="color:#166534;font-weight:900;">前往 Office.com 登录</a>';
      document.getElementById('regForm').reset();
    } else {
      msg.className='message error';
      msg.innerText = '❌ ' + (data.message || '失败');
      if((data.message||'').includes('禁止注册')) showBan();
    }

    if(turnstileOn && typeof turnstile !== 'undefined') turnstile.reset();
  } catch(err) {
    msg.className='message error';
    msg.style.display='block';
    msg.innerText='网络异常，请稍后重试';
    if(turnstileOn && typeof turnstile !== 'undefined') turnstile.reset();
  } finally {
    btn.disabled=false;
    btn.innerText='立即创建账号';
    validateForm();
  }
});

window.hideBan = hideBan;
</script>
</body>
</html>`;
}

function adminLayout({ title, content, adminPath, active }) {
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>${title}</title>
<style>${baseStyles}
html,body{max-width:100%;overflow-x:hidden;}
body{background:var(--bg);padding:0;margin:0;}
.nav{background:var(--surface);border-bottom:1px solid var(--border);padding:0 28px;display:grid;grid-template-columns:1fr auto 1fr;align-items:center;height:56px;}
.nav-left{display:flex;align-items:center;gap:16px;}
.nav-logo{font-weight:700;font-size:15px;color:var(--text-main);letter-spacing:-.3px;}
.nav a{color:var(--text-sub);font-size:13px;font-weight:500;display:flex;align-items:center;gap:5px;}
.nav a:hover{color:var(--text-main);}
.tabs{display:flex;gap:2px;justify-content:center;}
.tab{padding:7px 14px;border-radius:8px;color:var(--text-sub);text-decoration:none;font-weight:500;font-size:13px;transition:background .15s,color .15s;}
.tab:hover{background:#f3f4f6;color:var(--text-main);}
.tab.active{background:#f3f4f6;color:var(--text-main);font-weight:700;}
.container{max-width:1280px;margin:24px auto;padding:0 20px;}
.section{background:var(--surface);border-radius:12px;border:1px solid var(--border);padding:20px;margin-bottom:16px;}
.section h3{margin:0 0 14px;font-size:14px;font-weight:700;color:var(--text-main);}
.badge{padding:2px 8px;border-radius:6px;background:#f3f4f6;color:var(--text-sub);font-weight:600;font-size:11px;letter-spacing:.3px;border:1px solid var(--border);}
.table-wrap{overflow-x:auto;}
input[type=checkbox]{width:15px;height:15px;accent-color:var(--primary);}
.modal{position:fixed;top:0;left:0;width:100%;height:100%;display:none;align-items:center;justify-content:center;background:rgba(0,0,0,0.3);backdrop-filter:blur(4px);z-index:1000;}
.modal .dialog{background:var(--surface);border-radius:12px;border:1px solid var(--border);padding:20px;min-width:320px;max-width:92vw;max-height:88vh;overflow:visible;box-shadow:var(--shadow-lg);animation:fadeInUp .2s;}
.modal .dialog.no-padding{padding:0;overflow:hidden;border:none;}
.modal .header{display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;padding-bottom:12px;border-bottom:1px solid var(--border);}
.modal .header h3{font-size:14px;font-weight:700;}
.modal .footer{display:flex;justify-content:flex-end;gap:8px;margin-top:16px;padding-top:12px;border-top:1px solid var(--border);}
.modal-close{width:28px;height:28px;padding:0;border-radius:6px;background:#f3f4f6;color:var(--text-sub);display:flex;align-items:center;justify-content:center;font-weight:700;line-height:1;font-size:14px;}
.modal-close:hover{background:#e5e7eb;}
.btn-ghost{background:#f3f4f6;color:var(--text-main);}
.btn-ghost:hover{background:#e5e7eb;}
.btn-danger{background:#dc2626;}
.btn-danger:hover{background:#b91c1c;}
label.inline{display:flex;align-items:center;gap:8px;margin:6px 0;font-size:13px;cursor:pointer;}
.pagination{display:flex;align-items:center;gap:6px;flex-wrap:wrap;font-size:13px;}
.page-input{width:80px;}
.search-box{display:flex;gap:8px;flex-wrap:wrap;align-items:center;}
.toast-wrap{position:fixed;top:24px;left:50%;transform:translateX(-50%);z-index:3000;display:flex;flex-direction:column;align-items:center;gap:8px;pointer-events:none;}
.toast{pointer-events:auto;display:flex;align-items:flex-start;gap:10px;background:var(--surface);color:var(--text-main);padding:12px 16px;border-radius:10px;box-shadow:var(--shadow-lg);border:1px solid var(--border);font-size:13px;line-height:1.5;animation:toastIn .2s ease;min-width:200px;max-width:320px;}
.toast .toast-icon{font-size:16px;flex:0 0 auto;}
.toast .toast-body{flex:1;}
.toast .toast-title{font-weight:700;font-size:12px;margin-bottom:1px;color:var(--text-sub);text-transform:uppercase;letter-spacing:.4px;}
.toast .toast-msg{font-size:13px;color:var(--text-main);}
.toast.success .toast-title{color:#16a34a;}
.toast.error .toast-title{color:#dc2626;}
.toast.info .toast-title{color:#2563eb;}
@keyframes toastIn{from{opacity:0;transform:translateY(-12px);}to{opacity:1;transform:translateY(0);}}
@keyframes spin{from{transform:rotate(0deg);}to{transform:rotate(360deg);}}
.global-loader { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(255, 255, 255, 0.7); backdrop-filter: blur(2px); z-index: 9999; display: none; align-items: center; justify-content: center; }
.spinner { width: 40px; height: 40px; border-radius: 50%; border: 3px solid rgba(17, 24, 39, 0.1); border-top-color: var(--primary); animation: spin 0.8s linear infinite; }
@media (max-width: 720px){
  .nav{grid-template-columns:1fr;height:auto;padding:12px 16px;gap:10px;}
  .tabs{justify-content:flex-start;flex-wrap:wrap;gap:4px;}
  .tab{font-size:13px;padding:7px 12px;}
  .container{margin:14px auto;padding:0 12px;}
  .section{padding:14px;}
  .input-compact{max-width:100%;}
  .modal .dialog{min-width:unset;width:94vw;}
  .toolbar{gap:6px;}
  .search-box{width:100%;}
  .pagination{gap:4px;}
  .page-input{width:70px;}
  .toast-wrap{bottom:12px;right:12px;left:12px;align-items:stretch;}
  .toast{max-width:none;}
}
@media (max-width: 720px){
  .table-wrap{overflow-x:visible;}
  .table thead{display:none;}
  .table tr{display:block;background:var(--surface);border-radius:10px;border:1px solid var(--border);margin-bottom:8px;overflow:hidden;}
  .table td{display:flex;justify-content:space-between;align-items:flex-start;gap:10px;width:100%;padding:9px 12px;word-break:break-word;border-bottom:1px solid #f3f4f6;}
  .table td:last-child{border-bottom:none;}
  .table td::before{content:attr(data-label);font-weight:700;color:var(--text-sub);font-size:11px;min-width:80px;text-transform:uppercase;letter-spacing:.4px;}
  .table td:first-child{justify-content:flex-start;}
  .table td:first-child::before{content:'';min-width:0;}
  .table td code{word-break:break-all;}
  .tag{white-space:normal;}
}
</style>
</head>
<body>
<div class="nav">
  <div class="nav-left">
    <span class="nav-logo">Office 365 Admin</span>
    <span class="badge">安全</span>
  </div>
  <div class="tabs">
    <a class="tab ${active === 'users' ? 'active' : ''}" href="${adminPath}/users">用户</a>
    <a class="tab ${active === 'invites' ? 'active' : ''}" href="${adminPath}/invites">邀请码</a>
    <a class="tab ${active === 'globals' ? 'active' : ''}" href="${adminPath}/globals">全局账户</a>
    <a class="tab ${active === 'settings' ? 'active' : ''}" href="${adminPath}/settings">设置</a>
  </div>
  <div style="display:flex;justify-content:flex-end;gap:12px;">
    <a href="${GITHUB_LINK}" target="_blank" style="display:flex;align-items:center;gap:6px;color:var(--text-muted);font-size:12px;">${GITHUB_ICON}<span>GitHub</span></a>
    <a href="${adminPath}/logout" style="display:flex;align-items:center;gap:6px;color:var(--text-muted);font-size:12px;cursor:pointer;"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"></path><polyline points="16 17 21 12 16 7"></polyline><line x1="21" y1="12" x2="9" y2="12"></line></svg><span>退出</span></a>
  </div>
</div>
<div class="container">${content}</div>
<div class="toast-wrap" id="toastWrap"></div>
<div class="global-loader" id="globalLoader"><div class="spinner"></div></div>
<script>
function showToast(message, type = 'info', duration = 2200) {
  const wrap = document.getElementById('toastWrap');
  if (!wrap) return;
  const icons = { success: '✅', error: '❌', info: 'ℹ️' };
  const titles = { success: '成功', error: '错误', info: '提示' };
  const el = document.createElement('div');
  el.className = 'toast ' + type;
  el.innerHTML = '<span class="toast-icon">' + (icons[type] || icons.info) + '</span><div class="toast-body"><div class="toast-title">' + (titles[type] || titles.info) + '</div><div class="toast-msg">' + message + '</div></div>';
  wrap.appendChild(el);
  setTimeout(() => {
    el.style.opacity = '0';
    el.style.transform = 'translateY(-10px) scale(0.95)';
    el.style.transition = 'all .22s ease';
    setTimeout(() => el.remove(), 230);
  }, duration);
}
window.showToast = showToast;
</script>
</body>
</html>`;
}

/* -------------------- Pages -------------------- */
function render404(adminPath) {
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>404 Not Found</title>
<style>${baseStyles}
body{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:100vh;text-align:center;}
h1{font-size:64px;font-weight:900;color:var(--text-main);margin:0;line-height:1;}
.desc{font-size:16px;color:var(--text-sub);margin:16px 0 32px;}
</style>
</head>
<body>
  <h1>404</h1>
  <div class="desc">页面未找到 / Page Not Found</div>
  <a href="/" style="display:inline-flex;padding:10px 20px;background:var(--primary);color:#fff;border-radius:var(--radius);font-weight:600;text-decoration:none;">返回首页</a>
</body>
</html>`;
}

function renderSetup(adminPath) {
  return `<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>初始化安装</title>
<style>${baseStyles}
body{display:flex;justify-content:center;align-items:center;min-height:100vh;padding:20px;}
.card{width:100%;max-width:520px;}
h2{margin:0 0 12px 0;}
.desc{color:#6b7280;font-size:14px;margin-bottom:16px;}
.row{margin-bottom:14px;}
.helper{font-size:12px;color:#6b7280;margin-top:6px;line-height:1.5;}
</style></head><body>
<div class="card">
  <h2>首次安装</h2>
  <div class="desc">初始化后台账号与路径。</div>
  <form id="setupForm">
    <div class="row"><span class="label">用户名</span><input type="text" id="user" required placeholder="admin" pattern="[a-zA-Z0-9_\\-]{3,32}"><div class="helper">3-32位字母/数字/_/-</div></div>
    <div class="row"><span class="label">密码</span><input type="password" id="pwd" required placeholder="至少 8 位"></div>
    <div class="row"><span class="label">后台路径</span><input type="text" id="path" value="${adminPath}" required pattern="\\/[a-zA-Z0-9\\-_/]+"></div>
    <button type="submit" id="btn">保存并进入后台</button>
  </form>
  <div id="msg" class="message" style="display:none;"></div>
  <div class="footer" style="margin-top:14px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;">${GITHUB_ICON}<a href="${GITHUB_LINK}" target="_blank">CF-M365-Admin</a></div>
</div>
<script>
document.getElementById('setupForm').addEventListener('submit', async (e)=>{
  e.preventDefault();
  const username = (document.getElementById('user').value || '').trim();
  const pwd = document.getElementById('pwd').value;
  const path = (document.getElementById('path').value || '/admin').trim();
  const btn = document.getElementById('btn');
  const msg = document.getElementById('msg');

  if(!/^[a-zA-Z0-9_\\-]{3,32}$/.test(username)){
    msg.innerText='用户名格式不正确（3-32位，仅字母/数字/_/-）';
    msg.className='message error'; msg.style.display='block'; return;
  }
  if(!pwd || pwd.length<8){
    msg.innerText='密码至少 8 位';
    msg.className='message error'; msg.style.display='block'; return;
  }
  btn.disabled=true; btn.innerText='正在保存...'; msg.style.display='none';
  const res = await fetch('${adminPath}/setup',{
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({username, password:pwd, adminPath:path})
  });
  const data = await res.json();
  if(data.success){ window.location.href = path; }
  else {
    msg.className='message error'; msg.style.display='block';
    msg.innerText=data.message||'保存失败';
    btn.disabled=false; btn.innerText='保存并进入后台';
  }
});
</script>
</body></html>`;
}

function renderLogin(adminPath) {
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Office 365 Admin</title>
<style>
*{box-sizing:border-box;margin:0;padding:0;}
:root{
  --primary:#111827;
  --primary-hover:#1f2937;
  --border:#e5e7eb;
  --bg:#f9fafb;
  --surface:#ffffff;
  --text-main:#111827;
  --text-sub:#6b7280;
  --text-muted:#9ca3af;
  --radius:12px;
  --shadow:0 4px 24px rgba(0,0,0,0.08);
}
body{
  font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','Inter',Roboto,sans-serif;
  background:var(--bg);
  min-height:100vh;
  display:flex;
  flex-direction:column;
  align-items:center;
  justify-content:center;
  padding:20px;
  color:var(--text-main);
}
@keyframes fadeUp{from{opacity:0;transform:translateY(14px);}to{opacity:1;transform:translateY(0);}}
.login-card{
  width:100%;
  max-width:400px;
  background:var(--surface);
  border:1px solid var(--border);
  border-radius:20px;
  box-shadow:var(--shadow);
  padding:36px 32px 28px;
  animation:fadeUp .3s ease;
}
.logo-area{
  display:flex;
  flex-direction:column;
  align-items:center;
  gap:10px;
  margin-bottom:28px;
}
.logo-title{
  font-size:20px;
  font-weight:800;
  color:var(--text-main);
  letter-spacing:-.5px;
}
.logo-sub{
  font-size:13px;
  color:var(--text-sub);
  margin-top:-4px;
}
.field{margin-bottom:14px;}
.field-label{
  display:block;
  font-size:12px;
  font-weight:600;
  color:var(--text-sub);
  text-transform:uppercase;
  letter-spacing:.5px;
  margin-bottom:6px;
}
.field-input{
  width:100%;
  padding:11px 14px;
  border:1px solid var(--border);
  border-radius:var(--radius);
  font-size:14px;
  color:var(--text-main);
  background:var(--surface);
  outline:none;
  transition:border-color .15s,box-shadow .15s;
}
.field-input:focus{
  border-color:#6b7280;
  box-shadow:0 0 0 3px rgba(107,114,128,0.12);
}
.field-input::placeholder{color:var(--text-muted);}
.submit-btn{
  width:100%;
  padding:12px;
  margin-top:6px;
  background:var(--primary);
  color:#fff;
  border:none;
  border-radius:var(--radius);
  font-size:15px;
  font-weight:700;
  cursor:pointer;
  transition:background .15s,opacity .15s;
  letter-spacing:-.1px;
}
.submit-btn:hover{background:var(--primary-hover);}
.submit-btn:disabled{opacity:.45;cursor:not-allowed;}
.msg-box{
  margin-top:12px;
  padding:10px 14px;
  border-radius:9px;
  font-size:13px;
  display:none;
  line-height:1.6;
}
.msg-box.error{background:#fef2f2;color:#dc2626;border:1px solid #fecaca;}
.footer-link{
  margin-top:22px;
  display:flex;
  align-items:center;
  justify-content:center;
  gap:6px;
  color:var(--text-muted);
  font-size:12px;
  text-decoration:none;
  transition:color .15s;
}
.footer-link:hover{color:var(--text-sub);}
@media(max-width:480px){
  .login-card{padding:28px 20px 22px;}
  .logo-title{font-size:18px;}
}
</style>
</head>
<body>
<div class="login-card">
  <div class="logo-area">
    <div class="logo-title">Office 365 Admin</div>
    <div class="logo-sub">管理员登录</div>
  </div>
  <form id="loginForm" autocomplete="on">
    <div class="field">
      <label class="field-label" for="user">用户名</label>
      <input class="field-input" type="text" id="user" name="username" required placeholder="管理员用户名" autocomplete="username">
    </div>
    <div class="field">
      <label class="field-label" for="pwd">密码</label>
      <input class="field-input" type="password" id="pwd" name="password" required placeholder="管理员密码" autocomplete="current-password">
    </div>
    <button class="submit-btn" type="submit" id="btn">登 录</button>
  </form>
  <div class="msg-box" id="msg"></div>
  <a class="footer-link" href="${GITHUB_LINK}" target="_blank">
    <svg viewBox="0 0 16 16" width="15" height="15" fill="currentColor"><path d="M8 0C3.58 0 0 3.58 0 8a8 8 0 0 0 5.47 7.59c.4.07.55-.17.55-.38l-.01-1.49C3.99 14.91 3.48 13.5 3.48 13.5c-.36-.92-.88-1.17-.88-1.17-.72-.5.06-.49.06-.49.79.06 1.2.82 1.2.82.71 1.21 1.86.86 2.31.66.07-.52.28-.86.5-1.06-2-.22-4.1-1-4.1-4.43 0-.98.35-1.78.92-2.41-.09-.22-.4-1.11.09-2.31 0 0 .76-.24 2.49.92a8.64 8.64 0 0 1 4.53 0c1.72-1.16 2.48-.92 2.48-.92.5 1.2.19 2.09.1 2.31.57.63.92 1.43.92 2.41 0 3.44-2.1 4.2-4.11 4.42.29.25.54.73.54 1.48l-.01 2.2c0 .21.15.46.55.38A8 8 0 0 0 16 8c0-4.42-3.58-8-8-8z"/></svg>
    CF-M365-Admin
  </a>
</div>
<script>
document.getElementById('loginForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const username = (document.getElementById('user').value || '').trim();
  const pwd = document.getElementById('pwd').value;
  const btn = document.getElementById('btn');
  const msg = document.getElementById('msg');
  btn.disabled = true;
  btn.innerText = '验证中...';
  msg.style.display = 'none';
  try {
    const res = await fetch('${adminPath}', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password: pwd })
    });
    const data = await res.json();
    if (data.success) {
      window.location.href = '${adminPath}/users';
    } else {
      msg.className = 'msg-box error';
      msg.style.display = 'block';
      msg.innerText = data.message || '用户名或密码错误';
      btn.disabled = false;
      btn.innerText = '登 录';
    }
  } catch {
    msg.className = 'msg-box error';
    msg.style.display = 'block';
    msg.innerText = '网络异常，请稍后重试';
    btn.disabled = false;
    btn.innerText = '登 录';
  }
});
</script>
</body>
</html>`;
}

function renderUsersPage(adminPath) {
  return adminLayout({
    title: '用户管理',
    adminPath,
    active: 'users',
    content: `
<div class="section">
  <div class="toolbar">
    <button id="btnRefresh" class="btn-ghost">${ICONS.refresh} 刷新</button>
    <button id="btnAddUser">${ICONS.plus} 新增用户</button>
    <button id="btnLic" class="btn-ghost">${ICONS.chart} 查看订阅</button>
    <button id="btnPwd" class="btn-ghost">${ICONS.key} 重置密码</button>
    <button id="btnDel" class="btn-danger">${ICONS.trash} 批量删除</button>
  </div>
  <div class="toolbar search-box" style="flex-wrap:wrap;gap:8px;">
    <div id="globalFilters" style="display:contents;"></div>
    <div style="width:1px;height:20px;background:var(--border);margin:0 2px;flex-shrink:0;"></div>
    <span class="label" style="margin:0;white-space:nowrap;">搜索：</span>
    <div class="custom-select" style="min-width:100px;max-width:120px;">
      <div class="select-trigger" id="searchFieldTrigger" style="padding:6px 10px;font-size:12px;">
        <span id="searchFieldDisplay">用户名</span>
        <div class="select-arrow"></div>
      </div>
      <div class="options-container" id="searchFieldOptions" style="min-width:100px;">
        <div class="option selected" data-value="displayName">用户名</div>
        <div class="option" data-value="userPrincipalName">账号</div>
        <div class="option" data-value="license">订阅</div>
        <div class="option" data-value="_globalLabel">全局</div>
      </div>
    </div>
    <input type="hidden" id="searchField" value="displayName">
    <input id="searchText" class="input-compact" placeholder="输入关键词，支持模糊" style="max-width:180px;">
    <button id="btnSearch" class="btn-ghost">搜索</button>
    <button id="btnClear" class="btn-ghost">清空</button>
  </div>
  <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin:8px 0;">
    <div style="display:flex;align-items:center;gap:6px;">
      <span style="font-size:12px;color:var(--text-sub);">每页</span>
      <div class="custom-select" style="min-width:70px;">
        <div class="select-trigger" id="pageSizeTrigger" style="padding:6px 10px;font-size:12px;">
          <span id="pageSizeDisplay">20</span>
          <div class="select-arrow"></div>
        </div>
        <div class="options-container" id="pageSizeOptions" style="min-width:70px;">
          <div class="option selected" data-value="20">20</div>
          <div class="option" data-value="30">30</div>
          <div class="option" data-value="50">50</div>
          <div class="option" data-value="100">100</div>
        </div>
      </div>
      <span style="font-size:12px;color:var(--text-sub);">条</span>
    </div>
    <div style="display:flex;align-items:center;gap:6px;">
      <span id="pageInfo" style="font-size:12px;color:var(--text-sub);white-space:nowrap;"></span>
      <button id="prevPage" class="btn-ghost" style="padding:5px 10px;font-size:12px;">上一页</button>
      <button id="nextPage" class="btn-ghost" style="padding:5px 10px;font-size:12px;">下一页</button>
      <input class="page-input" id="jumpPage" type="number" min="1" placeholder="页码" style="width:64px;padding:5px 8px;font-size:12px;">
      <button id="goPage" class="btn-ghost" style="padding:5px 10px;font-size:12px;">跳转</button>
    </div>
  </div>
  <div id="status" style="color:#107c10;font-weight:700;margin-bottom:6px;"></div>
  <div class="table-wrap">
    <table class="table" id="userTable">
      <thead>
        <tr>
          <th><input type="checkbox" id="chkAll"></th>
          <th data-sort="displayName">用户名 <span class="arrow" id="arr-displayName">↕</span></th>
          <th data-sort="userPrincipalName">账号 <span class="arrow" id="arr-userPrincipalName">↕</span></th>
          <th data-sort="_licSort">订阅 <span class="arrow" id="arr-_licSort">↕</span></th>
          <th data-sort="createdDateTime">创建时间 <span class="arrow" id="arr-createdDateTime">↕</span></th>
          <th data-sort="_globalLabel">全局 <span class="arrow" id="arr-_globalLabel">↕</span></th>
          <th>UUID</th>
        </tr>
      </thead>
      <tbody id="userBody"></tbody>
    </table>
  </div>
</div>

<div class="modal" id="modalPwd">
  <div class="dialog" style="max-width:440px;">
    <div class="header" style="margin-bottom:20px;"><h3 style="margin:0;font-size:16px;">重置密码</h3><button class="modal-close" onclick="closeModal('modalPwd')">✕</button></div>
    
    <div style="display:flex;flex-direction:column;gap:12px;margin-bottom:20px;">
      <label class="pwd-card active" id="pwdCardAuto">
        <div style="display:flex;align-items:center;gap:12px;">
          <input type="radio" name="pwdType" value="auto" checked style="width:18px;height:18px;accent-color:var(--primary);">
          <div>
            <div style="font-weight:700;font-size:14px;color:var(--text-main);">自动生成强密码</div>
            <div style="font-size:12px;color:var(--text-sub);margin-top:2px;">系统将生成12位包含大小写、数字及符号的随机密码</div>
          </div>
        </div>
      </label>
      
      <label class="pwd-card" id="pwdCardCustom">
        <div style="display:flex;align-items:center;gap:12px;">
          <input type="radio" name="pwdType" value="custom" style="width:18px;height:18px;accent-color:var(--primary);">
          <div>
            <div style="font-weight:700;font-size:14px;color:var(--text-main);">自定义密码</div>
            <div style="font-size:12px;color:var(--text-sub);margin-top:2px;">手动指定一个新密码</div>
          </div>
        </div>
        <div id="customPwdWrap" style="display:none;margin-top:12px;padding-top:12px;border-top:1px dashed var(--border);">
          <input type="text" id="customPwd" placeholder="输入新密码 (至少8位, 包含大小写数字符号)" style="width:100%;">
        </div>
      </label>
    </div>
    
    <div class="footer"><button class="btn-ghost" onclick="closeModal('modalPwd')">取消</button><button id="confirmPwd">确认重置</button></div>
    <div id="pwdResult" style="display:none;margin-top:16px;padding:12px;background:#f8f9fa;border:1px solid #e5e7eb;border-radius:8px;font-family:monospace;font-size:13px;color:#374151;white-space:pre-wrap;max-height:160px;overflow:auto;"></div>
  </div>
</div>

<style>
.pwd-card { display:block; padding:14px 16px; border:2px solid var(--border); border-radius:12px; cursor:pointer; transition:all .2s ease; background:var(--surface); }
.pwd-card:hover { border-color:#9ca3af; background:#f9fafb; }
.pwd-card.active { border-color:var(--primary); background:#f8fafc; }
</style>

<div class="modal" id="modalLic">
  <div class="dialog" style="max-width:520px;">
    <div class="header"><h3 style="margin:0;">订阅余量</h3><button class="modal-close" onclick="closeModal('modalLic')">✕</button></div>
    <div id="licContent">加载中...</div>
  </div>
</div>

<div class="modal" id="modalDelUserConfirm">
  <div class="dialog no-padding" style="max-width:420px;border:1px solid #dc2626;">
    <div class="header" style="background:#dc2626;margin:0;padding:14px 20px;border-bottom:none;border-radius:12px 12px 0 0;">
      <h3 style="margin:0;color:#fff;">确认删除用户</h3>
      <button class="modal-close" style="background:rgba(255,255,255,0.2);color:#fff;" onclick="closeModal('modalDelUserConfirm')">✕</button>
    </div>
    <div style="padding:20px;color:#374151;line-height:1.8;">即将删除 <strong id="delUserCount" style="color:#dc2626;font-size:18px;">0</strong> 个用户，<strong>此操作不可恢复</strong>，请确认。</div>
    <div class="footer" style="margin:0;padding:16px 20px;background:#f9fafb;border-radius:0 0 12px 12px;">
      <button class="btn-ghost" onclick="closeModal('modalDelUserConfirm')">取消</button>
      <button id="confirmDelUsers" class="btn-danger">确认删除</button>
    </div>
  </div>
</div>

<div class="modal" id="modalAddUser">
  <div class="dialog" style="max-width:520px;">
    <div class="header"><h3 style="margin:0;">新增用户</h3><button class="modal-close" onclick="closeModal('modalAddUser')">✕</button></div>
    <div class="row">
      <span class="label">选择全局</span>
      <div class="custom-select">
        <div class="select-trigger" id="addGlobalTrigger">
          <span id="addGlobalDisplay">加载中...</span>
          <div class="select-arrow"></div>
        </div>
        <div class="options-container" id="addGlobalOptions"></div>
      </div>
      <input type="hidden" id="addGlobal">
    </div>
    <div class="row">
      <span class="label">选择订阅</span>
      <div class="custom-select">
        <div class="select-trigger" id="addSkuTrigger">
          <span id="addSkuDisplay">请先选择全局</span>
          <div class="select-arrow"></div>
        </div>
        <div class="options-container" id="addSkuOptions"></div>
      </div>
      <input type="hidden" id="addSku">
    </div>
    <div class="row" id="addServicePlansRow" style="display:none;">
      <span class="label">禁用应用权限 (可选)</span>
      <div id="addServicePlansWrap" style="max-height:320px;overflow:auto;border:1px solid var(--border);border-radius:var(--radius);padding:8px;background:var(--bg);">
        <!-- 动态加载 Service Plans -->
      </div>
    </div>
    <div class="row"><span class="label">用户名</span><input id="addUsername" placeholder="例如 user123"></div>
    <div class="row"><span class="label">密码</span><input id="addPassword" type="text" placeholder="至少8位，包含大/小写/数字/符号中的3类"></div>
    <div class="row"><label class="inline"><input type="checkbox" id="addForcePwd" checked> 首次登录需修改密码</label></div>
    <div class="footer"><button class="btn-ghost" onclick="closeModal('modalAddUser')">取消</button><button id="confirmAddUser">创建</button></div>
  </div>
</div>

<script>
const adminPath='${adminPath}';
const PLAN_FRIENDLY_NAMES=${JSON.stringify(PLAN_FRIENDLY_NAMES)};
const SUBSCRIPTION_FRIENDLY_NAMES=${JSON.stringify(SUBSCRIPTION_FRIENDLY_NAMES)};
let globalsCache=[]; let usersCache=[]; let sortKey='displayName'; let sortDir=1; let currentPage=1; let pageSize=20; let filterGlobal='ALL'; let searchField='displayName'; let searchText='';

function closeModal(id){document.getElementById(id).style.display='none';}
function openModal(id){document.getElementById(id).style.display='flex';}

function updateArrows(){
  document.querySelectorAll('#userTable th[data-sort]').forEach(th=>{
    const key=th.getAttribute('data-sort');
    th.classList.remove('active');
    const arrow=document.getElementById('arr-'+key);
    if(arrow) arrow.innerText='↕';
    if(key===sortKey){th.classList.add('active'); if(arrow) arrow.innerText=sortDir===1?'↑':'↓';}
  });
}

function renderGlobalsFilter(){
  const wrap=document.getElementById('globalFilters');
  wrap.innerHTML='<span class="label" style="margin:0;">按全局筛选：</span>';
  const allPill=document.createElement('div');
  allPill.className='pill active';
  allPill.innerText='全部';
  allPill.onclick=()=>{
    filterGlobal='ALL';
    document.querySelectorAll('.pill').forEach(p=>p.classList.remove('active'));
    allPill.classList.add('active');
    renderUserRows();
  };
  wrap.appendChild(allPill);
  globalsCache.forEach(g=>{
    const pill=document.createElement('div');
    pill.className='pill';
    pill.innerText=g.label;
    pill.onclick=()=>{
      filterGlobal=g.id;
      document.querySelectorAll('.pill').forEach(p=>p.classList.remove('active'));
      pill.classList.add('active');
      renderUserRows();
    };
    wrap.appendChild(pill);
  });
}

function applyFilterSort(list){
  let data=[...list];
  if(filterGlobal!=='ALL') data=data.filter(u=>u._globalId===filterGlobal);
  if(searchText){
    const txt=searchText.toLowerCase();
    data=data.filter(u=>{
      if(searchField==='displayName') return (u.displayName||'').toLowerCase().includes(txt);
      if(searchField==='userPrincipalName') return (u.userPrincipalName||'').toLowerCase().includes(txt);
      if(searchField==='_globalLabel') return (u._globalLabel||'').toLowerCase().includes(txt);
      if(searchField==='license') return (u._licSort||'').toLowerCase().includes(txt);
      return true;
    });
  }
  data.sort((a,b)=>{
    const va=a[sortKey]||'';
    const vb=b[sortKey]||'';
    if(typeof va==='string') return sortDir*va.localeCompare(vb,'zh-CN');
    return sortDir*((va>vb)-(va<vb));
  });
  return data;
}

function renderUserRows(){
  updateArrows();
  document.getElementById('chkAll').checked=false;
  const body=document.getElementById('userBody');
  const data=applyFilterSort(usersCache);
  const total=data.length;
  const totalPages=Math.max(1,Math.ceil(total/pageSize));
  currentPage=Math.min(currentPage,totalPages);
  const start=(currentPage-1)*pageSize;
  const pageData=data.slice(start,start+pageSize);

  if(!pageData.length){
    body.innerHTML='<tr><td colspan="7" style="text-align:center;">暂无数据</td></tr>';
  } else {
    body.innerHTML=pageData.map(u=>{
      const lic=(u.assignedLicenses||[]).map(l=>{
        const nm = l.name || l.skuId;
        const disp = SUBSCRIPTION_FRIENDLY_NAMES[nm] || nm;
        return '<span class="tag">'+disp+'</span>';
      }).join('') || '<span style="color:#9ca3af;">无</span>';
      return '<tr><td data-label="选择"><input type="checkbox" class="chk" data-g="'+u._globalId+'" value="'+u.id+'"></td><td data-label="用户名"><strong>'+ (u.displayName||'') +'</strong></td><td data-label="账号">'+u.userPrincipalName+'</td><td data-label="订阅">'+lic+'</td><td data-label="创建时间">'+new Date(u.createdDateTime).toLocaleString()+'</td><td data-label="全局">'+u._globalLabel+'</td><td data-label="UUID" style="font-size:11px;color:#9ca3af;">'+u.id+'</td></tr>';
    }).join('');
  }
  document.getElementById('pageInfo').innerText='第 '+currentPage+' / '+totalPages+' 页 · 共 '+total+' 条';
}

function getSelected(){return Array.from(document.querySelectorAll('.chk:checked')).map(c=>({id:c.value,g:c.getAttribute('data-g')}));}
function generatePass(){const chars='abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*'; let p=''; for(let i=0;i<12;i++) p+=chars[Math.floor(Math.random()*chars.length)]; return p+'Aa1!';}

function setupCustomSelect(triggerId, optionsId, inputId, displayId, onChange) {
  const trigger = document.getElementById(triggerId);
  const container = document.getElementById(optionsId);
  const wrapper = trigger.parentElement;
  
  const optCount = container.querySelectorAll('.option').length;
  if (optCount <= 1) {
    trigger.classList.add('disabled');
    trigger.style.cursor = 'default';
    trigger.onclick = (e) => e.stopPropagation();
    return;
  } else {
    trigger.classList.remove('disabled');
    trigger.style.cursor = 'pointer';
  }
  
  trigger.onclick = (e) => { e.stopPropagation(); container.classList.toggle('open'); wrapper.classList.toggle('open'); };
  document.addEventListener('click', (e) => { if(!wrapper.contains(e.target)) {container.classList.remove('open'); wrapper.classList.remove('open');} });
  if(wrapper) {
    let leaveTimer;
    wrapper.addEventListener('mouseleave', () => { leaveTimer = setTimeout(() => { container.classList.remove('open'); wrapper.classList.remove('open'); }, 200); });
    wrapper.addEventListener('mouseenter', () => clearTimeout(leaveTimer));
  }
  container.addEventListener('click', (e) => {
    const opt = e.target.closest('.option');
    if (!opt) return;
    document.getElementById(displayId).innerText = opt.innerText;
    document.getElementById(inputId).value = opt.getAttribute('data-value');
    container.classList.remove('open');
    wrapper.classList.remove('open');
    document.querySelectorAll('#' + optionsId + ' .option').forEach(o => o.classList.remove('selected'));
    opt.classList.add('selected');
    if(onChange) onChange();
  });
}
setupCustomSelect('addGlobalTrigger', 'addGlobalOptions', 'addGlobal', 'addGlobalDisplay', refreshSkuOptions);
setupCustomSelect('addSkuTrigger', 'addSkuOptions', 'addSku', 'addSkuDisplay', refreshAddServicePlans);
setupCustomSelect('pageSizeTrigger', 'pageSizeOptions', null, 'pageSizeDisplay', () => {
  pageSize = parseInt(document.getElementById('pageSizeDisplay').innerText) || 20;
  currentPage = 1;
  renderUserRows();
});
setupCustomSelect('searchFieldTrigger', 'searchFieldOptions', 'searchField', 'searchFieldDisplay');

function fillAddUserModal(){
  const container = document.getElementById('addGlobalOptions');
  container.innerHTML = globalsCache.map(g => '<div class="option" data-value="'+g.id+'">'+g.label+'</div>').join('');
  
  if(globalsCache.length > 0) {
    // 优先选中下拉选项中的第一项
    document.querySelectorAll('#addGlobalOptions .option').forEach(o => o.classList.remove('selected'));
    const firstOption = container.firstChild;
    if (firstOption) {
      firstOption.classList.add('selected');
      document.getElementById('addGlobal').value = firstOption.getAttribute('data-value');
      document.getElementById('addGlobalDisplay').innerText = firstOption.innerText;
    }
  } else {
    document.getElementById('addGlobal').value = '';
    document.getElementById('addGlobalDisplay').innerText = '无可用全局';
  }
  refreshSkuOptions();
}

function refreshSkuOptions(){
  const globalId = document.getElementById('addGlobal').value;
  const g = globalsCache.find(x => x.id === globalId);
  const skuMap = g?.skuMap || {};
  const container = document.getElementById('addSkuOptions');
  const keys = Object.keys(skuMap);
  container.innerHTML = keys.map(k => '<div class="option" data-value="'+k+'">'+(SUBSCRIPTION_FRIENDLY_NAMES[k]||k)+'</div>').join('');
  
  if(keys.length > 0) {
    document.querySelectorAll('#addSkuOptions .option').forEach(o => o.classList.remove('selected'));
    const firstOption = container.firstChild;
    if (firstOption) {
      firstOption.classList.add('selected');
      document.getElementById('addSku').value = firstOption.getAttribute('data-value');
      document.getElementById('addSkuDisplay').innerText = firstOption.innerText;
    }
  } else {
    document.getElementById('addSku').value = '';
    document.getElementById('addSkuDisplay').innerText = '暂无可用订阅';
  }
  refreshAddServicePlans();
}

async function fetchGlobals(){const res=await fetch(adminPath+'/api/globals'); const data=await res.json(); globalsCache=data; renderGlobalsFilter(); fillAddUserModal();}
async function fetchUsers(){document.getElementById('globalLoader').style.display='flex'; const res=await fetch(adminPath+'/api/users'); const data=await res.json(); usersCache=data; renderUserRows(); document.getElementById('globalLoader').style.display='none';}

document.getElementById('btnAddUser').onclick=()=>{fillAddUserModal(); document.getElementById('addPassword').value=generatePass(); openModal('modalAddUser');};

async function refreshAddServicePlans() {
  const globalId = document.getElementById('addGlobal').value;
  const skuName = document.getElementById('addSku').value;
  const wrap = document.getElementById('addServicePlansWrap');
  const row = document.getElementById('addServicePlansRow');
  
  if (!globalId || !skuName) {
    row.style.display = 'none';
    return;
  }
  
  const g = globalsCache.find(x => x.id === globalId);
  const skuId = (g?.skuMap || {})[skuName];
  if (!skuId) return;

  row.style.display = 'block';
  wrap.innerHTML = '<div style="color:var(--text-muted);font-size:12px;text-align:center;padding:10px;">正在拉取可用应用...</div>';
  
  try {
    const res = await fetch(adminPath + '/api/service_plans', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ globalId, skuId })
    });
    const data = await res.json();
    if (data.success && data.plans && data.plans.length > 0) {
      let html = '<div style="display:flex;flex-direction:column;gap:6px;">';
      data.plans.forEach(p => {
        html += '<label style="display:flex;align-items:center;gap:6px;font-size:13px;cursor:pointer;">';
        html += '<input type="checkbox" class="addDisabledPlanChk" value="' + p.servicePlanId + '"> <span style="color:var(--text-main);">禁用 ' + (PLAN_FRIENDLY_NAMES[p.servicePlanName] || p.servicePlanName) + '</span>';
        html += '</label>';
      });
      html += '</div>';
      wrap.innerHTML = html;
    } else {
      wrap.innerHTML = '<div style="color:var(--text-muted);font-size:12px;text-align:center;padding:10px;">该订阅未匹配到可禁用的核心应用</div>';
    }
  } catch (e) {
    wrap.innerHTML = '<div style="color:#dc2626;font-size:12px;text-align:center;padding:10px;">拉取应用失败</div>';
  }
}


document.getElementById('confirmAddUser').onclick=async()=>{
  const disabledPlans = Array.from(document.querySelectorAll('.addDisabledPlanChk:checked')).map(c => c.value);
  const payload={
    globalId:document.getElementById('addGlobal').value, 
    skuName:document.getElementById('addSku').value, 
    username:(document.getElementById('addUsername').value||'').trim(), 
    password:document.getElementById('addPassword').value||'', 
    forceChangePasswordNextSignIn:document.getElementById('addForcePwd').checked,
    disabledPlans
  };
  if(!payload.globalId||!payload.skuName||!payload.username||!payload.password){showToast('请填写完整信息','error'); return;}
  const res=await fetch(adminPath+'/api/users/create',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
  const data=await res.json();
  if(data.success){closeModal('modalAddUser'); showToast('用户创建成功：'+data.email,'success',3200); await fetchUsers();}
  else showToast(data.message||'创建失败','error',3200);
};

document.getElementById('btnRefresh').onclick=fetchUsers;
document.getElementById('btnPwd').onclick=()=>{if(getSelected().length===0){showToast('请选择用户','error'); return;} openModal('modalPwd');};
document.getElementById('btnLic').onclick=async()=>{
  openModal('modalLic');
  document.getElementById('licContent').innerText='查询中...';
  const res=await fetch(adminPath+'/api/licenses');
  const data=await res.json();
  document.getElementById('licContent').innerHTML=data.map(i=>{
    const remain=i.total-i.used;
    const pct=i.total?Math.round(i.used/i.total*100):0;
    const exp=i.expiresAt?new Date(i.expiresAt).toLocaleString():'-';
    const friendlyName = SUBSCRIPTION_FRIENDLY_NAMES[i.skuPartNumber] || i.skuPartNumber;
    return '<div style="margin:8px 0;"><strong>'+i.globalLabel+' / '+friendlyName+'</strong><div style="color:#6b7280;font-size:12px;margin-top:4px;line-height:1.6;">总量 '+i.total+'，已用 '+i.used+'，剩余 '+remain+'，使用率 '+pct+'%</div><div style="color:#6b7280;font-size:12px;margin-top:2px;line-height:1.6;">订阅到期时间：'+exp+'</div><div style="margin-top:6px;height:6px;background:#e5e7eb;border-radius:8px;overflow:hidden;"><div style="width:'+pct+'%;height:100%;background:var(--primary);"></div></div></div>';
  }).join('') || '暂无数据';
};
document.getElementById('btnDel').onclick=()=>{
  const sel=getSelected();
  if(!sel.length){showToast('请选择用户','error'); return;}
  document.getElementById('delUserCount').innerText=sel.length;
  openModal('modalDelUserConfirm');
};

document.getElementById('confirmDelUsers').onclick=async()=>{
  closeModal('modalDelUserConfirm');
  const sel=getSelected();
  if(!sel.length) return;
  const btn = document.getElementById('confirmDelUsers');
  btn.disabled = true;
  document.getElementById('status').innerText='删除中...';
  await Promise.all(sel.map(item=>fetch(adminPath+'/api/users/'+item.g+'/'+item.id,{method:'DELETE'})));
  btn.disabled = false;
  showToast('删除完成','success');
  document.getElementById('status').innerText='';
  // 如果删除后当前页为空且不是第一页，自动退回上一页
  const data=applyFilterSort(usersCache);
  const remainingInPage = data.length - sel.length;
  if(remainingInPage <= (currentPage-1)*pageSize && currentPage > 1) {
    currentPage--;
  }
  await fetchUsers();
};
document.getElementById('chkAll').onchange=(e)=>{document.querySelectorAll('.chk').forEach(c=>c.checked=e.target.checked);};
document.querySelectorAll('input[name="pwdType"]').forEach(r=>{
  r.onchange=()=>{
    const isCustom = r.value === 'custom';
    document.getElementById('customPwdWrap').style.display = isCustom ? 'block' : 'none';
    document.getElementById('pwdCardCustom').className = isCustom ? 'pwd-card active' : 'pwd-card';
    document.getElementById('pwdCardAuto').className = isCustom ? 'pwd-card' : 'pwd-card active';
  };
});

document.getElementById('confirmPwd').onclick=async()=>{
  const sel=getSelected();
  if(!sel.length){showToast('请选择用户','error'); return;}
  const type=document.querySelector('input[name="pwdType"]:checked').value;
  let pwd='';
  if(type==='custom'){
    pwd=document.getElementById('customPwd').value; 
    if(!pwd){showToast('请输入新密码','error'); return;}
  }
  
  const btn = document.getElementById('confirmPwd');
  const resBox = document.getElementById('pwdResult');
  btn.disabled = true;
  btn.innerText = '正在重置...';
  resBox.style.display = 'block';
  resBox.innerText = '开始重置密码...\\n';
  
  const result=[];
  let successCount = 0;
  for(const s of sel){
    const user = usersCache.find(u => u.id === s.id);
    const finalPwd=type==='auto'?generatePass():pwd;
    try {
      const res = await fetch(adminPath+'/api/users/'+s.g+'/'+s.id+'/password',{method:'PATCH',headers:{'Content-Type':'application/json'},body:JSON.stringify({password:finalPwd})});
      const data = await res.json();
      if(data.success) {
        result.push('✅ ' + (user?.userPrincipalName || s.id) + ' => ' + finalPwd);
        successCount++;
      } else {
        result.push('❌ ' + (user?.userPrincipalName || s.id) + ' => 失败: ' + (data.message || '未知错误'));
      }
    } catch(e) {
      result.push('❌ ' + (user?.userPrincipalName || s.id) + ' => 请求异常');
    }
    resBox.innerText = '进度: ' + result.length + ' / ' + sel.length + '\\n\\n' + result.join('\\n');
  }
  btn.disabled = false;
  btn.innerText = '确认重置';
  showToast('重置完成，成功 ' + successCount + ' 个', 'success');
};

document.querySelectorAll('#userTable th[data-sort]').forEach(th=>{th.onclick=()=>{const key=th.getAttribute('data-sort'); if(sortKey===key) sortDir*=-1; else {sortKey=key; sortDir=1;} renderUserRows();};});
updateArrows();
document.getElementById('prevPage').onclick=()=>{if(currentPage>1){currentPage--; renderUserRows();}};
document.getElementById('nextPage').onclick=()=>{const data=applyFilterSort(usersCache); const totalPages=Math.max(1,Math.ceil(data.length/pageSize)); if(currentPage<totalPages){currentPage++; renderUserRows();}};
document.getElementById('goPage').onclick=()=>{const val=parseInt(document.getElementById('jumpPage').value)||1; const data=applyFilterSort(usersCache); const totalPages=Math.max(1,Math.ceil(data.length/pageSize)); currentPage=Math.min(Math.max(1,val),totalPages); renderUserRows();};
document.getElementById('btnSearch').onclick=()=>{searchField=document.getElementById('searchField').value; searchText=document.getElementById('searchText').value.trim(); currentPage=1; renderUserRows();};
document.getElementById('btnClear').onclick=()=>{document.getElementById('searchText').value=''; searchText=''; currentPage=1; renderUserRows();};
(async()=>{await fetchGlobals(); await fetchUsers();})();
</script>
`,
  });
}

function renderGlobalsPage(adminPath) {
  const ICON_EDIT = `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" width="14" height="14"><path d="M4 13.5V16h2.5L15 7.5l-2.5-2.5L4 13.5z"/><path d="M12.5 5l2.5 2.5"/></svg>`;
  const ICON_TRASH = `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" width="14" height="14"><path d="M4 6h12"/><path d="M7 6v-2h6v2"/><path d="M8 9v5"/><path d="M12 9v5"/><path d="M6 6l1 10h6l1-10"/></svg>`;
  const ICON_PLUS = `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" width="15" height="15"><path d="M10 4v12M4 10h12"/></svg>`;
  const ICON_DOWNLOAD = `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" width="15" height="15"><path d="M10 3v9"/><path d="M6.5 9.5L10 13l3.5-3.5"/><path d="M4 16h12"/></svg>`;
  return adminLayout({
    title: '全局账户',
    adminPath,
    active: 'globals',
    content: `
<style>
.g-controls{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:16px;}
.g-sort-wrap{display:flex;align-items:center;gap:6px;margin-left:auto;}
.g-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:16px;}
.g-card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:18px 20px;transition:box-shadow .18s,transform .18s;cursor:default;}
.g-card:hover{box-shadow:var(--shadow-md);transform:translateY(-2px);}
.g-card-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;gap:8px;}
.g-card-name{font-weight:700;font-size:15px;color:var(--text-main);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex:1;}
.g-card-actions{display:flex;gap:6px;flex-shrink:0;}
.g-card-actions button{padding:6px 10px;font-size:12px;gap:4px;}
.g-card-actions button svg{width:13px;height:13px;}
.g-meta{display:flex;flex-direction:column;gap:8px;}
.g-row{display:flex;flex-direction:column;gap:2px;}
.g-row-label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--text-muted);}
.g-row-val{font-size:13px;color:var(--text-main);word-break:break-all;line-height:1.5;}
.g-sku-badge{display:inline-flex;align-items:center;gap:4px;padding:3px 8px;background:#f0f4ff;border:1px solid #c7d7fe;border-radius:6px;font-size:12px;color:#3730a3;font-weight:600;}
.g-empty{text-align:center;padding:48px 20px;color:var(--text-muted);font-size:14px;}
@media(max-width:600px){.g-grid{grid-template-columns:1fr;} .g-sort-wrap{margin-left:0;width:100%;}}
</style>

<div class="section">
  <div class="g-controls">
    <button id="btnAdd">${ICON_PLUS} 新增全局</button>
    <input id="gSearch" class="input-compact" placeholder="搜索名称 / 域 / 租户 ID" style="max-width:240px;">
    <div class="g-sort-wrap">
      <span class="label" style="margin:0;white-space:nowrap;">排序：</span>
      <select id="gSortSel" style="width:auto;padding:6px 10px;font-size:12px;">
        <option value="label">名称</option>
        <option value="defaultDomain">域</option>
        <option value="tenantId">租户 ID</option>
        <option value="skuCount">SKU 数</option>
      </select>
      <button id="gSortDir" class="btn-ghost" style="padding:6px 10px;font-size:12px;" title="切换升/降序">↑ 升序</button>
    </div>
  </div>
  <div id="gGrid" class="g-grid"></div>
</div>

<div class="modal" id="modalG">
  <div class="dialog" style="max-width:720px;">
    <div class="header"><h3 id="gTitle" style="margin:0;">新增全局</h3><button class="modal-close" onclick="closeModal('modalG')">✕</button></div>
    <div class="row"><span class="label">展示名称（用户可见）</span><input id="gLabel"></div>
    <div class="row"><span class="label">默认邮箱后缀 (不含 @)</span><input id="gDomain"></div>
    <div class="row"><span class="label">租户 ID</span><input id="gTenant"></div>
    <div class="row"><span class="label">客户端 ID</span><input id="gClientId"></div>
    <div class="row"><span class="label">客户端密钥</span><input id="gSecret" type="password"></div>
    <div class="row"><span class="label">SKU JSON (键为展示名, 值为 SKU ID)</span><textarea id="gSku" rows="4" placeholder='例如 {"E5开发版":"xxx","A1教育":"yyy"}'></textarea></div>
    <div class="toolbar" style="margin-top:6px;">
      <button id="btnFetchSku" class="btn-ghost" disabled>${ICON_DOWNLOAD} 点我获取 SKU</button>
      <span style="color:#6b7280;font-size:12px;line-height:1.4;">填入租户ID / 客户端ID / 客户端密钥后即可获取</span>
    </div>
    <div class="footer"><button class="btn-ghost" onclick="closeModal('modalG')">取消</button><button id="btnSaveG">保存</button></div>
  </div>
</div>

<script>
const adminPath='${adminPath}';
const SUBSCRIPTION_FRIENDLY_NAMES=${JSON.stringify(SUBSCRIPTION_FRIENDLY_NAMES)};
const ICON_EDIT_STR=${JSON.stringify(ICON_EDIT)};
const ICON_TRASH_STR=${JSON.stringify(ICON_TRASH)};
let editingId=null; let gSortKey='label', gSortDir=1; let gSearchText=''; let globalsData=[];

function closeModal(id){document.getElementById(id).style.display='none';}
function openModal(id){document.getElementById(id).style.display='flex';}

function renderGlobals(){
  let list=[...globalsData];
  if(gSearchText){
    const t=gSearchText.toLowerCase();
    list=list.filter(x=>(x.label||'').toLowerCase().includes(t)||(x.defaultDomain||'').toLowerCase().includes(t)||(x.tenantId||'').toLowerCase().includes(t));
  }
  list.sort((a,b)=>{
    const va=a[gSortKey]||''; const vb=b[gSortKey]||'';
    if(typeof va==='string') return gSortDir*va.localeCompare(vb,'zh-CN');
    return gSortDir*((va>vb)-(va<vb));
  });
  const grid=document.getElementById('gGrid');
  if(!list.length){
    grid.innerHTML='<div class="g-empty">暂无全局账户，点击「新增全局」开始配置</div>';
    return;
  }
  grid.innerHTML=list.map(g=>{
    const skuNames=Object.keys(g.skuMap||{});
    const skuHtml=skuNames.length
      ? skuNames.map(s=>'<span class="g-sku-badge">'+(SUBSCRIPTION_FRIENDLY_NAMES[s]||s)+'</span>').join(' ')
      : '<span style="color:var(--text-muted);font-size:12px;">未配置订阅</span>';
    const shortTenant=(g.tenantId||'').length>20?(g.tenantId.slice(0,10)+'…'+g.tenantId.slice(-8)):g.tenantId;
    return '<div class="g-card">'
      +'<div class="g-card-head">'
        +'<div class="g-card-name" title="'+g.label+'">'+g.label+'</div>'
        +'<div class="g-card-actions">'
          +'<button class="btn-ghost" onclick="editG(\\''+g.id+'\\')">'+ICON_EDIT_STR+' 编辑</button>'
          +'<button class="btn-danger" onclick="delG(\\''+g.id+'\\')">'+ICON_TRASH_STR+'</button>'
        +'</div>'
      +'</div>'
      +'<div class="g-meta">'
        +'<div class="g-row"><span class="g-row-label">邮箱域</span><span class="g-row-val">@'+g.defaultDomain+'</span></div>'
        +'<div class="g-row"><span class="g-row-label">租户 ID</span><span class="g-row-val" title="'+g.tenantId+'">'+shortTenant+'</span></div>'
        +'<div class="g-row"><span class="g-row-label">订阅 ('+skuNames.length+')</span><div style="margin-top:4px;display:flex;flex-wrap:wrap;gap:4px;">'+skuHtml+'</div></div>'
      +'</div>'
    +'</div>';
  }).join('');
}

async function loadGlobals(){
  const res=await fetch(adminPath+'/api/globals');
  const data=await res.json();
  globalsData=data.map(g=>({...g, skuCount:Object.keys(g.skuMap||{}).length}));
  renderGlobals();
}

document.getElementById('btnAdd').onclick=()=>{
  editingId=null;
  ['gLabel','gDomain','gTenant','gClientId','gSecret','gSku'].forEach(id=>{const el=document.getElementById(id); if(el) el.value='';});
  document.getElementById('gTitle').innerText='新增全局';
  refreshFetchBtn();
  openModal('modalG');
};

window.editG=async(id)=>{
  const res=await fetch(adminPath+'/api/globals/'+id);
  const g=await res.json();
  editingId=id;
  document.getElementById('gTitle').innerText='编辑全局';
  document.getElementById('gLabel').value=g.label||'';
  document.getElementById('gDomain').value=g.defaultDomain||'';
  document.getElementById('gTenant').value=g.tenantId||'';
  document.getElementById('gClientId').value=g.clientId||'';
  document.getElementById('gSecret').value=g.clientSecret||'';
  document.getElementById('gSku').value=JSON.stringify(g.skuMap||{},null,2);
  refreshFetchBtn();
  openModal('modalG');
};

window.delG=async(id)=>{
  const g=globalsData.find(x=>x.id===id);
  if(!confirm('确认删除全局「'+(g?g.label:id)+'」？')) return;
  await fetch(adminPath+'/api/globals/'+id,{method:'DELETE'});
  showToast('已删除','success');
  loadGlobals();
};

document.getElementById('btnSaveG').onclick=async()=>{
  const payload={
    label:document.getElementById('gLabel').value.trim(),
    defaultDomain:document.getElementById('gDomain').value.trim(),
    tenantId:document.getElementById('gTenant').value.trim(),
    clientId:document.getElementById('gClientId').value.trim(),
    clientSecret:document.getElementById('gSecret').value.trim(),
    skuMap:document.getElementById('gSku').value
  };
  const method=editingId?'PATCH':'POST';
  const url=adminPath+'/api/globals'+(editingId?'/'+editingId:'');
  const res=await fetch(url,{method,headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
  const d=await res.json();
  if(d.success){closeModal('modalG'); loadGlobals(); showToast('保存成功','success');}
  else showToast(d.message||'保存失败','error');
};

function canFetchSku(){
  const t=(document.getElementById('gTenant').value||'').trim();
  const c=(document.getElementById('gClientId').value||'').trim();
  const s=(document.getElementById('gSecret').value||'').trim();
  return !!(t&&c&&s);
}
function refreshFetchBtn(){document.getElementById('btnFetchSku').disabled=!canFetchSku();}
['gTenant','gClientId','gSecret'].forEach(id=>{const el=document.getElementById(id); if(el) el.addEventListener('input',refreshFetchBtn);});

document.getElementById('btnFetchSku').onclick=async()=>{
  if(!canFetchSku()){showToast('请先填写租户ID、客户端ID、客户端密钥','error'); return;}
  const payload={tenantId:(document.getElementById('gTenant').value||'').trim(),clientId:(document.getElementById('gClientId').value||'').trim(),clientSecret:(document.getElementById('gSecret').value||'').trim()};
  const res=await fetch(adminPath+'/api/fetch_skus',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
  const data=await res.json();
  if(data.success){document.getElementById('gSku').value=JSON.stringify(data.map||{},null,2); showToast('SKU 获取成功','success');}
  else showToast(data.message||'获取失败','error');
};

// 搜索与排序
document.getElementById('gSearch').addEventListener('input',(e)=>{gSearchText=e.target.value.trim(); renderGlobals();});
document.getElementById('gSortSel').onchange=(e)=>{gSortKey=e.target.value; renderGlobals();};
document.getElementById('gSortDir').onclick=()=>{
  gSortDir*=-1;
  document.getElementById('gSortDir').innerText=gSortDir===1?'↑ 升序':'↓ 降序';
  renderGlobals();
};

loadGlobals();
</script>
`,
  });
}

function renderInvitesPage(adminPath, globals) {
  return adminLayout({
    title: '邀请码管理',
    adminPath,
    active: 'invites',
    content: `
<div class="section">
  <div class="toolbar">
    <button id="btnRefreshInvites" class="btn-ghost">${ICONS.refresh} 刷新</button>
    <button id="btnGen">${ICONS.spark} 生成邀请码</button>
    <button id="btnExport" class="btn-ghost">${ICONS.download} 导出所选</button>
    <button id="btnDelInvites" class="btn-danger">${ICONS.trash} 删除所选</button>
  </div>
  <div class="toolbar search-box" style="flex-wrap:wrap;gap:8px;">
    <span class="label" style="margin:0;white-space:nowrap;">搜索：</span>
    <div class="custom-select" style="min-width:100px;max-width:120px;">
      <div class="select-trigger" id="iSearchFieldTrigger" style="padding:6px 10px;font-size:12px;">
        <span id="iSearchFieldDisplay">全部内容</span>
        <div class="select-arrow"></div>
      </div>
      <div class="options-container" id="iSearchFieldOptions" style="min-width:100px;">
        <div class="option selected" data-value="all">全部内容</div>
        <div class="option" data-value="code">邀请码</div>
        <div class="option" data-value="limit">限制次数</div>
        <div class="option" data-value="used">已用</div>
        <div class="option" data-value="status">状态</div>
        <div class="option" data-value="scope">限制范围</div>
      </div>
    </div>
    <input type="hidden" id="iSearchField" value="all">
    <input id="iSearchText" class="input-compact" placeholder="输入关键词，支持模糊" style="max-width:180px;">
    <button id="iSearchBtn" class="btn-ghost">搜索</button>
    <button id="iClearBtn" class="btn-ghost">清空</button>
  </div>
  <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin:8px 0;">
    <div style="display:flex;align-items:center;gap:6px;">
      <span style="font-size:12px;color:var(--text-sub);">每页</span>
      <div class="custom-select" style="min-width:70px;">
        <div class="select-trigger" id="pageSizeInviteTrigger" style="padding:6px 10px;font-size:12px;">
          <span id="pageSizeInviteDisplay">20</span>
          <div class="select-arrow"></div>
        </div>
        <div class="options-container" id="pageSizeInviteOptions" style="min-width:70px;">
          <div class="option selected" data-value="20">20</div>
          <div class="option" data-value="30">30</div>
          <div class="option" data-value="50">50</div>
          <div class="option" data-value="100">100</div>
        </div>
      </div>
      <span style="font-size:12px;color:var(--text-sub);">条</span>
    </div>
    <div style="display:flex;align-items:center;gap:6px;">
      <span id="pageInfoInvite" style="font-size:12px;color:var(--text-sub);white-space:nowrap;"></span>
      <button id="prevInvite" class="btn-ghost" style="padding:5px 10px;font-size:12px;">上一页</button>
      <button id="nextInvite" class="btn-ghost" style="padding:5px 10px;font-size:12px;">下一页</button>
      <input class="page-input" id="jumpInvite" type="number" min="1" placeholder="页码" style="width:64px;padding:5px 8px;font-size:12px;">
      <button id="goInvite" class="btn-ghost" style="padding:5px 10px;font-size:12px;">跳转</button>
    </div>
  </div>
  <div class="table-wrap">
    <table class="table" style="table-layout:fixed;">
      <thead><tr>
        <th style="width:40px;text-align:center;padding-left:14px;"><input type="checkbox" id="chkInviteAll" style="vertical-align:middle;"></th>
        <th style="width:15%;white-space:nowrap;" data-sort="code">邀请码 <span class="arrow" id="iarr-code">↕</span></th>
        <th style="width:8%;text-align:center;white-space:nowrap;" data-sort="limit">限制次数 <span class="arrow" id="iarr-limit">↕</span></th>
        <th style="width:7%;text-align:center;white-space:nowrap;" data-sort="used">已用 <span class="arrow" id="iarr-used">↕</span></th>
        <th style="width:8%;text-align:center;white-space:nowrap;" data-sort="status">状态 <span class="arrow" id="iarr-status">↕</span></th>
        <th style="width:40%;white-space:nowrap;" data-sort="scope">限制范围 <span class="arrow" id="iarr-scope">↕</span></th>
        <th style="width:11%;white-space:nowrap;" data-sort="createdAt">生成时间 <span class="arrow" id="iarr-createdAt">↕</span></th>
        <th style="width:11%;white-space:nowrap;" data-sort="usedAt">最近使用 <span class="arrow" id="iarr-usedAt">↕</span></th>
      </tr></thead>
      <tbody id="inviteBody"></tbody>
    </table>
  </div>
</div>

<div class="modal" id="modalGen">
  <div class="dialog" style="max-width:620px;">
    <div class="header"><h3 style="margin:0;">生成邀请码</h3><button class="modal-close" onclick="closeModal('modalGen')">✕</button></div>
    <div class="row"><span class="label">生成数量</span><input id="cQty" type="number" value="10" min="1"></div>
    <div class="row"><span class="label">每个邀请码可使用次数</span><input id="cLimit" type="number" value="1" min="1"></div>
    <div class="row"><span class="label">限制注册范围 (至少选一项)</span><div id="scopeWrap" style="max-height:280px;overflow:auto;border:1px solid #e5e7eb;border-radius:12px;padding:12px;background:#f8f9ff;"></div></div>
    <div class="row" id="genServicePlansRow" style="display:none;margin-top:14px;">
      <span class="label">禁用应用权限 (仅单选订阅生效)</span>
      <div id="genServicePlansWrap" style="max-height:320px;overflow:auto;border:1px solid #e5e7eb;border-radius:12px;padding:12px;background:#f8f9ff;">
      </div>
    </div>
    <div class="footer"><button class="btn-ghost" onclick="closeModal('modalGen')">取消生成</button><button id="doGen">确定生成</button></div>
  </div>
</div>

<div class="modal" id="modalDelInviteConfirm">
  <div class="dialog no-padding" style="max-width:420px;border:1px solid #dc2626;">
    <div class="header" style="background:#dc2626;margin:0;padding:14px 20px;border-bottom:none;border-radius:12px 12px 0 0;">
      <h3 style="margin:0;color:#fff;">确认删除邀请码</h3>
      <button class="modal-close" style="background:rgba(255,255,255,0.2);color:#fff;" onclick="closeModal('modalDelInviteConfirm')">✕</button>
    </div>
    <div style="padding:20px;color:#374151;line-height:1.8;">即将删除 <strong id="delInviteCount" style="color:#dc2626;font-size:18px;">0</strong> 条邀请码，<strong>此操作不可恢复</strong>，请确认。</div>
    <div class="footer" style="margin:0;padding:16px 20px;background:#f9fafb;border-radius:0 0 12px 12px;">
      <button class="btn-ghost" onclick="closeModal('modalDelInviteConfirm')">取消</button>
      <button id="confirmDelInvites" class="btn-danger">确认删除</button>
    </div>
  </div>
</div>

<script>
const adminPath='${adminPath}';
const globalsList=${JSON.stringify(globals)};
const SUBSCRIPTION_FRIENDLY_NAMES=${JSON.stringify(SUBSCRIPTION_FRIENDLY_NAMES)};
const PLAN_FRIENDLY_NAMES=${JSON.stringify(PLAN_FRIENDLY_NAMES)};
function closeModal(id){document.getElementById(id).style.display='none';}
function openModal(id){document.getElementById(id).style.display='flex';}
let invitesCache=[]; let sortKey='code'; let sortDir=1; let invitePage=1; let invitePageSize=20; let iSearchField='code'; let iSearchText=''; let pendingDelCodes=[];

function updateIArrows(){
  ['code','limit','used','status','scope','createdAt','usedAt'].forEach(k=>{
    const th=document.querySelector('th[data-sort="'+k+'"]');
    const arr=document.getElementById('iarr-'+k);
    if(th){ th.classList.remove('active'); if(arr) arr.innerText='↕'; }
    if(k===sortKey){ if(th) th.classList.add('active'); if(arr) arr.innerText=sortDir===1?'↑':'↓'; }
  });
}

async function refreshGenServicePlans() {
  const chks = Array.from(document.querySelectorAll('.scopeChk:checked'));
  const row = document.getElementById('genServicePlansRow');
  const wrap = document.getElementById('genServicePlansWrap');
  
  if (chks.length !== 1) {
    row.style.display = 'none';
    return;
  }
  
  const c = chks[0];
  const globalId = c.getAttribute('data-g');
  const skuName = c.getAttribute('data-sku');
  const g = globalsList.find(x => x.id === globalId);
  const skuId = (g?.skuMap || {})[skuName];
  
  if (!skuId) return;

  row.style.display = 'block';
  wrap.innerHTML = '<div style="color:var(--text-muted);font-size:12px;text-align:center;padding:10px;">正在拉取可用应用...</div>';
  
  try {
    const res = await fetch(adminPath + '/api/service_plans', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ globalId, skuId })
    });
    const data = await res.json();
    if (data.success && data.plans && data.plans.length > 0) {
      let html = '<div style="display:flex;flex-direction:column;gap:6px;">';
      data.plans.forEach(p => {
        html += '<label style="display:flex;align-items:center;gap:6px;font-size:13px;cursor:pointer;">';
        html += '<input type="checkbox" class="genDisabledPlanChk" value="' + p.servicePlanId + '" data-name="' + p.servicePlanName + '"> <span style="color:var(--text-main);">禁用 ' + (PLAN_FRIENDLY_NAMES[p.servicePlanName] || p.servicePlanName) + '</span>';
        html += '</label>';
      });
      html += '</div>';
      wrap.innerHTML = html;
    } else {
      wrap.innerHTML = '<div style="color:var(--text-muted);font-size:12px;text-align:center;padding:10px;">该订阅未匹配到可禁用的核心应用</div>';
    }
  } catch (e) {
    wrap.innerHTML = '<div style="color:#dc2626;font-size:12px;text-align:center;padding:10px;">拉取应用失败</div>';
  }
}

function buildScopeOptions(){
  const wrap=document.getElementById('scopeWrap');
  const validGlobals=globalsList.filter(g=>Object.keys(g.skuMap||{}).length>0);
  const autoCheck=validGlobals.length===1;
  wrap.innerHTML=validGlobals.map(g=>{
    const sku=Object.keys(g.skuMap||{});
    return sku.map(s=>{
      const chk=autoCheck?'checked':'';
      return '<label style="display:flex;align-items:center;gap:10px;padding:7px 10px;border-radius:8px;cursor:pointer;background:rgba(79,70,229,0.04);margin-bottom:4px;">'
        +'<input type="checkbox" class="scopeChk" data-g="'+g.id+'" data-sku="'+s+'" '+chk+' style="width:15px;height:15px;flex:0 0 auto;">'
        +'<span style="font-size:13px;color:#374151;display:flex;align-items:center;gap:6px;">'
        +'<strong style="color:#4338ca;">🌐 '+g.label+'</strong>'
        +'<span style="color:#9ca3af;">/</span>'
        +'<span>'+(SUBSCRIPTION_FRIENDLY_NAMES[s]||s)+'</span>'
        +'</span>'
        +'</label>';
    }).join('');
  }).join('') || '<div style="color:#9ca3af;font-size:13px;padding:8px;">暂无全局/订阅</div>';
  
  document.querySelectorAll('.scopeChk').forEach(c => {
    c.addEventListener('change', refreshGenServicePlans);
  });
  if (autoCheck) refreshGenServicePlans();
}

document.getElementById('btnGen').onclick=()=>{
  document.getElementById('cQty').value=10;
  document.getElementById('cLimit').value=1;
  buildScopeOptions();
  openModal('modalGen');
};
document.getElementById('btnRefreshInvites').onclick=loadInvites;
document.getElementById('chkInviteAll').onchange=(e)=>{document.querySelectorAll('.inviteChk').forEach(c=>c.checked=e.target.checked);};

document.getElementById('doGen').onclick=async()=>{
  const scopes=Array.from(document.querySelectorAll('.scopeChk:checked')).map(c=>({globalId:c.getAttribute('data-g'), skuName:c.getAttribute('data-sku')}));
  if(!scopes.length){showToast('至少选择一个可用范围','error'); return;}
  
  const disabledPlans = Array.from(document.querySelectorAll('.genDisabledPlanChk:checked')).map(c => c.value);
    const disabledPlanNames = Array.from(document.querySelectorAll('.genDisabledPlanChk:checked')).map(c => {
      const name = c.getAttribute('data-name');
      return PLAN_FRIENDLY_NAMES[name] || name;
    });
  
  const codeLen=parseInt(document.querySelector('input[name="codeLen"]:checked')?.value||'16');
  const payload={codeLen, quantity:parseInt(document.getElementById('cQty').value)||1, limit:parseInt(document.getElementById('cLimit').value)||1, scopes, disabledPlans, disabledPlanNames};
  const res=await fetch(adminPath+'/api/invites/generate',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
  const data=await res.json();
  if(data.success){showToast('生成完成，新增 '+data.count+' 条','success'); closeModal('modalGen'); document.getElementById('chkInviteAll').checked=false; await loadInvites();}
  else showToast(data.message||'生成失败','error');
};

document.getElementById('btnDelInvites').onclick=()=>{
  const sel=Array.from(document.querySelectorAll('.inviteChk:checked')).map(c=>c.value);
  if(!sel.length){showToast('请选择邀请码','error'); return;}
  pendingDelCodes=sel;
  document.getElementById('delInviteCount').innerText=sel.length;
  openModal('modalDelInviteConfirm');
};
document.getElementById('confirmDelInvites').onclick=async()=>{
  closeModal('modalDelInviteConfirm');
  await fetch(adminPath+'/api/invites/bulk',{method:'DELETE',headers:{'Content-Type':'application/json'},body:JSON.stringify({codes:pendingDelCodes})});
  showToast('删除完成','success');
  document.getElementById('chkInviteAll').checked=false;
  
  // 智能退页：如果当前页删空且不是第一页
  const remainingInPage = invitesCache.length - pendingDelCodes.length;
  if(remainingInPage <= (invitePage-1)*invitePageSize && invitePage > 1) {
    invitePage--;
  }
  
  await loadInvites();
};

document.getElementById('btnExport').onclick=()=>{
  const sel=Array.from(document.querySelectorAll('.inviteChk:checked')).map(c=>c.value);
  if(!sel.length){showToast('请选择邀请码','error'); return;}
  const blob=new Blob([sel.join('\\n')],{type:'text/plain'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url; a.download='invites.txt'; a.click();
  URL.revokeObjectURL(url);
  showToast('导出完成','success');
};

document.querySelectorAll('th[data-sort]').forEach(th=>{th.onclick=()=>{const k=th.getAttribute('data-sort'); if(k===sortKey) sortDir*=-1; else {sortKey=k; sortDir=1;} renderInvites();};});

function setupInviteCustomSelect(triggerId, optionsId, displayId, inputId, onChange) {
  const trigger = document.getElementById(triggerId);
  const container = document.getElementById(optionsId);
  const wrapper = trigger.parentElement;
  trigger.onclick = (e) => { e.stopPropagation(); container.classList.toggle('open'); wrapper.classList.toggle('open'); };
  document.addEventListener('click', (e) => { if(!wrapper.contains(e.target)) {container.classList.remove('open'); wrapper.classList.remove('open');} });
  if(wrapper) {
    let leaveTimer;
    wrapper.addEventListener('mouseleave', () => { leaveTimer = setTimeout(() => { container.classList.remove('open'); wrapper.classList.remove('open'); }, 200); });
    wrapper.addEventListener('mouseenter', () => clearTimeout(leaveTimer));
  }
  container.addEventListener('click', (e) => {
    const opt = e.target.closest('.option');
    if (!opt) return;
    document.getElementById(displayId).innerText = opt.innerText;
    if (inputId) document.getElementById(inputId).value = opt.getAttribute('data-value');
    container.classList.remove('open');
    wrapper.classList.remove('open');
    document.querySelectorAll('#' + optionsId + ' .option').forEach(o => o.classList.remove('selected'));
    opt.classList.add('selected');
    if (onChange) onChange(opt);
  });
}

setupInviteCustomSelect('pageSizeInviteTrigger', 'pageSizeInviteOptions', 'pageSizeInviteDisplay', null, (opt) => {
  invitePageSize = parseInt(opt.getAttribute('data-value')) || 20;
  invitePage = 1;
  renderInvites();
});
setupInviteCustomSelect('iSearchFieldTrigger', 'iSearchFieldOptions', 'iSearchFieldDisplay', 'iSearchField');

function setupInvitePageSizeSelect(triggerId, optionsId, displayId) {
  const trigger = document.getElementById(triggerId);
  const container = document.getElementById(optionsId);
  const wrapper = trigger.parentElement;
  trigger.onclick = (e) => { e.stopPropagation(); container.classList.toggle('open'); wrapper.classList.toggle('open'); };
  document.addEventListener('click', (e) => { if(!wrapper.contains(e.target)) {container.classList.remove('open'); wrapper.classList.remove('open');} });
  if(wrapper) {
    let leaveTimer;
    wrapper.addEventListener('mouseleave', () => { leaveTimer = setTimeout(() => { container.classList.remove('open'); wrapper.classList.remove('open'); }, 200); });
    wrapper.addEventListener('mouseenter', () => clearTimeout(leaveTimer));
  }
  container.addEventListener('click', (e) => {
    const opt = e.target.closest('.option');
    if (!opt) return;
    document.getElementById(displayId).innerText = opt.innerText;
    container.classList.remove('open');
    wrapper.classList.remove('open');
    document.querySelectorAll('#' + optionsId + ' .option').forEach(o => o.classList.remove('selected'));
    opt.classList.add('selected');
    invitePageSize = parseInt(opt.getAttribute('data-value')) || 20;
    invitePage = 1;
    renderInvites();
  });
}
setupInvitePageSizeSelect('pageSizeInviteTrigger', 'pageSizeInviteOptions', 'pageSizeInviteDisplay');

document.getElementById('prevInvite').onclick=()=>{if(invitePage>1){invitePage--; renderInvites();}};
document.getElementById('nextInvite').onclick=()=>{const total=Math.max(1,Math.ceil(invitesCache.length/invitePageSize)); if(invitePage<total){invitePage++; renderInvites();}};
document.getElementById('goInvite').onclick=()=>{const val=parseInt(document.getElementById('jumpInvite').value)||1; const total=Math.max(1,Math.ceil(invitesCache.length/invitePageSize)); invitePage=Math.min(Math.max(1,val), total); renderInvites();};
document.getElementById('iSearchBtn').onclick=()=>{iSearchField=document.getElementById('iSearchField').value; iSearchText=document.getElementById('iSearchText').value.trim(); invitePage=1; renderInvites();};
document.getElementById('iClearBtn').onclick=()=>{document.getElementById('iSearchText').value=''; iSearchText=''; invitePage=1; renderInvites();};

function renderInvites(){
  updateIArrows();
  document.getElementById('chkInviteAll').checked=false;
  let list=[...invitesCache];
  if(iSearchText){
    const t=iSearchText.toLowerCase();
    list=list.filter(c=>{
      const st = c.used >= c.limit ? '已用完' : '可用';
      const scopeText = (c.allowed || []).map(s => {
        const g = globalsList.find(x => x.id === s.globalId);
        const disp = SUBSCRIPTION_FRIENDLY_NAMES[s.skuName] || s.skuName;
        return (g ? g.label : '') + ' ' + s.skuName + ' ' + disp;
      }).join(' ');
      const disabledText = (c.disabledPlanNames || []).join(' ') + ' ' + (c.disabledPlans || []).join(' ');
      
      if(iSearchField==='code') return (c.code||'').toLowerCase().includes(t);
      if(iSearchField==='limit') return String(c.limit).includes(t);
      if(iSearchField==='used') return String(c.used).includes(t);
      if(iSearchField==='status') return st.toLowerCase().includes(t);
      if(iSearchField==='scope') return scopeText.toLowerCase().includes(t) || disabledText.toLowerCase().includes(t);
      if(iSearchField==='all') {
        return (c.code||'').toLowerCase().includes(t) ||
               String(c.limit).includes(t) ||
               String(c.used).includes(t) ||
               st.toLowerCase().includes(t) ||
               scopeText.toLowerCase().includes(t) ||
               disabledText.toLowerCase().includes(t);
      }
      return true;
    });
  }
  list.sort((a,b)=>{
    if(sortKey==='status') return sortDir*(((a.used>=a.limit)?1:0)-((b.used>=b.limit)?1:0));
    if(sortKey==='scope'){
      const sa=(a.allowed||[]).map(s=>s.globalId+s.skuName).join(',');
      const sb=(b.allowed||[]).map(s=>s.globalId+s.skuName).join(',');
      return sortDir*sa.localeCompare(sb);
    }
    const va=a[sortKey]||0, vb=b[sortKey]||0;
    if(typeof va==='string') return sortDir*va.localeCompare(vb);
    return sortDir*((va>vb)-(va<vb));
  });
  const total=list.length;
  const totalPages=Math.max(1,Math.ceil(total/invitePageSize));
  invitePage=Math.min(invitePage,totalPages);
  const start=(invitePage-1)*invitePageSize;
  const pageData=list.slice(start,start+invitePageSize);
  const body=document.getElementById('inviteBody');
  body.innerHTML=pageData.map(c=>{
    const status=c.used>=c.limit?'<span class="tag" style="background:#fee2e2;color:#991b1b;">已用完</span>':'<span class="tag" style="background:#dcfce7;color:#166534;">可用</span>';
    let scope=(c.allowed||[]).map(s=>{
      const g=globalsList.find(x=>x.id===s.globalId);
      const disp = SUBSCRIPTION_FRIENDLY_NAMES[s.skuName] || s.skuName;
      return '<span class="tag" style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100%;display:inline-block;vertical-align:bottom;">'+(g?g.label:'?')+' / '+disp+'</span>';
    }).join('') || '<span style="color:#9ca3af;">未设置</span>';
    if (c.disabledPlanNames && c.disabledPlanNames.length > 0) {
      const disabledHtml = c.disabledPlanNames.map(name => '<span style="white-space:nowrap;display:inline-block;">'+name+'</span>').join(', ');
      scope += '<br><div class="tag" style="background:#fff1f2;color:#991b1b;border-color:#fecdd3;margin-top:4px;display:-webkit-box;-webkit-box-orient:vertical;-webkit-line-clamp:2;overflow:hidden;white-space:normal;line-height:1.5;">禁: '+disabledHtml+'</div>';
    }
    return '<tr><td data-label="选择" style="text-align:center;"><input type="checkbox" class="inviteChk" value="'+c.code+'"></td><td data-label="邀请码" style="overflow:hidden;text-overflow:ellipsis;"><code style="white-space:nowrap;display:inline-block;font-size:12px;">'+c.code+'</code></td><td data-label="限制次数" style="text-align:center;">'+c.limit+'</td><td data-label="已用" style="text-align:center;">'+c.used+'</td><td data-label="状态" style="text-align:center;">'+status+'</td><td data-label="限制范围" style="word-break:break-all;line-height:1.6;">'+scope+'</td><td data-label="生成时间" style="font-size:12px;color:var(--text-sub);">'+new Date(c.createdAt).toLocaleString()+'</td><td data-label="最近使用" style="font-size:12px;color:var(--text-sub);">'+(c.usedAt?new Date(c.usedAt).toLocaleString():'-')+'</td></tr>';
  }).join('') || '<tr><td colspan="8" style="text-align:center;">暂无邀请码</td></tr>';
  document.getElementById('pageInfoInvite').innerText='第 '+invitePage+' / '+totalPages+' 页 · 共 '+total+' 条';
}
async function loadInvites(){const res=await fetch(adminPath+'/api/invites?sort='+sortKey); const data=await res.json(); invitesCache=data; renderInvites();}
loadInvites();
</script>
`,
  });
}

function renderSettingsPage(adminPath, cfg) {
  const protectedPrefixes = (cfg.protectedPrefixes || []).join(',');
  return adminLayout({
    title: '设置',
    adminPath,
    active: 'settings',
    content: `
<div class="section">
  <h3 style="margin-top:0;">后台账号</h3>
  <div class="row"><span class="label">用户名</span><input id="sAdminUser" value="${cfg.adminUsername || 'admin'}" placeholder="例如: admin"></div>
  <div class="row"><span class="label">新密码</span><input id="sAdminPwd" type="password" placeholder="留空不修改 (至少 8 位)"></div>
  <div style="color:#6b7280;font-size:12px;line-height:1.6;margin-top:6px;">修改后当前会话不受影响，下次按新账号登录。</div>
</div>

<div class="section">
  <h3 style="margin-top:0;">基础设置</h3>
  <div class="row"><span class="label">后台路径</span><input id="sPath" value="${cfg.adminPath}" placeholder="/admin"></div>
  <div class="row"><span class="label">Turnstile Site Key (留空关闭)</span><input id="sSite" value="${cfg.turnstile.siteKey||''}"></div>
  <div class="row"><span class="label">Turnstile Secret Key (留空关闭)</span><input id="sSecret" value="${cfg.turnstile.secretKey||''}"></div>
  <div class="row"><label class="inline"><input type="checkbox" id="sInvite" ${cfg.invite?.enabled?'checked':''}> 启用邀请码注册</label></div>
</div>

<div class="section">
  <h3 style="margin-top:0;">额外保护账户（禁止注册）</h3>
  <div class="row"><span class="label">额外保护账户（逗号分隔）</span><textarea id="sProtectPrefixes" rows="3" placeholder="例如: admin,superadmin,root">${protectedPrefixes}</textarea></div>
  <div style="color:#6b7280;font-size:12px;line-height:1.6;">匹配邮箱 @ 前缀。命中的用户名禁止前台注册，且禁止删除。内置常见敏感用户名，建议保留。</div>
  <div class="toolbar" style="margin-top:14px;"><button id="btnSaveSetting">${ICONS.save} 保存设置</button></div>
</div>

<script>
const adminPath='${adminPath}';
function parseCommaList(v){return (v||'').split(',').map(s=>s.trim()).filter(Boolean);}
document.getElementById('btnSaveSetting').onclick=async()=>{
  const adminUsername=(document.getElementById('sAdminUser').value||'').trim();
  const adminPassword=document.getElementById('sAdminPwd').value||'';
  const payload={adminPath:(document.getElementById('sPath').value||'/admin').trim(), adminUsername, adminPassword: adminPassword ? adminPassword : undefined, turnstile:{siteKey:(document.getElementById('sSite').value||'').trim(), secretKey:(document.getElementById('sSecret').value||'').trim()}, protectedPrefixes:parseCommaList(document.getElementById('sProtectPrefixes').value), inviteEnabled:document.getElementById('sInvite').checked};
  const res=await fetch(adminPath+'/api/config',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
  const data=await res.json();
  if(data.success){showToast('保存成功','success'); if(data.newPath && data.newPath !== adminPath) location.href = data.newPath + '/settings';}
  else showToast(data.message||'保存失败','error');
};
</script>
`,
  });
}

/* -------------------- Register logic -------------------- */
async function handleRegister(env, req, cfg) {
  const form = await req.formData();
  const username = (form.get('username') || '').trim();
  const password = form.get('password') || '';
  const skuName = form.get('skuName');
  const globalId = form.get('globalId');
  const inviteCode = form.get('inviteCode');
  const turnstileToken = form.get('cf-turnstile-response');
  const clientIp = req.headers.get('CF-Connecting-IP');

  const global = (cfg.globals || []).find((g) => g.id === globalId);
  if (!global) return jsonResponse({ success: false, message: '请选择有效全局' }, 400);
  const skuId = (global.skuMap || {})[skuName];
  if (!skuId) return jsonResponse({ success: false, message: '请选择有效订阅' }, 400);
  if (!/^[a-zA-Z0-9]+$/.test(username)) return jsonResponse({ success: false, message: '用户名格式错误' }, 400);

  let disabledPlansToApply = [];
  if (cfg.invite?.enabled) {
    const invites = await getInvites(env);
    const idx = invites.findIndex((c) => c.code === inviteCode);
    if (idx === -1) return jsonResponse({ success: false, message: '邀请码无效' }, 400);
    const c = invites[idx];
    if (c.used >= c.limit) return jsonResponse({ success: false, message: '邀请码已用完' }, 400);
    const matched = (c.allowed || []).some((a) => a.globalId === globalId && a.skuName === skuName);
    if (!matched) return jsonResponse({ success: false, message: '邀请码不允许当前全局/订阅' }, 400);
    
    if (c.disabledPlans && Array.isArray(c.disabledPlans)) {
      disabledPlansToApply = c.disabledPlans;
    }
    
    c.used += 1;
    c.usedAt = Date.now();
    invites[idx] = c;
    await saveInvites(env, invites);
  }

  if (cfg.turnstile?.secretKey) {
    if (!turnstileToken) {
      return jsonResponse({ success: false, message: '请完成人机验证' }, 400);
    }
    const ver = await fetch('https://challenges.cloudflare.com/turnstile/v0/siteverify', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        secret: cfg.turnstile.secretKey,
        response: turnstileToken,
        remoteip: clientIp || undefined,
      }),
    });
    const verData = await ver.json().catch(() => ({}));
    if (!ver.ok) {
      return jsonResponse({ success: false, message: '人机验证服务异常，请稍后重试' }, 502);
    }
    if (!verData.success) {
      return jsonResponse({ success: false, message: '人机验证失败', errors: verData['error-codes'] || [] }, 400);
    }
  }

  const userEmail = `${username}@${global.defaultDomain}`;
  if (isProtectedUpn(userEmail, env, cfg)) {
    return jsonResponse({ success: false, message: '该用户名被禁止注册！请勿尝试注册非法用户名！' }, 403);
  }

  if (password.toLowerCase().includes(username.toLowerCase())) {
    return jsonResponse({ success: false, message: '密码不能包含用户名' }, 400);
  }
  if (!checkPasswordComplexity(password)) {
    return jsonResponse({ success: false, message: '密码不符合复杂度' }, 400);
  }

  const token = await getAccessTokenForGlobal(global, fetch);

  const createResp = await fetch('https://graph.microsoft.com/v1.0/users', {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      accountEnabled: true,
      displayName: username,
      mailNickname: username,
      userPrincipalName: userEmail,
      passwordProfile: { forceChangePasswordNextSignIn: false, password },
      usageLocation: 'CN',
    }),
  });
  if (!createResp.ok) {
    const err = await createResp.json().catch(() => ({}));
    return jsonResponse({ success: false, message: err.error?.message || '创建失败' }, 400);
  }
  const newUser = await createResp.json();

  const licResp = await fetch(`https://graph.microsoft.com/v1.0/users/${newUser.id}/assignLicense`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ addLicenses: [{ disabledPlans: disabledPlansToApply, skuId }], removeLicenses: [] }),
  });
  if (!licResp.ok) {
    const err = await licResp.json().catch(() => ({}));
    // 回滚：订阅分配失败时，尝试删除刚创建的用户，避免留下无订阅的废弃账号
    try {
      await fetch(`https://graph.microsoft.com/v1.0/users/${newUser.id}`, {
        method: 'DELETE',
        headers: { Authorization: `Bearer ${token}` },
      });
    } catch (rollbackErr) {
      console.error('[handleRegister] 回滚删除用户失败:', rollbackErr?.message || rollbackErr);
    }
    return jsonResponse({ success: false, message: '订阅分配失败，账号已自动撤销: ' + (err.error?.message || '未知') }, 400);
  }
  return jsonResponse({ success: true, email: userEmail });
}

/* -------------------- Worker -------------------- */
export default {
  async fetch(request, env) {
    if (!env || !env.CONFIG_KV) {
      return new Response('系统配置错误：未绑定 KV 命名空间。请在 Cloudflare Workers 设置中绑定 KV 命名空间，且变量名必须为 CONFIG_KV', {
        status: 500,
        headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
      });
    }

    const url = new URL(request.url);
    let cfg = await getConfig(env);
    const adminPath = cfg.adminPath || '/admin';
    const installed = !!(await env.CONFIG_KV.get(KV.INSTALL_LOCK));
    const isSetupPath = url.pathname === `${adminPath}/setup`;
    const isLoginPath = url.pathname === adminPath || url.pathname === `${adminPath}/`;

    if (!installed && !isSetupPath) return redirect(`${adminPath}/setup`);

    if (isSetupPath) {
      if (request.method === 'GET') return htmlResponse(renderSetup(adminPath));
      if (request.method === 'POST') {
        const body = await request.json().catch(() => ({}));
        const username = (body.username || '').toString().trim();
        const password = (body.password || '').toString();
        if (!/^[a-zA-Z0-9_\\-]{3,32}$/.test(username)) {
          return jsonResponse({ success: false, message: '用户名格式不正确（3-32位，仅字母/数字/_/-）' }, 400);
        }
        if (!password || password.length < 8) {
          return jsonResponse({ success: false, message: '密码至少 8 位' }, 400);
        }
        const newPath = (body.adminPath || '/admin').toString().trim() || '/admin';
        const hash = await sha256(password);
        cfg = mergeConfig({ ...cfg, adminUsername: username, adminPasswordHash: hash, adminPath: newPath });
        await setConfig(env, cfg);
        await env.CONFIG_KV.put(KV.INSTALL_LOCK, '1');
        return jsonResponse({ success: true });
      }
    }

    // 登出逻辑处理
    if (url.pathname === `${adminPath}/logout`) {
      const cookies = parseCookies(request);
      const token = cookies.ADMIN_SESSION;
      if (token) {
        await env.CONFIG_KV.delete(KV.SESS_PREFIX + token);
      }
      return new Response(null, {
        status: 302,
        headers: {
          'Location': '/',
          'Set-Cookie': 'ADMIN_SESSION=; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=0; Expires=Thu, 01 Jan 1970 00:00:00 GMT'
        }
      });
    }

    // 登录页面处理
    if (isLoginPath) {
      if (request.method === 'GET') {
        if (await verifySession(env, request)) {
          return redirect(`${adminPath}/users`);
        }
        return htmlResponse(renderLogin(adminPath));
      }
      if (request.method === 'POST') {
        const body = await request.json().catch(() => ({}));
        const username = (body.username || '').toString().trim();
        const pwdHash = await sha256((body.password || '').toString());
        const cfgUser = (cfg.adminUsername || 'admin').toString().trim();
        if (username.toLowerCase() !== cfgUser.toLowerCase() || pwdHash !== cfg.adminPasswordHash) {
          return jsonResponse({ success: false, message: '用户名或密码错误' }, 401);
        }
        const token = await createSession(env);
        return new Response(JSON.stringify({ success: true }), {
          headers: {
            'Content-Type': 'application/json',
            'Set-Cookie': `ADMIN_SESSION=${token}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=604800`,
          },
        });
      }
    }

    if (url.pathname === `${adminPath}/users`) {
      if (!(await verifySession(env, request))) return redirect(adminPath);
      return htmlResponse(renderUsersPage(adminPath));
    }
    if (url.pathname === `${adminPath}/globals`) {
      if (!(await verifySession(env, request))) return redirect(adminPath);
      return htmlResponse(renderGlobalsPage(adminPath));
    }
    if (url.pathname === `${adminPath}/invites`) {
      if (!(await verifySession(env, request))) return redirect(adminPath);
      return htmlResponse(renderInvitesPage(adminPath, cfg.globals || []));
    }
    if (url.pathname === `${adminPath}/settings`) {
      if (!(await verifySession(env, request))) return redirect(adminPath);
      return htmlResponse(renderSettingsPage(adminPath, cfg));
    }

    if (url.pathname.startsWith(`${adminPath}/api/`)) {
      if (!(await verifySession(env, request))) return jsonResponse({ error: 'unauthorized' }, 401);

      if (url.pathname === `${adminPath}/api/fetch_skus` && request.method === 'POST') {
        const body = await request.json().catch(() => ({}));
        const tenantId = (body.tenantId || '').trim();
        const clientId = (body.clientId || '').trim();
        const clientSecret = (body.clientSecret || '').trim();
        if (!tenantId || !clientId || !clientSecret) return jsonResponse({ success: false, message: '缺少租户/客户端信息' }, 400);
        try {
          const tmp = { tenantId, clientId, clientSecret };
          const token = await getAccessTokenForGlobal(tmp, fetch);
          const resp = await fetch('https://graph.microsoft.com/v1.0/subscribedSkus', { headers: { Authorization: `Bearer ${token}` } });
          if (!resp.ok) {
            const err = await resp.json().catch(() => ({}));
            return jsonResponse({ success: false, message: err?.error?.message || '获取失败' }, 400);
          }
          const data = await resp.json();
          const map = {};
          (data.value || []).forEach((s) => {
            const friendly = SUBSCRIPTION_FRIENDLY_NAMES[s.skuPartNumber];
            const key = friendly || s.skuPartNumber;
            map[key] = s.skuId;
          });
          return jsonResponse({ success: true, map });
        } catch (e) {
          return jsonResponse({ success: false, message: e.message || '获取失败' }, 400);
        }
      }

      if (url.pathname === `${adminPath}/api/service_plans` && request.method === 'POST') {
        const body = await request.json().catch(() => ({}));
        const globalId = body.globalId;
        const skuId = body.skuId;
        const global = (cfg.globals || []).find((g) => g.id === globalId);
        if (!global) return jsonResponse({ success: false, message: '全局不存在' }, 400);
        try {
          const plans = await fetchServicePlans(global, skuId, fetch);
          return jsonResponse({ success: true, plans });
        } catch (e) {
          return jsonResponse({ success: false, message: e.message || '获取失败' }, 400);
        }
      }

      if (url.pathname === `${adminPath}/api/globals` && request.method === 'GET') {
        const list = (cfg.globals || []).map((g) => ({ ...g, clientSecret: undefined }));
        return jsonResponse(list);
      }
      if (url.pathname === `${adminPath}/api/globals` && request.method === 'POST') {
        const body = await request.json().catch(() => ({}));
        const id = crypto.randomUUID();
        const item = {
          id,
          label: body.label || '未命名',
          defaultDomain: body.defaultDomain || '',
          tenantId: body.tenantId || '',
          clientId: body.clientId || '',
          clientSecret: body.clientSecret || '',
          skuMap: sanitizeSkuMap(body.skuMap),
        };
        cfg.globals = cfg.globals || [];
        cfg.globals.push(item);
        await setConfig(env, cfg);
        return jsonResponse({ success: true, id });
      }
      if (url.pathname.match(new RegExp(`^${adminPath}/api/globals/[^/]+$`)) && request.method === 'GET') {
        const gid = url.pathname.split('/').pop();
        const g = (cfg.globals || []).find((x) => x.id === gid);
        if (!g) return jsonResponse({ error: 'not found' }, 404);
        return jsonResponse(g);
      }
      if (url.pathname.match(new RegExp(`^${adminPath}/api/globals/[^/]+$`)) && request.method === 'PATCH') {
        const gid = url.pathname.split('/').pop();
        const body = await request.json().catch(() => ({}));
        const idx = (cfg.globals || []).findIndex((x) => x.id === gid);
        if (idx === -1) return jsonResponse({ error: 'not found' }, 404);
        cfg.globals[idx] = {
          ...cfg.globals[idx],
          label: body.label || cfg.globals[idx].label,
          defaultDomain: body.defaultDomain || cfg.globals[idx].defaultDomain,
          tenantId: body.tenantId || cfg.globals[idx].tenantId,
          clientId: body.clientId || cfg.globals[idx].clientId,
          clientSecret: body.clientSecret || cfg.globals[idx].clientSecret,
          skuMap: body.skuMap ? sanitizeSkuMap(body.skuMap) : cfg.globals[idx].skuMap,
        };
        await setConfig(env, cfg);
        return jsonResponse({ success: true });
      }
      if (url.pathname.match(new RegExp(`^${adminPath}/api/globals/[^/]+$`)) && request.method === 'DELETE') {
        const gid = url.pathname.split('/').pop();
        cfg.globals = (cfg.globals || []).filter((x) => x.id !== gid);
        await setConfig(env, cfg);
        return jsonResponse({ success: true });
      }

      if (url.pathname === `${adminPath}/api/config` && request.method === 'POST') {
        const body = await request.json().catch(() => ({}));
        const newPath = (body.adminPath || adminPath).toString().trim() || adminPath;
        if (body.adminUsername !== undefined) {
          const u = (body.adminUsername || '').toString().trim();
          if (!/^[a-zA-Z0-9_\\-]{3,32}$/.test(u)) {
            return jsonResponse({ success: false, message: '用户名格式不正确（3-32位，仅字母/数字/_/-）' }, 400);
          }
          cfg.adminUsername = u;
        }
        if (body.adminPassword) {
          const p = body.adminPassword.toString();
          if (p.length < 8) return jsonResponse({ success: false, message: '密码至少 8 位' }, 400);
          cfg.adminPasswordHash = await sha256(p);
        }
        cfg.turnstile = body.turnstile || cfg.turnstile;
        cfg.protectedUsers = Array.isArray(body.protectedUsers) ? body.protectedUsers : (cfg.protectedUsers || []);
        cfg.protectedPrefixes = Array.isArray(body.protectedPrefixes) ? body.protectedPrefixes : (cfg.protectedPrefixes || []);
        cfg.invite = { ...(cfg.invite || {}), enabled: !!body.inviteEnabled };
        cfg.adminPath = newPath;
        cfg = mergeConfig(cfg);
        await setConfig(env, cfg);
        return jsonResponse({ success: true, newPath });
      }

      if (url.pathname === `${adminPath}/api/users/create` && request.method === 'POST') {
        const body = await request.json().catch(() => ({}));
        const username = (body.username || '').trim();
        const password = body.password || '';
        const skuName = body.skuName;
        const globalId = body.globalId;
        const forceChangePasswordNextSignIn = !!body.forceChangePasswordNextSignIn;
        const global = (cfg.globals || []).find((g) => g.id === globalId);
        if (!global) return jsonResponse({ success: false, message: '请选择有效全局' }, 400);
        const skuId = (global.skuMap || {})[skuName];
        if (!skuId) return jsonResponse({ success: false, message: '请选择有效订阅' }, 400);
        if (!/^[a-zA-Z0-9]+$/.test(username)) return jsonResponse({ success: false, message: '用户名格式错误' }, 400);
        if (password.toLowerCase().includes(username.toLowerCase())) return jsonResponse({ success: false, message: '密码不能包含用户名' }, 400);
        if (!checkPasswordComplexity(password)) return jsonResponse({ success: false, message: '密码不符合复杂度' }, 400);
        const userEmail = `${username}@${global.defaultDomain}`;
        if (isProtectedUpn(userEmail, env, cfg)) return jsonResponse({ success: false, message: '该用户名被禁止创建' }, 403);

        const token = await getAccessTokenForGlobal(global, fetch);
        const createResp = await fetch('https://graph.microsoft.com/v1.0/users', {
          method: 'POST',
          headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({
            accountEnabled: true,
            displayName: username,
            mailNickname: username,
            userPrincipalName: userEmail,
            passwordProfile: { forceChangePasswordNextSignIn, password },
            usageLocation: 'CN',
          }),
        });
        if (!createResp.ok) {
          const err = await createResp.json().catch(() => ({}));
          return jsonResponse({ success: false, message: err.error?.message || '创建失败' }, 400);
        }
        const newUser = await createResp.json();
        const licResp = await fetch(`https://graph.microsoft.com/v1.0/users/${newUser.id}/assignLicense`, {
          method: 'POST',
          headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({ addLicenses: [{ disabledPlans: [], skuId }], removeLicenses: [] }),
        });
        if (!licResp.ok) {
          const err = await licResp.json().catch(() => ({}));
          return jsonResponse({ success: false, message: '账号已创建但订阅分配失败: ' + (err.error?.message || '未知') }, 400);
        }
        return jsonResponse({ success: true, email: userEmail });
      }

      if (url.pathname === `${adminPath}/api/users` && request.method === 'GET') {
        // 并发拉取所有全局的用户，同时支持 Graph API 分页（@odata.nextLink）
        const globalResults = await Promise.all(
          (cfg.globals || []).map(async (g) => {
            try {
              const token = await getAccessTokenForGlobal(g, fetch);
              let arr = await fetchAllGraphUsers(token, fetch);
              arr = filterProtectedUsers(arr, env, cfg);
              const idToName = Object.entries(g.skuMap || {}).reduce((m, [k, v]) => { m[v] = k; return m; }, {});
              arr.forEach((u) => {
                u.assignedLicenses = (u.assignedLicenses || []).map((l) => ({ ...l, name: idToName[l.skuId] || l.skuId || '' }));
                u._licSort = (u.assignedLicenses || []).map((l) => l.name || '').join(',');
                u._globalId = g.id;
                u._globalLabel = g.label;
              });
              return arr;
            } catch (e) {
              console.error(`[api/users] 全局 ${g.label} (${g.id}) 拉取失败:`, e?.message || e);
              return [];
            }
          })
        );
        const result = globalResults.flat();
        return jsonResponse(result);
      }

      if (url.pathname.match(new RegExp(`^${adminPath}/api/users/[^/]+/[^/]+$`)) && request.method === 'DELETE') {
        const parts = url.pathname.split('/');
        const userId = parts.pop();
        const gId = parts.pop();
        const g = (cfg.globals || []).find((x) => x.id === gId);
        if (!g) return jsonResponse({ error: 'not found' }, 404);
        const token = await getAccessTokenForGlobal(g, fetch);
        const checkResp = await fetch(`https://graph.microsoft.com/v1.0/users/${userId}?$select=userPrincipalName`, { headers: { Authorization: `Bearer ${token}` } });
        if (!checkResp.ok) return jsonResponse({ error: 'cannot_verify_user' }, 502);
        const user = await checkResp.json();
        const upn = user.userPrincipalName || '';
        if (isProtectedUpn(upn, env, cfg)) return jsonResponse({ error: 'forbidden' }, 403);
        const delResp = await fetch(`https://graph.microsoft.com/v1.0/users/${userId}`, { method: 'DELETE', headers: { Authorization: `Bearer ${token}` } });
        if (!delResp.ok) {
          const t = await delResp.text().catch(() => '');
          return jsonResponse({ error: 'delete_failed', details: t.slice(0, 300) }, delResp.status);
        }
        return jsonResponse({ success: true });
      }

      if (url.pathname.match(new RegExp(`^${adminPath}/api/users/[^/]+/[^/]+/password$`)) && request.method === 'PATCH') {
        const parts = url.pathname.split('/');
        const userId = parts[parts.length - 2];
        const gId = parts[parts.length - 3];
        const body = await request.json().catch(() => ({}));
        const g = (cfg.globals || []).find((x) => x.id === gId);
        if (!g) return jsonResponse({ error: 'not found' }, 404);
        const token = await getAccessTokenForGlobal(g, fetch);
        await fetch(`https://graph.microsoft.com/v1.0/users/${userId}`, {
          method: 'PATCH',
          headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({ passwordProfile: { forceChangePasswordNextSignIn: false, password: body.password } }),
        });
        return jsonResponse({ success: true });
      }

      if (url.pathname === `${adminPath}/api/licenses` && request.method === 'GET') {
        // 并发拉取所有全局的订阅信息
        const globalResults = await Promise.all(
          (cfg.globals || []).map(async (g) => {
            try {
              const token = await getAccessTokenForGlobal(g, fetch);
              const expiryBySkuId = {};
              try {
                const subResp = await fetch('https://graph.microsoft.com/v1.0/directory/subscriptions?$select=skuId,skuPartNumber,nextLifecycleDateTime,status', {
                  headers: { Authorization: `Bearer ${token}` },
                });
                if (subResp.ok) {
                  const subData = await subResp.json().catch(() => ({}));
                  (subData.value || []).forEach((cs) => {
                    const skuId = (cs.skuId || '').toString().toLowerCase();
                    const dt = cs.nextLifecycleDateTime;
                    if (!skuId || !dt) return;
                    if (!expiryBySkuId[skuId] || new Date(dt) < new Date(expiryBySkuId[skuId])) expiryBySkuId[skuId] = dt;
                  });
                }
              } catch {}
              const resp = await fetch('https://graph.microsoft.com/v1.0/subscribedSkus', { headers: { Authorization: `Bearer ${token}` } });
              const data = await resp.json().catch(() => ({}));
              return (data.value || []).map((s) => {
                const skuIdLower = (s.skuId || '').toString().toLowerCase();
                return {
                  globalId: g.id,
                  globalLabel: g.label,
                  skuPartNumber: s.skuPartNumber,
                  skuId: s.skuId,
                  total: s.prepaidUnits?.enabled || 0,
                  used: s.consumedUnits || 0,
                  expiresAt: expiryBySkuId[skuIdLower] || null,
                };
              });
            } catch (e) {
              console.error(`[api/licenses] 全局 ${g.label} (${g.id}) 拉取失败:`, e?.message || e);
              return [];
            }
          })
        );
        return jsonResponse(globalResults.flat());
      }

      if (url.pathname === `${adminPath}/api/invites` && request.method === 'GET') {
        await ensureInvites(env);
        return jsonResponse(await getInvites(env));
      }
      if (url.pathname === `${adminPath}/api/invites/generate` && request.method === 'POST') {
        const body = await request.json().catch(() => ({}));
        // 支持 16 位或 32 位固定长度邀请码，默认 16 位
        const codeLen = [16, 32].includes(parseInt(body.codeLen)) ? parseInt(body.codeLen) : 16;
        const qty = Math.max(1, parseInt(body.quantity) || 1);
        const limit = Math.max(1, parseInt(body.limit) || 1);
        const scopes = body.scopes || [];
        const disabledPlans = Array.isArray(body.disabledPlans) ? body.disabledPlans : [];
        const disabledPlanNames = Array.isArray(body.disabledPlanNames) ? body.disabledPlanNames : [];
        if (!scopes.length) return jsonResponse({ success: false, message: '请选择限制范围' }, 400);
        
        await ensureInvites(env);
        const invites = await getInvites(env);
        let created = 0;
        let attempts = 0;
        const maxAttempts = qty * 20;
        while (created < qty && attempts < maxAttempts) {
          attempts++;
          let code = generateFixedLengthCode(codeLen);
          if (disabledPlans.length > 0) {
            code = 'LMT-' + code; // 限制权限的邀请码增加特殊前缀
          } else {
            code = 'INV-' + code; // 普通邀请码增加特殊前缀
          }
          if (inviteCodeExists(invites, code)) continue;
          invites.push({ code, limit, used: 0, createdAt: Date.now(), usedAt: null, allowed: scopes, disabledPlans, disabledPlanNames });
          created++;
        }
        await saveInvites(env, invites);
        return jsonResponse({ success: true, count: created });
      }
      if (url.pathname === `${adminPath}/api/invites/bulk` && request.method === 'DELETE') {
        const body = await request.json().catch(() => ({ codes: [] }));
        const codes = body.codes || [];
        const invites = await getInvites(env);
        const filtered = invites.filter((c) => !codes.includes(c.code));
        await saveInvites(env, filtered);
        return jsonResponse({ success: true, removed: codes.length });
      }
    }

    if (request.method === 'GET' && url.pathname === '/') {
      const globals = (cfg.globals || []).map((g) => ({ id: g.id, label: g.label }));
      const selectedGlobalId = url.searchParams.get('g') || globals[0]?.id || '';
      const selectedGlobal = (cfg.globals || []).find((g) => g.id === selectedGlobalId) || (cfg.globals || [])[0];
      let skuDisplayList = [];
      if (selectedGlobal) {
        try {
          const subscribed = await fetchSubscribedSkus(selectedGlobal, fetch);
          const bySkuId = new Map(subscribed.map((s) => [String(s.skuId).toLowerCase(), s]));
          const skuMap = selectedGlobal.skuMap || {};
          skuDisplayList = Object.keys(skuMap).map((name) => {
            const skuId = String(skuMap[name] || '').toLowerCase();
            const sku = bySkuId.get(skuId);
            const rem = sku ? remainingFromSubscribedSku(sku) : 0;
            const disp = SUBSCRIPTION_FRIENDLY_NAMES[name] || name;
            return { name, remaining: rem, label: `${disp}（剩余总量：${rem}）` };
          });
          skuDisplayList.sort((a, b) => (b.remaining - a.remaining) || a.name.localeCompare(b.name));
        } catch {
          const skuMap = selectedGlobal.skuMap || {};
          skuDisplayList = Object.keys(skuMap).map((name) => ({ name, remaining: 0, label: `${name}（剩余总量：0）` }));
        }
      }
      return htmlResponse(renderRegisterPage({
        globals,
        selectedGlobalId,
        skuDisplayList,
        protectedPrefixes: cfg.protectedPrefixes || [],
        turnstileSiteKey: cfg.turnstile?.siteKey || '',
        inviteMode: !!cfg.invite?.enabled,
        adminPath,
      }));
    }

    if (request.method === 'POST' && url.pathname === '/') {
      cfg = await getConfig(env);
      return handleRegister(env, request, cfg);
    }

    // 后台其它未知路径，若未登录则重定向到登录页，若已登录则重定向到 /users
    if (url.pathname.startsWith(adminPath)) {
      if (!(await verifySession(env, request))) return redirect(adminPath);
      return redirect(`${adminPath}/users`);
    }

    // 前台其它未知路径，重定向到注册主页
    if (request.method === 'GET' && !url.pathname.startsWith(adminPath) && url.pathname !== '/') {
      return redirect('/');
    }

    return htmlResponse(render404(adminPath), 404);
  },
};
