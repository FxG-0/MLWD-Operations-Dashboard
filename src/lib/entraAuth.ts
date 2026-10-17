import { PublicClientApplication, Configuration, AccountInfo } from "@azure/msal-browser";

// Main app sign-in — separate Entra ID App Registration from the OneDrive/Graph
// integration in msAuth.ts. This one gates who can use the dashboard at all.
const HALLIBURTON_TENANT_ID = "b7be7686-6f97-4db7-9081-a23cf09a96b5";
const APP_CLIENT_ID = "3c402a5a-d1f3-4387-86c5-1418152cc654";

const msalConfig: Configuration = {
  auth: {
    clientId: APP_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${HALLIBURTON_TENANT_ID}`,
    redirectUri: typeof window !== "undefined" ? window.location.origin : undefined,
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

export const entraMsal = new PublicClientApplication(msalConfig);

let initialized = false;
let initPromise: Promise<void> | null = null;

async function ensureInit(): Promise<void> {
  if (initialized) return;
  if (!initPromise) {
    initPromise = entraMsal.initialize().then(() => {
      initialized = true;
    });
  }
  await initPromise;
}

const loginRequest = {
  scopes: ["User.Read"],
};

/**
 * Call once on app startup, before rendering anything auth-dependent.
 * Completes the redirect flow if the page just came back from Microsoft's
 * login page, and returns the signed-in account if one exists.
 */
export async function entraHandleReturn(): Promise<AccountInfo | null> {
  await ensureInit();
  const result = await entraMsal.handleRedirectPromise();
  if (result && result.account) {
    entraMsal.setActiveAccount(result.account);
    return result.account;
  }
  const existing = entraMsal.getActiveAccount() || entraMsal.getAllAccounts()[0] || null;
  if (existing) entraMsal.setActiveAccount(existing);
  return existing;
}

/**
 * Kicks off sign-in via full-page redirect to Microsoft's login page.
 * The page will reload; entraHandleReturn() picks up the result after.
 */
export async function entraSignInRedirect(): Promise<void> {
  await ensureInit();
  await entraMsal.loginRedirect(loginRequest);
}

export async function entraSignOut(): Promise<void> {
  await ensureInit();
  const account = entraMsal.getActiveAccount();
  await entraMsal.logoutRedirect({
    account: account || undefined,
    postLogoutRedirectUri: window.location.origin,
  });
}

/**
 * Gets a fresh access token for calling this app's own backend API.
 * Silent refresh first, falls back to a redirect if the session's stale.
 */
export async function getAuthHeader(): Promise<Record<string, string>> {
  await ensureInit();
  let account = entraMsal.getActiveAccount();
  if (!account) {
    account = entraMsal.getAllAccounts()[0];
    if (account) entraMsal.setActiveAccount(account);
  }
  if (!account) return {};

  try {
    const result = await entraMsal.acquireTokenSilent({ ...loginRequest, account });
    return { Authorization: `Bearer ${result.idToken}` };
  } catch (err) {
    await entraMsal.acquireTokenRedirect(loginRequest);
    return {};
  }
}

export interface EntraProfile {
  oid: string;
  email: string;
  name: string;
}

export function getEntraProfile(): EntraProfile | null {
  const account = entraMsal.getActiveAccount();
  if (!account) return null;
  const claims: any = account.idTokenClaims || {};
  return {
    oid: claims.oid || account.homeAccountId,
    email: claims.preferred_username || account.username || "",
    name: claims.name || account.name || account.username || "User",
  };
}

/**
 * fetch() wrapper that automatically attaches the bearer token.
 * Use this for every call to this app's own /api/* routes that need auth.
 */
export async function authFetch(url: string, options: RequestInit = {}): Promise<Response> {
  const authHeader = await getAuthHeader();
  return fetch(url, {
    ...options,
    headers: {
      ...(options.headers || {}),
      ...authHeader,
    },
  });
}
