import { PublicClientApplication, SilentRequest, InteractionRequiredAuthError } from "@azure/msal-browser";

/**
 * MSAL configuration for the Outlook add-in.
 *
 * IMPORTANT: Replace the clientId below with your own Azure AD / Entra ID
 * application (client) ID after registering the app in the Azure portal.
 *
 * Required API permissions (delegated):
 *   - User.Read           (sign-in and read user profile)
 *   - MailboxSettings.ReadWrite (read/write auto-reply settings)
 */
const msalConfig = {
  auth: {
    clientId: "YOUR_CLIENT_ID_HERE", // <-- Replace with your Entra app registration client ID
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "localStorage" as const,
  },
};

const GRAPH_SCOPES = ["User.Read", "MailboxSettings.ReadWrite"];

let msalInstance: PublicClientApplication | null = null;

/** Initialise the MSAL PublicClientApplication singleton. */
async function getMsalInstance(): Promise<PublicClientApplication> {
  if (!msalInstance) {
    msalInstance = new PublicClientApplication(msalConfig);
    await msalInstance.initialize();
  }
  return msalInstance;
}

/**
 * Acquire a Microsoft Graph access token.
 *
 * Tries silent acquisition first (cached token / SSO). Falls back to
 * interactive popup if the user hasn't consented yet.
 */
export async function getAccessToken(): Promise<string> {
  const pca = await getMsalInstance();

  // Try Office SSO token first (if running inside Outlook)
  try {
    if (typeof Office !== "undefined" && Office.auth) {
      const ssoToken = await Office.auth.getAccessToken({ allowSignInPrompt: true });
      // The SSO token is an ID token — exchange it for a Graph token via OBO
      // if you have a backend. For client-only, fall through to MSAL popup.
      if (ssoToken) {
        // For simplicity, we use MSAL directly. SSO token can be used as a
        // login hint to avoid re-prompting.
      }
    }
  } catch {
    // SSO not available — continue with MSAL
  }

  const accounts = pca.getAllAccounts();
  const silentRequest: SilentRequest = {
    scopes: GRAPH_SCOPES,
    account: accounts[0] || undefined,
  };

  try {
    const result = await pca.acquireTokenSilent(silentRequest);
    return result.accessToken;
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError || accounts.length === 0) {
      const result = await pca.acquireTokenPopup({ scopes: GRAPH_SCOPES });
      return result.accessToken;
    }
    throw err;
  }
}

/** Sign the user out and clear cached tokens. */
export async function signOut(): Promise<void> {
  const pca = await getMsalInstance();
  const accounts = pca.getAllAccounts();
  if (accounts.length > 0) {
    await pca.logoutPopup({ account: accounts[0] });
  }
}

/** Check whether a cached session exists (user previously signed in). */
export async function isSignedIn(): Promise<boolean> {
  const pca = await getMsalInstance();
  return pca.getAllAccounts().length > 0;
}
