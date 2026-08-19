import {
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
} from "@azure/msal-browser";

const clientId = import.meta.env.VITE_AZURE_CLIENT_ID as string | undefined;
const tenantId = import.meta.env.VITE_AZURE_TENANT_ID as string | undefined;

export const graphScopes = [
  "User.Read",
  "Files.Read.All",
  "Sites.Read.All",
];

export const authConfigured = Boolean(clientId && tenantId && !clientId.startsWith("your-"));

let pca: PublicClientApplication | null = null;

export function getMsal(): PublicClientApplication {
  if (!authConfigured) {
    throw new Error("Missing VITE_AZURE_CLIENT_ID or VITE_AZURE_TENANT_ID in .env");
  }
  if (!pca) {
    pca = new PublicClientApplication({
      auth: {
        clientId: clientId!,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        redirectUri: window.location.origin,
      },
      cache: {
        cacheLocation: "localStorage",
      },
    });
  }
  return pca;
}

export async function initAuth(): Promise<AccountInfo | null> {
  if (!authConfigured) return null;
  const app = getMsal();
  await app.initialize();
  const result = await app.handleRedirectPromise();
  if (result?.account) {
    app.setActiveAccount(result.account);
    return result.account;
  }
  const existing = app.getActiveAccount() ?? app.getAllAccounts()[0];
  if (existing) {
    app.setActiveAccount(existing);
    return existing;
  }
  return null;
}

export async function signIn(): Promise<AccountInfo> {
  const app = getMsal();
  const result = await app.loginPopup({ scopes: graphScopes });
  app.setActiveAccount(result.account);
  return result.account;
}

export async function signOut(): Promise<void> {
  const app = getMsal();
  const account = app.getActiveAccount();
  if (account) {
    await app.logoutPopup({ account });
  }
}

export async function getToken(): Promise<string> {
  const app = getMsal();
  const account = app.getActiveAccount();
  if (!account) {
    throw new Error("Sign in with Microsoft first.");
  }
  let result: AuthenticationResult;
  try {
    result = await app.acquireTokenSilent({ account, scopes: graphScopes });
  } catch {
    result = await app.acquireTokenPopup({ account, scopes: graphScopes });
  }
  return result.accessToken;
}
