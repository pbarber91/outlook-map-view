import {
  createNestablePublicClientApplication,
  InteractionRequiredAuthError,
  BrowserAuthError,
  IPublicClientApplication,
  AuthenticationResult,
  AccountInfo,
} from "@azure/msal-browser";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";

declare const __AZURE_CLIENT_ID__: string;
declare const __AZURE_TENANT_ID__: string;

const fallbackClientId = "45f4ed01-b835-4aa3-b143-8606bcb85d60";
const GRAPH_SCOPES = ["User.Read", "Calendars.Read"];
const AUTH_TIMEOUT_MS = 30000;

let msalInstance: IPublicClientApplication | undefined;
let accessTokenPromise: Promise<string> | null = null;
let interactivePromise: Promise<void> | null = null;

function getClientId(): string {
  const value =
    typeof __AZURE_CLIENT_ID__ === "string" && __AZURE_CLIENT_ID__.trim().length > 0
      ? __AZURE_CLIENT_ID__.trim()
      : fallbackClientId;

  return value;
}

function getTenantAuthority(): string {
  const tenant =
    typeof __AZURE_TENANT_ID__ === "string" && __AZURE_TENANT_ID__.trim().length > 0
      ? __AZURE_TENANT_ID__.trim()
      : "common";

  return `https://login.microsoftonline.com/${tenant}`;
}

function getRedirectUri(): string {
  return `${window.location.origin}/auth.html`;
}

function withTimeout<T>(promise: Promise<T>, ms: number, message: string): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = window.setTimeout(() => reject(new Error(message)), ms);

    promise
      .then((value) => {
        window.clearTimeout(timer);
        resolve(value);
      })
      .catch((error) => {
        window.clearTimeout(timer);
        reject(error);
      });
  });
}

async function initMsal(): Promise<IPublicClientApplication> {
  if (!msalInstance) {
    msalInstance = await createNestablePublicClientApplication({
      auth: {
        clientId: getClientId(),
        authority: getTenantAuthority(),
        redirectUri: getRedirectUri(),
      },
      cache: {
        cacheLocation: "localStorage",
      },
    });
  }

  return msalInstance;
}

function getPreferredAccount(app: IPublicClientApplication): AccountInfo | null {
  const active = app.getActiveAccount();
  if (active) {
    return active;
  }

  const accounts = app.getAllAccounts();
  return accounts.length > 0 ? accounts[0] : null;
}

async function completePendingRedirect(app: IPublicClientApplication): Promise<void> {
  const redirectResult = await withTimeout(
    app.handleRedirectPromise(),
    AUTH_TIMEOUT_MS,
    "Authentication redirect did not complete in time."
  );

  if (redirectResult?.account) {
    app.setActiveAccount(redirectResult.account);
    return;
  }

  const existing = getPreferredAccount(app);
  if (existing) {
    app.setActiveAccount(existing);
  }
}

function isInteractionInProgressError(error: unknown): boolean {
  if (!error || typeof error !== "object") {
    return false;
  }

  const maybe = error as { errorCode?: string; message?: string };
  return (
    maybe.errorCode === "interaction_in_progress" ||
    maybe.message?.includes("interaction_in_progress") === true
  );
}

async function ensureSignedIn(app: IPublicClientApplication): Promise<void> {
  const existing = getPreferredAccount(app);
  if (existing) {
    app.setActiveAccount(existing);
    return;
  }

  if (interactivePromise) {
    return interactivePromise;
  }

  interactivePromise = (async () => {
    const loginResult = await withTimeout(
      app.loginPopup({
        scopes: GRAPH_SCOPES,
        redirectUri: getRedirectUri(),
      }),
      AUTH_TIMEOUT_MS,
      "Sign-in popup did not complete in time."
    );

    if (loginResult.account) {
      app.setActiveAccount(loginResult.account);
    }
  })();

  try {
    await interactivePromise;
  } finally {
    interactivePromise = null;
  }
}

async function acquireTokenSilentlyOrInteractive(app: IPublicClientApplication): Promise<AuthenticationResult> {
  const account = getPreferredAccount(app);

  if (!account) {
    await ensureSignedIn(app);
  }

  const resolvedAccount = getPreferredAccount(app);
  if (!resolvedAccount) {
    throw new Error("Sign-in completed, but no account was returned.");
  }

  try {
    const silentResult = await withTimeout(
      app.acquireTokenSilent({
        scopes: GRAPH_SCOPES,
        account: resolvedAccount,
        redirectUri: getRedirectUri(),
      }),
      AUTH_TIMEOUT_MS,
      "Silent token acquisition timed out."
    );

    if (silentResult.account) {
      app.setActiveAccount(silentResult.account);
    }

    return silentResult;
  } catch (error) {
    if (!(error instanceof InteractionRequiredAuthError)) {
      throw error;
    }
  }

  if (interactivePromise) {
    await interactivePromise;
  } else {
    interactivePromise = (async () => {
      const refreshed = await withTimeout(
        app.loginPopup({
          scopes: GRAPH_SCOPES,
          redirectUri: getRedirectUri(),
        }),
        AUTH_TIMEOUT_MS,
        "Sign-in popup did not complete in time."
      );

      if (refreshed.account) {
        app.setActiveAccount(refreshed.account);
      }
    })();

    try {
      await interactivePromise;
    } finally {
      interactivePromise = null;
    }
  }

  const retryAccount = getPreferredAccount(app);
  if (!retryAccount) {
    throw new Error("Authentication completed, but no account is available.");
  }

  const retrySilent = await withTimeout(
    app.acquireTokenSilent({
      scopes: GRAPH_SCOPES,
      account: retryAccount,
      redirectUri: getRedirectUri(),
    }),
    AUTH_TIMEOUT_MS,
    "Silent token acquisition timed out after sign-in."
  );

  if (retrySilent.account) {
    app.setActiveAccount(retrySilent.account);
  }

  return retrySilent;
}

async function getAccessTokenInternal(): Promise<string> {
  const app = await initMsal();
  await completePendingRedirect(app);

  try {
    const result = await acquireTokenSilentlyOrInteractive(app);
    return result.accessToken;
  } catch (error) {
    if (isInteractionInProgressError(error)) {
      throw new Error("Sign-in is already in progress. Close any old sign-in popups, then refresh and try again.");
    }

    if (error instanceof BrowserAuthError && error.errorCode === "interaction_in_progress") {
      throw new Error("Sign-in is already in progress. Close any old sign-in popups, then refresh and try again.");
    }

    throw error;
  }
}

async function getAccessToken(): Promise<string> {
  if (!accessTokenPromise) {
    accessTokenPromise = getAccessTokenInternal().finally(() => {
      accessTokenPromise = null;
    });
  }

  return accessTokenPromise;
}

export async function completeRedirectIfNeeded(): Promise<void> {
  const app = await initMsal();
  await completePendingRedirect(app);
}

export async function ensureGraphAccessInteractiveRedirect(): Promise<void> {
  await getAccessToken();
}

export async function getGraphClient(): Promise<Client> {
  const accessToken = await getAccessToken();

  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });
}