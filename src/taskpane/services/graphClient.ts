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
let interactivePromise: Promise<AuthenticationResult> | null = null;

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

async function runInteractiveRequest(
  app: IPublicClientApplication
): Promise<AuthenticationResult> {
  if (interactivePromise) {
    return interactivePromise;
  }

  interactivePromise = (async () => {
    let account = getPreferredAccount(app);

    if (!account) {
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

      account = getPreferredAccount(app);
    }

    if (!account) {
      throw new Error("Sign-in completed, but no account was returned.");
    }

    const tokenResult = await withTimeout(
      app.acquireTokenPopup({
        scopes: GRAPH_SCOPES,
        account,
        redirectUri: getRedirectUri(),
      }),
      AUTH_TIMEOUT_MS,
      "Token popup did not complete in time."
    );

    if (tokenResult.account) {
      app.setActiveAccount(tokenResult.account);
    }

    return tokenResult;
  })();

  try {
    return await interactivePromise;
  } finally {
    interactivePromise = null;
  }
}

async function getAccessTokenInternal(): Promise<string> {
  const app = await initMsal();
  await completePendingRedirect(app);

  const account = getPreferredAccount(app);

  if (account) {
    try {
      const silentResult = await withTimeout(
        app.acquireTokenSilent({
          scopes: GRAPH_SCOPES,
          account,
          redirectUri: getRedirectUri(),
        }),
        AUTH_TIMEOUT_MS,
        "Silent token acquisition timed out."
      );

      if (silentResult.account) {
        app.setActiveAccount(silentResult.account);
      }

      return silentResult.accessToken;
    } catch (error) {
      if (!(error instanceof InteractionRequiredAuthError)) {
        throw error;
      }
    }
  }

  try {
    const interactiveResult = await runInteractiveRequest(app);
    return interactiveResult.accessToken;
  } catch (error) {
    if (isInteractionInProgressError(error) && interactivePromise) {
      const result = await interactivePromise;
      return result.accessToken;
    }

    if (error instanceof BrowserAuthError && error.errorCode === "interaction_in_progress") {
      throw new Error("Sign-in is already in progress. Please wait for the sign-in window to finish.");
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