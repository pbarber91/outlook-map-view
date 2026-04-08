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

function getClientId(): string {
  return typeof __AZURE_CLIENT_ID__ === "string" && __AZURE_CLIENT_ID__.trim().length > 0
    ? __AZURE_CLIENT_ID__.trim()
    : fallbackClientId;
}

function getTenantAuthority(): string {
  const tenant =
    typeof __AZURE_TENANT_ID__ === "string" && __AZURE_TENANT_ID__.trim().length > 0
      ? __AZURE_TENANT_ID__.trim()
      : "common";

  return `https://login.microsoftonline.com/${tenant}`;
}

function isStandaloneWindow(): boolean {
  if (typeof window === "undefined") {
    return false;
  }

  return new URLSearchParams(window.location.search).get("standalone") === "1";
}

function getStandaloneRedirectUri(): string {
  return `${window.location.origin}/taskpane.html?standalone=1`;
}

function getPopupRedirectUri(): string {
  return `${window.location.origin}/popup-complete.html`;
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
        redirectUri: isStandaloneWindow() ? getStandaloneRedirectUri() : getPopupRedirectUri(),
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

export async function completeRedirectIfNeeded(): Promise<boolean> {
  const app = await initMsal();

  const result = await app.handleRedirectPromise();
  if (result?.account) {
    app.setActiveAccount(result.account);
  }

  const account = getPreferredAccount(app);
  if (account) {
    app.setActiveAccount(account);
    return true;
  }

  return false;
}

async function getAccessTokenViaPopup(app: IPublicClientApplication, account: AccountInfo): Promise<AuthenticationResult> {
  return withTimeout(
    app.acquireTokenPopup({
      scopes: GRAPH_SCOPES,
      account,
      redirectUri: getPopupRedirectUri(),
    }),
    AUTH_TIMEOUT_MS,
    "Sign-in popup did not complete in time."
  );
}

async function getAccessTokenViaRedirect(app: IPublicClientApplication, account?: AccountInfo): Promise<never> {
  if (account) {
    await app.acquireTokenRedirect({
      scopes: GRAPH_SCOPES,
      account,
      redirectUri: getStandaloneRedirectUri(),
    });
  } else {
    await app.loginRedirect({
      scopes: GRAPH_SCOPES,
      redirectUri: getStandaloneRedirectUri(),
    });
  }

  throw new Error("Redirecting to Microsoft 365 sign-in...");
}

export async function ensureGraphAccessInteractiveRedirect(): Promise<void> {
  const app = await initMsal();
  const account = getPreferredAccount(app);

  if (account) {
    app.setActiveAccount(account);

    try {
      await withTimeout(
        app.acquireTokenSilent({
          scopes: GRAPH_SCOPES,
          account,
        }),
        AUTH_TIMEOUT_MS,
        "Silent token acquisition timed out."
      );
      return;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        await getAccessTokenViaRedirect(app, account);
        return;
      }

      throw error;
    }
  }

  await getAccessTokenViaRedirect(app);
}

async function getAccessTokenInternal(): Promise<string> {
  const app = await initMsal();

  const existing = getPreferredAccount(app);
  if (existing) {
    app.setActiveAccount(existing);

    try {
      const silentResult = await withTimeout(
        app.acquireTokenSilent({
          scopes: GRAPH_SCOPES,
          account: existing,
        }),
        AUTH_TIMEOUT_MS,
        "Silent token acquisition timed out."
      );

      if (silentResult.account) {
        app.setActiveAccount(silentResult.account);
      }

      return silentResult.accessToken;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        if (isStandaloneWindow()) {
          await getAccessTokenViaRedirect(app, existing);
        }

        const popupResult = await getAccessTokenViaPopup(app, existing);
        if (popupResult.account) {
          app.setActiveAccount(popupResult.account);
        }
        return popupResult.accessToken;
      }

      if (isInteractionInProgressError(error)) {
        throw new Error("Sign-in is already in progress. Finish the Microsoft sign-in window, then try again.");
      }

      throw error;
    }
  }

  if (isStandaloneWindow()) {
    await getAccessTokenViaRedirect(app);
  }

  const loginResult = await withTimeout(
    app.loginPopup({
      scopes: GRAPH_SCOPES,
      redirectUri: getPopupRedirectUri(),
    }),
    AUTH_TIMEOUT_MS,
    "Sign-in popup did not complete in time."
  );

  if (loginResult.account) {
    app.setActiveAccount(loginResult.account);
  }

  return loginResult.accessToken;
}

async function getAccessToken(): Promise<string> {
  if (!accessTokenPromise) {
    accessTokenPromise = getAccessTokenInternal().finally(() => {
      accessTokenPromise = null;
    });
  }

  return accessTokenPromise;
}

export async function getGraphClient(): Promise<Client> {
  const accessToken = await getAccessToken();

  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });
}
