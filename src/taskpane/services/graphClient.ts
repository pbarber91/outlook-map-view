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
      },
      cache: {
        cacheLocation: "localStorage",
      },
    });
  }

  return msalInstance;
}

function getPreferredAccount(app: IPublicClientApplication): AccountInfo | null {
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

async function runInteractiveAuth(
  app: IPublicClientApplication,
  account?: AccountInfo | null
): Promise<AuthenticationResult> {
  if (interactivePromise) {
    return interactivePromise;
  }

  interactivePromise = withTimeout(
    account
      ? app.acquireTokenPopup({
          scopes: GRAPH_SCOPES,
          account,
        })
      : app.loginPopup({
          scopes: GRAPH_SCOPES,
        }),
    AUTH_TIMEOUT_MS,
    "Sign-in popup did not complete in time."
  ).finally(() => {
    interactivePromise = null;
  });

  return interactivePromise;
}

export async function completeRedirectIfNeeded(): Promise<boolean> {
  const app = await initMsal();
  const account = getPreferredAccount(app);
  return !!account;
}

export async function ensureGraphAccessInteractiveRedirect(): Promise<void> {
  const app = await initMsal();
  const existing = getPreferredAccount(app);

  try {
    if (existing) {
      await withTimeout(
        app.acquireTokenSilent({
          scopes: GRAPH_SCOPES,
          account: existing,
        }),
        AUTH_TIMEOUT_MS,
        "Silent token acquisition timed out."
      );
      return;
    }
  } catch (error) {
    if (!(error instanceof InteractionRequiredAuthError)) {
      if (isInteractionInProgressError(error)) {
        throw new Error(
          "Sign-in is already in progress. Finish the Microsoft sign-in window, then try again."
        );
      }
      throw error;
    }
  }

  await runInteractiveAuth(app, existing);
}

async function getAccessTokenInternal(): Promise<string> {
  const app = await initMsal();
  const existing = getPreferredAccount(app);

  if (existing) {
    try {
      const silentResult = await withTimeout(
        app.acquireTokenSilent({
          scopes: GRAPH_SCOPES,
          account: existing,
        }),
        AUTH_TIMEOUT_MS,
        "Silent token acquisition timed out."
      );

      return silentResult.accessToken;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        const popupResult = await runInteractiveAuth(app, existing);
        return popupResult.accessToken;
      }

      if (isInteractionInProgressError(error)) {
        throw new Error(
          "Sign-in is already in progress. Finish the Microsoft sign-in window, then try again."
        );
      }

      if (error instanceof BrowserAuthError && error.errorCode === "interaction_in_progress") {
        throw new Error(
          "Sign-in is already in progress. Finish the Microsoft sign-in window, then try again."
        );
      }

      throw error;
    }
  }

  const loginResult = await runInteractiveAuth(app, null);
  if (loginResult.accessToken) {
    return loginResult.accessToken;
  }

  const accountAfterLogin = getPreferredAccount(app);
  if (!accountAfterLogin) {
    throw new Error("Microsoft 365 sign-in completed, but no account is available.");
  }

  const silentAfterLogin = await withTimeout(
    app.acquireTokenSilent({
      scopes: GRAPH_SCOPES,
      account: accountAfterLogin,
    }),
    AUTH_TIMEOUT_MS,
    "Silent token acquisition timed out after sign-in."
  );

  return silentAfterLogin.accessToken;
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