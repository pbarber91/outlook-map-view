import {
  createNestablePublicClientApplication,
  InteractionRequiredAuthError,
  IPublicClientApplication,
  AuthenticationResult,
  AccountInfo,
} from "@azure/msal-browser";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";

declare const __AZURE_CLIENT_ID__: string;

const fallbackClientId = "45f4ed01-b835-4aa3-b143-8606bcb85d60";

function getClientId(): string {
  const value =
    typeof __AZURE_CLIENT_ID__ === "string" && __AZURE_CLIENT_ID__.trim().length > 0
      ? __AZURE_CLIENT_ID__.trim()
      : fallbackClientId;

  return value;
}

let msalInstance: IPublicClientApplication | undefined;

async function initMsal(): Promise<IPublicClientApplication> {
  if (!msalInstance) {
    msalInstance = await createNestablePublicClientApplication({
      auth: {
        clientId: getClientId(),
        authority: "https://login.microsoftonline.com/common",
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

async function acquireSilentOrInteractiveToken(): Promise<AuthenticationResult> {
  const app = await initMsal();
  const account = getPreferredAccount(app);

  const tokenRequest = {
    scopes: ["User.Read", "Calendars.Read"],
    account: account ?? undefined,
  };

  try {
    const silentResult = await app.acquireTokenSilent(tokenRequest);

    if (silentResult.account) {
      app.setActiveAccount(silentResult.account);
    }

    return silentResult;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      const popupResult = await app.acquireTokenPopup({
        scopes: ["User.Read", "Calendars.Read"],
      });

      if (popupResult.account) {
        app.setActiveAccount(popupResult.account);
      }

      return popupResult;
    }

    throw error;
  }
}

async function getAccessToken(): Promise<string> {
  const result = await acquireSilentOrInteractiveToken();
  return result.accessToken;
}

export async function getGraphClient(): Promise<Client> {
  const accessToken = await getAccessToken();

  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });
}