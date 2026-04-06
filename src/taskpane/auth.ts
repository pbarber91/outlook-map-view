import { PublicClientApplication } from "@azure/msal-browser";

declare const __AZURE_CLIENT_ID__: string;
declare const __AZURE_TENANT_ID__: string;

const fallbackClientId = "45f4ed01-b835-4aa3-b143-8606bcb85d60";

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

function showError(message: string): void {
  const el = document.getElementById("auth-error");
  if (el) {
    el.textContent = message;
  }
}

async function run(): Promise<void> {
  try {
    const app = new PublicClientApplication({
      auth: {
        clientId: getClientId(),
        authority: getTenantAuthority(),
        redirectUri: `${window.location.origin}/auth.html`,
      },
      cache: {
        cacheLocation: "localStorage",
      },
    });

    await app.initialize();
    const result = await app.handleRedirectPromise();

    if (result?.account) {
      app.setActiveAccount(result.account);
    }

    window.close();

    window.setTimeout(() => {
      showError("Sign-in completed. You can close this window.");
    }, 500);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Authentication failed.";
    showError(message);
    console.error("auth callback failed", error);
  }
}

void run();