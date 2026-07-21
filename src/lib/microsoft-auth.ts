import { PublicClientApplication, Configuration, AuthenticationResult, AccountInfo } from "@azure/msal-browser";

const msalConfig: Configuration = {
  auth: {
    clientId: "96a4ec44-e1d7-4b8c-a3dd-7f725865598a",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin + window.location.pathname,
  },
  cache: {
    cacheLocation: "sessionStorage"
  }
};

const msalInstance = new PublicClientApplication(msalConfig);
let isInitialized = false;
let initializePromise: Promise<void> | null = null;
let activeAccount: AccountInfo | null = null;
let cachedAccessToken: string | null = null;

const initializeMsal = async () => {
  if (isInitialized) return;
  if (!initializePromise) {
    initializePromise = (async () => {
      try {
        await msalInstance.initialize();
        const response = await msalInstance.handleRedirectPromise();
        if (response) {
          activeAccount = response.account;
          msalInstance.setActiveAccount(activeAccount);
          cachedAccessToken = response.accessToken;
        }
      } catch (e) {
        console.warn("MSAL initialization/redirect warning:", e);
      }
      isInitialized = true;
      
      // Check if there are any accounts already signed in
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        activeAccount = accounts[0];
        msalInstance.setActiveAccount(activeAccount);
        
        try {
          const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: ["Mail.Send", "User.Read"],
            account: activeAccount
          });
          cachedAccessToken = tokenResponse.accessToken;
        } catch (err) {
          console.error("Silent token acquisition failed", err);
          activeAccount = null;
          cachedAccessToken = null;
        }
      }
    })();
  }
  await initializePromise;
};

export const initMicrosoftAuth = async (
  onAuthSuccess?: (user: any, token: string) => void,
  onAuthFailure?: () => void
) => {
  await initializeMsal();
  if (activeAccount && cachedAccessToken) {
    if (onAuthSuccess) {
      onAuthSuccess({ email: activeAccount.username, name: activeAccount.name }, cachedAccessToken);
    }
  } else {
    if (onAuthFailure) onAuthFailure();
  }
};

export const microsoftSignIn = async (): Promise<{ user: any; accessToken: string } | null> => {
  await initializeMsal();
  try {
    const loginRequest = {
      scopes: ["Mail.Send", "User.Read"],
      prompt: "select_account"
    };
    
    const response = await msalInstance.loginPopup(loginRequest);
    if (response) {
      activeAccount = response.account;
      msalInstance.setActiveAccount(activeAccount);
      cachedAccessToken = response.accessToken;
      return {
        user: { email: activeAccount.username, name: activeAccount.name },
        accessToken: cachedAccessToken
      };
    }
    return null;
  } catch (error) {
    console.error("Microsoft sign in error:", error);
    throw error;
  }
};

export const getMicrosoftAccessToken = async (): Promise<string | null> => {
  return cachedAccessToken;
};

export const microsoftLogout = async () => {
  await initializeMsal();
  if (activeAccount) {
    await msalInstance.logoutRedirect({
      account: activeAccount
    });
  }
  activeAccount = null;
  cachedAccessToken = null;
};
