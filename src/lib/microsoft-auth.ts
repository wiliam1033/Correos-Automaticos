import { PublicClientApplication, Configuration, AuthenticationResult, AccountInfo } from "@azure/msal-browser";

const msalConfig: Configuration = {
  auth: {
    clientId: "96a4ec44-e1d7-4b8c-a3dd-7f725865598a",
    authority: "https://login.microsoftonline.com/8b75b7be-1c8e-45c0-bedc-10aaf1377ed6",
    redirectUri: window.location.origin,
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
        await msalInstance.handleRedirectPromise();
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
      scopes: ["Mail.Send", "User.Read"]
    };
    const response: AuthenticationResult = await msalInstance.loginPopup(loginRequest);
    activeAccount = response.account;
    msalInstance.setActiveAccount(activeAccount);
    cachedAccessToken = response.accessToken;
    
    return {
      user: {
        email: response.account?.username,
        name: response.account?.name
      },
      accessToken: response.accessToken
    };
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
    await msalInstance.logoutPopup({
      account: activeAccount
    });
  }
  activeAccount = null;
  cachedAccessToken = null;
};
