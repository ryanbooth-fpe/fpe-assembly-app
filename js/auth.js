const msalConfig = {
    auth: {
        clientId: 'f3e17927-201a-4e5a-a71d-9dcee3193bb4',
        authority: 'https://login.microsoftonline.com/eb6d316c-6482-464f-8290-f2ddee200fea',
        redirectUri: window.location.origin
    },
    cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false
    }
};

const loginRequest = {
    scopes: ['User.Read', 'Sites.ReadWrite.All']
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

async function getToken() {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) throw new Error('Not signed in');

    try {
        const response = await msalInstance.acquireTokenSilent({
            ...loginRequest,
            account: accounts[0]
        });
        return response.accessToken;
    } catch (error) {
        if (error instanceof msal.InteractionRequiredAuthError) {
            const response = await msalInstance.acquireTokenPopup(loginRequest);
            return response.accessToken;
        }
        throw error;
    }
}

async function getCurrentUser() {
    const accounts = msalInstance.getAllAccounts();
    return accounts.length > 0 ? accounts[0] : null;
}

async function login() {
    const response = await msalInstance.loginPopup(loginRequest);
    return response.account;
}

async function logout() {
    await msalInstance.logoutPopup({ postLogoutRedirectUri: window.location.origin });
}
