/*
    <copyright file="adal-config.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import { AuthenticationContext, adalFetch, withAdalLogin, AdalConfig } from 'react-adal';
import { getAzureActiveDirectorySettings } from "./api/incident-api";

getAzureActiveDirectorySettings();

export const adalConfig: AdalConfig = {
    tenant: localStorage.getItem("TenantId")!,
    clientId: localStorage.getItem("ClientId")!,
    endpoints: {
        api: localStorage.getItem("TokenEndpoint")!,
    },
    postLogoutRedirectUri: window.location.origin,
    cacheLocation: 'localStorage'
};

export const authContext = new AuthenticationContext(adalConfig);
export const getToken = () => authContext.getCachedToken(adalConfig.clientId);
export const adalApiFetch = (fetch:any, url:any, options:any) =>
    adalFetch(authContext, adalConfig!.endpoints!.api, fetch, url, options);

export const withAdalLoginApi = withAdalLogin(authContext, adalConfig!.endpoints!.api);