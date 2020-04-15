// <copyright file="incident-api.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import axios from "./axios-decorator";

const baseAxiosUrl = window.location.origin;

/**
* Get localized resource strings from API
*/
export const getResourceStrings = async (token: string): Promise<any> => {

    let url = baseAxiosUrl + "/api/resource/resourcestrings";
    let resourceStringsResponse = await axios.get(url, token);
    return resourceStringsResponse;
}

/**
* Get card configuration saved in storage.
* @param  {String | Null} token Custom JWT token.
*/
export const getConfigurationsAsync = async (token: string): Promise<any> => {

    let url = baseAxiosUrl + "/api/Storage";
    let configurationsResponse = await axios.get(url, token);
    return configurationsResponse;
}

/**
* Get Azure Active Directory settings for authentication.
*/
export const getAzureActiveDirectorySettingsAsync = async (): Promise<any> => {

    let url = baseAxiosUrl + "/api/Settings";
    let settingsResponse = await axios.get(url, "");
    if (settingsResponse.status === 200) {
        localStorage.setItem("TenantId", settingsResponse.data.TenantId);
        localStorage.setItem("ClientId", settingsResponse.data.ClientId);
        localStorage.setItem("TokenEndpoint", settingsResponse.data.TokenEndpoint);
    }
    return true;
}

/**
* Save card configuration details in Azure storage.
* @param  {configurationDetails | Null} configurationDetails configuration details.
* @param  {String | Null} token Custom JWT token.
*/
export const saveConfigurationsAsync = async (configurationDetails: any, token: string): Promise<any> => {

    let url = baseAxiosUrl + "/api/Storage";
    let saveConfigurationsResult = await axios.post(url, configurationDetails, token);
    return saveConfigurationsResult;
}

export const getAzureActiveDirectorySettings = async () => {
    await getAzureActiveDirectorySettingsAsync();
};
