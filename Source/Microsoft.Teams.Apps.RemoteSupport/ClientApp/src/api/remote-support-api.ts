/*
    <copyright file="remote-support-api.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import axios from "./axios-decorator";
import { AxiosResponse } from "axios";

const baseAxiosUrl = window.location.origin;

/**
* Get all team members.
* @param  {String} teamId Team ID for getting members
* @param  {String | Null} token Custom JWT token
*/
export const getMembersInTeam = async (token: string): Promise<any> => {

    let url = baseAxiosUrl + "/api/remotesupport/teammembers";
    let teamMemberResponse = await axios.get(url, token);
    if (teamMemberResponse.status === 401) {
        redirectToErrorPage(teamMemberResponse, token);
    }
    else {
        return teamMemberResponse;
    }
}

/**
* Get on call support experts details.
*/
export const getOnCallExpertsInTeam = async (token: string): Promise<any> => {

    let url = baseAxiosUrl + "/api/remotesupport/oncallexperts";
    let onCallSupportResponse = await axios.get(url, token);
    if (onCallSupportResponse.status === 401) {
        redirectToErrorPage(onCallSupportResponse, token);
    }
    else {
        return onCallSupportResponse;
    }
}

/**
* Get localized resource strings from API
*/
export const getResourceStrings = async (token: string, locale?: string | null): Promise<any> => {

    let url = baseAxiosUrl + "/api/resource/resourcestrings";
    let resourceStringsResponse = await axios.get(url, token, locale);
    if (resourceStringsResponse.status === 401) {
        redirectToErrorPage(resourceStringsResponse, token);
    }
    else {
        return resourceStringsResponse;
    }
}

/**
* Save on call support details.
* @param  {OnCallSupportDetail | Null} onCallSupportDetails On call support details.
* @param  {String | Null} token Custom JWT token.
*/
export const saveOnCallSupportDetails = async (onCallSupportDetails: any, token: string): Promise<any> => {

    let url = baseAxiosUrl + "/api/remotesupport/saveoncallsupportdetails";
    let saveOnCallSupportResponse = await axios.post(url, onCallSupportDetails, token);
    if (saveOnCallSupportResponse.status === 401) {
        redirectToErrorPage(saveOnCallSupportResponse, token);
    }
    else {
        return saveOnCallSupportResponse;
    }
}

const redirectToErrorPage = (response: AxiosResponse<any>, token: string) => {
    if (response.data) {
        window.location.href = "/error?code=" + response.data.code + "&token=" + token;
    }
    else {
        window.location.href = "/error?token=" + token;
    }
}