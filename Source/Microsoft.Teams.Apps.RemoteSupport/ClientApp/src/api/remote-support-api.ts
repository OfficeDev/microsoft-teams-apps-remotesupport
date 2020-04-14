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
    return teamMemberResponse;
}

/**
* Get on call support experts details.
*/
export const getOnCallExpertsInTeam = async (token: string): Promise<any> => {

    let url = baseAxiosUrl + "/api/remotesupport/oncallexperts";
    let onCallSupportResponse = await axios.get(url, token);
    return onCallSupportResponse;
}

/**
* Get localized resource strings from API
*/
export const getResourceStrings = async (token: string, locale?: string | null): Promise<any> => {

    let url = baseAxiosUrl + "/api/resource/resourcestrings";
    let resourceStringsResponse = await axios.get(url, token, locale);
    return resourceStringsResponse;
}

/**
* Save on call support details.
* @param  {OnCallSupportDetail | Null} onCallSupportDetails On call support details.
* @param  {String | Null} token Custom JWT token.
*/
export const saveOnCallSupportDetails = async (onCallSupportDetails: any, token: string): Promise<any> => {

    let url = baseAxiosUrl + "/api/remotesupport/saveoncallsupportdetails";
    let saveOnCallSupportResponse = await axios.post(url, onCallSupportDetails, token);
    return saveOnCallSupportResponse;
}

/**
* Handle error occurred during API call.
* @param  {Object} error Error response object
*/
export const handleError = (error: any, token: any): any => {
	const errorStatus = error.status;
	if (errorStatus === 403) {
        window.location.href = "/error?code=403&token=" + token;
    }
    else if (errorStatus === 401) {
        window.location.href = "/error?code=401&token=" + token;
    }
    else {
        window.location.href = "/error?token=" + token;
	}
}
