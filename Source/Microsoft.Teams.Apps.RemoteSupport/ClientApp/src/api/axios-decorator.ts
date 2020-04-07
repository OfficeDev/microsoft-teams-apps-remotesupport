/*
    <copyright file="axios-decorator.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import axios, { AxiosResponse, AxiosRequestConfig } from "axios";

export class AxiosJWTDecorator {

	/**
	* Post data to API
	* @param  {String} url Resource URI
	* @param  {Object} data Request body data
	* @param  {String} token Custom JWT token
	*/
	public async post<T = any, R = AxiosResponse<T>>(
		url: string,
		data?: any,
		token?: string
	): Promise<R> {
        try {
            let config: AxiosRequestConfig = axios.defaults;
            config.headers["Authorization"] = `Bearer ${token}`;

            return await axios.post(url, data, config);
        } catch (error) {
            this.handleError(error);
            throw error;
        }                                                                                                                                                                                               
	}

	/**
	* Post data to API
	* @param  {String} url Resource URI
	* @param  {Object} data Request body data
	* @param  {String} token Custom JWT token
	*/
	public async Put<T = any, R = AxiosResponse<T>>(
		url: string,
		data?: any,
		token?: string
	): Promise<R> {
		try {
			let config: AxiosRequestConfig = axios.defaults;
			config.headers["Authorization"] = `Bearer ${token}`;

			return await axios.put(url, data, config);
		} catch (error) {
			this.handleError(error);
			throw error;
		}
	}

	/**
	* Get data to API
	* @param  {String} token Custom JWT token
	*/
	public async get<T = any, R = AxiosResponse<T>>(
		url: string,
        token?: string,
        locale?: string | null
	): Promise<R> {
		try {
			let config: AxiosRequestConfig = axios.defaults;
            config.headers["Authorization"] = `Bearer ${token}`;
            if (locale) {
                config.headers["Accept-Language"] = `${locale}`;
            }
            return await axios.get(url, config);
		} catch (error) {
			this.handleError(error);
			throw error;
		}
	}

	/**
	* Handle error occurred during API call.
	* @param  {Object} error Error response object
	*/
	private handleError(error: any): void {
		if (error.hasOwnProperty("response")) {
			const errorStatus = error.response.status;
			if (errorStatus === 403) {
				window.location.href = "/error/403";
			} else if (errorStatus === 401) {
				window.location.href = "/error/401";
			} else {
				window.location.href = "/error";
			}
		} else {
			window.location.href = "/error";
		}
	}
}

const axiosJWTDecoratorInstance = new AxiosJWTDecorator();
export default axiosJWTDecoratorInstance;