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
			return error.response;
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
			return error.response;
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
			return error.response;
		}
	}
}

const axiosJWTDecoratorInstance = new AxiosJWTDecorator();
export default axiosJWTDecoratorInstance;