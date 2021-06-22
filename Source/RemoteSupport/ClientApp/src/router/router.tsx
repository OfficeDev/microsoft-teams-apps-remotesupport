/*
    <copyright file="router.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import * as React from "react";
import { BrowserRouter, Route, Switch } from "react-router-dom";
import ManageExperts from '../components/manage-experts';
import ErrorPage from '../components/error-page';

export const AppRoute: React.FunctionComponent<{}> = () => {
	return (
		<BrowserRouter>
			<Switch>
				<Route path='/manage-experts' component={ManageExperts} />
				<Route exact path="/error" component={ErrorPage} />
			</Switch>
		</BrowserRouter>
	);
};