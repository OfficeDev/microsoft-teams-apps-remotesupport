/*
    <copyright file="router.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import * as React from "react";
import { BrowserRouter, Route, Switch } from "react-router-dom";
import Home from "../components/home";

const AppRoute = () => {
    return (
        <BrowserRouter>
            <Switch>
                <Route exact path="/" component={Home} />
            </Switch>
        </BrowserRouter>
    );
}
export default AppRoute;
