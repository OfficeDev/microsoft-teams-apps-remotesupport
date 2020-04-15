/*
    <copyright file="app.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import * as React from "react";
import ReactDOM from "react-dom";
import AppRoute from "./router/router";
import { runWithAdal } from 'react-adal';
import { authContext } from './adal-config';

export default class App extends React.Component<{}, {}> {
    constructor(props: any) {
        super(props);
    }

    render(): JSX.Element {
        return (
            <div className="appContainer">
                <AppRoute />
            </div>
        );
    }
}

/* renders the component */
const DO_NOT_LOGIN = false;

runWithAdal(authContext, () => {
    ReactDOM.render(
        <App />, document.getElementById("container"));
}, DO_NOT_LOGIN);