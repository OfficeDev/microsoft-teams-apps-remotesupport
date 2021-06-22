/*
    <copyright file="index.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import React from "react";
import ReactDOM from "react-dom";
import { BrowserRouter as Router } from "react-router-dom";
import App from "./app";
import { getAzureActiveDirectorySettingsAsync } from "./api/incident-api";

ReactDOM.render(
    <Router>
        <App />
    </Router>, document.getElementById("root"));
