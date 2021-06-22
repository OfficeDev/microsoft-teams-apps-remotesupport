/*
    <copyright file="app.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/


import * as React from "react";
import { AppRoute } from "./router/router";
import { Provider, themes } from "@fluentui/react";

export interface IAppState {
	theme: string;
	themeStyle: any;
}

export default class App extends React.Component<{}, IAppState> {

	constructor(props: any) {
		super(props);
		this.state = {
			theme: "",
			themeStyle: themes.teams,
		}
	}

	/**
	* Renders the component
	*/
	public render(): JSX.Element {

		return (
			<Provider theme={this.state.themeStyle}>
				<div className="appContainer">
					<AppRoute />
				</div>
			</Provider>
		);
	}

}
