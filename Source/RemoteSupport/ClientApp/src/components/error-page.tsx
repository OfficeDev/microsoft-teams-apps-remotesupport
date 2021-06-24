/*
    <copyright file="error-page.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import * as React from "react";
import { Text, Flex, Icon, Provider, themes } from "@fluentui/react";
import * as microsoftTeams from "@microsoft/teams-js";
import { getResourceStrings } from "../api/remote-support-api";

interface errorPageState {
    theme: string | "",
    themeStyle: any | "",
    resourceStrings: any | "",
}

const DarkTheme = "dark";
const ContrastTheme = "contrast";

export default class ErrorPage extends React.Component<{}, errorPageState> {
    code: string | null = null;
    token: string | null = null;
    locale?: string | null;

    constructor(props: any) {
        super(props);
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.token = params.get("token");
        this.code = params.get("code");
        this.state = {
            theme: "",
            themeStyle: themes.teams,
            resourceStrings: {}
        };
    }

    /** Called once component is mounted. */
    async componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            let theme = context.theme || "";
            this.updateTheme(theme);
            this.locale = context.locale;
            this.setState({
                theme: theme
            });
            this.getResourceStrings();
        });

        microsoftTeams.registerOnThemeChangeHandler((theme) => {
            this.updateTheme(theme);
            this.setState({
                theme: theme,
            }, () => {
                this.forceUpdate();
            });
        });
    }

    /**
	* Set current theme state received from teams context
	* @param  {String} theme Current theme name
	*/
    private updateTheme = (theme: string) => {
        if (theme === DarkTheme) {
            this.setState({
                themeStyle: themes.teamsDark
            });
        } else if (theme === ContrastTheme) {
            this.setState({
                themeStyle: themes.teamsHighContrast
            });
        } else {
            this.setState({
                themeStyle: themes.teams
            });
        }
    }

    /** 
  *  Get resource strings according to user locale.
  * */
    getResourceStrings = async () => {
        const resourceStringsResponse = await getResourceStrings(this.token!, this.locale);

        if (resourceStringsResponse.status === 200) {
            this.setState({ resourceStrings: resourceStringsResponse.data });
        }
    }

    render() {
        let message = this.state.resourceStrings.errorMessage;
        if (this.code === "401") {
            message = `${this.state.resourceStrings.unauthorizedAccess}`;
        }
        return (
            <Provider theme={this.state.themeStyle}>
                <div className="container-div">
                    <Flex gap="gap.small" hAlign="center" vAlign="center" className="error-container">
                        <Flex gap="gap.small" hAlign="center" vAlign="center">
                            <Flex.Item>
                                <div
                                    style={{
                                        position: "relative",
                                    }}
                                >
                                    <Icon outline color="red" name="error" />
                                </div>
                            </Flex.Item>

                            <Flex.Item grow>
                                <Flex column gap="gap.small" vAlign="stretch">
                                    <div>
                                        <Text weight="bold" error content={message} /><br />
                                    </div>
                                </Flex>
                            </Flex.Item>
                        </Flex>
                    </Flex>
                </div>
            </Provider>
        );
    }
}