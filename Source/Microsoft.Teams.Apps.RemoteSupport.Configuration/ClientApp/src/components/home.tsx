/*
    <copyright file="home.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import * as React from "react";
import { Text, Input, Provider, Flex, themes, Button, Segment, Loader } from '@fluentui/react';
import "../styles/theme.css";
import { getResourceStrings, getConfigurationsAsync } from "../api/incident-api";
import BuildYourForm from '../components/build-form-action';
import { getToken, authContext } from '../adal-config';


/** State interface. */
interface IState {
    TeamId: string,
    isAuthenticated: boolean | null,
    isError: boolean | null,
    CardItems: Array<any>,
    Message: string,
    MessageType: boolean,
    resourceStrings: any,
    resourceStringsLoaded: boolean,
}


/** Component for displaying home page of incident report configuration application. */
class Home extends React.Component<{}, IState>
{
    state: IState;
    bearer: string = "";
    /**
     * Constructor to initialize component.
     * @param props Props of component.
     */
    constructor(props: {}) {
        super(props);
        this.state = {
            TeamId: "",
            isAuthenticated: null,
            isError: null,
            CardItems: [],
            Message: "",
            MessageType: false,
            resourceStrings: {},
            resourceStringsLoaded: false,
        };

        this.bearer = getToken();
    }

    async componentDidMount() {
        this.getResourceStrings();
        this.getConfigurationsAsync();
    }

/** 
*  Get resource strings according to user locale.
* */
    getResourceStrings = async () => {
        const resourceStringsResponse = await getResourceStrings(this.bearer!);
        if (resourceStringsResponse) {
            if (resourceStringsResponse.status === 200) {
                this.setState({ resourceStrings: resourceStringsResponse.data });
            }
            this.setState({ resourceStringsLoaded: true });   
        }
    }

    getConfigurationsAsync = async () => {
        const configurationsResponse = await getConfigurationsAsync(this.bearer);
        if (configurationsResponse.status === 401) {
            this.setState({ isAuthenticated: false, isError: false });
        }
        else if (configurationsResponse.status === 204) {
            this.setState({ isAuthenticated: true, isError: false });
        }
        else if (configurationsResponse.status === 200) {
            const results = await configurationsResponse.data;
            this.setState({ TeamId: results.TeamLink, CardItems: JSON.parse(results.CardTemplate), isAuthenticated: true, isError: false });
        }
        else {
            this.setState({ isAuthenticated: true, isError: true });
        }
    }

    logout = () => {
        authContext.logOut();
    }

    onPublish = (result: boolean, msg: string) => {
        if (result) {
            this.setState({ Message: msg, MessageType: result });
        }
        else {
            this.setState({ Message: msg, MessageType: result });
        }
       
    }

    /** Render function. */
    render() {
        const renderDetails = () => {
                return (
                    <div className="container-div">

                        <Flex gap="gap.small" padding="padding.medium">
                            <Flex.Item size="size.half" >
                                <Text align="end" content={this.state.resourceStrings.common.teamLink + "* "} className="team-link" />
                            </Flex.Item>
                            <Flex.Item size="size.half" className="team-textbox-width" >
                                <Input fluid placeholder={this.state.resourceStrings.common.teamLink}
                                    value={this.state.TeamId}
                                    onChange={(e: any) => { this.setState({ TeamId: e.target.value }) }} />
                            </Flex.Item>
                        </Flex>
                        <Flex hAlign="center">
                            <BuildYourForm HomeState={this.state} onPublish={this.onPublish} resourceStrings={this.state.resourceStrings} />
                        </Flex>
                        {this.state.MessageType ? <Text className="medium-margin-top" align="center" success content={this.state.Message} />
                            : <Text className="medium-margin-top" align="center" error content={this.state.Message} />}
                    </div>
                );
        }

        const renderComponent = () => {
            if (this.state.resourceStringsLoaded) {
                return (
                    <Provider theme={themes.teams} >
                        <Segment content={
                            <Flex vAlign="center" gap="gap.small">
                                <Text size="large" content={this.state.resourceStrings.common.mainHeader} />
                                <Flex.Item push>
                                    <div>
                                        <Text content={authContext.getCachedUser().userName} className="small-margin-right" />
                                        <Button content={this.state.resourceStrings.common.btnLogout} onClick={this.logout} />
                                    </div>
                                </Flex.Item>
                            </Flex>
                        }
                            inverted />

                        {this.state.isAuthenticated === true && this.state.isError === false && <div>{renderDetails()}</div>}
                        {this.state.isAuthenticated === false && <Text content={this.state.resourceStrings.common.notAuthorized} align="center" error />}
                        {this.state.isError === true && <Text content={this.state.resourceStrings.common.genericError} align="center" error />}
                        {this.state.isAuthenticated === null && <Text content={this.state.resourceStrings.common.loading} align="center" />}
                    </Provider>
                );
            }
            else {
                return (
                    <Provider>
                        <Loader />
                    </Provider>
                );
            }
        }

        return (
            renderComponent()
        );
    };
}

export default Home;