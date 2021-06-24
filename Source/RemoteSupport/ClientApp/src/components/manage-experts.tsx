/*
    <copyright file="manage-experts.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import * as React from "react";
import { ApplicationInsights, SeverityLevel } from "@microsoft/applicationinsights-web";
import { ReactPlugin, withAITracking } from "@microsoft/applicationinsights-react-js";
import * as microsoftTeams from "@microsoft/teams-js";
import { createBrowserHistory } from "history";
import { Dropdown, Button, Loader, Flex, Text, List, Provider, themes } from "@fluentui/react";
import { getResourceStrings, saveOnCallSupportDetails, getMembersInTeam, getOnCallExpertsInTeam, handleError } from "../api/remote-support-api";
import { OnCallSupportDetail } from "../models/on-call-support-detail"
import Constants from "../constants/constants";
import "../styles/site.css";

interface IState {
    loading: boolean,
    theme: string | null,
    themeStyle: any;
    teamMembers: any[];
    onCallExperts: any[];
    allMembers: any[];
    selectedMembers: any[];
    isSubmitExpertListLoading: boolean;
    errorMessage: string | null;
    resourceStrings: any;
    resourceStringsLoaded: boolean;
}

const browserHistory = createBrowserHistory({ basename: "" });
let reactPlugin = new ReactPlugin();

/** Component for displaying on call support team details. */
class ManageExperts extends React.Component<{}, IState>
{
    customAPIAuthenticationToken?: string | null = null;
    locale?: string | null;
    telemetry?: any = null;
    appInsights: ApplicationInsights;
    theme: string | null = null;
    userEmail?: any = null;
    userObjectId?: string | null = null;
    activityId: string | null = null;

    constructor(props: {}) {
        super(props);
        this.state = {
            loading: false,
            theme: null,
            themeStyle: themes.teams,
            teamMembers: [],
            onCallExperts: [],
            allMembers: [],
            selectedMembers: [],
            isSubmitExpertListLoading: false,
            errorMessage: "",
            resourceStrings: {},
            resourceStringsLoaded: false,
        };
        
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.telemetry = params.get("telemetry");
        this.customAPIAuthenticationToken = params.get("token");
        this.theme = params.get("theme");
        this.activityId = params.get("activityId");
        this.locale = params.get("locale");

        // Initialize application insights for logging events and errors.
        try {
            this.appInsights = new ApplicationInsights({
                config: {
                    instrumentationKey: this.telemetry,
                    extensions: [reactPlugin],
                    extensionConfig: {
                        [reactPlugin.identifier]: { history: browserHistory }
                    }
                }
            });
            this.appInsights.loadAppInsights();
        }
        catch (exception) {
            this.appInsights = new ApplicationInsights({
                config: {
                    instrumentationKey: undefined,
                    extensions: [reactPlugin],
                    extensionConfig: {
                        [reactPlugin.identifier]: { history: browserHistory }
                    }
                }
            });
            console.log(exception);
        }
    }

    /** Called once component is mounted. */
    async componentDidMount() {
        microsoftTeams.initialize();
        this.updateTheme(this.theme!);
        this.setState({
            theme: this.theme!
        });

        microsoftTeams.getContext((context) => {
            this.userObjectId = context.userObjectId;
            this.userEmail = context.upn;
            this.locale = context.locale;
        });

        microsoftTeams.registerOnThemeChangeHandler((theme) => {
            this.updateTheme(theme);
            this.setState({
                theme: theme,
            }, () => {
                this.forceUpdate();
            });
        });

        this.getResourceStrings();
        this.getMembersInTeam();
    }

    /**
	* Set current theme state received from teams context
	* @param  {String} theme Current theme name
	*/
    private updateTheme = (theme: string) => {
        if (theme === Constants.dark) {
            this.setState({
                themeStyle: themes.teamsDark
            });
        } else if (theme === Constants.contrast) {
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
        this.appInsights.trackTrace({ message: `'getResourceStrings' - Request initiated`, severityLevel: SeverityLevel.Information, properties: { User: this.userObjectId } });
        const resourceStringsResponse = await getResourceStrings(this.customAPIAuthenticationToken!, this.locale);
        if (resourceStringsResponse) {
            this.setState({ resourceStringsLoaded: true });

            if (resourceStringsResponse.status === 200) {
                this.setState({ resourceStrings: resourceStringsResponse.data });
            }
            else {
                handleError(resourceStringsResponse, this.customAPIAuthenticationToken);
            }
        }
    }

    /** 
    *  Get all team members.
    * */
    getMembersInTeam = async () => {
        this.appInsights.trackTrace({ message: `'getMembersInTeam' - Request initiated`, severityLevel: SeverityLevel.Information });
        this.setState({ loading: true });
        const teamMemberResponse = await getMembersInTeam(this.customAPIAuthenticationToken!);
        if (teamMemberResponse) {
            if (teamMemberResponse.status === 200) {
                this.setState({ allMembers: teamMemberResponse.data });
                this.getOnCallExpertsInTeam();
            }
            else {
                handleError(teamMemberResponse, this.customAPIAuthenticationToken);
            }
        }
        this.setState({ loading: false });
    }

    /** 
     *  Get on call experts configured in team.
     * */
    getOnCallExpertsInTeam = async () => {
        this.appInsights.trackTrace({ message: `'getOnCallExpertsInTeam' - Request initiated`, severityLevel: SeverityLevel.Information });
        this.setState({ loading: true });
        const onCallExpertsResponse = await getOnCallExpertsInTeam(this.customAPIAuthenticationToken!);
        if (onCallExpertsResponse) {
            if (onCallExpertsResponse.status === 200) {
                // Convert API response into list items.
                let onCallExperts: Array<any> = [];
                let teamMembers = this.state.allMembers;
                if (onCallExpertsResponse.data != null && onCallExpertsResponse.data[0] != null) {
                    let onCallExpertsDetails = JSON.parse(onCallExpertsResponse.data[0].OnCallSMEs);
                    onCallExpertsDetails.forEach((onCallExpertsDetail) => {

                        let member = teamMembers.find(element => element.aadobjectid == onCallExpertsDetail.objectid);
                        const actions = (
                            <Flex gap="gap.large" vAlign="center">
                                <Button icon="icon-close" text iconOnly title="Close" onClick={() => this.removeUserFromExpertList(onCallExpertsDetail.objectid)} />
                            </Flex>
                        );
                        if (member != undefined) {
                            onCallExperts.push({
                                key: onCallExpertsDetail.objectid,
                                header: member.header,
                                endMedia: actions,
                            });
                        }
                        teamMembers = teamMembers.filter(item => item !== member)
                    });
                }
                this.setState({ onCallExperts: onCallExperts });
                this.setState({ teamMembers: teamMembers });


            }
            else if (onCallExpertsResponse.status === 204) {
                this.setState({ teamMembers: this.state.allMembers });
            }
            else {
                handleError(onCallExpertsResponse, this.customAPIAuthenticationToken);
            }
        }
        this.setState({ loading: false });
    }

    getA11ySelectionMessage = {
        onAdd: item => {
            let selectedMembers = this.state.selectedMembers;
            selectedMembers.push(item);
            this.setState({ selectedMembers: selectedMembers });
            return "";

        },
        onRemove: item => {
            let selectedMembers = this.state.selectedMembers;
            selectedMembers.splice(selectedMembers.indexOf(item), 1);
            this.setState({ selectedMembers: selectedMembers });
            return "";
        }
    };

    removeUserFromExpertList(objectId) {
        let teamMembers = this.state.teamMembers;
        let member = this.state.allMembers.find(element => element.aadobjectid == objectId);
        teamMembers.push(member);
        this.setState({ teamMembers: teamMembers });

        let onCallExperts = this.state.onCallExperts;
        let onCallExpert = onCallExperts.find(element => element.key == objectId);
        onCallExperts = onCallExperts.filter(item => item !== onCallExpert)
        this.setState({ onCallExperts: onCallExperts });
    }

    submitExpertList = async () => {
        this.setState({ isSubmitExpertListLoading: true });
        let onCallExperts = this.state.onCallExperts;
        let onCallSupportExpertDetails: Array<any> = [];

        onCallExperts.forEach((onCallExpert) => {
            onCallSupportExpertDetails.push({
                name: onCallExpert.header.split(" ")[0],
                objectid: onCallExpert.key,
            })
        });

        if (onCallSupportExpertDetails.length > 15) {
            this.setState({ isSubmitExpertListLoading: false, errorMessage: this.state.resourceStrings.maxOnCallExpertsAllowedText });
            return;
        }
        else {
            this.setState({ errorMessage: "" });
        }
        let member = this.state.allMembers.find(element => element.aadobjectid == this.userObjectId);

        let onCallSupportDetail: OnCallSupportDetail = {
            ModifiedByName: member.header,
            ModifiedByObjectId: this.userObjectId != null ? this.userObjectId.toString() : null,
            ModifiedOn: new Date(),
            OnCallSMEs: JSON.stringify(onCallSupportExpertDetails),
        };

        this.appInsights.trackTrace({ message: `'submitExpertList' - Request initiated`, severityLevel: SeverityLevel.Information, properties: { UserEmail: this.userEmail } });
        const expertListResponse = await saveOnCallSupportDetails(onCallSupportDetail, this.customAPIAuthenticationToken!);
        if (expertListResponse.status === 200) {

            let allMembers = this.state.allMembers;
            let onCallExpertsList: Array<string> = [];
            this.state.onCallExperts.forEach((onCallExpertsDetail) => {
                let member = allMembers.find(element => element.aadobjectid == onCallExpertsDetail.key);
                onCallExpertsList.push(member.aadobjectid);
            });

            let toBot = { Command: Constants.updateExpertListCommand, OnCallExpertsList: onCallExpertsList, OnCallSupportCardActivityId: this.activityId, OnCallSupportId: expertListResponse.data };

            microsoftTeams.getContext((context) => {
                microsoftTeams.tasks.submitTask(toBot);
            });
        }
        else {
            this.setState({ isSubmitExpertListLoading: false, errorMessage: this.state.resourceStrings.errorMessage });
            handleError(expertListResponse, this.customAPIAuthenticationToken);
        }
    }

    addExpertInList = () => {
        let onCallExperts = this.state.onCallExperts;
        let teamMembers = this.state.teamMembers;
        let selectedMembers = this.state.selectedMembers;

        selectedMembers.forEach((selectedMember) => {
            let member = teamMembers.find(element => element.aadobjectid == selectedMember.aadobjectid);
            const actions = (
                <Flex gap="gap.large" vAlign="center">
                    <Button icon="icon-close" text iconOnly title="Close" onClick={() => this.removeUserFromExpertList(selectedMember.aadobjectid)} />
                </Flex>
            );
            onCallExperts.push({
                key: selectedMember.aadobjectid,
                header: selectedMember.header,
                endMedia: actions,
            });
            teamMembers = teamMembers.filter(item => item !== member);
        });

        onCallExperts.forEach((onCallExpert) => {
            let member = teamMembers.find(element => element.aadobjectid == onCallExpert.objectid);
            teamMembers = teamMembers.filter(item => item !== member)
        });

        this.setState({ onCallExperts: onCallExperts });
        this.setState({ teamMembers: teamMembers });
        this.setState({ selectedMembers: [] });
    };

    renderManageExperts() {
        return (
            <div className="container-subdiv-main">
                <Flex gap="gap.large" vAlign="center" className="title">
                    <Text content={this.state.resourceStrings.expertNameTitle} />
                </Flex>
                <Flex gap="gap.large" vAlign="center">
                    <Flex.Item align="start" size="size.small" grow>
                        <Dropdown
                            multiple
                            search
                            items={this.state.teamMembers}
                            placeholder={this.state.resourceStrings.expertListPlaceHolderText}
                            getA11ySelectionMessage={this.getA11ySelectionMessage}
                            noResultsMessage={this.state.resourceStrings.noMatchesFoundText}
                            value={this.state.selectedMembers}
                        />
                    </Flex.Item>
                    <Flex.Item align="start" className="margin-button" size="size.small" grow>
                        <Provider
                            theme={{
                                componentVariables: {
                                    Button: siteVars => ({
                                        color: siteVars.colorScheme.brand.foreground,
                                        colorHover: siteVars.colorScheme.brand.foreground,
                                        colorFocus: siteVars.colorScheme.default.foreground,
                                        colorDisabled: siteVars.colorScheme.brandForegroundDisabled,
                                        backgroundColor: siteVars.colorScheme.default.background,
                                        backgroundColorActive: siteVars.colorScheme.brandBorderPressed,
                                        backgroundColorHover: siteVars.colorScheme.brand.backgroundHover1,
                                        backgroundColorFocus: siteVars.colorScheme.default.background,
                                        backgroundColorDisabled: siteVars.colorScheme.brand.backgroundDisabled,
                                        borderColor: siteVars.colorScheme.brandBorder2,
                                        borderColorHover: siteVars.colorScheme.brandBorderHover
                                    })
                                }
                            }}
                        >
                            <Button content={this.state.resourceStrings.addButtonText} onClick={this.addExpertInList} className="small-margin-left" />
                        </Provider>
                    </Flex.Item>
                </Flex>
                <div className="container-subdiv">
                    <Flex gap="gap.large" vAlign="center" className="list-title" >
                        {this.state.onCallExperts !== null && this.state.onCallExperts.length > 0 && <Text content={this.state.resourceStrings.expertListTitle} />}
                    </Flex>
                    <Flex >
                        {this.state.onCallExperts !== null && this.state.onCallExperts.length > 0 &&
                            <List className="list-width"
                                items={this.state.onCallExperts}
                            />
                        }
                    </Flex>
                </div>
                <div className="error">
                    <Flex gap="gap.small">
                        {this.state.errorMessage !== null && <Text className="small-margin-left" content={this.state.errorMessage} error />}
                    </Flex>
                </div>
                <div className="footer">
                    <Flex gap="gap.small">
                        <Button primary content={this.state.resourceStrings.saveButtonText} loading={this.state.isSubmitExpertListLoading} disabled={this.state.isSubmitExpertListLoading} onClick={this.submitExpertList} />
                    </Flex>
                </div>
            </div>
        );
    }

    render() {
        let contents = this.state.loading
            ? <p><em><Loader /></em></p>
            : this.renderManageExperts();
        if (this.state.resourceStringsLoaded) {
            return (
                <Provider theme={this.state.themeStyle}>
                    <div className="container-div">
                        {contents}
                    </div>
                </Provider>
            );
        }
        else {
            return (
                <Provider theme={this.state.themeStyle}>
                    <Loader />
                </Provider>
            );
        }
    }
}

export default withAITracking(reactPlugin, ManageExperts);
