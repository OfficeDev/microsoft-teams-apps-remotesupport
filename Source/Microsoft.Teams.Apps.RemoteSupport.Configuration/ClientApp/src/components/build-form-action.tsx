/*
    <copyright file="build-form-action.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import React from 'react';
import { Dialog, Button, Dropdown, Flex, Text, Divider, Input } from '@fluentui/react';
import { isNullorWhiteSpace, userControls } from '../constants';
import { saveConfigurationsAsync } from "../api/incident-api";
import InputTextForm from './input-text-form';
import ChoiceSetForm from './choiceset-form';
import DatePickerForm from './datepicker-form';
import RadioButtonForm from './radiobutton-form';
import CheckBoxForm from './checkbox-form';
import { InputTextPreview } from './input-text-preview';
import { CheckBoxPreview } from './checkbox-preview';
import { ChoiceSetPreview } from './choiceset-preview';
import { DatePickerPreview } from './datepicker-preview';
import { RadioButtonPreview } from './radiobutton-preview';
import { getToken } from '../adal-config';


interface ITeamState {
    SelectedItem: string,
    jsx: Array<JSX.Element>,
    json: Array<any>,
    Error: string,
    open: boolean | undefined,
}

interface ITeamProps {
    HomeState: {
        TeamId: string,
        CardItems: Array<any>
    },
    onPublish: (result: boolean, msg: string) => void,
    resourceStrings: any,
}


/** Component for displaying home page of incident report configuration application. */
class BuildYourForm extends React.Component<ITeamProps, ITeamState>
{
    state: ITeamState;
    bearer: string = "";
    requestType: Array<any> = [];
    inputItems: Array<any> = [];
    /**
     * Constructor to initialize component.
     * @param props Props of component.
     */
    constructor(props: ITeamProps) {
        super(props);
        this.state = {
            SelectedItem: '',
            jsx: [],
            json: [],
            open: undefined,
            Error: '',
        };

        this.bearer = getToken();
        this.requestType = [
            props.resourceStrings.common.normalSeverity,
            props.resourceStrings.common.urgentSeverity,
        ];

        this.inputItems = [
            {
                header: props.resourceStrings.buildForm.inputText,
                controlId: userControls.inputText
            },
            {
                header: props.resourceStrings.buildForm.dropDown,
                controlId: userControls.dropDown
            },
            {
                header: props.resourceStrings.buildForm.inputDate,
                controlId: userControls.inputDate
            },
            {
                header: props.resourceStrings.buildForm.radioButton,
                controlId: userControls.radioButton
            },
            {
                header: props.resourceStrings.buildForm.checkBox,
                controlId: userControls.checkBox
            },
        ]
    }

    onOpen = () => {
        let items = this.props.HomeState.CardItems;
        this.addComponentsToState(items);
    }
    OnChangeSelection = {
        onAdd: item => {
            this.setState({ SelectedItem: item.controlId });
            return `${item} has been selected.`
        },
        onRemove: item => {
            this.setState({ SelectedItem: '' });
            return `${item} has been removed.`
        },
    };

    OnDeleteComponent = (index: number) => {
        let jsonItems = this.state.json;
        jsonItems.splice(index, 1);

        this.addComponentsToState(jsonItems);

    }

    ValidateUrl = () => {
        let link = this.props.HomeState.TeamId;
        if (isNullorWhiteSpace(link)) {
            this.setState({ open: false });
            this.props.onPublish(false, this.props.resourceStrings.common.invalidTeamLink);
            return;
        }

        if (!link.includes("teams.microsoft.com/")) {
            this.setState({ open: false });
            this.props.onPublish(false, this.props.resourceStrings.common.invalidTeamLink);
            return;
        }

        this.setState({ open: true });
    }

    onPublish = (event: any) => {
        this.setState({ open: false });

        this.saveConfigurationsAsync(this.props.HomeState.TeamId, JSON.stringify(this.state.json))
            .then(response => {
                this.props.onPublish(response, (response) ? this.props.resourceStrings.common.successPublish : this.props.resourceStrings.common.genericError)
            });

    }

    onCancel = () => {
        this.setState({open : false});
    }

    OnAddComponent = (properties: any) : boolean => {
        let keyVal = this.state.json.length;
        if (keyVal > 3) {
            this.setState({ Error: this.props.resourceStrings.buildForm.maxComponents })
            return false;
        }

        if (isNullorWhiteSpace(properties.displayName)) {

            this.setState({ Error: this.props.resourceStrings.buildForm.notEmptyDisplayName })
            return false;
        }
        else if (properties.displayName.length > 50) {
            this.setState({ Error: this.props.resourceStrings.buildForm.maxLengthDisplayName })
            return false;
        }


        let items = this.state.json;
        let uniqueDisplayNameCheck = items.find(element => (element.id.substr(0, element.id.indexOf('_')).toUpperCase() === properties.displayName.toUpperCase()))

        if (uniqueDisplayNameCheck) {
            this.setState({ Error: this.props.resourceStrings.buildForm.uniqueDisplayName })
            return false;
        }
        // switch case push component based on type
        switch (properties.type) {
            case 'Input.Text':

                items.push({
                    "type": "Input.Text",
                    "placeholder": properties.placeholder,
                    "maxLength": 500,
                    "id": properties.displayName + "_InputText" + keyVal
                });

                break;
            case 'Input.ChoiceSet':
                if (properties.style === 'expanded') {
                    let choices = properties.options.map(x => ({ "title": x, "value": x }));
                    items.push({
                        "type": "Input.ChoiceSet",
                        "placeholder": properties.placeholder,
                        "choices": choices,
                        "style": properties.style,
                        "id": properties.displayName + "_ChoiceSet" + keyVal
                    });
                }
                else if (properties.isMultiSelect === true) {
                    let choices = properties.options.map(x => ({ "title": x, "value": x }));
                    items.push({
                        "type": "Input.ChoiceSet",
                        "placeholder": properties.placeholder,
                        "choices": choices,
                        "isMultiSelect": true,
                        "id": properties.displayName + "_ChoiceSet" + keyVal
                    });
                }
                else {
                    let choices = properties.options.map(x => ({ "title": x, "value": x }));
                    items.push({
                        "type": "Input.ChoiceSet",
                        "placeholder": properties.placeholder,
                        "choices": choices,
                        "id": properties.displayName + "_ChoiceSet" + keyVal
                    });
                }

                break;
            case 'Input.Date':
                items.push({
                    "type": "Input.Date",
                    "id": properties.displayName + "_Date" + keyVal
                });
                break;

        };

        this.addComponentsToState(items);
        return true;
    }

    addComponentsToState = (items: Array<any>) => {
        let components: Array<JSX.Element> = [];
        items.forEach((element, index) => {
            switch (element.type) {
                case 'Input.Text':
                    components.push(
                        <InputTextPreview key={index} keyVal={index} placeholder={element.placeholder} displayName={element.id.substr(0, element.id.indexOf('_'))} OnDeleteComponent={this.OnDeleteComponent} />
                    );
                    break;
                case 'Input.ChoiceSet':
                    if (element.style === 'expanded') {
                        components.push(
                            <RadioButtonPreview key={index} displayName={element.id.substr(0, element.id.indexOf('_'))} keyVal={index} OnDeleteComponent={this.OnDeleteComponent} options={element.choices.map(x => x.title)} />
                        );
                    }
                    else if (element.isMultiSelect === true) {
                        components.push(
                            <CheckBoxPreview key={index} displayName={element.id.substr(0, element.id.indexOf('_'))} keyVal={index} OnDeleteComponent={this.OnDeleteComponent} options={element.choices.map(x => x.title)} />
                        );
                    }
                    else {
                        components.push(
                            <ChoiceSetPreview key={index} displayName={element.id.substr(0, element.id.indexOf('_'))} keyVal={index} OnDeleteComponent={this.OnDeleteComponent} options={element.choices.map(x => x.title)} placeholder={element.placeholder} />
                        );
                    }
                    break;
                case 'Input.Date':
                    components.push(
                        <DatePickerPreview displayName={element.id.substr(0, element.id.indexOf('_'))} key={index} keyVal={index} OnDeleteComponent={this.OnDeleteComponent} />
                    );
                    break;
            }
        });
        this.setState({ json: items, jsx: components, Error: '' });
    }

    /** 
     *  Save configuration details.
     *  @param teamId.
     *  @cardJson json array string that represents each item of adaptive card.
     * */
    saveConfigurationsAsync = async (teamId: string, cardJson: string): Promise<boolean> => {
        let configurationDetail = {
            TeamLink: teamId,
            CardTemplate: cardJson
        };

        const saveConfigurationsResult = await saveConfigurationsAsync(configurationDetail, this.bearer);
        if (saveConfigurationsResult.status === 200) {
            return true;
        }
        else {
            return false;
        }
    }

    /** Render function. */
    render() {
        const renderComponent = () => {
            switch (Number(this.state.SelectedItem)) {
                case userControls.inputText:
                    return (<InputTextForm OnAddComponent={this.OnAddComponent} resourceStrings={this.props.resourceStrings.common} />);
                case userControls.dropDown:
                    return (<ChoiceSetForm OnAddComponent={this.OnAddComponent} resourceStrings={this.props.resourceStrings} />);
                case userControls.inputDate:
                    return (<DatePickerForm OnAddComponent={this.OnAddComponent} resourceStrings={this.props.resourceStrings.common} />);
                case userControls.radioButton:
                    return (<RadioButtonForm OnAddComponent={this.OnAddComponent} resourceStrings={this.props.resourceStrings} />);
                case userControls.checkBox:
                    return (<CheckBoxForm OnAddComponent={this.OnAddComponent} resourceStrings={this.props.resourceStrings} />);
            }
        }

        return (
            <Dialog
                open={this.state.open}
                styles={{ width: '75vw', height: '85vh', overflow: 'auto' }}
                confirmButton='Publish'
                onOpen={this.onOpen}
                onConfirm={this.onPublish}
                cancelButton="Cancel"
                onCancel={this.onCancel}
                content=
                {
                    <>
                        <Flex className='dialog-container-div' gap='gap.small' padding='padding.medium'>
                            <Flex.Item size='size.half'>
                            <div>
                                    <Text size="large" content={this.props.resourceStrings.buildForm.headerTitle} weight="semibold" />
                                <Divider key={1} size={1} />
                                {this.state.Error.length > 0 && <Text error content={this.state.Error} />}
                                <Flex column gap='gap.small' padding='padding.medium'>
                                    <Dropdown clearable items={this.inputItems}
                                            placeholder={this.props.resourceStrings.buildForm.componentDropdown}
                                        getA11ySelectionMessage={this.OnChangeSelection} />
                                    {renderComponent()}
                                </Flex>
                            </div>
                            </Flex.Item>
                           
                            <Divider vertical />
                            <Flex.Item size='size.half'>
                            <div id='components'>
                                    <Text content={this.props.resourceStrings.buildForm.previewTitle} size="large" weight="semibold" />
                                <Divider fitted key={1} size={1} styles={{ padding: "0.5rem" }} />
                                <Flex column gap='gap.small' >
                                        <Input fluid placeholder={this.props.resourceStrings.buildForm.staticInput1} styles={{ marginTop: '1rem' }} />
                                        <Text content={this.props.resourceStrings.buildForm.staticText2} />
                                        <Input fluid placeholder={this.props.resourceStrings.buildForm.staticInput2} />
                                        <Text content={this.props.resourceStrings.buildForm.staticDropdown} />
                                        <Dropdown items={this.requestType} placeholder={this.props.resourceStrings.buildForm.staticDropdownPlaceholder} />
                                    {this.state.jsx}
                                </Flex>
                            </div>
                            </Flex.Item>
                        </Flex>
                    </>
                }
                trigger={<Button content={this.props.resourceStrings.buildForm.btnBuildForm} onClick={this.ValidateUrl}/>}
            />
        );
    };
}
export default BuildYourForm;







