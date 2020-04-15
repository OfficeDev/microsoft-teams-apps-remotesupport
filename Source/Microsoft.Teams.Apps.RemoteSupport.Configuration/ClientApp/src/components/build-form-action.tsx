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
    selectedItem: string,
    jsx: Array<JSX.Element>,
    json: Array<any>,
    error: string,
    open: boolean | undefined,
}

interface ITeamProps {
    homeState: {
        teamId: string,
        cardItems: Array<any>
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
            selectedItem: '',
            jsx: [],
            json: [],
            open: undefined,
            error: '',
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
        let items = this.props.homeState.cardItems;
        items = items.filter(agg => agg.type !== 'TextBlock');
        this.addComponentsToState(items);
    }

    onChangeSelection = {
        onAdd: item => {
            this.setState({ selectedItem: item.controlId });
            return `${item} has been selected.`
        },
        onRemove: item => {
            this.setState({ selectedItem: '' });
            return `${item} has been removed.`
        },
    };

    onDeleteComponent = (index: number) => {
        let jsonItems = this.state.json;
        jsonItems.splice(index, 1);

        this.addComponentsToState(jsonItems);

    }

    validateUrl = () => {
        let link = this.props.homeState.teamId;
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

        this.saveConfigurationsAsync(this.props.homeState.teamId, JSON.stringify(this.state.json))
            .then(response => {
                this.props.onPublish(response, (response) ? this.props.resourceStrings.common.successPublish : this.props.resourceStrings.common.genericError)
            });

    }

    onCancel = () => {
        this.setState({open : false});
    }

    onAddComponent = (properties: any) : boolean => {
        let keyVal = this.state.json.length;
        if (keyVal > 3) {
            this.setState({ error: this.props.resourceStrings.buildForm.maxComponents })
            return false;
        }

        if (isNullorWhiteSpace(properties.displayName)) {

            this.setState({ error: this.props.resourceStrings.buildForm.notEmptyDisplayName })
            return false;
        }
        else if (properties.displayName.length > 50) {
            this.setState({ error: this.props.resourceStrings.buildForm.maxLengthDisplayName })
            return false;
        }

        let items = this.state.json;
        let uniqueDisplayNameCheck = (properties.displayName.toUpperCase() === this.props.resourceStrings.buildForm.titleText.toUpperCase()
            || properties.displayName.toUpperCase() === this.props.resourceStrings.buildForm.descriptionText.toUpperCase()
            || properties.displayName.toUpperCase() === this.props.resourceStrings.buildForm.staticDropdown.toUpperCase()
            || items.find(element => (element.displayName.toUpperCase() === properties.displayName.toUpperCase())));

        if (uniqueDisplayNameCheck) {
            this.setState({ error: this.props.resourceStrings.buildForm.uniqueDisplayName })
            return false;
        }
        // switch case push component based on type
        switch (properties.type) {
            case 'Input.Text':
                items.push({
                    "type": "Input.Text",
                    "placeholder": properties.placeholder,
                    "maxLength": 500,
                    "id": properties.displayName + keyVal,
                    "displayName": properties.displayName
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
                        "id": properties.displayName + keyVal,
                        "displayName": properties.displayName
                    });
                }
                else if (properties.isMultiSelect === true) {
                    let choices = properties.options.map(x => ({ "title": x, "value": x }));
                    items.push({
                        "type": "Input.ChoiceSet",
                        "placeholder": properties.placeholder,
                        "choices": choices,
                        "isMultiSelect": true,
                        "id": properties.displayName + keyVal,
                        "displayName": properties.displayName
                    });
                }
                else {
                    let choices = properties.options.map(x => ({ "title": x, "value": x }));
                    items.push({
                        "type": "Input.ChoiceSet",
                        "placeholder": properties.placeholder,
                        "choices": choices,
                        "id": properties.displayName + keyVal,
                        "displayName": properties.displayName
                    });
                }
                break;

            case 'Input.Date':
                items.push({
                    "type": "Input.Date",
                    "id": properties.displayName + keyVal,
                    "displayName": properties.displayName
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
                        <InputTextPreview key={index} keyVal={index} placeholder={element.placeholder} displayName={element.displayName} onDeleteComponent={this.onDeleteComponent} />
                    );
                    break;

                case 'Input.ChoiceSet':
                    if (element.style === 'expanded') {
                        components.push(
                            <RadioButtonPreview key={index} displayName={element.displayName} keyVal={index} onDeleteComponent={this.onDeleteComponent} options={element.choices.map(x => x.title)} />
                        );
                    }
                    else if (element.isMultiSelect === true) {
                        components.push(
                            <CheckBoxPreview key={index} displayName={element.displayName} keyVal={index} onDeleteComponent={this.onDeleteComponent} options={element.choices.map(x => x.title)} />
                        );
                    }
                    else {
                        components.push(
                            <ChoiceSetPreview key={index} displayName={element.displayName} keyVal={index} onDeleteComponent={this.onDeleteComponent} options={element.choices.map(x => x.title)} placeholder={element.placeholder} />
                        );
                    }
                    break;

                case 'Input.Date':
                    components.push(
                        <DatePickerPreview displayName={element.displayName} key={index} keyVal={index} onDeleteComponent={this.onDeleteComponent} />
                    );
                    break;

            }
        });
        this.setState({ json: items, jsx: components, error: '' });
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
        return saveConfigurationsResult.status === 200;
    }

    /** Render function. */
    render() {
        const renderComponent = () => {
            switch (Number(this.state.selectedItem)) {
                case userControls.inputText:
                    return (<InputTextForm onAddComponent={this.onAddComponent} resourceStrings={this.props.resourceStrings.common} />);

                case userControls.dropDown:
                    return (<ChoiceSetForm onAddComponent={this.onAddComponent} resourceStrings={this.props.resourceStrings} />);

                case userControls.inputDate:
                    return (<DatePickerForm onAddComponent={this.onAddComponent} resourceStrings={this.props.resourceStrings.common} />);

                case userControls.radioButton:
                    return (<RadioButtonForm onAddComponent={this.onAddComponent} resourceStrings={this.props.resourceStrings} />);

                case userControls.checkBox:
                    return (<CheckBoxForm onAddComponent={this.onAddComponent} resourceStrings={this.props.resourceStrings} />);
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
                                {this.state.error.length > 0 && <Text error content={this.state.error} />}
                                <Flex column gap='gap.small' padding='padding.medium'>
                                    <Dropdown clearable items={this.inputItems}
                                            placeholder={this.props.resourceStrings.buildForm.componentDropdown}
                                        getA11ySelectionMessage={this.onChangeSelection} />
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
                                        <Input fluid placeholder={this.props.resourceStrings.buildForm.titleText} styles={{ marginTop: '1rem' }} />
                                        <Text content={this.props.resourceStrings.buildForm.descriptionText} />
                                        <Input fluid placeholder={this.props.resourceStrings.buildForm.descriptionPlaceholderText} />
                                        <Text content={this.props.resourceStrings.buildForm.staticDropdown} />
                                        <Dropdown items={this.requestType} placeholder={this.props.resourceStrings.buildForm.staticDropdownPlaceholder} />
                                    {this.state.jsx}
                                </Flex>
                            </div>
                            </Flex.Item>
                        </Flex>
                    </>
                }
                trigger={<Button content={this.props.resourceStrings.buildForm.btnBuildForm} onClick={this.validateUrl}/>}
            />
        );
    };
}

export default BuildYourForm;
