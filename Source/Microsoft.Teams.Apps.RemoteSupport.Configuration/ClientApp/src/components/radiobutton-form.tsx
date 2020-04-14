/*
    <copyright file="radiobutton-form.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import React from "react";
import { Icon, Input, Text, Flex, Button } from '@fluentui/react';
import { isNullorWhiteSpace } from "../constants";
import "../styles/theme.css";

interface IPropertiesProps {
    onAddComponent: (properties: any) => boolean,
    resourceStrings: any,
}

interface IChoices {
    type: string,
    choicesJsx: Array<JSX.Element>,
    options: Array<string>,
    option: string,
    displayName: string,
    style: string,
    error: string
}

const RadioButtonForm: React.FunctionComponent<IPropertiesProps> = (props) => {
    const [properties, setProperties] = React.useState<IChoices>({ type: 'Input.ChoiceSet', choicesJsx: [], option: '', options: [], style: 'expanded', displayName:'', error: '' });

    const onAddComponent = (event: any) => {
        if (properties.options.length < 2) {
            setProperties({
                choicesJsx: properties.choicesJsx,
                type: properties.type,
                options: properties.options,
                option: properties.option,
                style: properties.style,
                displayName: properties.displayName,
                error: props.resourceStrings.common.minimumItems
            });
            return;
        }

        let result = props.onAddComponent(properties);
        if (result) {
            setProperties({ type: 'Input.ChoiceSet', choicesJsx: [], option: '', options: [], style: 'expanded', displayName: '', error: '' });
        }
    }

    const onDeleteOption = (key: number, option: string) => {
        let result = properties.choicesJsx.find(x => x.key === option + key.toString());
        properties.choicesJsx.splice(properties.choicesJsx.indexOf(result!), 1);
        properties.options.splice(key, 1); 
     
        setProperties({
            choicesJsx: properties.choicesJsx,
            type: properties.type,
            options: properties.options,
            option: '',
            style: properties.style,
            displayName: properties.displayName,
            error: ''
        });

    }

    const onAddOption = (event: any) => {
        let choicesJsx = properties.choicesJsx;
        let keyVal = choicesJsx.length;
        let options = properties.options;

        if (keyVal > 2) {
            setProperties({
                choicesJsx: choicesJsx,
                type: properties.type,
                options: options,
                option: '',
                style: properties.style,
                displayName: properties.displayName,
                error: props.resourceStrings.radiobutton.maxRadioChoices
            });
            return;
        }

        if (isNullorWhiteSpace(properties.option)) {
            setProperties({
                choicesJsx: choicesJsx,
                type: properties.type,
                options: options,
                option: '',
                style: properties.style,
                displayName: properties.displayName,
                error: props.resourceStrings.common.nonEmptyItem
            });
            return;
        }
        else {
            let uniqueItemsCheck = properties.options.find(item => item.toUpperCase() === properties.option.toUpperCase());
            if (uniqueItemsCheck) {
                setProperties({
                    choicesJsx: choicesJsx,
                    type: properties.type,
                    options: options,
                    option: properties.option,
                    style: properties.style,
                    displayName: properties.displayName,
                    error: props.resourceStrings.common.duplicateItem
                });
                return;
            }
        }

        choicesJsx.push(
            <Flex key={properties.option + keyVal} padding="padding.medium">
                <Flex.Item>
                    <div>
                        <Text content={properties.option} />
                        <Icon name="trash-can" className="common-padding" onClick={() => onDeleteOption(keyVal, properties.option)} />
                    </div>
                </Flex.Item>
            </Flex>
        );

        options.push(properties.option);
        setProperties({
            choicesJsx: choicesJsx,
            type: properties.type,
            options: options,
            option: '',
            style: properties.style,
            displayName: properties.displayName,
            error: ''
        });
    }

    return (
        <Flex column gap="gap.small" padding="padding.medium">
            <Flex.Item>
                <>
                    <Text content={props.resourceStrings.common.displayName} />
                    <Input fluid placeholder={props.resourceStrings.common.displayNamePlaceholder} value={properties.displayName} onChange={(e: any) => { setProperties({ ...properties, displayName: e.target.value }) }} />
                    <Text content={props.resourceStrings.radiobutton.radioOptions} />
                    <div>
                        <Input placeholder="Text" value={properties.option} onChange={(e: any) => { setProperties({ ...properties, option: e.target.value }) }} />
                        <Icon name="add" onClick={onAddOption} className = "add-icon" />
                    </div>
                </>
            </Flex.Item>
            <Flex.Item>
                <Flex column>
                    <Text content={properties.error} error />
                    {properties.choicesJsx}
                </Flex>
            </Flex.Item>
            <Button content={props.resourceStrings.common.btnAddComponent} onClick={onAddComponent} />
        </Flex>
    );
}

export default RadioButtonForm;