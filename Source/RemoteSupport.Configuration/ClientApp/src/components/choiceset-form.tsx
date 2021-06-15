/*
    <copyright file="choiceset-form.tsx" company="Microsoft">
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
    placeholder: string,
    displayName: string,
    choicesJsx: Array<JSX.Element>,
    options: Array<string>,
    option: string,
    error: string
}

const ChoiceSetForm: React.FunctionComponent<IPropertiesProps> = (props) => {
    const [properties, setProperties] = React.useState<IChoices>({ type: 'Input.ChoiceSet', choicesJsx: [], option: '', placeholder: '', options: [], displayName: '', error : '' });

    const onAddComponent = (event: any) => {
        if (properties.options.length < 2) {
            setProperties({
                placeholder: properties.placeholder,
                choicesJsx: properties.choicesJsx,
                type: properties.type,
                options: properties.options,
                option: properties.option,
                displayName: properties.displayName,
                error: props.resourceStrings.common.minimumItems
            });
            return;
        }

        let result = props.onAddComponent(properties);
        if (result) {
            setProperties({ type: 'Input.ChoiceSet', choicesJsx: [], option: '', placeholder: '', options: [], displayName: '', error: '' });
        }
    }

    const onDeleteOption = (key: string, option: string) => {
        let result = properties.choicesJsx.find(x => x.key === option + key);
        properties.choicesJsx.splice(properties.choicesJsx.indexOf(result!), 1);

        setProperties({
            placeholder: properties.placeholder,
            choicesJsx: properties.choicesJsx,
            type: properties.type,
            options: properties.options,
            option: '',
            displayName: properties.displayName,
            error: ''
        });
    }

    const onAddOption = (event: any) => {
        let choicesJsx = properties.choicesJsx;
        let keyVal = choicesJsx.length;
        let options = properties.options;

        if (keyVal > 9) {
            setProperties({
                placeholder: properties.placeholder,
                choicesJsx: properties.choicesJsx,
                type: properties.type,
                options: properties.options,
                option: '',
                displayName: properties.displayName,
                error: props.resourceStrings.dropdown.maxDropdownChoices
            });
            return;
        }

        if (isNullorWhiteSpace(properties.option)) {
            setProperties({
                placeholder: properties.placeholder,
                choicesJsx: properties.choicesJsx,
                type: properties.type,
                options: properties.options,
                option: '',
                displayName: properties.displayName,
                error: props.resourceStrings.common.nonEmptyItem
            });
            return;
        }
        else {
            let uniqueItemsCheck = properties.options.find(item => item.toUpperCase() === properties.option.toUpperCase());
            if (uniqueItemsCheck) {
                setProperties({
                    placeholder: properties.placeholder,
                    choicesJsx: properties.choicesJsx,
                    type: properties.type,
                    options: properties.options,
                    option: '',
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
                        <Icon name="trash-can" className="common-padding" onClick={() => onDeleteOption(keyVal.toString(), properties.option)} />
                    </div>
                </Flex.Item>
            </Flex>
        );

        options.push(properties.option);
        setProperties({
            placeholder: properties.placeholder,
            choicesJsx: choicesJsx,
            type: properties.type,
            options: options,
            option: '',
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

                    <Text content={props.resourceStrings.common.placeholder} />
                    <Input fluid placeholder={props.resourceStrings.common.placeholderText} value={properties.placeholder} onChange={(e: any) => { setProperties({ ...properties, placeholder: e.target.value }) }} />

                    <Text content={props.resourceStrings.dropdown.dropdownOptions} />
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

export default ChoiceSetForm;