/*
    <copyright file="choiceset-preview.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import React from "react";
import { Flex, Icon, Dropdown, Text } from '@fluentui/react';
import "../styles/theme.css";

interface IPreviewProps {
    keyVal: number,
    placeholder: string,
    displayName : string,
    options: Array<string>,
    onDeleteComponent: (keyVal: number) => void
}

export const ChoiceSetPreview: React.FunctionComponent<IPreviewProps> = (props) => {
    const onDeleteComponent = (keyVal: number) => {
        props.onDeleteComponent(keyVal);
    }

    return (
        <Flex key={props.keyVal} className="preview-item">
            <Flex.Item grow>
                <div>
                    <Text content={props.displayName} /><br />
                    <Dropdown items={props.options} placeholder={props.placeholder} className="medium-margin-top" />
                </div>
            </Flex.Item>
            <Icon className="common-padding" name='trash-can' onClick={() => onDeleteComponent(props.keyVal)} />
        </Flex>
    );
}