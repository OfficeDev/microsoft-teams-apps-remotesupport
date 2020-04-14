/*
    <copyright file="radiobutton-preview.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import React from "react";
import { Flex, Icon, RadioGroup, Text } from '@fluentui/react';
import "../styles/theme.css";

interface IPreviewProps {
    keyVal: number,
    displayName: string,
    options: Array<string>,
    onDeleteComponent: (keyVal: number) => void
}

export const RadioButtonPreview: React.FunctionComponent<IPreviewProps> = (props) => {
    const onDeleteComponent = (keyVal: number) => {
        props.onDeleteComponent(keyVal);
    }

    return (
        <Flex key={props.keyVal} className="preview-item">
            <Flex.Item grow>
                <div>
                    <Text content={props.displayName} />
                    <RadioGroup items={props.options} vertical className="medium-margin-top" />
                </div>
            </Flex.Item>
            <Icon className="common-padding" name='trash-can' onClick={() => onDeleteComponent(props.keyVal)} />
        </Flex>
    );
}