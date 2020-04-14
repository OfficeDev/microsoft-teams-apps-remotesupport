/*
    <copyright file="datepicker-preview.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import React from "react";
import { Icon, Text, Flex } from '@fluentui/react';
import "../styles/theme.css";

interface IPreviewProps {
    keyVal: number,
    displayName : string,
    OnDeleteComponent: (keyVal: number) => void
}

export const DatePickerPreview: React.FunctionComponent<IPreviewProps> = (props) => {
    const OnDeleteComponent = (keyVal: number) => {
        props.OnDeleteComponent(keyVal);
    }

    return (
        <Flex key={props.keyVal} gap='gap.medium' vAlign="center" className="preview-item">
            <Flex.Item grow>
                <div>
                    <Text content={props.displayName} /><br />
                    <Text content="dd-mm-yyyy" className="medium-margin-top" />
                    <Icon name="calendar" className="common-padding" color="blue" />
                </div>
            </Flex.Item>
            <Icon className="common-padding" name='trash-can' onClick={() => OnDeleteComponent(props.keyVal)} />
        </Flex>
    );
}
