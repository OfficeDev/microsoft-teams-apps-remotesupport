/*
    <copyright file="input-text-preview.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import React from 'react';
import { Input, Flex, Icon, Text } from '@fluentui/react';
import "../styles/theme.css";

interface IPreviewProps {
    keyVal: number,
    placeholder: string,
    displayName: string,
    onDeleteComponent: (keyVal: number) => void
}

export const InputTextPreview: React.FunctionComponent<IPreviewProps> = (props) => {
    const onDeleteComponent = (keyVal: number) => {
        props.onDeleteComponent(keyVal);
    }

    return (
        <Flex key={props.keyVal} gap='gap.medium' vAlign="center" className="preview-item">
            <Flex.Item grow>
                <div>
                    <Text content={props.displayName} /><br />
                    <Input fluid key={'i' + props.keyVal} placeholder={props.placeholder} className="medium-margin-top" />
                </div>
            </Flex.Item>
            <Icon className="common-padding" name='trash-can' onClick={() => onDeleteComponent(props.keyVal)} />
        </Flex>
    );
}