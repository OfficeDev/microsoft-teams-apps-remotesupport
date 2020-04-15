/*
    <copyright file="checkbox-preview.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import React from "react";
import { Flex, Icon, Checkbox, Text } from '@fluentui/react';
import "../styles/theme.css";

interface IPreviewProps {
    keyVal: number,
    displayName: string,
    options: Array<string>,
    onDeleteComponent: (keyVal: number) => void
}

export const CheckBoxPreview: React.FunctionComponent<IPreviewProps> = (props) => {
    const onDeleteComponent = (keyVal: number) => {
        props.onDeleteComponent(keyVal);
    }

    return (
        <Flex key={props.keyVal} className="preview-item">
            <Flex.Item key={props.keyVal} grow>
                <div>
                    <Text content={props.displayName} /><br/>
                    <div id={props.keyVal.toString()} style={{ "marginTop": "0.5rem"}}>
                        {props.options.map(function(option, index){
                            return <><Checkbox label={option} key={index}/><br/></>
                        })}
                    </div>
                </div>
            </Flex.Item>
            <Icon className="common-padding" name='trash-can' onClick={() => onDeleteComponent(props.keyVal)} />
        </Flex>
    );
}