/*
    <copyright file="input-text-form.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import React, { useState } from "react";
import { Button, Input, Text, Flex } from '@fluentui/react';

interface IPropertiesProps {
    onAddComponent: (properties: any) => boolean,
    resourceStrings: any,
}

const InputTextForm: React.FunctionComponent<IPropertiesProps> = (props) => {
    const [properties, setProperties] = useState({ type: 'Input.Text', placeholder: '', displayName: '' });
    const onAddComponent = (event: any) => {
        let result = props.onAddComponent(properties);
        if (result) {
            setProperties({ type: 'Input.Text', placeholder: '', displayName: '' });
        }
    }

    return (
        <>
            <Flex.Item grow>
                <>
                    <Text content={props.resourceStrings.displayName} />
                    <Input fluid placeholder={props.resourceStrings.displayNamePlaceholder} value={properties.displayName} onChange={(e: any) => { setProperties({ ...properties, displayName: e.target.value }) }} />
                    <Text content={props.resourceStrings.placeholder} />
                    <Input fluid placeholder={props.resourceStrings.placeholderText} value={properties.placeholder} onChange={(e: any) => { setProperties({ ...properties, placeholder: e.target.value }) }} />
                </>
            </Flex.Item>
            <Button content={props.resourceStrings.btnAddComponent} onClick={onAddComponent} />
        </>
    );
}

export default InputTextForm;