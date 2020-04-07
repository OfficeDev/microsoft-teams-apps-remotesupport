/*
    <copyright file="datepicker-form.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import React, { useState } from "react";
import { Button, Icon, Text, Flex, Input } from '@fluentui/react';
import "../styles/theme.css";

interface IPropertiesProps {
    OnAddComponent: (properties: any) => boolean,
    resourceStrings: any,
}


const DatePickerForm: React.FunctionComponent<IPropertiesProps> = (props) => {

    const [properties, setProperties] = useState({ type: 'Input.Date', displayName: '' });
    const OnAddComponent = (event: any) => {
        let result = props.OnAddComponent(properties);
        if (result) {
            setProperties({ type: 'Input.Date', displayName: '' });
        }
    }

    return (
        <>
            <Flex.Item>
                <>
                    <Text content={props.resourceStrings.displayName} />
                    <Input fluid placeholder={props.resourceStrings.displayNamePlaceholder} value={properties.displayName} onChange={(e: any) => { setProperties({ ...properties, displayName: e.target.value }) }} />
                    <div>
                        <Text content="dd-mm-yyyy" />
                        <Icon name="calendar" className="common-padding" color="blue" />
                    </div>
                </>
            </Flex.Item>
            <Button content={props.resourceStrings.btnAddComponent} onClick={OnAddComponent} />
        </>
    );
}

export default DatePickerForm;