// <copyright file="constants.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

export const DarkTheme: string = "dark";
export const ContrastTheme: string = "contrast";

export const isNullorWhiteSpace = (input: string): boolean => {
    return !input || !input.trim();
}

export enum userControls {
    inputText = 1,
    dropDown = 2,
    inputDate = 3,
    radioButton = 4,
    checkBox = 5
}
