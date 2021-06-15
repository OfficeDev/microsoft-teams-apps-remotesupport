/*
    <copyright file="on-call-support-detail.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

export class OnCallSupportDetail {
    ModifiedByName: string | "" = "";
    ModifiedByObjectId?: string | null = null;
    ModifiedOn: Date | null = null;
    OnCallSMEs: string | "" = "";
}