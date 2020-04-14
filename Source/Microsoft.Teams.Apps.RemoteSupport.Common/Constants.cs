// <copyright file="Constants.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common
{
    /// <summary>
    /// Constant values that are used in multiple files.
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// New request action.
        /// </summary>
        public const string NewRequestAction = "NEW REQUEST";

        /// <summary>
        /// Send request action.
        /// </summary>
        public const string SendRequestAction = "SEND REQUEST";

        /// <summary>
        /// Edit request action.
        /// </summary>
        public const string EditRequestAction = "EDIT REQUEST";

        /// <summary>
        /// Withdraw request action.
        /// </summary>
        public const string WithdrawRequestAction = "WITHDRAW REQUEST";

        /// <summary>
        /// Manage experts action.
        /// </summary>
        public const string ManageExpertsAction = "EXPERT LIST";

        /// <summary>
        /// Update experts list action.
        /// </summary>
        public const string UpdateExpertListAction = "UPDATE EXPERT LIST";

        /// <summary>
        /// Message back card action.
        /// </summary>
        public const string MessageBackActionType = "messageBack";

        /// <summary>
        /// Task fetch action Type.
        /// </summary>
        public const string FetchActionType = "task/fetch";

        /// <summary>
        /// submit action Type.
        /// </summary>
        public const string SubmitActionType = "task/submit";

        /// <summary>
        /// Described adaptive card version to be used. Version can be upgraded or changed using this value.
        /// </summary>
        public const string AdaptiveCardVersion = "1.2";

        /// <summary>
        /// Update request action.
        /// </summary>
        public const string UpdateRequestAction = "UPDATE REQUEST";

        /// <summary>
        /// Ticket detail table name.
        /// </summary>
        public const string TicketDetailTable = "TicketDetail";

        /// <summary>
        /// Card configuration table name.
        /// </summary>
        public const string CardConfigurationTable = "CardConfiguration";

        /// <summary>
        ///  OnCallSupportDetail.
        /// </summary>
        public const string OnCallSupportDetailTable = "OnCallSupportDetail";

        /// <summary>
        /// No command.
        /// </summary>
        public const string NoCommand = "NO";

        /// <summary>
        /// Partition key for "OnCallSupportDetail" table.
        /// </summary>
        public const string OnCallSupportDetailPartitionKey = "OnCallSupport";

        /// <summary>
        /// Partition key for "TicketDetail" table.
        /// </summary>
        public const string TicketDetailPartitionKey = "Ticket";

        /// <summary>
        /// Partition key for "CardConfiguration" table.
        /// </summary>
        public const string CardConfigurationPartitionKey = "Card";

        /// <summary>
        /// Urgent request type text.
        /// </summary>
        public const string UrgentString = "Urgent";

        /// <summary>
        /// Date time format to support adaptive card text feature.
        /// </summary>
        /// <remarks>
        /// refer adaptive card text feature https://docs.microsoft.com/en-us/adaptive-cards/authoring-cards/text-features#datetime-formatting-and-localization.
        /// </remarks>
        public const string Rfc3339DateTimeFormat = "yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'";

        /// <summary>
        /// Cancel command.
        /// </summary>
        public const string CancelCommand = "CANCEL";

        /// <summary>
        /// Activity table store partition key name.
        /// </summary>
        public const string TicketIdGeneratorPartitionKey = "TicketId";

        /// <summary>
        ///  Urgent command id in the manifest file.
        /// </summary>
        public const string UrgentCommandId = "urgentrequests";

        /// <summary>
        /// Assigned requests command id in the manifest file.
        /// </summary>
        public const string AssignedCommandId = "assignedrequests";

        /// <summary>
        /// Unassigned requests command id in the manifest file.
        /// </summary>
        public const string UnassignedCommandId = "unassignedrequests";

        /// <summary>
        /// Active requests command id in the manifest file.
        /// </summary>
        public const string ActiveCommandId = "activerequests";

        /// <summary>
        /// Closed requests command id in the manifest file.
        /// </summary>
        public const string ClosedCommandId = "closedrequests";
    }
}