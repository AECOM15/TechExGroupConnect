// <copyright file="ColumnData.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

namespace Microsoft.Teams.Apps.DIConnect.Models.CardSetting
{
    using System;
    using Newtonsoft.Json;

    /// <summary>
    /// Adaptive column data class.
    /// </summary>
    public class ColumnData
    {
        /// <summary>
        /// Gets or sets image url value.
        /// </summary>
        [JsonProperty("imageUrl")]
        public Uri ImageUrl { get; set; }

        /// <summary>
        /// Gets or sets prompt question text value.
        /// </summary>
        [JsonProperty("promptQuestion")]
        public string PromptQuestion { get; set; }
    }
}