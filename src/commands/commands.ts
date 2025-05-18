/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { generateIndicators } from "../indicators";
import { getAllPageConfig } from "../page-config";

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
async function RefeshIndicators(event: Office.AddinCommands.Event) {

  console.log(generateIndicators(await getAllPageConfig()))

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("RefeshIndicators", RefeshIndicators);
