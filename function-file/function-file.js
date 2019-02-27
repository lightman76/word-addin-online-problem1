/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

console.log("Loading function file")
// The initialize function must be run each time a new page is loaded
Office.initialize = reason => {
  console.log("Function file initialized!")

};

// Add any ui-less function here
