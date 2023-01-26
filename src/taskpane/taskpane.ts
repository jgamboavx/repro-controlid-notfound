/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;

  // Uncomment below to see the error when trying to hide the CustomTab

  // Office.ribbon
  //   .requestUpdate({
  //     tabs: [
  //       {
  //         id: "ConstosoCustomTab",
  //         visible: false,
  //       },
  //     ],
  //   })
  //   .then(() => console.log("Done"))
  //   .catch((err) => console.log(err));

  // Uncomment below to see Office.js finding the same Tab it fails to above, and sucessfully disabling a control within it

  // Office.ribbon
  //   .requestUpdate({
  //     tabs: [
  //       {
  //         id: "ConstosoCustomTab",
  //         groups: [
  //           {
  //             id: "ConstosoCustomTabGroup",
  //             controls: [{ id: "ConstosoCustomTabButton", enabled: false }],
  //           },
  //         ],
  //       },
  //     ],
  //   })
  //   .then(() => console.log("Done"))
  //   .catch((err) => console.log(err));
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
