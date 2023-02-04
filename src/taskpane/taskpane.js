/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("insert-paragraph").onclick = insertParagraph;
    }
});

async function insertParagraph() {
    await Word.run(async (context) => {
        const docBody = context.document.body;
        docBody.insertParagraph(
            "Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
            "Start"
        );
        await context.sync();
    }).catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}

export async function run() {
    return Word.run(async (context) => {
        /**
         * Insert your Word code here
         */

        // insert a paragraph at the end of the document.
        const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

        // change the paragraph color to blue.
        paragraph.font.color = "blue";

        await context.sync();
    });
}
