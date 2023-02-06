/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
        if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
            console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
        }
        // Assign event handlers and other initialization logic.
        document.getElementById("insert-paragraph").onclick = insertParagraph;

        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run;
        document.getElementById("run1").onclick = run;
        document.getElementById("apply-style").onclick = applyStyle;
    }
});

async function applyStyle() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to style text.
        const firstParagraph = context.document.body.paragraphs.getFirst();
        firstParagraph.styleBuiltIn = Word.Style.intenseReference;


        await context.sync();
    })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
}
async function insertParagraph() {
    await Word.run(async (context) => {
        const docBody = context.document.body;
        //文首添加
        docBody.insertParagraph(
            "Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
            "Start"
        );
        //文末添加
        docBody.insertParagraph("233333", "End");


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
