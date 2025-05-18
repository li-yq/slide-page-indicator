/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { generateIndicators, PartPageIndicators } from "../indicators";
import { getAllPageConfig } from "../page-config";

/* global Office */

Office.onReady(() => {
    // If needed, Office.js is ready to be called.
    OfficeExtension.config.extendedErrorLogging = true;


});


function getSelectedSlideIndex() {
    return new OfficeExtension.Promise<number>(function(resolve, reject) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function(asyncResult) {
            try {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    reject(console.error(asyncResult.error.message));
                } else {
                    resolve((asyncResult.value as any).slides[0].index);
                }
            }
            catch (error) {
                reject(console.log(error));
            }
        });
    });
}


/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
async function RefeshIndicators(event: Office.AddinCommands.Event) {

    const indicator_data = (generateIndicators(await getAllPageConfig()))
    // for (let index = 0; index < indicator_data.length; index++) {
    
    // get current slide index
    // let index = 3
    
    let index =await getSelectedSlideIndex()-1
        await UpdatePageIndicator(index, indicator_data[index]);
    // }

    // Be sure to indicate when the add-in command function is complete.
    event.completed();
}

// Register the function with Office.
Office.actions.associate("RefeshIndicators", RefeshIndicators);


async function UpdatePageIndicator(index, indicators: PartPageIndicators) {
    try {
        await PowerPoint.run(async (context) => {

            console.log("on page " + index)
            const shapes = context.presentation.slides.getItemAt(index).shapes;
            shapes.load("items/name")
            await context.sync()
            for (const shape of shapes.items) {
                if (shape.name === "toc-header-content") {
                    shape.delete();
                    console.log("remove old header")
                }
            }
            shapes.load(["items/name", "items/id"])
            await context.sync()

            if (indicators.length == 0) {
                console.log("empty page")
                return
            }

            const groups: Array<PowerPoint.Shape> = []
            for (let section_index = 0; section_index < indicators.length; section_index++) {
                const indicator = indicators[section_index];

                const color = indicator.active ? "#FFFFFF" : "#B0B4C3"

                const title_shape = shapes.addTextBox(indicator.title)
                title_shape.top = 3
                title_shape.textFrame.textRange.font.size = 12
                title_shape.textFrame.autoSizeSetting = "AutoSizeShapeToFitText"
                title_shape.textFrame.wordWrap=false

                title_shape.name = `toc-header-sec-${section_index}`
                title_shape.textFrame.textRange.font.color = color
                console.log("add section name " + indicator.title)

                const char_list = Array(indicator.total_pages).fill("○")
                if (indicator.active_page >= 0) {
                    char_list[indicator.active_page] = "●"
                }
                const index_shape = shapes.addTextBox(char_list.join(" "))
                index_shape.top = 18
                index_shape.textFrame.textRange.font.size = 7
                index_shape.textFrame.autoSizeSetting = "AutoSizeShapeToFitText"
                index_shape.textFrame.wordWrap=false
                
                index_shape.name = `toc-header-idx-${section_index}`
                console.log("add section index" + index_shape.name)

                index_shape.textFrame.textRange.font.color = color
                console.log("set style")

                // groups.push(title_shape)
                // groups.push(index_shape)

                title_shape.load(["name", "id"])
                index_shape.load(["name", "id"])
                await context.sync();

                const grouped_shape = shapes.addGroup([title_shape, index_shape])
                grouped_shape.name = `toc-header-${section_index}`
                // grouped_shape.textFrame.textRange.font.color = indicator.active ? "#FFFFFF" : "#B0B4C3"
                await context.sync();
                console.log("group & style")
                grouped_shape.load("id")
                groups.push(grouped_shape)


            }

            shapes.load(["items/name", "items/id"])
            await context.sync()
            console.log(groups)
            shapes.addGroup(groups).name = "toc-header-content"
            await context.sync();


        });
    } catch (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log(error.debugInfo);

        }
    }
}


