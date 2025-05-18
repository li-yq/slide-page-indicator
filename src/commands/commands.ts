/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { generateIndicators, PartPageIndicators } from "../indicators";
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

  const indicator_data = (generateIndicators(await getAllPageConfig()))
  UpdateAllPageIndicator(indicator_data);

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("RefeshIndicators", RefeshIndicators);


async function UpdateAllPageIndicator(all_indicators: Array<PartPageIndicators>) {
  await PowerPoint.run(async (context) => {

    for (let index = 0; index < all_indicators.length; index++) {
      const indicators = all_indicators[index]


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

      if (indicators.length == 0) {
        console.log("empty page")
        return
      }

      const groups: Array<PowerPoint.Shape> = []
      for (let section_index = 0; section_index < indicators.length; section_index++) {
        let indicator = indicators[section_index];
        let title_shape = shapes.addTextBox(indicator.title)
        title_shape.top = 3
        title_shape.textFrame.textRange.font.size = 12
        title_shape.name = `toc-header-sec-${section_index}`
        console.log("add section name "+ indicator.title)

        let char_list = Array(indicator.total_pages).fill("○")
        if (indicator.active_page >= 0) {
          char_list[indicator.active_page] = "●"
        }
        let index_shape = shapes.addTextBox(char_list.join(" "))
        index_shape.top = 18
        index_shape.textFrame.textRange.font.size = 7
        index_shape.name = `toc-header-idx-${section_index}`
        console.log("add section index")

        let grouped_shape = shapes.addGroup([title_shape, index_shape])
        grouped_shape.textFrame.textRange.font.color = indicator.active ? "#FFFFFF" : "#B0B4C3"
        console.log("group & style")
      }

      shapes.addGroup(groups).name = "toc-header-content"
      console.log("group")
      await context.sync();

    }
  });
}


