/* global PowerPoint console */

export interface PageConfig {
  new_part: boolean;
  new_section: boolean;
  section_name: string;
  skip: boolean;
  hide: boolean;
}

// default config
export const defaultPageConfig: PageConfig = {
  new_part: false,
  new_section: false,
  section_name: "",
  skip: false,
  hide: false,
};

export async function setPageConfig(config: PageConfig) {
  try {
    await PowerPoint.run(async (context) => {
      // const slide = context.presentation.getSelectedSlides().getItemAt(0);
      // set all selected slides
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();
      for (const slide of slides.items) {
        slide.tags.add("INDICATORCONFIG", JSON.stringify(config));
      }
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function getPageConfig(): Promise<PageConfig> {
  let config: PageConfig = defaultPageConfig;
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      slide.load("tags/key, tags/value");
      await context.sync();
      
      for (const tag of slide.tags.items) {
        if (tag.key === "INDICATORCONFIG") {
          config = JSON.parse(tag.value);
          break;
        }
      }
    });
  } catch (error) {
    console.log("Error: " + error);
  }
  return config;
}

export async function getAllPageConfig(): Promise<Array<PageConfig>> {
  const configs: Array<PageConfig> = [];
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      for (const slide of slides.items) {
        slide.load("tags/key, tags/value");
      }
      await context.sync();

      for (const slide of slides.items) {
        let config = defaultPageConfig;
        for (const tag of slide.tags.items) {
          if (tag.key === "INDICATORCONFIG") {
            config = JSON.parse(tag.value);
            break;
          }
        }
        configs.push(config);
      }
    });
  } catch (error) {
    console.log("Error: " + error);
  }
  return configs;
}