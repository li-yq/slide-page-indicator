/* global PowerPoint console */

interface PageConfig {
  text: string;
  fontSize: number;
}

export async function setPageConfig(config: PageConfig) {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      // save the text as slide tag
      slide.tags.add("INDICATORCONFIG", JSON.stringify(config));
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function getPageConfig(): Promise<PageConfig> {
  let config: PageConfig = { text: "NAN.", fontSize: 12 };
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

export async function setText(text: string) {
  return setPageConfig({ text: text, fontSize: 12 });
}

export async function getText(): Promise<string> {
  const config = await getPageConfig();
  return config.text;
}