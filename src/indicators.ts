import { PageConfig } from "./page-config";

export type PartPageIndicators = Array<SectionPageIndicators>;

export interface SectionPageIndicators {
    title: string;
    active: boolean;
    total_pages: number;
    active_page: number;
}

export function generateIndicators(page_configs: Array<PageConfig>): Array<PartPageIndicators> {
    const parts: Array<Array<PageConfig>> = [];
    let currentPart: Array<PageConfig> = [];

    for (const config of page_configs) {
        if (config.new_part && currentPart.length > 0) {
            parts.push(currentPart);
            currentPart = [];
        }
        currentPart.push(config);
    }
    if (currentPart.length > 0) {
        parts.push(currentPart);
    }

    return parts.flatMap(generatePartIndicators);
}

function generatePartIndicators(page_configs: Array<PageConfig>): Array<PartPageIndicators> {
    const section: Array<Array<PageConfig>> = [];
    let currentSection: Array<PageConfig> = [];

    for (const config of page_configs) {
        if (config.new_section && currentSection.length > 0) {
            section.push(currentSection);
            currentSection = [];
        }
        currentSection.push(config);
    }
    if (currentSection.length > 0) {
        section.push(currentSection);
    }

    const section_all_indicators = section.map(generateSectionIndicators)
    const inactive_section_indicators: Array<SectionPageIndicators> = []
    for (const section_all_indicator of section_all_indicators) {
        inactive_section_indicators.push({
            ...section_all_indicator[0],
            active: false,
            active_page: -1,
        })
    }

    const indicators: Array<PartPageIndicators> = [];
    // iterate over each section
    for (const current_section_index in section_all_indicators) {
        const current_section_indicators = section_all_indicators[current_section_index]
        // then iter over slide page
        for (let page_idx_in_section = 0; page_idx_in_section < current_section_indicators.length; page_idx_in_section++) {
            let current_page_indicators: PartPageIndicators = []
            for (const idx in section) {
                if (idx == current_section_index) {
                    current_page_indicators.push(section_all_indicators[current_section_index][page_idx_in_section])
                } else {
                    current_page_indicators.push(inactive_section_indicators[idx])
                }
            }
            indicators.push(current_page_indicators)
        }
    }
    return indicators;
}


function generateSectionIndicators(page_configs: Array<PageConfig>): Array<SectionPageIndicators> {
    const section_name = page_configs[0].section_name;
    const total_pages = page_configs.filter((p) => !p.skip).length;
    const indicators: Array<SectionPageIndicators> = [];
    let currentPage = 0;
    for (const page of page_configs) {
        if (!page.skip) {
            indicators.push({
                title: section_name,
                active: true,
                total_pages: total_pages,
                active_page: currentPage,
            });
            currentPage += 1;
        } else {
            indicators.push({
                title: section_name,
                active: true,
                total_pages: total_pages,
                active_page: -1,
            });
        }
    }
    return indicators;
}
