import * as React from "react";
import { useState } from "react";
import { Button, Checkbox, Field, Input } from "@fluentui/react-components";
import { defaultPageConfig, PageConfig } from "../../page-config";


interface PageSettingsProps {
  configSetter: (config: PageConfig) => void;
  configGetter: () => Promise<PageConfig>;
}


const PageSettings: React.FC<PageSettingsProps> = (props: PageSettingsProps) => {
  const [config, setConfig] = useState<PageConfig>(defaultPageConfig);

  const handleConfigSet = async () => {
    await props.configSetter(config);
  };
  const handleConfigGet = async () => {
    setConfig(await props.configGetter());
  };


  const handleInputChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const target = event.target;
    const value = target.type === 'checkbox' ? target.checked : target.value;
    const name = target.name;
    setConfig({
      ...config,
      [name]: value
    });
  }

  const handleTextChange = async (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setConfig(JSON.parse(event.target.value));
  };

  return (
    <div>

      <Field label="Start a new Part?">
        <Checkbox checked={config.new_part} name="new_part" onChange={handleInputChange} />
      </Field>
      <Field label="Start a new Section?">
        <Checkbox checked={config.new_section} name="new_section" onChange={handleInputChange} />
      </Field>
      <Field label="Section name">
        <Input value={config.section_name} name="section_name" onChange={handleInputChange} disabled={!config.new_section} />
      </Field>
      <Field label="Skip this page?">
        <Checkbox checked={config.skip} name="skip" onChange={handleInputChange} />
      </Field>
      <Field label="Hide indicators on this page?">
        <Checkbox checked={config.hide} name="hide" onChange={handleInputChange} />
      </Field>

      <Button onClick={handleConfigGet}>Load</Button>
      <Button onClick={handleConfigSet}>Apply</Button>

    </div>
  );
};

export default PageSettings;
