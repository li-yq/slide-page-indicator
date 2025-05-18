import * as React from "react";
import { useState } from "react";
import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components";
import { defaultPageConfig, PageConfig } from "../../PageConfig";

/* global HTMLTextAreaElement */

interface PageSettingsProps {
  configSetter: (config: PageConfig) => void;
  configGetter: () => Promise<PageConfig>;
}

const useStyles = makeStyles({
  instructions: {
    // fontWeight: tokens.fontWeightSemibold,
    // marginTop: "20px",
    // marginBottom: "10px",
  },
  textPromptAndInsertion: {
    // display: "flex",
    // flexDirection: "column",
    // alignItems: "center",
  },
  textAreaField: {
    // marginLeft: "20px",
    // marginTop: "30px",
    // marginBottom: "20px",
    // marginRight: "20px",
    // maxWidth: "50%",
  },
});

const PageSettings: React.FC<PageSettingsProps> = (props: PageSettingsProps) => {
  const [config, setConfig] = useState<PageConfig>(defaultPageConfig);

  const handleTextSet = async () => {
    await props.configSetter(config);
  };
  const handleTextGet = async () => {
    setConfig(await props.configGetter());
  };

  const handleTextChange = async (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setConfig(JSON.parse(event.target.value));
  };

  const styles = useStyles();

  return (
    <div className={styles.textPromptAndInsertion}>
      <Field className={styles.textAreaField} size="large" label="Enter text to be inserted into the document.">
        <Textarea size="large" value={JSON.stringify(config)} onChange={handleTextChange} />
      </Field>
      <Field className={styles.instructions}>Click the button to insert text.</Field>
      <Button appearance="primary" disabled={false} size="large" onClick={handleTextSet}>
        Set text
      </Button>
      <Button appearance="primary" disabled={false} size="large" onClick={handleTextGet}>
        Get text
      </Button>
    </div>
  );
};

export default PageSettings;
