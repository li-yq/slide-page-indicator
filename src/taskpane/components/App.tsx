import * as React from "react";
import { makeStyles } from "@fluentui/react-components";
import { getPageConfig, setPageConfig } from "../../PageConfig";
import PageSettings from "./PageSettings";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC = () => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <PageSettings configGetter={getPageConfig} configSetter={setPageConfig} />
    </div>
  );
};

export default App;
