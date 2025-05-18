import * as React from "react";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { getText, setText } from "../taskpane";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC = () => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <TextInsertion textGetter={getText} textSetter={setText} />
    </div>
  );
};

export default App;
