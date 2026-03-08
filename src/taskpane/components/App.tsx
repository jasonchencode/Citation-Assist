import * as React from "react";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { analyzeSelection } from "../taskpane";

interface AppProps {
  title: string;
  documentId?: string;
  userId?: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <TextInsertion
        analyzeSelection={analyzeSelection}
        documentId={props.documentId}
        userId={props.userId}
      />
    </div>
  );
};

export default App;
