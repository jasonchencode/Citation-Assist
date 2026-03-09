import * as React from "react";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { analyzeSelection, reselectText } from "../taskpane";

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
        reselectText={reselectText}
        documentId={props.documentId ?? "trustops-handbook-v1"}
        userId={props.userId ?? "candidate_1"}
      />
    </div>
  );
};

export default App;
