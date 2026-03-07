import * as React from "react";
import { Button, Field, tokens, makeStyles } from "@fluentui/react-components";

interface TextInsertionProps {
  analyzeSelection: () => Promise<void>;
}

const useStyles = makeStyles({
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
  },
  analyzeSection: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
});

const TextInsertion: React.FC<TextInsertionProps> = (props: TextInsertionProps) => {
  const [isAnalyzing, setIsAnalyzing] = React.useState(false);

  const handleAnalyzeSelection = async () => {
    setIsAnalyzing(true);
    try {
      await props.analyzeSelection();
    } finally {
      setIsAnalyzing(false);
    }
  };

  const styles = useStyles();

  return (
    <div className={styles.analyzeSection}>
      <Field className={styles.instructions}>
        Select text in the document, then click the button to highlight it and log it to the console.
      </Field>
      <Button
        appearance="primary"
        disabled={isAnalyzing}
        size="large"
        onClick={handleAnalyzeSelection}
      >
        {isAnalyzing ? "Analyzing…" : "Analyze Selection"}
      </Button>
    </div>
  );
};

export default TextInsertion;
