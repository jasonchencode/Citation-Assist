import * as React from "react";
import {
  Button,
  Field,
  tokens,
  makeStyles,
  Spinner,
  Link,
} from "@fluentui/react-components";
import type { AnalyzeSelectionResult } from "../taskpane";
import { getHealth, getDocument } from "../api";

interface TextInsertionProps {
  analyzeSelection: (options?: { documentId?: string; userId?: string }) => Promise<AnalyzeSelectionResult>;
  documentId?: string;
  userId?: string;
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
  statusRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    marginTop: "8px",
    marginBottom: "8px",
  },
  citation: {
    marginTop: "12px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: "4px",
    maxWidth: "100%",
  },
  citationTitle: {
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: "4px",
  },
  error: {
    color: tokens.colorPaletteRedForeground1,
    marginTop: "8px",
  },
  documentPreview: {
    marginTop: "8px",
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground4,
    borderRadius: "4px",
    maxHeight: "120px",
    overflow: "auto",
    fontSize: tokens.fontSizeBase200,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
});

const TextInsertion: React.FC<TextInsertionProps> = (props: TextInsertionProps) => {
  const [isAnalyzing, setIsAnalyzing] = React.useState(false);
  const [result, setResult] = React.useState<AnalyzeSelectionResult | null>(null);
  const [health, setHealth] = React.useState<{ ok: boolean; status: number; error?: string } | null>(null);
  const [documentContent, setDocumentContent] = React.useState<unknown>(null);
  const [loadingDoc, setLoadingDoc] = React.useState(false);

  React.useEffect(() => {
    getHealth().then(setHealth);
  }, []);

  const handleAnalyzeSelection = async () => {
    setIsAnalyzing(true);
    setResult(null);
    try {
      const value = await props.analyzeSelection({
        documentId: props.documentId,
        userId: props.userId,
      });
      setResult(value);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleViewDocument = async () => {
    setLoadingDoc(true);
    setDocumentContent(null);
    try {
      const res = await getDocument();
      setDocumentContent(res.body ?? res.error ?? { status: res.status });
    } finally {
      setLoadingDoc(false);
    }
  };

  const styles = useStyles();

  return (
    <div className={styles.analyzeSection}>
      {/* Backend health */}
      <Field className={styles.statusRow} label="Backend">
        {health === null ? (
          <Spinner size="tiny" />
        ) : health.ok ? (
          <span style={{ color: tokens.colorPaletteGreenForeground1 }}>Available</span>
        ) : (
          <span className={styles.error}>
            Unavailable {health.status ? `(${health.status})` : ""} {health.error ?? ""}
          </span>
        )}
      </Field>

      <Field className={styles.instructions}>
        Select text in the document, then click the button to analyze it and show the citation.
      </Field>
      <Button
        appearance="primary"
        disabled={isAnalyzing}
        size="large"
        onClick={handleAnalyzeSelection}
      >
        {isAnalyzing ? "Analyzing…" : "Analyze Selection"}
      </Button>

      {/* Last analyze result */}
      {result?.citation && (
        <div className={styles.citation}>
          <div className={styles.citationTitle}>Citation</div>
          <div>{result.citation.citation_text}</div>
          <div style={{ marginTop: "4px", fontSize: tokens.fontSizeBase200 }}>
            Confidence: {(result.citation.confidence * 100).toFixed(0)}%
          </div>
          {result.citation.url && (
            <Link href={result.citation.url} target="_blank" rel="noopener noreferrer" style={{ marginTop: "4px" }}>
              Source
            </Link>
          )}
        </div>
      )}
      {result?.error && (
        <div className={styles.error}>Analysis failed: {result.error}</div>
      )}

      {/* Debug: document content */}
      <Field className={styles.instructions} style={{ marginTop: "20px" }}>
        Debug: backend document content
      </Field>
      <Button appearance="secondary" size="small" onClick={handleViewDocument} disabled={loadingDoc}>
        {loadingDoc ? "Loading…" : "View document"}
      </Button>
      {documentContent !== null && (
        <pre className={styles.documentPreview}>
          {typeof documentContent === "string"
            ? documentContent
            : JSON.stringify(documentContent, null, 2)}
        </pre>
      )}
    </div>
  );
};

export default TextInsertion;
