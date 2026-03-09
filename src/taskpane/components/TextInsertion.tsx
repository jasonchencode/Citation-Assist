import * as React from "react";
import {
  Button,
  Field,
  tokens,
  makeStyles,
  Link,
} from "@fluentui/react-components";
import type { AnalyzeSelectionResult } from "../taskpane";
import type { AnalyzeResponse } from "../api";
import { getDocument } from "../api";

export interface CitationEntry {
  id: string;
  text: string;
  citation: AnalyzeResponse;
}

interface TextInsertionProps {
  analyzeSelection: (options?: { documentId?: string; userId?: string }) => Promise<AnalyzeSelectionResult>;
  reselectText: (text: string) => void | Promise<void>;
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
  citationList: {
    marginTop: "12px",
    width: "100%",
    maxHeight: "280px",
    overflowY: "auto",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  citationCard: {
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: "4px",
    maxWidth: "100%",
    flexShrink: 0,
  },
  citationTitle: {
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: "4px",
  },
  cardActions: {
    marginTop: "8px",
    display: "flex",
    gap: "8px",
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

let nextId = 0;
function generateId(): string {
  nextId += 1;
  return `citation-${nextId}`;
}

const TextInsertion: React.FC<TextInsertionProps> = (props: TextInsertionProps) => {
  const [isAnalyzing, setIsAnalyzing] = React.useState(false);
  const [citations, setCitations] = React.useState<CitationEntry[]>([]);
  const [lastError, setLastError] = React.useState<string | null>(null);
  const [documentContent, setDocumentContent] = React.useState<unknown>(null);
  const [loadingDoc, setLoadingDoc] = React.useState(false);

  const handleAnalyzeSelection = async () => {
    setIsAnalyzing(true);
    setLastError(null);
    try {
      const value = await props.analyzeSelection({
        documentId: props.documentId,
        userId: props.userId,
      });
      if (value.citation) {
        setCitations((prev) => [
          { id: generateId(), text: value.text, citation: value.citation! },
          ...prev,
        ]);
      } else if (value.error) {
        setLastError(value.error);
      } else if (!value.text) {
        setLastError("No text selected.");
      }
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleRemove = (id: string) => {
    setCitations((prev) => prev.filter((e) => e.id !== id));
  };

  const handleReselect = (text: string) => {
    void Promise.resolve(props.reselectText(text));
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

      {lastError && (
        <div className={styles.error}>Analysis failed: {lastError}</div>
      )}

      {citations.length > 0 && (
        <div className={styles.citationList}>
          {citations.map((entry) => (
            <div key={entry.id} className={styles.citationCard}>
              <div className={styles.citationTitle}>Citation</div>
              <div>{entry.citation.citation_text}</div>
              <div style={{ marginTop: "4px", fontSize: tokens.fontSizeBase200 }}>
                Confidence: {(entry.citation.confidence * 100).toFixed(0)}%
              </div>
              {entry.citation.url && (
                <Link
                  href={entry.citation.url}
                  target="_blank"
                  rel="noopener noreferrer"
                  style={{ marginTop: "4px", display: "block" }}
                >
                  Source
                </Link>
              )}
              <div className={styles.cardActions}>
                <Button
                  appearance="secondary"
                  size="small"
                  onClick={() => handleReselect(entry.text)}
                >
                  Re-select
                </Button>
                <Button
                  appearance="secondary"
                  size="small"
                  onClick={() => handleRemove(entry.id)}
                >
                  Remove
                </Button>
              </div>
            </div>
          ))}
        </div>
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
