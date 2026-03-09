import * as React from "react";
import {
  Button,
  Field,
  tokens,
  makeStyles,
  Link,
} from "@fluentui/react-components";
import { removeCitation, checkCitationExists } from "../taskpane";
import type { AnalyzeSelectionResult } from "../taskpane";
import type { AnalyzeResponse } from "../api";

export interface CitationEntry {
  id: string;
  text: string;
  citation: AnalyzeResponse;
  highlightColor?: string;
}

interface TextInsertionProps {
  analyzeSelection: (options?: { documentId?: string; userId?: string }) => Promise<AnalyzeSelectionResult>;
  reselectText: (text: string) => void | Promise<void>;
  documentId?: string;
  userId?: string;
}

const useStyles = makeStyles({
  analyzeSection: {
    display: "flex",
    flexDirection: "column",
    alignItems: "stretch",
    padding: "24px 20px 24px",
    boxSizing: "border-box",
    width: "100%",
    height: "100%",
  },
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "4px",
    marginBottom: "12px",
    textAlign: "left",
    width: "100%",
  },
  analyzeButton: {
    alignSelf: "flex-start",
    marginBottom: "8px",
  },
  error: {
    color: tokens.colorPaletteRedForeground1,
    marginTop: "4px",
    marginBottom: "8px",
    alignSelf: "flex-start",
  },
  refreshButton: {
    alignSelf: "flex-start",
    marginTop: "8px",
    marginBottom: "8px",
  },
  citationList: {
    marginTop: "4px",
    width: "100%",
    flex: 1,
    overflowY: "auto",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    paddingBottom: "24px",
    boxSizing: "border-box",
  },
  citationCard: {
    padding: "16px 18px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: "6px",
    maxWidth: "100%",
    flexShrink: 0,
    boxShadow: tokens.shadow4,
  },
  citationTitle: {
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: "6px",
  },
  citationHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    marginBottom: "6px",
  },
  colorDot: {
    width: "12px",
    height: "12px",
    borderRadius: "999px",
    display: "inline-block",
    marginLeft: "8px",
    flexShrink: 0,
  },
  citationText: {
    fontSize: tokens.fontSizeBase300,
    lineHeight: "1.4",
    marginBottom: "6px",
  },
  citationMeta: {
    marginTop: "2px",
    fontSize: tokens.fontSizeBase200,
  },
  cardActions: {
    marginTop: "10px",
    display: "flex",
    gap: "8px",
    flexWrap: "wrap",
  },
});

let nextId = 0;
function generateId(): string {
  nextId += 1;
  return `citation-${nextId}`;
}

const highlightColorMap: Record<string, string> = {
  Yellow: "#f8f803",
  Turquoise: "#63e7e8",
  Pink: "#ea07fd",
  Blue: "#010bfd",
  Red: "#db0906",
  DarkBlue: "#00038b",
  Teal: "#348382",
  Green: "#3a6104",
  Violet: "#800080",
};

const TextInsertion: React.FC<TextInsertionProps> = (props: TextInsertionProps) => {
  const [isAnalyzing, setIsAnalyzing] = React.useState(false);
  const [citations, setCitations] = React.useState<CitationEntry[]>([]);
  const [lastError, setLastError] = React.useState<string | null>(null);
  const [isRefreshing, setIsRefreshing] = React.useState(false);

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
          {
            id: generateId(),
            text: value.text,
            citation: value.citation!,
            highlightColor: value.highlightColor,
          },
          ...prev,
        ]);
        void Promise.resolve(props.reselectText(value.text));
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

  const handleRefresh = async () => {
    if (citations.length === 0) return;
    setIsRefreshing(true);
    try {
      const stillExist = await Promise.all(
        citations.map(async (entry) => {
          try {
            return await checkCitationExists(entry.text);
          } catch {
            return true;
          }
        })
      );
      setCitations((prev) => prev.filter((_, i) => stillExist[i]));
    } catch {
      // Silently ignore.
    } finally {
      setIsRefreshing(false);
    }
  };

  const styles = useStyles();

  return (
    <div className={styles.analyzeSection}>
      <Field className={styles.instructions}>
        Select text in the document, then click the button to analyze it and show the citation.
      </Field>
      <Button
        className={styles.analyzeButton}
        appearance="primary"
        disabled={isAnalyzing}
        size="large"
        onClick={handleAnalyzeSelection}
      >
        {isAnalyzing ? "Analyzing…" : "Analyze Selection"}
      </Button>

      {lastError && (
        <div className={styles.error}>
          <span>{lastError}</span>
          <Button
            appearance="secondary"
            size="small"
            onClick={handleAnalyzeSelection}
            style={{ marginLeft: "8px" }}
          >
            Retry
          </Button>
        </div>
      )}

      {citations.length > 0 && (
        <>
          <Button
            className={styles.refreshButton}
            appearance="secondary"
            size="medium"
            onClick={handleRefresh}
            disabled={isRefreshing}
          >
            {isRefreshing ? "Refreshing…" : "Refresh"}
          </Button>
          <div className={styles.citationList}>
          {citations.map((entry) => (
            <div key={entry.id} className={styles.citationCard}>
              <div className={styles.citationHeader}>
                <div className={styles.citationTitle}>Citation</div>
                {entry.highlightColor && (
                  <span
                    className={styles.colorDot}
                    style={{
                      backgroundColor:
                        highlightColorMap[entry.highlightColor] ?? "#FFE066",
                    }}
                  />
                )}
              </div>
              <div className={styles.citationText}>{entry.citation.citation_text}</div>
              <div className={styles.citationMeta}>
                Confidence: {(entry.citation.confidence * 100).toFixed(0)}%
              </div>
              {entry.citation.url && (
                <Link
                  href={entry.citation.url}
                  target="_blank"
                  rel="noopener noreferrer"
                  style={{ marginTop: "4px", display: "inline-block" }}
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
                  onClick={() => {
                    try {
                      void Promise.resolve(removeCitation(entry.text)).catch(() => {});
                    } catch {
                      // Ignore.
                    } finally {
                      handleRemove(entry.id);
                    }
                  }}
                >
                  Remove
                </Button>
              </div>
            </div>
          ))}
          </div>
        </>
      )}
    </div>
  );
};

export default TextInsertion;
