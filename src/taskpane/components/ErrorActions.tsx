import { Button, MessageBar, MessageBarActions, MessageBarBody } from "@fluentui/react-components";
import { ArrowClockwise24Regular, Copy24Regular } from "@fluentui/react-icons";
import { useState } from "react";

interface ErrorActionsProps {
  error: string;
  disabled?: boolean;
  onRetry?: () => void;
}

export function ErrorActions({ error, disabled, onRetry }: ErrorActionsProps) {
  const [copied, setCopied] = useState(false);

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(error);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 2000);
    } catch {
      setCopied(false);
    }
  };

  return (
    <MessageBar intent="error" className="error-actions">
      <MessageBarBody>
        <strong>Error:</strong> {error}
      </MessageBarBody>
      <MessageBarActions
        containerAction={
          <>
            {onRetry ? (
              <Button
                appearance="secondary"
                size="small"
                icon={<ArrowClockwise24Regular />}
                disabled={disabled}
                onClick={onRetry}
              >
                Retry
              </Button>
            ) : null}
            <Button
              appearance="secondary"
              size="small"
              icon={<Copy24Regular />}
              onClick={() => void handleCopy()}
            >
              {copied ? "Copied" : "Copy"}
            </Button>
          </>
        }
      />
    </MessageBar>
  );
}