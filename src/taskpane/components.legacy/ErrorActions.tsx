import { DefaultButton, MessageBar, MessageBarType } from "@fluentui/react";
import * as React from "react";

interface ErrorActionsProps {
  error: string;
  disabled?: boolean;
  onRetry?: () => void;
}

function copyToClipboard(text: string): boolean {
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text);
      return true;
    }
  } catch (_error) {
    // Fall through to execCommand.
  }

  var textarea = document.createElement("textarea");
  textarea.value = text;
  textarea.style.position = "fixed";
  textarea.style.left = "-9999px";
  document.body.appendChild(textarea);
  textarea.select();
  var ok = false;
  try {
    ok = document.execCommand("copy");
  } catch (_execError) {
    ok = false;
  }
  document.body.removeChild(textarea);
  return ok;
}

export function ErrorActions(props: ErrorActionsProps): React.ReactElement {
  var _a = React.useState(false);
  var copied = _a[0];
  var setCopied = _a[1];

  function handleCopy(): void {
    if (copyToClipboard(props.error)) {
      setCopied(true);
      window.setTimeout(function () {
        setCopied(false);
      }, 2000);
    }
  }

  return (
    <MessageBar
      messageBarType={MessageBarType.error}
      className="error-actions"
      actions={
        <div className="error-actions-buttons" style={{ display: "flex", flexWrap: "wrap" }}>
          {props.onRetry ? (
            <DefaultButton
              disabled={props.disabled}
              iconProps={{ iconName: "Refresh" }}
              onClick={props.onRetry}
            >
              Retry
            </DefaultButton>
          ) : null}
          <DefaultButton iconProps={{ iconName: "Copy" }} onClick={handleCopy}>
            {copied ? "Copied" : "Copy"}
          </DefaultButton>
        </div>
      }
    >
      <strong>Error:</strong> {props.error}
    </MessageBar>
  );
}