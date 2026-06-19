import { DefaultButton, PrimaryButton, Text } from "@fluentui/react";
import * as React from "react";
import type { PendingEdit } from "../../types/agent";

interface EditPreviewProps {
  edit: PendingEdit;
  disabled?: boolean;
  onApply: () => void;
  onReject: () => void;
  onUndo?: () => void;
}

export function EditPreview(props: EditPreviewProps): React.ReactElement {
  var isPending = props.edit.status === "pending";
  var canUndo = props.edit.status === "applied" && Boolean(props.edit.undo);

  return (
    <div className="edit-preview">
      <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
        {props.edit.description}
      </Text>
      <div className="edit-preview-grid">
        <div className="edit-preview-pane">
          <Text variant="small" block styles={{ root: { fontWeight: 600 } }}>
            Before
          </Text>
          <pre>{props.edit.before || "(empty)"}</pre>
        </div>
        <div className="edit-preview-pane">
          <Text variant="small" block styles={{ root: { fontWeight: 600 } }}>
            After
          </Text>
          <pre>{props.edit.after}</pre>
        </div>
      </div>
      {isPending ? (
        <div className="edit-preview-actions">
          <PrimaryButton disabled={props.disabled} onClick={props.onApply}>
            Apply
          </PrimaryButton>
          <DefaultButton disabled={props.disabled} onClick={props.onReject}>
            Reject
          </DefaultButton>
        </div>
      ) : (
        <div className="edit-preview-actions">
          <Text variant="small">
            {props.edit.status === "applied"
              ? "Edit applied."
              : props.edit.status === "undone"
                ? "Edit undone."
                : "Edit rejected."}
          </Text>
          {canUndo && props.onUndo ? (
            <DefaultButton disabled={props.disabled} onClick={props.onUndo}>
              Undo
            </DefaultButton>
          ) : null}
        </div>
      )}
    </div>
  );
}