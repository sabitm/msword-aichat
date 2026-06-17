import { Button, Text } from "@fluentui/react-components";
import type { PendingEdit } from "../../types/agent";

interface EditPreviewProps {
  edit: PendingEdit;
  disabled?: boolean;
  onApply: () => void;
  onReject: () => void;
}

export function EditPreview({ edit, disabled, onApply, onReject }: EditPreviewProps) {
  const isPending = edit.status === "pending";

  return (
    <div className="edit-preview">
      <Text size={200} weight="semibold">
        {edit.description}
      </Text>
      <div className="edit-preview-grid">
        <div className="edit-preview-pane">
          <Text size={100} weight="semibold">
            Before
          </Text>
          <pre>{edit.before || "(empty)"}</pre>
        </div>
        <div className="edit-preview-pane">
          <Text size={100} weight="semibold">
            After
          </Text>
          <pre>{edit.after}</pre>
        </div>
      </div>
      {isPending ? (
        <div className="edit-preview-actions">
          <Button appearance="primary" disabled={disabled} onClick={onApply}>
            Apply
          </Button>
          <Button appearance="secondary" disabled={disabled} onClick={onReject}>
            Reject
          </Button>
        </div>
      ) : (
        <Text size={200}>
          {edit.status === "applied" ? "Edit applied." : "Edit rejected."}
        </Text>
      )}
    </div>
  );
}