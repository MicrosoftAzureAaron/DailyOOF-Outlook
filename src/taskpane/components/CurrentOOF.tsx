import React, { useState, useRef } from "react";
import { Button, Spinner } from "@fluentui/react-components";
import { ArrowSyncRegular } from "@fluentui/react-icons";
import { getMailboxSettings } from "../services/graphService";

interface Props {
  connected: boolean;
  updateStatus: (text: string, type?: "info" | "error" | "success") => void;
}

export const CurrentOOF: React.FC<Props> = ({ connected, updateStatus }) => {
  const [busy, setBusy] = useState(false);
  const [state, setState] = useState("-");
  const iframeRef = useRef<HTMLIFrameElement>(null);

  const handleRefresh = async () => {
    if (!connected) {
      updateStatus("Connect on the Actions tab first.", "error");
      return;
    }
    setBusy(true);
    updateStatus("Fetching current OOF message...");
    try {
      const settings = await getMailboxSettings();
      const ars = settings.automaticRepliesSetting;
      setState(ars.status);

      const html = ars.externalReplyMessage || ars.internalReplyMessage
        || "<p style='color:#888;font-family:Segoe UI;'>No OOF message is currently set.</p>";

      if (iframeRef.current) {
        const doc = iframeRef.current.contentDocument;
        if (doc) {
          doc.open();
          doc.write(html);
          doc.close();
        }
      }
      updateStatus("Current OOF message loaded", "success");
    } catch (err: unknown) {
      updateStatus(`Failed: ${(err as Error).message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  const stateClassName = state === "scheduled" ? "scheduled"
    : state === "alwaysEnabled" ? "enabled"
    : state === "disabled" ? "disabled" : "";

  return (
    <div>
      <div className="section">
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div className="section-title">Current OOF Message</div>
          <Button
            size="small"
            icon={busy ? <Spinner size="tiny" /> : <ArrowSyncRegular />}
            onClick={handleRefresh}
            disabled={busy || !connected}
          >
            Refresh
          </Button>
        </div>
        <div className="status-grid" style={{ marginBottom: 8 }}>
          <span className="status-label">State:</span>
          <span className={`status-value ${stateClassName}`}>{state}</span>
        </div>
        <iframe
          ref={iframeRef}
          className="preview-frame"
          style={{ minHeight: 300 }}
          title="Current OOF Preview"
          sandbox="allow-same-origin"
        />
      </div>
    </div>
  );
};
