import React, { useState, useCallback, useRef } from "react";
import { Button, Select, Spinner, TabList, Tab } from "@fluentui/react-components";
import {
  ArrowDownloadRegular,
  SendRegular,
  SaveRegular,
  ArrowUploadRegular,
} from "@fluentui/react-icons";
import { TEMPLATES } from "../types";
import type { AppConfig } from "../types";
import { loadTemplate, resolvePlaceholders } from "../services/templateService";
import { setAutoReplyMessage, getMailboxSettings } from "../services/graphService";

interface Props {
  config: AppConfig;
  updateConfig: (patch: Partial<AppConfig>) => void;
  userEmail: string;
  connected: boolean;
  updateStatus: (text: string, type?: "info" | "error" | "success") => void;
}

export const MessageTemplates: React.FC<Props> = ({
  config,
  updateConfig,
  userEmail,
  connected,
  updateStatus,
}) => {
  const [selectedTemplate, setSelectedTemplate] = useState(TEMPLATES[0].id);
  const [messageHtml, setMessageHtml] = useState("");
  const [subTab, setSubTab] = useState<"edit" | "preview">("edit");
  const [busy, setBusy] = useState(false);
  const previewRef = useRef<HTMLIFrameElement>(null);

  // Load and resolve a template
  const handleLoadTemplate = useCallback(async () => {
    setBusy(true);
    updateStatus("Loading template...");
    try {
      const raw = await loadTemplate(selectedTemplate);
      const resolved = resolvePlaceholders(raw, config, null, userEmail);
      setMessageHtml(resolved);
      updateStatus(`Template loaded: ${TEMPLATES.find((t) => t.id === selectedTemplate)?.name}`, "success");
    } catch (err: unknown) {
      updateStatus(`Failed to load template: ${(err as Error).message}`, "error");
    } finally {
      setBusy(false);
    }
  }, [selectedTemplate, config, userEmail, updateStatus]);

  // Apply message to Exchange
  const handleApply = async (scope: "internal" | "external" | "both") => {
    if (!messageHtml.trim()) {
      updateStatus("No message to apply. Load a template first.", "error");
      return;
    }
    setBusy(true);
    updateStatus(`Applying ${scope} message...`);
    try {
      await setAutoReplyMessage(messageHtml, scope);
      updateStatus(`${scope.charAt(0).toUpperCase() + scope.slice(1)} message applied`, "success");
    } catch (err: unknown) {
      updateStatus(`Failed: ${(err as Error).message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  // Save online message to editor
  const handleFetchOnline = async () => {
    setBusy(true);
    updateStatus("Fetching current online message...");
    try {
      const settings = await getMailboxSettings();
      const msg = settings.automaticRepliesSetting.externalReplyMessage || settings.automaticRepliesSetting.internalReplyMessage || "";
      if (!msg) {
        updateStatus("No OOF message is currently set online.", "info");
      } else {
        setMessageHtml(msg);
        updateStatus("Online message loaded into editor", "success");
      }
    } catch (err: unknown) {
      updateStatus(`Failed: ${(err as Error).message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  // Update preview when switching to preview tab
  const handleSubTabChange = (tab: "edit" | "preview") => {
    setSubTab(tab);
    if (tab === "preview" && previewRef.current) {
      const doc = previewRef.current.contentDocument;
      if (doc) {
        doc.open();
        doc.write(messageHtml || "<p style='color:#888;font-family:Segoe UI;'>No message to preview.</p>");
        doc.close();
      }
    }
  };

  return (
    <div>
      {/* Template Selector */}
      <div className="section">
        <div className="section-title">Template</div>
        <div className="field-row">
          <Select
            value={selectedTemplate}
            onChange={(_, data) => setSelectedTemplate(data.value)}
            style={{ flex: 1 }}
          >
            {TEMPLATES.map((t) => (
              <option key={t.id} value={t.id}>{t.name}</option>
            ))}
          </Select>
          <Button
            appearance="primary"
            icon={<ArrowDownloadRegular />}
            onClick={handleLoadTemplate}
            disabled={busy}
          >
            {busy ? <Spinner size="tiny" /> : "Load"}
          </Button>
        </div>
      </div>

      {/* Edit / Preview Tabs */}
      <div className="section">
        <TabList
          selectedValue={subTab}
          onTabSelect={(_, data) => handleSubTabChange(data.value as "edit" | "preview")}
          size="small"
        >
          <Tab value="edit">Edit</Tab>
          <Tab value="preview">Preview</Tab>
        </TabList>

        {subTab === "edit" ? (
          <textarea
            className="message-editor"
            value={messageHtml}
            onChange={(e) => setMessageHtml(e.target.value)}
            placeholder="Load a template or type HTML here..."
          />
        ) : (
          <iframe
            ref={previewRef}
            className="preview-frame"
            title="Message Preview"
            sandbox="allow-same-origin"
          />
        )}
      </div>

      {/* Apply Buttons */}
      <div className="section">
        <div className="section-title">Apply Message</div>
        <div className="button-row">
          <Button
            size="small"
            icon={<SendRegular />}
            onClick={() => handleApply("internal")}
            disabled={busy || !connected}
          >
            Internal
          </Button>
          <Button
            size="small"
            icon={<SendRegular />}
            onClick={() => handleApply("external")}
            disabled={busy || !connected}
          >
            External
          </Button>
          <Button
            size="small"
            appearance="primary"
            icon={<SendRegular />}
            onClick={() => handleApply("both")}
            disabled={busy || !connected}
          >
            Both
          </Button>
        </div>
      </div>

      {/* Fetch Online */}
      <div className="section">
        <Button
          size="small"
          appearance="subtle"
          icon={<ArrowUploadRegular />}
          onClick={handleFetchOnline}
          disabled={busy || !connected}
        >
          Import Current Online Message
        </Button>
      </div>
    </div>
  );
};
