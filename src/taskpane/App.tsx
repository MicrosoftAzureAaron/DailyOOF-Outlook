import React, { useState, useCallback, useEffect } from "react";
import { TabList, Tab, Spinner } from "@fluentui/react-components";
import {
  FlashRegular,
  SettingsRegular,
  MailTemplateRegular,
  MailReadRegular,
} from "@fluentui/react-icons";
import { QuickActions } from "./components/QuickActions";
import { Configuration } from "./components/Configuration";
import { MessageTemplates } from "./components/MessageTemplates";
import { CurrentOOF } from "./components/CurrentOOF";
import { loadConfig, saveConfig } from "./services/configService";
import { isSignedIn } from "./services/authService";
import type { AppConfig } from "./types";
import { DEFAULT_CONFIG } from "./types";
import "./styles/App.css";

type TabId = "quick" | "config" | "templates" | "current";

export const App: React.FC = () => {
  const [activeTab, setActiveTab] = useState<TabId>("quick");
  const [config, setConfig] = useState<AppConfig>(DEFAULT_CONFIG);
  const [status, setStatus] = useState<{ text: string; type: "info" | "error" | "success" }>({ text: "Ready", type: "info" });
  const [connected, setConnected] = useState(false);
  const [userEmail, setUserEmail] = useState("");
  const [loading, setLoading] = useState(true);

  // Load config on mount
  useEffect(() => {
    try {
      const saved = loadConfig();
      setConfig(saved);
    } catch {
      // use defaults
    }
    isSignedIn().then(setConnected).catch(() => {}).finally(() => setLoading(false));
  }, []);

  // Persist config changes
  const updateConfig = useCallback(async (patch: Partial<AppConfig>) => {
    setConfig((prev) => {
      const next = { ...prev, ...patch };
      saveConfig(next).catch(() => {});
      return next;
    });
  }, []);

  const updateStatus = useCallback((text: string, type: "info" | "error" | "success" = "info") => {
    setStatus({ text, type });
  }, []);

  if (loading) {
    return (
      <div style={{ display: "flex", justifyContent: "center", alignItems: "center", height: "100vh" }}>
        <Spinner label="Loading..." />
      </div>
    );
  }

  return (
    <div className="app-container">
      <TabList
        selectedValue={activeTab}
        onTabSelect={(_, data) => setActiveTab(data.value as TabId)}
        size="small"
      >
        <Tab value="quick" icon={<FlashRegular />}>Actions</Tab>
        <Tab value="config" icon={<SettingsRegular />}>Config</Tab>
        <Tab value="templates" icon={<MailTemplateRegular />}>Templates</Tab>
        <Tab value="current" icon={<MailReadRegular />}>Current</Tab>
      </TabList>

      <div className="app-content">
        {activeTab === "quick" && (
          <QuickActions
            config={config}
            connected={connected}
            setConnected={setConnected}
            userEmail={userEmail}
            setUserEmail={setUserEmail}
            updateStatus={updateStatus}
          />
        )}
        {activeTab === "config" && (
          <Configuration config={config} updateConfig={updateConfig} updateStatus={updateStatus} />
        )}
        {activeTab === "templates" && (
          <MessageTemplates
            config={config}
            updateConfig={updateConfig}
            userEmail={userEmail}
            connected={connected}
            updateStatus={updateStatus}
          />
        )}
        {activeTab === "current" && (
          <CurrentOOF connected={connected} updateStatus={updateStatus} />
        )}
      </div>

      <div className={`status-bar ${status.type}`}>{status.text}</div>
    </div>
  );
};
