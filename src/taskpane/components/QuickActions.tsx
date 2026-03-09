import React, { useState } from "react";
import { Button, Spinner } from "@fluentui/react-components";
import {
  PlugConnectedRegular,
  PlugDisconnectedRegular,
  CalendarClockRegular,
  ArrowSyncRegular,
  BeachRegular,
  MailReadRegular,
} from "@fluentui/react-icons";
import { getAccessToken, signOut } from "../services/authService";
import {
  getUserProfile,
  getMailboxSettings,
  setAutoReplyStatus,
  setScheduledAutoReply,
  calculateScheduleWindow,
  calculateVacationWindow,
} from "../services/graphService";
import type { AppConfig, AutoReplyStatus } from "../types";

interface Props {
  config: AppConfig;
  connected: boolean;
  setConnected: (v: boolean) => void;
  userEmail: string;
  setUserEmail: (v: string) => void;
  updateStatus: (text: string, type?: "info" | "error" | "success") => void;
}

export const QuickActions: React.FC<Props> = ({
  config,
  connected,
  setConnected,
  userEmail,
  setUserEmail,
  updateStatus,
}) => {
  const [busy, setBusy] = useState(false);
  const [arcState, setArcState] = useState<string>("-");
  const [arcStart, setArcStart] = useState<string>("-");
  const [arcEnd, setArcEnd] = useState<string>("-");
  const [returnDate, setReturnDate] = useState("");

  // ── Connect ───────────────────────────────────────────────
  const handleConnect = async () => {
    setBusy(true);
    updateStatus("Connecting...");
    try {
      await getAccessToken();
      const profile = await getUserProfile();
      setUserEmail(profile.mail || profile.userPrincipalName);
      setConnected(true);
      updateStatus(`Connected as ${profile.displayName}`, "success");

      // Fetch current status on connect
      const settings = await getMailboxSettings();
      const ars = settings.automaticRepliesSetting;
      setArcState(ars.status);
      setArcStart(ars.scheduledStartDateTime?.dateTime || "-");
      setArcEnd(ars.scheduledEndDateTime?.dateTime || "-");
    } catch (err: unknown) {
      updateStatus(`Connection failed: ${(err as Error).message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  const handleDisconnect = async () => {
    try {
      await signOut();
      setConnected(false);
      setUserEmail("");
      updateStatus("Disconnected");
    } catch {
      updateStatus("Disconnect failed", "error");
    }
  };

  // ── Enable Scheduled ──────────────────────────────────────
  const handleEnableScheduled = async () => {
    if (!config.startOfShift || !config.endOfShift) {
      updateStatus("Configure shift hours on the Config tab first.", "error");
      return;
    }
    setBusy(true);
    updateStatus("Setting scheduled auto reply...");
    try {
      const { startTime, endTime } = calculateScheduleWindow(
        config.startOfShift, config.endOfShift, config.workDays
      );
      await setScheduledAutoReply(startTime, endTime);
      setArcState("scheduled");
      setArcStart(startTime.toLocaleString());
      setArcEnd(endTime.toLocaleString());
      updateStatus("Scheduled auto reply enabled", "success");
    } catch (err: unknown) {
      updateStatus(`Failed: ${(err as Error).message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  // ── Set Vacation ──────────────────────────────────────────
  const handleSetVacation = async () => {
    if (!returnDate) {
      updateStatus("Select a return date first.", "error");
      return;
    }
    if (!config.startOfShift || !config.endOfShift) {
      updateStatus("Configure shift hours on the Config tab first.", "error");
      return;
    }
    setBusy(true);
    updateStatus("Setting vacation OOF...");
    try {
      const { startTime, endTime } = calculateVacationWindow(
        new Date(returnDate), config.startOfShift, config.endOfShift
      );
      await setScheduledAutoReply(startTime, endTime);
      setArcState("scheduled");
      setArcStart(startTime.toLocaleString());
      setArcEnd(endTime.toLocaleString());
      updateStatus(`Vacation OOF set until ${returnDate}`, "success");
    } catch (err: unknown) {
      updateStatus(`Failed: ${(err as Error).message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  // ── Refresh Status ────────────────────────────────────────
  const handleRefresh = async () => {
    setBusy(true);
    updateStatus("Refreshing status...");
    try {
      const settings = await getMailboxSettings();
      const ars = settings.automaticRepliesSetting;
      setArcState(ars.status);
      setArcStart(ars.scheduledStartDateTime?.dateTime
        ? new Date(ars.scheduledStartDateTime.dateTime).toLocaleString()
        : "-");
      setArcEnd(ars.scheduledEndDateTime?.dateTime
        ? new Date(ars.scheduledEndDateTime.dateTime).toLocaleString()
        : "-");
      updateStatus("Status refreshed", "success");
    } catch (err: unknown) {
      updateStatus(`Refresh failed: ${(err as Error).message}`, "error");
    } finally {
      setBusy(false);
    }
  };

  const stateClassName = arcState === "scheduled" ? "scheduled"
    : arcState === "alwaysEnabled" ? "enabled"
    : arcState === "disabled" ? "disabled" : "";

  return (
    <div>
      {/* Connection */}
      <div className="section">
        <div className="section-title">Connection</div>
        <div className={`connection-badge ${connected ? "connected" : "disconnected"}`}>
          {connected ? "Connected" : "Not Connected"}
          {userEmail && ` — ${userEmail}`}
        </div>
        <div className="button-row">
          {!connected ? (
            <Button appearance="primary" icon={<PlugConnectedRegular />} onClick={handleConnect} disabled={busy}>
              {busy ? <Spinner size="tiny" /> : "Connect"}
            </Button>
          ) : (
            <Button appearance="secondary" icon={<PlugDisconnectedRegular />} onClick={handleDisconnect} disabled={busy}>
              Disconnect
            </Button>
          )}
        </div>
      </div>

      {/* Quick Actions */}
      <div className="section">
        <div className="section-title">Quick Actions</div>
        <div className="button-row">
          <Button
            appearance="primary"
            icon={<CalendarClockRegular />}
            onClick={handleEnableScheduled}
            disabled={busy || !connected}
          >
            Enable Scheduled
          </Button>
          <Button
            appearance="secondary"
            icon={<ArrowSyncRegular />}
            onClick={handleRefresh}
            disabled={busy || !connected}
          >
            Refresh Status
          </Button>
        </div>
      </div>

      {/* Vacation */}
      <div className="section">
        <div className="section-title">Vacation / Extended OOF</div>
        <div className="field-row">
          <label>Return Date</label>
          <input
            type="date"
            value={returnDate}
            onChange={(e) => setReturnDate(e.target.value)}
            style={{ padding: "4px 8px", borderRadius: "4px", border: "1px solid #ccc" }}
          />
        </div>
        <Button
          appearance="primary"
          icon={<BeachRegular />}
          onClick={handleSetVacation}
          disabled={busy || !connected}
        >
          Set Vacation OOF
        </Button>
      </div>

      {/* Current Status */}
      <div className="section">
        <div className="section-title">Current Status</div>
        <div className="status-grid">
          <span className="status-label">State:</span>
          <span className={`status-value ${stateClassName}`}>{arcState}</span>
          <span className="status-label">Start:</span>
          <span className="status-value">{arcStart}</span>
          <span className="status-label">End:</span>
          <span className="status-value">{arcEnd}</span>
        </div>
      </div>
    </div>
  );
};
