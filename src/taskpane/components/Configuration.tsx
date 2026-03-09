import React from "react";
import { Button, Checkbox, Input, Label, Select } from "@fluentui/react-components";
import { SaveRegular } from "@fluentui/react-icons";
import type { AppConfig } from "../types";

const HOURS = Array.from({ length: 12 }, (_, i) => (i + 1).toString());
const MINUTES = ["00", "15", "30", "45"];
const AMPM = ["AM", "PM"];

const ALL_DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
const PRESETS: Record<string, string[]> = {
  "Mon–Fri": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"],
  "Sun–Wed (4×10)": ["Sunday", "Monday", "Tuesday", "Wednesday"],
  "Wed–Sat (4×10)": ["Wednesday", "Thursday", "Friday", "Saturday"],
};

interface Props {
  config: AppConfig;
  updateConfig: (patch: Partial<AppConfig>) => void;
  updateStatus: (text: string, type?: "info" | "error" | "success") => void;
}

/** Parse "HH:mm" into { hour12, minute, ampm } for display. */
function parse24(time: string | null): { hour: string; min: string; ampm: string } {
  if (!time) return { hour: "9", min: "00", ampm: "AM" };
  const [h, m] = time.split(":").map(Number);
  const ampm = h >= 12 ? "PM" : "AM";
  const h12 = h > 12 ? h - 12 : h === 0 ? 12 : h;
  return { hour: h12.toString(), min: MINUTES.reduce((a, b) => Math.abs(Number(b) - m) < Math.abs(Number(a) - m) ? b : a), ampm };
}

/** Convert 12h parts to "HH:mm" (24h). */
function to24(hour: string, min: string, ampm: string): string {
  let h = parseInt(hour, 10);
  if (ampm === "PM" && h !== 12) h += 12;
  if (ampm === "AM" && h === 12) h = 0;
  return `${h.toString().padStart(2, "0")}:${min}`;
}

export const Configuration: React.FC<Props> = ({ config, updateConfig, updateStatus }) => {
  const start = parse24(config.startOfShift);
  const end = parse24(config.endOfShift);

  const saveShiftTime = (which: "start" | "end", field: "hour" | "min" | "ampm", value: string) => {
    const cur = which === "start" ? start : end;
    const updated = { ...cur, [field]: value };
    const time24 = to24(updated.hour, updated.min, updated.ampm);
    updateConfig(which === "start" ? { startOfShift: time24 } : { endOfShift: time24 });
    updateStatus("Shift hours updated", "success");
  };

  const toggleDay = (day: string) => {
    const days = config.workDays.includes(day)
      ? config.workDays.filter((d) => d !== day)
      : [...config.workDays, day];
    updateConfig({ workDays: days });
  };

  const applyPreset = (days: string[]) => {
    updateConfig({ workDays: days });
    updateStatus("Work days preset applied", "success");
  };

  return (
    <div>
      {/* Identity */}
      <div className="section">
        <div className="section-title">Identity</div>
        <div className="field-row">
          <Label weight="semibold" style={{ minWidth: 80 }}>Full Name</Label>
          <Input
            value={config.fullName}
            onChange={(_, data) => updateConfig({ fullName: data.value })}
            placeholder="Auto-detected from profile"
            style={{ flex: 1 }}
          />
        </div>
        <div className="field-row">
          <Label weight="semibold" style={{ minWidth: 80 }}>Role</Label>
          <Input
            value={config.role}
            onChange={(_, data) => updateConfig({ role: data.value })}
            style={{ flex: 1 }}
          />
        </div>
      </div>

      {/* Shift Hours */}
      <div className="section">
        <div className="section-title">Office Hours</div>
        <div className="time-row">
          <label>Start</label>
          <Select value={start.hour} onChange={(_, d) => saveShiftTime("start", "hour", d.value)}>
            {HOURS.map((h) => <option key={h} value={h}>{h}</option>)}
          </Select>
          <span>:</span>
          <Select value={start.min} onChange={(_, d) => saveShiftTime("start", "min", d.value)}>
            {MINUTES.map((m) => <option key={m} value={m}>{m}</option>)}
          </Select>
          <Select value={start.ampm} onChange={(_, d) => saveShiftTime("start", "ampm", d.value)}>
            {AMPM.map((a) => <option key={a} value={a}>{a}</option>)}
          </Select>
        </div>
        <div className="time-row">
          <label>End</label>
          <Select value={end.hour} onChange={(_, d) => saveShiftTime("end", "hour", d.value)}>
            {HOURS.map((h) => <option key={h} value={h}>{h}</option>)}
          </Select>
          <span>:</span>
          <Select value={end.min} onChange={(_, d) => saveShiftTime("end", "min", d.value)}>
            {MINUTES.map((m) => <option key={m} value={m}>{m}</option>)}
          </Select>
          <Select value={end.ampm} onChange={(_, d) => saveShiftTime("end", "ampm", d.value)}>
            {AMPM.map((a) => <option key={a} value={a}>{a}</option>)}
          </Select>
        </div>
      </div>

      {/* Work Days */}
      <div className="section">
        <div className="section-title">Work Days</div>
        <div className="day-grid">
          {ALL_DAYS.map((day) => (
            <Checkbox
              key={day}
              label={day.slice(0, 3)}
              checked={config.workDays.includes(day)}
              onChange={() => toggleDay(day)}
            />
          ))}
        </div>
        <div className="button-row">
          {Object.entries(PRESETS).map(([label, days]) => (
            <Button key={label} size="small" appearance="subtle" onClick={() => applyPreset(days)}>
              {label}
            </Button>
          ))}
        </div>
      </div>

      {/* Signature Options */}
      <div className="section">
        <div className="section-title">Signature Options</div>
        <Checkbox
          label="Include Signature"
          checked={config.includeSignature}
          onChange={(_, data) => updateConfig({ includeSignature: !!data.checked })}
        />
        <Checkbox
          label="Include Office Hours"
          checked={config.includeOfficeHours}
          onChange={(_, data) => updateConfig({ includeOfficeHours: !!data.checked })}
        />
        <Checkbox
          label="Include Work Days"
          checked={config.includeWorkDays}
          onChange={(_, data) => updateConfig({ includeWorkDays: !!data.checked })}
        />
        <Checkbox
          label="Include Timezone"
          checked={config.includeTimezone}
          onChange={(_, data) => updateConfig({ includeTimezone: !!data.checked })}
        />
      </div>
    </div>
  );
};
