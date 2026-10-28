import { useId, useRef, useState } from "react";
import type { InspectResult, Kind, SlotKey } from "./types";
import { PLATFORM_LABELS, slotPlatform } from "./types";

export interface SlotState {
  file: File | null;
  checking: boolean;
  result: InspectResult | null;
  kind: Kind;
  /** True once the user picks Mobile Web / Native themselves; stops auto-detection from overriding it. */
  kindTouched: boolean;
}

interface Props {
  slot: SlotKey;
  name: string;
  hint: string;
  nativeCapable: boolean;
  state: SlotState;
  disabled: boolean;
  onSelect: (file: File) => void;
  onRemove: () => void;
  onKindChange: (kind: Kind) => void;
}

export function UploadSlot({ slot, name, hint, nativeCapable, state, disabled, onSelect, onRemove, onKindChange }: Props) {
  const inputId = useId();
  const statusId = useId();
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);

  const { file, checking, result, kind } = state;
  const platform = slotPlatform(slot, kind);
  const label = PLATFORM_LABELS[platform];
  const status: "empty" | "checking" | "ok" | "error" = !file
    ? "empty"
    : checking
      ? "checking"
      : result?.ok
        ? "ok"
        : "error";
  const guessed = result?.guessedPlatform;
  const kindMismatch = status === "ok" && guessed && guessed !== platform && guessed.split("_")[0] === slot;

  function pick(files: FileList | null) {
    const f = files?.[0];
    if (f) onSelect(f);
    if (inputRef.current) inputRef.current.value = "";
  }

  return (
    <section
      className={`slot slot--${status}${dragging ? " slot--drag" : ""}`}
      aria-labelledby={`${inputId}-title`}
      onDragOver={(e) => {
        e.preventDefault();
        if (!disabled) setDragging(true);
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragging(false);
        if (!disabled) pick(e.dataTransfer.files);
      }}
    >
      <h3 id={`${inputId}-title`} className="slot__title">
        {name} VPAT
      </h3>
      <p className="slot__hint">{hint}</p>

      {nativeCapable && (
        <fieldset className="kind" disabled={disabled}>
          <legend className="visually-hidden">{name} VPAT type</legend>
          {(["web", "native"] as const).map((k) => (
            <label key={k} className={`chip${kind === k ? " chip--on" : ""}`}>
              <input
                type="radio"
                name={`${inputId}-kind`}
                value={k}
                checked={kind === k}
                onChange={() => onKindChange(k)}
              />
              {k === "web" ? "Mobile Web" : "Native app"}
            </label>
          ))}
        </fieldset>
      )}
      <p className="slot__label">
        Labeled in report as <strong>{label}:</strong>
      </p>

      <input
        ref={inputRef}
        id={inputId}
        type="file"
        accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        className="visually-hidden"
        tabIndex={-1}
        aria-hidden="true"
        disabled={disabled}
        data-slot={slot}
        onChange={(e) => pick(e.target.files)}
      />

      <div id={statusId} className="slot__status" aria-live="polite">
        {status === "empty" && <p className="muted">No file selected. Drop a .docx here or choose a file.</p>}
        {file && <p className="slot__filename" title={file.name}>{file.name}</p>}
        {status === "checking" && <p className="badge badge--info">Checking document…</p>}
        {status === "ok" && result && (
          <>
            <p className="badge badge--ok">
              <span aria-hidden="true">✓ </span>Valid VPAT: {result.criteriaCount} criteria found
            </p>
            {result.tables && (
              <ul className="slot__tables">
                {result.tables.map((t) => (
                  <li key={t.heading}>
                    {t.heading}: {t.count}
                  </li>
                ))}
              </ul>
            )}
          </>
        )}
        {status === "error" && result && (
          <p className="badge badge--error" role="alert">
            <span aria-hidden="true">✕ </span>
            {result.error}
          </p>
        )}
        {status === "ok" && (kindMismatch || (result?.warnings?.length ?? 0) > 0) && (
          <ul className="slot__warnings">
            {kindMismatch && guessed && (
              <li>
                <strong>Warning:</strong> This file looks like a {PLATFORM_LABELS[guessed]} VPAT, but{" "}
                {kind === "native" ? "Native app" : "Mobile Web"} is selected.
              </li>
            )}
            {result?.warnings?.map((w, i) => (
              <li key={i}>
                <strong>Warning:</strong> {w}
              </li>
            ))}
          </ul>
        )}
      </div>

      <div className="slot__actions">
        <button
          type="button"
          className="btn btn--secondary"
          disabled={disabled}
          aria-describedby={statusId}
          onClick={() => inputRef.current?.click()}
        >
          {file ? "Replace file" : "Choose file"}
          <span className="visually-hidden"> for {name}</span>
        </button>
        {file && (
          <button type="button" className="btn btn--link" onClick={onRemove} disabled={disabled}>
            Remove<span className="visually-hidden"> {name} file</span>
          </button>
        )}
      </div>
    </section>
  );
}
