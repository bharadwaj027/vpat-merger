import { useMemo, useState } from "react";
import type { Issue, MergeReport } from "./types";
import { PLATFORM_LABELS } from "./types";

const SEVERITY_LABEL: Record<Issue["severity"], string> = {
  error: "Errors",
  warning: "Warnings",
  info: "Notes",
};

export function Summary({ report }: { report: MergeReport }) {
  const [severity, setSeverity] = useState<"all" | Issue["severity"]>("all");
  const issues = useMemo(
    () => (report.issues ?? []).filter((i) => severity === "all" || i.severity === severity),
    [report.issues, severity],
  );
  const counts = report.issueCounts ?? { error: 0, warning: 0, info: 0 };
  const combined = report.combined!;
  const failedChecks = (report.checks ?? []).filter((c) => !c.passed);
  const platforms = (report.platforms ?? []).map((key) => ({
    key,
    label: report.platformLabels?.[key] ?? PLATFORM_LABELS[key],
  }));
  const scopeDiffs = report.scopeDifferences ?? [];
  const allLabel = platforms.length === 3 ? "all three" : `both (${platforms.map((p) => p.label).join(" + ")})`;

  return (
    <div className="summary">
      <p className="callout">
        This summary reports <strong>data quality and merging issues</strong> found while combining the source VPATs.
        It is not an independent accessibility assessment of the product.
      </p>

      <h3>Criteria</h3>
      <dl className="stats">
        {report.sources!.map((s) => (
          <div key={s.platform} className="stat">
            <dt>{s.platformLabel} source</dt>
            <dd>{s.criteriaCount}</dd>
            <dd className="stat__sub" title={s.filename}>
              {s.tables.map((t) => `${t.heading.replace(/^Table \d+: /, "")}: ${t.count}`).join(" · ")}
            </dd>
          </div>
        ))}
        <div className="stat stat--accent">
          <dt>Unique criteria in combined VPAT</dt>
          <dd>{combined.uniqueCriteria}</dd>
        </div>
        <div className="stat">
          <dt>Present in {allLabel} sources</dt>
          <dd>{combined.inAllSources}</dd>
        </div>
        <div className={`stat${combined.missingFromSome.length ? " stat--warn" : ""}`}>
          <dt>Missing from one or more sources</dt>
          <dd>{combined.missingFromSome.length}</dd>
        </div>
      </dl>

      {combined.missingFromSome.length > 0 && (
        <table className="data">
          <caption>Criteria missing from a source (shown as “Not provided in source VPAT”)</caption>
          <thead>
            <tr>
              <th scope="col">Criterion</th>
              <th scope="col">Missing from</th>
            </tr>
          </thead>
          <tbody>
            {combined.missingFromSome.map((m) => (
              <tr key={m.key}>
                <th scope="row">{m.title}</th>
                <td>{m.missing.join(", ")}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}

      <h3>Scope of evaluation</h3>
      {scopeDiffs.length === 0 ? (
        <p className="badge badge--ok">
          <span aria-hidden="true">✓ </span>The same pages were evaluated on every platform.
        </p>
      ) : (
        <>
          <p className="badge badge--warn" role="status">
            <strong>{scopeDiffs.length}</strong> {scopeDiffs.length === 1 ? "page is" : "pages are"} not evaluated on
            every platform. Pages are matched by title, so the same page with a different URL is not flagged. In the
            Word document these rows are highlighted in yellow in the Scope of Evaluation table, with a comment naming
            the platforms they are missing from.
          </p>
          <table className="data">
            <caption>Pages missing from one or more platforms</caption>
            <thead>
              <tr>
                <th scope="col">Page</th>
                <th scope="col">Evaluated on</th>
                <th scope="col">Missing from</th>
              </tr>
            </thead>
            <tbody>
              {scopeDiffs.map((d) => (
                <tr key={d.page}>
                  <th scope="row">{d.page}</th>
                  <td>{d.presentIn.join(", ")}</td>
                  <td className="lvl--missing">{d.missingFrom.join(", ")}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </>
      )}

      <h3>Output document checks</h3>
      <p className={report.outputValid ? "badge badge--ok" : "badge badge--error"}>
        {report.outputValid
          ? `All ${report.checks!.length} checks passed: the generated document was reopened and compared with every source.`
          : `${failedChecks.length} of ${report.checks!.length} checks failed. Review before using the document.`}
      </p>
      <ul className="checks">
        {report.checks!.map((c) => (
          <li key={c.name} className={c.passed ? "check check--pass" : "check check--fail"}>
            <span className="check__icon" aria-hidden="true">
              {c.passed ? "✓" : "✕"}
            </span>
            <span className="visually-hidden">{c.passed ? "Passed: " : "Failed: "}</span>
            {c.name}
            {c.details.length > 0 && (
              <ul>
                {c.details.map((d, i) => (
                  <li key={i}>{d}</li>
                ))}
              </ul>
            )}
          </li>
        ))}
      </ul>

      <h3>Issues for review</h3>
      <fieldset className="filters">
        <legend className="visually-hidden">Filter issues by severity</legend>
        {(["all", "error", "warning", "info"] as const).map((s) => (
          <label key={s} className={`chip${severity === s ? " chip--on" : ""}`}>
            <input
              type="radio"
              name="severity"
              value={s}
              checked={severity === s}
              onChange={() => setSeverity(s)}
            />
            {s === "all" ? `All (${counts.error + counts.warning + counts.info})` : `${SEVERITY_LABEL[s]} (${counts[s]})`}
          </label>
        ))}
      </fieldset>
      {issues.length === 0 ? (
        <p className="muted">No issues in this category.</p>
      ) : (
        <ul className="issues">
          {issues.map((i, idx) => (
            <li key={idx} className={`issue issue--${i.severity}`}>
              <span className="issue__sev">{i.severity === "info" ? "Note" : i.severity}</span>
              {i.platformLabel && <span className="issue__tag">{i.platformLabel}</span>}
              <span>{i.message}</span>
            </li>
          ))}
        </ul>
      )}

      <details className="matrix">
        <summary>Conformance levels by platform ({report.matrix!.length} criteria)</summary>
        <div className="table-scroll" tabIndex={0} role="region" aria-label="Conformance levels by platform">
          <table className="data">
            <thead>
              <tr>
                <th scope="col">Criterion</th>
                {platforms.map((p) => (
                  <th scope="col" key={p.key}>
                    {p.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {report.matrix!.map((row) => (
                <tr key={row.key}>
                  <th scope="row">{row.title}</th>
                  {platforms.map((p) => (
                    <td key={p.key} className={levelClass(row.results[p.key])}>
                      {row.results[p.key] ?? "Not provided in source VPAT"}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </details>
    </div>
  );
}

function levelClass(v: string | null | undefined): string {
  switch (v) {
    case "Supports":
      return "lvl lvl--supports";
    case "Partially Supports":
      return "lvl lvl--partial";
    case "Does Not Support":
      return "lvl lvl--no";
    case null:
    case undefined:
      return "lvl lvl--missing";
    default:
      return "lvl";
  }
}
