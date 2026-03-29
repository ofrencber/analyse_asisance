#!/usr/bin/env python3
"""
MCDM Karar Destek Sistemi — Command-Line Interface
======================================
Run MCDM analyses without the Streamlit UI.

Usage:
    python mcdm_cli.py --config analysis.json --output results/
    python mcdm_cli.py --config batch.json --batch
    python mcdm_cli.py --list-methods

Config file format (single analysis):
{
    "data": "data.csv",
    "criteria_types": {"K1": "min", "K2": "max", "K3": "max"},
    "weight_method": "Entropy",
    "ranking_method": "TOPSIS",
    "compare_methods": ["VIKOR", "EDAS"],
    "fuzzy_spread": 0.10,
    "sensitivity_iterations": 200,
    "outputs": ["excel", "imrad_tr", "imrad_en", "executive_summary", "dashboard"]
}

Batch config format:
{
    "data": "data.csv",
    "criteria_types": {"K1": "min", "K2": "max"},
    "batch": [
        {"weight_method": "Entropy", "ranking_method": "TOPSIS"},
        {"weight_method": "CRITIC", "ranking_method": "VIKOR"},
        {"weight_method": "MEREC", "ranking_method": "Fuzzy TOPSIS"}
    ],
    "outputs": ["json"]
}
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path
from typing import Any, Dict, List

import numpy as np
import pandas as pd


def _load_data(path: str) -> pd.DataFrame:
    """Load CSV or Excel file."""
    p = Path(path)
    if p.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(p)
    elif p.suffix.lower() == ".csv":
        return pd.read_csv(p)
    else:
        raise ValueError(f"Unsupported file format: {p.suffix}")


def _run_single(
    data: pd.DataFrame,
    cfg: Dict[str, Any],
    output_dir: Path,
    label: str = "",
) -> Dict[str, Any]:
    """Run a single analysis and save outputs."""
    import mcdm_engine as me

    criteria = list(cfg.get("criteria") or [c for c in data.select_dtypes(include=["number"]).columns])
    criteria_types = cfg.get("criteria_types", {})
    # Auto-fill missing criteria types as benefit
    for c in criteria:
        if c not in criteria_types:
            criteria_types[c] = "max"

    # Apply thresholds if present
    thresholds = cfg.get("thresholds", {})
    eliminated = []
    if thresholds:
        data, eliminated = me.apply_threshold_filter(data, criteria, criteria_types, thresholds)
        if len(data) < 2:
            return {"error": "Fewer than 2 alternatives after threshold filtering", "eliminated": eliminated}

    config = me.AnalysisConfig(
        criteria=criteria,
        criteria_types=criteria_types,
        weight_method=cfg.get("weight_method", "Entropy"),
        weight_mode=cfg.get("weight_mode", "objective"),
        ranking_method=cfg.get("ranking_method"),
        compare_methods=cfg.get("compare_methods", []),
        fuzzy_spread=float(cfg.get("fuzzy_spread", 0.10)),
        sensitivity_iterations=int(cfg.get("sensitivity_iterations", 200)),
        sensitivity_sigma=float(cfg.get("sensitivity_sigma", 0.12)),
        run_heavy_robustness=bool(cfg.get("run_heavy_robustness", True)),
        manual_weights=cfg.get("manual_weights"),
    )

    t0 = time.time()
    result = me.run_full_analysis(data, config)
    elapsed = time.time() - t0
    result["selected_data"] = data
    result["eliminated_alternatives"] = eliminated

    prefix = f"{label}_" if label else ""
    outputs = cfg.get("outputs", ["json"])

    # Always produce JSON summary
    summary = _build_json_summary(result, criteria, elapsed, eliminated)
    json_path = output_dir / f"{prefix}result.json"
    json_path.write_text(json.dumps(summary, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
    print(f"  {'✓' if True else '✗'} JSON: {json_path}")

    # Optional outputs
    import mcdm_article

    if "excel" in outputs:
        # Excel requires streamlit context for full generation; export JSON tables instead
        _export_tables_excel(result, criteria, output_dir, prefix)

    if "imrad_tr" in outputs:
        b = mcdm_article.generate_imrad_docx(result, data, lang="TR")
        if b:
            p = output_dir / f"{prefix}IMRAD_TR.docx"
            p.write_bytes(b)
            print(f"  ✓ IMRAD TR: {p} ({len(b):,} bytes)")

    if "imrad_en" in outputs:
        b = mcdm_article.generate_imrad_docx(result, data, lang="EN")
        if b:
            p = output_dir / f"{prefix}IMRAD_EN.docx"
            p.write_bytes(b)
            print(f"  ✓ IMRAD EN: {p} ({len(b):,} bytes)")

    if "executive_summary" in outputs or "exec" in outputs:
        b = mcdm_article.generate_executive_summary_docx(result, data, lang="TR")
        if b:
            p = output_dir / f"{prefix}Executive_Summary.docx"
            p.write_bytes(b)
            print(f"  ✓ Executive Summary: {p} ({len(b):,} bytes)")

    if "dashboard" in outputs:
        html = mcdm_article.generate_dashboard_html(result, data, lang="TR")
        if html:
            p = output_dir / f"{prefix}Dashboard.html"
            p.write_text(html, encoding="utf-8")
            print(f"  ✓ Dashboard: {p} ({len(html):,} chars)")

    return summary


def _build_json_summary(result: dict, criteria: list, elapsed: float, eliminated: list) -> dict:
    """Build a JSON-serializable summary of analysis results."""
    weights = result.get("weights", {})
    ranking = result.get("ranking", {})
    sensitivity = result.get("sensitivity", {})
    confidence = result.get("decision_confidence", {})

    wt = weights.get("table")
    rt = ranking.get("table")

    summary: Dict[str, Any] = {
        "weight_method": weights.get("method"),
        "ranking_method": ranking.get("method"),
        "criteria": criteria,
        "n_alternatives": len(result.get("selected_data", pd.DataFrame())),
        "n_criteria": len(criteria),
        "analysis_time_seconds": round(elapsed, 3),
    }

    if isinstance(wt, pd.DataFrame) and not wt.empty:
        w_col = "Ağırlık" if "Ağırlık" in wt.columns else ("Weight" if "Weight" in wt.columns else None)
        c_col = "Kriter" if "Kriter" in wt.columns else ("Criterion" if "Criterion" in wt.columns else None)
        if w_col and c_col:
            summary["weights"] = dict(zip(wt[c_col], wt[w_col].round(6)))

    if isinstance(rt, pd.DataFrame) and not rt.empty:
        a_col = next((c for c in rt.columns if c in ("Alternatif", "Alternative")), None)
        s_col = next((c for c in rt.columns if c in ("Skor", "Score")), None)
        r_col = next((c for c in rt.columns if c in ("Sıra", "Rank")), None)
        if a_col and s_col and r_col:
            summary["ranking"] = [
                {"alternative": str(row[a_col]), "score": round(float(row[s_col]), 6), "rank": int(row[r_col])}
                for _, row in rt.sort_values(r_col).iterrows()
            ]
            summary["leader"] = str(rt.sort_values(r_col).iloc[0][a_col])

    if sensitivity:
        mc = sensitivity.get("monte_carlo_summary")
        summary["stability"] = float(sensitivity.get("top_stability", 0))
        if isinstance(mc, pd.DataFrame) and not mc.empty:
            summary["mc_summary"] = mc.to_dict(orient="records")

    if confidence:
        summary["confidence_level"] = confidence.get("verdict")

    if eliminated:
        summary["eliminated_alternatives"] = eliminated

    return summary


def _export_tables_excel(result: dict, criteria: list, output_dir: Path, prefix: str) -> None:
    """Export result tables as a multi-sheet Excel file without Streamlit dependency."""
    try:
        path = output_dir / f"{prefix}Tables.xlsx"
        _engine = "xlsxwriter"
        try:
            import xlsxwriter  # noqa: F401
        except ImportError:
            _engine = "openpyxl"
        with pd.ExcelWriter(path, engine=_engine) as writer:
            # Weight table
            wt = (result.get("weights") or {}).get("table")
            if isinstance(wt, pd.DataFrame) and not wt.empty:
                wt.to_excel(writer, sheet_name="Weights", index=False)

            # Ranking table
            rt = (result.get("ranking") or {}).get("table")
            if isinstance(rt, pd.DataFrame) and not rt.empty:
                rt.to_excel(writer, sheet_name="Ranking", index=False)

            # Detail tables
            details = (result.get("ranking") or {}).get("details", {})
            for key, val in details.items():
                if key in ("result_table", "score_direction"):
                    continue
                if isinstance(val, pd.DataFrame) and not val.empty:
                    sheet = key[:31]  # Excel max sheet name length
                    val.to_excel(writer, sheet_name=sheet, index=False)
                elif isinstance(val, np.ndarray) and val.size > 0:
                    df = pd.DataFrame(val)
                    sheet = key[:31]
                    df.to_excel(writer, sheet_name=sheet, index=False)

            # Comparison
            comp = result.get("comparison", {})
            for ckey in ["rank_table", "score_table", "spearman_matrix"]:
                cv = comp.get(ckey)
                if isinstance(cv, pd.DataFrame) and not cv.empty:
                    cv.to_excel(writer, sheet_name=ckey[:31], index=("matrix" in ckey))

            # Sensitivity
            sens = result.get("sensitivity", {})
            for skey in ["monte_carlo_summary", "local_sensitivity"]:
                sv = sens.get(skey)
                if isinstance(sv, pd.DataFrame) and not sv.empty:
                    sv.to_excel(writer, sheet_name=skey[:31], index=False)

        print(f"  ✓ Excel Tables: {path}")
    except Exception as e:
        print(f"  ✗ Excel Tables: {e}")


def _list_methods():
    """Print all available methods."""
    import mcdm_engine as me
    print("\n=== OBJECTIVE WEIGHTING METHODS ===")
    for m in me.OBJECTIVE_WEIGHT_METHODS:
        print(f"  {m}")
    print("\n=== CLASSICAL RANKING METHODS ===")
    for m in me.CLASSICAL_MCDM_METHODS:
        print(f"  {m}")
    print("\n=== FUZZY RANKING METHODS ===")
    for m in me.FUZZY_MCDM_METHODS:
        print(f"  {m}")
    print(f"\nTotal: {len(me.OBJECTIVE_WEIGHT_METHODS)} weighting + "
          f"{len(me.CLASSICAL_MCDM_METHODS)} classical + "
          f"{len(me.FUZZY_MCDM_METHODS)} fuzzy = "
          f"{len(me.OBJECTIVE_WEIGHT_METHODS) + len(me.CLASSICAL_MCDM_METHODS) + len(me.FUZZY_MCDM_METHODS)} methods")


def main():
    parser = argparse.ArgumentParser(
        description="MCDM Karar Destek Sistemi CLI — Run multi-criteria decision analyses from the command line.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--config", "-c", type=str, help="Path to JSON config file")
    parser.add_argument("--output", "-o", type=str, default="./mcdm_output", help="Output directory (default: ./mcdm_output)")
    parser.add_argument("--batch", "-b", action="store_true", help="Run in batch mode (config must have 'batch' array)")
    parser.add_argument("--list-methods", action="store_true", help="List all available methods and exit")
    parser.add_argument("--seed", type=int, default=42, help="Random seed for reproducibility (default: 42)")

    args = parser.parse_args()

    if args.list_methods:
        _list_methods()
        return

    if not args.config:
        parser.error("--config is required (or use --list-methods)")

    config_path = Path(args.config)
    if not config_path.is_file():
        print(f"Error: Config file not found: {config_path}", file=sys.stderr)
        sys.exit(1)

    cfg = json.loads(config_path.read_text(encoding="utf-8"))
    data_path = cfg.get("data")
    if not data_path:
        print("Error: 'data' field required in config", file=sys.stderr)
        sys.exit(1)

    # Resolve data path relative to config file
    data_file = Path(data_path)
    if not data_file.is_absolute():
        data_file = config_path.parent / data_file

    if not data_file.is_file():
        print(f"Error: Data file not found: {data_file}", file=sys.stderr)
        sys.exit(1)

    data = _load_data(str(data_file))
    print(f"Data loaded: {data.shape[0]} rows × {data.shape[1]} columns from {data_file.name}")

    # Set first non-numeric column as index if present
    if data.iloc[:, 0].dtype == object:
        data = data.set_index(data.columns[0])

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    np.random.seed(args.seed)

    if args.batch:
        batch_configs = cfg.get("batch", [])
        if not batch_configs:
            print("Error: 'batch' array required in config for batch mode", file=sys.stderr)
            sys.exit(1)

        print(f"\n{'='*60}")
        print(f"BATCH MODE: {len(batch_configs)} analyses")
        print(f"{'='*60}")

        all_summaries = []
        for i, bc in enumerate(batch_configs):
            # Merge batch item with base config
            merged = {**cfg, **bc}
            merged.pop("batch", None)
            wm = merged.get("weight_method", "?")
            rm = merged.get("ranking_method", "?")
            label = f"{i+1:02d}_{wm}_{rm}".replace(" ", "_")
            print(f"\n[{i+1}/{len(batch_configs)}] {wm} + {rm}")
            summary = _run_single(data.copy(), merged, output_dir, label=label)
            summary["_label"] = label
            all_summaries.append(summary)

        # Batch comparison summary
        comp_path = output_dir / "batch_comparison.json"
        comp_path.write_text(json.dumps(all_summaries, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
        print(f"\n✓ Batch comparison: {comp_path}")

        # Quick comparison table
        print(f"\n{'Label':30s} | {'Leader':8s} | {'Score':8s} | {'Stability':10s}")
        print("-" * 65)
        for s in all_summaries:
            leader = s.get("leader", "?")
            score = s.get("ranking", [{}])[0].get("score", 0) if s.get("ranking") else 0
            stab = s.get("stability", 0)
            print(f"  {s.get('_label', '?'):28s} | {leader:8s} | {score:8.4f} | {stab*100:8.1f}%")

    else:
        print(f"\n{'='*60}")
        print(f"SINGLE ANALYSIS: {cfg.get('weight_method', '?')} + {cfg.get('ranking_method', '?')}")
        print(f"{'='*60}\n")
        _run_single(data, cfg, output_dir)

    print(f"\n✓ All outputs saved to: {output_dir}/")


if __name__ == "__main__":
    main()
