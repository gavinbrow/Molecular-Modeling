"""
app.py - Flask web application for the ORCA workflow manager.

Provides a browser UI to:
  - Enter SMILES codes for molecules
  - Configure ORCA calculation settings
  - Run the full pipeline (geometry -> input -> ORCA -> report)
  - Download the Excel report with molecule images

Start with:  python app.py
Then open:   http://localhost:5000
"""

import os
import re
import time
import threading
from pathlib import Path
from uuid import uuid4

from flask import (
    Flask,
    render_template,
    request,
    jsonify,
    send_file,
)

from pipeline import (
    ORCA_EXE,
    INP_DIR,
    OUT_DIR,
    MID_DIR,
    validate_smiles,
    smiles_to_xyz,
    generate_inp,
    make_run_stamp,
    run_orca_job,
    parse_out_file,
    build_report,
    fmt_hhmmss,
)

app = Flask(__name__)

# ── Job tracking ──────────────────────────────────────────
jobs: dict = {}
_active_lock = threading.Lock()
_active_job_id: str | None = None


def _sanitize_name(name: str) -> str:
    """Turn a user-provided name into a safe filesystem token."""
    safe = re.sub(r"[^\w\-]", "_", name.strip())
    return (safe or "mol")[:50]


# ── Routes ────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/orca_check")
def api_orca_check():
    """Let the UI know whether ORCA is reachable."""
    return jsonify({"available": ORCA_EXE.exists(), "path": str(ORCA_EXE)})


@app.route("/api/validate", methods=["POST"])
def api_validate():
    """Validate one or more SMILES strings (no job created)."""
    data = request.get_json(force=True)
    results = []
    for mol in data.get("molecules", []):
        smiles = mol.get("smiles", "").strip()
        results.append(
            {
                "name": mol.get("name", ""),
                "smiles": smiles,
                "valid": validate_smiles(smiles) if smiles else False,
            }
        )
    return jsonify({"results": results})


@app.route("/api/submit", methods=["POST"])
def api_submit():
    """Accept a job submission and kick off the pipeline in a background thread."""
    global _active_job_id

    with _active_lock:
        if _active_job_id and jobs.get(_active_job_id, {}).get("status") == "running":
            return jsonify({"error": "A job is already running. Please wait for it to finish."}), 409

    data = request.get_json(force=True)
    molecules = data.get("molecules", [])
    settings = data.get("settings", {})

    if not molecules:
        return jsonify({"error": "No molecules provided."}), 400

    # Pre-validate every SMILES
    for mol in molecules:
        smi = mol.get("smiles", "").strip()
        if not smi:
            return jsonify({"error": f"Missing SMILES for '{mol.get('name', '?')}'."}), 400
        if not validate_smiles(smi):
            return jsonify({"error": f"Invalid SMILES: {smi}"}), 400

    job_id = uuid4().hex[:8]

    with _active_lock:
        _active_job_id = job_id

    jobs[job_id] = {
        "status": "running",
        "phase": "starting",
        "current": 0,
        "total": len(molecules),
        "current_name": "",
        "elapsed": "00:00:00",
        "molecules": [
            {
                "name": _sanitize_name(m.get("name") or f"mol_{i + 1}"),
                "smiles": m["smiles"].strip(),
                "status": "pending",
                "gibbs": None,
                "error": None,
            }
            for i, m in enumerate(molecules)
        ],
        "report_path": None,
        "stamp": None,
        "error": None,
    }

    t = threading.Thread(target=_run_pipeline, args=(job_id, molecules, settings), daemon=True)
    t.start()

    return jsonify({"job_id": job_id})


@app.route("/api/status/<job_id>")
def api_status(job_id):
    if job_id not in jobs:
        return jsonify({"error": "Job not found."}), 404
    return jsonify(jobs[job_id])


@app.route("/api/download/<job_id>")
def api_download(job_id):
    if job_id not in jobs:
        return jsonify({"error": "Job not found."}), 404
    rp = jobs[job_id].get("report_path")
    if not rp or not Path(rp).exists():
        return jsonify({"error": "Report not available yet."}), 404
    return send_file(rp, as_attachment=True, download_name="orca_report.xlsx")


@app.route("/api/image/<job_id>/<name>")
def api_image(job_id, name):
    """Serve a molecule's 2D structure PNG."""
    if job_id not in jobs:
        return "", 404
    stamp = jobs[job_id].get("stamp")
    if not stamp:
        return "", 404
    safe = _sanitize_name(name)
    img_path = OUT_DIR / stamp / safe / f"{safe}.png"
    if not img_path.exists():
        return "", 404
    return send_file(str(img_path), mimetype="image/png")


# ── Background pipeline ──────────────────────────────────

def _run_pipeline(job_id: str, molecules: list[dict], settings: dict):
    global _active_job_id
    job = jobs[job_id]
    start_time = time.time()

    try:
        stamp = make_run_stamp()
        job["stamp"] = stamp
        out_run_dir = OUT_DIR / stamp
        mid_run_dir = MID_DIR / stamp
        out_run_dir.mkdir(parents=True, exist_ok=True)
        mid_run_dir.mkdir(parents=True, exist_ok=True)
        INP_DIR.mkdir(parents=True, exist_ok=True)

        image_map: dict[str, str] = {}
        smiles_map: dict[str, str] = {}
        inp_paths: dict[str, Path] = {}

        # --- Phase 1: generate geometries + input files ---
        job["phase"] = "geometry"
        for i, mol_data in enumerate(molecules):
            name = _sanitize_name(mol_data.get("name") or f"mol_{i + 1}")
            smiles = mol_data["smiles"].strip()
            job["current"] = i + 1
            job["current_name"] = name
            job["elapsed"] = fmt_hhmmss(time.time() - start_time)
            job["molecules"][i]["status"] = "generating"

            try:
                xyz_block, png_bytes, _ = smiles_to_xyz(smiles)
                inp_content = generate_inp(name, xyz_block, settings)

                inp_path = INP_DIR / f"{name}.inp"
                inp_path.write_text(inp_content, encoding="utf-8")
                inp_paths[name] = inp_path

                img_dir = out_run_dir / name
                img_dir.mkdir(parents=True, exist_ok=True)
                img_path = img_dir / f"{name}.png"
                img_path.write_bytes(png_bytes)
                image_map[name] = str(img_path)
                smiles_map[name] = smiles

                job["molecules"][i]["status"] = "generated"
            except Exception as exc:
                job["molecules"][i]["status"] = "error"
                job["molecules"][i]["error"] = str(exc)

        # --- Phase 2: run ORCA jobs ---
        job["phase"] = "orca"
        env = os.environ.copy()

        for i, mol_data in enumerate(molecules):
            name = _sanitize_name(mol_data.get("name") or f"mol_{i + 1}")

            if job["molecules"][i]["status"] == "error":
                continue  # skip molecules that failed geometry generation
            if name not in inp_paths:
                continue

            job["current"] = i + 1
            job["current_name"] = name
            job["elapsed"] = fmt_hhmmss(time.time() - start_time)
            job["molecules"][i]["status"] = "running"

            rc = run_orca_job(
                name,
                inp_paths[name],
                mid_run_dir,
                out_run_dir,
                env,
                status_dict=job,
                pipeline_start=start_time,
            )

            job["elapsed"] = fmt_hhmmss(time.time() - start_time)

            if rc == 0:
                out_file = out_run_dir / name / f"{name}.out"
                if out_file.exists():
                    parsed = parse_out_file(out_file)
                    job["molecules"][i]["gibbs"] = parsed.get("gibbs_eh")
                    job["molecules"][i]["status"] = (
                        "completed" if parsed["normal_term"] else "warning"
                    )
                else:
                    job["molecules"][i]["status"] = "completed"
            else:
                job["molecules"][i]["status"] = "error"
                job["molecules"][i]["error"] = f"ORCA exited with code {rc}"

        # --- Phase 3: generate Excel report ---
        job["phase"] = "report"
        job["elapsed"] = fmt_hhmmss(time.time() - start_time)

        out_files = []
        for i, mol_data in enumerate(molecules):
            name = _sanitize_name(mol_data.get("name") or f"mol_{i + 1}")
            out_file = out_run_dir / name / f"{name}.out"
            if out_file.exists():
                out_files.append(out_file)

        records = [parse_out_file(f) for f in out_files]
        xlsx_path = out_run_dir / "orca_report.xlsx"

        if records:
            build_report(records, image_map, smiles_map, xlsx_path)
            job["report_path"] = str(xlsx_path)

        job["status"] = "completed"
        job["phase"] = "done"
        job["elapsed"] = fmt_hhmmss(time.time() - start_time)

    except Exception as exc:
        job["status"] = "failed"
        job["error"] = str(exc)
        job["elapsed"] = fmt_hhmmss(time.time() - start_time)
    finally:
        with _active_lock:
            if _active_job_id == job_id:
                _active_job_id = None


# ── Entry point ───────────────────────────────────────────

if __name__ == "__main__":
    print(f"ORCA path: {ORCA_EXE}  ({'FOUND' if ORCA_EXE.exists() else 'NOT FOUND'})")
    print(f"INP dir:   {INP_DIR}")
    print(f"OUT dir:   {OUT_DIR}")
    print(f"MID dir:   {MID_DIR}")
    print()
    app.run(debug=True, port=5000)
