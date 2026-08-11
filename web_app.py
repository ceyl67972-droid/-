from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import uuid
from pathlib import Path

from flask import Flask, jsonify, request, send_file


BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "web"
JOBS_DIR = BASE_DIR / "web_jobs"
WORKER = BASE_DIR / "web_worker.py"
RESULTS_DIR = Path.home() / "Desktop" / "流水核对结果"

app = Flask(__name__, static_folder=str(STATIC_DIR), static_url_path="/assets")
app.config["MAX_CONTENT_LENGTH"] = 4 * 1024 * 1024 * 1024

jobs: dict[str, dict[str, object]] = {}
jobs_lock = threading.Lock()
processing_lock = threading.Lock()


def safe_filename(filename: str, fallback: str) -> str:
    name = Path(filename or fallback).name
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")
    return name or fallback


def public_job(job: dict[str, object]) -> dict[str, object]:
    return {
        "id": job["id"],
        "status": job["status"],
        "created_at": job["created_at"],
        "started_at": job.get("started_at"),
        "finished_at": job.get("finished_at"),
        "logs": job.get("logs", [])[-120:],
        "stats": job.get("stats"),
        "error": job.get("error"),
        "download_ready": bool(job.get("result_path")),
        "result_name": job.get("result_name"),
    }


def run_job(job_id: str, manifest_path: Path) -> None:
    with processing_lock:
        with jobs_lock:
            job = jobs[job_id]
            job["status"] = "running"
            job["started_at"] = time.time()

        command = [sys.executable, "-u", str(WORKER), str(manifest_path)]
        try:
            process = subprocess.Popen(
                command,
                cwd=str(BASE_DIR),
                env={**os.environ, "PYTHONIOENCODING": "utf-8"},
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
            )
            result_payload = None
            assert process.stdout is not None
            for raw_line in process.stdout:
                line = raw_line.rstrip()
                if not line:
                    continue
                if line.startswith("__RESULT__"):
                    result_payload = json.loads(line.removeprefix("__RESULT__"))
                    continue
                with jobs_lock:
                    jobs[job_id].setdefault("logs", []).append(line)
            return_code = process.wait()
            if return_code != 0 or result_payload is None:
                raise RuntimeError("核对任务未正常完成，请查看运行日志")

            with jobs_lock:
                job = jobs[job_id]
                job["status"] = "completed"
                job["finished_at"] = time.time()
                job["stats"] = result_payload["stats"]
                job["result_path"] = result_payload["result_path"]
                job["result_name"] = Path(result_payload["result_path"]).name
        except Exception as exc:
            with jobs_lock:
                job = jobs[job_id]
                job["status"] = "failed"
                job["finished_at"] = time.time()
                job["error"] = str(exc)


@app.get("/")
def index():
    return send_file(STATIC_DIR / "index.html")


@app.get("/api/health")
def health():
    return jsonify({"status": "ok", "local": True})


@app.post("/api/jobs")
def create_job():
    excel_file = request.files.get("excel")
    pdf_files = request.files.getlist("pdfs")
    if excel_file is None or not excel_file.filename:
        return jsonify({"error": "请选择一份 Excel 序时账"}), 400
    if not pdf_files or not any(item.filename for item in pdf_files):
        return jsonify({"error": "请至少选择一份 PDF 对账单"}), 400
    if Path(excel_file.filename).suffix.lower() != ".xlsx":
        return jsonify({"error": "序时账必须是 .xlsx 文件"}), 400

    try:
        tolerance = int(request.form.get("date_tolerance", "62"))
    except ValueError:
        return jsonify({"error": "日期容差必须是整数"}), 400
    if not 0 <= tolerance <= 366:
        return jsonify({"error": "日期容差需在 0 到 366 天之间"}), 400

    job_id = uuid.uuid4().hex
    job_dir = JOBS_DIR / job_id
    input_dir = job_dir / "input"
    input_dir.mkdir(parents=True, exist_ok=True)

    excel_name = safe_filename(excel_file.filename, "ledger.xlsx")
    excel_path = input_dir / excel_name
    excel_file.save(excel_path)

    pdf_paths = []
    for index, uploaded in enumerate(pdf_files, 1):
        if not uploaded.filename:
            continue
        if Path(uploaded.filename).suffix.lower() != ".pdf":
            return jsonify({"error": f"{uploaded.filename} 不是 PDF 文件"}), 400
        filename = safe_filename(uploaded.filename, f"statement-{index}.pdf")
        destination = input_dir / f"{index:03d}-{filename}"
        uploaded.save(destination)
        pdf_paths.append(str(destination))

    output_path = job_dir / f"{Path(excel_name).stem}_核对结果.xlsx"
    manifest = {
        "excel_path": str(excel_path),
        "pdf_paths": pdf_paths,
        "output_path": str(output_path),
        "date_tolerance": tolerance,
    }
    manifest_path = job_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")

    job = {
        "id": job_id,
        "status": "queued",
        "created_at": time.time(),
        "logs": [f"已接收 1 份序时账和 {len(pdf_paths)} 份对账单"],
    }
    with jobs_lock:
        jobs[job_id] = job
    threading.Thread(target=run_job, args=(job_id, manifest_path), daemon=True).start()
    return jsonify(public_job(job)), 202


@app.get("/api/jobs/<job_id>")
def get_job(job_id: str):
    with jobs_lock:
        job = jobs.get(job_id)
        if job is None:
            return jsonify({"error": "任务不存在"}), 404
        return jsonify(public_job(job))


@app.get("/api/jobs/<job_id>/download")
def download_result(job_id: str):
    with jobs_lock:
        job = jobs.get(job_id)
        if job is None or not job.get("result_path"):
            return jsonify({"error": "结果文件尚未生成"}), 404
        result_path = Path(str(job["result_path"]))
        result_name = str(job["result_name"])
    return send_file(result_path, as_attachment=True, download_name=result_name)


@app.post("/api/jobs/<job_id>/save-local")
def save_result_locally(job_id: str):
    with jobs_lock:
        job = jobs.get(job_id)
        if job is None or not job.get("result_path"):
            return jsonify({"error": "结果文件尚未生成"}), 404
        source = Path(str(job["result_path"]))
        result_name = str(job["result_name"])

    if not source.exists():
        return jsonify({"error": "结果文件已不存在，请重新核对"}), 404

    RESULTS_DIR.mkdir(parents=True, exist_ok=True)
    destination = RESULTS_DIR / result_name
    try:
        shutil.copy2(source, destination)
    except PermissionError:
        return jsonify({"error": "结果文件正在被占用，请关闭 Excel 后重试"}), 409

    try:
        os.startfile(RESULTS_DIR)
    except OSError:
        pass
    return jsonify({"status": "saved", "path": str(destination)})


def cleanup_old_jobs() -> None:
    JOBS_DIR.mkdir(exist_ok=True)
    cutoff = time.time() - 7 * 24 * 60 * 60
    for directory in JOBS_DIR.iterdir():
        try:
            if directory.is_dir() and directory.stat().st_mtime < cutoff:
                shutil.rmtree(directory, ignore_errors=True)
        except OSError:
            continue


if __name__ == "__main__":
    cleanup_old_jobs()
    try:
        from waitress import serve

        serve(app, host="127.0.0.1", port=8765, threads=6)
    except ImportError:
        app.run(host="127.0.0.1", port=8765, debug=False)
