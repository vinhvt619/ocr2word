from __future__ import annotations

import time
from threading import Lock
from typing import Any


class ProgressRegistry:
    """Small in-memory registry for progress polling during a conversion request."""

    def __init__(self, *, ttl_seconds: int = 3600, max_jobs: int = 250) -> None:
        self._ttl_seconds = ttl_seconds
        self._max_jobs = max_jobs
        self._jobs: dict[str, dict[str, Any]] = {}
        self._lock = Lock()

    def _prune_locked(self, now: float) -> None:
        expired = [
            job_id
            for job_id, job in self._jobs.items()
            if now - float(job.get("updated_at", now)) > self._ttl_seconds
        ]
        for job_id in expired:
            self._jobs.pop(job_id, None)
        if len(self._jobs) <= self._max_jobs:
            return
        oldest = sorted(self._jobs, key=lambda key: float(self._jobs[key].get("updated_at", 0)))
        for job_id in oldest[: len(self._jobs) - self._max_jobs]:
            self._jobs.pop(job_id, None)

    def start(self, job_id: str, *, total_files: int) -> None:
        now = time.time()
        with self._lock:
            self._prune_locked(now)
            self._jobs[job_id] = {
                "job_id": job_id,
                "status": "uploading",
                "percent": 0,
                "message": "Đang nhận tài liệu…",
                "current_file": None,
                "completed_files": 0,
                "total_files": total_files,
                "updated_at": now,
            }

    def update(self, job_id: str | None, **changes: Any) -> None:
        if not job_id:
            return
        now = time.time()
        with self._lock:
            job = self._jobs.get(job_id)
            if job is None:
                return
            if "percent" in changes:
                requested = max(0, min(100, round(float(changes["percent"]))))
                changes["percent"] = max(int(job.get("percent", 0)), requested)
            job.update(changes)
            job["updated_at"] = now

    def get(self, job_id: str) -> dict[str, Any] | None:
        now = time.time()
        with self._lock:
            self._prune_locked(now)
            job = self._jobs.get(job_id)
            if job is None:
                return None
            return {key: value for key, value in job.items() if key != "updated_at"}
