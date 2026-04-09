from __future__ import annotations

import json
import re
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Dict, List, Optional
import unicodedata


@dataclass(frozen=True)
class TaigaConfig:
    base_url: str
    project_slug: str
    username: str
    password: str
    qc_names: tuple[str, ...]

    @classmethod
    def from_path(cls, path: Path) -> "TaigaConfig":
        data = json.loads(path.read_text(encoding="utf-8"))
        required = ["baseUrl", "projectSlug", "username", "password"]
        missing = [key for key in required if not data.get(key)]
        if missing:
            raise ValueError(
                "Missing required Taiga config keys: " + ", ".join(sorted(missing))
            )
        qc_names = tuple(
            str(name).strip()
            for name in data.get("qcNames", [])
            if str(name).strip()
        )
        return cls(
            base_url=str(data["baseUrl"]).rstrip("/"),
            project_slug=str(data["projectSlug"]).strip(),
            username=str(data["username"]).strip(),
            password=str(data["password"]),
            qc_names=qc_names,
        )


class TaigaClient:
    def __init__(self, config: TaigaConfig) -> None:
        self.config = config
        self.api_base = f"{self.config.base_url}/api/v1"
        self.auth_token: Optional[str] = None

    def login(self) -> None:
        payload = {
            "type": "normal",
            "username": self.config.username,
            "password": self.config.password,
        }
        response = self._request_json(
            method="POST",
            path="/auth",
            data=payload,
            authorized=False,
        )
        token = response.get("auth_token")
        if not token:
            raise RuntimeError("Taiga login succeeded but no auth token was returned.")
        self.auth_token = token

    def enrich(self, item_type: str, ref: str) -> Dict[str, str]:
        normalized_type = item_type.strip().casefold()
        if not ref:
            return {"Status": "", "QC PIC": ""}

        if not self.auth_token:
            self.login()

        lookup_order = self._lookup_order(normalized_type)
        last_error: Optional[Exception] = None

        for resource_type in lookup_order:
            try:
                if resource_type == "issue":
                    payload = self._request_json(
                        method="GET",
                        path="/issues/by_ref",
                        query={"ref": ref, "project__slug": self.config.project_slug},
                        authorized=True,
                    )
                    status = self._extract_status(payload)
                    pic = self._extract_issue_pic(payload)
                    return {"Status": status, "QC PIC": self._filter_qc_names(pic)}

                payload = self._request_json(
                    method="GET",
                    path="/userstories/by_ref",
                    query={"ref": ref, "project__slug": self.config.project_slug},
                    authorized=True,
                )
                status = self._extract_status(payload)
                pic = self._extract_assigned_to(payload)
                return {"Status": status, "QC PIC": self._filter_qc_names(pic)}
            except RuntimeError as exc:
                last_error = exc
                if " 404 " not in f" {exc} " and "No UserStory matches" not in str(exc) and "No Issue matches" not in str(exc):
                    raise

        if last_error:
            raise last_error
        return {"Status": "", "QC PIC": ""}

    def _is_issue_type(self, normalized_type: str) -> bool:
        return normalized_type in {"issue", "fix bug", "bug", "bug fix"}

    def _lookup_order(self, normalized_type: str) -> List[str]:
        if self._is_issue_type(normalized_type):
            return ["issue", "userstory"]
        return ["userstory", "issue"]

    def _extract_status(self, payload: Dict[str, Any]) -> str:
        for key in ("status_extra_info", "status_extra", "status"):
            value = payload.get(key)
            if isinstance(value, dict) and value.get("name"):
                return str(value["name"]).strip()
        if payload.get("status"):
            return str(payload["status"]).strip()
        return ""

    def _extract_assigned_to(self, payload: Dict[str, Any]) -> str:
        assigned_to_extra = payload.get("assigned_to_extra_info")
        if isinstance(assigned_to_extra, dict):
            return self._user_display_name(assigned_to_extra)
        return ""

    def _extract_issue_pic(self, payload: Dict[str, Any]) -> str:
        issue_id = payload.get("id")
        watcher_names: List[str] = []
        if isinstance(issue_id, int):
            try:
                watchers = self._request_json(
                    method="GET",
                    path=f"/issues/{issue_id}/watchers",
                    authorized=True,
                )
            except Exception:
                watchers = []
            if isinstance(watchers, list):
                for watcher in watchers:
                    if isinstance(watcher, dict):
                        name = self._user_display_name(watcher)
                        if name:
                            watcher_names.append(name)

        if watcher_names:
            return ", ".join(dict.fromkeys(watcher_names))
        return self._extract_assigned_to(payload)

    def _filter_qc_names(self, raw_names: str) -> str:
        if not raw_names:
            return ""
        if not self.config.qc_names:
            return raw_names

        kept: List[str] = []
        for raw_name in [item.strip() for item in raw_names.split(",") if item.strip()]:
            canonical = self._match_qc_name(raw_name)
            if canonical:
                kept.append(canonical)
        return ", ".join(dict.fromkeys(kept))

    def _user_display_name(self, payload: Dict[str, Any]) -> str:
        for key in ("full_name_display", "full_name", "username"):
            value = payload.get(key)
            if value:
                return str(value).strip()
        return ""

    def _match_qc_name(self, raw_name: str) -> str:
        candidate = self._normalize_name(raw_name)
        if not candidate:
            return ""

        best_name = ""
        best_score = 0.0
        candidate_tokens = set(candidate.split())

        for qc_name in self.config.qc_names:
            normalized_qc = self._normalize_name(qc_name)
            if not normalized_qc:
                continue

            if candidate == normalized_qc:
                return qc_name

            qc_tokens = set(normalized_qc.split())
            if candidate_tokens and qc_tokens:
                if candidate_tokens.issubset(qc_tokens) or qc_tokens.issubset(candidate_tokens):
                    return qc_name

            score = SequenceMatcher(None, candidate, normalized_qc).ratio()
            if candidate in normalized_qc or normalized_qc in candidate:
                score = max(score, 0.91)

            overlap = len(candidate_tokens & qc_tokens)
            if overlap:
                score = max(score, min(0.9, 0.65 + overlap * 0.12))

            if score > best_score:
                best_score = score
                best_name = qc_name

        if best_score >= 0.78:
            return best_name
        return ""

    def _normalize_name(self, value: str) -> str:
        text = unicodedata.normalize("NFKD", value)
        text = "".join(ch for ch in text if not unicodedata.combining(ch))
        text = text.casefold()
        text = re.sub(r"[^a-z0-9]+", " ", text)
        return re.sub(r"\s+", " ", text).strip()

    def _request_json(
        self,
        method: str,
        path: str,
        data: Optional[Dict[str, Any]] = None,
        query: Optional[Dict[str, Any]] = None,
        authorized: bool = True,
    ) -> Any:
        url = f"{self.api_base}{path}"
        if query:
            url = f"{url}?{urllib.parse.urlencode(query)}"

        headers = {"Content-Type": "application/json"}
        if authorized:
            if not self.auth_token:
                raise RuntimeError("Taiga request attempted before login.")
            headers["Authorization"] = f"Bearer {self.auth_token}"

        body = None
        if data is not None:
            body = json.dumps(data).encode("utf-8")

        request = urllib.request.Request(url, data=body, headers=headers, method=method)
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                return json.load(response)
        except urllib.error.HTTPError as exc:
            details = exc.read().decode("utf-8", errors="ignore")
            raise RuntimeError(
                f"Taiga API error {exc.code} on {path}: {details or exc.reason}"
            ) from exc
