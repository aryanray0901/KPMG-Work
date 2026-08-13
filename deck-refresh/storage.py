"""Vercel Blob-backed persistence for session working directories.

Vercel Functions run in ephemeral, isolated containers. A session's files
saved to local disk in one request are not guaranteed to exist on the
container that serves the next request for the same session -- even two
requests fired moments apart (e.g. parallel slide-thumbnail fetches) can
land on different containers. This module mirrors a session's local
directory to Vercel Blob storage so state survives across requests.

Locally, or whenever BLOB_READ_WRITE_TOKEN isn't set, every function here
is a no-op and the app relies purely on local disk, exactly as before.
"""
from __future__ import annotations

import mimetypes
import os

import requests

BLOB_TOKEN = os.environ.get("BLOB_READ_WRITE_TOKEN")
ENABLED = bool(BLOB_TOKEN)

_BASE_URL = "https://blob.vercel-storage.com"
_API_VERSION = "10"
_TIMEOUT = 20


def _headers(extra=None):
    headers = {
        "authorization": f"Bearer {BLOB_TOKEN}",
        "x-api-version": _API_VERSION,
    }
    if extra:
        headers.update(extra)
    return headers


def put_file(pathname, local_path):
    """Upload a local file to Blob storage at an exact, stable pathname."""
    if not ENABLED:
        return
    content_type = mimetypes.guess_type(local_path)[0] or "application/octet-stream"
    with open(local_path, "rb") as f:
        data = f.read()
    resp = requests.put(
        f"{_BASE_URL}/{pathname}",
        data=data,
        headers=_headers({
            "access": "public",
            "x-content-type": content_type,
            "x-allow-overwrite": "1",
        }),
        timeout=_TIMEOUT,
    )
    resp.raise_for_status()


def list_prefix(prefix):
    """Return metadata dicts (with pathname/url) for every blob under a prefix."""
    if not ENABLED:
        return []
    blobs = []
    cursor = None
    while True:
        params = {"prefix": prefix, "limit": "1000"}
        if cursor:
            params["cursor"] = cursor
        resp = requests.get(_BASE_URL, params=params, headers=_headers(), timeout=_TIMEOUT)
        resp.raise_for_status()
        data = resp.json()
        blobs.extend(data.get("blobs", []))
        if not data.get("hasMore"):
            break
        cursor = data.get("cursor")
    return blobs


def download_to(url, local_path):
    if not ENABLED:
        return False
    resp = requests.get(url, timeout=_TIMEOUT)
    if resp.status_code != 200:
        return False
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    with open(local_path, "wb") as f:
        f.write(resp.content)
    return True


def delete_prefix(prefix):
    if not ENABLED:
        return
    blobs = list_prefix(prefix)
    urls = [b["url"] for b in blobs]
    if not urls:
        return
    requests.post(
        f"{_BASE_URL}/delete",
        json={"urls": urls},
        headers=_headers(),
        timeout=_TIMEOUT,
    )
