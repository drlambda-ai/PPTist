"""File storage for decks: Vercel Blob when BLOB_READ_WRITE_TOKEN is set, else a local folder (dev).

Every file is written once under a new random name, so nothing is ever overwritten or served stale from a cache.
"""
import json
import os
import urllib.parse
import urllib.request
import uuid

TOKEN = os.environ.get("BLOB_READ_WRITE_TOKEN", "")
LOCAL = os.environ.get("DECK_DATA", os.path.join(os.path.dirname(os.path.dirname(__file__)), ".deck-data"))
LOCAL_PREFIX = "/blob/"  # the dev server serves LOCAL here


def put(name, data, content_type="application/octet-stream"):
    """Store bytes, return the URL the browser and later calls read them from."""
    path = f"decks/{uuid.uuid4().hex}/{name}"
    if not TOKEN:
        full = os.path.join(LOCAL, path)
        os.makedirs(os.path.dirname(full), exist_ok=True)
        open(full, "wb").write(data)
        return LOCAL_PREFIX + path
    req = urllib.request.Request(
        "https://vercel.com/api/blob/?" + urllib.parse.urlencode({"pathname": path}), data=data, method="PUT",
        headers={"authorization": f"Bearer {TOKEN}", "x-api-version": "13", "x-vercel-blob-access": "public",
                 "x-content-type": content_type, "x-add-random-suffix": "0"})
    with urllib.request.urlopen(req, timeout=120) as r:
        return json.load(r)["url"]


def drop(url):
    """Delete a file this app stored (only our own store's deck/ files; best effort)."""
    host = urllib.parse.urlparse(url).hostname or ""
    if not TOKEN or not host.endswith(".blob.vercel-storage.com") or "/deck/" not in url:
        return False
    req = urllib.request.Request("https://vercel.com/api/blob/delete", data=json.dumps({"urls": [url]}).encode(),
                                 method="POST", headers={"authorization": f"Bearer {TOKEN}", "x-api-version": "13",
                                                         "content-type": "application/json"})
    with urllib.request.urlopen(req, timeout=30):
        return True


def get(url):
    if url.startswith(LOCAL_PREFIX):
        rel = os.path.normpath(url[len(LOCAL_PREFIX):])
        if rel.startswith(".."):
            raise ValueError("bad path")
        return open(os.path.join(LOCAL, rel), "rb").read()
    host = urllib.parse.urlparse(url).hostname or ""
    if not host.endswith(".blob.vercel-storage.com"):  # only our own store: never fetch arbitrary URLs
        raise ValueError("file must be uploaded to this app's storage")
    with urllib.request.urlopen(url, timeout=120) as r:
        return r.read()
