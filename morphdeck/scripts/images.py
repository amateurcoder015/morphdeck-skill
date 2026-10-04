#!/usr/bin/env python3
"""Resolve image references for morphdeck specs.

An image field in a spec can be:
    "photos/team.jpg"           a local file (relative to the spec)
    "https://…/pic.jpg"         a URL, downloaded once
    "stock:black hole galaxy"   a stock-photo search

Stock search uses Pexels when PEXELS_API_KEY is set, otherwise Openverse (no key;
Creative Commons images from Flickr, Wikimedia and others). Every download is
cached in `<spec dir>/images/` and comes back with a credit line, which the deck
puts in that slide's speaker notes.

CLI (handy for trying a query before putting it in a spec):
    python3 images.py "black hole galaxy" out_dir [--shape wide|square|tall]
"""
import hashlib
import io
import json
import os
import re
import shutil
import ssl
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request

UA = "morphdeck/1.0 (+https://github.com/amateurcoder015/morphdeck-skill)"
MIN_WIDTH = 800
MAX_TRIES = 6


def _ssl_context():
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        return ssl.create_default_context()


def _get(url, headers=None, timeout=25):
    headers = {"User-Agent": UA, **(headers or {})}
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=_ssl_context()) as r:
            return r.read()
    except urllib.error.URLError as e:
        # python.org builds on macOS often ship without root certificates; curl
        # uses the system keychain instead.
        if "CERTIFICATE_VERIFY_FAILED" not in str(e) or not shutil.which("curl"):
            raise
        cmd = ["curl", "-sfL", "--max-time", str(timeout)]
        for k, v in headers.items():
            cmd += ["-H", f"{k}: {v}"]
        return subprocess.run(cmd + [url], check=True, capture_output=True).stdout


def _slug(text):
    return re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")[:48] or "image"


def _valid_image(data):
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(data))
        im.verify()
        return True
    except Exception:
        return False


def _openverse(query, shape):
    params = {"q": query, "license_type": "commercial,modification", "page_size": "8",
              "mature": "false"}
    if shape:
        params["aspect_ratio"] = shape
    data = json.loads(_get("https://api.openverse.org/v1/images/?" + urllib.parse.urlencode(params), timeout=45))
    for r in data.get("results", []):
        if max(r.get("width") or 0, r.get("height") or 0) < MIN_WIDTH:
            continue
        who = r.get("creator") or "unknown"
        lic = f"CC {r.get('license', '').upper()} {r.get('license_version') or ''}".strip()
        if r.get("license") in ("pdm", "cc0"):
            lic = "public domain"
        yield r["url"], f"“{r.get('title') or query}” by {who} ({lic}) — {r.get('foreign_landing_url') or r['url']}"


def _pexels(query, shape, key):
    orient = {"wide": "landscape", "tall": "portrait", "square": "square"}.get(shape, "landscape")
    params = {"query": query, "per_page": "10", "orientation": orient}
    data = json.loads(_get("https://api.pexels.com/v1/search?" + urllib.parse.urlencode(params),
                           headers={"Authorization": key}))
    for p in data.get("photos", []):
        yield p["src"]["large2x"], f"Photo by {p['photographer']} on Pexels — {p['url']}"


def stock(query, cache_dir, shape="wide"):
    """Return (path, credit) for the first usable stock photo for `query`."""
    os.makedirs(cache_dir, exist_ok=True)
    base = os.path.join(cache_dir, f"stock-{_slug(query)}-{shape or 'any'}")
    if os.path.exists(base + ".jpg") and os.path.exists(base + ".txt"):
        return base + ".jpg", open(base + ".txt").read().strip()
    key = os.environ.get("PEXELS_API_KEY")
    def sources():
        yield from (_pexels(query, shape, key) if key else _openverse(query, shape))
        if shape and not key:  # nothing in that orientation: take any and crop
            yield from _openverse(query, None)
    for attempt, (url, credit) in enumerate(sources()):
        if attempt >= MAX_TRIES:
            break
        try:
            data = _get(url, timeout=20)
        except Exception:
            continue
        if _valid_image(data):
            with open(base + ".jpg", "wb") as f:
                f.write(data)
            with open(base + ".txt", "w") as f:
                f.write(credit)
            return base + ".jpg", credit
    raise RuntimeError(f"no usable stock image found for '{query}'")


def resolve(ref, spec_dir, shape="wide"):
    """Turn a spec image reference into (local_path, credit_or_None)."""
    cache = os.path.join(spec_dir, "images")
    if ref.startswith("stock:"):
        return stock(ref[6:].strip(), cache, shape)
    if ref.startswith(("http://", "https://")):
        os.makedirs(cache, exist_ok=True)
        ext = os.path.splitext(urllib.parse.urlparse(ref).path)[1] or ".jpg"
        path = os.path.join(cache, "url-" + hashlib.sha1(ref.encode()).hexdigest()[:12] + ext)
        if not os.path.exists(path):
            with open(path, "wb") as f:
                f.write(_get(ref))
        return path, f"Image: {ref}"
    path = ref if os.path.isabs(ref) else os.path.join(spec_dir, ref)
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    return path, None


def prefetch(jobs, spec_dir, workers=6):
    """Resolve many (ref, shape) pairs in parallel so a deck build isn't serial on the network."""
    from concurrent.futures import ThreadPoolExecutor

    def one(job):
        try:
            resolve(job[0], spec_dir, job[1])
        except Exception as e:
            print(f"warning: image '{job[0]}' failed: {e}", file=sys.stderr)

    jobs = list(dict.fromkeys(jobs))
    if jobs:
        with ThreadPoolExecutor(max_workers=workers) as ex:
            list(ex.map(one, jobs))


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("query")
    ap.add_argument("out_dir")
    ap.add_argument("--shape", default="wide", choices=["wide", "square", "tall"])
    a = ap.parse_args()
    try:
        p, c = stock(a.query, a.out_dir, a.shape)
    except RuntimeError as e:
        sys.exit(str(e))
    print(p)
    print(c)
