#!/usr/bin/env python3
"""Resolve image references for morphdeck specs.

An image field in a spec can be:
    "photos/team.jpg"           a local file (relative to the spec)
    "https://…/pic.jpg"         a URL, downloaded once
    "stock:black hole galaxy"   a stock-photo search
    "ai:<prompt>"               an AI-generated image (Cloudflare Workers AI, FLUX.1 schnell)

Stock search uses Pexels when PEXELS_API_KEY is set, otherwise Openverse (no key;
Creative Commons images from Flickr, Wikimedia and others). Every download is
cached in `<spec dir>/images/` and comes back with a credit line, which the deck
puts in that slide's speaker notes.

AI images need CF_ACCOUNT_ID and CF_API_TOKEN (environment, or a .env file in the
spec's folder, the current folder or ~/.config/morphdeck/.env). Cloudflare's free
allowance is 10,000 neurons a day, about 170 images. Without keys, "ai:" refs fall
back to a stock search on the same words.

CLI (handy for trying a query before putting it in a spec):
    python3 images.py "black hole galaxy" out_dir [--shape wide|square|tall]
    python3 images.py "ai:glowing accretion disk around a black hole, cinematic" out_dir
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


def _post_json(url, payload, headers, timeout=90):
    body = json.dumps(payload).encode()
    headers = {"User-Agent": UA, "Content-Type": "application/json", **headers}
    req = urllib.request.Request(url, data=body, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=_ssl_context()) as r:
            return json.loads(r.read())
    except urllib.error.HTTPError as e:
        return json.loads(e.read() or b"{}")
    except urllib.error.URLError as e:
        if "CERTIFICATE_VERIFY_FAILED" not in str(e) or not shutil.which("curl"):
            raise
        cmd = ["curl", "-sL", "--max-time", str(timeout), "-X", "POST", "--data-binary", "@-"]
        for k, v in headers.items():
            cmd += ["-H", f"{k}: {v}"]
        out = subprocess.run(cmd + [url], input=body, check=True, capture_output=True).stdout
        return json.loads(out)


def _load_env(*dirs):
    """Read KEY=VALUE lines from the first .env files found (no python-dotenv needed)."""
    for d in dirs:
        path = os.path.join(d, ".env")
        if not os.path.isfile(path):
            continue
        for ln in open(path):
            ln = ln.strip()
            if ln and not ln.startswith("#") and "=" in ln:
                k, v = ln.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


CF_MODEL = "@cf/black-forest-labs/flux-1-schnell"
_NOTED = set()


def ai_image(prompt, cache_dir, steps=6):
    """Generate with Cloudflare Workers AI. Returns (path, credit). Output is 1024x1024; the deck crops it."""
    os.makedirs(cache_dir, exist_ok=True)
    path = os.path.join(cache_dir, f"ai-{_slug(prompt)[:32]}-{hashlib.sha1(prompt.encode()).hexdigest()[:8]}.jpg")
    credit = f"AI-generated image (Cloudflare Workers AI, FLUX.1 schnell). Prompt: {prompt}"
    if os.path.exists(path):
        return path, credit
    acct, token = os.environ.get("CF_ACCOUNT_ID"), os.environ.get("CF_API_TOKEN")
    if not (acct and token):
        raise KeyError("CF_ACCOUNT_ID / CF_API_TOKEN not set")
    data = _post_json(f"https://api.cloudflare.com/client/v4/accounts/{acct}/ai/run/{CF_MODEL}",
                      {"prompt": prompt, "steps": steps}, {"Authorization": f"Bearer {token}"})
    b64 = (data.get("result") or {}).get("image")
    if not b64:
        raise RuntimeError(f"Cloudflare error: {data.get('errors') or data}")
    import base64
    with open(path, "wb") as f:
        f.write(base64.b64decode(b64))
    return path, credit


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


def web_copy(path, max_px=1920, max_bytes=900_000):
    """Return a slide-sized JPEG of `path` (cached next to it) so decks stay small."""
    if os.path.getsize(path) <= max_bytes:
        return path
    from PIL import Image
    out = os.path.splitext(path)[0] + f"-{max_px}.jpg"
    if os.path.exists(out):
        return out
    im = Image.open(path)
    im.thumbnail((max_px, max_px))
    if im.mode not in ("RGB", "L"):
        im = im.convert("RGB")
    im.save(out, "JPEG", quality=85, optimize=True, progressive=True)
    return out


def resolve(ref, spec_dir, shape="wide"):
    """Turn a spec image reference into (local_path, credit_or_None)."""
    cache = os.path.join(spec_dir, "images")
    if ref.startswith("ai:"):
        prompt = ref[3:].strip()
        _load_env(spec_dir, os.getcwd(), os.path.expanduser("~/.config/morphdeck"))
        try:
            return ai_image(prompt, cache)
        except KeyError:
            # no Cloudflare keys: fall back to a stock search on the prompt's leading words
            words = re.sub(r"[^\w\s]", " ", prompt).split()[:6]
            if prompt not in _NOTED:
                _NOTED.add(prompt)
                print(f"note: no Cloudflare keys, using stock search for '{' '.join(words)}'", file=sys.stderr)
            return stock(" ".join(words), cache, shape)
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


def _resolve_into(ref, out_dir, shape):
    if ref.startswith("ai:"):
        _load_env(os.getcwd(), os.path.expanduser("~/.config/morphdeck"))
        return ai_image(ref[3:].strip(), out_dir)
    return stock(ref[6:].strip(), out_dir, shape)


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("query")
    ap.add_argument("out_dir")
    ap.add_argument("--shape", default="wide", choices=["wide", "square", "tall"])
    a = ap.parse_args()
    try:
        ref = a.query if a.query.startswith(("ai:", "stock:")) else "stock:" + a.query
        p, c = _resolve_into(ref, a.out_dir, a.shape)
    except (RuntimeError, KeyError) as e:
        sys.exit(str(e))
    print(p)
    print(c)
