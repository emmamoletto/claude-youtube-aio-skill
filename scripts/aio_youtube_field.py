#!/usr/bin/env python3
"""
aio_youtube_field.py: measure how often YouTube is cited in Google's AI
Overviews for a list of keywords, whose videos get cited, and how much of
that space belongs to your own channels.

SERP data comes from an SEO connector. SE Ranking (DATA_getSerpResults,
advanced) is parsed natively; any other provider's response is parsed in
generic mode, which looks for the AI Overview block and the YouTube links
inside it. Claude runs one call per keyword and this script ingests the saved
response. Video and channel data come from the YouTube Data API v3.

Subcommands, all working on one audit folder (--work):

  init      Create the folder with the keyword list and your channels.
  status    Show which keywords still need a SERP check for a given day.
  ingest    Parse one saved SERP response into runs/<date>.json.
  resolve   Look up every cited video and its channel (YouTube API).
  channels  Print the cited channels, for classification.
  build     Write the Excel report, aggregating every day collected.

Only the Python standard library is needed, except for `build`, which needs
openpyxl. The API key is read from YOUTUBE_API_KEY.
"""

import argparse
import glob
import json
import os
import re
import shutil
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter, defaultdict
from datetime import date, datetime, timezone

API = "https://www.googleapis.com/youtube/v3/"
QUOTA_USED = 0


# --------------------------------------------------------------- utilities

def load_json(path, default=None):
    if not os.path.exists(path):
        return default
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save_json(path, data):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def slug(text):
    s = re.sub(r"[^\w]+", "-", text.lower(), flags=re.UNICODE).strip("-")
    return s[:80] or "keyword"


def norm_kw(text):
    return re.sub(r"\s+", " ", text.strip().lower())


def read_keywords(work):
    path = os.path.join(work, "keywords.txt")
    if not os.path.exists(path):
        raise SystemExit(f"No keywords.txt in {work}. Run init first.")
    seen, out = set(), []
    with open(path, encoding="utf-8") as f:
        for line in f:
            kw = line.strip()
            if kw and not kw.startswith("#") and norm_kw(kw) not in seen:
                seen.add(norm_kw(kw))
                out.append(kw)
    return out


# ---------------------------------------------------------------- YouTube

def api_key():
    key = os.environ.get("YOUTUBE_API_KEY")
    if not key:
        raise SystemExit("Missing API key: set the YOUTUBE_API_KEY environment variable.")
    return key


def api_get(endpoint, params, key, retries=3):
    global QUOTA_USED
    url = API + endpoint + "?" + urllib.parse.urlencode(dict(params, key=key))
    last = None
    for attempt in range(retries):
        try:
            with urllib.request.urlopen(url, timeout=30) as resp:
                QUOTA_USED += 1
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "replace")
            try:
                reason = json.loads(body)["error"]["errors"][0].get("reason", "")
            except Exception:
                reason = ""
            if e.code in (500, 503) and attempt < retries - 1:
                time.sleep(2 * (attempt + 1))
                continue
            raise SystemExit(f"YouTube API error {e.code} on {endpoint}: {reason or body[:300]}")
        except urllib.error.URLError as e:
            last = e
            time.sleep(2 * (attempt + 1))
    raise SystemExit(f"Network error calling {endpoint}: {last}")


def parse_channel_ref(ref):
    ref = ref.strip()
    m = re.search(r"(UC[0-9A-Za-z_-]{22})", ref)
    if m:
        return "id", m.group(1)
    m = re.search(r"@([0-9A-Za-z_.\-·]+)", ref)
    if m:
        return "forHandle", "@" + m.group(1)
    return "forHandle", ref if ref.startswith("@") else "@" + ref


def iso_duration_seconds(iso):
    m = re.fullmatch(r"P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?", iso or "")
    if not m:
        return 0
    d, h, mi, s = (int(x) if x else 0 for x in m.groups())
    return d * 86400 + h * 3600 + mi * 60 + s


def fmt_seconds(sec):
    h, rem = divmod(int(sec), 3600)
    m, s = divmod(rem, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"


# ------------------------------------------------------------ SERP parsing

YT_RE = re.compile(
    r"https?://(?:www\.|m\.)?(?:youtube\.com/(?:watch\?[^\s\"'<>)\]]*?v=|shorts/|embed/|live/)"
    r"|youtu\.be/)([0-9A-Za-z_-]{11})[^\s\"'<>)\]]*")


def parse_t(value):
    """'243', '243s', '4m3s', '1h2m3s' -> seconds."""
    if not value:
        return None
    value = value.strip()
    if value.isdigit():
        return int(value)
    m = re.fullmatch(r"(?:(\d+)h)?(?:(\d+)m)?(?:(\d+)s)?", value)
    if m and any(m.groups()):
        h, mi, s = (int(x) if x else 0 for x in m.groups())
        return h * 3600 + mi * 60 + s
    return None


def youtube_links(text):
    """Every YouTube video link in a string: (video_id, url, timestamp_s, is_shorts_url)."""
    out = []
    for m in YT_RE.finditer(text or ""):
        url = m.group(0).rstrip(".,;")
        q = urllib.parse.urlparse(url)
        params = urllib.parse.parse_qs(q.query)
        t = parse_t((params.get("t") or params.get("start") or [None])[0])
        out.append((m.group(1), url, t, "/shorts/" in url))
    return out


def find_results(obj):
    """Yield every SERP result block ({request_metadata, items}) in an MCP response."""
    if isinstance(obj, dict):
        if "items" in obj and isinstance(obj.get("items"), list) and "request_metadata" in obj:
            yield obj
            return
        for v in obj.values():
            yield from find_results(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from find_results(v)


def load_mcp_file(path):
    raw = open(path, encoding="utf-8").read()
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        start = min([i for i in (raw.find("{"), raw.find("[")) if i >= 0], default=-1)
        if start < 0:
            raise SystemExit(f"No JSON found in {path}")
        data = json.loads(raw[start:])
    # MCP content wrapper: [{"type": "text", "text": "<json>"}]
    if isinstance(data, list) and data and all(isinstance(x, dict) and x.get("type") == "text" for x in data):
        parts = []
        for x in data:
            try:
                parts.append(json.loads(x.get("text", "")))
            except json.JSONDecodeError:
                pass
        if parts:
            data = parts[0] if len(parts) == 1 else parts
    return data


def parse_serp(block):
    meta = block.get("request_metadata", {})
    items = block.get("items", [])
    aio_items = [i for i in items if i.get("type") == "ai_overview"]
    videos = {}
    sources_total = 0
    for aio in aio_items:
        text = aio.get("overview") or ""
        sources = aio.get("sources") or []
        sources_total += len(sources)
        for pos, src in enumerate(sources, start=1):
            for vid, url, t, is_short in youtube_links(src.get("url", "")):
                v = videos.setdefault(vid, {"video_id": vid, "urls": [], "in_sources": False,
                                            "in_text": False, "source_position": None,
                                            "source_title": "", "timestamps": [],
                                            "shorts_url": False})
                v["in_sources"] = True
                v["source_position"] = v["source_position"] or pos
                v["source_title"] = v["source_title"] or src.get("title", "")
                if url not in v["urls"]:
                    v["urls"].append(url)
                if t is not None and t not in v["timestamps"]:
                    v["timestamps"].append(t)
                v["shorts_url"] = v["shorts_url"] or is_short
        for vid, url, t, is_short in youtube_links(text):
            v = videos.setdefault(vid, {"video_id": vid, "urls": [], "in_sources": False,
                                        "in_text": False, "source_position": None,
                                        "source_title": "", "timestamps": [],
                                        "shorts_url": False})
            v["in_text"] = True
            if url not in v["urls"]:
                v["urls"].append(url)
            if t is not None and t not in v["timestamps"]:
                v["timestamps"].append(t)
            v["shorts_url"] = v["shorts_url"] or is_short
    organic_yt = sum(1 for i in items if i.get("type") == "organic"
                     and "youtube.com" in (i.get("domain") or i.get("url") or ""))
    crawled = (meta.get("crawled_at") or "")[:10]
    return {
        "keyword": meta.get("keyword", ""),
        "crawled_at": meta.get("crawled_at", ""),
        "crawled_date": crawled,
        "language_code": meta.get("language_code", ""),
        "aio_present": bool(aio_items),
        "aio_sources": sources_total,
        "aio_text_chars": sum(len(i.get("overview") or "") for i in aio_items),
        "aio_label_mentions_ai_mode": any("AI Mode" in (i.get("overview") or "")[:300] for i in aio_items),
        "youtube_videos": list(videos.values()),
        "organic_youtube_results": organic_yt,
        "serp_features": block.get("summary", {}).get("SERP_features", []),
    }


AIO_TYPES = {"ai_overview", "ai_overviews", "aio", "ai_mode", "google_ai_overview", "sge"}
URL_KEYS = {"url", "link", "source_url", "href", "source_link"}


def find_aio_blocks(obj):
    """Yield every AI Overview subtree in a provider response of unknown shape."""
    if isinstance(obj, dict):
        if str(obj.get("type", "")).strip().lower() in AIO_TYPES:
            yield obj
            return
        for k, v in obj.items():
            if "ai_overview" in str(k).lower() and v:
                yield v
            else:
                yield from find_aio_blocks(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from find_aio_blocks(v)


def url_fields(obj):
    """Every (url, title) stored under a url-like key in a subtree."""
    if isinstance(obj, dict):
        title = obj.get("title") if isinstance(obj.get("title"), str) else ""
        for k, v in obj.items():
            if str(k).lower() in URL_KEYS and isinstance(v, str):
                yield v, title
            else:
                yield from url_fields(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from url_fields(v)


def find_organic_youtube(obj):
    if isinstance(obj, dict):
        if str(obj.get("type", "")).lower() == "organic":
            target = " ".join(str(obj.get(k, "")) for k in ("domain", "url", "link"))
            return 1 if "youtube.com" in target else 0
        return sum(find_organic_youtube(v) for v in obj.values())
    if isinstance(obj, list):
        return sum(find_organic_youtube(v) for v in obj)
    return 0


def parse_generic(data, keyword):
    """Provider-agnostic parsing: AI Overview block -> YouTube links in it."""
    blocks = list(find_aio_blocks(data))
    videos = {}
    sources, titles = [], {}
    for block in blocks:
        for u, title in url_fields(block):
            if u not in sources:
                sources.append(u)
            titles[u] = titles.get(u) or title
    source_set = set(sources)
    for block in blocks:
        for vid, url, t, is_short in youtube_links(json.dumps(block, ensure_ascii=False)):
            v = videos.setdefault(vid, {"video_id": vid, "urls": [], "in_sources": False,
                                        "in_text": False, "source_position": None,
                                        "source_title": "", "timestamps": [],
                                        "shorts_url": False})
            in_src = any(src.startswith(url) or url.startswith(src) for src in source_set
                         if vid in src)
            if in_src:
                v["in_sources"] = True
                pos = next((i for i, src in enumerate(sources, start=1) if vid in src), None)
                v["source_position"] = v["source_position"] or pos
                v["source_title"] = v["source_title"] or titles.get(sources[pos - 1], "") if pos else v["source_title"]
            else:
                v["in_text"] = True
            if url not in v["urls"]:
                v["urls"].append(url)
            if t is not None and t not in v["timestamps"]:
                v["timestamps"].append(t)
            v["shorts_url"] = v["shorts_url"] or is_short
    text = json.dumps(blocks, ensure_ascii=False)
    return {
        "keyword": keyword,
        "crawled_at": "",
        "crawled_date": "",
        "language_code": "",
        "aio_present": bool(blocks),
        "aio_sources": len(sources),
        "aio_text_chars": len(text) if blocks else 0,
        "aio_label_mentions_ai_mode": "AI Mode" in text[:2000],
        "youtube_videos": list(videos.values()),
        "organic_youtube_results": find_organic_youtube(data),
        "serp_features": [],
        "parser": "generic",
    }


# -------------------------------------------------------------- commands

def cmd_init(args):
    os.makedirs(args.work, exist_ok=True)
    kw_path = os.path.join(args.work, "keywords.txt")
    if args.keywords:
        kws = []
        for item in args.keywords:
            if os.path.exists(item):
                kws.extend(l.strip() for l in open(item, encoding="utf-8") if l.strip())
            else:
                kws.append(item.strip())
        with open(kw_path, "w", encoding="utf-8") as f:
            f.write("\n".join(kws) + "\n")
    cfg_path = os.path.join(args.work, "config.json")
    cfg = load_json(cfg_path, {})
    cfg.update({
        "own_channels": args.own or cfg.get("own_channels", []),
        "language_code": args.language or cfg.get("language_code", ""),
        "country_code": args.country or cfg.get("country_code", ""),
        "created_at": cfg.get("created_at") or datetime.now(timezone.utc).isoformat(timespec="seconds"),
    })
    save_json(cfg_path, cfg)
    n = len(read_keywords(args.work)) if os.path.exists(kw_path) else 0
    print(f"Ready: {args.work} with {n} keywords, {len(cfg['own_channels'])} own channel(s), "
          f"market {cfg['language_code']}-{cfg['country_code']}. "
          f"SERP calls per full run: {n} (with SE Ranking advanced SERPs: {n * 10} credits).")


def cmd_status(args):
    kws = read_keywords(args.work)
    day = args.date or date.today().isoformat()
    run = load_json(os.path.join(args.work, "runs", f"{day}.json"), {"keywords": {}})
    done = {norm_kw(k) for k in run["keywords"]}
    todo = [k for k in kws if norm_kw(k) not in done]
    days = sorted(os.path.basename(p)[:-5] for p in glob.glob(os.path.join(args.work, "runs", "*.json")))
    print(json.dumps({
        "date": day, "keywords_total": len(kws), "checked_today": len(kws) - len(todo),
        "to_check": todo, "serp_calls_needed": len(todo),
        "se_ranking_credits_needed": len(todo) * 10, "days_collected": days,
    }, ensure_ascii=False, indent=1))


def cmd_ingest(args):
    data = load_mcp_file(args.serp_file)
    blocks = list(find_results(data)) if args.format in ("auto", "seranking") else []
    if args.format == "seranking" and not blocks:
        raise SystemExit(f"No SE Ranking SERP results in {args.serp_file}. Is it an advanced "
                         f"DATA_getSerpResults response? For other providers use --format generic.")
    if blocks:
        records = [parse_serp(b) for b in blocks]
    else:
        if not args.keyword:
            raise SystemExit("Not an SE Ranking response: pass --keyword so the generic parser "
                             "knows which keyword this SERP belongs to.")
        records = [parse_generic(data, args.keyword)]
        if not records[0]["aio_present"]:
            print("  note: no AI Overview block found. If the provider returned one under another "
                  "name, check the raw file before trusting this result.", file=sys.stderr)
    kws = {norm_kw(k): k for k in read_keywords(args.work)} if os.path.exists(
        os.path.join(args.work, "keywords.txt")) else {}
    for rec in records:
        if args.keyword and len(records) == 1:
            rec["keyword"] = args.keyword
        if not rec["keyword"]:
            raise SystemExit("The response has no keyword: pass --keyword.")
        key = kws.get(norm_kw(rec["keyword"]), rec["keyword"])
        day = args.date or rec["crawled_date"] or date.today().isoformat()
        run_path = os.path.join(args.work, "runs", f"{day}.json")
        run = load_json(run_path, {"date": day, "keywords": {}})
        rec["ingested_at"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
        run["keywords"][key] = rec
        save_json(run_path, run)
        raw_dir = os.path.join(args.work, "raw", day)
        os.makedirs(raw_dir, exist_ok=True)
        shutil.copyfile(args.serp_file, os.path.join(raw_dir, slug(key) + ".json"))
        yt = rec["youtube_videos"]
        print(f"{day} | {key} | AI answer: {'yes' if rec['aio_present'] else 'no'} | "
              f"sources: {rec['aio_sources']} | YouTube videos: {len(yt)} | "
              f"with timestamp: {sum(1 for v in yt if v['timestamps'])}")
        if key not in kws.values() and kws:
            print(f"  note: '{key}' is not in keywords.txt", file=sys.stderr)


def all_runs(work):
    runs = []
    for p in sorted(glob.glob(os.path.join(work, "runs", "*.json"))):
        runs.append(load_json(p))
    return runs


def cmd_resolve(args):
    key = api_key()
    cfg = load_json(os.path.join(args.work, "config.json"), {})
    vcache_path = os.path.join(args.work, "videos_cache.json")
    ccache_path = os.path.join(args.work, "channels_cache.json")
    vcache = load_json(vcache_path, {})
    ccache = load_json(ccache_path, {})

    # Own channels: resolve each reference to a channel ID once.
    own = cfg.get("own_channel_ids", {})
    for ref in cfg.get("own_channels", []):
        if ref in own.values() or any(r == ref for r in cfg.get("own_resolved_refs", [])):
            continue
        kind, value = parse_channel_ref(ref)
        data = api_get("channels", {"part": "snippet", kind: value}, key)
        items = data.get("items") or []
        if not items:
            print(f"Own channel not found: {ref}", file=sys.stderr)
            continue
        own[items[0]["id"]] = items[0]["snippet"].get("title", "")
        cfg.setdefault("own_resolved_refs", []).append(ref)
    cfg["own_channel_ids"] = own
    save_json(os.path.join(args.work, "config.json"), cfg)

    ids = sorted({v["video_id"] for run in all_runs(args.work)
                  for rec in run["keywords"].values() for v in rec["youtube_videos"]})
    missing = [i for i in ids if i not in vcache]
    for i in range(0, len(missing), 50):
        batch = missing[i:i + 50]
        data = api_get("videos", {"part": "snippet,contentDetails,statistics",
                                  "id": ",".join(batch), "maxResults": 50}, key)
        found = set()
        for it in data.get("items", []):
            sn, cd, st = it["snippet"], it.get("contentDetails", {}), it.get("statistics", {})
            dur = iso_duration_seconds(cd.get("duration"))
            vcache[it["id"]] = {
                "title": sn.get("title", ""), "channel_id": sn.get("channelId", ""),
                "channel_title": sn.get("channelTitle", ""), "published_at": sn.get("publishedAt", ""),
                "duration_seconds": dur, "duration": fmt_seconds(dur) if dur else "",
                "views": int(st["viewCount"]) if "viewCount" in st else None,
                "audio_language": sn.get("defaultAudioLanguage", ""),
                "available": True,
            }
            found.add(it["id"])
        for vid in batch:
            if vid not in found:
                vcache[vid] = {"available": False, "title": "", "channel_id": "", "channel_title": ""}
    save_json(vcache_path, vcache)

    ch_ids = sorted({v["channel_id"] for v in vcache.values() if v.get("channel_id")})
    missing_ch = [c for c in ch_ids if c not in ccache]
    for i in range(0, len(missing_ch), 50):
        batch = missing_ch[i:i + 50]
        data = api_get("channels", {"part": "snippet,statistics", "id": ",".join(batch),
                                    "maxResults": 50}, key)
        for it in data.get("items", []):
            sn, st = it["snippet"], it.get("statistics", {})
            ccache[it["id"]] = {
                "title": sn.get("title", ""), "handle": sn.get("customUrl", ""),
                "description": sn.get("description", "")[:600], "country": sn.get("country", ""),
                "subscribers": int(st["subscriberCount"]) if st.get("subscriberCount") else None,
                "subscribers_hidden": bool(st.get("hiddenSubscriberCount")),
                "video_count": int(st["videoCount"]) if st.get("videoCount") else None,
            }
    save_json(ccache_path, ccache)
    unavailable = sum(1 for i in ids if not vcache.get(i, {}).get("available"))
    print(f"Videos cited: {len(ids)} ({len(missing)} looked up now, {unavailable} unavailable). "
          f"Channels: {len(ch_ids)} ({len(missing_ch)} looked up now). Own channels: {len(own)}. "
          f"Quota used: {QUOTA_USED} units.")


def observations(work):
    """One row per (day, keyword, video) plus the per-keyword records."""
    vcache = load_json(os.path.join(work, "videos_cache.json"), {})
    rows, recs = [], []
    for run in all_runs(work):
        for kw, rec in run["keywords"].items():
            recs.append((run["date"], kw, rec))
            for v in rec["youtube_videos"]:
                meta = vcache.get(v["video_id"], {})
                rows.append({"date": run["date"], "keyword": kw, **v,
                             "channel_id": meta.get("channel_id", ""),
                             "channel_title": meta.get("channel_title", "")})
    return recs, rows, vcache


def cmd_channels(args):
    cfg = load_json(os.path.join(args.work, "config.json"), {})
    own = cfg.get("own_channel_ids", {})
    ccache = load_json(os.path.join(args.work, "channels_cache.json"), {})
    types = load_json(os.path.join(args.work, "channel-types.json"), {})
    recs, rows, vcache = observations(args.work)
    if rows and not ccache:
        raise SystemExit("Run resolve first.")
    apps = Counter(r["channel_id"] for r in rows if r["channel_id"])
    kws = defaultdict(set)
    titles = defaultdict(set)
    for r in rows:
        if r["channel_id"]:
            kws[r["channel_id"]].add(r["keyword"])
            titles[r["channel_id"]].add(vcache.get(r["video_id"], {}).get("title", ""))
    out = []
    for cid, n in apps.most_common():
        c = ccache.get(cid, {})
        if args.only_unclassified and (cid in types or cid in own):
            continue
        out.append({
            "channel_id": cid, "title": c.get("title", ""), "handle": c.get("handle", ""),
            "own_channel": cid in own, "appearances": n,
            "subscribers": c.get("subscribers"), "video_count": c.get("video_count"),
            "country": c.get("country", ""), "description": c.get("description", "")[:400],
            "cited_video_titles": sorted(t for t in titles[cid] if t)[:5],
            "keywords": sorted(kws[cid])[:5],
            "already_classified": cid in types,
        })
    print(json.dumps({"channels": len(out), "items": out}, ensure_ascii=False, indent=1))


# ------------------------------------------------------------------ build

LABELS = {
    "it": {
        "summary": "Riepilogo", "days": "Per giorno", "kw": "Keyword", "detail": "Dettaglio",
        "videos": "Video", "channels": "Canali", "info": "Info",
        "date": "Data", "period": "Totale periodo", "kw_checked": "Keyword controllate",
        "with_aio": "Con risposta AI", "pct_aio": "% con risposta AI",
        "with_yt": "Con YouTube", "pct_yt": "% con YouTube (su risposte AI)",
        "yt_apps": "Apparizioni YouTube", "unique_videos": "Video unici",
        "own_apps": "Apparizioni tuoi canali", "own_share": "Quota tuoi canali",
        "other_share": "Quota altri canali", "with_ts": "Citazioni con timestamp",
        "type_share": "Quota per tipo di canale (proposto)",
        "keyword": "Keyword", "days_checked": "Giorni controllati", "days_aio": "Giorni con risposta AI",
        "days_yt": "Giorni con YouTube", "pct_days_yt": "% giorni con YouTube",
        "videos_cited": "Video citati", "channels_cited": "Canali citati",
        "own_cited": "Tuo canale citato", "last_check": "Ultimo controllo",
        "aio": "Risposta AI", "yt": "YouTube", "n_videos": "N. video", "sources": "Fonti AI",
        "video_titles": "Video citati", "channel_names": "Canali", "timestamps": "Timestamp citati",
        "organic_yt": "Risultati YouTube organici",
        "title": "Titolo", "url": "URL", "channel": "Canale", "channel_type": "Tipo canale (proposto)",
        "own": "Tuo canale", "duration": "Durata", "format": "Formato", "published": "Pubblicato",
        "views": "Views", "keywords_cited": "Keyword in cui è citato", "appearances": "Apparizioni",
        "days_seen": "Giorni in cui compare", "ts_count": "Citazioni con timestamp",
        "handle": "Handle", "subscribers": "Iscritti", "channel_videos": "Video sul canale",
        "country": "Paese", "proposed": "Tipo proposto", "reason": "Motivo",
        "confirmed": "Tipo confermato", "yes": "Sì", "no": "No",
        "types": ["Creator", "Editore", "Brand", "Altro"],
        "own_type": "Tuo canale", "unknown": "Da classificare",
        "long": "Long", "short": "Short", "short_est": "Short (stima)",
        "info_rows": ["Keyword nella lista", "Giorni raccolti", "Primo giorno", "Ultimo giorno",
                      "Tuoi canali", "Generato il",
                      "Nota", "La risposta AI è il blocco AI Overview della SERP (ai_overview in SE Ranking); "
                              "a volte il testo riporta 'AI Mode', quindi non distingue AI Overview e AI Mode."],
    },
    "en": {
        "summary": "Summary", "days": "By day", "kw": "Keywords", "detail": "Detail",
        "videos": "Videos", "channels": "Channels", "info": "Info",
        "date": "Date", "period": "Whole period", "kw_checked": "Keywords checked",
        "with_aio": "With AI answer", "pct_aio": "% with AI answer",
        "with_yt": "With YouTube", "pct_yt": "% with YouTube (of AI answers)",
        "yt_apps": "YouTube appearances", "unique_videos": "Unique videos",
        "own_apps": "Your channels' appearances", "own_share": "Your channels' share",
        "other_share": "Other channels' share", "with_ts": "Timestamped citations",
        "type_share": "Share by channel type (proposed)",
        "keyword": "Keyword", "days_checked": "Days checked", "days_aio": "Days with AI answer",
        "days_yt": "Days with YouTube", "pct_days_yt": "% days with YouTube",
        "videos_cited": "Videos cited", "channels_cited": "Channels cited",
        "own_cited": "Your channel cited", "last_check": "Last check",
        "aio": "AI answer", "yt": "YouTube", "n_videos": "Videos", "sources": "AI sources",
        "video_titles": "Videos cited", "channel_names": "Channels", "timestamps": "Timestamps cited",
        "organic_yt": "Organic YouTube results",
        "title": "Title", "url": "URL", "channel": "Channel", "channel_type": "Channel type (proposed)",
        "own": "Your channel", "duration": "Duration", "format": "Format", "published": "Published",
        "views": "Views", "keywords_cited": "Keywords where cited", "appearances": "Appearances",
        "days_seen": "Days seen", "ts_count": "Timestamped citations",
        "handle": "Handle", "subscribers": "Subscribers", "channel_videos": "Videos on channel",
        "country": "Country", "proposed": "Proposed type", "reason": "Reason",
        "confirmed": "Confirmed type", "yes": "Yes", "no": "No",
        "types": ["Creator", "Publisher", "Brand", "Other"],
        "own_type": "Your channel", "unknown": "To classify",
        "long": "Long", "short": "Short", "short_est": "Short (estimate)",
        "info_rows": ["Keywords in list", "Days collected", "First day", "Last day",
                      "Your channels", "Generated at",
                      "Note", "The AI answer is the AI Overview block of the SERP (ai_overview in SE Ranking); "
                              "its text sometimes says 'AI Mode', so it does not separate AI Overviews from AI Mode."],
    },
}

TYPE_ALIASES = {"creator": 0, "editore": 1, "publisher": 1, "media": 1,
                "brand": 2, "altro": 3, "other": 3}


def type_label(t, L):
    idx = TYPE_ALIASES.get(str(t or "").strip().lower())
    return L["types"][idx] if idx is not None else (t or L["unknown"])


def mmss(sec):
    return fmt_seconds(sec) if sec is not None else ""


def cmd_build(args):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.worksheet.datavalidation import DataValidation
    except ImportError:
        raise SystemExit("openpyxl is missing: run  pip install openpyxl")
    L = LABELS[args.lang]
    work = args.work
    cfg = load_json(os.path.join(work, "config.json"), {})
    own = cfg.get("own_channel_ids", {})
    ccache = load_json(os.path.join(work, "channels_cache.json"), {})
    types = load_json(os.path.join(work, "channel-types.json"), {})
    recs, rows, vcache = observations(work)
    if not recs:
        raise SystemExit("No runs yet: ingest at least one SERP.")
    if rows and not vcache:
        print("Warning: videos not resolved yet, channel columns will be empty. Run resolve.", file=sys.stderr)

    def ctype(cid):
        if cid in own:
            return L["own_type"]
        if not cid:
            return L["unknown"]
        return type_label(types.get(cid, {}).get("type"), L)

    def video_format(vid, shorts_url):
        if shorts_url:
            return L["short"]
        d = vcache.get(vid, {}).get("duration_seconds")
        if d is not None and 0 < d <= 60:
            return L["short_est"]
        return L["long"] if d else ""

    wb = Workbook()
    hfill, hfont = PatternFill("solid", fgColor="1F2937"), Font(bold=True, color="FFFFFF")
    wrap = Alignment(wrap_text=True, vertical="top")

    def table(ws, headers, data_rows, widths, link_cols=()):
        ws.append(headers)
        for i, w in enumerate(widths, start=1):
            c = ws.cell(row=1, column=i)
            c.fill, c.font = hfill, hfont
            c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.column_dimensions[c.column_letter].width = w
        for r in data_rows:
            ws.append(r)
            for c in ws[ws.max_row]:
                c.alignment = wrap
            for col in link_cols:
                c = ws.cell(row=ws.max_row, column=col)
                if c.value:
                    c.hyperlink = c.value
                    c.font = Font(color="2563EB", underline="single")
        ws.freeze_panes = "B2"
        if ws.max_row > 1:
            ws.auto_filter.ref = ws.dimensions

    def pct(a, b):
        return round(a / b, 4) if b else None

    def stats(rec_subset, row_subset):
        n = len(rec_subset)
        aio = sum(1 for _, _, r in rec_subset if r["aio_present"])
        yt = sum(1 for _, _, r in rec_subset if r["aio_present"] and r["youtube_videos"])
        apps = len(row_subset)
        own_apps = sum(1 for r in row_subset if r["channel_id"] in own)
        by_type = Counter(ctype(r["channel_id"]) for r in row_subset)
        return {
            "n": n, "aio": aio, "pct_aio": pct(aio, n), "yt": yt, "pct_yt": pct(yt, aio),
            "apps": apps, "unique": len({r["video_id"] for r in row_subset}),
            "own": own_apps, "own_share": pct(own_apps, apps),
            "other_share": pct(apps - own_apps, apps) if apps else None,
            "ts": sum(1 for r in row_subset if r["timestamps"]),
            "by_type": by_type,
        }

    days = sorted({d for d, _, _ in recs})
    type_order = [L["own_type"]] + L["types"] + [L["unknown"]]

    # Summary: whole period, then by day.
    ws = wb.active
    ws.title = L["summary"]
    tot = stats(recs, rows)
    lines = [
        (L["kw_checked"], tot["n"]), (L["with_aio"], tot["aio"]), (L["pct_aio"], tot["pct_aio"]),
        (L["with_yt"], tot["yt"]), (L["pct_yt"], tot["pct_yt"]), (L["yt_apps"], tot["apps"]),
        (L["unique_videos"], tot["unique"]), (L["own_apps"], tot["own"]),
        (L["own_share"], tot["own_share"]), (L["other_share"], tot["other_share"]),
        (L["with_ts"], tot["ts"]),
    ]
    ws.append([L["period"] + f" ({days[0]} / {days[-1]}, {len(days)})", ""])
    ws.cell(row=1, column=1).font = Font(bold=True, size=13)
    for k, v in lines:
        ws.append([k, v])
        if isinstance(v, float):
            ws.cell(row=ws.max_row, column=2).number_format = "0.0%"
    ws.append([])
    ws.append([L["type_share"], ""])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
    for t in type_order:
        if tot["by_type"].get(t):
            ws.append([t, pct(tot["by_type"][t], tot["apps"])])
            ws.cell(row=ws.max_row, column=2).number_format = "0.0%"
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 16
    note_kw = "Nota: con un solo giorno il campione è debole, le risposte AI cambiano spesso." if args.lang == "it" \
        else "Note: a single day is a weak sample, AI answers change often."
    if len(days) == 1:
        ws.append([])
        ws.append([note_kw])

    # By day.
    ws = wb.create_sheet(L["days"])
    drows = []
    for d in days:
        s = stats([x for x in recs if x[0] == d], [r for r in rows if r["date"] == d])
        drows.append([d, s["n"], s["aio"], s["pct_aio"], s["yt"], s["pct_yt"], s["apps"],
                      s["unique"], s["own"], s["own_share"], s["ts"]])
    table(ws, [L["date"], L["kw_checked"], L["with_aio"], L["pct_aio"], L["with_yt"], L["pct_yt"],
               L["yt_apps"], L["unique_videos"], L["own_apps"], L["own_share"], L["with_ts"]],
          drows, [12, 12, 12, 12, 12, 14, 12, 12, 12, 12, 12])
    for row in ws.iter_rows(min_row=2):
        for idx in (3, 5, 9):
            row[idx].number_format = "0.0%"

    # Keywords, aggregated over days.
    ws = wb.create_sheet(L["kw"])
    by_kw = defaultdict(list)
    for d, kw, rec in recs:
        by_kw[kw].append((d, rec))
    krows = []
    for kw, lst in sorted(by_kw.items(), key=lambda x: -sum(1 for _, r in x[1] if r["youtube_videos"])):
        kw_rows = [r for r in rows if r["keyword"] == kw]
        days_yt = sum(1 for _, r in lst if r["aio_present"] and r["youtube_videos"])
        krows.append([
            kw, len(lst), sum(1 for _, r in lst if r["aio_present"]), days_yt, pct(days_yt, len(lst)),
            len({r["video_id"] for r in kw_rows}),
            ", ".join(sorted({r["channel_title"] for r in kw_rows if r["channel_title"]})),
            L["yes"] if any(r["channel_id"] in own for r in kw_rows) else L["no"],
            max(d for d, _ in lst),
        ])
    table(ws, [L["keyword"], L["days_checked"], L["days_aio"], L["days_yt"], L["pct_days_yt"],
               L["videos_cited"], L["channels_cited"], L["own_cited"], L["last_check"]],
          krows, [36, 10, 12, 12, 12, 10, 50, 10, 12])
    for row in ws.iter_rows(min_row=2):
        row[4].number_format = "0%"

    # Detail: one row per keyword per day.
    ws = wb.create_sheet(L["detail"])
    det = []
    for d, kw, rec in sorted(recs, key=lambda x: (x[0], x[1]), reverse=True):
        vids = rec["youtube_videos"]
        det.append([
            d, kw, L["yes"] if rec["aio_present"] else L["no"],
            L["yes"] if vids else L["no"], len(vids), rec["aio_sources"],
            "\n".join(vcache.get(v["video_id"], {}).get("title") or v["source_title"] or v["video_id"] for v in vids),
            "\n".join(sorted({vcache.get(v["video_id"], {}).get("channel_title", "") for v in vids} - {""})),
            "\n".join(f"{vcache.get(v['video_id'], {}).get('title') or v['video_id']} @ {mmss(t)}"
                      for v in vids for t in v["timestamps"]),
            rec.get("organic_youtube_results", 0),
        ])
    table(ws, [L["date"], L["keyword"], L["aio"], L["yt"], L["n_videos"], L["sources"],
               L["video_titles"], L["channel_names"], L["timestamps"], L["organic_yt"]],
          det, [12, 32, 10, 10, 8, 8, 50, 30, 40, 12])

    # Videos.
    ws = wb.create_sheet(L["videos"])
    by_vid = defaultdict(list)
    for r in rows:
        by_vid[r["video_id"]].append(r)
    vrows = []
    for vid, lst in sorted(by_vid.items(), key=lambda x: -len(x[1])):
        m = vcache.get(vid, {})
        cid = m.get("channel_id", "")
        shorts = any(r["shorts_url"] for r in lst)
        vrows.append([
            m.get("title") or lst[0]["source_title"] or vid,
            f"https://www.youtube.com/watch?v={vid}",
            m.get("channel_title", ""), ctype(cid), L["yes"] if cid in own else L["no"],
            m.get("duration", ""), video_format(vid, shorts), (m.get("published_at") or "")[:10],
            m.get("views"), ", ".join(sorted({r["keyword"] for r in lst})), len(lst),
            len({r["date"] for r in lst}), sum(1 for r in lst if r["timestamps"]),
        ])
    table(ws, [L["title"], L["url"], L["channel"], L["channel_type"], L["own"], L["duration"],
               L["format"], L["published"], L["views"], L["keywords_cited"], L["appearances"],
               L["days_seen"], L["ts_count"]],
          vrows, [45, 22, 26, 16, 9, 9, 12, 12, 10, 40, 11, 11, 11], link_cols=(2,))

    # Channels.
    ws = wb.create_sheet(L["channels"])
    by_ch = defaultdict(list)
    for r in rows:
        if r["channel_id"]:
            by_ch[r["channel_id"]].append(r)
    crows = []
    for cid, lst in sorted(by_ch.items(), key=lambda x: -len(x[1])):
        c = ccache.get(cid, {})
        t = types.get(cid, {})
        crows.append([
            c.get("title") or lst[0]["channel_title"], c.get("handle", ""),
            f"https://www.youtube.com/channel/{cid}",
            c.get("subscribers"), c.get("video_count"), c.get("country", ""),
            len(lst), len({r["keyword"] for r in lst}), len({r["video_id"] for r in lst}),
            L["own_type"] if cid in own else type_label(t.get("type"), L),
            t.get("reason", ""), L["own_type"] if cid in own else "",
            L["yes"] if cid in own else L["no"],
        ])
    table(ws, [L["channel"], L["handle"], L["url"], L["subscribers"], L["channel_videos"], L["country"],
               L["appearances"], L["keywords_cited"], L["videos_cited"], L["proposed"], L["reason"],
               L["confirmed"], L["own"]],
          crows, [30, 20, 22, 11, 10, 8, 11, 11, 11, 14, 50, 14, 9], link_cols=(3,))
    if ws.max_row > 1:
        dv = DataValidation(type="list", formula1='"' + ",".join(L["types"] + [L["own_type"]]) + '"',
                            allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"L2:L{ws.max_row}")

    # Info.
    ws = wb.create_sheet(L["info"])
    vals = [len(read_keywords(work)) if os.path.exists(os.path.join(work, "keywords.txt")) else "",
            len(days), days[0], days[-1], ", ".join(own.values()) or "-",
            datetime.now(timezone.utc).isoformat(timespec="seconds")]
    ir = L["info_rows"]
    for k, v in zip(ir[:6], vals):
        ws.append([k, v])
    ws.append([ir[6], ir[7]])
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 90
    for c in ws["A"]:
        c.font = Font(bold=True)

    out = args.out or os.path.join(work, f"aio-youtube-{days[-1]}.xlsx")
    wb.save(out)
    unclassified = [cid for cid in by_ch if cid not in own and cid not in types]
    print(f"Saved {out}: {len(days)} day(s), {len(recs)} keyword checks, {len(rows)} YouTube "
          f"appearances, {len(by_vid)} videos, {len(by_ch)} channels "
          f"({len(unclassified)} not classified).")


# -------------------------------------------------------------------- main

def main():
    p = argparse.ArgumentParser(description="YouTube in Google AI answers: field check.")
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("init", help="create the audit folder")
    s.add_argument("--work", required=True)
    s.add_argument("--keywords", nargs="*", help="keywords, or a .txt file with one per line")
    s.add_argument("--own", nargs="*", help="your channels: URL, @handle or UC... ID")
    s.add_argument("--language", help="SERP language code, e.g. it")
    s.add_argument("--country", help="SERP country code, e.g. IT")
    s.set_defaults(func=cmd_init)

    s = sub.add_parser("status", help="keywords still to check for a day")
    s.add_argument("--work", required=True)
    s.add_argument("--date", help="YYYY-MM-DD, default today")
    s.set_defaults(func=cmd_status)

    s = sub.add_parser("ingest", help="parse a saved SERP response")
    s.add_argument("--work", required=True)
    s.add_argument("--serp-file", required=True)
    s.add_argument("--format", choices=["auto", "seranking", "generic"], default="auto",
                   help="auto: SE Ranking if recognised, else generic (needs --keyword)")
    s.add_argument("--keyword", help="the keyword of this SERP; required for non-SE Ranking responses")
    s.add_argument("--date", help="YYYY-MM-DD, default the SERP crawl date, else today")
    s.set_defaults(func=cmd_ingest)

    s = sub.add_parser("resolve", help="look up cited videos and channels")
    s.add_argument("--work", required=True)
    s.set_defaults(func=cmd_resolve)

    s = sub.add_parser("channels", help="print cited channels for classification")
    s.add_argument("--work", required=True)
    s.add_argument("--only-unclassified", action="store_true")
    s.set_defaults(func=cmd_channels)

    s = sub.add_parser("build", help="write the Excel report")
    s.add_argument("--work", required=True)
    s.add_argument("--out")
    s.add_argument("--lang", choices=["it", "en"], default="en")
    s.set_defaults(func=cmd_build)

    args = p.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
