#!/usr/bin/env python3
"""
Tapestry Journal Scraper
========================
Downloads all observation assets from tapestryjournal.com, organises them
into a date/child directory structure, renames files, and stamps file system
(and EXIF) creation dates to match each observation's recorded date.

Usage:
    python tapestry_scraper.py -e EMAIL -p PASSWORD -o ./export
    python tapestry_scraper.py -e EMAIL -p PASSWORD -o ./export --list-children
    python tapestry_scraper.py -e EMAIL -p PASSWORD -o ./export --child CHILD_ID
    python tapestry_scraper.py -e EMAIL -p PASSWORD -o ./export --verbose
"""

from __future__ import annotations

import argparse
import ctypes
import ctypes.wintypes
import json
import logging
import mimetypes
import os
import platform
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from tqdm import tqdm

# ── Logging ───────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("tapestry")

# ── Constants ─────────────────────────────────────────────────────────────────

BASE_URL          = "https://tapestryjournal.com"
LOGIN_URL         = f"{BASE_URL}/login"
OBSERVATIONS_PAGE = f"{BASE_URL}/s/observations"

MEDIA_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".webp",
                    ".mp4", ".mov", ".avi", ".mkv", ".m4v",
                    ".pdf", ".mp3", ".m4a", ".heic", ".heif"}

# ── Utilities ─────────────────────────────────────────────────────────────────


def slugify(text: str) -> str:
    """Convert *text* to a filesystem-safe slug."""
    text = str(text).strip().lower()
    text = re.sub(r"[^\w\s-]", "", text)
    text = re.sub(r"[\s_-]+", "_", text)
    return text[:60].strip("_")


def parse_obs_date(obs: dict) -> datetime | None:
    """Extract and parse the observation date from an observation dict."""
    for key in ("createdAt", "observation_time", "created_at", "date",
                "page_added", "scheduledAt", "obs_date"):
        raw = obs.get(key)
        if raw:
            dt = parse_date(str(raw))
            if dt:
                return dt
    return None


def parse_date(raw: str) -> datetime | None:
    """Try to parse a date string in many common formats."""
    if not raw:
        return None
    raw = raw.strip()
    for fmt in (
        "%Y-%m-%dT%H:%M:%S.%fZ",
        "%Y-%m-%dT%H:%M:%SZ",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
        "%B %d, %Y",
        "%d %B %Y",
        "%d %b %Y",
    ):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            pass
    return None


def get_ext(url: str) -> str:
    """Return a file extension for *url*, guessing from path."""
    suffix = Path(urlparse(url).path).suffix.lower()
    if suffix in MEDIA_EXTENSIONS:
        return suffix
    return ""


def set_file_times(path: Path, dt: datetime) -> None:
    """
    Set atime + mtime to *dt*.  On Windows also sets the creation time via
    the Win32 SetFileTime API.
    """
    ts = dt.timestamp()
    os.utime(path, (ts, ts))

    if platform.system() != "Windows":
        return

    GENERIC_WRITE        = 0x40000000
    OPEN_EXISTING        = 3
    FILE_ATTRIBUTE_NORMAL = 0x80
    EPOCH_AS_FILETIME    = 116_444_736_000_000_000
    HNS                  = 10_000_000  # 100-nanosecond intervals per second

    class FILETIME(ctypes.Structure):
        _fields_ = [
            ("dwLowDateTime",  ctypes.wintypes.DWORD),
            ("dwHighDateTime", ctypes.wintypes.DWORD),
        ]

    k32 = ctypes.windll.kernel32  # type: ignore[attr-defined]
    handle = k32.CreateFileW(
        str(path), GENERIC_WRITE, 0, None,
        OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, None,
    )
    if handle == ctypes.wintypes.HANDLE(-1).value:
        log.debug("Could not open %s to set creation time", path)
        return

    ft_val = int(ts * HNS) + EPOCH_AS_FILETIME
    ft = FILETIME(ft_val & 0xFFFFFFFF, ft_val >> 32)
    k32.SetFileTime(handle, ctypes.byref(ft), None, None)
    k32.CloseHandle(handle)


def embed_image_metadata(path: Path, dt: datetime, obs: dict) -> None:
    """
    Write rich metadata into a JPEG or PNG:
      - EXIF: DateTimeOriginal, ImageDescription, Artist, GPS (if available)
      - XMP packet: dc:title, dc:description, dc:subject (keywords), dc:creator
    Google Photos reads both EXIF ImageDescription and XMP dc:description for
    the caption shown in the info panel, and dc:subject for keyword tags.
    """
    ext = path.suffix.lower()
    if ext in (".jpg", ".jpeg"):
        _embed_jpeg_metadata(path, dt, obs)
    # PNG/WEBP: filesystem timestamps are enough; EXIF in PNG is non-standard


def _obs_text_fields(obs: dict) -> tuple[str, str, str, list[str]]:
    """
    Return (title, notes, child_name, keywords) extracted from an obs dict.
    keywords = child name + any learning-goal / tag fields Tapestry may include.
    """
    title = str(obs.get("title") or "").strip()

    # Child name — API v4 returns a "children" list of objects
    child_name = ""
    children_field = obs.get("children")
    if isinstance(children_field, list) and children_field:
        names = [str(c.get("fullName") or c.get("name") or "").strip()
                 for c in children_field if isinstance(c, dict)]
        child_name = ", ".join(n for n in names if n)
    if not child_name:
        _child = obs.get("child")
        child_name = str(
            obs.get("child_name")
            or (_child.get("name", "") if isinstance(_child, dict) else _child or "")
        ).strip()

    # Observation body text — API v4 uses "notes" and "additionalInformation"
    notes = str(
        obs.get("notes") or obs.get("additionalInformation")
        or obs.get("body") or obs.get("text") or obs.get("description") or ""
    ).strip()

    # Keywords: child names + EYFS framework areas / tags
    keywords: list[str] = []
    for name in (child_name.split(", ") if child_name else []):
        if name and name not in keywords:
            keywords.append(name)
    for key in ("frameworks", "tags", "areas", "labels", "goals",
                "framework_tags", "learning_areas"):
        val = obs.get(key)
        if isinstance(val, list):
            for v in val:
                tag = str(v.get("name") or v.get("title") or v
                          if isinstance(v, dict) else v).strip()
                if tag and tag not in keywords:
                    keywords.append(tag)
        elif isinstance(val, str) and val.strip():
            keywords.append(val.strip())

    return title, notes, child_name, keywords


def _gps_rationals(value: float) -> tuple:
    """Convert a decimal degree value to EXIF GPS rational tuple (d, m, s)."""
    value = abs(value)
    d = int(value)
    m = int((value - d) * 60)
    s = round((value - d - m / 60) * 3600, 5)
    # Represent as (numerator, denominator) pairs
    return ((d, 1), (m, 1), (int(s * 10000), 10000))


def _embed_jpeg_metadata(path: Path, dt: datetime, obs: dict) -> None:
    """Write EXIF + XMP into a JPEG file."""
    try:
        import piexif
        from PIL import Image

        title, notes, child_name, keywords = _obs_text_fields(obs)

        # ── 1. EXIF ───────────────────────────────────────────────────────
        img = Image.open(path)
        exif_dict: dict = {"0th": {}, "Exif": {}, "GPS": {}}
        if "exif" in img.info:
            try:
                exif_dict = piexif.load(img.info["exif"])
            except Exception:
                pass

        stamp = dt.strftime("%Y:%m:%d %H:%M:%S").encode()
        exif_dict.setdefault("Exif", {})[piexif.ExifIFD.DateTimeOriginal]  = stamp
        exif_dict["Exif"][piexif.ExifIFD.DateTimeDigitized]                = stamp
        exif_dict.setdefault("0th",  {})[piexif.ImageIFD.DateTime]         = stamp

        # ImageDescription: "Title — notes" (Google Photos shows this)
        description_parts = [p for p in (title, notes) if p]
        if description_parts:
            desc = " — ".join(description_parts)
            exif_dict["0th"][piexif.ImageIFD.ImageDescription] = (
                desc[:2000].encode("ascii", errors="replace")
            )

        if child_name:
            exif_dict["0th"][piexif.ImageIFD.Artist] = (
                child_name.encode("ascii", errors="replace")
            )

        # GPS — Tapestry may expose lat/lng on the obs or obs["location"]
        loc = obs.get("location") or {}
        if isinstance(loc, str):
            parts = loc.split(",")
            if len(parts) == 2:
                try:
                    loc = {"lat": float(parts[0]), "lng": float(parts[1])}
                except ValueError:
                    loc = {}
        lat = obs.get("latitude")  or (loc.get("lat")  if isinstance(loc, dict) else None)
        lng = obs.get("longitude") or (loc.get("lng") or loc.get("lon")
                                        if isinstance(loc, dict) else None)
        if lat is not None and lng is not None:
            try:
                lat, lng = float(lat), float(lng)
                exif_dict.setdefault("GPS", {})[piexif.GPSIFD.GPSLatitudeRef]  = (b"N" if lat >= 0 else b"S")
                exif_dict["GPS"][piexif.GPSIFD.GPSLatitude]                    = _gps_rationals(lat)
                exif_dict["GPS"][piexif.GPSIFD.GPSLongitudeRef]                = (b"E" if lng >= 0 else b"W")
                exif_dict["GPS"][piexif.GPSIFD.GPSLongitude]                   = _gps_rationals(lng)
            except Exception:
                pass

        img.save(path, exif=piexif.dump(exif_dict))

        # ── 2. XMP packet ─────────────────────────────────────────────────
        _inject_xmp_into_jpeg(path, _build_xmp_packet(dt, title, notes,
                                                       child_name, keywords))

        # ── 3. IPTC Caption-Abstract ───────────────────────────────────────
        # Google Photos reads IPTC 2:120 for its editable Description field
        description = " — ".join(p for p in (title, notes) if p)
        _inject_iptc_into_jpeg(path, caption=description, byline=child_name)

    except ImportError:
        pass
    except Exception as exc:
        log.debug("Image metadata write skipped for %s: %s", path.name, exc)


def _build_xmp_packet(dt: datetime, title: str, notes: str,
                      child_name: str, keywords: list[str]) -> str:
    """Return a complete XMP packet string for embedding in a JPEG."""
    from xml.sax.saxutils import escape

    date_str = dt.strftime("%Y-%m-%dT%H:%M:%S")

    def alt(tag: str, value: str) -> str:
        if not value:
            return ""
        return (f"<{tag}><rdf:Alt>"
                f'<rdf:li xml:lang="x-default">{escape(value)}</rdf:li>'
                f"</rdf:Alt></{tag}>")

    def seq(tag: str, values: list[str]) -> str:
        if not values:
            return ""
        items = "".join(f"<rdf:li>{escape(v)}</rdf:li>" for v in values)
        return f"<{tag}><rdf:Seq>{items}</rdf:Seq></{tag}>"

    def bag(tag: str, values: list[str]) -> str:
        if not values:
            return ""
        items = "".join(f"<rdf:li>{escape(v)}</rdf:li>" for v in values)
        return f"<{tag}><rdf:Bag>{items}</rdf:Bag></{tag}>"

    description = " — ".join(p for p in (title, notes) if p)

    return (
        '<?xpacket begin="\xef\xbb\xbf" id="W5M0MpCehiHzreSzNTczkc9d"?>'
        '<x:xmpmeta xmlns:x="adobe:ns:meta/">'
        '<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">'
        '<rdf:Description rdf:about="" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:xmp="http://ns.adobe.com/xap/1.0/">'
        + alt("dc:title", title)
        + alt("dc:description", description)
        + seq("dc:creator", [child_name] if child_name else [])
        + bag("dc:subject", keywords)
        + f"<xmp:CreateDate>{date_str}</xmp:CreateDate>"
        + "</rdf:Description>"
        "</rdf:RDF>"
        "</x:xmpmeta>"
        '<?xpacket end="w"?>'
    )


def _inject_iptc_into_jpeg(path: Path, caption: str, byline: str = "") -> None:
    """
    Inject IPTC IIM records into a JPEG APP13 (Photoshop 3.0) segment.
    Google Photos reads IPTC Caption-Abstract (2:120) as the photo description
    and By-line (2:80) as the photographer/creator.
    """
    import struct

    def iptc_record(dataset: int, value: bytes) -> bytes:
        return bytes([0x1C, 0x02, dataset]) + struct.pack(">H", len(value)) + value

    iptc_data = b""
    if caption:
        iptc_data += iptc_record(120, caption.encode("utf-8"))  # Caption-Abstract
    if byline:
        iptc_data += iptc_record(80, byline.encode("utf-8"))    # By-line
    if not iptc_data:
        return

    # Pad IPTC block to even length
    if len(iptc_data) % 2:
        iptc_data += b"\x00"

    # Wrap in Photoshop 3.0 / 8BIM resource block
    resource = (
        b"8BIM"
        + b"\x04\x04"               # Resource ID: IPTC-NAA
        + b"\x00\x00"               # Pascal string name (empty, even-padded)
        + struct.pack(">I", len(iptc_data))
        + iptc_data
    )
    app13_payload = b"Photoshop 3.0\x00" + resource
    app13 = b"\xff\xed" + struct.pack(">H", len(app13_payload) + 2) + app13_payload

    data = path.read_bytes()
    if len(data) < 2 or data[:2] != b"\xff\xd8":
        return

    # Strip any existing APP13 segments, keep everything else
    out = bytearray(data[:2])  # SOI
    i = 2
    while i < len(data):
        if i + 2 > len(data):
            out += data[i:]
            break
        if data[i] != 0xFF:
            out += data[i:]
            break
        marker = data[i: i + 2]
        if marker in (b"\xff\xd8", b"\xff\xd9", b"\xff\x01") or (
            0xD0 <= data[i + 1] <= 0xD7
        ):
            out += marker
            i += 2
            continue
        if i + 4 > len(data):
            out += data[i:]
            break
        seg_len = struct.unpack(">H", data[i + 2: i + 4])[0]
        seg_end = i + 2 + seg_len
        if marker != b"\xff\xed":   # keep everything except APP13
            out += data[i: seg_end]
        i = seg_end

    # Insert new APP13 right after SOI
    path.write_bytes(bytes(out[:2]) + app13 + bytes(out[2:]))


def _inject_xmp_into_jpeg(path: Path, xmp_str: str) -> None:
    """
    Insert (or replace) an XMP APP1 segment in a JPEG file.
    XMP APP1 = FF E1 <length> <namespace_uri>\0 <xmp_xml>
    """
    import struct

    XMP_NS = b"http://ns.adobe.com/xap/1.0/\x00"

    data = path.read_bytes()
    if len(data) < 2 or data[:2] != b"\xff\xd8":
        return  # Not a valid JPEG

    # ── Strip any existing XMP APP1 segments ──────────────────────────────
    out = bytearray(data[:2])  # SOI
    i = 2
    while i < len(data):
        if i + 2 > len(data):
            out += data[i:]
            break
        if data[i] != 0xFF:
            out += data[i:]
            break
        marker = data[i: i + 2]
        # Markers without a length field
        if marker in (b"\xff\xd8", b"\xff\xd9", b"\xff\x01") or (
            0xD0 <= data[i + 1] <= 0xD7
        ):
            out += marker
            i += 2
            continue
        if i + 4 > len(data):
            out += data[i:]
            break
        seg_len = struct.unpack(">H", data[i + 2: i + 4])[0]
        seg_end = i + 2 + seg_len
        seg = data[i: seg_end]
        # Drop existing XMP APP1 (keep EXIF APP1 and everything else)
        is_xmp = (marker == b"\xff\xe1"
                  and len(seg) > 4 + len(XMP_NS)
                  and seg[4: 4 + len(XMP_NS)] == XMP_NS)
        if not is_xmp:
            out += seg
        i = seg_end

    # ── Build and insert new XMP APP1 right after SOI ─────────────────────
    xmp_payload = XMP_NS + xmp_str.encode("utf-8")
    app1 = (b"\xff\xe1"
            + struct.pack(">H", len(xmp_payload) + 2)
            + xmp_payload)
    path.write_bytes(bytes(out[:2]) + app1 + bytes(out[2:]))


def _patch_mp4_mvhd_time(path: Path, dt: datetime) -> None:
    """
    Directly patch the creation_time and modification_time fields in the mvhd
    (Movie Header) box of an MP4 file.  Google Photos reads this binary field
    rather than the ©day text tag.

    Timestamps in mvhd are seconds since 1904-01-01 00:00:00 UTC.
    """
    import struct

    # Seconds between 1904-01-01 and 1970-01-01
    MAC_EPOCH_OFFSET = 2082844800

    ts = int(dt.timestamp()) + MAC_EPOCH_OFFSET

    data = bytearray(path.read_bytes())

    def find_box(buf: bytearray, start: int, end: int, name: bytes) -> int:
        """Return the offset of box *name* within buf[start:end], or -1."""
        i = start
        while i + 8 <= end:
            box_size = struct.unpack_from(">I", buf, i)[0]
            box_name = buf[i + 4: i + 8]
            if box_size < 8:
                break
            if box_name == name:
                return i
            i += box_size
        return -1

    # Locate moov box at top level
    moov_off = find_box(data, 0, len(data), b"moov")
    if moov_off < 0:
        log.debug("mvhd patch: moov box not found in %s", path.name)
        return
    moov_size = struct.unpack_from(">I", data, moov_off)[0]

    # Locate mvhd inside moov
    mvhd_off = find_box(data, moov_off + 8, moov_off + moov_size, b"mvhd")
    if mvhd_off < 0:
        log.debug("mvhd patch: mvhd box not found in %s", path.name)
        return

    # mvhd layout: size(4) + "mvhd"(4) + version(1) + flags(3) + timestamps…
    version = data[mvhd_off + 8]
    if version == 0:
        # 32-bit timestamps at offsets +12 (creation) and +16 (modification)
        struct.pack_into(">I", data, mvhd_off + 12, ts & 0xFFFFFFFF)
        struct.pack_into(">I", data, mvhd_off + 16, ts & 0xFFFFFFFF)
    elif version == 1:
        # 64-bit timestamps at offsets +12 (creation) and +20 (modification)
        struct.pack_into(">Q", data, mvhd_off + 12, ts)
        struct.pack_into(">Q", data, mvhd_off + 20, ts)
    else:
        log.debug("mvhd patch: unknown mvhd version %d in %s", version, path.name)
        return

    path.write_bytes(bytes(data))
    log.debug("mvhd patch: stamped %s with %s", path.name, dt.isoformat())


def _inject_xmp_into_mp4(path: Path, xmp_str: str) -> None:
    """
    Embed (or replace) an XMP packet in an MP4/MOV file as a uuid box.
    Adobe XMP UUID: BE7ACFCB-97A9-42E8-9C71-999491E3AFAC
    Google Photos reads dc:description from this for the file description.
    """
    import struct

    XMP_UUID = bytes.fromhex("BE7ACFCB97A942E89C71999491E3AFAC")

    data = bytearray(path.read_bytes())

    # Walk box list and strip any existing XMP uuid box
    out = bytearray()
    i = 0
    while i < len(data):
        if i + 8 > len(data):
            out += data[i:]
            break
        box_size = struct.unpack_from(">I", data, i)[0]
        box_name = bytes(data[i + 4: i + 8])
        if box_size < 8:
            out += data[i:]
            break
        is_xmp_uuid = (
            box_name == b"uuid"
            and i + 24 <= len(data)
            and bytes(data[i + 8: i + 24]) == XMP_UUID
        )
        if is_xmp_uuid:
            i += box_size
            continue
        out += data[i: i + box_size]
        i += box_size

    # Append new uuid box
    xmp_bytes = xmp_str.encode("utf-8")
    box_content = b"uuid" + XMP_UUID + xmp_bytes
    new_box = struct.pack(">I", len(box_content) + 4) + box_content
    out += new_box

    path.write_bytes(bytes(out))
    log.debug("XMP uuid box written to %s", path.name)


def embed_video_metadata(path: Path, dt: datetime, obs: dict) -> None:
    """
    Write metadata tags into MP4/M4V/MOV files via mutagen, then patch the
    mvhd box creation_time so Google Photos picks up the correct date, and
    embed XMP so Google Photos shows the description.
    """
    title, notes, child_name, keywords = _obs_text_fields(obs)

    try:
        from mutagen.mp4 import MP4, MP4Tags

        audio = MP4(str(path))
        audio.tags = audio.tags or MP4Tags()  # type: ignore[assignment]

        # ISO date string for ©day
        audio["©day"] = [dt.strftime("%Y-%m-%dT%H:%M:%S")]
        if title:
            audio["©nam"] = [title]
        description = " — ".join(p for p in (title, notes) if p)
        if description:
            audio["©cmt"] = [description[:4000]]
        if child_name:
            audio["©ART"] = [child_name]   # artist field

        audio.save()
    except ImportError:
        pass
    except Exception as exc:
        log.debug("Video metadata write skipped for %s: %s", path.name, exc)

    # Patch mvhd binary timestamps — this is what Google Photos actually reads
    try:
        _patch_mp4_mvhd_time(path, dt)
    except Exception as exc:
        log.debug("mvhd patch failed for %s: %s", path.name, exc)

    # Embed XMP uuid box — Google Photos reads dc:description from here
    try:
        _inject_xmp_into_mp4(path, _build_xmp_packet(dt, title, notes,
                                                      child_name, keywords))
    except Exception as exc:
        log.debug("XMP inject failed for %s: %s", path.name, exc)


# ── Session & login ───────────────────────────────────────────────────────────


class TapestrySession:
    """Authenticated requests session for tapestryjournal.com."""

    def __init__(self) -> None:
        self.s = requests.Session()
        self.s.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
        })
        self._csrf: str = ""
        self._school_slug: str = ""

    @property
    def _obs_url(self) -> str:
        """Base URL for the observations section, including school slug."""
        if self._school_slug:
            return f"{BASE_URL}/s/{self._school_slug}/observations"
        return OBSERVATIONS_PAGE

    @property
    def _api4_headers(self) -> dict:
        """Headers required by the React /api/4/ endpoints."""
        return {
            "X-Requested-With": "XMLHttpRequest",
            "X-CSRF-TOKEN": self._csrf,
            "X-TAPESTRY-VERSION": "3",
            "Accept": "application/json",
        }

    # ── Session cache ─────────────────────────────────────────────────────────
    #
    # Logging in from scratch on every scheduled run is noisy for the account
    # holder, so the authenticated cookies (plus the CSRF token and school slug
    # the API needs) can be persisted and reused until the server rejects them.

    def load_session(self, path: Path) -> bool:
        """Restore a cached session from *path*. Returns True only if the
        server still accepts it."""
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except FileNotFoundError:
            return False
        except Exception as exc:
            log.debug("Ignoring unreadable session cache %s: %s", path, exc)
            return False

        for c in data.get("cookies", []):
            self.s.cookies.set(c["name"], c["value"],
                               domain=c.get("domain", ""), path=c.get("path", "/"))
        self._csrf = data.get("csrf", "")
        self._school_slug = data.get("school_slug", "")

        try:
            r = self.s.get(f"{BASE_URL}/api/4/children/list",
                           headers=self._api4_headers, timeout=20,
                           allow_redirects=False)
            ok = (r.status_code == 200
                  and "json" in r.headers.get("Content-Type", ""))
        except requests.RequestException as exc:
            log.debug("Session cache check failed: %s", exc)
            ok = False

        if ok:
            log.info("Reusing cached session from %s", path)
            return True

        log.info("Cached session no longer accepted — logging in again")
        self.s.cookies.clear()
        self._csrf = ""
        self._school_slug = ""
        return False

    def save_session(self, path: Path) -> None:
        """Write the current session to *path*, readable by the owner only."""
        data = {
            "cookies": [
                {"name": c.name, "value": c.value,
                 "domain": c.domain, "path": c.path}
                for c in self.s.cookies
            ],
            "csrf": self._csrf,
            "school_slug": self._school_slug,
        }
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".tmp")
        fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f)
        os.replace(tmp, path)

    # ── Internal helpers ──────────────────────────────────────────────────────

    def _soup(self, url: str, **kwargs) -> tuple[BeautifulSoup, requests.Response]:
        r = self.s.get(url, timeout=30, **kwargs)
        r.raise_for_status()
        return BeautifulSoup(r.text, "lxml"), r

    def _extract_csrf(self, html: str) -> str:
        """Pull CSRF token from a page's HTML."""
        # 1. JSON config blob embedded by Tapestry
        m = re.search(r'"csrf_token"\s*:\s*"([^"]+)"', html)
        if m:
            return m.group(1)
        # 2. <meta name="csrf-token" content="...">
        soup = BeautifulSoup(html, "lxml")
        meta = soup.find("meta", attrs={"name": "csrf-token"})
        if meta and meta.get("content"):
            return meta["content"]
        # 3. Hidden _token input
        inp = soup.find("input", attrs={"name": "_token"})
        if inp and inp.get("value"):
            return inp["value"]
        return ""

    def _json_from_page(self, html: str) -> dict | list | None:
        """
        Tapestry embeds page data as JSON in a few patterns:
          window.__INITIAL_STATE__ = {...}
          var pageData = {...}
          <script id="page-data" type="application/json">...</script>
        Returns the parsed object, or None.
        """
        # Inline <script type="application/json">
        soup = BeautifulSoup(html, "lxml")
        for tag in soup.find_all("script",
                                  attrs={"type": "application/json"}):
            try:
                return json.loads(tag.string)
            except Exception:
                pass

        # window.__X__ = {...}; or var x = {...};
        for pattern in (
            r'window\.__[A-Z_]+__\s*=\s*(\{[\s\S]+?\});\s*(?:\n|$)',
            r'var\s+(?:pageData|initialData|appData)\s*=\s*(\{[\s\S]+?\});\s*(?:\n|$)',
        ):
            m = re.search(pattern, html)
            if m:
                try:
                    return json.loads(m.group(1))
                except Exception:
                    pass
        return None

    # ── Login ─────────────────────────────────────────────────────────────────

    def login(self, email: str, password: str) -> None:
        log.info("Fetching login page …")
        soup, r = self._soup(LOGIN_URL)
        csrf = self._extract_csrf(r.text)
        if not csrf:
            raise RuntimeError("Could not find CSRF token on login page.")
        self._csrf = csrf

        # ── Read the actual form fields from the HTML ──────────────────────
        # Rather than hardcoding field names, find the login form and use
        # whatever input names it declares.
        form = soup.find("form")
        payload: dict = {}
        if form:
            for inp in form.find_all("input"):
                name  = inp.get("name", "")
                value = inp.get("value", "")
                if not name:
                    continue
                if name == "_token":
                    payload[name] = csrf
                elif inp.get("type", "").lower() == "email" or name in ("email", "username", "login"):
                    payload[name] = email
                elif inp.get("type", "").lower() == "password" or name == "password":
                    payload[name] = password
                elif inp.get("type", "").lower() in ("hidden", "checkbox", "radio"):
                    payload[name] = value  # carry through hidden/default values
        # Fallback to known field names if form parsing didn't find them
        payload.setdefault("_token",   csrf)
        payload.setdefault("email",    email)
        payload.setdefault("password", password)
        payload.setdefault("remember", "1")

        log.debug("Login payload fields: %s", list(payload.keys()))

        # ── POST credentials ───────────────────────────────────────────────
        # Use allow_redirects=False so we can inspect the direct response
        # before following the chain — a 302 away from /login = success.
        resp = self.s.post(
            LOGIN_URL,
            data=payload,
            headers={"Referer": LOGIN_URL,
                     "Content-Type": "application/x-www-form-urlencoded"},
            allow_redirects=False,
            timeout=30,
        )
        log.debug("Login POST status: %s  Location: %s",
                  resp.status_code, resp.headers.get("Location", "—"))

        # Follow redirects manually so we can detect where we end up
        final = self.s.get(
            urljoin(BASE_URL, resp.headers.get("Location", "/")),
            allow_redirects=True,
            timeout=30,
        ) if resp.status_code in (301, 302, 303, 307, 308) else resp

        log.debug("Final URL after login: %s", final.url)

        # ── Failure detection ──────────────────────────────────────────────
        still_on_login = final.url.rstrip("/") == LOGIN_URL.rstrip("/")
        has_error_text = re.search(r"\bincorrect\b|\bthese credentials\b",
                                   final.text, re.I)
        if still_on_login or has_error_text:
            # Find a human-readable error message (avoid the giant JS blob)
            err_soup = BeautifulSoup(final.text, "lxml")
            msg = ""
            for sel in (".alert", "[role=alert]", ".invalid-feedback",
                        ".text-danger", ".error-message", "p.error"):
                el = err_soup.select_one(sel)
                if el:
                    msg = el.get_text(" ", strip=True)
                    if len(msg) < 300:   # ignore if it's the JS blob
                        break
                    msg = ""
            raise RuntimeError(
                f"Login failed — still on {final.url}\n"
                + (f"  Site message: {msg}\n" if msg else "")
                + "  • Double-check your email and password.\n"
                + "  • If your password contains special characters (!$&etc.)\n"
                + "    wrap it in single quotes:  -p 'your!password'\n"
                + "  • Run with --verbose to see full debug output."
            )

        # Extract school slug from the redirect URL
        # e.g. https://tapestryjournal.com/s/cherry-garden-framework-test/observations
        m = re.search(r"/s/([^/]+)/", final.url)
        if m:
            self._school_slug = m.group(1)
            log.info("School slug: %s", self._school_slug)

        # The React app embeds the CSRF token in a hidden div as part of its
        # initial state: <div class="hidden">{..."csrfToken":"..."}...</div>
        # This is the token that every API/4 request must carry.
        soup_final = BeautifulSoup(final.text, "lxml")
        for div in soup_final.find_all("div", class_="hidden"):
            txt = div.get_text().strip()
            if txt.startswith("{") and "csrfToken" in txt:
                try:
                    cfg = json.loads(txt)
                    tok = cfg.get("csrfToken") or cfg.get("csrf_token")
                    if tok:
                        self._csrf = tok
                        log.debug("CSRF token extracted from React config div")
                        break
                except Exception:
                    pass

        # Fallback: look in meta / JS config blob
        if not self._csrf:
            new_csrf = self._extract_csrf(final.text)
            if new_csrf:
                self._csrf = new_csrf

        log.info("Logged in successfully. Session URL: %s", final.url)

    # ── Children ──────────────────────────────────────────────────────────────

    def get_children(self) -> list[dict]:
        """Return list of children visible to this account."""
        try:
            r = self.s.get(f"{BASE_URL}/api/4/children/list",
                           headers=self._api4_headers, timeout=20)
            r.raise_for_status()
            data = r.json()
            # Response is a list of child objects
            kids = data if isinstance(data, list) else data.get("children", data.get("data", []))
            return kids if isinstance(kids, list) else []
        except Exception as exc:
            log.debug("Children API failed: %s", exc)
            return []

    # ── Observation list ──────────────────────────────────────────────────────

    def get_observations(self, child_id: str | None = None) -> list[dict]:
        """
        Return all observations as a list of dicts.
        Tries JSON API first; falls back to HTML page scraping.
        """
        obs = self._try_json_api(child_id)
        if obs is not None:
            self._backfill_truncated_media(obs)
            return obs
        return self._scrape_observation_list(child_id)

    def _try_json_api(self, child_id: str | None) -> list[dict] | None:
        """
        Fetch observations from the React app's /api/4/observations/list endpoint.
        Uses cursor-based pagination (nextCursor field in each response page).
        """
        url = f"{BASE_URL}/api/4/observations/list"
        all_obs: list[dict] = []
        cursor = None
        page = 0

        while True:
            params: dict = {"perPage": 50}
            if child_id:
                params["children.child_id"] = child_id
            if cursor:
                params["cursor"] = cursor

            try:
                r = self.s.get(url, params=params, headers=self._api4_headers,
                               timeout=30, allow_redirects=False)
                log.debug("API /api/4/observations/list page %d → %s",
                          page + 1, r.status_code)

                if r.status_code == 401:
                    log.warning("API returned 401 — session may have expired")
                    return None
                if r.status_code != 200:
                    return None
                if "json" not in r.headers.get("Content-Type", ""):
                    return None

                data = r.json()
            except Exception as exc:
                log.debug("API fetch error: %s", exc)
                return None

            # Observations are in data["observations"] or the top-level list
            obs_list = (data.get("observations")
                        or data.get("data")
                        or (data if isinstance(data, list) else []))

            if not obs_list:
                break

            all_obs.extend(obs_list)
            page += 1
            log.info("  Page %d — %d observations so far", page, len(all_obs))

            cursor = data.get("nextCursor")
            if not cursor:
                break

            time.sleep(0.25)

        log.info("API returned %d observations total", len(all_obs))
        return all_obs if all_obs else None

    def _backfill_truncated_media(self, observations: list[dict]) -> None:
        """
        The /api/4/observations/list endpoint only returns a *preview* of each
        observation's media — at most 3 items — while the true total is given by
        the "mediaCount" field.  For any observation whose preview is truncated,
        fetch the full detail (/api/4/observations/get/{id}) and replace its
        media (and documents) with the complete lists, so every asset downloads.
        """
        truncated = [
            o for o in observations
            if isinstance(o.get("mediaCount"), int)
            and o["mediaCount"] > len(o.get("media") or [])
        ]
        if not truncated:
            return

        log.info(
            "%d observation(s) have more media than the list preview returned "
            "(preview caps at 3) — fetching full media lists …",
            len(truncated),
        )
        for o in tqdm(truncated, desc="Fetching full media lists"):
            obs_id = o.get("id")
            if obs_id is None:
                continue
            url = f"{BASE_URL}/api/4/observations/get/{obs_id}"
            try:
                r = self.s.get(url, headers=self._api4_headers, timeout=30)
                r.raise_for_status()
                detail = r.json()
            except Exception as exc:
                log.warning(
                    "Could not fetch full media for observation %s "
                    "(preview had %d of %s): %s",
                    obs_id, len(o.get("media") or []), o.get("mediaCount"), exc,
                )
                continue

            if not isinstance(detail, dict):
                continue
            before = len(o.get("media") or [])
            if isinstance(detail.get("media"), list):
                o["media"] = detail["media"]
            if isinstance(detail.get("documents"), list):
                o["documents"] = detail["documents"]
            log.debug(
                "Observation %s: media %d → %d (mediaCount=%s)",
                obs_id, before, len(o.get("media") or []), o.get("mediaCount"),
            )
            time.sleep(0.15)

    # ── HTML scraping ─────────────────────────────────────────────────────────

    def _scrape_observation_list(self,
                                  child_id: str | None = None) -> list[dict]:
        """
        Scrape the observations listing page to collect observation IDs/URLs,
        then visit each detail page to extract full metadata and media.
        """
        log.info("Scraping observation list from HTML …")
        links: list[dict] = []
        page = 1

        while True:
            params: dict = {"page": page}
            if child_id:
                params["child_id"] = child_id

            soup, r = self._soup(self._obs_url, params=params)

            # Check if we've been redirected to login
            if "login" in r.url:
                raise RuntimeError(
                    "Session expired or login failed — redirected to login page."
                )

            # On the first page dump the HTML for diagnosis (debug mode only)
            if page == 1 and log.isEnabledFor(logging.DEBUG):
                dump = Path("debug_observations_page.html")
                dump.write_text(r.text, encoding="utf-8")
                log.debug("Saved observations page HTML → %s", dump.resolve())
                # Also show any script tags that hint at page data / API URLs
                for script in soup.find_all("script", src=False)[:5]:
                    txt = (script.string or "")[:200].strip()
                    if txt:
                        log.debug("Inline script snippet: %s …", txt[:120])

            # ── Extract observation links ──────────────────────────────────
            # Tapestry renders observation cards with links like
            # /s/<school-slug>/observations/<id>
            obs_path_re = re.compile(
                r"/s/[^/]+/observations/(\d+)"
                if self._school_slug else
                r"/s/observations/(\d+)"
            )
            found: list[dict] = []
            for a in soup.select("a[href]"):
                href = a["href"]
                m = obs_path_re.search(href)
                if m:
                    obs_id = m.group(1)
                    full_url = urljoin(BASE_URL, href)
                    if not any(x["id"] == obs_id for x in found):
                        found.append({"id": obs_id, "_url": full_url})

            if not found:
                # Try looking inside JSON embedded in the page
                page_data = self._json_from_page(r.text)
                if page_data:
                    obs_list = (page_data if isinstance(page_data, list)
                                else page_data.get("observations", []))
                    for item in obs_list:
                        if isinstance(item, dict) and "id" in item:
                            found.append(item)

            if not found:
                break

            links.extend(found)
            log.info("  List page %d — %d observations found so far",
                     page, len(links))

            # Pagination: look for a "next page" link
            next_a = soup.select_one(
                "a[rel='next'], .pagination .next a, "
                "a.next-page, li.next a"
            )
            if not next_a:
                break

            page += 1
            time.sleep(0.3)

        # Deduplicate by id
        seen: set[str] = set()
        unique: list[dict] = []
        for obs in links:
            if obs["id"] not in seen:
                seen.add(obs["id"])
                unique.append(obs)

        log.info("Found %d unique observation links; fetching details …",
                 len(unique))

        detailed: list[dict] = []
        for obs in tqdm(unique, desc="Fetching observation details"):
            detailed.append(self._fetch_observation_detail(obs))
            time.sleep(0.2)

        return detailed

    def _fetch_observation_detail(self, obs: dict) -> dict:
        """
        Fetch a single observation detail page and enrich *obs* with:
          - observation_time  (ISO-like string)
          - child_name
          - title
          - media             (list of {"url": ..., "type": ...})
        """
        url = obs.get("_url", f"{self._obs_url}/{obs['id']}")
        try:
            soup, r = self._soup(url)
        except Exception as exc:
            log.warning("Could not fetch observation %s: %s", obs["id"], exc)
            return obs

        # ── 1. Try embedded JSON first ─────────────────────────────────────
        page_data = self._json_from_page(r.text)
        if isinstance(page_data, dict):
            obs_data = (page_data.get("observation")
                        or page_data.get("entry")
                        or page_data)
            if isinstance(obs_data, dict):
                obs.update(obs_data)
                # Normalise media list
                obs.setdefault("media", self._collect_media_from_dict(obs_data))
                return obs

        # ── 2. HTML parsing ────────────────────────────────────────────────

        # Date ─────────────────────────────────────────────────────────────
        for sel in (
            "time[datetime]",
            "[data-observation-date]",
            "[data-date]",
            ".observation-date",
            ".obs-date",
            ".entry-date",
            ".date",
        ):
            el = soup.select_one(sel)
            if el:
                raw = (el.get("datetime")
                       or el.get("data-observation-date")
                       or el.get("data-date")
                       or el.get_text(strip=True))
                if raw:
                    obs.setdefault("observation_time", raw)
                    break

        # If still no date, try parsing text like "12 January 2024"
        if "observation_time" not in obs:
            date_pattern = re.compile(
                r"\b(\d{1,2}\s+\w+\s+\d{4}|\d{4}-\d{2}-\d{2})\b"
            )
            for el in soup.select("p, span, div, h2, h3"):
                m = date_pattern.search(el.get_text())
                if m:
                    obs["observation_time"] = m.group(1)
                    break

        # Child name ────────────────────────────────────────────────────────
        for sel in (
            "[data-child-name]",
            ".child-name",
            ".child_name",
            ".childname",
            ".profile-name",
        ):
            el = soup.select_one(sel)
            if el:
                obs.setdefault(
                    "child_name",
                    el.get("data-child-name") or el.get_text(strip=True),
                )
                break

        # Title / heading ──────────────────────────────────────────────────
        for sel in ("h1", ".observation-title", ".entry-title", ".obs-title"):
            el = soup.select_one(sel)
            if el:
                obs.setdefault("title", el.get_text(strip=True))
                break

        # ── 3. Collect media URLs ──────────────────────────────────────────
        media: list[dict] = []

        # Images
        for img in soup.select("img"):
            src = img.get("data-src") or img.get("data-original") or img.get("src", "")
            if src and _looks_like_asset(src):
                media.append({"url": src, "type": "image"})

        # Videos – <video>, <source>, or direct <a> links
        for tag in soup.select("video[src], video source[src], source[src]"):
            src = tag.get("src", "")
            if src and _looks_like_asset(src):
                media.append({"url": src, "type": "video"})

        for a in soup.select("a[href]"):
            href = a["href"]
            if _looks_like_asset(href):
                ext = get_ext(href)
                media.append({
                    "url": href,
                    "type": "video" if ext in {".mp4", ".mov", ".avi",
                                                ".mkv", ".m4v"} else "file",
                })

        # Deduplicate media by URL
        seen_urls: set[str] = set()
        unique_media: list[dict] = []
        for item in media:
            u = item["url"]
            if u not in seen_urls:
                seen_urls.add(u)
                unique_media.append(item)

        obs["media"] = unique_media
        return obs

    @staticmethod
    def _collect_media_from_dict(data: dict) -> list[dict]:
        """Extract media items from a JSON observation dict."""
        media: list[dict] = []
        for key in ("media", "assets", "files", "images", "videos",
                    "attachments"):
            items = data.get(key)
            if not isinstance(items, list):
                continue
            for item in items:
                if isinstance(item, str) and _looks_like_asset(item):
                    media.append({"url": item})
                elif isinstance(item, dict):
                    url = (item.get("original_url") or item.get("url")
                           or item.get("src") or item.get("path") or "")
                    if url and _looks_like_asset(url):
                        media.append(item)
        return media

    # ── Downloading ───────────────────────────────────────────────────────────

    def download(self, url: str, dest: Path) -> bool:
        """Download *url* → *dest*. Skips if file already exists. Returns True on success."""
        if dest.exists() and dest.stat().st_size > 0:
            log.debug("Already exists: %s", dest.name)
            return True
        dest.parent.mkdir(parents=True, exist_ok=True)
        try:
            with self.s.get(url, stream=True, timeout=120) as r:
                r.raise_for_status()
                # Use Content-Disposition filename if no extension yet
                tmp = dest.with_suffix(".part")
                with open(tmp, "wb") as f:
                    for chunk in r.iter_content(65536):
                        if chunk:
                            f.write(chunk)
            tmp.rename(dest)
            return True
        except Exception as exc:
            log.warning("Download failed %s → %s: %s", url, dest.name, exc)
            if dest.with_suffix(".part").exists():
                dest.with_suffix(".part").unlink(missing_ok=True)
            return False


# ── Helper ────────────────────────────────────────────────────────────────────


def _looks_like_asset(url: str) -> bool:
    """Return True if the URL looks like a media/document asset."""
    if not url or url.startswith("data:"):
        return False
    path = urlparse(url).path.lower()
    # Must have a recognisable media extension OR be on a CDN/storage path
    ext = Path(path).suffix
    if ext in MEDIA_EXTENSIONS:
        return True
    # Cloudinary / S3 / similar CDN paths that may lack extensions
    cdn_patterns = (
        "cloudinary.com", "s3.amazonaws.com", "storage.googleapis.com",
        "/uploads/", "/media/", "/assets/", "/files/",
    )
    return any(p in url for p in cdn_patterns)


# ── Organiser ─────────────────────────────────────────────────────────────────


def _remove_legacy_media_files(
    obs_dir: Path, date_prefix: str, keep: set[str]
) -> int:
    """
    Delete media files left by the earlier positional-index naming scheme
    ("<date>_001.jpg", "<date>_002.mp4", …) that are not part of the current
    id-named set *keep*.  Lets folders written by older runs migrate to stable
    id filenames without leaving duplicated assets behind.  Returns the number
    of files removed.
    """
    legacy_re = re.compile(rf"^{re.escape(date_prefix)}_\d{{3}}\.[A-Za-z0-9]+$")
    removed = 0
    for f in obs_dir.iterdir():
        if not f.is_file() or f.name in keep or not legacy_re.match(f.name):
            continue
        try:
            f.unlink()
            removed += 1
            log.debug("Migrated: removed legacy-named file %s", f.name)
        except OSError as exc:
            log.debug("Could not remove %s: %s", f.name, exc)
    return removed


def organise(
    observations: list[dict],
    output_dir: Path,
    session: TapestrySession,
) -> None:
    """
    Directory layout:
        <output>/<child_name>/<YYYY-MM-DD_<title>>/
            observation.json
            YYYY-MM-DD_<mediaId>.jpg
            YYYY-MM-DD_<mediaId>.mp4
            …
    Media files are named after each asset's stable Tapestry id, so every run
    maps an asset to the same filename and only fetches what is missing.
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    n_ok = n_skip = n_fail = n_migrated = 0

    for obs in tqdm(observations, desc="Organising observations"):
        obs_id = str(obs.get("id", "unknown"))

        # ── Date ──────────────────────────────────────────────────────────
        obs_dt = parse_obs_date(obs)
        if obs_dt is None:
            log.warning("Observation %s — no parseable date, skipping. Keys: %s",
                        obs_id, list(obs.keys()))
            n_skip += 1
            continue

        date_prefix = obs_dt.strftime("%Y-%m-%d")

        # ── Child ─────────────────────────────────────────────────────────
        # API v4 returns a "children" list of objects with fullName
        children_list = obs.get("children")
        if isinstance(children_list, list) and children_list:
            child_raw = (children_list[0].get("fullName")
                         or children_list[0].get("name") or "")
        else:
            _child = obs.get("child")
            child_raw = (obs.get("child_name")
                         or (_child.get("name", "") if isinstance(_child, dict)
                             else _child or ""))
        child_slug = slugify(str(child_raw)) or "unknown_child"

        # ── Folder ────────────────────────────────────────────────────────
        title_raw = obs.get("title") or f"obs_{obs_id}"
        title_slug = slugify(str(title_raw))
        folder_name = f"{date_prefix}_{title_slug}" if title_slug else date_prefix

        obs_dir = output_dir / child_slug / folder_name
        obs_dir.mkdir(parents=True, exist_ok=True)

        # ── Metadata JSON ─────────────────────────────────────────────────
        meta = obs_dir / "observation.json"
        if not meta.exists():
            meta.write_text(
                json.dumps(obs, indent=2, ensure_ascii=False, default=str),
                encoding="utf-8",
            )
        set_file_times(meta, obs_dt)

        # ── Media ─────────────────────────────────────────────────────────
        # API v4: "media" list (photos/videos) + "documents" list
        media_items: list[dict] = []
        for item in obs.get("media", []):
            if isinstance(item, dict) and item.get("url"):
                media_items.append(item)
        for item in obs.get("documents", []):
            if isinstance(item, dict) and item.get("url"):
                media_items.append(item)
        # Legacy / fallback keys
        for key in ("assets", "files", "images", "videos", "attachments"):
            for item in obs.get(key, []):
                if isinstance(item, str):
                    media_items.append({"url": item})
                elif isinstance(item, dict) and item.get("url"):
                    media_items.append(item)

        written: set[str] = set()
        obs_failed = False
        for idx, item in enumerate(media_items, start=1):
            url = (item.get("original_url") or item.get("url")
                   or item.get("src") or "")
            if not url:
                continue
            if not url.startswith("http"):
                url = urljoin(BASE_URL, url)

            ext = get_ext(url)
            if not ext:
                item_type = str(item.get("type", "")).lower()
                ext = ".jpg" if item_type == "image" else (
                      ".mp4" if item_type == "video" else ".bin")

            # Name each file after the asset's stable Tapestry id so it always
            # maps to the same filename regardless of position or sibling count.
            # Fall back to a positional index only for legacy items with no id.
            media_id = item.get("id")
            stem = (f"{date_prefix}_{media_id}" if media_id is not None
                    else f"{date_prefix}_{idx:03d}")
            filename = f"{stem}{ext}"
            if filename in written:      # two assets sharing an id + ext (rare)
                filename = f"{stem}_{idx:03d}{ext}"
            written.add(filename)
            dest = obs_dir / filename

            if session.download(url, dest):
                set_file_times(dest, obs_dt)
                if ext in (".jpg", ".jpeg"):
                    embed_image_metadata(dest, obs_dt, obs)
                elif ext in (".mp4", ".m4v", ".mov"):
                    embed_video_metadata(dest, obs_dt, obs)
                n_ok += 1
            else:
                n_fail += 1
                obs_failed = True

        # Migrate folders written by the old positional-index scheme by dropping
        # superseded "<date>_NNN.ext" files — but only once every asset for this
        # observation is safely downloaded, so a transient failure never deletes
        # the sole local copy of an asset.
        if not obs_failed:
            n_migrated += _remove_legacy_media_files(obs_dir, date_prefix, written)

        time.sleep(0.1)

    log.info(
        "Complete — downloaded: %d  no-date skipped: %d  failed: %d"
        "  legacy files migrated: %d",
        n_ok, n_skip, n_fail, n_migrated,
    )


# ── CLI ───────────────────────────────────────────────────────────────────────


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="tapestry_scraper",
        description="Download all Tapestry Journal observations for an account.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
examples:
  %(prog)s -e you@example.com -p hunter2 -o ./tapestry_export
  %(prog)s -e you@example.com -p hunter2 -o ./export --child 12345
  %(prog)s -e you@example.com -p hunter2 -o ./export --list-children
  %(prog)s -e you@example.com -p hunter2 -o ./export --verbose
        """,
    )
    p.add_argument("-e", "--email",    metavar="EMAIL",
                   default=os.environ.get("TAPESTRY_EMAIL"),
                   help="Tapestry login email address  (env: TAPESTRY_EMAIL)")
    p.add_argument("-p", "--password", metavar="PASSWORD",
                   default=os.environ.get("TAPESTRY_PASSWORD"),
                   help="Tapestry login password  (env: TAPESTRY_PASSWORD; "
                        "prefer the env var so the password never shows in `ps`)")
    p.add_argument("-o", "--output",   default="./tapestry_export",
                   metavar="DIR",
                   help="Output directory  (default: ./tapestry_export)")
    p.add_argument("--child", metavar="CHILD_ID",
                   help="Only download observations for this child (use --list-children to find IDs)")
    p.add_argument("--list-children", action="store_true",
                   help="Print available children and exit")
    p.add_argument("--limit", metavar="N", type=int,
                   help="Only process the first N observations (useful for testing)")
    p.add_argument("--session-cache", metavar="FILE",
                   default=os.environ.get("TAPESTRY_SESSION_CACHE"),
                   help="Reuse the login session stored in FILE, logging in only "
                        "when it has expired  (env: TAPESTRY_SESSION_CACHE)")
    p.add_argument("-v", "--verbose",  action="store_true",
                   help="Enable debug logging")
    return p


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    if not args.email or not args.password:
        parser.error("email and password are required "
                     "(-e/-p, or TAPESTRY_EMAIL/TAPESTRY_PASSWORD)")

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    output_dir = Path(args.output).expanduser().resolve()

    ts = TapestrySession()
    cache = Path(args.session_cache).expanduser() if args.session_cache else None
    if not (cache and ts.load_session(cache)):
        ts.login(args.email, args.password)
    if cache:
        ts.save_session(cache)

    if args.list_children:
        children = ts.get_children()
        if not children:
            print("No children found (the API may not expose this or "
                  "there are no children on this account).")
        else:
            print(f"\n{'ID':<14}  Name")
            print("─" * 40)
            for c in children:
                cid  = c.get("id", "?")
                name = c.get("name") or c.get("full_name") or c.get("display_name", "?")
                print(f"{cid!s:<14}  {name}")
        return

    observations = ts.get_observations(child_id=args.child)

    if not observations:
        log.warning(
            "No observations found. The account may be empty, or the site "
            "structure has changed. Try --verbose to see what's happening."
        )
        sys.exit(0)

    if args.limit:
        observations = observations[: args.limit]
        log.info("--limit %d: processing %d observation(s)", args.limit, len(observations))

    log.info("Processing %d observations → %s", len(observations), output_dir)
    organise(observations, output_dir, ts)


if __name__ == "__main__":
    main()
