#!/usr/bin/env python3
"""
create_api_tags.py
==================
Apply tags to APIs in the Noname / Akamai API Security Management API from a
hostname-association spreadsheet, with a full dry run and optional
auto-creation of one API group per distinct tag.

Built for a sheet shaped like:

    Application Name | Hostnames | Department | Market | Serviced Market

i.e. one row per hostname, with several independent columns that each describe
that host. Any of those columns can become a tag dimension. Tags are emitted
namespaced -- "dept:CPP", "market:Canada", "app:VCE" -- so the dimensions stay
distinguishable in the inventory and the generated group filters stay
unambiguous. Use --no-prefix for bare tag values instead.

Endpoint contract (Noname API Management, OpenAPI 3.0)
------------------------------------------------------
  GET   /api/v3/apis           -> inventory; paginate limit/offset, `moreEntities`
                                  each API carries `tags`: a FLAT LIST OF STRINGS
  PATCH /api/v4/apis/tags      -> append tags for a bulk of IDs   <- used here
  PUT   /api/v4/apis/tags      -> REPLACE tags for a bulk of IDs  (--replace)
  PATCH /api/v4/apis/{id}/tags -> append tags for a single ID     (fallback)
  GET   /api/v4/groups         -> list existing groups
  POST  /api/v4/groups         -> create a group (CreateGroupInput)
  DELETE/api/v4/groups/{id}    -> delete a group (used by --rollback)

CreateGroupInput:
  name           string  (required)  -- must NOT contain '/'
  description    string
  parentGroupId  string
  group_type     enum    (required)  APPLICATION | OTHER
  filters        array   (required)  of {field, operator, value}
      field    : api_owner | host | infrastructure_tags | method | path |
                 resources | sources | tags
      operator : in | contains | notContains | equals | notEquals |
                 startsWith | endsWith | blank | notBlank
      value    : string for equals/contains/*With, list for `in`

Credentials come from env NONAME_API_BASE / NONAME_API_TOKEN, else prompted
(token via getpass -- never echoed, never written to disk).

Quick start
-----------
  # 0. What's in the workbook? No network calls.
  python create_api_tags.py --sheet hosts.xlsx --list-sheets
  python create_api_tags.py --sheet hosts.xlsx --tab Main --inspect-sheet

  # 1. Dry run: tag on department only (the densest, cleanest dimension)
  python create_api_tags.py --sheet hosts.xlsx --tab Main \
      --tag-from "Department:dept"

  # 2. All three dimensions at once
  python create_api_tags.py --sheet hosts.xlsx --tab Main \
      --tag-from "Department:dept" \
      --tag-from "Market:market" \
      --tag-from "Application Name:app"

  # 3. Pull in a second tab, with the first tab winning any conflict
  python create_api_tags.py --sheet hosts.xlsx \
      --tab Main --tab "IRR-Akamai" \
      --tag-from "Department:dept" --tag-from "CORP Area:dept"

  # 4. Canary, then full apply, then groups
  python create_api_tags.py --sheet hosts.xlsx --tab Main \
      --tag-from "Department:dept" --apply --limit 25
  python create_api_tags.py --sheet hosts.xlsx --tab Main \
      --tag-from "Department:dept" --apply --create-groups \
      --group-parent "Departments"

  # 5. Undo a group-creation run
  python create_api_tags.py --rollback tag-run-20260731-140302.ledger.json
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
import time
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime, timezone
from getpass import getpass

try:
    import requests
except ImportError:
    sys.exit("Missing dependency: pip install requests")

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

# Applied to everything written to console or log file. A debug transcript is
# meant to be attachable to a support ticket, so credentials must never reach
# it -- not from headers, not from bodies, not from an echoed command line.
REDACTIONS = [
    (re.compile(r"(Bearer\s+)\S+", re.I), r"\1<redacted>"),
    (re.compile(r'("(?:token|apiKey|api_key|password|secret|authorization)"'
                r'\s*:\s*")[^"]*'), r"\1<redacted>"),
    (re.compile(r"(NONAME_API_TOKEN=)\S+"), r"\1<redacted>"),
]


def redact(text):
    out = str(text)
    for pattern, replacement in REDACTIONS:
        out = pattern.sub(replacement, out)
    return out


class Log:
    """Console + optional file logging at three levels.

    0 = normal, 1 = --verbose (decisions and per-batch detail),
    2 = --debug (full HTTP tracing).
    """

    NORMAL, VERBOSE, DEBUG = 0, 1, 2

    def __init__(self):
        self.level = self.NORMAL
        self.fh = None
        self.requests = Counter()
        self.timings = defaultdict(list)
        self.started = time.time()

    def configure(self, level=0, path=None, argv=None, base=None):
        self.level = level
        if path:
            self.fh = open(path, "a", encoding="utf-8")
            stamp = datetime.now(timezone.utc).isoformat()
            self.fh.write(f"\n{'=' * 72}\n")
            self.fh.write(f"run started {stamp}\n")
            if base:
                self.fh.write(f"tenant      {base}\n")
            if argv:
                self.fh.write(f"argv        {redact(' '.join(argv))}\n")
            self.fh.write(f"level       {level}\n")
            self.fh.write(f"{'=' * 72}\n")
            self.fh.flush()

    def _emit(self, text, to_console=True):
        text = redact(text)
        if to_console:
            print(text)
        if self.fh:
            self.fh.write(text + "\n")
            self.fh.flush()

    def out(self, text=""):
        self._emit(text)

    def raw(self, text, end="\n"):
        """Progress output that should not gain a newline on the console."""
        print(redact(text), end=end, flush=True)
        if self.fh and end == "\n":
            self.fh.write(redact(text) + "\n")

    def verbose(self, text):
        if self.level >= self.VERBOSE:
            self._emit(f"    . {text}")
        elif self.fh:
            self._emit(f"    . {text}", to_console=False)

    def debug(self, text):
        if self.level >= self.DEBUG:
            self._emit(f"    > {text}")
        elif self.fh:
            self._emit(f"    > {text}", to_console=False)

    def record(self, method, path, status, elapsed):
        self.requests[(method, path.split("?")[0], status)] += 1
        self.timings[(method, path.split("?")[0])].append(elapsed)

    def summary(self):
        if not self.requests:
            return
        total = sum(self.requests.values())
        self._emit("\n" + "=" * 72)
        self._emit(f"REQUEST SUMMARY  ({total} request(s) in "
                   f"{time.time() - self.started:.1f}s)")
        self._emit("=" * 72)
        self._emit(f"  {'METHOD':<7}{'ENDPOINT':<34}{'STATUS':>7}{'COUNT':>7}")
        self._emit("  " + "-" * 62)
        for (method, path, status), count in sorted(self.requests.items()):
            self._emit(f"  {method:<7}{path[:34]:<34}{status:>7}{count:>7}")
        self._emit("\n  Latency by endpoint (seconds):")
        for (method, path), times in sorted(self.timings.items()):
            self._emit(f"    {method:<7}{path[:34]:<34}"
                       f"n={len(times):<5} min={min(times):.2f} "
                       f"avg={sum(times) / len(times):.2f} max={max(times):.2f}")
        if self.fh:
            self._emit(f"\n  Transcript written to {self.fh.name}")

    def close(self):
        if self.fh:
            self.fh.close()
            self.fh = None


LOG = Log()


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

EP_INVENTORY = "/api/v3/apis"
EP_TAGS_BULK = "/api/v4/apis/tags"
EP_TAGS_ONE = "/api/v4/apis/{id}/tags"
EP_GROUPS = "/api/v4/groups"
EP_GROUPS_V3 = "/api/v3/groups"
EP_GROUP_ID = "/api/v4/groups/{id}"

# Keys whose nested objects are people or config, not groups.
NON_GROUP_KEYS = {
    "owner", "owners", "createdby", "updatedby", "modifiedby", "user",
    "users", "author", "filters", "filter", "metadata", "permissions",
    "roles", "team", "teams", "contact", "assignee",
}
ID_KEYS = ("id", "groupId", "group_id", "uuid", "guid")
CHILD_KEYS = ("children", "nodes", "subGroups", "subgroups", "childGroups")
# Top-level keys a paged inventory response has used for its row array.
ENTITY_KEYS = ("entities", "apis", "data", "items", "results", "records")

# Endpoint + method combinations to try for tag ASSIGNMENT. Only
# PATCH /api/v4/apis/tags had ever been probed, so an identical fault every
# time may simply mean that one handler is broken.
# PATCH appends tags, PUT replaces them. The v3 pair is marked deprecated in
# the spec ("Please use /api/v4/apis/tags") but is kept as a fallback.
TAG_WRITE_ROUTES = [
    ("PATCH", "/api/v4/apis/tags"),
    ("PATCH", "/api/v3/apis/tags"),
]
TAG_REPLACE_ROUTES = [
    ("PUT", "/api/v4/apis/tags"),
    ("PUT", "/api/v3/apis/tags"),
]

EP_TAGS_CATALOG = "/api/v4/tags"
EP_TAGS_CATALOG_V3 = "/api/v3/tags"

# Candidate request bodies for the bulk tag endpoint. The spec upload turned
# out to be a different API, so rather than guess once and fail across 13k
# APIs, the writer probes these against a two-API batch and adopts whichever
# the tenant accepts. `needs_ids` shapes reference tag IDs from the catalogue
# instead of free-text names.
def uuid_dashed(value):
    """018c78a642ab948bf4800ce68e912465 -> 018c78a6-42ab-948b-f480-0ce68e912465"""
    v = str(value).replace("-", "")
    if len(v) != 32:
        return str(value)
    return f"{v[:8]}-{v[8:12]}-{v[12:16]}-{v[16:20]}-{v[20:]}"


# CONFIRMED against the Noname Security API Management spec (v3.0).
#
#   PATCH /api/v4/apis/tags   ApiTagsController_updateApiTagsOfApi   (append)
#   PUT   /api/v4/apis/tags   ApiTagsController_replaceApiTagsOfApi  (replace)
#
#   editTagsOfApisBody:  { terms: ApiIds, tagIds: string[] }   both required
#   ApiIds:              { ids: string[] }
#
# `terms` is an OBJECT wrapping an ids array -- NOT an array of filter terms.
# Sending a filter array satisfied "terms should not be empty" and then failed
# inside the handler, which is the 500 we chased for several rounds.
#
# Tags are referenced by catalogue ID, never by name:
#   POST /api/v4/tags   createClassificationTag: { name }   -> { id, name }
# Note the tag object carries only name/id/legacyId, so tags are statically
# assigned -- there is no rule-based tag that matches APIs on its own.
BULK_SHAPES = [
    ("terms.ids + tagIds", "ids", True,
     lambda v, t: {"terms": {"ids": list(v)}, "tagIds": list(t)}),
    # Fallbacks, kept only in case a deployment predates the v4 contract.
    ("ids + tagIds", "ids", True,
     lambda v, t: {"ids": list(v), "tagIds": list(t)}),
    ("terms.ids + tags", "ids", False,
     lambda v, t: {"terms": {"ids": list(v)}, "tags": list(t)}),
]

# PATCH|PUT /api/v4/apis/{id}/tags -- editTagsOfApiById: { tagIds: string[] }
SINGLE_SHAPES = [
    ("tagIds", True, lambda t: {"tagIds": list(t)}),
]

HOST_BATCH_SIZE = 100

# Candidate bodies for POST /api/v4/tags. If tags here are rule-based
# ("classification" tags), the tag itself carries the filter and APIs matching
# it acquire the tag automatically -- there is no per-API assignment step, and
# a name-only tag necessarily matches nothing.
# createClassificationTag accepts only { name }. Tags are static.
TAG_CREATE_SHAPES = [
    ("name", lambda n, f: {"name": n}),
]

# Legacy single-shape constants, still honoured by --bulk-keys. The v4 bulk-tag body was not in the spec excerpt I
# had on hand -- these match the `ids` array convention used by the other v3/v4
# bulk endpoints. Run --probe to confirm against the tenant before a big apply;
# if the tenant differs, change these two constants only.
BULK_IDS_KEY = "ids"
BULK_TAGS_KEY = "tags"
MAX_CONSECUTIVE_FAILURES = 5
# Unique per run. A fixed name meant probe tags left on APIs by an earlier
# run looked like collateral damage from this one, and a correct body got
# rejected because of it. The prefix stays stable so cleanup still matches.
SCRATCH_PREFIX = "zz-shape-probe"
SCRATCH_TAG = f"{SCRATCH_PREFIX}-{datetime.now().strftime('%Y%m%d%H%M%S')}"
# Repeated identical 5xx means a server-side fault, not a wrong body shape.
MAX_IDENTICAL_SERVER_ERRORS = 2
MAX_LOG_BODY = 2000
# Set from CLI before apply_tags runs.
PROBE_BUDGET = [40, 20]
# [name_from, known_prefixes] -- set from CLI before groups are built.
GROUP_NAMING = ["value", ()]

# GroupFilterInput lists the full operator enum, but the server restricts it
# per field: `tags` rejects `equals` and expects contains | notContains | in.
# `in` takes an array value and is exact; `contains` is a substring match and
# would make dept:CPP also match dept:CPPX, so it is only a fallback.
GROUP_TAG_OPERATORS = ["in", "contains"]


def tag_filter(tag, operator):
    """One GroupFilterInput selecting APIs carrying `tag`."""
    value = [tag] if operator == "in" else tag
    return {"field": "tags", "operator": operator, "value": value}
# Observed on a live tenant: the handler resolves the APIs, then dies reading
# their existing tags. No request body gets past it, so probing further shapes
# only wastes time.
KNOWN_BROKEN_FAULT = "retrieve the tags associated with the specified api id"

def allowed_return_fields(resp):
    """Pull the valid returnFields list out of a 400 body.

    The API answers an unrecognised field with the full set of accepted ones,
    which is enough to correct the request and retry instead of failing.
    """
    try:
        body = resp.json()
    except ValueError:
        return None
    message = body.get("message") if isinstance(body, dict) else None
    if isinstance(message, list):
        message = " ".join(str(m) for m in message)
    if not message or "returnFields" not in message:
        return None
    match = re.search(r"following values:\s*(.+)", str(message), re.S)
    if not match:
        return None
    fields = [f.strip().strip("'\"") for f in match.group(1).split(",")]
    return [f for f in fields if f and " " not in f] or None


INVENTORY_PAGE_SIZE = 2000
TAG_BATCH_SIZE = 200
# Confirmed valid against a live tenant. `api_owner` is NOT a valid value --
# the API rejects the whole request with a 400 if any entry is unrecognised,
# and the rewritten matcher only needs host / path / id anyway.
RETURN_FIELDS = ["id", "host", "path", "method", "tags"]

# Placeholder values that look like data but are not. Compared case-folded.
DEFAULT_SENTINELS = {
    "not found", "notfound", "n/a", "na", "none", "null", "unknown",
    "tbd", "tba", "-", "--", "?", "not applicable", "not detected",
    "not enforced", "no data",
}

HOST_ALIASES = ["hostname", "hostnames", "host", "hosts", "fqdn", "domain",
                "domains", "server", "url", "endpointhost"]
ID_ALIASES = ["apiid", "id", "guid", "apiguid", "endpointid"]
PATH_ALIASES = ["path", "pathprefix", "endpoint", "uri", "route"]

# Columns never worth offering as a tag dimension.
NON_TAG_HEADERS = set(HOST_ALIASES) | set(ID_ALIASES) | set(PATH_ALIASES)


def _norm_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").strip().lower())


def clean_value(value):
    """NFKC-normalise (kills non-breaking spaces), collapse whitespace, trim."""
    if value is None:
        return None
    text = unicodedata.normalize("NFKC", str(value))
    text = re.sub(r"\s+", " ", text).strip()
    return text or None


def is_sentinel(value, sentinels):
    return value is not None and value.casefold() in sentinels


# ---------------------------------------------------------------------------
# Spreadsheet loading
# ---------------------------------------------------------------------------


def list_workbook(path):
    """Return [(sheet_name, headers, row_count)] without loading everything."""
    if path.lower().endswith((".csv", ".tsv", ".txt")):
        headers, rows = _load_delimited(path)
        return [("(single)", headers, len(rows))]

    try:
        from openpyxl import load_workbook
    except ImportError:
        sys.exit("Reading .xlsx needs openpyxl: pip install openpyxl")

    wb = load_workbook(path, read_only=True, data_only=True)
    out = []
    for name in wb.sheetnames:
        ws = wb[name]
        it = ws.iter_rows(values_only=True)
        headers = []
        for raw in it:
            if raw and any(c is not None and str(c).strip() for c in raw):
                headers = [str(c).strip() if c is not None else "" for c in raw]
                break
        count = sum(
            1 for raw in it
            if raw and any(c is not None and str(c).strip() for c in raw)
        )
        out.append((name, [h for h in headers if h], count))
    wb.close()
    return out


def _load_delimited(path):
    delim = "\t" if path.lower().endswith(".tsv") else ","
    with open(path, newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh, delimiter=delim)
        headers = list(reader.fieldnames or [])
        rows = [dict(r) for r in reader]
    return headers, rows


def load_tab(path, tab=None):
    """Return (headers, rows) for one worksheet (or the whole csv)."""
    if path.lower().endswith((".csv", ".tsv", ".txt")):
        return _load_delimited(path)

    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=True, data_only=True)
    if tab and tab not in wb.sheetnames:
        wb.close()
        sys.exit(f"No sheet named {tab!r}. Available: {', '.join(wb.sheetnames)}")
    ws = wb[tab] if tab else wb[wb.sheetnames[0]]

    it = ws.iter_rows(values_only=True)
    headers = []
    for raw in it:
        if raw and any(c is not None and str(c).strip() for c in raw):
            headers = [str(c).strip() if c is not None else "" for c in raw]
            break

    rows = []
    for raw in it:
        if not raw or all(c is None or str(c).strip() == "" for c in raw):
            continue
        rows.append({
            headers[i]: raw[i]
            for i in range(min(len(headers), len(raw)))
            if headers[i]
        })
    wb.close()
    return [h for h in headers if h], rows


def find_column(headers, aliases, override=None):
    if override:
        if override not in headers:
            sys.exit(f"Column {override!r} not in sheet. Headers: "
                     f"{', '.join(map(str, headers))}")
        return override
    lookup = {}
    for h in headers:
        lookup.setdefault(_norm_header(h), h)
    for alias in aliases:
        if alias in lookup:
            return lookup[alias]
    return None


# ---------------------------------------------------------------------------
# Tag dimensions
# ---------------------------------------------------------------------------


class Dimension:
    """One spreadsheet column promoted to a tag namespace."""

    def __init__(self, column, prefix=None):
        self.column = column
        self.prefix = prefix

    def tag(self, raw, use_prefix=True):
        value = clean_value(raw)
        if value is None:
            return None
        if use_prefix and self.prefix:
            return f"{self.prefix}:{value}"
        return value

    def __repr__(self):
        return f"{self.column} -> {self.prefix + ':' if self.prefix else ''}<value>"


def parse_specs(specs):
    """'Department:dept' -> ('Department', 'dept'). Prefix is optional."""
    out = []
    for spec in specs:
        if ":" in spec:
            column, prefix = spec.rsplit(":", 1)
            out.append((column.strip(), prefix.strip()))
        else:
            out.append((spec.strip(), None))
    return out


def resolve_dimensions(specs, headers):
    """Resolve parsed specs against ONE tab's headers.

    A column missing from this tab is simply not a dimension here -- with
    several --tab values the columns legitimately differ (e.g. 'Department'
    on one tab, 'CORP Area' on another). Returns (dims, matched_columns).
    """
    dims, matched = [], set()
    for column, prefix in specs:
        found = column if column in headers else find_column(
            headers, [_norm_header(column)])
        if not found:
            continue
        dims.append(Dimension(found, prefix))
        matched.add(column)
    return dims, matched


def suggest_dimensions(headers, rows, sentinels, limit=12):
    """Score each column's usefulness as a tag dimension."""
    out = []
    for header in headers:
        if _norm_header(header) in NON_TAG_HEADERS:
            continue
        values = [clean_value(r.get(header)) for r in rows]
        usable = [v for v in values if v and not is_sentinel(v, sentinels)]
        blanks = len(values) - len(usable)
        distinct = len(set(usable))
        if not usable:
            out.append((header, 0, 0, blanks, "empty -- skip"))
            continue
        if distinct == len(usable):
            note = "every value unique -- not a grouping dimension"
        elif distinct > 200:
            note = f"{distinct} distinct -- very high cardinality"
        else:
            note = "usable"
        out.append((header, len(usable), distinct, blanks, note))
    out.sort(key=lambda t: -t[1])
    return out[:limit]


# ---------------------------------------------------------------------------
# Rules
# ---------------------------------------------------------------------------


class Rule:
    """One host + one tag, traceable back to a sheet row."""

    __slots__ = ("tab", "row_no", "tag", "dimension", "value",
                 "api_id", "host", "path", "matched")

    def __init__(self, tab, row_no, tag, dimension, value,
                 api_id=None, host=None, path=None):
        self.tab = tab
        self.row_no = row_no
        self.tag = tag
        self.dimension = dimension
        self.value = value
        self.api_id = api_id
        self.host = (host or "").lower() or None
        self.path = path
        self.matched = 0

    def criteria(self):
        bits = []
        if self.api_id:
            bits.append(f"id={self.api_id}")
        if self.host:
            bits.append(f"host={self.host}")
        if self.path:
            bits.append(f"path~{self.path}")
        return " AND ".join(bits) or "(no criteria)"

    def matches(self, api):
        if self.api_id:
            return str(api.get("id")) == self.api_id
        if not (self.host or self.path):
            return False
        if self.host and str(api.get("host", "")).lower() != self.host:
            return False
        if self.path and not str(api.get("path", "")).startswith(self.path):
            return False
        return True


def build_rules(tabs, dimensions_by_tab, host_col_by_tab, id_col_by_tab,
                path_col_by_tab, sentinels, use_prefix):
    """Walk every loaded tab and emit one Rule per (host, dimension) pair.

    Also returns cross-row conflicts: the same host given two different values
    for the same dimension namespace.
    """
    rules = []
    skipped = []
    seen = defaultdict(dict)   # host -> {prefix: (value, tab, row)}
    conflicts = []

    for tab_name, rows in tabs:
        dims = dimensions_by_tab[tab_name]
        host_col = host_col_by_tab.get(tab_name)
        id_col = id_col_by_tab.get(tab_name)
        path_col = path_col_by_tab.get(tab_name)

        for row_no, row in enumerate(rows, start=2):
            api_id = clean_value(row.get(id_col)) if id_col else None
            host = clean_value(row.get(host_col)) if host_col else None
            path = clean_value(row.get(path_col)) if path_col else None

            if not (api_id or host):
                skipped.append((tab_name, row_no, "no hostname or API id"))
                continue

            emitted = 0
            for dim in dims:
                value = clean_value(row.get(dim.column))
                if value is None:
                    continue
                if is_sentinel(value, sentinels):
                    continue

                key = host.lower() if host else f"id:{api_id}"
                namespace = dim.prefix or dim.column
                prior = seen[key].get(namespace)
                if prior and prior[0] != value:
                    conflicts.append({
                        "host": host or api_id,
                        "dimension": namespace,
                        "first": {"value": prior[0], "tab": prior[1],
                                  "row": prior[2]},
                        "second": {"value": value, "tab": tab_name,
                                   "row": row_no},
                    })
                    continue  # first sighting wins; --tab order sets priority
                seen[key][namespace] = (value, tab_name, row_no)

                tag = dim.tag(value, use_prefix)
                rules.append(Rule(tab_name, row_no, tag, namespace, value,
                                  api_id, host, path))
                emitted += 1

            if emitted == 0:
                reason = (f"host {host or api_id!r}: no usable value on "
                          "any tag dimension")
                skipped.append((tab_name, row_no, reason))
                LOG.verbose(f"skip [{tab_name}] row {row_no}: {reason}")

    return rules, skipped, conflicts


class FailedResponse:
    """Stands in for a requests.Response when the request never completed."""

    def __init__(self, exc):
        self.exc = exc
        self.status_code = 0
        self.ok = False
        self.content = b""
        self.text = f"{exc.__class__.__name__}: {exc}"

    def json(self):
        raise ValueError("no response body")


# ---------------------------------------------------------------------------
# Group payload harvesting
# ---------------------------------------------------------------------------


def harvest_group_records(payload):
    """Every group anywhere in the payload, with id, name and parentage.

    Parentage comes from nesting and drives children-before-parents deletion.
    """
    out = []

    def walk(obj, parent_id, parent_name, depth):
        if isinstance(obj, list):
            for item in obj:
                walk(item, parent_id, parent_name, depth)
            return
        if not isinstance(obj, dict):
            return
        name = obj.get("name") or obj.get("groupName")
        gid = next((str(obj[k]) for k in ID_KEYS
                    if obj.get(k) not in (None, "")), None)
        if name and gid:
            out.append({"id": gid, "name": str(name), "parent_id": parent_id,
                        "parent_name": parent_name, "depth": depth,
                        "filters": obj.get("filters")})
            parent_id, parent_name, depth = gid, str(name), depth + 1
        for key, value in obj.items():
            if str(key).strip().lower() in NON_GROUP_KEYS:
                continue
            if isinstance(value, (dict, list)):
                walk(value, parent_id, parent_name, depth)

    walk(payload, None, None, 0)
    seen, unique = set(), []
    for rec in out:
        if rec["id"] not in seen:
            seen.add(rec["id"])
            unique.append(rec)
    return unique


def harvest_group_names(payload):
    """Every group name anywhere in the payload, whatever the wrapper."""
    names = set()

    def walk(obj):
        if isinstance(obj, list):
            for item in obj:
                walk(item)
            return
        if not isinstance(obj, dict):
            return
        name = obj.get("name") or obj.get("groupName")
        has_id = any(obj.get(k) not in (None, "") for k in ID_KEYS)
        if name and has_id:
            names.add(str(name))
        for key, value in obj.items():
            if str(key).strip().lower() in NON_GROUP_KEYS:
                continue
            if isinstance(value, (dict, list)):
                walk(value)

    walk(payload)
    return names


# ---------------------------------------------------------------------------
# API client
# ---------------------------------------------------------------------------


class Client:
    def __init__(self, base, token, timeout=60, verify=True,
                 page_size=INVENTORY_PAGE_SIZE):
        self.base = base.rstrip("/")
        self.timeout = timeout
        self.page_size = page_size
        self.session = requests.Session()
        self.session.verify = verify
        self.session.headers.update({
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        })

    def request(self, method, path, max_retries=4, timeout=None, **kwargs):
        url = self.base + path
        delay = 2.0
        timeout = timeout or self.timeout
        for attempt in range(max_retries + 1):
            label = f"{method} {path}"
            if attempt:
                label += f" (attempt {attempt + 1}/{max_retries + 1})"

            if LOG.level >= Log.DEBUG or LOG.fh:
                LOG.debug(f"--> {label}")
                params = kwargs.get("params")
                if params:
                    LOG.debug(f"    params  {params}")
                body = kwargs.get("data")
                if body:
                    LOG.debug(f"    body    {len(body)} bytes")
                    LOG.debug(f"    payload {str(body)[:MAX_LOG_BODY]}"
                              + (" ...truncated" if len(str(body)) > MAX_LOG_BODY
                                 else ""))

            started = time.time()
            try:
                resp = self.session.request(method, url, timeout=timeout,
                                            **kwargs)
            except requests.RequestException as exc:
                elapsed = time.time() - started
                LOG.record(method, path, 0, elapsed)
                LOG.debug(f"<-- {exc.__class__.__name__} after {elapsed:.2f}s")
                if attempt == max_retries:
                    raise
                LOG.out(f"    ! {exc.__class__.__name__} on {method} {path} "
                        f"(timeout {timeout}s); retry in {delay:.0f}s")
                time.sleep(delay)
                delay *= 2
                continue

            elapsed = time.time() - started
            LOG.record(method, path, resp.status_code, elapsed)
            size = len(resp.content or b"")
            LOG.debug(f"<-- {resp.status_code} in {elapsed:.2f}s, {size} bytes")
            if LOG.level >= Log.DEBUG or LOG.fh:
                interesting = {k: v for k, v in resp.headers.items()
                               if k.lower() in ("content-type", "retry-after",
                                                "x-request-id", "x-trace-id",
                                                "x-ratelimit-remaining")}
                if interesting:
                    LOG.debug(f"    headers {interesting}")
                if not resp.ok or LOG.level >= Log.DEBUG:
                    LOG.debug(f"    response {resp.text[:MAX_LOG_BODY]}"
                              + (" ...truncated"
                                 if len(resp.text) > MAX_LOG_BODY else ""))

            if resp.status_code in (429, 500, 502, 503, 504) \
                    and attempt < max_retries:
                wait = float(resp.headers.get("Retry-After") or delay)
                LOG.out(f"    ! HTTP {resp.status_code} on {method} {path}; "
                        f"retry in {wait:.0f}s")
                time.sleep(wait)
                delay *= 2
                continue
            return resp
        raise RuntimeError("unreachable")

    def preflight(self):
        try:
            resp = self.request("GET", f"{EP_INVENTORY}?limit=1&offset=0")
        except requests.RequestException as exc:
            sys.exit(f"\nPreflight could not reach {self.base}: "
                     f"{exc.__class__.__name__}.\n"
                     "That is the cheapest call this script makes, so the "
                     "tenant is slow or\nunreachable rather than the request "
                     "being wrong. Try again, raise\n--timeout, or check VPN "
                     "and tenant status.")
        if resp.status_code == 401:
            sys.exit("Preflight failed: 401 Unauthorized. Token expired or invalid "
                     "(Management API tokens are short-lived -- regenerate it).")
        if resp.status_code == 403:
            sys.exit("Preflight failed: 403 Forbidden. The token's role lacks "
                     "inventory read access.")
        if resp.status_code == 404:
            sys.exit(f"Preflight failed: 404 at {self.base}{EP_INVENTORY}. "
                     "Check the base URL.")
        if not resp.ok:
            sys.exit(f"Preflight failed: HTTP {resp.status_code} {resp.text[:300]}")
        print(f"  Preflight OK ({self.base})")

    def probe(self):
        print("\nEndpoint probe (read-only):")
        for path in (f"{EP_INVENTORY}?limit=1&offset=0", EP_GROUPS):
            resp = self.request("GET", path)
            print(f"  GET  {path.split('?')[0]:<28} -> {resp.status_code}")
        print(f'\n  Bulk tag body will be sent as: '
              f'{{"{BULK_IDS_KEY}": [...], "{BULK_TAGS_KEY}": [...]}}')
        print("  If the tenant rejects that shape with a 400, adjust "
              "BULK_IDS_KEY / BULK_TAGS_KEY at the top of this script.")

    def fetch_inventory(self, return_fields=None, quiet=False,
                        max_pages=None):
        """Page the inventory, negotiating how `returnFields` is serialised.

        The parameter is validated as an ARRAY server-side ("each value in
        returnFields must be..."), so a comma-joined string is read as one
        bogus field name and rejected even when every name in it is valid.
        Encoding conventions differ between deployments, so try each in turn
        and fall back to omitting the parameter entirely -- heavier payloads,
        but it always works.
        """
        fields = list(return_fields or RETURN_FIELDS)
        styles = ["repeat", "csv", "bracket", "omit"]
        style = 0
        trimmed = False
        apis, offset, page = [], 0, 0

        if not quiet:
            print("  Fetching inventory", end="", flush=True)
        while True:
            params = {"limit": self.page_size, "offset": offset}
            if fields and styles[style] == "repeat":
                params["returnFields"] = fields          # ?f=a&f=b
            elif fields and styles[style] == "csv":
                params["returnFields"] = ",".join(fields)
            elif fields and styles[style] == "bracket":
                params["returnFields[]"] = fields

            resp = self.request("GET", EP_INVENTORY, params=params)

            if resp.status_code == 400:
                allowed = allowed_return_fields(resp)

                # An actually-invalid name: drop it and retry once.
                if allowed and not trimmed:
                    keep = [f for f in fields if f in allowed]
                    drop = [f for f in fields if f not in allowed]
                    if keep and drop:
                        trimmed = True
                        fields = keep
                        if not quiet:
                            print(f"\n  ! tenant rejected returnFields "
                                  f"{', '.join(drop)}; retrying without them")
                            print("  Fetching inventory", end="", flush=True)
                        continue

                # Every name is valid, so it is the encoding that is wrong.
                if style + 1 < len(styles):
                    style += 1
                    label = ("omitting returnFields (full objects)"
                             if styles[style] == "omit"
                             else f"returnFields as {styles[style]}")
                    if not quiet:
                        print(f"\n  ! 400 on returnFields encoding; retrying "
                              f"with {label}")
                        print("  Fetching inventory", end="", flush=True)
                    continue

            if not resp.ok:
                sys.exit(f"\nInventory fetch failed: HTTP {resp.status_code} "
                         f"{resp.text[:500]}")

            body = resp.json()
            if isinstance(body, list):
                chunk = body
            else:
                chunk = next((body[k] for k in ENTITY_KEYS
                              if isinstance(body.get(k), list)), [])
            apis.extend(chunk)
            page += 1
            if not quiet:
                print(".", end="", flush=True)
            LOG.verbose(f"inventory page {page}: offset {offset}, "
                        f"{len(chunk):,} row(s), {len(apis):,} total")
            if max_pages and page >= max_pages:
                break

            more = body.get("moreEntities") if isinstance(body, dict) else None
            if more is None:
                more = len(chunk) == self.page_size
            if not more or not chunk:
                break
            offset += self.page_size

        if not quiet:
            print(f" {len(apis):,} APIs across {page} page(s)")
        if not apis and not quiet:
            print("  ! Inventory came back empty. Re-run with --dump-raw "
                  "if that looks wrong.")
        return apis

    def fetch_groups(self, dump_raw=False):
        """Existing group names, tolerating several response shapes.

        Tenants have returned the tree as a bare list, as {"entities": [...]},
        under "nodes", and as a single root object. Guessing wrong here is
        expensive: an empty result looks like "no groups exist" and every
        create then collides with a 409.
        """
        for endpoint in (EP_GROUPS, EP_GROUPS_V3):
            resp = self.request("GET", endpoint)
            if not resp.ok:
                continue
            try:
                body = resp.json()
            except ValueError:
                continue
            if dump_raw:
                print(f"\n--- raw {endpoint} (first 2000 chars) ---")
                print(json.dumps(body)[:2000])
                print("--- end raw ---\n")
            names = harvest_group_names(body)
            if names:
                return names
        print("  ! Could not read existing groups; duplicate detection is off. "
              "Creates will rely on the tenant returning 409 for duplicates.")
        return set()

    # -- tag catalogue -----------------------------------------------------

    def fetch_tag_catalog(self):
        """{tag name: tag id} if the tenant models tags as first-class objects.

        The inventory exposes both `tags` and `tagIds`, so some deployments
        expect writes to reference IDs rather than names.
        """
        for endpoint in (EP_TAGS_CATALOG, EP_TAGS_CATALOG_V3):
            resp = self.request("GET", endpoint)
            if not resp.ok:
                continue
            try:
                body = resp.json()
            except ValueError:
                continue
            catalog = {}

            def walk(obj):
                if isinstance(obj, list):
                    for item in obj:
                        walk(item)
                    return
                if not isinstance(obj, dict):
                    return
                name = obj.get("name") or obj.get("tagName")
                tid = next((obj[k] for k in ID_KEYS
                            if obj.get(k) not in (None, "")), None)
                if name and tid is not None:
                    catalog[str(name)] = str(tid)
                for key, value in obj.items():
                    if isinstance(value, (dict, list)):
                        walk(value)

            walk(body)
            if catalog:
                return catalog, endpoint
        return {}, None

    def create_tag(self, name, endpoint=EP_TAGS_CATALOG):
        return self.request("POST", endpoint,
                            data=json.dumps({"name": name}))

    # -- tag writes --------------------------------------------------------

    def safe_request(self, method, path, **kwargs):
        """Like request(), but a transport failure becomes a response-like
        object instead of an exception. Probing must never crash the run."""
        try:
            return self.request(method, path, **kwargs)
        except requests.RequestException as exc:
            return FailedResponse(exc)

    def write_tags(self, selectors, values, shape, route=None, retries=1,
                   timeout=None):
        """Send one bulk tag write using a shape from BULK_SHAPES.

        `selectors` are hostnames or API ids depending on shape[1].
        """
        build = shape[3]
        body = build(list(selectors), list(values))
        method, path = route or TAG_WRITE_ROUTES[0]
        return self.safe_request(method, path, max_retries=retries,
                                 timeout=timeout, data=json.dumps(body))

    def write_tags_single(self, api_id, values, shape, retries=1):
        build = shape[2]
        body = build(list(values))
        return self.safe_request("PATCH", EP_TAGS_ONE.format(id=api_id),
                                 max_retries=retries, data=json.dumps(body))

    def fetch_group_records(self):
        for endpoint in (EP_GROUPS, EP_GROUPS_V3):
            resp = self.safe_request("GET", endpoint, max_retries=1)
            if not resp.ok:
                continue
            try:
                body = resp.json()
            except ValueError:
                continue
            records = harvest_group_records(body)
            if records:
                return records
        return []

    def create_group(self, body):
        return self.safe_request("POST", EP_GROUPS, data=json.dumps(body))

    def delete_group(self, group_id):
        return self.request("DELETE", EP_GROUP_ID.format(id=group_id))


# ---------------------------------------------------------------------------
# Planning
# ---------------------------------------------------------------------------


def load_inventory(client, cache_path=None, max_age_hours=24,
                   return_fields=None):
    """Fetch the inventory, reusing a local cache when one is available.

    Paging 35k endpoints takes many requests; a timeout on the last page
    otherwise discards every page before it. The cache turns a retry into a
    local read.
    """
    if cache_path and os.path.exists(cache_path):
        age = (time.time() - os.path.getmtime(cache_path)) / 3600
        if age <= max_age_hours:
            try:
                with open(cache_path, encoding="utf-8") as fh:
                    cached = json.load(fh)
                apis = cached.get("apis") if isinstance(cached, dict) else cached
                if apis:
                    print(f"  Using cached inventory from {cache_path} "
                          f"({len(apis):,} APIs, {age:.1f}h old)")
                    print("  Delete that file or pass --refresh-inventory for "
                          "a fresh pull.")
                    return apis
            except (ValueError, OSError) as exc:
                print(f"  ! Could not read {cache_path} ({exc}); fetching fresh")
        else:
            print(f"  Cache {cache_path} is {age:.1f}h old; fetching fresh")

    apis = client.fetch_inventory(return_fields)

    if cache_path and apis:
        try:
            with open(cache_path, "w", encoding="utf-8") as fh:
                json.dump({"fetched": datetime.now(timezone.utc).isoformat(),
                           "base": client.base, "apis": apis}, fh)
            print(f"  Inventory cached to {cache_path}")
        except OSError as exc:
            print(f"  ! Could not write {cache_path}: {exc}")
    return apis


def build_plan(rules, apis):
    by_id = {str(a.get("id")): a for a in apis}
    by_host = defaultdict(list)
    for api in apis:
        by_host[str(api.get("host", "")).lower()].append(api)

    plan = defaultdict(set)
    already = defaultdict(set)
    resolved = defaultdict(set)        # tag -> api_ids, new AND already there
    resolved_hosts = defaultdict(set)  # tag -> hostnames, for filter-based writes
    unresolved_hosts = set()

    for rule in rules:
        if rule.api_id:
            candidates = [by_id[rule.api_id]] if rule.api_id in by_id else []
        else:
            candidates = by_host.get(rule.host, [])

        hit = False
        for api in candidates:
            if not rule.matches(api):
                continue
            hit = True
            rule.matched += 1
            api_id = str(api.get("id"))
            resolved[rule.tag].add(api_id)
            host = str(api.get("host", "")).lower()
            if host:
                resolved_hosts[rule.tag].add(host)
            existing = a_tags(api)
            if rule.tag in existing:
                already[rule.tag].add(api_id)
            else:
                plan[api_id].add(rule.tag)
        if not hit and rule.host:
            unresolved_hosts.add(rule.host)

    LOG.verbose(f"plan: {len(plan):,} API(s) need at least one new tag; "
                f"{len(unresolved_hosts):,} sheet host(s) absent from inventory")
    for tag in sorted(resolved, key=lambda t: -len(resolved[t])):
        LOG.verbose(f"  {tag}: {len(resolved[tag]):,} API(s) across "
                    f"{len(resolved_hosts.get(tag, ())):,} host(s), "
                    f"{len(already.get(tag, ())):,} already tagged")
    if unresolved_hosts:
        sample = sorted(unresolved_hosts)[:10]
        LOG.verbose(f"  hosts not in inventory (first {len(sample)}): "
                    + ", ".join(sample))
    return plan, already, unresolved_hosts, resolved, resolved_hosts


def invert_plan(plan):
    by_tag = defaultdict(list)
    for api_id, tags in plan.items():
        for tag in tags:
            by_tag[tag].append(api_id)
    return by_tag


def group_label(tag, name_from="value", known_prefixes=()):
    """Display name for a tag's group.

    "value" strips the dimension namespace, so tag `dept:CPP` becomes group
    "CPP" while the filter still matches the full tag. "tag" keeps it verbatim.
    """
    if name_from == "tag":
        return tag
    for pfx in known_prefixes:
        if pfx and tag.startswith(f"{pfx}:"):
            return tag[len(pfx) + 1:]
    return tag.split(":", 1)[1] if ":" in tag else tag


def sanitize_group_name(tag, prefix=""):
    """Group names must not contain '/'."""
    name = f"{prefix}{tag}" if prefix else str(tag)
    name = name.replace("/", "-")
    name = re.sub(r"\s+", " ", name).strip()
    return name[:120]


# ---------------------------------------------------------------------------
# Reporting
# ---------------------------------------------------------------------------


def print_dry_run(rules, plan, already, by_tag, skipped, conflicts,
                  unresolved_hosts, apis, create_groups, existing_group_names,
                  group_prefix, group_tags=None, sample=5):
    print("\n" + "=" * 72)
    print("DRY RUN -- no writes performed")
    print("=" * 72)

    print(f"\nInventory scanned      : {len(apis):,} APIs")
    print(f"Rules from spreadsheet : {len(rules):,}")
    print(f"Rows skipped           : {len(skipped):,}")
    print(f"Distinct tags to apply : {len(by_tag):,}")
    print(f"APIs receiving tags    : {len(plan):,}")
    print(f"Tag assignments to add : {sum(len(v) for v in plan.values()):,}")
    print(f"Already tagged (no-op) : {sum(len(v) for v in already.values()):,}")

    per_dim = defaultdict(set)
    for rule in rules:
        per_dim[rule.dimension].add(rule.tag)
    if per_dim:
        print("\nTags per dimension:")
        for dim in sorted(per_dim):
            print(f"  {dim:<20} {len(per_dim[dim]):>5} distinct tag(s)")

    if conflicts:
        print(f"\n!! {len(conflicts)} CONFLICT(S): the same host was given two "
              "different values")
        print("   for the same dimension. The FIRST --tab listed wins; the "
              "second is dropped.")
        by_pair = defaultdict(int)
        for c in conflicts:
            by_pair[(c["dimension"], c["first"]["value"],
                     c["second"]["value"])] += 1
        for (dim, first, second), count in sorted(by_pair.items(),
                                                  key=lambda kv: -kv[1])[:10]:
            print(f'   {dim}: kept "{first}" over "{second}"  ({count} host(s))')
        if len(by_pair) > 10:
            print(f"   ... and {len(by_pair) - 10} more pairing(s)")
        print("   Full list is in the --plan-out file.")

    if by_tag:
        print(f"\nPer-tag breakdown (top 25 of {len(by_tag)}):")
        print(f"  {'TAG':<44} {'NEW':>8} {'EXISTS':>8}")
        print("  " + "-" * 62)
        for tag in sorted(by_tag, key=lambda t: -len(by_tag[t]))[:25]:
            print(f"  {tag[:44]:<44} {len(by_tag[tag]):>8,} "
                  f"{len(already.get(tag, [])):>8,}")

    if unresolved_hosts:
        print(f"\n{len(unresolved_hosts):,} hostname(s) from the sheet are NOT "
              "in the inventory.")
        print("  These are hosts the platform has not observed traffic for, "
              "or a naming mismatch.")
        print("  They cannot be tagged. See the --plan-out file for the list.")

    if skipped:
        print(f"\n{len(skipped):,} spreadsheet row(s) skipped:")
        reasons = defaultdict(int)
        for _, _, reason in skipped:
            reasons[re.sub(r"'[^']*'", "'...'", reason)] += 1
        for reason, count in sorted(reasons.items(), key=lambda kv: -kv[1])[:8]:
            print(f"  {count:>6,}  {reason}")

    if plan:
        print(f"\nSample of planned changes (first {sample}):")
        for api_id, tags in list(plan.items())[:sample]:
            print(f"  {api_id}  += {sorted(tags)}")

    if create_groups:
        print("\n" + "-" * 72)
        print("GROUP AUTO-CREATION (one group per distinct tag)")
        print("-" * 72)
        new_groups, dupes = [], []
        for tag in sorted(group_tags if group_tags is not None else by_tag):
            name = sanitize_group_name(
                group_label(tag, GROUP_NAMING[0], GROUP_NAMING[1]),
                group_prefix)
            (dupes if name in existing_group_names else new_groups).append((name, tag))
        print(f"  Groups to create : {len(new_groups)}")
        print(f"  Already present  : {len(dupes)} (will be skipped)")
        if len(new_groups) > 60:
            print(f"\n  !! {len(new_groups)} groups is a lot. Consider running "
                  "one dimension")
            print("     at a time, or dropping the high-cardinality ones.")
        for name, tag in new_groups[:15]:
            print(f'    + "{name}"  <- filters: tags '
                  f'{GROUP_TAG_OPERATORS[0]} "{tag}"')
        if len(new_groups) > 15:
            print(f"    ... and {len(new_groups) - 15} more")

    print("\n" + "=" * 72)
    print("Re-run with --apply to execute. Add --limit N for a canary batch.")
    print("=" * 72)


def write_plan_file(path, rules, plan, by_tag, skipped, conflicts,
                    unresolved_hosts):
    payload = {
        "generated": datetime.now(timezone.utc).isoformat(),
        "summary": {
            "rules": len(rules),
            "apis_affected": len(plan),
            "assignments": sum(len(v) for v in plan.values()),
            "distinct_tags": len(by_tag),
            "conflicts": len(conflicts),
            "hosts_not_in_inventory": len(unresolved_hosts),
        },
        "tags": {tag: sorted(ids) for tag, ids in by_tag.items()},
        "conflicts": conflicts,
        "hosts_not_in_inventory": sorted(unresolved_hosts),
        "skipped_rows": [{"tab": t, "row": n, "reason": why}
                         for t, n, why in skipped],
    }
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=2)
    print(f"\nPlan written to {path}")


# ---------------------------------------------------------------------------
# Apply
# ---------------------------------------------------------------------------


class TagResolver:
    """Maps tag names to catalogue IDs, creating entries when needed."""

    def __init__(self, client):
        self.client = client
        self.catalog, self.endpoint = client.fetch_tag_catalog()
        if self.endpoint:
            print(f"  Tag catalogue at {self.endpoint}: "
                  f"{len(self.catalog):,} existing tag(s)")

    @property
    def available(self):
        return self.endpoint is not None

    def ids_for(self, names, create=True):
        if not self.available:
            return None
        out = []
        for name in names:
            if name not in self.catalog:
                if not create:
                    return None
                resp = self.client.create_tag(name, self.endpoint)
                if not resp.ok:
                    return None
                try:
                    body = resp.json() if resp.content else {}
                except ValueError:
                    return None
                tid = next((body[k] for k in ID_KEYS
                            if isinstance(body, dict)
                            and body.get(k) not in (None, "")), None)
                if tid is None:
                    return None
                self.catalog[name] = str(tid)
                LOG.verbose(f"created catalogue tag '{name}' -> id {tid}")
            else:
                LOG.debug(f"catalogue hit '{name}' -> id {self.catalog[name]}")
            out.append(self.catalog[name])
        return out


def error_detail(resp, width=300):
    """Readable validator output. NestJS returns `message` as a list."""
    try:
        body = resp.json()
    except ValueError:
        return resp.text[:width]
    message = body.get("message") if isinstance(body, dict) else body
    if isinstance(message, list):
        return "; ".join(str(m) for m in message)[:width]
    return str(message or body)[:width]


def scope_check(client, probe_tag, target_ids):
    """Confirm a filter-based write hit ONLY the APIs we aimed at.

    A `terms` body selects by filter. Get the filter wrong and the write can
    land on far more than intended -- on this estate, potentially every API.
    Re-read a page of inventory and fail the shape if the probe tag shows up
    anywhere outside the target set.
    """
    try:
        page = client.fetch_inventory(quiet=True, max_pages=1)
    except SystemExit:
        return None
    if not page:
        return None
    targets = {str(i) for i in target_ids}
    stray = []
    for api in page:
        tags = a_tags(api)
        if probe_tag in tags and str(api.get("id")) not in targets:
            stray.append(str(api.get("id")))
    return stray


def a_tags(api):
    """Tag names on an API, whether the field holds strings or objects."""
    out = set()
    for tag in (api.get("tags") or []):
        if isinstance(tag, dict):
            name = tag.get("name") or tag.get("tagName")
            if name:
                out.add(str(name))
        else:
            out.add(str(tag))
    return out


def _probe(client, route, shape, resolver, scratch, targets, timeout):
    """One probe. Returns (response, elapsed)."""
    needs_ids = shape[2]
    values = resolver.ids_for([scratch]) if needs_ids else [scratch]
    if values is None:
        return None, 0.0
    started = time.time()
    resp = client.write_tags(targets, values, shape, route, retries=0,
                             timeout=timeout)
    return resp, time.time() - started


def cleanup_probe_tag(client, name):
    """Delete this run's scratch tag so it never pollutes a later probe."""
    resolver = TagResolver(client)
    tag_id = resolver.catalog.get(name)
    if not resolver.available or not tag_id:
        return
    resp = client.safe_request("DELETE", f"{resolver.endpoint}/{tag_id}",
                               max_retries=0)
    if resp.ok:
        LOG.verbose(f"removed probe tag '{name}'")
    else:
        LOG.out(f"  (probe tag '{name}' left behind: HTTP {resp.status_code}; "
                f"clear it with --delete-tags {SCRATCH_PREFIX} --apply)")


def negotiate_bulk_shape(client, resolver, sample_ids, sample_hosts,
                         expected_ids, replace=False, max_probes=40,
                         probe_timeout=20):
    """Find a route + body the tenant accepts, in two phases.

    Phase 1 sends one representative body per route to see which routes exist
    at all -- six cheap requests instead of seventy-eight. Phase 2 sweeps body
    shapes only on routes that responded with something other than 404/405.
    Probes use a short timeout: a slow tenant should not turn a sweep into a
    twenty-minute stall.
    """
    scratch = SCRATCH_TAG
    routes = TAG_REPLACE_ROUTES if replace else TAG_WRITE_ROUTES
    LOG.out(f"\n  Negotiating tag write (probe timeout {probe_timeout}s, "
            f"budget {max_probes} probes)")
    LOG.out(f"  Probing with scratch tag '{scratch}' so a mis-scoped filter "
            "cannot\n  spray a real tag across the estate.")

    def targets_for(shape):
        return sample_hosts if shape[1] == "hosts" else sample_ids

    probes = 0
    representative = next((sh for sh in BULK_SHAPES if sh[1] == "hosts"),
                          BULK_SHAPES[0])

    # ---- phase 1: which routes are even there? ----
    LOG.out(f"\n  Phase 1 - route check ({len(routes)} routes)")
    alive = []
    for method, path in routes:
        probes += 1
        resp, elapsed = _probe(client, (method, path), representative,
                               resolver, scratch, targets_for(representative),
                               probe_timeout)
        if resp is None:
            continue
        status = resp.status_code
        note = ""
        if status in (404, 405):
            note = "  not present"
        elif status == 0:
            note = "  unreachable"
        else:
            alive.append((method, path))
        LOG.out(f"    {method:<6}{path:<26}{status:>5}  {elapsed:5.1f}s{note}")
        if resp.ok:
            targets = targets_for(representative)
            stray = scope_check(client, scratch, targets)
            if stray:
                LOG.out(f"      REJECTED despite {status}: the write also "
                        f"tagged {len(stray)} API(s) outside the "
                        f"{len(targets)} target(s).")
                LOG.out(f"      e.g. {', '.join(stray[:3])}")
                LOG.out("      Either the selector was ignored, or a probe tag "
                        "from an earlier\n      run is still attached. "
                        "--delete-tags zz-shape-probe --apply clears the "
                        "latter.")
                continue
            if stray is None:
                LOG.out("      (blast radius unverified; proceeding)")
            LOG.out(f"\n  Adopted: {method} {path} with "
                    f"{representative[0]}")
            cleanup_probe_tag(client, scratch)
            return representative, representative[1], (method, path), \
                representative[2]

    if not alive:
        LOG.out("\n  No tag-write route responded. Nothing was written.")
        return None, None, None, False

    # ---- phase 2: sweep bodies on the routes that answered ----
    LOG.out(f"\n  Phase 2 - body shapes on {len(alive)} live route(s)")
    for route in alive:
        method, path = route
        LOG.out(f"\n    {method} {path}")
        errors = Counter()
        for shape in BULK_SHAPES:
            if shape is representative:
                continue
            targets = targets_for(shape)
            if not targets:
                continue
            if probes >= max_probes:
                LOG.out(f"\n  Probe budget of {max_probes} reached. Raise it "
                        "with --max-probes if you want a longer sweep.")
                return None, None, None, False
            probes += 1
            resp, elapsed = _probe(client, route, shape, resolver, scratch,
                                   targets, probe_timeout)
            if resp is None:
                continue
            LOG.out(f"      [{probes:>2}/{max_probes}] {shape[0]:<28}"
                    f"{resp.status_code:>5}  {elapsed:5.1f}s")
            if not resp.ok:
                detail = error_detail(resp, 200)
                LOG.verbose(f"          {detail}")
                errors[detail] += 1
                if errors[detail] >= 3:
                    LOG.out(f"      same error 3x; next route")
                    LOG.out(f"          {detail}")
                    break
                continue

            stray = scope_check(client, scratch, targets)
            if stray:
                LOG.out(f"          REJECTED: also tagged {len(stray)} API(s) "
                        f"outside the {len(targets)} target(s)")
                continue
            LOG.out(f"\n  Adopted: {method} {path} with {shape[0]} "
                    f"(selects by {shape[1]})")
            cleanup_probe_tag(client, scratch)
            return shape, shape[1], route, shape[2]

    LOG.out(f"\n  No route/body combination accepted ({probes} probes).")
    return None, None, None, False


def apply_tags(client, by_tag, by_tag_hosts, replace, limit, ledger,
               allow_single=False, max_failures=MAX_CONSECUTIVE_FAILURES):
    print("\n" + "=" * 72)
    print("APPLYING TAGS" + (" (REPLACE mode)" if replace else " (append mode)"))
    print("=" * 72)

    resolver = TagResolver(client)
    applied = failed = 0
    budget = limit if limit else None
    shape = selector = route = None
    needs_ids = False
    consecutive = 0

    for tag in sorted(by_tag, key=lambda t: -len(by_tag[t])):
        api_ids = by_tag[tag]
        hosts = sorted(by_tag_hosts.get(tag, []))
        if budget is not None:
            if budget <= 0:
                break
            api_ids = api_ids[:budget]

        if shape is None:
            probe_hosts = hosts[:1]
            expected = list(by_tag[tag]) if probe_hosts else api_ids[:2]
            shape, selector, route, needs_ids = negotiate_bulk_shape(
                client, resolver, api_ids[:2], probe_hosts, expected, replace,
                max_probes=PROBE_BUDGET[0], probe_timeout=PROBE_BUDGET[1])
            if shape is None:
                LOG.out("\n  No route/body combination was accepted, so no "
                        "tags can be written.")
                LOG.out("  Capture the full matrix for support with:")
                LOG.out("      --discover-tag-api --sample 8 "
                        "--log-file tag-endpoint.log")
                return applied, failed
            ledger["shape"] = shape[0]
            ledger["selector"] = selector
            ledger["route"] = f"{route[0]} {route[1]}"

        if selector == "hosts":
            units = hosts
            batch_size = HOST_BATCH_SIZE
            unit_label = "host"
        else:
            units = api_ids
            batch_size = TAG_BATCH_SIZE
            unit_label = "API"

        if not units:
            continue

        print(f"\n  tag '{tag}' -> {len(units):,} {unit_label}(s) "
              f"({len(api_ids):,} API(s))")
        values = resolver.ids_for([tag]) if needs_ids else [tag]
        if values is None:
            print(f"    ! could not resolve a catalogue ID for '{tag}'")
            failed += len(api_ids)
            continue

        starts = list(range(0, len(units), batch_size))
        for n, start in enumerate(starts, 1):
            batch = units[start:start + batch_size]
            resp = client.write_tags(batch, values, shape, route)

            if resp.ok:
                consecutive = 0
                applied += len(batch)
                if budget is not None:
                    budget -= len(batch)
                ledger["tags"].append({"tag": tag, "selector": selector,
                                       "targets": batch,
                                       "mode": f"bulk:{shape[0]}"})
                print(f"    batch {n}/{len(starts)}: {len(batch)} OK "
                      f"({resp.status_code})")
                continue

            consecutive += 1
            failed += len(batch)
            ledger["failed"].append({"tag": tag, "targets": batch,
                                     "status": resp.status_code,
                                     "body": resp.text[:300]})
            print(f"    batch {n}/{len(starts)}: HTTP {resp.status_code} "
                  f"{error_detail(resp)}")

            if consecutive >= max_failures:
                print(f"\n  ! {consecutive} consecutive failures -- stopping.")
                print("  Re-run after fixing: already-tagged APIs are skipped.")
                print(f"\n  Tag assignments applied: {applied:,}   "
                      f"failed: {failed:,}")
                return applied, failed

    print(f"\n  Tag write batches applied: {applied:,}   failed: {failed:,}")
    return applied, failed


def discover_tag_api(client, apis, probe_tag, sample=1):
    """Probe candidate bodies, and probe across several APIs.

    Two different questions get answered here. First, which request body does
    the endpoint accept -- a 4xx tells you the shape is wrong. Second, is a
    5xx universal or record-specific: the same write against APIs that differ
    in whether they already carry tags separates a broken endpoint from
    broken data on particular records.
    """
    print("\n" + "=" * 72)
    print("TAG API DISCOVERY")
    print("=" * 72)

    resolver = TagResolver(client)
    if not resolver.available:
        print(f"  No tag catalogue at {EP_TAGS_CATALOG} or "
              f"{EP_TAGS_CATALOG_V3}")

    tagged = [a for a in apis if a.get("tags")]
    untagged = [a for a in apis if not a.get("tags")]
    picks, seen = [], set()
    for pool in (untagged, tagged):
        for api in pool:
            if len(picks) >= max(sample, 1):
                break
            aid = str(api.get("id"))
            if aid not in seen:
                seen.add(aid)
                picks.append(api)

    print(f"  Probe tag : {probe_tag}")
    print(f"  Sample    : {len(picks)} API(s) -- "
          f"{sum(1 for a in picks if a.get('tags'))} already tagged, "
          f"{sum(1 for a in picks if not a.get('tags'))} untagged")
    print("  This writes a throwaway tag. Remove it afterwards with "
          "--cleanup-scratch.")

    first = str(picks[0].get("id"))

    print(f"\n  Request bodies, against API {first}")
    print(f"  {'':4}{'SHAPE':<26}{'STATUS':<8}")
    accepted = []
    first_host = str(picks[0].get("host") or "")
    for route in TAG_WRITE_ROUTES:
        method, path = route
        LOG.out(f"\n    {method} {path}")
        seen = Counter()
        for shape in BULK_SHAPES:
            name, selector, needs_ids, _ = shape
            target = [first_host] if selector == "hosts" else [first]
            if not target[0]:
                continue
            values = resolver.ids_for([probe_tag]) if needs_ids else [probe_tag]
            if values is None:
                continue
            resp = client.write_tags(target, values, shape, route, retries=0)
            mark = "OK " if resp.ok else "   "
            LOG.out(f"      {mark}{name:<28}{resp.status_code}")
            if resp.ok:
                accepted.append((route, shape))
            else:
                detail = error_detail(resp, 300)
                seen[detail] += 1
                if seen[detail] == 1:
                    LOG.out(f"          {detail}")
                if resp.status_code == 404:
                    LOG.out("          route not present; skipping its "
                            "remaining bodies")
                    break
                if seen[detail] >= 3:
                    LOG.out("          same error 3x; skipping remaining "
                            "bodies on this route")
                    break

    print(f"\n  Single endpoint  PATCH {EP_TAGS_ONE.format(id='<id>')}")
    for shape in SINGLE_SHAPES:
        name, needs_ids, _ = shape
        values = resolver.ids_for([probe_tag]) if needs_ids else [probe_tag]
        if values is None:
            print(f"    {name:<26}skipped (needs tag catalogue)")
            continue
        resp = client.write_tags_single(first, values, shape, retries=0)
        mark = "OK  " if resp.ok else "    "
        print(f"    {mark}{name:<26}{resp.status_code}")
        if not resp.ok:
            print(f"        {error_detail(resp, 400)}")

    # Same write, different records.
    probe_route, probe_shape = (accepted[0] if accepted
                                else (TAG_WRITE_ROUTES[0], BULK_SHAPES[0]))
    if len(picks) > 1:
        LOG.out(f"\n  {probe_route[0]} {probe_route[1]} with "
                f"{probe_shape[0]}, across {len(picks)} different APIs:")
        print(f"    {'API':<26}{'TAGGED?':<10}{'STATUS':<8}")
        outcomes = Counter()
        for api in picks:
            aid = str(api.get("id"))
            values = (resolver.ids_for([probe_tag]) if probe_shape[2]
                      else [probe_tag])
            target = ([str(api.get("host") or "")] if probe_shape[1] == "hosts"
                      else [aid])
            resp = client.write_tags(target, values, probe_shape, probe_route,
                                     retries=0)
            outcomes[resp.status_code] += 1
            has = "yes" if api.get("tags") else "no"
            print(f"    {aid[:26]:<26}{has:<10}{resp.status_code}")
        print(f"\n    outcomes: "
              f"{', '.join(f'{v}x HTTP {k}' for k, v in outcomes.most_common())}")
        if len(outcomes) == 1 and next(iter(outcomes)) >= 500:
            print("    Every record fails identically -- this is the endpoint, "
                  "not the data.")
            print("    Worth raising with the product team with this output "
                  "attached.")
        elif len(outcomes) > 1:
            print("    Mixed results -- the failure is record-specific, so the "
                  "run can")
            print("    proceed and skip the records that fault.")

    print("\n  Send me this output and I will pin the script to whatever "
          "worked.")


def cleanup_scratch(client, name=SCRATCH_TAG):
    """Remove probe tags left in the catalogue by earlier negotiation runs."""
    resolver = TagResolver(client)
    if not resolver.available:
        print("No tag catalogue reachable; nothing to clean up.")
        return
    targets = {n: i for n, i in resolver.catalog.items()
               if n == name or n.startswith(name)}
    if not targets:
        print(f"No '{name}*' tags in the catalogue.")
        return
    print(f"Deleting {len(targets)} scratch tag(s):")
    for tag_name, tag_id in targets.items():
        resp = client.request("DELETE", f"{resolver.endpoint}/{tag_id}",
                              max_retries=0)
        state = "removed" if resp.ok else f"HTTP {resp.status_code}"
        print(f"  {tag_name:<32} {state}")
    print("\nIf DELETE is unsupported, remove them from the portal instead.")


def create_groups(client, by_tag, existing_names, group_type, prefix,
                  parent_name, description, ledger, name_from="value",
                  known_prefixes=()):
    print("\n" + "=" * 72)
    print("CREATING GROUPS")
    print("=" * 72)

    parent_id = None
    if parent_name:
        safe_parent = sanitize_group_name(parent_name)
        resp = client.create_group({
            "name": safe_parent,
            "description": "Container for tag-derived groups.",
            "group_type": group_type,
            "filters": [{"field": "tags", "operator": "in",
                         "value": sorted(by_tag.keys())}],
        })
        if resp.ok:
            body = resp.json() if resp.content else {}
            parent_id = body.get("id") or body.get("groupId")
            ledger["groups"].append({"name": safe_parent, "id": parent_id})
            LOG.out(f'  Parent "{safe_parent}" created (id={parent_id})')
        elif resp.status_code == 409:
            print(f'  Parent "{safe_parent}" already exists -- nesting skipped')
        else:
            print(f"  ! Parent creation failed: HTTP {resp.status_code} "
                  f"{resp.text[:200]} -- continuing without nesting")

    operators = list(GROUP_TAG_OPERATORS)
    created = skipped = failed = 0
    for tag in sorted(by_tag):
        name = sanitize_group_name(group_label(tag, name_from, known_prefixes),
                                   prefix)
        if name in existing_names:
            skipped += 1
            LOG.out(f'  = "{name}" already exists')
            continue

        body = {
            "name": name,
            "description": description or f"APIs tagged {tag}",
            "group_type": group_type,
            "filters": [tag_filter(tag, operators[0])],
        }
        if parent_id:
            body["parentGroupId"] = parent_id

        resp = client.create_group(body)

        # The accepted operator differs by deployment; learn it once.
        if not resp.ok and resp.status_code == 400 and len(operators) > 1:
            detail = error_detail(resp, 200)
            if "operator" in detail.lower():
                for alt in operators[1:]:
                    LOG.out(f'    operator "{operators[0]}" rejected; '
                            f'retrying with "{alt}"')
                    LOG.verbose(f"      {detail}")
                    body["filters"] = [tag_filter(tag, alt)]
                    resp = client.create_group(body)
                    if resp.ok or resp.status_code == 409:
                        operators[:] = [alt] + [o for o in operators if o != alt]
                        break
        if resp.ok:
            created += 1
            payload = resp.json() if resp.content else {}
            gid = payload.get("id") or payload.get("groupId")
            ledger["groups"].append({"name": name, "id": gid, "tag": tag})
            LOG.out(f'  + "{name}"  <- tags {operators[0]} "{tag}"  '
                    f'(id={gid})')
        elif resp.status_code == 409:
            skipped += 1
            print(f'  = "{name}" already exists (409)')
        else:
            failed += 1
            print(f'  ! "{name}": HTTP {resp.status_code} {resp.text[:200]}')

    print(f"\n  Groups created: {created}   skipped: {skipped}   failed: {failed}")
    return created, skipped, failed


def create_groups_from_hosts(client, tag_hosts, existing_names, group_type,
                             prefix, parent_name, description, ledger,
                             split_at=None, dry_run=False):
    """One group per tag value, selecting APIs by hostname filter directly.

    The tag write endpoint is not involved. Group filters already accept
    `host`, so the department scoping the customer actually wants can be
    built straight from the register even while tagging is broken.
    """
    print("\n" + "=" * 72)
    print("GROUPS FROM HOST FILTERS" + (" (dry run)" if dry_run else ""))
    print("=" * 72)
    print("  Filters select on `host` directly, so no tags are required.")

    parent_id = None
    if parent_name and not dry_run:
        safe_parent = sanitize_group_name(parent_name)
        resp = client.create_group({
            "name": safe_parent,
            "description": "Container for register-derived groups.",
            "group_type": group_type,
            "filters": [{"field": "host", "operator": "in",
                         "value": sorted({h for hs in tag_hosts.values()
                                          for h in hs})[:1000]}],
        })
        if resp.ok:
            body = resp.json() if resp.content else {}
            parent_id = body.get("id") or body.get("groupId")
            ledger["groups"].append({"name": safe_parent, "id": parent_id})
            print(f'  Parent "{safe_parent}" created (id={parent_id})')
        elif resp.status_code == 409:
            print(f'  Parent "{safe_parent}" already exists')
        else:
            print(f"  ! Parent creation failed: HTTP {resp.status_code} "
                  f"{error_detail(resp)} -- continuing unnested")

    created = skipped = failed = 0
    for tag in sorted(tag_hosts):
        hosts = sorted(tag_hosts[tag])
        if not hosts:
            continue

        chunks = ([hosts] if not split_at
                  else [hosts[i:i + split_at]
                        for i in range(0, len(hosts), split_at)])

        for idx, chunk in enumerate(chunks, 1):
            name = sanitize_group_name(tag, prefix)
            if len(chunks) > 1:
                name = sanitize_group_name(f"{tag} ({idx}/{len(chunks)})", prefix)

            if name in existing_names:
                skipped += 1
                print(f'  = "{name}" already exists')
                continue

            body = {
                "name": name,
                "description": description or
                               f"Auto-created from register: {tag}",
                "group_type": group_type,
                "filters": [{"field": "host", "operator": "in",
                             "value": chunk}],
            }
            if parent_id:
                body["parentGroupId"] = parent_id

            if dry_run:
                print(f'  + "{name}"  <- host in [{len(chunk)} hostname(s)]')
                created += 1
                continue

            LOG.debug(f"group body {json.dumps(body)[:MAX_LOG_BODY]}")
            resp = client.create_group(body)
            if resp.ok:
                created += 1
                payload = resp.json() if resp.content else {}
                gid = payload.get("id") or payload.get("groupId")
                ledger["groups"].append({"name": name, "id": gid, "tag": tag,
                                         "hosts": len(chunk)})
                print(f'  + "{name}" ({len(chunk)} host(s), id={gid})')
            elif resp.status_code == 409:
                skipped += 1
                print(f'  = "{name}" already exists (409)')
            else:
                failed += 1
                print(f'  ! "{name}": HTTP {resp.status_code} '
                      f'{error_detail(resp)}')
                if len(chunk) > 50 and not split_at:
                    print("      That filter carries "
                          f"{len(chunk)} hostnames. If the tenant caps filter "
                          "size,\n      retry with --split-groups-at 100.")

    verb = "would create" if dry_run else "created"
    print(f"\n  Groups {verb}: {created}   skipped: {skipped}   "
          f"failed: {failed}")
    return created, skipped, failed


def inspect_tags(client, limit=5):
    """Show the actual shape of existing tag objects.

    The decisive question is whether a tag carries a filter of its own. If it
    does, tags are rule-based and there is no per-API assignment call to make
    work -- the tag is created with terms and matching APIs pick it up.
    """
    LOG.out("\n" + "=" * 72)
    LOG.out("TAG SCHEMA INSPECTION")
    LOG.out("=" * 72)

    for endpoint in (EP_TAGS_CATALOG, EP_TAGS_CATALOG_V3):
        resp = client.safe_request("GET", endpoint, max_retries=1)
        LOG.out(f"\n  GET {endpoint} -> HTTP {resp.status_code}")
        if not resp.ok:
            LOG.out(f"      {error_detail(resp)}")
            continue
        try:
            body = resp.json()
        except ValueError:
            LOG.out("      non-JSON response")
            continue

        nodes = []

        def walk(obj):
            if isinstance(obj, list):
                for item in obj:
                    walk(item)
            elif isinstance(obj, dict):
                if obj.get("name") and any(obj.get(k) not in (None, "")
                                           for k in ID_KEYS):
                    nodes.append(obj)
                for value in obj.values():
                    if isinstance(value, (dict, list)):
                        walk(value)

        walk(body)
        if not nodes:
            LOG.out(f"      no tag objects found; top level "
                    f"{list(body) if isinstance(body, dict) else type(body).__name__}")
            continue

        keys = sorted({k for n in nodes for k in n})
        LOG.out(f"      {len(nodes)} tag(s); fields present across them:")
        LOG.out(f"      {', '.join(keys)}")

        rule_keys = [k for k in keys
                     if k.lower() in ("terms", "filters", "rules", "conditions",
                                      "query", "criteria", "expression")]
        if rule_keys:
            LOG.out(f"\n      RULE-BEARING FIELD(S): {', '.join(rule_keys)}")
            LOG.out("      Tags are rule-based -- the tag carries its own "
                    "filter, so there is")
            LOG.out("      no per-API assignment call. Create the tag WITH "
                    "terms instead.")
        else:
            LOG.out("\n      No rule-bearing field on these tags, so they "
                    "look statically")
            LOG.out("      assigned and the association endpoint is the "
                    "blocker.")

        LOG.out(f"\n      First {min(limit, len(nodes))} tag object(s) in "
                "full:")
        for node in nodes[:limit]:
            LOG.out("        " + json.dumps(node)[:MAX_LOG_BODY])

        tid = next((node[k] for node in nodes for k in ID_KEYS
                    if node.get(k) not in (None, "")), None)
        if tid:
            detail = client.safe_request("GET", f"{endpoint}/{tid}",
                                         max_retries=1)
            LOG.out(f"\n  GET {endpoint}/<id> -> HTTP {detail.status_code}")
            if detail.ok:
                LOG.out("        " + detail.text[:MAX_LOG_BODY])
            else:
                LOG.out(f"      {error_detail(detail)}")
        return

    LOG.out("\n  No tag catalogue reachable.")


def create_rule_tags(client, tag_hosts, group_type, ledger, dry_run=False,
                     split_at=None):
    """Create one rule-based tag per value, each carrying a host filter."""
    LOG.out("\n" + "=" * 72)
    LOG.out("CREATING RULE-BASED TAGS" + (" (dry run)" if dry_run else ""))
    LOG.out("=" * 72)

    resolver = TagResolver(client)
    if not resolver.available:
        LOG.out("  No tag catalogue endpoint; cannot create tags.")
        return {}, None

    shape = None
    created = {}
    for tag in sorted(tag_hosts):
        hosts = sorted(tag_hosts[tag])
        if not hosts:
            continue
        if split_at:
            hosts = hosts[:split_at]
        terms = [{"field": "host", "operator": "in", "value": hosts}]

        if tag in resolver.catalog:
            LOG.out(f'  = "{tag}" already in catalogue '
                    f"(id={resolver.catalog[tag]})")
            created[tag] = resolver.catalog[tag]
            continue

        if dry_run:
            LOG.out(f'  + "{tag}"  <- terms: host in [{len(hosts)} hostname(s)]')
            created[tag] = "<dry-run>"
            continue

        candidates = [shape] if shape else TAG_CREATE_SHAPES
        for candidate in candidates:
            label, build = candidate
            body = build(tag, terms)
            LOG.debug(f"tag body {json.dumps(body)[:MAX_LOG_BODY]}")
            resp = client.safe_request("POST", resolver.endpoint,
                                       max_retries=1, data=json.dumps(body))
            if resp.ok:
                payload = resp.json() if resp.content else {}
                tid = next((payload[k] for k in ID_KEYS
                            if isinstance(payload, dict)
                            and payload.get(k) not in (None, "")), None)
                if shape is None:
                    shape = candidate
                    LOG.out(f"  Adopted tag body: {label}")
                created[tag] = str(tid) if tid else ""
                LOG.out(f'  + "{tag}" ({len(hosts)} host(s), id={tid})')
                break
            LOG.out(f'    {label:<22} -> HTTP {resp.status_code}  '
                    f'{error_detail(resp, 160)}')
        else:
            LOG.out(f'  ! "{tag}": no tag body accepted')

    if not dry_run:
        ledger["rule_tags"] = created
    LOG.out(f"\n  Tags: {len(created)}")
    return created, shape


# Never deleted by prefix matching, whatever the prefix.
PROTECTED_GROUP_NAMES = {"root", "all apis", "default", "unassigned"}
MIN_PURGE_PREFIX = 2


def _confirm(count, noun, yes):
    if yes:
        return True
    LOG.out(f"\n  About to permanently delete {count} {noun}. "
            "This cannot be undone.")
    answer = input(f"  Type the number {count} to confirm: ").strip()
    if answer != str(count):
        LOG.out("  Aborted.")
        return False
    return True


def delete_groups_by_prefix(client, prefix, apply_mode, yes, ledger,
                            ignore_case=False):
    """Delete groups whose name starts with `prefix`, children first."""
    LOG.out("\n" + "=" * 72)
    LOG.out("DELETE GROUPS" + ("" if apply_mode else " (dry run)"))
    LOG.out("=" * 72)

    records = client.fetch_group_records()
    LOG.out(f"  {len(records):,} group(s) in tenant")
    needle = prefix.lower() if ignore_case else prefix
    matches, protected = [], []
    for rec in records:
        name = rec["name"].lower() if ignore_case else rec["name"]
        if not name.startswith(needle):
            continue
        if rec["name"].strip().lower() in PROTECTED_GROUP_NAMES:
            protected.append(rec)
        else:
            matches.append(rec)

    # Deepest first so a child never outlives its parent.
    matches.sort(key=lambda r: -r["depth"])

    LOG.out(f'  matching "{prefix}": {len(matches):,}'
            + (f"  (protected, skipped: {len(protected)})" if protected else ""))
    for rec in matches[:30]:
        LOG.out(f'    {"  " * rec["depth"]}"{rec["name"]}"  id={rec["id"]}')
    if len(matches) > 30:
        LOG.out(f"    ... and {len(matches) - 30} more")

    if not matches:
        return 0
    if not apply_mode:
        LOG.out("\n  Re-run with --apply to delete these.")
        return 0
    if not _confirm(len(matches), "group(s)", yes):
        return 0

    deleted = failed = 0
    for rec in matches:
        resp = client.safe_request("DELETE", EP_GROUP_ID.format(id=rec["id"]),
                                   max_retries=1)
        if resp.ok or resp.status_code == 404:
            deleted += 1
            ledger.setdefault("deleted_groups", []).append(rec)
            LOG.out(f'  - "{rec["name"]}" removed')
        else:
            failed += 1
            LOG.out(f'  ! "{rec["name"]}": HTTP {resp.status_code} '
                    f"{error_detail(resp, 160)}")
    LOG.out(f"\n  Groups deleted: {deleted}   failed: {failed}")
    return deleted


def delete_tags_by_prefix(client, prefix, apply_mode, yes, ledger,
                          ignore_case=False):
    """Delete catalogue tags whose name starts with `prefix`.

    Groups filtering on a deleted tag are reported first -- they survive but
    stop matching, which is worse than an obvious failure.
    """
    LOG.out("\n" + "=" * 72)
    LOG.out("DELETE TAGS" + ("" if apply_mode else " (dry run)"))
    LOG.out("=" * 72)

    resolver = TagResolver(client)
    if not resolver.available:
        LOG.out("  No tag catalogue reachable; nothing to delete.")
        return 0

    needle = prefix.lower() if ignore_case else prefix
    matches = {n: i for n, i in resolver.catalog.items()
               if (n.lower() if ignore_case else n).startswith(needle)}
    LOG.out(f'  matching "{prefix}": {len(matches):,} of '
            f"{len(resolver.catalog):,} tag(s)")
    for name in sorted(matches)[:30]:
        LOG.out(f'    "{name}"  id={matches[name]}')
    if len(matches) > 30:
        LOG.out(f"    ... and {len(matches) - 30} more")

    if not matches:
        return 0

    # Warn about groups that would be left filtering on a tag that is gone.
    dependents = []
    for rec in client.fetch_group_records():
        for flt in (rec.get("filters") or []):
            if not isinstance(flt, dict) or flt.get("field") != "tags":
                continue
            value = flt.get("value")
            values = value if isinstance(value, list) else [value]
            if any(v in matches for v in values if v):
                dependents.append(rec["name"])
    if dependents:
        LOG.out(f"\n  ! {len(dependents)} group(s) filter on these tags and "
                "would stop matching:")
        for name in sorted(set(dependents))[:10]:
            LOG.out(f'      "{name}"')
        LOG.out("    Delete those groups too, or repoint them first.")

    if not apply_mode:
        LOG.out("\n  Re-run with --apply to delete these.")
        return 0
    if not _confirm(len(matches), "tag(s)", yes):
        return 0

    deleted = failed = 0
    for name in sorted(matches):
        resp = client.safe_request("DELETE",
                                   f"{resolver.endpoint}/{matches[name]}",
                                   max_retries=1)
        if resp.ok or resp.status_code == 404:
            deleted += 1
            ledger.setdefault("deleted_tags", []).append(
                {"name": name, "id": matches[name]})
            LOG.out(f'  - "{name}" removed')
        else:
            failed += 1
            LOG.out(f'  ! "{name}": HTTP {resp.status_code} '
                    f"{error_detail(resp, 160)}")
    LOG.out(f"\n  Tags deleted: {deleted}   failed: {failed}")
    if failed:
        LOG.out("  If DELETE is unsupported here, remove them in the portal.")
    return deleted


def purge(client, prefix, apply_mode, yes, ignore_case, do_groups, do_tags):
    """Remove groups then tags matching a prefix, in that order.

    Groups go first: a group filtering on a tag that no longer exists is a
    silent no-match, whereas an orphaned tag is merely untidy.
    """
    if len(prefix.strip()) < MIN_PURGE_PREFIX:
        sys.exit(f"Refusing to match on a prefix shorter than "
                 f"{MIN_PURGE_PREFIX} characters -- too easy to delete "
                 "everything by accident.")

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    ledger = {"started": stamp, "base": client.base, "prefix": prefix}
    try:
        if do_groups:
            delete_groups_by_prefix(client, prefix, apply_mode, yes, ledger,
                                    ignore_case)
        if do_tags:
            delete_tags_by_prefix(client, prefix, apply_mode, yes, ledger,
                                  ignore_case)
    finally:
        if apply_mode and (ledger.get("deleted_groups")
                           or ledger.get("deleted_tags")):
            path = f"purge-{stamp}.ledger.json"
            with open(path, "w", encoding="utf-8") as fh:
                json.dump(ledger, fh, indent=2)
            LOG.out(f"\nLedger: {path}")
            LOG.out("  It records every deleted object's name, id and filters "
                    "for reconstruction.")


def rollback(client, ledger_path):
    with open(ledger_path, encoding="utf-8") as fh:
        ledger = json.load(fh)

    groups = [g for g in ledger.get("groups", []) if g.get("id")]
    if not groups:
        print("Ledger contains no created groups with IDs -- nothing to roll back.")
        print("Note: tag assignments are NOT auto-reverted. The ledger lists every "
              "tag/API pair written, so they can be removed deliberately.")
        return

    ordered = sorted(groups, key=lambda g: 0 if g.get("tag") else 1)
    ok = bad = 0
    print(f"Deleting {len(ordered)} group(s) from {ledger_path}")
    for group in ordered:
        resp = client.delete_group(group["id"])
        if resp.ok or resp.status_code == 404:
            ok += 1
            print(f'  - "{group["name"]}" removed')
        else:
            bad += 1
            print(f'  ! "{group["name"]}": HTTP {resp.status_code}')
    print(f"\nRolled back {ok} group(s); {bad} failure(s).")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def resolve_credentials(args):
    base = args.base or os.environ.get("NONAME_API_BASE")
    if not base:
        base = input("Tenant base URL (https://<tenant>): ").strip()
    if not base.startswith(("http://", "https://")):
        base = "https://" + base

    token = os.environ.get("NONAME_API_TOKEN")
    if not token:
        token = getpass("Management API token (input hidden): ").strip()
    if not token:
        sys.exit("No token supplied.")
    return base, token


def main():
    ap = argparse.ArgumentParser(
        description="Apply spreadsheet-defined tags to APIs, with dry run and "
                    "optional per-tag group creation.",
        formatter_class=argparse.RawDescriptionHelpFormatter)

    ap.add_argument("--sheet", help="Path to the .xlsx / .csv / .tsv source")
    ap.add_argument("--tab", action="append", default=[],
                    help="Worksheet to read. Repeatable -- ORDER SETS PRIORITY, "
                         "the first --tab wins any conflict. Default: first sheet.")
    ap.add_argument("--tag-from", action="append", default=[],
                    metavar="COLUMN[:PREFIX]",
                    help="Promote a column to a tag dimension, e.g. "
                         "--tag-from 'Department:dept'. Repeatable.")
    ap.add_argument("--no-prefix", action="store_true",
                    help="Emit bare tag values instead of 'prefix:value'")
    ap.add_argument("--host-column", help="Override hostname column detection")
    ap.add_argument("--id-column", help="Override API-id column detection")
    ap.add_argument("--path-column", help="Override path column detection")
    ap.add_argument("--sentinel", action="append", default=[],
                    help="Extra placeholder value to treat as blank. Repeatable.")

    ap.add_argument("--base", help="Tenant base URL (else env NONAME_API_BASE)")
    ap.add_argument("--insecure", action="store_true",
                    help="Skip TLS verification (lab tenants only)")
    ap.add_argument("--timeout", type=int, default=60,
                    help="Per-request timeout in seconds (default 60)")
    ap.add_argument("--page-size", type=int, default=INVENTORY_PAGE_SIZE,
                    help=f"Inventory page size (default {INVENTORY_PAGE_SIZE}). "
                         "Lower it if large pages time out.")
    ap.add_argument("--inventory-cache", metavar="PATH",
                    help="Cache the fetched inventory here and reuse it on the "
                         "next run, so a timeout does not mean re-paging "
                         "everything.")
    ap.add_argument("--refresh-inventory", action="store_true",
                    help="Ignore any existing --inventory-cache and re-fetch")
    ap.add_argument("--cache-max-age", type=float, default=24.0,
                    help="Hours before a cached inventory is considered stale "
                         "(default 24)")
    ap.add_argument("--return-fields",
                    help="Comma-separated inventory returnFields override. "
                         f"Default: {','.join(RETURN_FIELDS)}")

    ap.add_argument("--list-sheets", action="store_true",
                    help="List every worksheet with headers and row counts, "
                         "then exit. No network calls.")
    ap.add_argument("--inspect-sheet", action="store_true",
                    help="Show column detection, suggested tag dimensions and "
                         "parsed rules, then exit. No network calls.")
    ap.add_argument("--probe", action="store_true",
                    help="Read-only endpoint check, then exit")
    ap.add_argument("--apply", action="store_true",
                    help="Actually write. Without this the run is a dry run.")
    ap.add_argument("--replace", action="store_true",
                    help="PUT (replace all tags) instead of PATCH (append). "
                         "Destructive -- existing tags on matched APIs are lost.")
    ap.add_argument("--limit", type=int,
                    help="Cap total tag assignments written (canary runs)")

    ap.add_argument("--create-groups", action="store_true",
                    help="Also create one group per distinct tag")
    ap.add_argument("--delete-tags", metavar="PREFIX",
                    help="Delete catalogue tags whose name starts with PREFIX. "
                         "Dry run unless --apply.")
    ap.add_argument("--delete-groups", metavar="PREFIX",
                    help="Delete groups whose name starts with PREFIX, "
                         "children first. Dry run unless --apply.")
    ap.add_argument("--purge", metavar="PREFIX",
                    help="Delete both groups and tags matching PREFIX -- "
                         "groups first. Dry run unless --apply. Use this to "
                         "reset before a clean run.")
    ap.add_argument("--purge-ignore-case", action="store_true",
                    help="Case-insensitive prefix match for the delete options")
    ap.add_argument("--yes", action="store_true",
                    help="Skip the interactive delete confirmation")
    ap.add_argument("--inspect-tags", action="store_true",
                    help="Show the real schema of existing tags -- whether "
                         "they carry their own filter -- then exit. Read-only.")
    ap.add_argument("--rule-tags", action="store_true",
                    help=argparse.SUPPRESS)
    ap.add_argument("--groups-from-hosts", action="store_true",
                    help="Build one group per tag value filtered on `host` "
                         "directly, writing no tags at all. Use when the tag "
                         "endpoint is unavailable.")
    ap.add_argument("--split-groups-at", type=int, metavar="N",
                    help="Split a group into numbered parts of at most N "
                         "hostnames, if the tenant caps filter size")
    ap.add_argument("--groups-only", action="store_true",
                    help="Create groups only; write no tags. Use when the tags "
                         "are already applied.")
    ap.add_argument("--discover-tag-api", action="store_true",
                    help="Probe every candidate tag request body against ONE "
                         "API and report which the tenant accepts, then exit. "
                         "Writes a throwaway tag on that one API.")
    ap.add_argument("--probe-tag", default=SCRATCH_PREFIX,
                    help=f"Tag name used for probing (default: {SCRATCH_TAG})")
    ap.add_argument("--sample", type=int, default=1,
                    help="How many different APIs --discover-tag-api should "
                         "probe, mixing tagged and untagged records "
                         "(default 1)")
    ap.add_argument("--cleanup-scratch", action="store_true",
                    help="Delete probe tags left in the catalogue, then exit")
    ap.add_argument("--allow-single-fallback", action="store_true",
                    help="On a failed batch of <=50, retry those APIs one at a "
                         "time. Off by default: at scale it is tens of "
                         "thousands of requests.")
    ap.add_argument("--probe-timeout", type=int, default=20,
                    help="Per-probe timeout in seconds during negotiation "
                         "(default 20). Probes should fail fast.")
    ap.add_argument("--max-probes", type=int, default=40,
                    help="Cap on negotiation probes (default 40)")
    ap.add_argument("--max-failures", type=int, default=MAX_CONSECUTIVE_FAILURES,
                    help="Abort after this many consecutive failed batches "
                         f"(default {MAX_CONSECUTIVE_FAILURES})")
    ap.add_argument("-v", "--verbose", action="count", default=0,
                    help="More detail. Repeat for full HTTP tracing "
                         "(-vv is the same as --debug).")
    ap.add_argument("--debug", action="store_true",
                    help="Full HTTP tracing: URLs, params, request and "
                         "response bodies, timings. Credentials are redacted.")
    ap.add_argument("--log-file", metavar="PATH",
                    help="Append a full debug transcript here regardless of "
                         "console level. Redacted, so it is safe to attach to "
                         "a support ticket.")
    ap.add_argument("--dump-raw", action="store_true",
                    help="Print the raw group response before parsing "
                         "(diagnostic for unexpected payload shapes)")
    ap.add_argument("--group-type", default="APPLICATION",
                    choices=["APPLICATION", "OTHER"])
    ap.add_argument("--group-name-from", choices=["value", "tag"],
                    default="value",
                    help="Group name source: 'value' names the group after the "
                         "department (dept:CPP -> \"CPP\"), 'tag' uses the tag "
                         "verbatim. The filter always matches the full tag.")
    ap.add_argument("--group-prefix", default="",
                    help='Prefix for generated group names, e.g. "Tag: "')
    ap.add_argument("--group-parent",
                    help="Create a parent container group and nest under it")
    ap.add_argument("--group-description",
                    help="Description applied to every generated group")

    ap.add_argument("--plan-out", help="Write the resolved plan to this JSON path")
    ap.add_argument("--rollback", metavar="LEDGER",
                    help="Delete the groups recorded in a prior run's ledger")

    args = ap.parse_args()

    level = Log.DEBUG if args.debug else min(args.verbose, Log.DEBUG)
    LOG.configure(level, args.log_file, sys.argv,
                  args.base or os.environ.get("NONAME_API_BASE"))
    if level >= Log.VERBOSE:
        LOG.out(f"Verbosity level {level} "
                f"({'debug' if level >= Log.DEBUG else 'verbose'})")
    if args.log_file:
        LOG.out(f"Transcript: {args.log_file}")

    if args.rollback:
        base, token = resolve_credentials(args)
        client = Client(base, token, timeout=args.timeout,
                        verify=not args.insecure,
                        page_size=args.page_size)
        client.preflight()
        rollback(client, args.rollback)
        return

    if args.probe:
        base, token = resolve_credentials(args)
        client = Client(base, token, timeout=args.timeout,
                        verify=not args.insecure,
                        page_size=args.page_size)
        client.preflight()
        client.probe()
        return

    if args.purge or args.delete_tags or args.delete_groups:
        base, token = resolve_credentials(args)
        client = Client(base, token, timeout=args.timeout,
                        verify=not args.insecure, page_size=args.page_size)
        client.preflight()
        prefix = args.purge or args.delete_groups or args.delete_tags
        purge(client, prefix, args.apply, args.yes, args.purge_ignore_case,
              do_groups=bool(args.purge or args.delete_groups),
              do_tags=bool(args.purge or args.delete_tags))
        return

    if args.inspect_tags:
        base, token = resolve_credentials(args)
        client = Client(base, token, timeout=args.timeout,
                        verify=not args.insecure, page_size=args.page_size)
        client.preflight()
        inspect_tags(client)
        return

    if args.cleanup_scratch:
        base, token = resolve_credentials(args)
        client = Client(base, token, timeout=args.timeout,
                        verify=not args.insecure,
                        page_size=args.page_size)
        client.preflight()
        cleanup_scratch(client, args.probe_tag)
        return

    if args.discover_tag_api:
        base, token = resolve_credentials(args)
        client = Client(base, token, timeout=args.timeout,
                        verify=not args.insecure,
                        page_size=args.page_size)
        client.preflight()
        apis = client.fetch_inventory(max_pages=1)
        if not apis:
            sys.exit("Inventory returned no APIs to probe against.")
        discover_tag_api(client, apis, args.probe_tag, args.sample)
        return

    if not args.sheet:
        ap.error("--sheet is required (or use --rollback / --probe)")
    if not os.path.exists(args.sheet):
        sys.exit(f"Spreadsheet not found: {args.sheet}")

    if args.list_sheets:
        print(f"{args.sheet}\n")
        for name, headers, count in list_workbook(args.sheet):
            marker = "  " if count else "  (empty) "
            print(f"{marker}{name}   {count:,} data row(s)")
            print(f"      columns: {', '.join(map(str, headers)) or '(none)'}")
        print("\nPick one or more with --tab, then --inspect-sheet.")
        return

    sentinels = set(DEFAULT_SENTINELS) | {s.casefold() for s in args.sentinel}
    tabs_to_read = args.tab or [None]
    specs = parse_specs(args.tag_from)
    matched_specs = set()

    tabs = []
    dims_by_tab, host_by_tab, id_by_tab, path_by_tab = {}, {}, {}, {}

    for tab in tabs_to_read:
        headers, rows = load_tab(args.sheet, tab)
        label = tab or "(first sheet)"
        tabs.append((label, rows))

        host_by_tab[label] = find_column(headers, HOST_ALIASES, args.host_column)
        id_by_tab[label] = find_column(headers, ID_ALIASES, args.id_column)
        path_by_tab[label] = find_column(headers, PATH_ALIASES, args.path_column)

        print(f"\nTab {label!r}: {len(rows):,} data row(s)")
        print(f"  columns    : {', '.join(map(str, headers))}")
        print(f"  hostname   -> {host_by_tab[label] or '(not found)'}")
        print(f"  api id     -> {id_by_tab[label] or '(not found)'}")
        print(f"  path       -> {path_by_tab[label] or '(not found)'}")

        if not (host_by_tab[label] or id_by_tab[label]):
            sys.exit(f"\nTab {label!r} has no hostname or API-id column. "
                     "Point at it with --host-column.")

        if args.tag_from:
            dims, matched = resolve_dimensions(specs, headers)
            dims_by_tab[label] = dims
            matched_specs |= matched
            print(f"  tag columns-> "
                  f"{', '.join(d.column for d in dims) or '(none on this tab)'}")
        else:
            dims_by_tab[label] = []

        if args.inspect_sheet or not args.tag_from:
            print("\n  Candidate tag dimensions:")
            print(f"    {'COLUMN':<26} {'USABLE':>8} {'DISTINCT':>9} "
                  f"{'BLANK':>7}  NOTE")
            for header, usable, distinct, blanks, note in suggest_dimensions(
                    headers, rows, sentinels):
                print(f"    {str(header)[:26]:<26} {usable:>8,} {distinct:>9,} "
                      f"{blanks:>7,}  {note}")

    unmatched = [c for c, _ in specs if c not in matched_specs]
    if unmatched:
        sys.exit(f"\n--tag-from column(s) not found on any selected tab: "
                 f"{', '.join(unmatched)}")

    if not args.tag_from:
        print("\nNo --tag-from given, so there is nothing to tag with.")
        print("Pick one or more columns from the table above, e.g.:")
        print("  --tag-from 'Department:dept' --tag-from 'Market:market'")
        return

    rules, skipped, conflicts = build_rules(
        tabs, dims_by_tab, host_by_tab, id_by_tab, path_by_tab,
        sentinels, not args.no_prefix)

    print(f"\nParsed {len(rules):,} rule(s); {len(skipped):,} row(s) skipped; "
          f"{len(conflicts):,} conflict(s).")

    if args.inspect_sheet:
        print("\nFirst 15 rules:")
        for rule in rules[:15]:
            print(f"  [{rule.tab}] row {rule.row_no:<5} tag='{rule.tag}'  "
                  f"{rule.criteria()}")
        if conflicts:
            print(f"\nFirst 10 of {len(conflicts)} conflict(s):")
            for c in conflicts[:10]:
                print(f'  {c["host"]}  {c["dimension"]}: '
                      f'"{c["first"]["value"]}" ({c["first"]["tab"]}) vs '
                      f'"{c["second"]["value"]}" ({c["second"]["tab"]})')
        print("\n(--inspect-sheet made no network calls.)")
        return

    if not rules:
        sys.exit("No usable rules. Check --tag-from against --inspect-sheet.")

    base, token = resolve_credentials(args)
    client = Client(base, token, timeout=args.timeout,
                        verify=not args.insecure,
                        page_size=args.page_size)
    client.preflight()

    fields = ([f.strip() for f in args.return_fields.split(",") if f.strip()]
              if args.return_fields else None)
    cache = None if args.refresh_inventory else args.inventory_cache
    apis = load_inventory(client, cache, args.cache_max_age, fields)
    if args.inventory_cache and args.refresh_inventory and apis:
        try:
            with open(args.inventory_cache, "w", encoding="utf-8") as fh:
                json.dump({"fetched": datetime.now(timezone.utc).isoformat(),
                           "base": client.base, "apis": apis}, fh)
        except OSError:
            pass
    (plan, already, unresolved, resolved,
     resolved_hosts) = build_plan(rules, apis)
    by_tag = invert_plan(plan)

    # Every tag that landed on a real API -- newly applied OR already present.
    # Group creation keys off this, not off the write plan, so a re-run after
    # tagging still builds the groups instead of finding nothing to do.
    group_tags = {tag: sorted(ids) for tag, ids in resolved.items()}

    existing_names = set()
    if args.create_groups or args.groups_from_hosts or args.rule_tags:
        existing_names = client.fetch_groups(dump_raw=args.dump_raw)
        print(f"  {len(existing_names):,} existing group(s) in tenant")

    if args.plan_out:
        write_plan_file(args.plan_out, rules, plan, by_tag, skipped,
                        conflicts, unresolved)

    GROUP_NAMING[0] = args.group_name_from
    GROUP_NAMING[1] = tuple(p for _, p in specs if p)

    if not args.apply:
        if args.rule_tags:
            LOG.out("\n  --rule-tags is not supported: the spec defines a tag "
                    "as { name } only.")
        if args.groups_from_hosts:
            create_groups_from_hosts(
                client, resolved_hosts, existing_names, args.group_type,
                args.group_prefix, args.group_parent, args.group_description,
                {"groups": []}, args.split_groups_at, dry_run=True)
            print("\n  Re-run with --apply to create these.")
            return
        print_dry_run(rules, plan, already, by_tag, skipped, conflicts,
                      unresolved, apis, args.create_groups, existing_names,
                      args.group_prefix, group_tags)
        return

    if args.groups_from_hosts and not resolved_hosts:
        sys.exit("\nNo hostnames resolved against the inventory, so there is "
                 "nothing to build groups from.")

    if args.create_groups and not group_tags:
        sys.exit("\nNo tags resolved onto any API, so there is nothing to "
                 "build groups from.")

    if args.replace and not args.groups_only:
        print("\n*** --replace uses PUT: existing tags on matched APIs will be "
              "OVERWRITTEN, not merged. ***")
        if input("    Type REPLACE to continue: ").strip() != "REPLACE":
            sys.exit("Aborted.")

    GROUP_NAMING[0] = args.group_name_from
    GROUP_NAMING[1] = tuple(p for _, p in specs if p)

    PROBE_BUDGET[0] = args.max_probes
    PROBE_BUDGET[1] = args.probe_timeout

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    ledger = {"started": stamp, "base": base, "tags": [], "groups": [],
              "failed": []}
    ledger_path = f"tag-run-{stamp}.ledger.json"

    try:
        if args.rule_tags:
            LOG.out("\n  --rule-tags is not supported: the spec defines a tag "
                    "as { name } only,\n  so tags are assigned to APIs rather "
                    "than matching them by rule. Tagging normally.")
            args.rule_tags = False
        if False:
            created, _ = create_rule_tags(client, resolved_hosts,
                                          args.group_type, ledger,
                                          split_at=args.split_groups_at)
            if created:
                create_groups(client, {t: [] for t in created}, existing_names,
                              args.group_type, args.group_prefix,
                              args.group_parent, args.group_description,
                              ledger)
        elif args.groups_from_hosts:
            create_groups_from_hosts(
                client, resolved_hosts, existing_names, args.group_type,
                args.group_prefix, args.group_parent, args.group_description,
                ledger, args.split_groups_at)
        elif args.groups_only:
            print("\n--groups-only: skipping tag writes.")
        elif by_tag:
            apply_tags(client, by_tag, resolved_hosts, args.replace,
                       args.limit, ledger, args.allow_single_fallback,
                       args.max_failures)
        else:
            print("\nNo tag writes needed -- every matched API already "
                  "carries its tag.")
        if args.create_groups and not args.groups_from_hosts:
            create_groups(client, group_tags, existing_names, args.group_type,
                          args.group_prefix, args.group_parent,
                          args.group_description, ledger,
                          args.group_name_from, GROUP_NAMING[1])
    finally:
        with open(ledger_path, "w", encoding="utf-8") as fh:
            json.dump(ledger, fh, indent=2)
        print(f"\nLedger: {ledger_path}")
        if args.create_groups:
            print(f"Undo groups with: --rollback {ledger_path}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        LOG.summary()
        LOG.close()
        sys.exit("\nInterrupted.")
    except SystemExit:
        LOG.summary()
        LOG.close()
        raise
    except Exception:
        import traceback
        LOG.out("\nUnhandled error:")
        LOG.out(traceback.format_exc())
        LOG.summary()
        LOG.close()
        sys.exit(1)
    else:
        LOG.summary()
        LOG.close()
