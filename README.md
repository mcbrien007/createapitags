# create_api_tags.py

Tag APIs from a hostname register, then build groups from those tags — against the
Akamai API Security (Noname) Management API.

Reads a spreadsheet mapping hostnames to an owning department, applies a tag to every
matching API in the inventory, then creates one group per department filtered on those
tags.

An API record is a **host + path + method** combination, so a single hostname usually
covers hundreds of APIs. The script resolves each row against the live inventory and
writes tags in batches of API IDs.

The result is two levels:

- Tag `dept:CPP` on every API whose hostname the register assigns to CPP
- Group `CPP`, filtered on `tags in ["dept:CPP"]`, nested under a parent container

> **Everything is dry-run by default.** No write happens without `--apply`.

---

## Contents

- [Install](#install)
- [Credentials](#credentials)
- [Preparing the spreadsheet](#preparing-the-spreadsheet)
- [Quick start](#quick-start)
- [Command reference](#command-reference)
- [Reading the dry run](#reading-the-dry-run)
- [How groups are built](#how-groups-are-built)
- [Cleanup and rollback](#cleanup-and-rollback)
- [Troubleshooting](#troubleshooting)
- [API contract reference](#api-contract-reference)

---

## Install

Python 3.8+ and two packages:

```bash
pip install requests openpyxl
```

`requests` handles the API calls. `openpyxl` is only needed for `.xlsx` registers — a
`.csv` or `.tsv` register needs nothing beyond the standard library.

Single file, no config file, no state beyond the ledger and cache files it writes into
the working directory.

---

## Credentials

```bash
export NONAME_API_BASE=https://<tenant>
export NONAME_API_TOKEN='<token>'
```

If either is unset the script prompts for it, reading the token through `getpass` so it
is never echoed. Tokens are never written to disk and are redacted from console output
and log files.

Management API tokens are short-lived — roughly two days. Generate one at
`POST /api/v3/users/generate-api-token`, or from **Settings → User Management** in the
portal.

| Preflight result | Meaning |
| --- | --- |
| `401` | Token expired or invalid |
| `403` | Token's role lacks inventory read |
| `ReadTimeout` | Tenant slow or unreachable — not a bad request |

---

## Preparing the spreadsheet

One row per hostname, plus at least one column describing that host. Column names are
detected automatically and can be overridden.

### Inspect before you run

```bash
python3 create_api_tags.py --sheet hosts.xlsx --list-sheets
python3 create_api_tags.py --sheet hosts.xlsx --tab Main --inspect-sheet
```

Neither call touches the network. `--inspect-sheet` profiles every column and scores it
as a tag dimension:

| Column | Usable | Distinct | Blank |
| --- | ---: | ---: | ---: |
| Department | 732 | 5 | 443 |
| Market | 646 | 61 | 529 |
| Application Name | 276 | 63 | 899 |
| Serviced Market | 0 | 0 | 1,175 |

A column where every value is unique, or where the count runs into the hundreds, is
flagged as a poor grouping dimension.

### Choose tag dimensions

`--tag-from "Column:prefix"` promotes a column to a tag namespace. Repeatable:

```bash
--tag-from "Department:dept"        # dept:CPP, dept:EDAA, ...
--tag-from "Market:market"          # market:Canada, ...
--tag-from "Application Name:app"
```

The prefix keeps dimensions separable so a department tag never collides with a market
tag. `--no-prefix` emits bare values instead.

> Start with one dimension. Three dimensions on a typical register produce 100+ groups.

### Multiple tabs and conflicts

Repeat `--tab` to merge worksheets. **Order sets priority** — the first `--tab` wins any
disagreement:

```bash
--tab Main --tab "IRR-Akamai" --dept-column "Department" --dept-column "CORP Area"
```

When two rows give the same host different values for one dimension, the later one is
dropped and recorded as a conflict. Conflicts are counted in the dry run and listed in
full in the `--plan-out` file.

### Values treated as blank

Filtered automatically: `Not Found`, `N/A`, `TBD`, `None`, `Unknown`, `Not Detected`,
`Not Enforced` and similar. Add your own with `--sentinel`.

Values are NFKC-normalized first, which catches non-breaking spaces that would otherwise
split one value into two tags.

---

## Quick start

Four steps, in order. Each is safe to repeat.

### 1. Dry run

No writes. Pages the inventory once and caches it so later steps are fast.

```bash
python3 create_api_tags.py \
  --sheet hosts.xlsx --tab Main \
  --tag-from "Department:dept" \
  --create-groups --group-parent "Departments" \
  --inventory-cache inv.json --plan-out plan.json \
  --log-file tag-run.log -v
```

### 2. Canary

Writes the first 500 tag assignments only. Check a few APIs in the portal before going
further.

```bash
python3 create_api_tags.py \
  --sheet hosts.xlsx --tab Main \
  --tag-from "Department:dept" \
  --apply --limit 500 \
  --inventory-cache inv.json --log-file tag-run.log -v
```

### 3. Full run

Tags everything, then builds the groups.

```bash
python3 create_api_tags.py \
  --sheet hosts.xlsx --tab Main \
  --tag-from "Department:dept" \
  --apply --create-groups --group-parent "Departments" \
  --refresh-inventory --inventory-cache inv.json \
  --log-file tag-run.log -v
```

> `--refresh-inventory` matters here. Without it the run reuses the cache from step 1,
> which predates the canary's tags, and those APIs get written again.

### 4. Verify

Open a group in the portal and confirm it resolves to the expected API count. An empty
group means the tag is not on the APIs yet — not a bad filter.

Re-running is safe at any point: APIs that already carry a tag are skipped, and groups
that already exist are left alone.

---

## Command reference

### Input

| Option | Meaning |
| --- | --- |
| `--sheet PATH` | Register file: `.xlsx`, `.csv` or `.tsv` |
| `--tab NAME` | Worksheet to read. Repeatable; first wins conflicts |
| `--tag-from COLUMN:PREFIX` | Promote a column to a tag dimension. Repeatable |
| `--no-prefix` | Emit bare values rather than `prefix:value` |
| `--host-column NAME` | Override hostname column detection |
| `--id-column NAME` | Override API-id column detection |
| `--sentinel VALUE` | Extra placeholder to treat as blank. Repeatable |

### Connection

| Option | Meaning |
| --- | --- |
| `--base URL` | Tenant URL, else `NONAME_API_BASE` |
| `--timeout N` | Per-request timeout, seconds. Default `60` |
| `--page-size N` | Inventory page size. Default `2000` |
| `--inventory-cache PATH` | Cache the inventory and reuse it next run |
| `--refresh-inventory` | Ignore the cache and re-fetch |
| `--cache-max-age H` | Hours before a cache is stale. Default `24` |
| `--return-fields LIST` | Override inventory `returnFields` |
| `--insecure` | Skip TLS verification. Lab tenants only |

### Writing

| Option | Meaning |
| --- | --- |
| `--apply` | Actually write. Without it the run is a dry run |
| `--limit N` | Cap tag assignments written. Use for canaries |
| `--replace` | `PUT` instead of `PATCH` — **overwrites** existing tags |
| `--max-failures N` | Abort after N consecutive failed batches. Default `5` |
| `--probe-timeout N` | Per-probe timeout during negotiation. Default `20` |
| `--max-probes N` | Cap on negotiation probes. Default `40` |

### Groups

| Option | Meaning |
| --- | --- |
| `--create-groups` | Create one group per distinct tag |
| `--groups-only` | Build groups without writing any tags |
| `--group-parent NAME` | Parent container group to nest under |
| `--group-name-from value\|tag` | Name from the department value (default) or the full tag |
| `--group-prefix TEXT` | Prefix for generated group names |
| `--group-type APPLICATION\|OTHER` | Group type. Default `APPLICATION` |
| `--group-description TEXT` | Description applied to every group |

### Diagnostics and cleanup

| Option | Meaning |
| --- | --- |
| `-v` / `-vv` / `--debug` | Verbose planning detail / full HTTP tracing |
| `--log-file PATH` | Debug transcript, redacted, safe to attach to a ticket |
| `--plan-out PATH` | Write the resolved plan as JSON |
| `--probe` | Read-only endpoint check |
| `--inspect-tags` | Dump the tag catalogue schema |
| `--discover-tag-api --sample N` | Probe every route and body against N APIs |
| `--rollback LEDGER` | Delete the groups a prior run created |
| `--delete-tags PREFIX` | Delete catalogue tags by name prefix |
| `--delete-groups PREFIX` | Delete groups by name prefix, children first |
| `--purge PREFIX` | Delete both, groups first |
| `--yes` | Skip the delete confirmation prompt |

---

## Reading the dry run

```text
Inventory scanned      : 35,587 APIs
Rules from spreadsheet : 730
Rows skipped           : 445
Distinct tags to apply : 5
APIs receiving tags    : 25,313
Tag assignments to add : 25,313
Already tagged (no-op) : 0
```

Four things worth checking every time:

| Field | What a bad number means |
| --- | --- |
| **Rows skipped** | Usually a wrong column mapping, not bad data |
| **Hosts not in inventory** | Register describes decommissioned estate, or hosts not fronted by the connector |
| **Conflicts** | Same host, two values for one dimension. First `--tab` wins |
| **Per-tag breakdown** | On a re-run everything should show as already tagged |

The `--plan-out` JSON holds the complete picture: every tag with its resolved API IDs,
every conflict, every unresolved hostname, and every skipped row with its reason. It is
the artifact to hand an application owner when asking them to fill gaps.

---

## How groups are built

`--create-groups` makes one group per distinct tag. Each filters on the tag; the name
comes from the department value:

```text
Departments   tags in ["dept:CPP", "dept:EDAA", "dept:EPP", "dept:GCS", "dept:GTIO"]
  CPP         tags in ["dept:CPP"]
  EDAA        tags in ["dept:EDAA"]
  GCS         tags in ["dept:GCS"]
```

The `dept:` namespace stays on the tag so it never collides with market or application
tags, but is stripped from the group name. `--group-name-from tag` keeps the namespaced
form.

Details that matter in practice:

- Group names cannot contain `/` — it is replaced with `-` automatically
- An existing group of the same name is **skipped, not updated** — delete it first if its
  filter is wrong
- `--group-parent` creates a container group and nests the rest beneath it
- Groups only cover tags the run saw land on APIs, so a stale inventory cache can
  silently miss a department

If the tags are already applied and you only want the groups:

```bash
python3 create_api_tags.py ... --apply --groups-only --create-groups
```

---

## Cleanup and rollback

Every `--apply` run writes `tag-run-<timestamp>.ledger.json` recording each tag written,
each group created with its ID and filters, and each failed batch with its response body.

### Undo a run's groups

```bash
python3 create_api_tags.py --rollback tag-run-20260929-140302.ledger.json
```

Deletes only the groups that run created, children before parents. Tag assignments are
not reverted; the ledger lists every tag and API pair so they can be removed
deliberately.

### Delete by prefix

```bash
python3 create_api_tags.py --delete-groups "dept:"            # preview
python3 create_api_tags.py --delete-groups "dept:" --apply
python3 create_api_tags.py --delete-tags "zz-shape-probe" --apply
python3 create_api_tags.py --purge "dept:" --apply            # both, groups first
```

Groups go first deliberately. A group filtering on a tag that no longer exists silently
matches nothing, whereas an orphaned tag is only untidy.

Guards built in:

- Prefixes shorter than two characters are refused
- Reserved group names (`root`, `default`, `unassigned`, `all apis`) are never matched
- `--apply` asks you to type the match count back, unless `--yes`
- Deleting tags warns which groups would be left filtering on nothing
- A `purge-<timestamp>.ledger.json` records each deleted object's name, ID and filters

> There is no undelete on either endpoint. The purge ledger is the only way to
> reconstruct a group you removed by mistake.

---

## Troubleshooting

| Symptom | Cause | Fix |
| --- | --- | --- |
| `401` on preflight | Token expired — they last ~2 days | Generate a new one |
| `403` on preflight | Role lacks inventory read | Check role assignment |
| `ReadTimeout` on preflight | Tenant slow or unreachable | Retry, raise `--timeout`, check VPN |
| `each value in returnFields must be one of…` | Invalid field name or wrong encoding | Handled automatically |
| `terms should not be empty` | Wrong request body shape | Handled automatically by negotiation |
| `Operator equals is not supported for field tags` | Server restricts operators per field | Handled automatically — uses `in` |
| `Retrieved 0 group(s)` | Unrecognised response wrapper | Re-run with `--dump-raw` |
| Group created but empty | Tag not on the APIs yet | Check the tag run succeeded |
| Groups skipped as duplicates | Same-named groups already exist | `--delete-groups` then re-run |

### When the tag write is rejected

Negotiation walks endpoint and method combinations, then body shapes, stopping at the
first one the tenant accepts **and** verifies for blast radius. It prints progress with
elapsed time per probe, so a slow tenant looks slow rather than hung.

If nothing is accepted, capture the full matrix:

```bash
python3 create_api_tags.py --discover-tag-api --sample 8 --log-file tag-endpoint.log
```

That probes every route and body against eight different APIs, mixing records that
already carry tags with ones that do not. All eight failing identically points at the
endpoint; mixed results point at specific records. Either way the transcript is the
evidence for a support case.

### Reading the logs

| Flag | Output |
| --- | --- |
| `-v` | Planning decisions, per-tag resolution, inventory paging |
| `-vv` / `--debug` | Full HTTP tracing with request and response bodies |
| `--log-file PATH` | Full debug transcript regardless of console level |

Bearer tokens and secret-like JSON keys are redacted from both console and file,
including the echoed command line. The transcript still contains hostnames and API IDs,
so treat it as customer data.

Every exit path — including <kbd>Ctrl</kbd>+<kbd>C</kbd> — prints a request summary with
status counts and per-endpoint latency.

---

## API contract reference

Confirmed against the API Management OpenAPI spec (v3.0) and a live tenant. Useful when
adapting the script or writing something else against the same API.

### Tag assignment

```http
PATCH /api/v4/apis/tags   # append tags to a set of APIs
PUT   /api/v4/apis/tags   # replace them
```

```json
{
  "terms": { "ids": ["<apiId>", "<apiId>"] },
  "tagIds": ["<tagId>"]
}
```

`terms` is an **object wrapping an ids array** — schema `ApiIds` — not an array of filter
terms. Sending a filter array passes the `terms should not be empty` check and then fails
inside the handler with a `500` about retrieving tags for the specified API ID. Both
fields are required.

The single-API form `PATCH|PUT /api/v4/apis/{id}/tags` takes `{"tagIds": [...]}`.

### Tag catalogue

Tags are **static objects referenced by ID**, never by name. `createClassificationTag`
accepts only `{"name": "..."}`, and a Tag carries just `id`, `name` and `legacyId` —
there is no rule-based tag that matches APIs on its own.

```http
GET  /api/v4/tags   # list
POST /api/v4/tags   # create, returns the id
```

### Groups

`CreateGroupInput` requires `name`, `group_type` (`APPLICATION` or `OTHER`) and
`filters`. A `GroupFilterInput` is `{field, operator, value}`:

- **field** — `api_owner`, `host`, `infrastructure_tags`, `method`, `path`, `resources`,
  `sources`, `tags`
- **operator** — `in`, `contains`, `notContains`, `equals`, `notEquals`, `startsWith`,
  `endsWith`, `blank`, `notBlank`

> The server restricts operators **per field**, more narrowly than this enum suggests.
> The `tags` field rejects `equals` and accepts only `contains`, `notContains` and `in`.

Use `in` with an array value — it is exact, whereas `contains` is a substring match that
would make `dept:CPP` also match `dept:CPPX`.

Each field carries at most one filter; different fields AND together. Group names cannot
contain `/`.

### Inventory

```http
GET /api/v3/apis?limit=2000&offset=0&returnFields=id&returnFields=host
```

`returnFields` is validated as an **array**, so it must be sent as repeated query
parameters. A comma-joined string is read as one invalid field name and rejected. The
field for API ownership is `apiOwner`, not `api_owner`.

Paginate on the `moreEntities` flag. API IDs are UUIDv7 with the dashes stripped — 32 hex
characters.
