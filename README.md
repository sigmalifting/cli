# SigmaLifting CLI

Command-line tools for creating, validating, editing, storing, inspecting, and exporting SigmaLifting training data.

This CLI is intentionally isolated from the app runtime. It operates on JSON bundle contracts, not on Realm, Expo APIs, or app services.

## Agent Skill

Companion agent skill:

- [`sigmalifting-cli`](https://github.com/sigmalifting/skills/tree/main/sigmalifting-cli)

## Why This Exists

Plain text is fine for explaining a training plan. It starts to fail when the plan needs to be used. If you run a program for a few weeks, your actual weights, reps, RPEs, missed sets, notes, and adjustments are now mixed with the original prescription. Sharing that as text means somebody else has to manually clean out your history before they can put in their own numbers.

SigmaLifting keeps those two things separate: a `program` is the reusable template, and a `process` is one lifter's run of that program. The CLI exists so agents and tools can create, validate, edit, store, and export that structure without turning it back into a wall of prose. Use text to explain the program; use SigmaLifting JSON to run it, share it cleanly, and let each lifter fill in their own training data.

The schema also gives agents a better way to think about training structure. Instead of flattening everything into free-form text, the CLI forces concrete decisions: what is the reusable program, what is one lifter's process, where blocks start and end, which days exist, how exercises are grouped, and what values are prescribed versus recorded. Ambiguity becomes visible instead of being hidden inside prose.

## What It Does

- Program bundle creation, validation, normalization, and mutation.
- Process bundle creation from a program bundle, validation, normalization, and mutation.
- XLSX import into normalized `program-import` and `process-import` bundles.
- XLSX export from `program-import` and `process-import` bundles.
- A self-contained RPE calculator surface.
- Machine-readable JSON output for all commands, including `help` and `doctor`.
- File input and output from local paths, `file://` URLs, and stdin where applicable.
- Local validation and business-logic checks included under `lib/validation`.

## Design Rules

The CLI follows these rules:

- Everything runtime-critical lives inside this package export.
- The CLI only knows about SigmaLifting JSON contracts.
- The CLI must be runnable in plain Node.
- The CLI must not depend on Realm or app storage choices.
- When developed inside an app repository, app builds should ignore this tool.

This means the CLI is best thought of as a bundle tool, not an app admin tool.

## Entry Points

- Main command surface: [`index.js`](index.js)
- Executable wrapper: [`bin/sigmalifting-cli.js`](bin/sigmalifting-cli.js)
- Bundle operations: [`lib/bundleOps.js`](lib/bundleOps.js)
- Validation layer: [`lib/validation/shared.js`](lib/validation/shared.js)
- RPE layer: [`lib/rpe.js`](lib/rpe.js)

## Command Surface

Top-level commands:

- `help`
- `doctor`
- `template`
- `validate`
- `normalize`
- `program`
- `process`
- `rpe`
- `schema`
- `inspect`
- `render`
- `xlsx`
- `store`

Program commands:

- `list`
- `show`
- `create`
- `copy`
- `update`
- `add-block`
- `update-block`
- `delete-block`
- `add-day`
- `rename-day`
- `delete-day`
- `update-schedule`
- `add-exercise`
- `update-exercise`
- `delete-exercise`
- `move-exercise`
- `copy-exercise`
- `copy-block`
- `copy-day`
- `add-custom-lift`
- `rename-custom-lift`
- `delete-custom-lift`

Process commands:

- `list`
- `show`
- `create-from-program`
- `update`
- `update-config`
- `update-one-rm`
- `update-config-and-one-rm`
- `update-set`
- `update-note`
- `get-recording`

RPE commands:

- `chart`
- `e1rm`
- `training-weight`
- `range`

Schema commands:

- `list`
- `show`

Inspect commands:

- `program-overview`
- `process-overview`
- `exercise-detail`
- `recording-detail`
- `one-rm-profile`

Render commands:

- `program-overview`
- `process-overview`
- `exercise-detail`
- `recording-detail`
- `one-rm-profile`

XLSX commands:

- `import`
- `export`

Store commands:

- `path`
- `init`

## Usage

Install from GitHub:

```bash
npm install -g https://github.com/sigmalifting/cli.git
sigmalifting-cli doctor --compact
sigmalifting-cli help --compact
```

Re-run the same install command to update. For a pinned install, append a commit SHA:

```bash
npm install -g https://github.com/sigmalifting/cli.git#<commit-sha>
```

Run from a package checkout:

```bash
npm run smoke
node ./bin/sigmalifting-cli.js help
node ./bin/sigmalifting-cli.js store path
node ./bin/sigmalifting-cli.js template program-import --out ./program.json
node ./bin/sigmalifting-cli.js validate program-import --file ./program.json
node ./bin/sigmalifting-cli.js xlsx import --file ./program.xlsx --out ./program-import.json
node ./bin/sigmalifting-cli.js xlsx export --file ./process-import.json --out ./process.xlsx
```

Run directly:

```bash
node ./bin/sigmalifting-cli.js help --compact
node ./bin/sigmalifting-cli.js template process-import --compact
```

Install from a package checkout:

```bash
npm install -g .
sigmalifting-cli doctor --compact
```

Build and install a local tarball:

```bash
mkdir -p ../sandbox/cli-pack
npm pack --pack-destination ../sandbox/cli-pack
npm install -g ../sandbox/cli-pack/sigmalifting-agent-cli-0.0.0-local.0.tgz
sigmalifting-cli help --compact
```

The CLI package is Apache-2.0 licensed and distributed from GitHub.

For coding agents, point the agent at the installed executable or call it by absolute path:

```bash
which sigmalifting-cli
sigmalifting-cli validate program-import --file ./program.json --compact
/absolute/path/to/sigmalifting-cli render program-overview --file ./program.json --raw
```

Input and output rules:

- `--file` accepts a local path, `file://` URL, or `-` for stdin.
- `--data` accepts inline JSON.
- `--data-file` accepts a local path or `file://` URL.
- `--out` accepts a local path or `file://` URL.
- `--exercise-id` selects an exercise for detail renders and process recording commands.
- `--compact` switches stdout to single-line JSON.
- `--raw` prints direct text for `render`.
- `--store-root` overrides the local JSON store root for one invocation.
- `xlsx import` requires a local workbook path or `file://` URL.
- `xlsx export` writes a workbook to `--out`.

All command responses are JSON objects with `ok: true` or `ok: false`.

## Local JSON Store

Program and process bundles are stored as local JSON by default whenever a command returns a `program-import` or `process-import` bundle. JSON is the source of truth for the CLI. XLSX files are import/export artifacts, not core storage.

Default storage root:

- macOS: `~/Library/Application Support/SigmaLifting/cli`
- Linux: `${XDG_DATA_HOME:-~/.local/share}/sigmalifting/cli`
- Windows: `%APPDATA%\SigmaLifting\cli`

Set `SIGMALIFTING_HOME` to override the root persistently, or pass `--store-root <dir>` for a single invocation.

Storage layout:

```text
SigmaLifting/cli/
  programs/
    candito-6-week__prog_abc123.json
  processes/
    april-meet-prep__proc_xyz789.json
  exports/
  index.json
```

Programs and processes are separate because that matches the app model. A process references one program snapshot, while a program can exist with zero, one, or many processes. Filenames combine a human-readable slug with the stable entity id, so duplicate names remain safe.

Common stored workflows:

```bash
sigmalifting-cli store init
sigmalifting-cli program create --name "Candito 6 Week"
sigmalifting-cli program list
sigmalifting-cli program show --program-id prog_abc123

sigmalifting-cli process create-from-program \
  --program-id prog_abc123 \
  --name "April Meet Prep" \
  --start-date 2026-04-27

sigmalifting-cli process list
sigmalifting-cli process update-set \
  --process-id proc_xyz789 \
  --exercise-id ex_abc123 \
  --week-index 0 \
  --set-index 0 \
  --data '{"weight":140,"reps":5,"rpe":8,"completed":true}'
```

File-based workflows still work. If `--out` is supplied, the CLI writes that file in addition to saving the bundle in the local JSON store.

## Data Model

The CLI currently works with these schema kinds:

- `program-import`
- `process-import`
- `program`
- `process`
- `exercise`

The important distinction is:

- `program` and `process` are bare model payloads.
- `program-import` and `process-import` are bundle payloads that include metadata and related arrays.

The CLI primarily operates on the import bundle forms because that is the cleanest database-agnostic interchange format.

## Architecture

The runtime is split into three layers.

1. CLI parsing and I/O in [`index.js`](index.js)
2. Pure bundle manipulation in [`lib/bundleOps.js`](lib/bundleOps.js)
3. Pure validation and math helpers in [`lib/validation/shared.js`](lib/validation/shared.js) and [`lib/rpe.js`](lib/rpe.js)

That separation is intentional:

- `index.js` should stay thin.
- `bundleOps.js` is the actual product logic.
- validation and RPE should remain independently testable.

## Self-Contained Runtime

The CLI includes the runtime pieces it needs:

- ID generation
- import size limits
- validation schemas
- business-logic validation
- RPE chart and calculations

That keeps the command-line surface portable and independent from mobile app storage.

## Verification

```bash
npm run smoke
npm pack --dry-run --json
sigmalifting-cli help --compact
sigmalifting-cli doctor --compact
```
