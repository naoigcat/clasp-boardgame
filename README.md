# Board Game

Google Apps Script that adds a custom menu to a Google Spreadsheet for fetching BoardGameGeek game data and updating related sheets (Rankings, Titles, Ratings).

Developed with TypeScript and [clasp](https://github.com/google/clasp), running inside Docker via [mise](https://mise.jdx.dev/) tasks.

## Architecture

The Apps Script project keeps its global entry points deliberately small and
groups the implementation by responsibility:

- `src/entrypoints/` — Apps Script functions discovered by name (`onOpen`, `update`)
- `src/services/` — update orchestration and sheet-specific business logic
- `src/infrastructure/` — Apps Script API adapters for HTTP, properties, triggers, and sheets
- `src/config/` — sheet layouts, endpoints, runtime limits, and source-specific title rules
- `src/shared/` — parsing and error-handling utilities shared by services

The unified `update` command refreshes rankings and ratings immediately, then
uses one time-driven trigger to finish game and title batches. This prevents
independent legacy triggers from updating the same spreadsheet concurrently.

## Update behavior

Open the spreadsheet and select **Functions > Update** to start an update. The
menu action deliberately does not try to finish every external request in one
Apps Script execution:

1. Rankings are imported from Board Game Arena and ratings are imported from
   Bodoge in the initial execution.
2. The coordinator saves the Games phase and creates one five-minute,
   time-driven trigger.
3. Each trigger execution refreshes stale BoardGameGeek game rows until about
   180 seconds have elapsed. A game is stale when it has never been updated or
   its last update attempt is seven days old or older. Remaining stale rows
   resume on the next trigger.
4. Once Games is complete, each subsequent execution normalizes Board Game
   Arena titles under the same 180-second soft limit.
5. The trigger and transient queue state are removed after the final title
   batch.

The project uses an Apps Script lock, so an overlapping menu click or trigger
execution is skipped rather than allowing two executions to write the same
sheet. Starting the menu command again intentionally restarts the queue: it
removes the prior queue trigger and its saved phase before beginning a new
cycle. Do not create an additional trigger for the update handler manually.

## Spreadsheet contract

Keep the following sheet names and a header row in row 1. The updater treats
these tabs as an application data contract; renaming a tab or changing the
managed column order requires a corresponding change in the configuration.

| Sheet    | Required input                                           | How the updater manages it                                                                                                                                                                                             |
| -------- | -------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Games    | A BoardGameGeek rich-text link in column A for each game | Reads rows until the first blank link, then refreshes the managed values in columns B through AA. Formula-dependent cells are cleared before values are rewritten so sheet formulas can recalculate from current data. |
| Rankings | Header row only                                          | Replaces rows below the header with the current Board Game Arena catalog. An empty parsed catalog leaves the previous snapshot untouched.                                                                              |
| Titles   | Header row only; existing URL rows are preserved         | Adds Board Game Arena URLs found in Rankings, stores the source Japanese title, its canonical matching title, and the most recent error.                                                                               |
| Ratings  | Header row only                                          | Replaces rows below the header with the configured Bodoge user's ratings after its pages have been fetched without a request error.                                                                                    |

The exact column positions and batch limits are centralized in
src/config/SheetSchema.ts and src/config/AppConfig.ts. Use those files when a
spreadsheet layout or an external-source rule must change; do not duplicate
magic column numbers in a service.

## Prerequisites

- [Docker](https://docs.docker.com/get-docker/)
- [mise](https://mise.jdx.dev/getting-started.html)

## Setup

First-time setup from a fresh clone:

```sh
mise run build     # Build the Docker image
mise run login     # Authenticate clasp (see Login below)
```

If you are connecting to an existing Apps Script project, skip clone and run `mise run pull` instead. To download a new project by its script ID (from the GAS project URL), run:

```sh
mise run clone {SCRIPT_ID}
```

Deploy the script to Google Apps Script:

```sh
mise run push
```

### Script properties

Some features call the [BoardGameGeek XML API](https://boardgamegeek.com/using_the_xml_api) with a Bearer token. Set the `TOKEN` script property in the Apps Script project:

1. Open the project in the [Apps Script editor](https://script.google.com/).
2. Go to **Project Settings** (gear icon) → **Script properties**.
3. Add a property named `TOKEN` with your BoardGameGeek application token as the value.

Without `TOKEN`, unauthenticated requests may be rate-limited or fail.

Set the optional `BODOGE_USER_ID` script property to import a Bodoge user's
played-game ratings. When it is absent or empty, the Ratings sheet is left
unchanged. The `UPDATE_STEP` property is maintained internally while a
multi-execution update is pending; do not create or edit it manually.

Store tokens and user IDs only in Apps Script script properties. They are
environment-specific configuration and must not be committed to the
repository.

### First update

After pushing the project and setting the required properties, reload the
target spreadsheet so Apps Script runs onOpen and shows the **Functions** menu.
Choose **Update** once, then allow the managed trigger to finish any remaining
Games and Titles batches. Execution logs in the Apps Script editor show skipped
overlaps and row-level external-service errors.

## Build

Rebuild Docker image.

```sh
mise run build
```

## Verify

Run the TypeScript check and unit tests before pushing a script change:

```sh
mise run typecheck
mise run test
```

Use the following routine for local changes:

1. Run `mise run format` after editing TypeScript or documentation.
2. Run `mise run typecheck` to validate the Apps Script types without
   generating JavaScript.
3. Run `mise run test` to execute the Node-based Apps Script harness.
4. Run `mise run push` only after the checks pass and you intend to deploy.

The test harness validates service behavior without contacting the external
sites. It is therefore suitable for checking parsing, queue transitions, and
sheet-write behavior locally, but it cannot detect a source site's HTML or API
format change.

## Failure handling

- A failed BoardGameGeek game lookup records its message and advances the row's
  attempt timestamp so later batches can refresh other stale games. The same
  refresh window later makes the failed row eligible again.
- A failed title lookup records its message and leaves the canonical title
  blank. Later batches prefer titles that have not yet failed so the queue can
  advance; a retry pass runs once only failed rows remain.
- Rankings are cleared only after a non-empty catalog has been fetched and
  parsed, preserving the last snapshot when the request or catalog parser
  fails.
- A request failure or unrecognized Bodoge page occurs before Ratings is
  cleared, preserving the last snapshot. Pagination ends only when Bodoge's
  empty played-games marker is present.
- For logged BoardGameGeek, title, or Board Game Arena catalog failures, inspect
  the Apps Script execution log and update the relevant parser or title rule
  before rerunning the update.

## Login

Login to clasp by accessing the URL output to the terminal, and pasting the authentication code to the terminal.

```sh
mise run login
```

## Clone

Download the script with `{SCRIPT_ID}` that can be obtained by the URL and save it.

```sh
mise run clone {SCRIPT_ID}
```

## Push

Compile TypeScript and upload the script.

```sh
mise run push
```

## Pull

Download the script.

```sh
mise run pull
```

## Remove

Stop containers and remove volumes (including clasp login credentials stored in the Docker volume).

```sh
mise run remove
```
