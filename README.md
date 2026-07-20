# Board Game

Google Apps Script that adds a custom menu to a Google Spreadsheet for fetching BoardGameGeek game data and updating related sheets (Rankings, Titles, Ratings).

Developed with TypeScript and [clasp](https://github.com/google/clasp), running inside Docker via [mise](https://mise.jdx.dev/) tasks.

## Architecture

The Apps Script project keeps its global entry points deliberately small and
groups the implementation by responsibility:

- `src/entrypoints/` — Apps Script functions discovered by name (`onOpen`, `update`)
- `src/services/` — update orchestration and sheet-specific business logic
- `src/infrastructure/` — Apps Script API adapters for HTTP, properties, triggers, and sheets
- `src/config/` — sheet layouts, endpoints, batch limits, and source-specific title rules
- `src/shared/` — parsing and error-handling utilities shared by services

The unified `update` command refreshes rankings and ratings immediately, then
uses one time-driven trigger to finish game and title batches. This prevents
independent legacy triggers from updating the same spreadsheet concurrently.

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
