# Jujutsu Workflow

This repository is configured for a colocated Git + Jujutsu workflow.

## Current setup

- Remote name: `origin`
- Trunk bookmark: `main`
- Default `jj` command: `status`
- Default fetch remote: `origin`
- Default push remote: `origin`
- New and fetched bookmarks from `origin` are auto-tracked

Repo-local config lives in `.jj/repo/config.toml`.

## Recommended commands

Fetch from GitHub:

```bash
jj sync
```

Show status:

```bash
jj
```

Create a new change on top of current work:

```bash
jj new
```

Name the current change:

```bash
jj describe -m "Your change summary"
```

See the local graph around trunk:

```bash
jj log -r "trunk():: | @ | @-"
```

Move `main` to the current change:

```bash
jj bookmark move main --to @
```

Push tracked bookmarks:

```bash
jj publish
```

## Typical update flow

1. `jj sync`
2. `jj new main`
3. edit files
4. `jj describe -m "..."`
5. `jj bookmark move main --to @`
6. `jj publish`
