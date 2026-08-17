# Homeschool Starter

A starting skeleton for running your homeschool out of a folder that Claude
Code can read and write — plans, journals, curriculum positions, reading
logs, and to-dos, all as plain text you own.

Nothing here is code. Every file is either a Markdown note (`.md`) or a
simple list (`.yaml`). You can open any of them in any text editor.

**New to this?** Start with the setup guide, which walks through the
accounts, the apps, and the phone: <https://claude.ai/code/artifact/9173616a-7f0e-4a9b-8f28-34165da7e4e4>

## First 15 minutes

1. **Copy this folder** to your computer and rename it `homeschool`.
2. **Open it in the Claude app** — the Code tab, "Local", Select folder.
3. **Say this:**

   > Set this folder up as a git repository, create a new PRIVATE repository
   > on my GitHub account called "homeschool", and push it. I've never used
   > git or GitHub before — walk me through anything I need to click.

4. **Then fill in your family.** Say:

   > Read CLAUDE.md and family.yaml. Interview me one question at a time and
   > fill in every placeholder for our family. Rename the example child files
   > in curricula/ to my children's names as we go.

5. **Confirm your repo is private.** Open it on github.com and look for the
   word **Private** next to the name. This folder will hold your children's
   names, ages, and daily schedules.

That's setup. From here on you just talk to it.

## What goes where

| Folder | Holds |
| --- | --- |
| `CLAUDE.md` | Standing instructions, read at the start of every conversation |
| `AGENTS.md` | A shortcut pointing at the same file — see below |
| `family.yaml` | Children, school-year dates, weekly rhythm |
| `TODO.md` | Open items, deadlines, pending decisions |
| `curricula/` | What each child learns, and where they currently are |
| `journal/` | What actually happened, by date. Append-only |
| `plans/` | The pages you print each week |
| `pacing/` | Weekly templates and generated reports |
| `resources/` | Charter, vendor, and logistics knowledge base |
| `experiments/` | One file per thing you're trying, with dated observations |

### One instructions file, two names

`AGENTS.md` is a symlink to `CLAUDE.md` — the same file under two names, not
a copy. Claude Code looks for `CLAUDE.md`; most other AI coding assistants
look for `AGENTS.md`, which has become the cross-tool convention. Pointing
one at the other means you write your instructions once and any assistant
you try reads the same rules, including the privacy list.

Edit `CLAUDE.md`. Never edit `AGENTS.md` directly, and don't paste a second
copy of the instructions into it — two files that drift apart is exactly the
problem the symlink prevents.

*On Windows:* git may check the symlink out as a small text file containing
the word `CLAUDE.md` instead of a working shortcut. Harmless, but if you
want it working, run `git config --global core.symlinks true` before you
clone — or just tell Claude "AGENTS.md didn't come through as a symlink,
fix it."

### Folder conventions

Each folder has its own README explaining the convention. Read those before
you reorganize anything — the conventions are what make the assistant
useful, and most of them exist because something went wrong once.

## Delete what you don't need

This skeleton is deliberately a little larger than most families need on day
one. `experiments/`, `pacing/`, and `resources/` earn their place over a few
months. Delete any folder you're not using — an empty folder full of
placeholder files is worse than no folder, because it invites the assistant
to invent content for it.

## Things to try in week one

```
log: [talk into your phone about how today went]

plan this week

pacing report — are we on track to finish math by June?

Review this whole folder and tell me if there are any contradictions.
```

## The one rule that matters most

When the assistant gets something wrong, don't just correct it in the
conversation — **add a line to `CLAUDE.md` so it never gets that wrong
again.** That single habit is what turns this from a filing system into
something that compounds.
