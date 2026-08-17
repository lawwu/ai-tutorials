# plans/ — the pages you actually print

Generated pages, one folder per child, plus a parents' brief for the week.

```
plans/
  week-brief-2026-08-17.md      the parents' one-pager
  child-one/2026-08-17.md       what this child does on this day
  child-two/2026-08-17.md
```

These are **outputs**. They get regenerated from `curricula/` and
`pacing/` whenever the week changes — so never record a decision only here,
because it will be overwritten. Decisions belong in `curricula/`,
`family.yaml`, or `TODO.md`.

## Decide the format once, then write it down

Different children need different pages, and the format is worth deciding
deliberately and then recording in `CLAUDE.md`. Two that work:

- **Older child, self-scheduling.** One page per school day, carrying that
  day's list, two or three guardrails, and a blank "My plan for today" table
  they fill in themselves. They own the order.
- **Younger child.** Same one page per day, but with a pre-filled "How the
  morning flows" table instead of a blank one, plus a short "For Mom" note.

Whatever you choose, put it in `CLAUDE.md` in a sentence like: *"Child One's
format, decided 2026-08-10: one page per school day at
`plans/child-one/YYYY-MM-DD.md`, with a blank planning table."* Then it
survives, and every future week matches.

## Ask for a week

```
plan this week
```

Once your curricula and weekly rhythm are filled in, that's the whole
prompt.
