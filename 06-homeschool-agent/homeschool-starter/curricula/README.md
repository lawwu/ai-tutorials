# curricula/ — what each child learns

One `<child>.yaml` per child, and that file is the **source of truth** for
where they currently are. When a lesson gets finished, the `position` field
is what changes.

## Files

- `<child>.yaml` — subjects, curricula, and current position
- `<child>-reading-list.md` — titles you're considering or have queued
- `<child>-reading-log.md` — books actually finished, with dates
- `scope/` — scope-and-sequence lists for each curriculum, used for pacing

## Why position matters

"Are we on track?" is answerable only if the assistant knows where each
child actually is. It reads the position, counts what's left in
`scope/`, divides by the instructional weeks remaining in `family.yaml`,
and gives you a real answer instead of a reassuring one.

Update positions as you go — or just dictate your day and let the assistant
update them for you. That's the intended workflow:

> log: finished 4A chapter 2 lesson 1 today, went well

## Content rules for reading lists

Write your own standards at the top of each reading list — reading level
band, content you won't allow, how much stretch vs. comfort reading. Those
rules bind every book the assistant ever suggests for that child, so they're
worth twenty minutes of thought.

Example of what belongs there:

```
Rules for this list:
- Band: ~650-850L, story-shaped
- No horror, no romance plots
- Roughly 80% comfortable reading to 20% stretch
- Preview anything dealing with death of a parent before assigning
```
