# scope/ — how long each curriculum actually is

One file per curriculum, listing its lessons in order. This is what makes
"are we on track?" answerable: current position + total lessons + weeks
remaining = a real verdict.

Name files after the curriculum, not the child — two children on the same
program share one scope file.

```
dimensions-math-4a.md
all-about-reading-4.md
story-of-the-world-2.md
```

## Format

Whatever's easiest to type from the table of contents. A flat numbered list
is plenty:

```
# Dimensions Math 4A — scope and sequence

## Chapter 1: Numbers to 100,000
1. Numbers to 100,000
2. Place Value
3. Comparing and Ordering Numbers
4. Rounding 5-Digit Numbers
...
```

You don't have to type these yourself. Photograph the table of contents and
say:

> Turn this table of contents into a scope file at
> curricula/scope/dimensions-math-4a.md

## A rule worth keeping

Have the assistant **write a small script for the pacing arithmetic and run
it**, rather than doing the math in its head. Same numbers every time, and
you can check the script. Ask for it once:

> Write scripts/pacing.py that reads family.yaml and the curricula files and
> computes remaining-lessons-per-week for each child. Then run it.
