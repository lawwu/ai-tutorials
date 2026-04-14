# Tutorial 3: AI Use Cases — Claude Cowork & Claude Code

**Video:** https://www.youtube.com/watch?v=zh5twqQLsB4  
**Format:** In-person, less technical audience  
**Date:** April 13, 2026

## Overview

Use-case focused tutorial covering Claude Cowork and Claude Code for a less technical audience. Demonstrates practical applications including data analysis, dashboard generation, and working with Claude's memory and skills systems.

**Topics:**
- Claude Cowork intro and use cases
- Claude Code 101 (whiteboard walkthrough)
- Demo: combining AQ + Impact campaign data into an Excel workbook and interactive HTML dashboard
- Demo: automated data quality tests with pytest
- Working with financial statements via skills
- Claude memory and CLAUDE.md

## Files

| File | Description |
|------|-------------|
| `notes-20260413.md` | Tutorial notes and agenda |
| `cc_101.excalidraw` | Claude Code 101 whiteboard diagram |
| `CLAUDE.md` | Example CLAUDE.md used in the demo |
| `combine_aq_impact.py` | Demo script — merges AQ + Impact Excel data into a formatted workbook and HTML dashboard |
| `test_data_quality.py` | Demo script — pytest-based data quality tests for the campaign data |
| `chrome_tabs.md` | Reference links open during the tutorial |

## Demo: Campaign Data Pipeline

The `combine_aq_impact.py` script was built live during the tutorial using Claude Code. It takes an Excel workbook with `AQ` and `Impact` sheets and produces:
1. `AQ_Impact_Combined.xlsx` — formatted Excel with color-coded sections
2. `AQ_Impact_Dashboard.html` — interactive filterable dashboard with charts

```bash
pip install pandas openpyxl
python combine_aq_impact.py "your_campaign_data.xlsx"
```

The `test_data_quality.py` script provides automated quality checks on the output:

```bash
pip install pytest pandas openpyxl
pytest test_data_quality.py -v
```

## Related Links

- Anthropic knowledge work finance skills: https://github.com/anthropics/knowledge-work-plugins/tree/main/finance/skills
- Claude Code docs: https://code.claude.com/docs/en/memory
- Claude models overview: https://platform.claude.com/docs/en/about-claude/models/overview
