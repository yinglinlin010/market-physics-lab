---
name: Market Physics Lab
description: A restrained Streamlit research workbench that makes time-series evidence and model comparisons easy to inspect.
colors:
  primary: "#137C78"
  ink: "#172A35"
  secondary-action: "#D16B3C"
  neutral-bg: "#F6F4EE"
  surface: "#FFFFFF"
  sidebar: "#EAF0ED"
  rule: "#DCE2DD"
  muted-text: "#5F727A"
  primary-soft: "#E9F3F1"
  formula-ink: "#134E4B"
  primary-hover: "#0D625F"
  baseline-series: "#9AA9A5"
typography:
  interface:
    fontFamily: "PingFang TC, Microsoft JhengHei, sans-serif"
    fontSize: "16px"
    fontWeight: 400
    lineHeight: 1.5
    letterSpacing: "normal"
  section-heading:
    fontFamily: "PingFang TC, Microsoft JhengHei, sans-serif"
    fontSize: "20px"
    fontWeight: 600
    lineHeight: 1.3
    letterSpacing: "-0.02em"
  title:
    fontFamily: "PingFang TC, Microsoft JhengHei, sans-serif"
    fontSize: "34px"
    fontWeight: 700
    lineHeight: 1.2
    letterSpacing: "-0.035em"
  equation:
    fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace"
    fontSize: "14px"
    fontWeight: 400
    lineHeight: 1.5
    letterSpacing: "normal"
  mobile-title:
    fontFamily: "PingFang TC, Microsoft JhengHei, sans-serif"
    fontSize: "26px"
    fontWeight: 700
    lineHeight: 1.2
    letterSpacing: "-0.035em"
  sidebar-heading:
    fontFamily: "PingFang TC, Microsoft JhengHei, sans-serif"
    fontSize: "18px"
    fontWeight: 600
    lineHeight: 1.3
    letterSpacing: "normal"
rounded:
  sm: "5px"
  md: "8px"
spacing:
  xs: "8px"
  sm: "14px"
  md: "18px"
  lg: "24px"
  xl: "32px"
components:
  button-primary:
    backgroundColor: "{colors.primary}"
    textColor: "{colors.surface}"
    typography: "{typography.interface}"
    rounded: "{rounded.md}"
    height: "45px"
  evidence-panel:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.ink}"
    rounded: "{rounded.md}"
    padding: "15px 18px"
  formula-panel:
    backgroundColor: "{colors.primary-soft}"
    textColor: "{colors.formula-ink}"
    typography: "{typography.equation}"
    rounded: "{rounded.sm}"
    padding: "8px 10px"
---

## Overview

**Creative North Star: The Research Workbench.** The screen makes one research task feel orderly: set the data and time boundary, run three comparable models, then inspect the evidence. The visual language stays quiet enough for formulas and plotted values to carry attention. Warm paper, evergreen actions, and a single rust comparison color distinguish the controls from the result series.

**Key Characteristics:**
- The product shell stays familiar: Streamlit controls in the left rail, evidence and results in the main work area.
- A shared grid, restrained borders, and tabular numbers keep dates and metrics scan-friendly.
- Teal identifies the main action and OU series; rust identifies PySR and state that deserves attention; gray-green marks the no-change baseline.
- Equations use a monospace face only where their characters benefit from alignment.

## Colors

The normative values are in the frontmatter. The warm neutral background separates the main work area from white evidence surfaces. The pale green sidebar gives configuration a stable home. Evergreen is reserved for the primary action and OU series, not decorative highlights. Rust marks the alternative formula and secondary emphasis. The gray-green line marks the baseline so it remains legible without competing with the observed series.

**The Evidence Color Rule.** A color names one stable role across the control, formula, and chart; never change a series color between views.

### Primary
- **Evergreen** marks the main action and OU series. Its darker hover shade is reserved for the primary button state.

### Secondary
- **Rust** marks the PySR series and secondary emphasis.

### Neutral
- **Deep Ink** carries headings and observed values; **Warm Paper** is the page ground; **Configuration Green** distinguishes the sidebar; **Evidence White** groups active experiment details; **Divider** marks section boundaries; **Secondary Text** lowers emphasis; **Formula Tint** and **Formula Ink** keep equations legible; **Baseline Gray Green** marks the no-change prediction.

## Typography

Use the platform's available Traditional Chinese sans-serif stack for the interface, headings, labels, and prose. Let font metrics fall back locally so the workbench remains usable offline. Reserve monospace for equations and compact values that benefit from fixed-width alignment. Keep plotted values and metrics tabular where the framework supports it.

**The Equation-Only Mono Rule.** Monospace is for formulas and machine-shaped values, not a general-purpose “technical” costume.

## Layout

The configuration rail is persistent on desktop and precedes the research content on narrow screens. The main column begins with the question and experiment state, then reports metrics before the aligned out-of-sample charts, equations, comparison, and conclusion. Detail that matters for auditability but not for the first reading stays in native expanders. On small screens, columns stack in reading order; labels and controls remain in Streamlit's standard responsive layout.

**The Shared-Test-Window Rule.** Put the three model series on the same chart axes and display their common test dates and target together.

## Elevation & Depth

The workbench is flat by default. Depth comes from a warm page ground, a distinct configuration rail, white evidence surfaces, and thin rules. Avoid decorative shadows; use borders and surface contrast to group data that must be checked.

## Shapes

Controls and evidence panels use modest, consistent corners. Buttons remain recognizably native and easy to locate. Formula blocks wrap long expressions rather than clipping them. Avoid pill-shaped containers for large content groups.

## Components

### Primary Action

The single filled action starts the comparison. It uses evergreen with high-contrast text, a compact radius, and a visible keyboard focus ring. Loading and failure feedback stays next to the action and names the recovery path.

### Evidence Panel

The experiment summary uses a white surface, thin neutral rule, and compact padding. It groups the selected pair, period, split, objective, and status without turning each value into a separate card.

### Formula Panel

The equation sits on a pale green surface in fixed-width type. Long formulas wrap; explanatory prose follows in ordinary interface type. Candidate origin and selection split stay visible alongside it.

### Metrics and Charts

Metrics use tabular numerals. The fixed mapping is observed value in deep ink, no-change in baseline gray-green, OU in evergreen, and PySR in rust. All three predictions share dates, target units, and axes.

## Do's and Don'ts

### Do:
- **Do** keep the data source, retrieval time, adjustment method, and date split available in the audit expander.
- **Do** preserve one series-color mapping in every plot and explanation.
- **Do** show the model's failed or unsupported state in plain language.
- **Do** use local/system fonts so an offline presentation does not fetch visual assets.

### Don't:
- **Don't** style the app as a laboratory instrument or trading terminal; the task needs a research workbench.
- **Don't** treat the synthetic dataset as stock evidence.
- **Don't** hide the baseline, validation selection rule, or non-winning result.
- **Don't** add shadows, gradients, or extra accent colors to data containers.
