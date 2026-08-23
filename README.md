# Apply Unique Colours

A SOLIDWORKS macro that makes a multi-body weldment readable at a glance by giving
every geometrically distinct body its own colour.

Works with SOLIDWORKS 2022, 2024 and 2025.

## What makes it useful

Bodies that are genuinely identical share a colour, so anything that *differs*
stands out immediately. On a weldment full of near-identical tubes, the one with a
hole in a different place is obvious at a glance instead of needing to be hunted
for.

Two bodies count as the same if the SOLIDWORKS geometry kernel can move one onto
the other — so position and orientation are irrelevant, while a hole moved a
fraction of a millimetre separates them. **A mirror image counts as a different
part**, because an opposite hand isn't interchangeable with its original on the
shop floor.

Colours are spread evenly around the hue wheel and corrected so every colour
carries the same perceived brightness. Without that correction blues and violets
read as near-black next to yellows — pure blue measures 0.07 relative luminance
against pure yellow's 0.93.

Everything is written to **the active display state only**, so other display states
and configurations are left untouched. In an assembly, nothing modifies a
referenced part file.

## Install

**The macro on its own:** download `Apply-Unique-Colours.swp` from the
[latest release](../../releases/latest), then run it with **Tools > Macro > Run**,
or add it to a toolbar with **Tools > Customize > Commands > Macro**.

**With [MacroShelf](https://github.com/james-debono/macroshelf-sw-addin):** get the
[MacroShelf Collection](https://github.com/james-debono/macroshelf-collection-sw-macro-library/releases/latest),
which packages this macro with its icon and hover text alongside every other macro
in the set. Point MacroShelf at the unzipped folder and it appears as a button.

## Using it

Open a part or assembly and run the macro. On a 56-body weldment it takes about 13
seconds — most of that is the display-state writes, which are the price of colours
that stay scoped to one display state.

To undo it, use
[Remove Body and Component Appearances](https://github.com/james-debono/remove-body-and-component-appearances-sw-macro),
which clears what this macro writes and nothing else.

## Worth knowing

**Pattern and mirror features have a "Propagate visual properties" option, on by
default.** When it's on, regenerating the feature stamps the seed body's colour
onto the derived bodies' *faces* — and a face appearance beats a body appearance,
so it overrides what this macro writes. The macro counts those features and tells
you when it finishes.

Switching that option off costs nothing here: identical bodies are grouped and
coloured alike regardless of how they were created.

If appearance behaviour ever looks wrong, check that checkbox before suspecting the
macro. It produced five separate-looking symptoms during development that all
turned out to be the same thing.

## Limitations

- **Assemblies group by file path and configuration, not geometry.** Two identical
  parts saved under different filenames get different colours. The geometry
  matching runs on part documents only.
- A multi-body part inside an assembly gets one colour for the whole component.
- Surface and sheet bodies are skipped.
- Suppressed and lightweight components aren't specially handled.
- **No undo.** Hundreds of appearance changes with no undo record.

## Related macros

- [Remove Body and Component Appearances](https://github.com/james-debono/remove-body-and-component-appearances-sw-macro)
  — the practical undo for this macro
- [Remove All Appearances](https://github.com/james-debono/remove-all-appearances-sw-macro)
  — clears every appearance in the display state

## Building from source

`src\Apply-Unique-Colours.vba` is the readable source. A `.swp` is a binary VBA
project, so it can only be produced from inside SOLIDWORKS — there is no build
step. Open the `.swp` via **Tools > Macro > Edit**, paste the source in, and save.

Technical detail, including why measurement-based matching cannot work here and
which approaches were tried and rejected, is in
[docs/DEVELOPMENT.md](docs/DEVELOPMENT.md). Read the **Dead ends** section before
changing the matching logic — most obvious ideas have already been tried and
measured.

## Licence

MIT — see [LICENSE](LICENSE). Free to use, modify and share. The full licence text
is also carried inside the macro itself, so a `.swp` passed on by itself still
carries its licence.

Created by James Debono, with AI assistance. Everything here was tested by
hand in SOLIDWORKS — nothing that touches the API can be verified any other way.

## Trademarks

SOLIDWORKS is a registered trademark of Dassault Systèmes SolidWorks Corporation.
This project is independent: it is not affiliated with, endorsed by, or sponsored
by Dassault Systèmes, and uses only the published SOLIDWORKS API.
