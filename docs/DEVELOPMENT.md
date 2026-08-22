# Apply Unique Colours — development notes

The API findings behind this macro. `README.md` covers what it does and how to
use it.

## How bodies are matched

`IBody2::GetCoincidenceTransform2(otherBody, transform)` asks the geometry kernel
whether one body can be moved onto another, and returns the transform it used.
The pipeline runs cheapest test first:

1. **Bucket by face count** — an exact integer with no rounding risk, which keeps
   the expensive comparison off almost every pair. 41 comparisons for 56 bodies.
2. **Size gate** — volume and area within one part per thousand. Deliberately
   generous; its only job is to skip pairs that obviously cannot match.
3. **Kernel comparison** — matched with rotation only means the same part;
   matched with a reflection means a mirror. Handedness comes from the
   determinant of the transform's 3×3 rotation part (`MathTransform.ArrayData`
   elements 0–8): `+1` rotation, `−1` reflection. **A mirror is treated as a
   different part** — an opposite hand is not interchangeable on the shop floor.

## Applying colour

Colours are written to the **active display state only**, through
`DisplayStateSetting`. Two things about that are not obvious:

- **Apply one entity at a time.** Passing several at once to
  `DisplayStateSetting::Entities` reports success and silently applies nothing to
  some of them.
- The palette is spread around the hue wheel and **luminance-corrected**, because
  uncorrected blues read as near-black beside yellows — pure blue measures 0.07
  relative luminance against pure yellow's 0.93.

## Propagate visual properties — read this before debugging appearances

Pattern and mirror features have a **"Propagate visual properties"** option, on by
default. When on, regenerating the feature stamps the seed body's colour onto the
derived bodies' **faces** — and a face appearance beats a body appearance, so it
overrides what this macro writes. Re-running does not help.

This produced five separate-looking symptoms during development that all turned
out to be the same cause: colour on the bottom faces of patterned inserts, one
appearance holding 5,301 faces, and removal taking 161 seconds and leaving bodies
coloured. Check that checkbox before suspecting anything else.

## Dead ends

Each was implemented, measured and rejected. **Do not revisit without new
evidence.**

- **Numeric shape invariants for matching** (0.4.7–0.5.18). Principal moments of
  inertia, tensor traces, cubic eigenvalue solvers, vertex-distance hashes,
  chiral pseudoscalars. Fifteen releases, each tuned then reverted, superseded by
  the kernel comparison. The fallback engine remains only for builds without
  `GetCoincidenceTransform2`.
- **Summed vertex hashes** (0.5.12). Fourth-power weighting buried small
  differences in accumulated rounding and made the result depend on vertex count.
  Sorting distances and comparing term by term has neither problem.
- **Chirality heuristics** (0.5.13). Unnecessary — the kernel returns the
  transform and its determinant gives handedness exactly.
- **Batched colour application** (0.6.0). Reports success, applies nothing. See
  above.
- **Reusing one appearance object across a group** (0.9.0). Measured zero
  benefit: 119 ms/body before, 117 ms after. Reads cost nothing.
- **Writing colour straight to the body** (0.10.0). Sixty times faster — 0.11 s
  against 6.55 s — but not scoped to a display state, so colours appeared in
  every display state. Rejected.

## Measured performance

56-body reference weldment:

| Phase | Time |
|---|---|
| Measure | 0.5 s |
| Group (41 kernel comparisons) | 6.0 s |
| Apply (display-state writes, ~117 ms each) | 6.5 s |
| **Total** | **~13 s** |

Naming an entity and reading its appearance back are effectively free. **The
display-state write is the entire cost — optimise write counts, never read
counts.**

## Known limitations

- **Assemblies group by file path and configuration, not geometry.** The matching
  engine runs on part documents only, so two identical parts saved under
  different filenames get different colours.
- A multi-body part inside an assembly gets one colour for the whole component.
- Surface and sheet bodies are skipped.
- **No undo.** Appearance changes leave no undo record.

## Verification status

Confirmed working in SOLIDWORKS on the reference weldment, with the applied
colours read back and checked rather than assumed.

## There is no build step

A `.swp` is a binary VBA project. Editing the `.vba` in `src\` changes nothing
that runs until the source is pasted into the SOLIDWORKS VBA editor and saved.
Treat the `.vba` as the source of truth and re-paste after every change.

Do not patch a `.swp` directly: it stores compiled p-code ahead of the source
text and VBA runs the p-code, so a patched file shows new code in the editor
while still running the old.
