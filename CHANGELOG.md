# Changelog

Semantic Versioning. `MAJOR` reaches 1 when the behaviour is settled enough
to promise not to break it; 0.x is an honest statement that it may still move.

---

Releases up to and including `V5.17` used a `V<minor>.<patch>` scheme. They map onto Semantic Versioning directly, so the old `V4.10` is now `0.4.10`.

Source for `0.2.0`, `0.5.7` and `0.5.9` was not preserved.

## 0.11.4 — 2026-08-21

- Moved to its own repository, so the macro is versioned and released on its
  own rather than alongside the two removal macros. The `Source` URL in the
  header points at the new repository.
- No functional change.

## 0.11.3 — 2026-08-20

- The `Source` URL in the header now points at the renamed repository,
  `apply-colours-sw-macro`. No functional change.

## 0.11.2 — 2026-08-13

- The completion dialog reported the previous version number. The version is held
  in two places — the header comment and a `MACRO_VERSION` constant — and only the
  header was updated in 0.11.1. `build-library.ps1` now fails the build if the two
  disagree.

## 0.11.1 — 2026-08-09

- Released under the **MIT licence**, with the full text carried in the code
  itself so a `.swp` passed on by itself still carries its licence. No functional
  change.

## 0.11.0

- **Warns about features that will undo the colouring**: Pattern and mirror features with "Propagate visual properties" switched on copy their seed body's colour onto the bodies they create, every time the model rebuilds. The copy is written onto the derived body's *faces*, and a face appearance beats the body appearance this macro writes - so re-running the macro does not help, and neither does applying a colour by hand. Confirmed by applying one manually and watching a rebuild overwrite it. Nothing can be done about it from here, so the macro counts those features instead and says so on completion. Without that, a mirrored body quietly reverting to its original's colour after an unrelated rebuild is close to undiagnosable.
- The cure is the checkbox in the feature's own options, left as the user's decision but worth switching off: propagation is redundant here, because identical bodies are grouped and coloured alike no matter how they were created. Turning it off loses nothing on a pattern and keeps mirrored bodies distinct, which propagation otherwise prevents. Part documents only for now, since that is where the behaviour has been observed.
- The same propagation explains a chain of symptoms previously chased separately: colour appearing on the bottom faces of patterned inserts, one appearance holding 5,301 faces, and the companion removal macro taking 161 seconds and leaving bodies coloured afterwards. All of it was pattern regeneration stamping the seed's colour onto derived faces.

## 0.10.2

- **No user settings**: The four behaviour toggles are gone and their behaviour is fixed in the code. `MATCH_MIRRORED` and `AVOID_EXISTING_FACE_COLOURS` were settled decisions - mirrored bodies get their own colour, existing face colours are avoided - and an option nobody should ever change is an invitation to break something. `USE_KERNEL_COMPARISON` duplicated a check the macro already makes for itself, and `SHAPE_FALLBACK` had no reason to be off. What remains is `SHOW_DIAGNOSTICS`, plus two named tolerances. The reasoning behind each fixed behaviour is kept as a comment where that behaviour lives.
- **Diagnostics off by default**: `SHOW_DIAGNOSTICS` now defaults to off, which also drops the verification pass and the per-body edge count query. Turn it on when a body is grouped or coloured unexpectedly - the Immediate window detail names which bodies were compared and why they did or did not match.
- **Header brought up to date**: Describes the two-stage matching and the treatment of mirrored bodies, and the summary changelog covers 0.6 through 0.10.

## 0.10.1

- **Direct write removed**: The 0.10.0 experiment is settled and the setting is gone. On a 56 body weldment the direct write applied colours in 0.11 s against 6.55 s through the display state, cutting the whole run from 13.24 s to 7.08 s - but it is not scoped to a display state. The verification pass saw only 3 of 56 colours, because it reads back through the display state and the colour was not there, and the colours turned up in every other display state in the document. Colours belonging to the display state they were applied in is the entire point of applying them this way, so a sixty-fold speed-up does not buy it back. Writing straight to the body remains only as a fallback for builds where display states are unavailable, exactly as it was before 0.10.0.
- **Where the time goes, established**: Naming an entity and reading its appearance back are both free - 56 of each complete inside the resolution of the VBA timer. The document write is nearly free too. The display state machinery is the entire cost of applying colour, at roughly 117 ms per body. Recorded here so the experiment need not be repeated.
- Behaviour is now identical to 0.9.0, less `FAST_APPLY`.

## 0.10.0

- **Direct colour write, for measurement**: `WRITE_METHOD` selects between writing through the active display state and writing straight onto the body or component. Whichever is chosen falls back to the other if it fails. Method 1 is set for testing; method 0 is the behaviour of every release since 0.4.5 and remains the safe choice, because a direct write may attach the colour to the configuration rather than the active display state.
- **`FAST_APPLY` removed**: 0.9.0 reused one appearance object across a group to avoid reading it back per body. Measured across two runs it saved nothing - 119 ms per body before, 117 ms after - because the read costs nothing to start with. The verification pass established why: naming an entity and reading its appearance back completed for 56 bodies inside the resolution of the VBA timer, so the display state *write* is the entire cost of applying colour. The reuse was removed rather than kept, since it carried a real risk of sharing an appearance object across bodies in exchange for a measured gain of zero.

## 0.9.0

- **Faster colour application**: Applying a colour took three API calls per body: naming the entity, reading its current appearance back, and writing the modified one. Every body in a group receives the same colour, so the appearance object read for the first body is now reused for the rest, removing one round trip per body after the first. Controlled by `FAST_APPLY`.
- **Applied colours are verified**: A colour application can report success without having changed anything, which is how the 0.6.0 batching fault stayed hidden until it was spotted by eye in the model. On diagnostic runs the macro now reads every colour back and reports how many actually took, timed separately so it does not distort the figure it is checking. This is what makes it safe to try faster ways of applying colour: a silent failure shows up as a number rather than as a body that happens to look wrong.

## 0.8.1

- **Kernel comparison silently disabled in 0.8.0**: While restructuring the comparison to report *how* two bodies matched, the call was rewritten as `If Not bodyA.GetCoincidenceTransform2(...) Then` to give an early exit. Applying `Not` directly to that call does not invert its result as it appears to, and every comparison came back as no match. The whole kernel engine was bypassed and grouping fell through to the shape fallback, which is mirror-blind - so mirrored bodies kept sharing a colour even though 0.8.0 was written to separate them. The result is now taken into a `Boolean` before being tested.
- **Verdict breakdown in the diagnostics**: Comparisons are reported as matched outright, matched as mirror, and no match, rather than a single "not coincident" count. The 0.8.0 fault showed as 37 comparisons of which 37 were not coincident, which the previous summary made easy to read past.

## 0.8.0

- **Mirrored bodies are now kept apart**: The geometry kernel counts a mirror image as coincident, so a mirrored copy of a body was silently sharing its original's colour. It also hands back the transform it used, and a reflection is exactly identifiable from it: the determinant of a rotation matrix is always +1, and a reflection makes it -1. The macro now reads that determinant and treats a mirror as a different part. No handedness heuristic is involved, which is what made this unreliable when it was attempted in 0.5.13 and abandoned in 0.5.15.
- **`MATCH_MIRRORED` now does what its name says**: It previously tried to *add* mirror matching by transforming a copy of the body, which was redundant since the kernel already matched mirrors, and the setting had no observable effect. It now controls whether the kernel's mirror match is accepted. Off by default: a mirrored part is usually a genuinely different part. The body-copying code it relied on has been removed.
- **The shape fallback is fenced off from mirrors**: Distances from the centre of mass are identical either side of a mirror, so the 0.7.0 fallback would have re-merged any mirror the kernel matched and we then rejected. It now runs only where the kernel found no transform at all.

## 0.7.0

- **Shape comparison as a second opinion**: `GetCoincidenceTransform2` carries a tolerance of its own, and copies of one imported part that were brought in through separate insert-part features fall outside it. Measured on a weldment: the same rivet nut derived twice differed by nine parts per million in volume and was split into separate colours, while a pair derived three parts per million apart matched. Where the kernel reports two bodies are not coincident, the macro now compares their shape directly - the sorted distance of every vertex from the body's centre of mass, matched term by term against a one micron tolerance. Those distances are unchanged by moving or rotating a body, but a feature that moves takes its vertices with it.
- **Why aggregate measurements cannot do this**: On the same model, two tubes that differ only in where a hole sits had volumes agreeing to fifteen decimal places, because moving a hole removes exactly as much material as before. The genuinely identical rivet nuts differed a billion times more than the genuinely different tubes did. Any threshold on volume or area therefore merges the tubes before it merges the nuts, which is the trap every release from 0.4.7 to 0.5.18 fell into. Comparing vertex positions has no such blind spot: measured on representative geometry, a hole moved 5 mm shifts the signature by 2.5 mm while re-derivation noise shifts it by 30 nm, either side of the tolerance by a factor of about 2500.
- **Sorted, not summed**: 0.5.12 tried a related idea and failed because it added the vertex distances together with fourth-power weighting, which buried small differences in accumulated rounding and made the result depend on how many vertices a body happened to have. Holding the distances individually, sorting them into a canonical order and comparing term by term avoids both problems.
- **Diagnostics**: Every pair the shape comparison rescues is logged beneath the kernel's rejection, and the summary counts them separately, so what the fallback merged can always be audited.
- Two new settings: `SHAPE_FALLBACK` and `SHAPE_TOLERANCE`.

## 0.6.3

- **Size gate rejected genuine matches**: The cheap volume and area check that runs before the kernel comparison was set at one part per million, on the assumption that identical bodies agree to near machine precision. That assumption does not hold for imported hardware: where the same part is brought in through several separate insert-part features, each derivation produces its own B-rep, and computed volumes were observed to differ by twelve parts per million on bodies with identical face and edge counts. Those pairs were discarded before the kernel was ever asked about them, so repeated hardware was split across several colours. The gate is now one part per thousand, which still discards the great majority of pairs while leaving the decision with the kernel.
- **Comparison diagnostics**: The readout now separates kernel comparisons from pairs skipped by the size gate, and reports how many comparisons the kernel judged not coincident. Each of those is logged with the two body names and their relative volume and area difference, which distinguishes "the gate blocked it" from "the kernel looked and said no".

## 0.6.2

- **Per-body diagnostic dump**: With `SHOW_DIAGNOSTICS` on, the macro writes one line per body to the VBA Immediate window giving its group, face count, edge count, volume, surface area and body name. This is for answering why two bodies were or were not matched, which the group-level summary could not show.

## 0.6.1

- **Repeated hardware left uncoloured**: 0.6.0 applied one colour per group by passing every body in the group to `DisplayStateSetting::Entities` at once. That call reports success but silently applies nothing when given more than one entity, and because it did not raise an error the fallback path never ran. Single-body groups were unaffected, so on a weldment the unique tubes were coloured while every repeated nutsert was skipped. Colours are applied one entity at a time again, as they were up to 0.5.18.
- **Colours appeared only after the dialog was dismissed**: The graphics redraw happened after the completion dialog was closed, so the model still looked untouched while the dialog was on screen. The view is now redrawn before the dialog opens.
- **Silent application failures are now visible**: The diagnostics readout reports how many bodies or components actually received a colour, against the total. A mismatch is what would have made the 0.6.0 fault obvious on the first run.

## 0.6.0

- **Body matching moved into the geometry kernel**: Parts are now grouped with `IBody2::GetCoincidenceTransform2`, which reports whether one body can be moved onto another. This is an exact test rather than a numeric approximation, so it removes the accuracy trade-off that every release from 0.4.7 onward was trying to tune. Two tubes whose holes sit in slightly different places no longer match at any feature size, while identical hardware at arbitrary positions and orientations still does. If a build does not expose the method, the macro falls back to the 0.5.18 invariant comparison automatically.
- **Luminance-corrected palette**: Red, green and blue contribute roughly 21%, 72% and 7% of perceived brightness, so plain HSV at full value made blues and violets read as near-black beside yellows — pure blue measures 0.07 relative luminance against pure yellow's 0.93. Colours are now generated at a fixed perceived brightness per layer: hue is darkened by reducing value, or lightened by reducing saturation, so every colour in a layer matches within rounding. Layer brightness spans 0.25 to 0.72, which drops the near-black and near-white bands entirely.
- **Two groups can no longer share a colour**: Hue collision avoidance previously checked only against colours already in the model, so two groups nudged away from the same obstacle could land on the same hue. Each hue is now also checked against every hue already assigned, and if nothing is clear the evenly spaced hue is kept rather than duplicating another group.
- **Assembly re-runs no longer degrade**: `ProcessAssembly` seeded its "colours to avoid" list from component-level appearances, which is exactly what the macro writes. A second run therefore treated its own output as off-limits, exhausted the search and collapsed the palette. Component colours are no longer read back.
- **Speed**: The second mass-properties evaluation per body is gone with the invariant engine. Bodies are bucketed by face count so the expensive comparison only runs on plausible candidates. Colours are applied once per group instead of once per body, and graphics updates are suspended for the duration, then restored — including when the macro fails.
- **Counters widened to `Long`**: `Integer` counters overflowed above 32,767 bodies or components.
- **Unit pinning**: The fallback engine sets `IMassProperty::UseSystemUnits`, so its results no longer shift with the document's unit scheme.
- **Settings block**: Comparison engine, mirror matching, face-colour avoidance and a diagnostics readout are exposed as constants at the top of the file.

## 0.5.18

- **Versioning**: Adopted Semantic Versioning and renumbered the release history.
- **Documentation**: Rewrote the macro header comment as user-facing documentation — what the macro does, how it groups geometry in parts and in assemblies, what it modifies, and how to run it — in place of the previous development notes. Rewrote this changelog in plain language.
- **File naming**: Renamed sources to `Apply Unique Colours <version>.txt` and standardised the spelling on "Colours" throughout.
- No functional change to the macro.

## 0.5.17

- **Principal moments read directly**: Replaced the copy-and-translate approach from 0.5.16 with `IMassProperty.PrincipleMomentsOfInertia`. Copying each body and applying a transform to move it to the origin could trigger re-tessellation on dense imported geometry, which changed the computed values between bodies that were otherwise identical. Principal moments are already taken about the centre of mass, so the translation was unnecessary.
- **Normalised by mass**: Divided the principal moments by body mass to give the squared radius of gyration. This removes any dependence on the material density assigned to a body, and leaves a purely geometric quantity that is invariant under translation and rotation.
- Grouping tolerance set to 0.01% relative plus 5e-9 absolute.

## 0.5.16

- **Vertex hash removed**: Dropped the vertex hash retained from 0.5.12. Imported models with fine tessellation, such as knurled inserts, produced vertex sets that varied with body orientation, so the hash split parts that were in fact identical. Grouping now relies on the tensor invariants J1, J2 and J3 alone, which are invariant under rotation.

## 0.5.15

- **Chiral detection removed**: Dropped the chiral hash from 0.5.13 and the face-centroid sampling from 0.5.14. Off-the-shelf hardware is routinely mirrored during assembly modelling, which flips its chirality; the chiral test treated those instances as different parts and split them. Tolerance tightened to 1e-10.

## 0.5.14

- **Face-centroid sampling**: Added face centroids to the topological hash, to catch cylindrical hole faces that have no start or end vertices for the vertex hash to pick up.
- **Density scaling**: Multiplied the density passed to `GetMassProperties` by 1000. At the default density of 1.0 the inertia difference produced by a small displaced hole is on the order of 1e-11 kg·m², below the practical precision floor; scaling the density lifts the difference clear of it.

## 0.5.13

- **Chiral hash**: Added a test to separate mirror-image bodies. Bodies that are exact mirrors — a tube with a hole offset 50 mm to the left, and the same tube with it offset 50 mm to the right — match on every invariant the macro computed: volume, area, tensor traces and centre of mass. The hash sums weighted vertex distances from the centre of mass into three direction vectors and takes their triple scalar product, which changes sign under mirroring but not under rotation.

## 0.5.12

- **Vertex hash**: Long symmetric weldment sections mask a repositioned hole; the volume change is under 0.0001% and the inertia tensor barely moves. Added a hash summing the Euclidean distance of every B-Rep vertex from the body's centre of mass — invariant under translation and rotation, but changed when a feature moves.

## 0.5.11

- **Tensor invariants instead of cubic roots**: Replaced the discriminant clamp from 0.5.9 with direct comparison of the invariants J1, J2 and J3. On rectangular tube sections the Y and Z axes are close to degenerate, so solving the characteristic cubic forced near-identical roots together and erased the small differences produced by offset holes. The invariants are rotationally invariant and need no root solving.

## 0.5.10

- **Transform to origin in memory**: Copied each body, applied a transform placing it at the origin, read its mass properties, then discarded the copy. Bodies positioned thousands of millimetres from the document origin lost the small decimal differences that distinguish off-centre holes when the parallel axis transformation was applied. The user's document is not modified.

## 0.5.9

- **Compatibility fallback**: Reverted to `IBody2::GetMassProperties(1.0)`. Clamped the cubic discriminant to zero when floating-point noise pushed it out of range, rather than aborting the calculation, so bodies that previously returned zero tensors are processed.

## 0.5.8

- **Distant bodies returning no inertia**: `IBody2::GetMassProperties` returns tensors about the global origin, so the parallel axis transformation for a body far from it lost precision. The resulting noise pushed the cubic solver's discriminant positive, and the solver aborted to avoid imaginary roots — leaving those bodies unprocessed entirely. Switched to `IModelDocExtension::CreateMassProperty`.

## 0.5.7

- **Configuration in the component key**: `ProcessAssembly` matched components on `GetPathName()` alone, so two configurations of the same part were treated as one component and shared a colour. The referenced configuration is now part of the key.

## 0.5.6

- **Edge count, restricted by volume**: Reinstated the edge count check for bodies above 1e-5 m³ (roughly 10 g). Mirror-image beams have identical principal moments and need a topological check, but applying that check to small hardware caused the false splits seen in 0.5.2. Noise floor widened to 1e-9.

## 0.5.5

- **Unit error in the tolerance**: The SolidWorks mass properties dialog displays in document units (g·mm²), but the `IMassProperty` API returns system units (kg·m²). A tolerance intended as 1 g·mm² was therefore acting as 1 kg·m² — a billion times larger — which merged tubes of different lengths. Rescaled to 1e-10.

## 0.5.4

- **Absolute floor plus relative term**: Replaced proportional tolerances with `mX * 1e-7 + 1.0`. Small features far from the origin introduced up to 0.01% relative noise when global tensors were converted to local moments.

## 0.5.3

- **Edge count reverted**: Removed the `GetEdgeCount()` check from 0.5.2. SolidWorks builds B-Rep faces through varying boolean sequences, so identical geometry — standard threaded inserts, for example — reported different edge counts when surfaces split against the weldment frame.

## 0.5.2

- Added `GetEdgeCount()` to the part grouping key.
- Tightened the principal moment tolerances to 1e-9.

## 0.5.1

- **Wider hue separation**: Reduced items per saturation/brightness layer from 21 to 8, increasing the hue gap between bodies sharing a layer.
- Tightened principal moment tolerances by a factor of 500.
- Added the macro version to the completion dialog.

## 0.5.0

- Version bumped after confirming the mass properties grouping worked across SolidWorks versions.
- Completion dialog reformatted to list total bodies, unique groups and skipped items on separate lines.
- Condensed the header comment to major releases only.

## 0.4.10

- **VBA tensor solver**: Older SolidWorks versions returned error 438 from the COM interface when retrieving specific moment vectors. Replaced the API call with a 3x3 tensor solver written in VBA, evaluating the characteristic cubic directly.

## 0.4.9

- **Interface rollback**: Reverted from `IMassProperty2` to `IMassProperty`. Older SolidWorks type libraries do not carry `AddBodies` through interface inheritance, which caused error 438.

## 0.4.8

- **Equidistant hues**: Replaced the golden-angle distribution with pre-calculated spacing — the macro counts unique groups and divides 360° by that number, giving the maximum available hue separation.
- **Interleaved layers**: Bucketed groups into layers of at most 21, each on a different saturation and brightness combination, so that similar hues do not read as identical under SolidWorks shading.

## 0.4.7

- **Principal moments of inertia**: Added Px, Py and Pz to the part grouping key via `IMassProperty2`. These are rotationally invariant but sensitive to where mass sits, so plates with identical volume, area and face count but holes in different positions now separate.

## 0.4.6

- **Display state COM crash**: Fixed automation error -2147417848 when instantiating late-bound display states on older SolidWorks versions, by using typed `swDisplayStateOpts_e` values and declared SolidWorks interfaces.

## 0.4.5

- **Display state targeting**: Reworked colour application to target the active display state through `DisplayStateSpecMaterialPropertyValues`, so the macro can be run independently in each display state under a single configuration.
- Removed the string-based selection mechanism, which was the source of error 438.

## 0.4.4

- **Selection errors removed**: Dropped `Entity.Select4`, which caused intermittent error 438 and error 13 on some SolidWorks builds, in favour of `SelectByID2` string matching.
- Re-enabled `MaterialPropertyValues` targeting: display states are linked to configurations, so applying to the active configuration reaches the paired display state.

## 0.4.3

- Added step tracking to the error handlers, so a failure reports which stage it occurred in.
- Replaced `Nothing` arguments to COM calls with empty object instances, and switched hard-cast `SldWorks.Entity` variables to late-bound `Object`, both to avoid interface mismatches on older versions.

## 0.4.2

- Fixed a silent runtime failure caused by passing String and Double arrays to methods expecting Variant arrays.
- Used the explicit `swSpecifyConfiguration` constant with typed Variant arrays for configuration names.

## 0.4.1

- **Per-configuration targeting**: Replaced implicit configuration targets with explicit entity selections and configuration name arrays. Passing `Empty` had caused colours to apply to every configuration rather than the active one.

## 0.4.0

- **Golden angle hues**: Switched hue assignment to the golden angle (137.5°), so sequentially discovered parts receive widely separated hues.
- **Saturation and brightness variation**: Cycled through 7 combinations. At 200 parts, dividing the wheel alone gives 1.8° steps, which are not distinguishable; varying lightness makes them so.

## 0.3.1

- **Leaf parts only**: Restricted assembly colouring to bottom-level parts, rather than applying overrides to parent subassemblies.

## 0.3.0

- **Face colour checks**: Added detection of existing face-level colours, and shifted generated colours away from them.
- **Assembly support**: Components grouped by reference, with appearances applied at component level.
- **Active configuration**: Targeted `swThisConfiguration` so other configurations keep their appearance.

## 0.2.0

- **Body grouping**: Grouped bodies by volume, surface area and face count, so identical bodies share a colour.

## 0.1.0

- Initial release. Applies a sequentially spaced colour to every solid body in the active part, generated in HSV.

---
