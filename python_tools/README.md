# Planform Statistics Tools (ArcGIS Pro Python Toolbox)

This is a from-scratch Python/ArcPy port of the legacy ArcMap VB.NET add-ins
(`Centerline.vb`, `ArcGISAddin1.vb`, `BankPolys.vb`) in the parent folder. It
runs in ArcGIS Pro (which has no ArcObjects/Add-In SDK), and replaces the
mouse-click selection workflow with ordinary geoprocessing tool parameters, so
it can also be run from the Python window, a script, or batched headlessly.

## Contents

- `PlanformStatisticsTools.pyt` — the toolbox; add this file in the Catalog
  pane (`Insert > Toolbox > Existing Toolbox`). It exposes three tools:
  - **Interpolate Centerline** — port of `Centerline.vb`
  - **Compute Bank Migration** — port of the `Migration` class in `ArcGISAddin1.vb`
  - **Create Bank Polygons** — port of `BankPolys.vb`
- `planform_common.py`, `centerline_tool.py`, `migration_tool.py`,
  `bank_polys_tool.py` — the algorithms, importable independently of the
  toolbox if you want to call them from your own scripts.

Keep all five files in the same folder — the `.pyt` adds its own directory to
`sys.path` so it can import the sibling modules.

## Requirements

- ArcGIS Pro with a **Standard or Advanced** license. The bank-buffer tool
  uses the `Copy Parallel Lines` geoprocessing tool (the ArcPy/GP equivalent
  of the legacy `IConstructCurve.ConstructOffset` call), which requires at
  least a Standard license.
- Each input line must already have a defined spatial reference, and only the
  *first* feature in each input layer/feature class is used (matching the
  original tools, which always selected a single feature per prompt).

## What changed vs. the legacy tools

- **No mouse selection / message boxes.** All bank lines, centerlines, and
  numeric parameters (point spacing, buffer width, angle threshold, etc.) are
  now tool parameters instead of `InputBox`/map-click prompts.
- **Centerline M-values now actually carry data.** The legacy tool created an
  M-aware polyline but called `SetMsAsDistance(False)` without ever assigning
  per-vertex M values, so the M dimension on the output shapefile was
  effectively unused (the downstream-distance values only ever made it into
  the companion `.txt` file). The Python version sets each vertex's M value to
  its downstream distance, so the M dimension is meaningful on its own.
- **Apex-line bend-translation adjustment was not ported.** `ArcGISAddin1.vb`
  declared Bezier-curve objects and had `DrawBezier`/`FindBestBezier`
  subroutines, but tracing the actual call path of `GetMigration` shows they
  were never invoked — the real interpolation always went through
  `FindMidCenterline`/`PolyDivide`. That part is fully ported. The optional
  apex-line correction (`AdjustCenterlineNearApexTrajectory`,
  `CreateBendTranslationSegment`), which still had debug `MsgBox` calls left
  in the original and only ran when a user opted in, was left out of this
  port as an unfinished/experimental feature rather than silently
  reimplementing it without the same level of validation the rest of the
  algorithm has had. If you relied on that option, say so and it can be added.
- **Output schema for bank polygons trims unused fields.** `BankPolys.vb`
  reused the same field-setup routine as the migration tool, so its output
  shapefiles carried several always-empty fields (`Mig_dist`, `m`, `old_m`,
  `Mig_1..4`). The Python port only writes the fields it actually populates
  (`i`, `LB_len`, `RB_len`, `LB_buf_ln`, `RB_buf_ln`).
- **One known quirk preserved on purpose**: when a search arc/normal line
  crosses a bank line more than once, the legacy code picked the crossing
  nearest the *start* of the bank line rather than nearest the current
  centerline point (an apparent leftover bug — a variable meant to track the
  previous position was never updated). This is preserved as-is since it only
  affects results when a bank line is crossed multiple times by one search
  arc, which shouldn't happen for well-behaved, monotonic bank digitizing.

## Usage notes

- **Interpolate Centerline**: pick the left and right bank polylines, a point
  spacing, and a maximum point count. Produces an M-aware polyline feature
  class plus a CSV with `m, width, theta, dtheta, r_curve` and the matching
  centerline/left-bank/right-bank coordinates at each station.
- **Compute Bank Migration**: pick the older and newer centerlines (same
  spatial reference) and a polygon half-width. Produces one polygon per
  segment of the newer centerline, with fields `Mig_dist, i, m, old_m,
  Mig_1..Mig_4` matching the legacy attribute schema.
- **Create Bank Polygons**: pick the centerline, the lidar-derived left/right
  banks (used to build the offset buffer and for the `*_buf_ln` length
  fields), the photo-derived left/right banks (used for the `*_len` length
  fields), a buffer width, a maximum search distance from the channel, and a
  curvature angle threshold (degrees). Produces the main polygon feature
  class plus `*_left`/`*_right` feature classes split along the lidar banks.

## Testing

This environment does not have ArcGIS Pro/ArcPy installed, so the code could
only be checked for valid Python syntax here, not run end-to-end. Please test
each tool in ArcGIS Pro against a known dataset (ideally one you previously
ran through the legacy tool) and compare outputs before relying on it.
