# PlanformStatisticsTools

Tools for measuring the planform characteristics of river channels — width, curvature, and lateral migration — at evenly spaced points along a channel, from bank lines or channel polygons digitized from aerial imagery.

> **⚠️ Status: active conversion to Python (ArcGIS Pro), not fully tested.**
> This repository is being converted from its original VB.NET/ArcObjects add-in (ArcMap 10.x) to a Python toolbox (`.pyt`) for ArcGIS Pro. The ported tools are functional but **have not been fully validated** against the original outputs. Use results with care and verify against known cases. New capabilities for wandering / anabranching rivers are in active development and are **not yet present** in the toolbox (see [Roadmap](#roadmap)).

---

## Overview

These tools automate the measurement of width, curvature, and migration rate at discrete, evenly spaced points along a single-thread river.

The toolbox provides three primary tools:

1. **Centerline** — interpolate a centerline at evenly spaced intervals from two digitized bank lines, recording width and a first-order radius of curvature at each point.
2. **Migration Distance** — estimate the mean lateral (normal) migration distance between two centerlines produced by Tool 1 (e.g., the same reach at two points in time).
3. **Bank Buffer Boxes** — generate polygons adjacent to the channel banks, keyed to each centerline point, useful for correlating a bank property with one of the observed planform statistics.

---

## Requirements

- **ArcGIS Pro 3.x** with an active license (Standard or Advanced).
- **`arcpy`** — ships with ArcGIS Pro; no `conda` installs or third-party packages are required for the current toolbox.
- No extensions (e.g., Spatial Analyst) are required.

---

## Installation

The toolbox is a single ArcGIS Pro Python Toolbox file, `PlanformStatisticsTools.pyt`. Unlike the legacy ArcMap add-in, there is nothing to compile or install — ArcGIS Pro treats the `.pyt` as a toolbox when you add it to a project.

1. Copy `PlanformStatisticsTools.pyt` into your ArcGIS Pro project's home folder (or any folder on disk).
2. In ArcGIS Pro, open the **Catalog** pane (*View* tab → *Catalog Pane*).
3. Under **Folders**, browse to the folder containing the `.pyt` (add a *Folder Connection* if needed). The file appears with a toolbox icon; expand it to see the three tools.
   - *Alternatively:* right-click **Toolboxes** in the Catalog pane → **Add Toolbox** → browse to `PlanformStatisticsTools.pyt`.
4. Double-click a tool to open its geoprocessing dialog.

> The original compiled ArcMap 10.x add-in (`.esriAddIn`) is preserved under [`legacy/`](legacy/) for users still working in ArcMap.

---

## Tools

### 1. Centerline

Finds evenly spaced points representing the center of two roughly parallel bank lines and connects them into a new line. From each previous point, the tool places a new point a user-specified distance away and varies the local angle until the distance to the closest point on each bank line is nearly equal, producing a smooth, evenly spaced centerline.

- **Inputs:** two bank polylines (oriented in the **same**, downstream direction) and a point spacing.
- **Recommended spacing:** about half a channel width. Much smaller spacings can cause the interpolated centerline to turn back on itself.
- **Output:** an M-aware centerline plus a table with the columns below.

| Column | Meaning |
|---|---|
| `OID` | Object ID |
| `m` | Down-channel distance |
| `width` | Local channel width (a + b) |
| `theta` | Local angle of the downstream segment, θᵢ |
| `dtheta` | Local change in angle, θᵢ − θᵢ₋₁ |
| `r_curve` | First-order radius of curvature, Δs/dθ (stores `-99999` where dθ = 0) |
| `cl_x`, `cl_y` | Centerline point coordinates |
| `left_x`, `left_y` | Corresponding left-bank point |
| `right_x`, `right_y` | Corresponding right-bank point |

> **Curvature note:** local curvature is best estimated as `dθ/ds = (θᵢ − θᵢ₋₁)/Δs` rather than inverting `r_curve` (which carries the `-99999` sentinel). First-order curvature is noisy for most digitized banks; smooth it or use a higher-order scheme before further analysis.

### 2. Migration Distance

Estimates the average lateral normal distance between the nodes of one centerline (from Tool 1) and a second centerline. Each line is split into segments at the midpoints between intersections; three intermediate centerlines are interpolated between the two inputs, and migration trajectories are built as four straight segments approximately normal to those intermediate lines. (This replaced the Bézier-curve method of the ArcGIS 8/9 releases, greatly reducing computation time.)

- **Inputs:** a "to" centerline and a reference centerline; output polygon name.
- **Output:** a polygon feature class with the columns below, centered on the nodes of the storage centerline (usually the newer one).

| Column | Meaning |
|---|---|
| `Mig_dist` | Total migration (sum of `Mig_1`–`Mig_4`) |
| `Mig_1`–`Mig_4` | Lengths of the four trajectory segments (1 nearest the storage line, 4 farthest) |
| `i` | Index of the apex on the storage centerline |
| `m` | Down-channel distance of the trajectory origin (storage centerline) |
| `old_m` | Down-channel distance of the trajectory endpoint (other centerline) |

> **Known limitation:** the optional **apex-trajectory adjustment** for bends that translate primarily downstream (v2.0 prompted for a digitized apex-trajectory shapefile to force intermediate centerlines onto the migration path) is **not yet reimplemented** in the Python port. It is scaffolded as an optional parameter and is a near-term roadmap item.

### 3. Bank Buffer Boxes

Creates a sampling corridor of user-specified width on the upland side of each bank line, then subdivides it with lines projected outward from the centerline (which should have equal-length segments, as produced by Tool 1). Subdividing lines are projected normal to the centerline except where curvature exceeds a threshold angle (10° works well), which prevents the lines from missing the corridor edge on the inside of sharp bends.

- **Inputs:** a centerline; corridor (buffer) width; two sets of bank lines (the first used only to measure bank length within boxes, the second to build the boxes — they may be identical but must extend past the centerline endpoints); a maximum expected centerline-to-corridor-edge distance; and a threshold angle.
- **Output:** three polygon feature classes — combined/union boxes (the chosen name), plus `_left` and `_right` bank boxes.

---

## Test data

[`tests/data/`](tests/data/) contains eight channel-polygon fixtures used to develop and validate the new wandering-river tools — four cases × two epochs (2006 and 2019):

| Case | Files |
|---|---|
| No islands (single-thread baseline) | `ChannelPoly_NoIslands_2006.json`, `…_2019.json` |
| One island | `ChannelPoly_OneIsland_2006.json`, `…_2019.json` |
| Two islands | `ChannelPoly_TwoIslands_2006.json`, `…_2019.json` |
| Multiple islands | `ChannelPoly_MultipleIslands_2006.json`, `…_2019.json` |

All fixtures are in **EPSG:6596** (NAD83(2011) StatePlane Washington North, meters). The GeoJSON `crs` tag is not reliably preserved on read, so load with the CRS set explicitly:

```python
import geopandas as gpd
gdf = gpd.read_file("tests/data/ChannelPoly_OneIsland_2006.json").set_crs(6596, allow_override=True)
```

See [`tests/data/README.md`](tests/data/README.md) for details.

---

## Repository structure

```
PlanformStatisticsTools.pyt   ArcGIS Pro Python toolbox (current)
tests/data/                   Channel-polygon test fixtures (8 GeoJSON, EPSG:6596)
legacy/                       Original ArcMap 10.x VB.NET add-in and documentation
  ├── Centerline.vb
  ├── ArcGISAddin1.vb
  ├── BankPolys.vb
  ├── README.md
  └── PlanformStatistics2.0.distribute.zip   (compiled .esriAddIn + v2.0 PowerPoint)
LICENSE                       MIT
```

The original v2.0 documentation slide deck (`PlanformStatisticsTools v. 2.0 (for ArcGIS 10).ppt`) is inside the distribution zip in `legacy/` and remains the most complete description of the underlying algorithms.

---

## Roadmap

Development is focused on extending the toolbox beyond single-thread channels:

- **Wandering / anabranching rivers** — accept channel polygons (with islands as interior rings), delineate islands, and resolve multiple flow paths into a continuous main centerline plus attributed secondary threads, with simple summary statistics (anabranching length/fraction, island counts and areas).
- **Migration vectors and bend kinematics** — node-to-node migration with bend decomposition and cutoff/elongation detection.

---

## Version history

- **ArcGIS 8 / 9** — original Planform Statistics Toolbox (Bézier-curve migration trajectories).
- **ArcGIS 10.x add-in, v2.0 (2012)** — VB.NET/ArcObjects release; four-segment intermediate-centerline migration. Preserved in [`legacy/`](legacy/) and tagged [`v1.0-arcmap`](../../releases/tag/v1.0-arcmap).
- **ArcGIS Pro (current)** — Python toolbox port and ongoing extension to wandering-river analysis.

---

## License

Released under the [MIT License](LICENSE).

The tools are provided **free of charge and "as is," without warranty**. Use at your own risk; users assume all responsibility for results and their application.

---

## Author & citation

**J. Wesley Lauer**, Department of Civil and Environmental Engineering, Seattle University — `lauerj@seattleu.edu`

Original development supported by the U.S. National Science Foundation through the National Center for Earth-surface Dynamics (NSF Grant OCE-0742476) and by Seattle University.