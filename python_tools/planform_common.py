"""
Shared geometry and I/O helpers for the Planform Statistics Tools ArcGIS Pro toolbox.

These functions translate the circular-arc / offset-line / cut-polygon operations
used by the original ArcMap VB.NET add-ins (Centerline.vb, ArcGISAddin1.vb,
BankPolys.vb) into ArcPy geometry calls so the underlying algorithms can run as
ordinary Python without ArcObjects or a mouse-driven UI.
"""
import math
import arcpy

PI = math.pi


def first_geometry(in_features):
    """Return the SHAPE@ geometry of the first feature in a feature class/layer."""
    with arcpy.da.SearchCursor(in_features, ["SHAPE@"]) as cur:
        for row in cur:
            return row[0]
    raise ValueError("No features found in {}".format(in_features))


def make_point(x, y, m=None):
    if m is None:
        return arcpy.Point(x, y)
    return arcpy.Point(x, y, None, m)


def make_polyline(points, sr):
    arr = arcpy.Array([make_point(x, y) for x, y in points])
    return arcpy.Polyline(arr, sr)


def tangent_angle(polyline, distance, as_percentage=False, step=1e-4):
    """Approximate the tangent direction (radians) of a polyline at a given
    position by finite differences. Stands in for ArcObjects' ILine.Angle,
    which the legacy tools derived from a two-point line."""
    length = polyline.length
    if as_percentage:
        d0 = max(distance - step, 0.0)
        d1 = min(distance + step, 1.0)
    else:
        step_dist = step * length if length else step
        d0 = max(distance - step_dist, 0.0)
        d1 = min(distance + step_dist, length)
    p0 = polyline.positionAlongLine(d0, as_percentage).firstPoint
    p1 = polyline.positionAlongLine(d1, as_percentage).firstPoint
    return math.atan2(p1.Y - p0.Y, p1.X - p0.X)


def query_normal_point(polyline, distance, as_percentage, offset):
    """Point offset perpendicular to `polyline` at `distance`. Positive offset
    is to the right of the line's direction of travel, matching ArcObjects'
    IPolyline.QueryNormal convention used throughout the legacy tools."""
    base = polyline.positionAlongLine(distance, as_percentage).firstPoint
    angle = tangent_angle(polyline, distance, as_percentage)
    normal_angle = angle - PI / 2.0
    return arcpy.Point(base.X + offset * math.cos(normal_angle),
                        base.Y + offset * math.sin(normal_angle))


def circular_arc_points(center, start_angle, sweep_angle, radius, n=361):
    """Densified vertices tracing a circular arc, used only to find where the
    arc crosses a bank line (mirrors ICircularArc + ITopologicalOperator.Intersect
    in the legacy findCL routine)."""
    cx, cy = center
    return [
        (cx + radius * math.cos(start_angle + sweep_angle * t / (n - 1)),
         cy + radius * math.sin(start_angle + sweep_angle * t / (n - 1)))
        for t in range(n)
    ]


def arc_point_at_fraction(center, start_angle, sweep_angle, radius, fraction):
    """Point at `fraction` (0..1) along the arc defined by circular_arc_points,
    computed directly instead of re-querying the densified polyline."""
    angle = start_angle + sweep_angle * fraction
    cx, cy = center
    return arcpy.Point(cx + radius * math.cos(angle), cy + radius * math.sin(angle))


def nearest_intersection_point(arc_polyline, bank_polyline):
    """Where arc_polyline crosses bank_polyline. When several crossings exist,
    return the one nearest the *start* of bank_polyline. This mirrors
    GetNearestIntersectionPoint in the original Centerline.vb, which measured
    distance from the bank's start point rather than from the current
    centerline point -- harmless for the typical case of 0 or 1 crossings."""
    pts_geom = arc_polyline.intersect(bank_polyline, 1)
    points = list(pts_geom)
    if not points:
        return None
    best_point = None
    best_frac = None
    for pnt in points:
        _, frac, _, _ = bank_polyline.queryPointAndDistance(pnt, True)
        if best_frac is None or frac < best_frac:
            best_frac = frac
            best_point = pnt
    return best_point


def construct_offset_polyline(polyline, distance, sr):
    """Parallel offset line, equivalent to ArcObjects' IConstructCurve.ConstructOffset.
    Uses the Copy Parallel Lines geoprocessing tool (Standard/Advanced license)
    since ArcPy geometry objects have no built-in offset method."""
    side = "RIGHT" if distance >= 0 else "LEFT"
    in_fc = "in_memory/offset_src"
    out_fc = "in_memory/offset_out"
    for fc in (in_fc, out_fc):
        if arcpy.Exists(fc):
            arcpy.management.Delete(fc)
    arcpy.management.CreateFeatureclass(
        "in_memory", "offset_src", "POLYLINE", spatial_reference=sr)
    with arcpy.da.InsertCursor(in_fc, ["SHAPE@"]) as cur:
        cur.insertRow([polyline])
    arcpy.management.CopyParallelLines(
        in_fc, out_fc, abs(distance), bevel_ratio=10,
        line_side=side, line_position="OUTER", merge_field="")
    offset_geom = first_geometry(out_fc)
    arcpy.management.Delete(in_fc)
    arcpy.management.Delete(out_fc)
    return offset_geom


def closest_point(ref_point, candidate_points):
    """Nearest point in candidate_points to ref_point (port of BankPolys.vb's
    ClosestPt)."""
    return min(
        candidate_points,
        key=lambda p: math.hypot(p.X - ref_point.X, p.Y - ref_point.Y))


def intersection_points(geom_a, geom_b):
    """List of arcpy.Point from a 0-dimension (point) intersection."""
    return list(geom_a.intersect(geom_b, 1))


def create_feature_class(out_path, geometry_type, sr, fields, has_m=False):
    """Create a new feature class at out_path with the given geometry type and
    a list of (name, type, precision, scale, length) field specs. Mirrors the
    role of CreateNewShapefile/SetupFields in the legacy tools, but works with
    both shapefiles and geodatabase feature classes depending on out_path."""
    out_dir = arcpy.management.CreateFeatureclass(
        *_split_path(out_path),
        geometry_type=geometry_type,
        spatial_reference=sr,
        has_m="ENABLED" if has_m else "DISABLED")
    fc = out_dir.getOutput(0)
    for name, ftype, precision, scale, length in fields:
        arcpy.management.AddField(
            fc, name, ftype,
            field_precision=precision, field_scale=scale, field_length=length)
    return fc


def _split_path(out_path):
    import os
    return os.path.dirname(out_path), os.path.basename(out_path)
