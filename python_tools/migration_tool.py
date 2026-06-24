"""
Port of the Migration tool in ArcGISAddin1.vb: measures lateral migration
distance between an older and a newer channel centerline by interpolating
three evenly-spaced intermediate centerlines (at 25%, 50%, and 75% of the
time interval) and summing the four leg distances between them.

Note: the legacy file declared IBezierCurve3 objects and had DrawBezier /
FindBestBezier subs, but the actual migration computation path (GetMigration)
never calls them -- it relies on FindMidCenterline/PolyDivide instead. They
are not ported here since they were unreachable dead code.
"""
import math
import os

import arcpy

from planform_common import PI, intersection_points


def find_mid_centerline(old_line, new_line, sr):
    """Centerline interpolated half-way (in fraction-of-length terms) between
    old_line and new_line. Port of FindMidCenterline."""
    crossing_pts = intersection_points(old_line, new_line)
    crossing_pts.sort(key=lambda p: new_line.queryPointAndDistance(p, True)[1])

    old_mid_pts = [old_line.firstPoint]
    new_mid_pts = [new_line.firstPoint]
    for prev_pt, pt in zip([None] + crossing_pts[:-1], crossing_pts):
        if prev_pt is None:
            continue
        start_old = old_line.queryPointAndDistance(prev_pt, True)[1]
        end_old = old_line.queryPointAndDistance(pt, True)[1]
        old_mid_pts.append(old_line.positionAlongLine(
            (start_old + end_old) / 2.0, True).firstPoint)

        start_new = new_line.queryPointAndDistance(prev_pt, True)[1]
        end_new = new_line.queryPointAndDistance(pt, True)[1]
        new_mid_pts.append(new_line.positionAlongLine(
            (start_new + end_new) / 2.0, True).firstPoint)
    old_mid_pts.append(old_line.lastPoint)
    new_mid_pts.append(new_line.lastPoint)

    mid_points = []
    for j in range(1, len(old_mid_pts)):
        start_old = old_line.queryPointAndDistance(old_mid_pts[j - 1], True)[1]
        end_old = old_line.queryPointAndDistance(old_mid_pts[j], True)[1]
        seg_old = old_line.segmentAlongLine(start_old, end_old, True)

        start_new = new_line.queryPointAndDistance(new_mid_pts[j - 1], True)[1]
        end_new = new_line.queryPointAndDistance(new_mid_pts[j], True)[1]
        seg_new = new_line.segmentAlongLine(start_new, end_new, True)

        mid_points.extend(poly_divide(seg_old, seg_new))

    return arcpy.Polyline(arcpy.Array([arcpy.Point(x, y) for x, y in mid_points]), sr)


def poly_divide(line1, line2):
    """Point-by-point average of two polylines, sampled at evenly spaced
    fractions along each. Port of PolyDivide."""
    max_points = max(_vertex_count(line1), _vertex_count(line2))
    points = [_midpoint(line1.firstPoint, line2.firstPoint)]
    for i in range(1, max_points + 1):
        fraction = i / (max_points + 1)
        p1 = line1.positionAlongLine(fraction, True).firstPoint
        p2 = line2.positionAlongLine(fraction, True).firstPoint
        points.append(_midpoint(p1, p2))
    points.append(_midpoint(line1.lastPoint, line2.lastPoint))
    return [(p.X, p.Y) for p in points]


def _vertex_count(polyline):
    return sum(part.count for part in polyline)


def _midpoint(p1, p2):
    return arcpy.Point((p1.X + p2.X) / 2.0, (p1.Y + p2.Y) / 2.0)


def compute_migration(old_line, new_line, sr, messages=None):
    """Port of GetMigration (apex-line bend-translation adjustment omitted --
    see module docstring). Returns a dict of per-vertex arrays keyed the same
    way as the legacy tool's Mig/Mig1-4/m/m_old arrays, indexed from new_line's
    vertices 1..N-1."""
    mid_center = find_mid_centerline(old_line, new_line, sr)
    mid_old = find_mid_centerline(old_line, mid_center, sr)
    mid_new = find_mid_centerline(mid_center, new_line, sr)

    new_pts = [pt for part in new_line for pt in part if pt is not None]

    mig, mig1, mig2, mig3, mig4 = [], [], [], [], []
    m_new, m_old = [], []

    for i in range(1, len(new_pts)):
        start_pt = new_pts[i]

        p_new_mid, _, leg1, _ = mid_new.queryPointAndDistance(start_pt, False)
        p_center, _, leg2, _ = mid_center.queryPointAndDistance(p_new_mid, False)
        p_old_mid, _, leg3, _ = mid_old.queryPointAndDistance(p_center, False)
        p_old, _, leg4, _ = old_line.queryPointAndDistance(p_old_mid, False)

        total = leg1 + leg2 + leg3 + leg4
        _, _, _, migrated_from_left = new_line.queryPointAndDistance(p_old, False)
        if not migrated_from_left:
            total, leg1, leg2, leg3, leg4 = -total, -leg1, -leg2, -leg3, -leg4

        mig.append(total)
        mig1.append(leg1)
        mig2.append(leg2)
        mig3.append(leg3)
        mig4.append(leg4)

        _, dist_new, _, _ = new_line.queryPointAndDistance(start_pt, False)
        _, dist_old, _, _ = old_line.queryPointAndDistance(p_old, False)
        m_new.append(dist_new)
        m_old.append(dist_old)

        if messages is not None and i % 25 == 0:
            messages.addMessage("Migration vertex {} of {}".format(i, len(new_pts) - 1))

    return {
        "mig": mig, "mig1": mig1, "mig2": mig2, "mig3": mig3, "mig4": mig4,
        "m": m_new, "m_old": m_old,
    }


MIGRATION_FIELDS = [
    ("Mig_dist", "DOUBLE", 10, 3, None),
    ("i", "DOUBLE", 10, 2, None),
    ("m", "DOUBLE", 10, 3, None),
    ("old_m", "DOUBLE", 10, 3, None),
    ("Mig_1", "DOUBLE", 10, 3, None),
    ("Mig_2", "DOUBLE", 10, 3, None),
    ("Mig_3", "DOUBLE", 10, 3, None),
    ("Mig_4", "DOUBLE", 10, 3, None),
]


def write_migration_polygons(out_path, new_line, migration, half_width, sr):
    """Port of CreateMultiplePolys: one rectangular polygon per segment of
    new_line, spanning `half_width` on either side, carrying the migration
    distance attributes computed for that segment's downstream vertex."""
    from planform_common import create_feature_class

    create_feature_class(out_path, "POLYGON", sr, MIGRATION_FIELDS)
    new_pts = [pt for part in new_line for pt in part if pt is not None]

    fields = ["SHAPE@"] + [f[0] for f in MIGRATION_FIELDS]
    with arcpy.da.InsertCursor(out_path, fields) as cur:
        for i in range(1, len(new_pts) - 1):
            seg_a = arcpy.Polyline(arcpy.Array([new_pts[i - 1], new_pts[i]]), sr)
            seg_b = arcpy.Polyline(arcpy.Array([new_pts[i], new_pts[i + 1]]), sr)

            pt1 = _normal_offset(seg_a, half_width)
            pt2 = _normal_offset(seg_a, -half_width)
            pt3 = _normal_offset(seg_b, -half_width)
            pt4 = _normal_offset(seg_b, half_width)

            ring = arcpy.Array([pt1, pt2, pt3, pt4, pt1])
            polygon = arcpy.Polygon(ring, sr)

            idx = i - 1
            cur.insertRow([
                polygon,
                migration["mig"][idx], i, migration["m"][idx], migration["m_old"][idx],
                migration["mig1"][idx], migration["mig2"][idx],
                migration["mig3"][idx], migration["mig4"][idx],
            ])


def _normal_offset(segment, offset):
    from planform_common import query_normal_point
    return query_normal_point(segment, 0.5, True, offset)
