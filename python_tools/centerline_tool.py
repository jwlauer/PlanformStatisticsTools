"""
Port of Centerline.vb: interpolates a centerline at evenly spaced intervals
between two roughly parallel bank lines, recording width, direction, and
curvature at each point.
"""
import csv
import math
import os

import arcpy

from planform_common import (
    PI,
    arc_point_at_fraction,
    circular_arc_points,
    make_polyline,
    nearest_intersection_point,
)


def find_centerline(left, right, spacing, max_points, sr, messages=None):
    """Port of program_utilities.findCL.

    Returns a dict with parallel lists describing the centerline:
    points, width, m, theta, dtheta, r_curve, left_points, right_points.
    """
    p0 = ((right.firstPoint.X + left.firstPoint.X) / 2.0,
          (right.firstPoint.Y + left.firstPoint.Y) / 2.0)
    theta_local = math.atan2(left.firstPoint.Y - right.firstPoint.Y,
                              left.firstPoint.X - right.firstPoint.X) - PI

    points = [p0]
    thetas = [theta_local]
    widths, ms, dthetas, rcurves = [], [], [], []

    i = 0
    found = True
    while found and i < max_points:
        i += 1
        cx, cy = points[-1]

        arc_pts = circular_arc_points((cx, cy), theta_local, PI, spacing)
        arc_line = make_polyline(arc_pts, sr)

        left_int = nearest_intersection_point(arc_line, left)
        right_int = nearest_intersection_point(arc_line, right)

        if left_int is None:
            max_fraction = 1.0
        else:
            _, max_fraction, _, _ = left.queryPointAndDistance(left_int, True)

        if right_int is None:
            min_fraction = 0.0
        else:
            _, min_fraction, _, _ = right.queryPointAndDistance(right_int, True)

        best_fraction = 0.0
        best_error = math.inf
        increment = (max_fraction - min_fraction) / 4.0
        for _ in range(5):
            increment = (max_fraction - min_fraction) / 4.0
            for j in range(5):
                fraction = min_fraction + increment * j
                pt = arc_point_at_fraction((cx, cy), theta_local, PI, spacing, fraction)
                _, _, dist_from_left, right_of_left = left.queryPointAndDistance(pt, False)
                _, _, dist_from_right, right_of_right = right.queryPointAndDistance(pt, False)
                dist_error = dist_from_left - dist_from_right
                if abs(dist_error) < abs(best_error) and right_of_left and not right_of_right:
                    best_fraction = fraction
                    best_error = dist_error
            min_fraction = best_fraction - increment
            max_fraction = best_fraction + increment

        pt = arc_point_at_fraction((cx, cy), theta_local, PI, spacing, best_fraction)
        _, frac_left, dist_from_left, right_of_left = left.queryPointAndDistance(pt, True)
        _, frac_right, dist_from_right, right_of_right = right.queryPointAndDistance(pt, True)

        found = not (frac_left >= 1.0 or frac_right >= 1.0 or right_of_right or not right_of_left)

        points.append((pt.X, pt.Y))
        dx = points[-1][0] - points[-2][0]
        dy = points[-1][1] - points[-2][1]
        theta_local = math.atan2(dy, dx) - PI / 2.0

        widths.append(dist_from_left + dist_from_right)
        ms.append(spacing * i)
        theta_val = theta_local + PI / 2.0
        if theta_val < 0:
            theta_val += 2 * PI
        thetas.append(theta_val)

        dtheta = thetas[-1] - thetas[-2]
        if dtheta >= PI:
            dtheta -= 2 * PI
        if dtheta <= -PI:
            dtheta += 2 * PI
        dthetas.append(dtheta)
        rcurves.append(spacing / dtheta if dtheta != 0 else None)

        if messages is not None and i % 10 == 0:
            messages.addMessage("Centerline point {}".format(i))

    left_points, right_points = [], []
    for x, y in points:
        pt = arcpy.Point(x, y)
        left_pt = left.queryPointAndDistance(pt, False)[0]
        right_pt = right.queryPointAndDistance(pt, False)[0]
        left_points.append((left_pt.X, left_pt.Y))
        right_points.append((right_pt.X, right_pt.Y))

    return {
        "points": points,
        "width": widths,
        "m": ms,
        "theta": thetas[1:],
        "dtheta": dthetas,
        "r_curve": rcurves,
        "left_points": left_points,
        "right_points": right_points,
    }


def write_centerline_feature_class(out_path, result, sr):
    """PolylineM feature class with M set to downstream distance (m)."""
    arcpy.management.CreateFeatureclass(
        os.path.dirname(out_path), os.path.basename(out_path),
        geometry_type="POLYLINE", spatial_reference=sr, has_m="ENABLED")
    pts = result["points"]
    m_values = [0.0] + result["m"]
    arr = arcpy.Array([arcpy.Point(x, y, None, m) for (x, y), m in zip(pts, m_values)])
    polyline = arcpy.Polyline(arr, sr, False, True)
    with arcpy.da.InsertCursor(out_path, ["SHAPE@"]) as cur:
        cur.insertRow([polyline])


def write_centerline_table(out_csv, result):
    """CSV with one row per interior point: m, width, theta, dtheta, r_curve,
    plus the matching point on the centerline and on each bank."""
    with open(out_csv, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["OID", "m", "width", "theta", "dtheta", "r_curve",
                          "cl_x", "cl_y", "left_x", "left_y", "right_x", "right_y"])
        pts = result["points"][1:]
        left_pts = result["left_points"][1:]
        right_pts = result["right_points"][1:]
        for idx, (m, width, theta, dtheta, r_curve, (cx, cy), (lx, ly), (rx, ry)) in enumerate(
                zip(result["m"], result["width"], result["theta"], result["dtheta"],
                    result["r_curve"], pts, left_pts, right_pts), start=1):
            writer.writerow([0, m, width, theta, dtheta,
                              r_curve if r_curve is not None else -99999,
                              cx, cy, lx, ly, rx, ry])
