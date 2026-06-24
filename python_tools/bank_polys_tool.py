"""
Port of BankPolys.vb: builds a polygon for each centerline segment bounded by
buffered (offset) versions of the bank lines, then measures how much of each
"photo" bank line and each "lidar" bank line falls inside that polygon.
"""
import math

import arcpy

from planform_common import (
    PI,
    closest_point,
    construct_offset_polyline,
    create_feature_class,
    intersection_points,
    query_normal_point,
)

MAIN_FIELDS = [
    ("i", "DOUBLE", 10, 2, None),
    ("LB_len", "DOUBLE", 10, 3, None),
    ("RB_len", "DOUBLE", 10, 3, None),
    ("LB_buf_ln", "DOUBLE", 10, 3, None),
    ("RB_buf_ln", "DOUBLE", 10, 3, None),
]
SIDE_FIELDS = [("i", "DOUBLE", 10, 2, None)]


def create_bank_buffer_polygons(
        centerline, left_lidar, right_lidar, left_photo, right_photo,
        buffer_width, max_bank_distance, angle_threshold_deg, sr,
        out_main, out_left, out_right, messages=None):
    """Port of CreateBankBufferPolygons."""
    right_offset = construct_offset_polyline(right_lidar, buffer_width, sr)
    left_offset = construct_offset_polyline(left_lidar, -buffer_width, sr)

    create_feature_class(out_main, "POLYGON", sr, MAIN_FIELDS)
    create_feature_class(out_left, "POLYGON", sr, SIDE_FIELDS)
    create_feature_class(out_right, "POLYGON", sr, SIDE_FIELDS)

    cl_pts = [pt for part in centerline for pt in part if pt is not None]
    angle_threshold = angle_threshold_deg * PI / 180.0

    seg0 = arcpy.Polyline(arcpy.Array([cl_pts[0], cl_pts[1]]), sr)
    cl_start = seg0.positionAlongLine(0.5, True).firstPoint
    pt3 = _first_point(intersection_points(
        arcpy.Polyline(arcpy.Array([cl_start, query_normal_point(seg0, 0.5, True, -max_bank_distance)]), sr),
        left_offset))
    pt4 = _first_point(intersection_points(
        arcpy.Polyline(arcpy.Array([cl_start, query_normal_point(seg0, 0.5, True, max_bank_distance)]), sr),
        right_offset))

    main_fields = ["SHAPE@"] + [f[0] for f in MAIN_FIELDS]
    side_fields = ["SHAPE@"] + [f[0] for f in SIDE_FIELDS]

    with arcpy.da.InsertCursor(out_main, main_fields) as main_cur, \
            arcpy.da.InsertCursor(out_left, side_fields) as left_cur, \
            arcpy.da.InsertCursor(out_right, side_fields) as right_cur:

        for i in range(1, len(cl_pts) - 2):
            seg_up = arcpy.Polyline(arcpy.Array([cl_pts[i - 1], cl_pts[i]]), sr)
            seg_mid = arcpy.Polyline(arcpy.Array([cl_pts[i], cl_pts[i + 1]]), sr)
            seg_down = arcpy.Polyline(arcpy.Array([cl_pts[i + 1], cl_pts[i + 2]]), sr)

            theta_upstream = _angle(seg_up)
            theta_downstream = _angle(seg_down)
            cl_pt_upstream = cl_start
            cl_start = seg_mid.positionAlongLine(0.5, True).firstPoint

            dtheta = 0.5 * (theta_downstream - theta_upstream)
            if dtheta < -PI:
                dtheta += 2 * PI
            if dtheta > PI:
                dtheta -= 2 * PI

            pt1, pt2 = pt3, pt4
            pt3, pt4 = _find_bank_points(
                seg_mid, cl_start, dtheta, angle_threshold,
                left_offset, right_offset, max_bank_distance, pt1, pt2)

            pt4 = _enforce_forward_progress(right_offset, pt2, pt4)
            pt3 = _enforce_forward_progress(left_offset, pt1, pt3)

            left_sub = _subcurve_between(left_offset, pt1, pt3)
            right_sub = _subcurve_between(right_offset, pt4, pt2)

            ring_pts = ([cl_pt_upstream, pt1] + list(left_sub) +
                        [pt3, cl_start, pt4] + list(right_sub) +
                        [pt2, cl_pt_upstream])
            polygon = arcpy.Polygon(arcpy.Array(ring_pts), sr)

            lb_len = _bank_length_in_polygon(left_photo, polygon)
            rb_len = _bank_length_in_polygon(right_photo, polygon)
            lb_buf_len = _bank_length_in_polygon(left_lidar, polygon)
            rb_buf_len = _bank_length_in_polygon(right_lidar, polygon)

            main_cur.insertRow([polygon, i, lb_len, rb_len, lb_buf_len, rb_buf_len])

            left_part, _ = polygon.cut(left_lidar)
            left_cur.insertRow([left_part, i])

            _, right_part = polygon.cut(right_lidar)
            right_cur.insertRow([right_part, i])

            if messages is not None and i % 25 == 0:
                messages.addMessage("Bank polygon {} of {}".format(i, len(cl_pts) - 3))


def _find_bank_points(seg_mid, cl_start, dtheta, angle_threshold,
                       left_offset, right_offset, max_dist, pt1, pt2):
    """Locate the left/right offset-bank points bounding this centerline
    segment, extending the normal line further on the inside of a sharp bend
    so the polygon doesn't pinch off. Port of the three-way branch in
    CreateBankBufferPolygons."""
    if dtheta > angle_threshold:
        right_line = arcpy.Polyline(arcpy.Array(
            [cl_start, query_normal_point(seg_mid, 0.5, True, max_dist)]), seg_mid.spatialReference)
        candidates = intersection_points(right_line, right_offset)
        pt4 = closest_point(cl_start, candidates) if candidates else pt2
        pt3 = left_offset.queryPointAndDistance(cl_start, False)[0]
    elif dtheta < -angle_threshold:
        left_line = arcpy.Polyline(arcpy.Array(
            [cl_start, query_normal_point(seg_mid, 0.5, True, -max_dist)]), seg_mid.spatialReference)
        candidates = intersection_points(left_line, left_offset)
        pt3 = closest_point(cl_start, candidates) if candidates else pt1
        pt4 = right_offset.queryPointAndDistance(cl_start, False)[0]
    else:
        right_line = arcpy.Polyline(arcpy.Array(
            [cl_start, query_normal_point(seg_mid, 0.5, True, max_dist)]), seg_mid.spatialReference)
        candidates = intersection_points(right_line, right_offset)
        pt4 = closest_point(cl_start, candidates) if candidates else pt2

        left_line = arcpy.Polyline(arcpy.Array(
            [cl_start, query_normal_point(seg_mid, 0.5, True, -max_dist)]), seg_mid.spatialReference)
        candidates = intersection_points(left_line, left_offset)
        pt3 = closest_point(cl_start, candidates) if candidates else pt1
    return pt3, pt4


def _enforce_forward_progress(offset_line, prev_pt, new_pt):
    """Don't let the new bank point fall upstream of the previous one."""
    start_dist = offset_line.queryPointAndDistance(prev_pt, False)[1]
    end_dist = offset_line.queryPointAndDistance(new_pt, False)[1]
    return prev_pt if start_dist > end_dist else new_pt


def _subcurve_between(offset_line, pt_a, pt_b):
    start = offset_line.queryPointAndDistance(pt_a, True)[1]
    end = offset_line.queryPointAndDistance(pt_b, True)[1]
    sub = offset_line.segmentAlongLine(start, end, True)
    return [pt for part in sub for pt in part if pt is not None]


def _bank_length_in_polygon(bank_line, polygon):
    clipped = bank_line.intersect(polygon, 2)
    return clipped.length


def _angle(segment):
    p0, p1 = segment.firstPoint, segment.lastPoint
    return math.atan2(p1.Y - p0.Y, p1.X - p0.X)


def _first_point(points):
    if not points:
        raise ValueError("Expected an intersection point but found none; "
                          "check that buffer_width and max_bank_distance are "
                          "large enough to reach the offset bank lines.")
    return points[0]
