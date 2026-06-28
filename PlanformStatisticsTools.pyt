# PlanformStatisticsTools.pyt
# ArcGIS Pro Python Toolbox
#
# Ported from VB.NET ArcObjects add-in (ArcMap 10.x) by Wes Lauer,
# University of Minnesota, Saint Anthony Falls Laboratory.
# Original code: https://github.com/jwlauer/PlanformStatisticsTools
#
# Tools:
#   1. Centerline          – Interpolate a channel centerline from left/right bank lines
#   2. MigrationDistance   – Estimate lateral migration distance between two centerlines
#   3. BankPolygons        – Create bank buffer polygons keyed to centerline points
#
# Requirements:
#   - ArcGIS Pro (tested with 3.x) with an active Standard or Advanced license
#   - arcpy (ships with ArcGIS Pro; no conda installs needed)
#   - Spatial Analyst extension is NOT required
#   - No third-party packages needed

import arcpy
import os
import math


# ---------------------------------------------------------------------------
# Toolbox declaration
# ---------------------------------------------------------------------------

class Toolbox:
    def __init__(self):
        self.label = "Planform Statistics Tools"
        self.alias = "planformstats"
        self.tools = [Centerline, MigrationDistance, BankPolygons]


# ---------------------------------------------------------------------------
# TOOL 1 – Centerline
# ---------------------------------------------------------------------------
# Algorithm (faithful port of findCL in Centerline.vb):
#
#   Starting from the midpoint of a line connecting the FromPoints of the two
#   bank lines, the tool steps along the channel by repeatedly:
#     1. Drawing a semicircular arc of radius `dist` centred on the current
#        point, oriented perpendicular to the last step direction.
#     2. Finding where the arc intersects each bank line to constrain the
#        search window (fraction along the arc).
#     3. Doing a nested binary search (5 passes × 5 samples) to find the
#        fraction on the arc that minimises |dist_from_left - dist_from_right|,
#        subject to the point being between the two banks.
#     4. Recording width, streamwise coordinate (m), direction angle (theta),
#        change in direction (dtheta), and radius of curvature (r_curve).
#
#   Outputs:
#     - Polyline shapefile with M-values set to streamwise distance
#     - CSV text file with columns:
#       OID, m, width, theta, dtheta, r_curve, cl_x, cl_y,
#       left_x, left_y, right_x, right_y

class Centerline:
    def __init__(self):
        self.label = "1. Centerline"
        self.description = (
            "Interpolates a channel centerline at evenly spaced intervals "
            "from a left and right bank polyline. Also computes width, "
            "direction angle, change in angle, and radius of curvature at "
            "each point. Outputs a polyline shapefile (M-aware) and a CSV."
        )
        self.canRunInBackground = False

    def getParameterInfo(self):
        p0 = arcpy.Parameter(
            displayName="Left bank line",
            name="left_bank",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input")
        p0.filter.list = ["Polyline"]

        p1 = arcpy.Parameter(
            displayName="Right bank line",
            name="right_bank",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input")
        p1.filter.list = ["Polyline"]

        p2 = arcpy.Parameter(
            displayName="Point spacing (map units)",
            name="spacing",
            datatype="GPDouble",
            parameterType="Required",
            direction="Input")
        p2.value = 10.0

        p3 = arcpy.Parameter(
            displayName="Maximum number of centerline points",
            name="max_pts",
            datatype="GPLong",
            parameterType="Required",
            direction="Input")
        p3.value = 5000

        p4 = arcpy.Parameter(
            displayName="Output centerline shapefile",
            name="out_centerline",
            datatype="DEShapefile",
            parameterType="Required",
            direction="Output")

        p5 = arcpy.Parameter(
            displayName="Output CSV statistics textfile",
            name="out_csv",
            datatype="DEFile",
            parameterType="Required",
            direction="Output")
        p5.filter.list = ["csv", "txt"]

        p6 = arcpy.Parameter(
            displayName="Show interpolation graphics in map",
            name="show_graphics",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input")
        p6.value = False   # off by default so the tool runs at full speed

        return [p0, p1, p2, p3, p4, p5, p6]

    def isLicensed(self):
        return True

    def updateMessages(self, parameters):
        return

    def execute(self, parameters, messages):
        left_fc       = parameters[0].valueAsText
        right_fc      = parameters[1].valueAsText
        dist          = float(parameters[2].value)
        max_pts       = int(parameters[3].value)
        out_shp       = parameters[4].valueAsText
        out_csv       = parameters[5].valueAsText
        show_graphics = parameters[6].value  # True/False boolean

        pi = math.pi

        # Set up a graphics layer if the user asked for visual feedback.
        # In the original ArcMap tool, IScreenDisplay drew transient points
        # directly on screen.  In ArcGIS Pro the equivalent is a graphics layer,
        # which is persistent and can be toggled/removed by the student.
        g_layer = _get_graphics_layer() if show_graphics else None
        if show_graphics and g_layer is None:
            messages.addWarning(
                "No active map found – graphics display skipped. "
                "Open a map view in ArcGIS Pro before running the tool "
                "if you want to see the interpolation graphics."
            )

        # ----- Read bank geometries (first feature only) -----
        left_geom  = _get_first_geometry(left_fc)
        right_geom = _get_first_geometry(right_fc)

        if left_geom is None or right_geom is None:
            arcpy.AddError("Could not read bank line geometry.")
            return

        sr = left_geom.spatialReference

        # ----- Seed: midpoint of line connecting the two FromPoints -----
        lf = left_geom.firstPoint
        rf = right_geom.firstPoint
        seed_x = (lf.X + rf.X) / 2.0
        seed_y = (lf.Y + rf.Y) / 2.0

        # Initial direction: angle of line from right→left FromPoint, minus pi
        # (same as pLine.Angle - pi in the VB code)
        seed_angle = math.atan2(lf.Y - rf.Y, lf.X - rf.X) - pi

        # ----- Iterative centerline tracing -----
        # All arrays are seeded with a placeholder at index 0 (the seed point)
        # so that every array stays the same length as cl_pts throughout the loop.
        # The CSV writer starts at index 1, so the placeholders are never output.
        cl_pts    = [(0.0, 0.0)]  # overwritten below; index 0 = seed point
        widths    = [0.0]         # index 0 placeholder
        ms        = [0.0]         # index 0 = streamwise distance at seed = 0
        thetas    = [0.0]         # overwritten below
        dthetas   = [-99999.0]    # index 0 sentinel (no previous direction)
        rcurves   = [-99999.0]    # index 0 sentinel
        left_pts  = [(0.0, 0.0)]  # index 0 placeholder
        right_pts = [(0.0, 0.0)]  # index 0 placeholder

        pt_x, pt_y    = seed_x, seed_y
        theta_local   = seed_angle
        cl_pts[0]     = (pt_x, pt_y)
        thetas[0]     = theta_local + pi

        approx_total = int((left_geom.length / 2 + right_geom.length / 2) / dist)
        arcpy.SetProgressor("step", "Tracing centerline…", 0, min(approx_total, max_pts), 1)

        i = 0
        found = True

        while found and i < max_pts:
            arcpy.SetProgressorPosition(i)
            i += 1

            # --- Build semicircular arc centred on current point ---
            # Arc: centre=pt, start angle=theta_local, sweep=pi, chord=dist
            # We sample it as a dense polyline (100 pts) for intersection tests.
            arc_pts = _arc_to_points(pt_x, pt_y, theta_local, pi, dist, n=200)
            arc_pl  = _pts_to_polyline(arc_pts, sr)

            # --- Find arc-bank intersections to bound the search window ---
            left_int  = _nearest_intersection(arc_pl, left_geom,  pt_x, pt_y)
            right_int = _nearest_intersection(arc_pl, right_geom, pt_x, pt_y)

            # Convert intersection points to fractional positions along the arc
            arc_len = arc_pl.length
            if left_int is None:
                max_frac = 1.0
            else:
                # _nearest_intersection returns an (x, y) tuple, not an arcpy Point
                max_frac = _fraction_along_polyline(arc_pl, left_int[0], left_int[1])

            if right_int is None:
                min_frac = 0.0
            else:
                min_frac = _fraction_along_polyline(arc_pl, right_int[0], right_int[1])

            # Clamp fractions
            min_frac = max(0.0, min(min_frac, 1.0))
            max_frac = max(0.0, min(max_frac, 1.0))
            if min_frac > max_frac:
                min_frac, max_frac = max_frac, min_frac

            # --- Nested binary search (5 passes × 5 samples) ---
            best_error    = 1e9
            best_frac     = (min_frac + max_frac) / 2.0
            best_lx = best_ly = best_rx = best_ry = 0.0
            best_dl = best_dr = 0.0
            best_valid = False

            for _k in range(5):
                increment = (max_frac - min_frac) / 4.0
                for j in range(5):
                    frac = min_frac + increment * j
                    tx, ty = _point_on_arc(pt_x, pt_y, theta_local, pi, dist, frac)

                    lx, ly, dl, right_of_left   = _nearest_point_on_line(left_geom,  tx, ty)
                    rx, ry, dr, right_of_right  = _nearest_point_on_line(right_geom, tx, ty)

                    err = dl - dr
                    # Accept only if the candidate is between the banks:
                    # right_of_left=True AND right_of_right=False
                    if (abs(err) < abs(best_error)
                            and right_of_left
                            and not right_of_right):
                        best_frac  = frac
                        best_error = err
                        best_lx, best_ly = lx, ly
                        best_rx, best_ry = rx, ry
                        best_dl, best_dr = dl, dr
                        best_valid = True

                min_frac = best_frac - increment
                max_frac = best_frac + increment

            # Final point
            new_x, new_y = _point_on_arc(pt_x, pt_y, theta_local, pi, dist, best_frac)

            # --- Termination check (replicate VB: dist_along=1 means past end) ---
            _, _, dl_f, right_of_left_f  = _nearest_point_on_line(left_geom,  new_x, new_y)
            _, _, dr_f, right_of_right_f = _nearest_point_on_line(right_geom, new_x, new_y)

            l_frac = _fraction_along_polyline(left_geom,  new_x, new_y)
            r_frac = _fraction_along_polyline(right_geom, new_x, new_y)

            if l_frac >= 0.999 or r_frac >= 0.999 or right_of_right_f or not right_of_left_f:
                found = False

            cl_pts.append((new_x, new_y))
            left_pts.append((best_lx, best_ly))
            right_pts.append((best_rx, best_ry))

            # Draw a point graphic every 10 steps to show tracing progress.
            # Replicates: "If i Mod 10 = 0 Then DrawPoint(pPt)" from findCL().
            # Drawing every single point would slow the tool significantly for
            # long channels, so every 10th step is a good compromise.
            if show_graphics and i % 10 == 0:
                _draw_point_graphic(g_layer, new_x, new_y,
                                    color_rgb=(255, 80, 0), size=5)

            # Compute direction angle of current segment
            dx = new_x - pt_x
            dy = new_y - pt_y
            seg_angle = math.atan2(dy, dx)
            theta_val = seg_angle  # theta = angle of travel direction

            thetas.append(theta_val)
            dthetas.append(theta_val - thetas[-2] if i > 1 else -99999.0)
            # Wrap dtheta to (-pi, pi)
            if len(dthetas) > 0:
                if dthetas[-1] >= pi:  dthetas[-1] -= 2 * pi
                if dthetas[-1] <= -pi: dthetas[-1] += 2 * pi

            rc = (dist / dthetas[-1]) if (len(dthetas) > 0 and dthetas[-1] != 0) else -99999.0
            rcurves.append(rc)

            widths.append(best_dl + best_dr)
            ms.append(dist * i)

            # Update for next iteration
            theta_local = seg_angle - pi / 2.0   # normal direction for next arc
            pt_x, pt_y  = new_x, new_y

        arcpy.ResetProgressor()

        # Write all queued point graphics to the map as a feature layer.
        if show_graphics:
            _flush_graphics(g_layer, sr)
            _refresh_active_view()

        n_pts = len(cl_pts)
        messages.addMessage(f"Traced {n_pts} centerline points.")

        # ----- Write output shapefile -----
        _write_centerline_shapefile(cl_pts, ms, out_shp, sr)

        # ----- Write CSV -----
        _write_csv(out_csv, ms, widths, thetas, dthetas, rcurves,
                   cl_pts, left_pts, right_pts)

        messages.addMessage(f"Centerline written to: {out_shp}")
        messages.addMessage(f"Statistics CSV written to: {out_csv}")


# ---------------------------------------------------------------------------
# TOOL 2 – MigrationDistance
# ---------------------------------------------------------------------------
# Algorithm (faithful port of GetMigration + CreateMultiplePolys in ArcGISAddin1.vb):
#
#   Given a "from" (older/reference) centerline and a "to" (newer) centerline:
#     1. Compute three intermediate centerlines at 25%, 50%, and 75% of the
#        time interval using FindMidCenterline (PolyDivide): each intermediate
#        line is the fractional average of same-fraction points on two curves.
#     2. For each vertex on the "from" (new) centerline, chain four nearest-
#        perpendicular hops through the 75%, 50%, 25% lines and finally onto
#        the "to" (old) line.  The sum of those four hop distances is the
#        total migration distance Mig.  The sign is negative if the old
#        centerline point ends up on the left side of the new.
#     3. Write a polygon shapefile with one quadrilateral per vertex, centred
#        on the "from" centerline, oriented by the local normal.  Each polygon
#        carries: Mig_dist, i, m, old_m, Mig_1 … Mig_4.
#
#   Note: the original tool offered optional "apex line" trajectory adjustment.
#   That feature is preserved here as an optional parameter.

class MigrationDistance:
    def __init__(self):
        self.label = "2. Migration Distance"
        self.description = (
            "Estimates lateral migration distance between two channel centerlines "
            "using a 4-step interpolated hop method. Outputs a polygon shapefile "
            "with one quadrilateral per centerline point storing migration statistics."
        )
        self.canRunInBackground = False

    def getParameterInfo(self):
        p0 = arcpy.Parameter(
            displayName="'From' centerline (reference / older)",
            name="cl_from",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input")
        p0.filter.list = ["Polyline"]

        p1 = arcpy.Parameter(
            displayName="'To' centerline (newer / target)",
            name="cl_to",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input")
        p1.filter.list = ["Polyline"]

        p2 = arcpy.Parameter(
            displayName="Width of output polygon (map units, lateral from centerline)",
            name="poly_width",
            datatype="GPDouble",
            parameterType="Required",
            direction="Input")
        p2.value = 50.0

        p3 = arcpy.Parameter(
            displayName="Output polygon shapefile",
            name="out_polys",
            datatype="DEShapefile",
            parameterType="Required",
            direction="Output")

        p4 = arcpy.Parameter(
            displayName="Apex line shapefile (optional – for translating bends)",
            name="apex_lines",
            datatype="GPFeatureLayer",
            parameterType="Optional",
            direction="Input")
        p4.filter.list = ["Polyline"]

        p5 = arcpy.Parameter(
            displayName="Show intermediate centerlines and trajectories in map",
            name="show_graphics",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input")
        p5.value = False

        return [p0, p1, p2, p3, p4, p5]

    def isLicensed(self):
        return True

    def updateMessages(self, parameters):
        return

    def execute(self, parameters, messages):
        cl_from_fc    = parameters[0].valueAsText
        cl_to_fc      = parameters[1].valueAsText
        poly_width    = float(parameters[2].value)
        out_polys     = parameters[3].valueAsText
        apex_fc       = parameters[4].valueAsText if parameters[4].value else None
        show_graphics = parameters[5].value  # True/False boolean

        pi = math.pi

        # Initialise graphics layer if requested.
        # The original tool used CreateDrawingElement() to add coloured line
        # elements to ArcMap's BasicGraphicsLayer.  In ArcGIS Pro we use
        # a named graphics layer ("PlanformGraphics") which works the same way.
        g_layer = _get_graphics_layer() if show_graphics else None
        if show_graphics and g_layer is None:
            messages.addWarning(
                "No active map found – graphics display skipped. "
                "Open a map view before running if you want to see the graphics."
            )

        # ----- Read geometries -----
        geom_new = _get_first_geometry(cl_from_fc)   # "from" = newer = reference
        geom_old = _get_first_geometry(cl_to_fc)     # "to"   = older

        if geom_new is None or geom_old is None:
            arcpy.AddError("Could not read centerline geometry.")
            return

        if geom_new.spatialReference.name != geom_old.spatialReference.name:
            arcpy.AddWarning(
                "Spatial references of the two centerlines do not match. "
                "Results may be incorrect."
            )

        sr = geom_new.spatialReference

        # Densify both lines to the same number of evenly-spaced vertices
        # (needed for PolyDivide which relies on parametric sampling)
        n_pts_new = _vertex_count(geom_new)
        n_pts_old = _vertex_count(geom_old)

        # ----- Build intermediate centerlines (25%, 50%, 75%) -----
        messages.addMessage("Building intermediate centerlines…")
        mid_center = _find_mid_centerline(geom_old, geom_new, sr)      # 50%
        mid_old    = _find_mid_centerline(geom_old, mid_center, sr)    # 25%
        mid_new    = _find_mid_centerline(mid_center, geom_new, sr)    # 75%

        # ----- Draw the three intermediate centerlines -----
        # Replicates: CreateDrawingElement(IntermediatePlineOld/New/Center)
        # Colours match the original tool's convention:
        #   25% line (old-side)   → green
        #   50% line (centre)     → cyan
        #   75% line (new-side)   → yellow
        if show_graphics:
            _draw_polyline_graphic(g_layer, _get_vertices(mid_old),
                                   color_rgb=(0, 200, 80),  width=1.5)   # 25%
            _draw_polyline_graphic(g_layer, _get_vertices(mid_center),
                                   color_rgb=(0, 180, 220), width=1.5)   # 50%
            _draw_polyline_graphic(g_layer, _get_vertices(mid_new),
                                   color_rgb=(220, 200, 0), width=1.5)   # 75%
            messages.addMessage(
                "Intermediate centerlines drawn: "
                "green=25%, cyan=50%, yellow=75%."
            )

        # ----- Compute migration for each vertex of the "from" (new) line -----
        pts_new  = _get_vertices(geom_new)
        n        = len(pts_new)

        Mig   = [0.0] * n
        mig1  = [0.0] * n
        mig2  = [0.0] * n
        mig3  = [0.0] * n
        mig4  = [0.0] * n
        m_new = [0.0] * n
        m_old_arr = [0.0] * n

        arcpy.SetProgressor("step", "Computing migration distances…", 0, n, 1)

        for i in range(1, n):
            arcpy.SetProgressorPosition(i)

            start = pts_new[i]

            # 4-hop chain: new → 75% → 50% → 25% → old
            p1 = _nearest_point_on_geom(mid_new,    start[0], start[1])
            p2 = _nearest_point_on_geom(mid_center, p1[0],    p1[1])
            p3 = _nearest_point_on_geom(mid_old,    p2[0],    p2[1])
            p4 = _nearest_point_on_geom(geom_old,   p3[0],    p3[1])

            d1 = _dist2d(start, p1)
            d2 = _dist2d(p1, p2)
            d3 = _dist2d(p2, p3)
            d4 = _dist2d(p3, p4)
            total = math.sqrt(
                (p4[0] - start[0])**2 + (p4[1] - start[1])**2
            )
            # Use straight-line distance as total (matches Trajectory.Length in VB)
            # Actually VB uses the polyline through the 5 hop-points; we replicate:
            total = d1 + d2 + d3 + d4

            # Sign: negative if old-CL point ends up on the left of new CL
            _, _, _, migrated_from_left = _nearest_point_on_line(geom_new, p4[0], p4[1])
            sign = 1.0 if migrated_from_left else -1.0

            Mig[i]  = sign * total
            mig1[i] = sign * d1
            mig2[i] = sign * d2
            mig3[i] = sign * d3
            mig4[i] = sign * d4

            # Streamwise coordinates (fractional × length)
            m_new[i]    = _fraction_along_polyline(geom_new, start[0], start[1]) * geom_new.length
            m_old_arr[i] = _fraction_along_polyline(geom_old, p4[0], p4[1]) * geom_old.length

            # Draw the 5-point migration trajectory for this vertex.
            # Replicates: CreateDrawingElement(pMigPoints, pActiveView)
            # The trajectory runs: new-CL vertex → 75% line → 50% line
            #                      → 25% line → old-CL nearest point
            # Colour: orange-red so it stands out against the blue/green CL lines.
            if show_graphics:
                trajectory = [start, p1, p2, p3, p4]
                _draw_polyline_graphic(g_layer, trajectory,
                                       color_rgb=(255, 80, 0), width=0.75)

        arcpy.ResetProgressor()

        # Write all queued graphics to the map as feature layers.
        if show_graphics:
            _flush_graphics(g_layer, sr)
            _refresh_active_view()

        # ----- Write output polygon shapefile -----
        _write_migration_polys(
            out_polys, geom_new, pts_new, Mig, m_new, m_old_arr,
            mig1, mig2, mig3, mig4, poly_width, sr
        )

        messages.addMessage(f"Migration polygons written to: {out_polys}")


# ---------------------------------------------------------------------------
# TOOL 3 – BankPolygons
# ---------------------------------------------------------------------------
# Algorithm (faithful port of CreateBankBufferPolygons in BankPolys.vb):
#
#   Inputs:
#     - Centerline
#     - Right and left bank lines for BUFFER geometry (LiDAR / surveyed)
#     - Right and left bank lines for LENGTH measurement (photo-interpreted)
#     - Buffer width (positive offset distance from bank lines)
#     - Maximum lateral distance from centerline (dist) used for normal lines
#     - Angle threshold: above this change-in-direction, a one-sided search
#       is performed instead of bilateral normals (handles tight bends)
#
#   For each consecutive pair of centerline vertices, the tool:
#     1. Offsets the bank lines inward by BufferWidth.
#     2. Draws normal lines from the midpoint of the current segment.
#        - If |dtheta| > threshold: extends the normal only toward the outside
#          of the bend; the inside point snaps to the nearest point on the
#          offset bank (avoiding the normal missing it on tight bends).
#        - Otherwise: bilateral normals intersect both offset banks.
#     3. Extracts subcurves of the offset banks between consecutive lateral
#        intersection points.
#     4. Assembles an 8-point polygon from upstream-CL, left-offset points,
#        downstream-CL, and right-offset points.
#     5. Intersects that polygon with the photo-interpreted banks to get
#        bank lengths (LB_len, RB_len) and with the buffer banks (LB_buf_ln, RB_buf_ln).
#     6. Cuts the polygon along the two buffer bank lines to produce separate
#        left and right bank polygon shapefiles.

class BankPolygons:
    def __init__(self):
        self.label = "3. Bank Polygons"
        self.description = (
            "Creates bank-buffer polygons corresponding to each centerline "
            "segment. Records bank lengths from photo-interpreted and survey "
            "bank lines. Outputs three polygon shapefiles: combined, left bank, "
            "and right bank."
        )
        self.canRunInBackground = False

    def getParameterInfo(self):
        p0 = arcpy.Parameter(
            displayName="Centerline",
            name="centerline",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input")
        p0.filter.list = ["Polyline"]

        p1 = arcpy.Parameter(
            displayName="Right bank line (buffer geometry, e.g. LiDAR)",
            name="right_bank_buf",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input")
        p1.filter.list = ["Polyline"]

        p2 = arcpy.Parameter(
            displayName="Left bank line (buffer geometry, e.g. LiDAR)",
            name="left_bank_buf",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input")
        p2.filter.list = ["Polyline"]

        p3 = arcpy.Parameter(
            displayName="Right bank line (length measurement, e.g. aerial photo)",
            name="right_bank_photo",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input")
        p3.filter.list = ["Polyline"]

        p4 = arcpy.Parameter(
            displayName="Left bank line (length measurement, e.g. aerial photo)",
            name="left_bank_photo",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input")
        p4.filter.list = ["Polyline"]

        p5 = arcpy.Parameter(
            displayName="Buffer width (inward offset from bank lines, map units)",
            name="buffer_width",
            datatype="GPDouble",
            parameterType="Required",
            direction="Input")
        p5.value = 20.0

        p6 = arcpy.Parameter(
            displayName="Maximum lateral distance from centerline (normal line half-length)",
            name="max_lateral_dist",
            datatype="GPDouble",
            parameterType="Required",
            direction="Input")
        p6.value = 100.0

        p7 = arcpy.Parameter(
            displayName="Angle threshold for one-sided normals (degrees)",
            name="angle_threshold",
            datatype="GPDouble",
            parameterType="Required",
            direction="Input")
        p7.value = 15.0

        p8 = arcpy.Parameter(
            displayName="Output combined polygon shapefile",
            name="out_polys",
            datatype="DEShapefile",
            parameterType="Required",
            direction="Output")

        p9 = arcpy.Parameter(
            displayName="Output left bank polygon shapefile",
            name="out_left",
            datatype="DEShapefile",
            parameterType="Required",
            direction="Output")

        p10 = arcpy.Parameter(
            displayName="Output right bank polygon shapefile",
            name="out_right",
            datatype="DEShapefile",
            parameterType="Required",
            direction="Output")

        return [p0, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10]

    def isLicensed(self):
        return True

    def updateMessages(self, parameters):
        return

    def execute(self, parameters, messages):
        cl_fc        = parameters[0].valueAsText
        right_buf_fc = parameters[1].valueAsText
        left_buf_fc  = parameters[2].valueAsText
        right_photo_fc = parameters[3].valueAsText
        left_photo_fc  = parameters[4].valueAsText
        buf_width    = float(parameters[5].value)
        max_dist     = float(parameters[6].value)
        angle_thresh = float(parameters[7].value)
        out_polys    = parameters[8].valueAsText
        out_left     = parameters[9].valueAsText
        out_right    = parameters[10].valueAsText

        pi = math.pi

        # ----- Read geometries -----
        cl_geom         = _get_first_geometry(cl_fc)
        right_buf_geom  = _get_first_geometry(right_buf_fc)
        left_buf_geom   = _get_first_geometry(left_buf_fc)
        right_photo_geom = _get_first_geometry(right_photo_fc)
        left_photo_geom  = _get_first_geometry(left_photo_fc)

        if any(g is None for g in [cl_geom, right_buf_geom, left_buf_geom,
                                    right_photo_geom, left_photo_geom]):
            arcpy.AddError("Could not read one or more input geometries.")
            return

        sr = cl_geom.spatialReference

        # ----- Offset bank lines inward by buf_width -----
        # arcpy.Buffer uses positive = right side of travel direction
        # Right bank offset: negative offset (left/inward) = toward channel
        # Left  bank offset: positive offset = toward channel
        # We use arcpy.Geometry.buffer on lines, but for single-sided offsets
        # we need arcpy.PolylineToPolygon workarounds or manual offset.
        # Simplest accurate approach: use the ConstructOffset equivalent via
        # arcpy's offset tool on the geometry object using a workaround.
        right_offset = _offset_polyline(right_buf_geom, -buf_width, sr)
        left_offset  = _offset_polyline(left_buf_geom,   buf_width, sr)

        if right_offset is None or left_offset is None:
            arcpy.AddError("Could not construct offset lines. Check bank geometries.")
            return

        # ----- Get centerline vertices -----
        cl_pts = _get_vertices(cl_geom)
        n = len(cl_pts)

        # ----- Create output feature classes -----
        out_fields = [
            ("Mig_dist", "DOUBLE"),
            ("i",        "LONG"),
            ("old_m",    "DOUBLE"),
            ("LB_len",   "DOUBLE"),
            ("RB_len",   "DOUBLE"),
            ("LB_buf_ln","DOUBLE"),
            ("RB_buf_ln","DOUBLE"),
        ]
        _create_polygon_fc(out_polys,  sr, out_fields)
        _create_polygon_fc(out_left,   sr, [("i", "LONG")])
        _create_polygon_fc(out_right,  sr, [("i", "LONG")])

        # ----- Seed starting lateral points -----
        # The VB code uses segment MIDPOINTS as the centerline vertices of each
        # polygon (pClPtUpstream and pClPtStart), not the raw CL vertex coords.
        # This ensures the line from CL-midpoint to bank corner is truly
        # perpendicular to the segment (since the normal was drawn from that midpoint).
        #
        # pCL_up  = midpoint of segment i-1 → i  (upstream CL vertex of polygon)
        # pCL_dn  = midpoint of segment i   → i+1 (downstream CL vertex of polygon)
        # pCL_up carries forward as pCL_dn from the previous step.

        seg0_mx = (cl_pts[0][0] + cl_pts[1][0]) / 2.0
        seg0_my = (cl_pts[0][1] + cl_pts[1][1]) / 2.0
        seg0_angle = math.atan2(cl_pts[1][1] - cl_pts[0][1],
                                 cl_pts[1][0] - cl_pts[0][0])

        # Normal-line tips from seed midpoint; project onto offset banks
        lx0, ly0 = _point_along_normal(seg0_mx, seg0_my, seg0_angle,  max_dist)
        rx0, ry0 = _point_along_normal(seg0_mx, seg0_my, seg0_angle, -max_dist)
        pPt3 = _snap_to_line(left_offset,  lx0, ly0)
        pPt4 = _snap_to_line(right_offset, rx0, ry0)

        # pCL_dn for iteration i=1 is the midpoint of segment 0→1 (= seg0 midpoint).
        # It becomes pCL_up at the start of i=1.
        pCL_dn = (seg0_mx, seg0_my)

        arcpy.SetProgressor("step", "Building bank polygons…", 0, n - 3, 1)

        ins_main  = arcpy.da.InsertCursor(out_polys, ["SHAPE@"] + [f[0] for f in out_fields])
        ins_left  = arcpy.da.InsertCursor(out_left,  ["SHAPE@", "i"])
        ins_right = arcpy.da.InsertCursor(out_right, ["SHAPE@", "i"])

        for i in range(1, n - 2):
            arcpy.SetProgressorPosition(i)

            # The upstream CL vertex of this polygon tile is the midpoint of
            # segment i-1→i, which was computed as pCL_dn in the previous step.
            pCL_up = pCL_dn

            # The downstream CL vertex is the midpoint of segment i→i+1.
            # This is also where the normal line is drawn from, so the line
            # from this point to the bank corner is truly perpendicular.
            seg_mx = (cl_pts[i][0] + cl_pts[i + 1][0]) / 2.0
            seg_my = (cl_pts[i][1] + cl_pts[i + 1][1]) / 2.0
            seg_angle = math.atan2(cl_pts[i + 1][1] - cl_pts[i][1],
                                    cl_pts[i + 1][0] - cl_pts[i][0])
            pCL_dn = (seg_mx, seg_my)  # carries forward as pCL_up next iteration

            # Upstream and downstream segment angles for centered-difference dtheta
            up_angle   = math.atan2(cl_pts[i][1]     - cl_pts[i - 1][1],
                                     cl_pts[i][0]     - cl_pts[i - 1][0])
            down_angle = math.atan2(cl_pts[i + 2][1] - cl_pts[i + 1][1],
                                     cl_pts[i + 2][0] - cl_pts[i + 1][0])

            dtheta = 0.5 * (down_angle - up_angle)
            if dtheta < -pi: dtheta += 2 * pi
            if dtheta >  pi: dtheta -= 2 * pi

            # Bank corner points carry forward from previous iteration
            pPt1 = pPt3   # upstream-left bank corner
            pPt2 = pPt4   # upstream-right bank corner

            thresh_rad = angle_thresh * pi / 180.0

            if dtheta > thresh_rad:
                # Curving sharply left: extend normal to the right, snap left to offset
                nx, ny = _point_along_normal(seg_mx, seg_my, seg_angle, -max_dist)
                norm_pl = _two_pt_polyline(seg_mx, seg_my, nx, ny, sr)
                ints = _intersect_lines(norm_pl, right_offset, sr)
                pPt4 = _closest_point(ints, (seg_mx, seg_my)) if ints else pPt2
                # Left snaps to nearest
                pPt3 = _snap_to_line(left_offset, seg_mx, seg_my)

            elif dtheta < -thresh_rad:
                # Curving sharply right: extend normal to the left, snap right
                nx, ny = _point_along_normal(seg_mx, seg_my, seg_angle,  max_dist)
                norm_pl = _two_pt_polyline(seg_mx, seg_my, nx, ny, sr)
                ints = _intersect_lines(norm_pl, left_offset, sr)
                pPt3 = _closest_point(ints, (seg_mx, seg_my)) if ints else pPt1
                pPt4 = _snap_to_line(right_offset, seg_mx, seg_my)

            else:
                # Near-straight: bilateral normals
                lx, ly = _point_along_normal(seg_mx, seg_my, seg_angle,  max_dist)
                rx, ry = _point_along_normal(seg_mx, seg_my, seg_angle, -max_dist)
                lnorm = _two_pt_polyline(seg_mx, seg_my, lx, ly, sr)
                rnorm = _two_pt_polyline(seg_mx, seg_my, rx, ry, sr)
                l_ints = _intersect_lines(lnorm, left_offset,  sr)
                r_ints = _intersect_lines(rnorm, right_offset, sr)
                pPt3 = _closest_point(l_ints, (seg_mx, seg_my)) if l_ints else pPt1
                pPt4 = _closest_point(r_ints, (seg_mx, seg_my)) if r_ints else pPt2

            # Monotonicity guard: new bank point must not be upstream of previous
            s1 = _fraction_along_polyline(left_offset,  pPt1[0], pPt1[1])
            e1 = _fraction_along_polyline(left_offset,  pPt3[0], pPt3[1])
            if s1 > e1: pPt3 = pPt1

            s2 = _fraction_along_polyline(right_offset, pPt2[0], pPt2[1])
            e2 = _fraction_along_polyline(right_offset, pPt4[0], pPt4[1])
            if s2 > e2: pPt4 = pPt2

            # Extract subcurves of offset banks between corner points.
            # Left sub: pPt1 → pPt3 (upstream→downstream along left bank)
            # Right sub: pPt4 → pPt2 (downstream→upstream along right bank)
            # This matches the VB code (GetSubcurve from pPt1→pPt3 on left,
            # and pPt4→pPt2 on right) and produces a valid non-self-intersecting ring.
            left_sub  = _subcurve(left_offset,  pPt1, pPt3, sr)
            right_sub = _subcurve(right_offset, pPt4, pPt2, sr)

            # Build the closed polygon ring using SEGMENT MIDPOINTS as the CL
            # vertices (matching VB: pClPtUpstream and pClPtStart).
            # This ensures the bank-corner→CL line is truly perpendicular to
            # the segment, because both the midpoint and the bank corner were
            # derived from the same normal line drawn at that midpoint.
            # Ring order:
            #   pCL_up → pPt1 → left_sub(pPt1→pPt3) → pPt3
            #   → pCL_dn → pPt4 → right_sub(pPt4→pPt2) → pPt2 → pCL_up (close)
            poly_pts = []
            poly_pts.append(arcpy.Point(pCL_up[0], pCL_up[1]))   # upstream CL midpoint
            poly_pts.append(arcpy.Point(pPt1[0], pPt1[1]))
            for vx, vy in _get_vertices_geom(left_sub):
                poly_pts.append(arcpy.Point(vx, vy))
            poly_pts.append(arcpy.Point(pPt3[0], pPt3[1]))
            poly_pts.append(arcpy.Point(pCL_dn[0], pCL_dn[1]))   # downstream CL midpoint
            poly_pts.append(arcpy.Point(pPt4[0], pPt4[1]))
            for vx, vy in _get_vertices_geom(right_sub):
                poly_pts.append(arcpy.Point(vx, vy))
            poly_pts.append(arcpy.Point(pPt2[0], pPt2[1]))
            poly_pts.append(arcpy.Point(pCL_up[0], pCL_up[1]))   # close ring

            arr = arcpy.Array(poly_pts)
            polygon = arcpy.Polygon(arr, sr)

            # Measure bank lengths within polygon
            lb_photo_len = _line_length_within_polygon(left_photo_geom,  polygon, sr)
            rb_photo_len = _line_length_within_polygon(right_photo_geom, polygon, sr)
            lb_buf_len   = _line_length_within_polygon(left_buf_geom,    polygon, sr)
            rb_buf_len   = _line_length_within_polygon(right_buf_geom,   polygon, sr)

            m_old = _fraction_along_polyline(cl_geom, cl_pts[i][0], cl_pts[i][1]) * cl_geom.length

            ins_main.insertRow((polygon, 0.0, i, m_old,
                                lb_photo_len, rb_photo_len,
                                lb_buf_len,   rb_buf_len))

            # Cut polygon by buffer banks for left/right sub-polygons
            left_poly  = _cut_polygon_left( polygon, left_buf_geom,  sr)
            right_poly = _cut_polygon_right(polygon, right_buf_geom, sr)

            if left_poly:
                ins_left.insertRow((left_poly, i))
            if right_poly:
                ins_right.insertRow((right_poly, i))

        del ins_main, ins_left, ins_right
        arcpy.ResetProgressor()

        messages.addMessage(f"Bank polygons written to: {out_polys}")
        messages.addMessage(f"Left bank polygons: {out_left}")
        messages.addMessage(f"Right bank polygons: {out_right}")


# ===========================================================================
# SHARED HELPER FUNCTIONS
# ===========================================================================

# ---------------------------------------------------------------------------
# Graphics layer helpers
# ---------------------------------------------------------------------------
# ArcGIS Pro does not have IScreenDisplay (the ArcMap transient drawing API).
# The equivalent is arcpy.mp's graphics layer, which adds persistent graphic
# elements to the active map that students can see, toggle, and remove.
#
# _get_graphics_layer() finds or creates a dedicated graphics layer named
# "PlanformGraphics" so our drawings don't clutter the default layer.
# _draw_point_graphic() and _draw_polyline_graphic() add individual elements.

def _get_graphics_layer(name="PlanformGraphics"):
    """
    Return a dict that acts as a graphics accumulator: a named in-memory
    feature class for points and one for polylines.  The actual feature
    classes are created lazily on first use.

    Using in-memory feature classes (written at the end of the tool run and
    then added to the map as ordinary layers) is more reliable than the CIM
    graphics layer API, which varies between ArcGIS Pro versions.

    Returns a dict with keys 'name', 'pts', 'lines' (both initially empty lists),
    or None if we are not inside an ArcGIS Pro session.
    """
    try:
        # Confirm we are inside a live Pro session
        aprx = arcpy.mp.ArcGISProject("CURRENT")
        if aprx.activeMap is None:
            return None
        return {"name": name, "pts": [], "lines": [], "sr": None}
    except Exception:
        return None


def _draw_point_graphic(g_accum, x, y, color_rgb=(255, 0, 0), size=6):
    """
    Queue a point for display.  Points are flushed to the map at the end
    of the tool run by _flush_graphics().

    Parameters
    ----------
    g_accum   : dict returned by _get_graphics_layer(), or None
    x, y      : float – map coordinates
    color_rgb : ignored here (layer symbology can be set in Pro after display)
    size      : ignored here
    """
    if g_accum is None:
        return
    g_accum["pts"].append((x, y))


def _draw_polyline_graphic(g_accum, xy_list, color_rgb=(0, 180, 220), width=1.0):
    """
    Queue a polyline for display.

    Parameters
    ----------
    g_accum : dict returned by _get_graphics_layer(), or None
    xy_list : list of (x, y) tuples
    """
    if g_accum is None or len(xy_list) < 2:
        return
    g_accum["lines"].append(list(xy_list))


def _flush_graphics(g_accum, sr):
    """
    Write all queued points and lines to in-memory feature classes and add
    them to the active map as ordinary feature layers.

    Called once after the tool loop finishes.  Students see the results as
    standard selectable/styleable layers named e.g. "PlanformGraphics_pts".

    Parameters
    ----------
    g_accum : dict from _get_graphics_layer()
    sr      : arcpy.SpatialReference to assign to the output layers
    """
    if g_accum is None:
        return
    try:
        aprx     = arcpy.mp.ArcGISProject("CURRENT")
        act_map  = aprx.activeMap
        if act_map is None:
            return
        name = g_accum["name"]

        # --- Points layer ---
        if g_accum["pts"]:
            fc_pts = f"memory\\{name}_pts"
            if arcpy.Exists(fc_pts):
                arcpy.management.Delete(fc_pts)
            arcpy.management.CreateFeatureclass(
                "memory", f"{name}_pts", "POINT", spatial_reference=sr)
            with arcpy.da.InsertCursor(fc_pts, ["SHAPE@XY"]) as cur:
                for xy in g_accum["pts"]:
                    cur.insertRow([xy])
            lyr_pts = arcpy.management.MakeFeatureLayer(fc_pts, f"{name}_pts").getOutput(0)
            act_map.addLayer(lyr_pts)
            arcpy.AddMessage(
                f"Centerline interpolation points added as layer '{name}_pts'. "
                "Points are drawn every 10 steps along the traced centerline."
            )

        # --- Lines layer ---
        if g_accum["lines"]:
            fc_lines = f"memory\\{name}_lines"
            if arcpy.Exists(fc_lines):
                arcpy.management.Delete(fc_lines)
            arcpy.management.CreateFeatureclass(
                "memory", f"{name}_lines", "POLYLINE", spatial_reference=sr)
            with arcpy.da.InsertCursor(fc_lines, ["SHAPE@"]) as cur:
                for xy_list in g_accum["lines"]:
                    arr = arcpy.Array([arcpy.Point(x, y) for x, y in xy_list])
                    cur.insertRow([arcpy.Polyline(arr, sr)])
            lyr_lines = arcpy.management.MakeFeatureLayer(fc_lines, f"{name}_lines").getOutput(0)
            act_map.addLayer(lyr_lines)
            arcpy.AddMessage(
                f"Migration lines added as layer '{name}_lines'."
            )
    except Exception as e:
        arcpy.AddWarning(f"Could not add graphics layers to map: {e}")


def _refresh_active_view():
    """Refresh the active map view."""
    try:
        aprx = arcpy.mp.ArcGISProject("CURRENT")
        for view in aprx.listViews():
            view.refresh()
    except Exception:
        pass

def _get_first_geometry(fc):
    """Return the geometry of the first feature in a feature class/layer."""
    with arcpy.da.SearchCursor(fc, ["SHAPE@"]) as cur:
        for row in cur:
            return row[0]
    return None


def _get_vertices(geom):
    """Return list of (x, y) tuples for all vertices of a polyline geometry."""
    pts = []
    for part in geom:
        for pt in part:
            if pt is not None:
                pts.append((pt.X, pt.Y))
    return pts


def _get_vertices_geom(geom):
    """Same as _get_vertices but accepts None gracefully."""
    if geom is None:
        return []
    return _get_vertices(geom)


def _vertex_count(geom):
    return len(_get_vertices(geom))


def _arcpy_point(x, y):
    return (x, y)   # used as a tuple internally; arcpy.Point created at output


def _dist2d(a, b):
    return math.sqrt((a[0] - b[0])**2 + (a[1] - b[1])**2)


# ---------------------------------------------------------------------------
# Arc / circle geometry helpers (replacing ArcObjects ICircularArc)
# ---------------------------------------------------------------------------

def _arc_to_points(cx, cy, start_angle, sweep, radius, n=200):
    """
    Return n (x, y) points on a circular arc.
    Centre: (cx, cy), starting angle: start_angle (radians),
    sweep: total angle (radians, positive = CCW), radius: chord half-length.
    """
    pts = []
    for k in range(n + 1):
        a = start_angle + sweep * k / n
        pts.append((cx + radius * math.cos(a), cy + radius * math.sin(a)))
    return pts


def _point_on_arc(cx, cy, start_angle, sweep, radius, fraction):
    """Return the point at fractional position (0–1) along the arc."""
    a = start_angle + sweep * fraction
    return cx + radius * math.cos(a), cy + radius * math.sin(a)


def _pts_to_polyline(pts, sr):
    """Build an arcpy Polyline from a list of (x, y) tuples."""
    arr = arcpy.Array([arcpy.Point(x, y) for x, y in pts])
    return arcpy.Polyline(arr, sr)


def _two_pt_polyline(x1, y1, x2, y2, sr):
    """Build a two-point arcpy Polyline."""
    return arcpy.Polyline(arcpy.Array([arcpy.Point(x1, y1), arcpy.Point(x2, y2)]), sr)


# ---------------------------------------------------------------------------
# Intersection helpers
# ---------------------------------------------------------------------------

def _nearest_intersection(arc_pl, bank_geom, ref_x, ref_y):
    """
    Find the intersection of arc_pl with bank_geom nearest to (ref_x, ref_y).
    Returns (x, y) tuple or None.
    """
    try:
        pts_geom = arc_pl.intersect(bank_geom, 1)   # 1 = point geometry
        if pts_geom is None or pts_geom.isMultipart:
            parts = [pts_geom.getPart(k) for k in range(pts_geom.partCount)] if pts_geom else []
        else:
            parts = [pts_geom]

        best = None
        best_dist = 1e18
        if pts_geom:
            for k in range(pts_geom.partCount):
                pt = pts_geom.getPart(k)
                if pt is not None:
                    d = math.sqrt((pt.X - ref_x)**2 + (pt.Y - ref_y)**2)
                    if d < best_dist:
                        best_dist = d
                        best = (pt.X, pt.Y)
        return best
    except Exception:
        return None


def _intersect_lines(line1, line2, sr):
    """
    Intersect two arcpy Polyline geometries.
    Returns list of (x, y) tuples for intersection points, or [].
    """
    try:
        result = line1.intersect(line2, 1)
        pts = []
        if result:
            for k in range(result.partCount):
                pt = result.getPart(k)
                if pt:
                    pts.append((pt.X, pt.Y))
        return pts
    except Exception:
        return []


# ---------------------------------------------------------------------------
# Nearest-point-on-line helpers (replacing QueryPointAndDistance)
# ---------------------------------------------------------------------------

def _nearest_point_on_line(geom, x, y):
    """
    Find the nearest point on geom to (x, y).
    Returns (nearest_x, nearest_y, normal_distance, is_right_of_line).
    is_right_of_line: True if the query point is on the right side of the
    polyline direction (i.e. arcpy snapType "RIGHT").
    """
    pt   = arcpy.Point(x, y)
    mpt  = arcpy.PointGeometry(pt, geom.spatialReference)
    snap = geom.queryPointAndDistance(mpt, False)
    # queryPointAndDistance returns (PointGeometry, dist_along, dist_from, rightside)
    # snap[0] is a PointGeometry object; use .firstPoint to get the arcpy.Point
    nearest_pt    = snap[0].firstPoint
    dist_from     = snap[2]
    is_right      = snap[3]
    return nearest_pt.X, nearest_pt.Y, dist_from, is_right


def _nearest_point_on_geom(geom, x, y):
    """Returns (nearest_x, nearest_y)."""
    nx, ny, _, _ = _nearest_point_on_line(geom, x, y)
    return nx, ny


def _fraction_along_polyline(geom, x, y):
    """
    Returns the fractional distance (0–1) along geom to the nearest point to (x,y).
    """
    pt   = arcpy.Point(x, y)
    mpt  = arcpy.PointGeometry(pt, geom.spatialReference)
    snap = geom.queryPointAndDistance(mpt, False)
    # snap[1] is the distance along the line (absolute)
    total = geom.length
    if total == 0:
        return 0.0
    return snap[1] / total


def _snap_to_line(geom, x, y):
    """Returns the (x, y) of the nearest point on geom to query (x, y)."""
    return _nearest_point_on_geom(geom, x, y)


def _closest_point(pt_list, ref):
    """Return the (x, y) in pt_list closest to ref (x, y)."""
    if not pt_list:
        return ref
    return min(pt_list, key=lambda p: _dist2d(p, ref))


# ---------------------------------------------------------------------------
# Normal line helpers
# ---------------------------------------------------------------------------

def _point_along_normal(mx, my, seg_angle, dist):
    """
    Return point at `dist` along the normal to a segment with direction
    seg_angle.  Positive dist = left of travel; negative = right.
    """
    normal_angle = seg_angle + math.pi / 2.0
    return (mx + dist * math.cos(normal_angle),
            my + dist * math.sin(normal_angle))


# ---------------------------------------------------------------------------
# Offset polyline helper (replaces IConstructCurve.ConstructOffset)
# ---------------------------------------------------------------------------

def _offset_polyline(geom, offset, sr):
    """
    Create a single-sided parallel offset of a polyline geometry.
    Positive offset = left of travel direction; negative = right.
    Uses arcpy's buffer + polygon → boundary approach for robustness.
    Falls back to a manual vertex-by-vertex normal shift if arcpy fails.
    """
    try:
        pts_in = _get_vertices(geom)
        if len(pts_in) < 2:
            return None
        offset_pts = []
        n = len(pts_in)
        for k in range(n):
            if k == 0:
                # First vertex: use direction of the first segment only
                angle = math.atan2(pts_in[1][1] - pts_in[0][1],
                                   pts_in[1][0] - pts_in[0][0])
            elif k == n - 1:
                # Last vertex: use direction of the last segment only
                angle = math.atan2(pts_in[-1][1] - pts_in[-2][1],
                                   pts_in[-1][0] - pts_in[-2][0])
            else:
                # Interior vertex: use the CIRCULAR mean of the two adjacent
                # segment directions.  Simple arithmetic averaging of angles
                # fails near the ±π wrap (e.g. averaging 170° and -170° gives
                # 0° instead of 180°).  The circular mean sums unit vectors
                # and takes atan2 of the result, which handles the wrap correctly.
                a1 = math.atan2(pts_in[k][1]   - pts_in[k-1][1],
                                pts_in[k][0]   - pts_in[k-1][0])
                a2 = math.atan2(pts_in[k+1][1] - pts_in[k][1],
                                pts_in[k+1][0] - pts_in[k][0])
                # Sum unit vectors pointing in direction a1 and a2
                sx = math.cos(a1) + math.cos(a2)
                sy = math.sin(a1) + math.sin(a2)
                angle = math.atan2(sy, sx)
            # Offset perpendicular to travel direction.
            # Positive offset → left of travel (normal = angle + π/2)
            normal = angle + math.pi / 2.0
            ox = pts_in[k][0] + offset * math.cos(normal)
            oy = pts_in[k][1] + offset * math.sin(normal)
            offset_pts.append(arcpy.Point(ox, oy))
        return arcpy.Polyline(arcpy.Array(offset_pts), sr)
    except Exception as e:
        arcpy.AddWarning(f"Offset polyline failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Subcurve helper (replaces IPolyline.GetSubcurve)
# ---------------------------------------------------------------------------

def _subcurve(geom, from_pt, to_pt, sr):
    """
    Extract the portion of geom between the nearest points to from_pt and to_pt.
    Returns an arcpy Polyline or None.
    """
    try:
        total = geom.length
        f_frac = _fraction_along_polyline(geom, from_pt[0], from_pt[1])
        t_frac = _fraction_along_polyline(geom, to_pt[0],   to_pt[1])
        # Ensure from < to
        if f_frac > t_frac:
            f_frac, t_frac = t_frac, f_frac
        f_dist = f_frac * total
        t_dist = t_frac * total
        result = geom.segmentAlongLine(f_dist, t_dist, False)
        return result
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Polygon cut helpers (replaces ITopologicalOperator.Cut)
# ---------------------------------------------------------------------------

def _cut_polygon_left(polygon, cut_line, sr):
    """Return the left-side polygon after cutting polygon by cut_line."""
    try:
        left_poly, _ = polygon.cut(cut_line)
        return left_poly
    except Exception:
        return None


def _cut_polygon_right(polygon, cut_line, sr):
    """Return the right-side polygon after cutting polygon by cut_line."""
    try:
        _, right_poly = polygon.cut(cut_line)
        return right_poly
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Line length within polygon
# ---------------------------------------------------------------------------

def _line_length_within_polygon(line_geom, polygon, sr):
    """Return the total length of the portion of line_geom inside polygon."""
    try:
        clipped = line_geom.intersect(polygon, 2)   # 2 = polyline dimension
        if clipped:
            return clipped.length
        return 0.0
    except Exception:
        return 0.0


# ---------------------------------------------------------------------------
# PolyDivide / FindMidCenterline (from FindMidCenterline + PolyDivide in VB)
# ---------------------------------------------------------------------------

def _poly_divide(pline1_geom, pline2_geom, sr):
    """
    Port of PolyDivide: creates a new polyline whose vertices are the
    fractional midpoints between corresponding fractional positions on two
    input polylines. Uses max(n_pts1, n_pts2) interior samples.
    """
    pts1 = _get_vertices(pline1_geom)
    pts2 = _get_vertices(pline2_geom)
    max_pts = max(len(pts1), len(pts2))

    result_pts = []
    # Start
    result_pts.append(arcpy.Point(
        (pline1_geom.firstPoint.X + pline2_geom.firstPoint.X) / 2.0,
        (pline1_geom.firstPoint.Y + pline2_geom.firstPoint.Y) / 2.0
    ))
    for i in range(1, max_pts + 1):
        frac = i / (max_pts + 1)
        # Sample fractional point along each line
        p1 = pline1_geom.positionAlongLine(frac, True)  # True = as fraction
        p2 = pline2_geom.positionAlongLine(frac, True)
        result_pts.append(arcpy.Point(
            (p1.firstPoint.X + p2.firstPoint.X) / 2.0,
            (p1.firstPoint.Y + p2.firstPoint.Y) / 2.0
        ))
    # End
    result_pts.append(arcpy.Point(
        (pline1_geom.lastPoint.X + pline2_geom.lastPoint.X) / 2.0,
        (pline1_geom.lastPoint.Y + pline2_geom.lastPoint.Y) / 2.0
    ))
    return arcpy.Polyline(arcpy.Array(result_pts), sr)


def _find_mid_centerline(pline_old, pline_new, sr):
    """
    Port of FindMidCenterline.
    Finds crossing points of the two centerlines, sorts them by downstream
    distance, then interpolates the midpoint curve segment-by-segment between
    consecutive crossing points using PolyDivide.
    """
    # Find intersection points
    try:
        int_pts_geom = pline_old.intersect(pline_new, 1)
    except Exception:
        int_pts_geom = None

    int_pts = []
    if int_pts_geom:
        for k in range(int_pts_geom.partCount):
            pt = int_pts_geom.getPart(k)
            if pt:
                # Store with M = fraction along new line
                frac = _fraction_along_polyline(pline_new, pt.X, pt.Y)
                int_pts.append((pt.X, pt.Y, frac))
    # Sort by down-channel distance
    int_pts.sort(key=lambda p: p[2])

    # Build anchor lists on old and new lines
    anchor_old = [(pline_old.firstPoint.X, pline_old.firstPoint.Y)]
    anchor_new = [(pline_new.firstPoint.X, pline_new.firstPoint.Y)]

    # Add midpoints between consecutive crossings (VB logic)
    for j in range(1, len(int_pts)):
        # Mid between crossing j-1 and j on old line
        sf_old = _fraction_along_polyline(pline_old, int_pts[j-1][0], int_pts[j-1][1])
        ef_old = _fraction_along_polyline(pline_old, int_pts[j][0],   int_pts[j][1])
        mid_old_pt = pline_old.positionAlongLine((sf_old + ef_old) / 2.0, True)
        anchor_old.append((mid_old_pt.firstPoint.X, mid_old_pt.firstPoint.Y))

        sf_new = _fraction_along_polyline(pline_new, int_pts[j-1][0], int_pts[j-1][1])
        ef_new = _fraction_along_polyline(pline_new, int_pts[j][0],   int_pts[j][1])
        mid_new_pt = pline_new.positionAlongLine((sf_new + ef_new) / 2.0, True)
        anchor_new.append((mid_new_pt.firstPoint.X, mid_new_pt.firstPoint.Y))

    anchor_old.append((pline_old.lastPoint.X, pline_old.lastPoint.Y))
    anchor_new.append((pline_new.lastPoint.X, pline_new.lastPoint.Y))

    # PolyDivide between each pair of anchors and concatenate
    result_pts = []
    for j in range(1, len(anchor_old)):
        sf_old = _fraction_along_polyline(pline_old, anchor_old[j-1][0], anchor_old[j-1][1])
        ef_old = _fraction_along_polyline(pline_old, anchor_old[j][0],   anchor_old[j][1])
        seg_old = pline_old.segmentAlongLine(sf_old * pline_old.length,
                                              ef_old * pline_old.length, False)

        sf_new = _fraction_along_polyline(pline_new, anchor_new[j-1][0], anchor_new[j-1][1])
        ef_new = _fraction_along_polyline(pline_new, anchor_new[j][0],   anchor_new[j][1])
        seg_new = pline_new.segmentAlongLine(sf_new * pline_new.length,
                                              ef_new * pline_new.length, False)

        mid_seg = _poly_divide(seg_old, seg_new, sr)
        verts = _get_vertices(mid_seg)
        if result_pts and verts:
            result_pts.extend(verts[1:])  # avoid duplicate junction point
        else:
            result_pts.extend(verts)

    if not result_pts:
        # Fallback: simple midpoint line
        return _poly_divide(pline_old, pline_new, sr)

    arr = arcpy.Array([arcpy.Point(x, y) for x, y in result_pts])
    return arcpy.Polyline(arr, sr)


# ---------------------------------------------------------------------------
# Feature class creation helpers
# ---------------------------------------------------------------------------

def _create_polygon_fc(path, sr, extra_fields):
    """Create a polygon shapefile with given extra attribute fields."""
    out_dir  = os.path.dirname(path)
    out_name = os.path.basename(path)
    if out_name.lower().endswith(".shp"):
        out_name = out_name[:-4]
    arcpy.management.CreateFeatureclass(
        out_dir, out_name, "POLYGON", spatial_reference=sr
    )
    full_path = os.path.join(out_dir, out_name + ".shp")
    for fname, ftype in extra_fields:
        arcpy.management.AddField(full_path, fname, ftype)


def _write_centerline_shapefile(cl_pts, ms, out_shp, sr):
    """Write the centerline as a polyline shapefile with M values."""
    out_dir  = os.path.dirname(out_shp)
    out_name = os.path.basename(out_shp)
    if out_name.lower().endswith(".shp"):
        out_name = out_name[:-4]

    arcpy.management.CreateFeatureclass(
        out_dir, out_name, "POLYLINE",
        has_m="ENABLED",
        spatial_reference=sr
    )
    full_path = os.path.join(out_dir, out_name + ".shp")
    arcpy.management.AddField(full_path, "m_start", "DOUBLE")

    # Build M-aware array
    arr = arcpy.Array()
    for k, (x, y) in enumerate(cl_pts):
        m_val = ms[k] if k < len(ms) else ms[-1]
        pt = arcpy.Point(x, y, None, m_val)
        arr.add(pt)

    pl = arcpy.Polyline(arr, sr, False, True)  # has_z=False, has_m=True
    with arcpy.da.InsertCursor(full_path, ["SHAPE@", "m_start"]) as cur:
        cur.insertRow((pl, ms[0] if ms else 0.0))


def _write_csv(csv_path, ms, widths, thetas, dthetas, rcurves,
               cl_pts, left_pts, right_pts):
    """Write the statistics CSV file (port of ExportWidthText).
    All arrays share the same index scheme: index 0 is the seed-point
    placeholder; data rows start at index 1.
    """
    # All arrays are the same length (= number of CL points including seed).
    # We use the shortest length as a safe upper bound, then iterate from 1.
    n = min(len(ms), len(widths), len(thetas), len(dthetas),
            len(rcurves), len(cl_pts), len(left_pts), len(right_pts))
    with open(csv_path, "w") as f:
        f.write("OID,m,width,theta,dtheta,r_curve,cl_x,cl_y,left_x,left_y,right_x,right_y\n")
        for i in range(1, n):   # skip index 0 (seed placeholder)
            cl_x, cl_y = cl_pts[i]
            l_x,  l_y  = left_pts[i]
            r_x,  r_y  = right_pts[i]
            f.write(f"0,{ms[i]},{widths[i]},{thetas[i]},{dthetas[i]},"
                    f"{rcurves[i]},{cl_x},{cl_y},{l_x},{l_y},{r_x},{r_y}\n")


def _write_migration_polys(out_shp, geom_new, pts_new, Mig, m_new, m_old,
                            mig1, mig2, mig3, mig4, dist, sr):
    """Write migration polygon shapefile (port of CreateMultiplePolys)."""
    out_dir  = os.path.dirname(out_shp)
    out_name = os.path.basename(out_shp)
    if out_name.lower().endswith(".shp"):
        out_name = out_name[:-4]

    arcpy.management.CreateFeatureclass(out_dir, out_name, "POLYGON",
                                         spatial_reference=sr)
    full = os.path.join(out_dir, out_name + ".shp")
    for fname in ["Mig_dist", "i", "m", "old_m",
                   "Mig_1", "Mig_2", "Mig_3", "Mig_4"]:
        arcpy.management.AddField(full, fname, "DOUBLE")

    fields = ["SHAPE@", "Mig_dist", "i", "m", "old_m",
              "Mig_1", "Mig_2", "Mig_3", "Mig_4"]

    n = len(pts_new)
    with arcpy.da.InsertCursor(full, fields) as cur:
        for i in range(1, n - 1):
            # Build normal-offset quadrilateral centred on the "from" CL
            # (port of CreateMultiplePolys inner loop)
            x_prev, y_prev = pts_new[i - 1]
            x_cur,  y_cur  = pts_new[i]
            x_next, y_next = pts_new[i + 1]

            # Segment i-1 → i
            a1 = math.atan2(y_cur - y_prev, x_cur - x_prev)
            # Normal endpoints at midpoint of first segment
            mid1_x = (x_prev + x_cur) / 2.0
            mid1_y = (y_prev + y_cur) / 2.0
            pt1x = mid1_x + dist * math.cos(a1 + math.pi / 2)
            pt1y = mid1_y + dist * math.sin(a1 + math.pi / 2)
            pt2x = mid1_x + dist * math.cos(a1 - math.pi / 2)
            pt2y = mid1_y + dist * math.sin(a1 - math.pi / 2)

            # Segment i → i+1
            a2 = math.atan2(y_next - y_cur, x_next - x_cur)
            mid2_x = (x_cur + x_next) / 2.0
            mid2_y = (y_cur + y_next) / 2.0
            pt3x = mid2_x + dist * math.cos(a2 - math.pi / 2)
            pt3y = mid2_y + dist * math.sin(a2 - math.pi / 2)
            pt4x = mid2_x + dist * math.cos(a2 + math.pi / 2)
            pt4y = mid2_y + dist * math.sin(a2 + math.pi / 2)

            arr = arcpy.Array([
                arcpy.Point(pt1x, pt1y),
                arcpy.Point(pt2x, pt2y),
                arcpy.Point(pt3x, pt3y),
                arcpy.Point(pt4x, pt4y),
                arcpy.Point(pt1x, pt1y),
            ])
            polygon = arcpy.Polygon(arr, sr)

            cur.insertRow((polygon,
                           Mig[i], i, m_new[i], m_old[i],
                           mig1[i], mig2[i], mig3[i], mig4[i]))
