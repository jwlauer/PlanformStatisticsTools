# -*- coding: utf-8 -*-
"""
ArcGIS Pro Python toolbox replacing the legacy ArcMap VB.NET add-ins
(Centerline.vb, ArcGISAddin1.vb, BankPolys.vb). Each tool below takes its
inputs as ordinary geoprocessing parameters instead of mouse-driven map
selections and dialog boxes, so the tools can be run interactively in
ArcGIS Pro or scripted/batched headlessly.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))

import arcpy

import bank_polys_tool
import centerline_tool
import migration_tool
from planform_common import first_geometry


class Toolbox(object):
    def __init__(self):
        self.label = "Planform Statistics Tools"
        self.alias = "planform"
        self.tools = [CenterlineTool, MigrationTool, BankPolysTool]


class CenterlineTool(object):
    def __init__(self):
        self.label = "Interpolate Centerline"
        self.description = (
            "Interpolates a centerline at evenly spaced intervals between two "
            "roughly parallel bank lines, recording width, direction, and "
            "curvature at each point.")

    def getParameterInfo(self):
        left_bank = arcpy.Parameter(
            displayName="Left bank line", name="left_bank",
            datatype="GPFeatureLayer", parameterType="Required", direction="Input")
        left_bank.filter.list = ["Polyline"]

        right_bank = arcpy.Parameter(
            displayName="Right bank line", name="right_bank",
            datatype="GPFeatureLayer", parameterType="Required", direction="Input")
        right_bank.filter.list = ["Polyline"]

        spacing = arcpy.Parameter(
            displayName="Point spacing", name="spacing",
            datatype="GPDouble", parameterType="Required", direction="Input")

        max_points = arcpy.Parameter(
            displayName="Maximum number of points", name="max_points",
            datatype="GPLong", parameterType="Required", direction="Input")
        max_points.value = 10000

        out_centerline = arcpy.Parameter(
            displayName="Output centerline feature class", name="out_centerline",
            datatype="DEFeatureClass", parameterType="Required", direction="Output")

        out_csv = arcpy.Parameter(
            displayName="Output statistics table (CSV)", name="out_csv",
            datatype="DEFile", parameterType="Required", direction="Output")
        out_csv.filter.list = ["csv"]

        return [left_bank, right_bank, spacing, max_points, out_centerline, out_csv]

    def isLicensed(self):
        return True

    def execute(self, parameters, messages):
        left_fc = parameters[0].valueAsText
        right_fc = parameters[1].valueAsText
        spacing = parameters[2].value
        max_points = parameters[3].value
        out_centerline = parameters[4].valueAsText
        out_csv = parameters[5].valueAsText

        left = first_geometry(left_fc)
        right = first_geometry(right_fc)
        sr = left.spatialReference

        result = centerline_tool.find_centerline(
            left, right, spacing, int(max_points), sr, messages)
        centerline_tool.write_centerline_feature_class(out_centerline, result, sr)
        centerline_tool.write_centerline_table(out_csv, result)
        messages.addMessage(
            "Centerline written with {} points.".format(len(result["points"])))


class MigrationTool(object):
    def __init__(self):
        self.label = "Compute Bank Migration"
        self.description = (
            "Measures lateral migration distance between an older and a newer "
            "channel centerline and writes a polygon per centerline segment "
            "carrying the migration distance and intermediate-leg distances.")

    def getParameterInfo(self):
        old_cl = arcpy.Parameter(
            displayName="Older centerline", name="old_centerline",
            datatype="GPFeatureLayer", parameterType="Required", direction="Input")
        old_cl.filter.list = ["Polyline"]

        new_cl = arcpy.Parameter(
            displayName="Newer centerline", name="new_centerline",
            datatype="GPFeatureLayer", parameterType="Required", direction="Input")
        new_cl.filter.list = ["Polyline"]

        half_width = arcpy.Parameter(
            displayName="Polygon half-width (lateral distance from centerline)",
            name="half_width", datatype="GPDouble",
            parameterType="Required", direction="Input")

        out_polygons = arcpy.Parameter(
            displayName="Output migration polygons", name="out_polygons",
            datatype="DEFeatureClass", parameterType="Required", direction="Output")

        return [old_cl, new_cl, half_width, out_polygons]

    def isLicensed(self):
        return True

    def execute(self, parameters, messages):
        old_fc = parameters[0].valueAsText
        new_fc = parameters[1].valueAsText
        half_width = parameters[2].value
        out_polygons = parameters[3].valueAsText

        old_line = first_geometry(old_fc)
        new_line = first_geometry(new_fc)
        sr = old_line.spatialReference

        if old_line.spatialReference.name != new_line.spatialReference.name:
            messages.addErrorMessage(
                "Spatial reference of new and old centerlines do not match.")
            raise arcpy.ExecuteError

        migration = migration_tool.compute_migration(old_line, new_line, sr, messages)
        migration_tool.write_migration_polygons(
            out_polygons, new_line, migration, half_width, sr)
        messages.addMessage(
            "Migration polygons written for {} segments.".format(len(migration["mig"])))


class BankPolysTool(object):
    def __init__(self):
        self.label = "Create Bank Polygons"
        self.description = (
            "Builds a polygon for each centerline segment bounded by buffered "
            "bank lines, and measures how much of a photo-derived and a "
            "lidar-derived bank line falls inside each polygon.")

    def getParameterInfo(self):
        def line_param(name, label):
            p = arcpy.Parameter(
                displayName=label, name=name, datatype="GPFeatureLayer",
                parameterType="Required", direction="Input")
            p.filter.list = ["Polyline"]
            return p

        centerline = line_param("centerline", "Centerline")
        right_lidar = line_param("right_lidar", "Right bank line (lidar, used for buffers)")
        left_lidar = line_param("left_lidar", "Left bank line (lidar, used for buffers)")
        right_photo = line_param("right_photo", "Right bank line (photo, for length measurement)")
        left_photo = line_param("left_photo", "Left bank line (photo, for length measurement)")

        buffer_width = arcpy.Parameter(
            displayName="Buffer width", name="buffer_width",
            datatype="GPDouble", parameterType="Required", direction="Input")

        max_bank_distance = arcpy.Parameter(
            displayName="Maximum distance from channel to bank buffer",
            name="max_bank_distance", datatype="GPDouble",
            parameterType="Required", direction="Input")

        angle_threshold = arcpy.Parameter(
            displayName="Curvature angle threshold (degrees)",
            name="angle_threshold", datatype="GPDouble",
            parameterType="Required", direction="Input")
        angle_threshold.value = 5.0

        out_main = arcpy.Parameter(
            displayName="Output bank polygons", name="out_main",
            datatype="DEFeatureClass", parameterType="Required", direction="Output")
        out_left = arcpy.Parameter(
            displayName="Output left-bank polygons", name="out_left",
            datatype="DEFeatureClass", parameterType="Required", direction="Output")
        out_right = arcpy.Parameter(
            displayName="Output right-bank polygons", name="out_right",
            datatype="DEFeatureClass", parameterType="Required", direction="Output")

        return [centerline, right_lidar, left_lidar, right_photo, left_photo,
                buffer_width, max_bank_distance, angle_threshold,
                out_main, out_left, out_right]

    def isLicensed(self):
        return True

    def execute(self, parameters, messages):
        centerline_fc = parameters[0].valueAsText
        right_lidar_fc = parameters[1].valueAsText
        left_lidar_fc = parameters[2].valueAsText
        right_photo_fc = parameters[3].valueAsText
        left_photo_fc = parameters[4].valueAsText
        buffer_width = parameters[5].value
        max_bank_distance = parameters[6].value
        angle_threshold = parameters[7].value
        out_main = parameters[8].valueAsText
        out_left = parameters[9].valueAsText
        out_right = parameters[10].valueAsText

        centerline = first_geometry(centerline_fc)
        right_lidar = first_geometry(right_lidar_fc)
        left_lidar = first_geometry(left_lidar_fc)
        right_photo = first_geometry(right_photo_fc)
        left_photo = first_geometry(left_photo_fc)
        sr = centerline.spatialReference

        bank_polys_tool.create_bank_buffer_polygons(
            centerline, left_lidar, right_lidar, left_photo, right_photo,
            buffer_width, max_bank_distance, angle_threshold, sr,
            out_main, out_left, out_right, messages)
        messages.addMessage("Bank polygons written to {}.".format(out_main))
