'Written by Wes Lauer
'University of Minnesota
'Saint Anthony Falls Laboratory
'2 3rd Avenue, SE, Minneapolis, MN 55414

'April 6, 2004
'Updated August 1, 2011

'This tool creates a set of line segments based on a single line that record the average lateral
'normal distance between the end points of the segments and a second line.

Option Explicit On

Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Framework
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.DataSourcesFile

Public Class Migration
    Inherits ESRI.ArcGIS.Desktop.AddIns.Tool

    'Private m_pGxDialog As IGxDialog
    Dim DefaultSpatialReference As ISpatialReference3
    Dim steps As Integer
    Dim pMxApp As IMxApplication
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pActiveView As IActiveView
    Dim pEnvelope As IEnvelope
    Dim pEnumFeat As IEnumFeature
    Dim pGeom As IGeometry
    Dim pGeom_old As IGeometry
    Dim pGeom_new As IGeometry
    Private m_pGxDialog As IGxDialog
    Private m_pGxObjectFilter As IGxObjectFilter

    Protected Overrides Sub OnActivate()
        MyBase.OnActivate()
        'Initialize all public variables upon tool selection
        steps = 0
        MsgBox("Select the 'to' centerline (to which distances will be measured)")

    End Sub

    Protected Overrides Sub OnMouseDown(ByVal arg As ESRI.ArcGIS.Desktop.AddIns.Tool.MouseEventArgs)
        Dim pMxApp As IMxApplication
        Dim pMxDoc As IMxDocument
        Dim pMap As IMap
        Dim pActiveView As IActiveView
        Dim pEnvelope As IEnvelope
        Dim pEnumFeat As IEnumFeature
        Dim pGc As IGeometryCollection

        '        Dim i As Long
        '       Dim R0 As Double
        '       Dim threshold_distance As Double
        Dim ShallIContinue As String = Nothing

        Dim ConsiderApexLines As String
        Dim pApexLineInFeatureclass As IFeatureClass
        Dim lApexLineTotalFeatureCount As Long
        '        Dim lApexLineEmptyFeatureCount As Long

        '      Dim xy() As Double
        Dim Mig1() As Double = Nothing ' distance from older centerline to first interpolated centerline
        Dim Mig2() As Double = Nothing ' distance from first interpolated centerline to second interpolated centerline
        Dim Mig3() As Double = Nothing ' distance from second interpolated centerline to third interpolated centerline
        Dim Mig4() As Double = Nothing ' distance from third interpolated centerline to newer centerline
        Dim Mig() As Double = Nothing  ' total offset (mig1 + mig2 + mig3 + mig4)
        Dim m_old() As Double = Nothing ' down channel coordinate of older centerline
        Dim m() As Double = Nothing ' down channel coordinate of newer centerline

        pMxApp = My.ArcMap.Application
        pMxDoc = My.ArcMap.Application.Document
        pMap = pMxDoc.FocusMap
        pActiveView = pMap

        Select Case steps
            Case 0
                'GET THE COORDINATES FOR THE OLDER CENTERLINE
                pEnvelope = pMxDoc.CurrentLocation.Envelope
                pEnvelope.Expand(pMxDoc.SearchTolerance, pMxDoc.SearchTolerance, False)

                'Refresh the old selection to erase it
                pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)
                'Perform the selection using a point created on mouse down
                pMap.SelectByShape(pEnvelope, pMxApp.SelectionEnvironment, True)
                'Refresh again to draw the new selection
                pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)
                pMxDoc = My.ArcMap.Document
                pEnumFeat = pMxDoc.FocusMap.FeatureSelection
                If pMxDoc.FocusMap.SelectionCount <> 1 Then
                    MsgBox("Nothing selected.  Stopping tool.")
                    My.ArcMap.Application.CurrentTool = Nothing
                    Return
                End If
                pEnumFeat.Reset()

                pGeom_old = pEnumFeat.Next.ShapeCopy
                Do
                    If pGeom_old.GeometryType <> ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline Then ShallIContinue = InputBox("Not a polyline.  Continue? (y/n)")
                Loop Until pGeom_old.GeometryType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline Or ShallIContinue = "n"

                pGc = pGeom_old

                If pGc.GeometryCount > 1 Then
                    MsgBox("Geometry Cleaned")
                    Call CleanGeometry(pGeom_old)
                End If
                pEnumFeat.Reset()

                If pGeom_old.SpatialReference.Name = "Unknown" Then
                    MsgBox("Please define a spatial reference for your input data.  Stopping Tool.")
                    My.ArcMap.Application.CurrentTool = Nothing
                    Return
                End If
                DefaultSpatialReference = pGeom_old.SpatialReference

                MsgBox("Select the 'from' centerline (used to store distance data)")

                steps = steps + 1


            Case 1

                'GET THE NEWER CENTERLINE

                pEnvelope = pMxDoc.CurrentLocation.Envelope
                pEnvelope.Expand(pMxDoc.SearchTolerance, pMxDoc.SearchTolerance, False)

                'Refresh the old selection to erase it
                pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)
                'Perform the selection using a point created on mouse down
                pMap.SelectByShape(pEnvelope, pMxApp.SelectionEnvironment, True)
                'Refresh again to draw the new selection
                pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)

                pMxDoc = My.ArcMap.Application.Document

                pEnumFeat = pMxDoc.FocusMap.FeatureSelection
                If pMxDoc.FocusMap.SelectionCount <> 1 Then
                    MsgBox("Nothing selected.  Stopping tool.")
                    My.ArcMap.Application.CurrentTool = Nothing
                    Return
                End If

                pEnumFeat.Reset()

                pGeom_new = pEnumFeat.Next.ShapeCopy

                Do
                    If pGeom_new.GeometryType <> ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline Then ShallIContinue = InputBox("Not a polyline.  Continue? (y/n)")
                Loop Until pGeom_new.GeometryType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline Or ShallIContinue = "n"

                pGc = pGeom_new

                If pGc.GeometryCount > 1 Then
                    MsgBox("Geometry Cleaned")
                    Call CleanGeometry(pGeom_new)
                End If

                If pGeom_new.SpatialReference.Name <> pGeom_old.SpatialReference.Name Then
                    MsgBox("Spatial Reference of new and old datasets do not match.  Stopping Tool.")
                    My.ArcMap.Application.CurrentTool = Nothing
                    Return
                End If

                'GET THE SHAPEFILE REPRESENTING THE APEX TRAJECTORY (ASSUMED TO BE A LINE CONNECTING OLD APEX WITH NEW APEX
                ConsiderApexLines = InputBox("Enter Y if apex lines are to be considered")
                If ConsiderApexLines = "Y" Then
                    pApexLineInFeatureclass = GetShapefile

                    If pApexLineInFeatureclass Is Nothing Then
                        MsgBox("Error selecting Shapefile.  Exiting.")
                        Exit Sub
                    End If

                    'Exit if featureclass has no shapes
                    lApexLineTotalFeatureCount = pApexLineInFeatureclass.FeatureCount(Nothing)
                    If lApexLineTotalFeatureCount = 0 Then
                        MsgBox("No features found in shapefile. Exiting")
                        Exit Sub
                    End If
                Else
                    pApexLineInFeatureclass = Nothing
                End If


                m_pGxObjectFilter = Nothing
                m_pGxDialog = Nothing
                '****************

                'COMPUTE OUTWARD NORMAL MIGRATION USING BEZIER CURVES
                Call GetMigration(Mig, m_old, m, Mig1, Mig2, Mig3, Mig4, pGeom_new, pGeom_old, pActiveView, pApexLineInFeatureclass, ConsiderApexLines)
                'pActiveView.PartialRefresh(esriViewDrawPhase.esriViewAll, Nothing, Nothing)
                'EXPORT THE CENTERLINE AS A NEW POLYLINE
                'Call CreateMultipleSegments(xy_centerline, Mig)

                Dim pPolyline As IPolyline
                pPolyline = pGeom_new

                Call CreateMultiplePolys(pPolyline, Mig, m, m_old, Mig1, Mig2, Mig3, Mig4)
                'MsgBox("test")
                steps = steps + 1

        End Select

        If steps = 2 Then My.ArcMap.Application.CurrentTool = Nothing
    End Sub

    Private Function GetShapefile() As IFeatureClass
        Dim pEnumGxObject As IEnumGxObject = Nothing
        Dim pFeatureClass As IFeatureClass = Nothing
        Dim pGxDataset As IGxDataset

        On Error GoTo ErrorHandler

        'Have the user select a shapefile
        m_pGxDialog = New GxDialog
        m_pGxObjectFilter = New GxFilterShapefiles

        m_pGxDialog.ObjectFilter = m_pGxObjectFilter
        m_pGxDialog.Title = "Select the Polyline Shapefile That Represents Bend Apex Paths"
        If m_pGxDialog.DoModalOpen(0, pEnumGxObject) Then
            pEnumGxObject.Reset()
            pGxDataset = pEnumGxObject.Next
            pFeatureClass = pGxDataset.Dataset
        End If
        GetShapefile = pFeatureClass

        Exit Function

ErrorHandler:
        GetShapefile = Nothing
    End Function
    Private Sub GetMigration(ByRef Mig() As Double, ByRef StreamwiseCoordinateOnOldCL() As Double, ByRef StreamwiseCoordinateonNewCL() As Double, ByRef mig1() As Double, ByRef mig2() As Double, ByRef mig3() As Double, ByRef mig4() As Double, ByVal pPg_new As IGeometry, ByVal pPg_old As IGeometry, ByRef pActiveView As IActiveView, ByVal pApexLineFeatureclass As IFeatureClass, ByVal ConsiderApexLines As String)
        Dim i As Integer
        Dim pPc_new As IPointCollection
        Dim pPl_new As IPolyline6
        Dim pPl_old As IPolyline6
        Dim pPl_old_subcurve As IPolyline6
        Dim pBezier As IBezierCurve3
        Dim pMigPoints As IPointCollection2 = New Polyline
        Dim pStartPt As IPoint
        Dim last_distance As Double
        Dim last_start_fraction As Double
        Dim from_fraction As Double
        Dim pApexLineFeature As IFeature
        Dim pApexLineFeatureCursor As IFeatureCursor
        Dim pApexLineGeometry As IGeometry
        Dim pApexPLine As IPolyline
        Dim dummy_boolean As Boolean
        Dim dummy_double As Double
        Dim dummy_double2 As Double
        Dim found As Boolean
        Dim found_on_apex_line As Boolean
        Dim pBezierToApexLine As IBezierCurve3
        Dim bMigratedFromLeft As Boolean
        Dim pModifiedOldCLPLine As IPolyline
        Dim pStatusBar As IStatusBar
        Dim pProgbar As IStepProgressor

        pStatusBar = My.ArcMap.Application.StatusBar
        pProgbar = pStatusBar.ProgressBar

        pPl_new = pPg_new
        pPl_old = pPg_old
        pPl_old_subcurve = pPg_old
        pPc_new = pPl_new
        pStartPt = New Point
        pBezier = New BezierCurve
        pBezierToApexLine = New BezierCurve
        pModifiedOldCLPLine = New Polyline

        pProgbar.Position = 0
        pStatusBar.ShowProgressBar("Working...", 0, pPc_new.PointCount - 1, 1, True)
        pPl_new.QueryPoint(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, 100, False, pStartPt)
        last_distance = 0
        last_start_fraction = 0
        from_fraction = 0


        'Get the line midway between the old and new centerlines
        Dim IntermediatePlineCenter As IPointCollection2 = New Polyline
        IntermediatePlineCenter = FindMidCenterline(pPl_old, pPl_new)
        'CreateDrawingElement(IntermediatePlineCenter, pActiveView)


        'Get the line midway between the old centerline and the interpolated midway line
        Dim IntermediatePlineOld As IPointCollection2 = New Polyline    'represents the entire centerline at 25% of the time interval
        IntermediatePlineOld = FindMidCenterline(pPl_old, IntermediatePlineCenter)
        'CreateDrawingElement(IntermediatePlineOld, pActiveView)

        'Get the line midway between the midway line and the new centerline
        Dim IntermediatePlineNew As IPointCollection2 = New Polyline    'represents the entire centerline at 75% of the time interval
        IntermediatePlineNew = FindMidCenterline(IntermediatePlineCenter, pPl_new)
        'CreateDrawingElement(IntermediatePlineNew, pActiveView)
        'MsgBox("done")


        'Check whether there are any apex lines and modify the intermediate centerlines near the apices of the bends that
        'are flagged as translating downstream rather than cross stream.

        If ConsiderApexLines = "Y" Then
            pApexLineFeatureCursor = pApexLineFeatureclass.Search(Nothing, False)
            pApexLineFeature = pApexLineFeatureCursor.NextFeature
            'lApexLineFeatureCount = 0
            'Dim pBendMidSegment As IPolyline
            Do While (Not pApexLineFeature Is Nothing)
                pApexLineGeometry = pApexLineFeature.Shape
                pApexPLine = pApexLineGeometry
                AdjustCenterlineNearApexTrajectory(IntermediatePlineCenter, pPl_old, pPl_new, pApexPLine, pActiveView)
                IntermediatePlineOld = FindMidCenterline(pPl_old, IntermediatePlineCenter)
                Dim newApexLine As IPolyline6 = New Polyline
                pApexPLine.GetSubcurve(0, 0.5, True, newApexLine)
                AdjustCenterlineNearApexTrajectory(IntermediatePlineOld, pPl_old, IntermediatePlineCenter, newApexLine, pActiveView)
                IntermediatePlineNew = FindMidCenterline(IntermediatePlineCenter, pPl_new)
                pApexPLine.GetSubcurve(0.5, 1, True, newApexLine)
                AdjustCenterlineNearApexTrajectory(IntermediatePlineNew, IntermediatePlineCenter, pPl_new, newApexLine, pActiveView)
                pApexLineFeature = pApexLineFeatureCursor.NextFeature
            Loop

        End If
        'Write the mid-lines to the screen
        Call CreateDrawingElement(IntermediatePlineOld, pActiveView)
        Call CreateDrawingElement(IntermediatePlineNew, pActiveView)
        Call CreateDrawingElement(IntermediatePlineCenter, pActiveView)

        Dim Trajectory As ICurve

        For i = 1 To pPc_new.PointCount - 1
            ReDim Preserve Mig(i)
            ReDim Preserve StreamwiseCoordinateOnOldCL(i)
            ReDim Preserve StreamwiseCoordinateonNewCL(i)
            ReDim Preserve mig1(i)
            ReDim Preserve mig2(i)
            ReDim Preserve mig3(i)
            ReDim Preserve mig4(i)

            pStatusBar.StepProgressBar()
            pStatusBar.Message(0) = "i = " & i
            found = False
            found_on_apex_line = False
            pStartPt = pPc_new.Point(i)
            Dim pPolyTest As IPolyline5 = IntermediatePlineNew
            Dim pNextPoint As IPoint = New Point
            Do While Not found
                pStatusBar.Message(0) = "i = " & i
                pMigPoints = New Polyline
                pMigPoints.AddPoint(pStartPt)
                pPolyTest = IntermediatePlineNew
                pPolyTest.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pMigPoints.Point(0), False, pNextPoint, dummy_double, mig1(i), dummy_boolean)
                pMigPoints.AddPoint(pNextPoint)
                pPolyTest = IntermediatePlineCenter
                pNextPoint = New Point
                pPolyTest.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pMigPoints.Point(1), False, pNextPoint, dummy_double, mig2(i), dummy_boolean)
                pMigPoints.AddPoint(pNextPoint)
                pPolyTest = IntermediatePlineOld
                pNextPoint = New Point
                pPolyTest.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pMigPoints.Point(2), False, pNextPoint, dummy_double, mig3(i), dummy_boolean)
                pMigPoints.AddPoint(pNextPoint)
                pPolyTest = pPl_old
                pNextPoint = New Point
                pPolyTest.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pMigPoints.Point(3), False, pNextPoint, dummy_double, mig4(i), dummy_boolean)
                pMigPoints.AddPoint(pNextPoint)
                '*******************************
                My.ArcMap.Document.ActiveView.Refresh()
                Call CreateDrawingElement(pMigPoints, pActiveView)
                Trajectory = pMigPoints
                Mig(i) = Trajectory.Length

                'determine the direction migration has been from and change sign of migration rate term accordingly
                'pPl_new.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pMigPoints.Point(4), False, Nothing, StreamwiseCoordinateOnOldCL(i), dummy_double2, bMigratedFromLeft)
                pPl_new.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pMigPoints.Point(4), False, Nothing, dummy_double, dummy_double2, bMigratedFromLeft)
                If Not bMigratedFromLeft Then
                    Mig(i) = -Mig(i)
                    mig1(i) = -mig1(i)
                    mig2(i) = -mig2(i)
                    mig3(i) = -mig3(i)
                    mig4(i) = -mig4(i)
                End If

                'get down channel distance for old and now centerline points
                pPl_new.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pMigPoints.Point(0), False, Nothing, StreamwiseCoordinateonNewCL(i), dummy_double2, bMigratedFromLeft)
                'pPolyTest = pPl_new
                'pPolyTest = pPl_old
                'pPolyTest.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pMigPoints.Point(0), False, Nothing, StreamwiseCoordinateonNewCL(i), dummy_double, dummy_boolean)
                'pPolyTest.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pMigPoints.Point(4), False, Nothing, StreamwiseCoordinateOnOldCL(i), dummy_double, dummy_boolean)
                pPl_old.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pMigPoints.Point(4), False, Nothing, StreamwiseCoordinateOnOldCL(i), dummy_double, dummy_boolean)
                found = True
            Loop

            'Get the approximate streamwise coordinate of the old centerline. 
            ReDim Preserve StreamwiseCoordinateOnOldCL(i)
            pPl_new.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pMigPoints.Point(4), False, Nothing, dummy_double, dummy_double2, bMigratedFromLeft)
            'pActiveView.Refresh
        Next i
        pStatusBar.HideProgressBar()
        pActiveView.Refresh()

    End Sub
    Private Function FindMidCenterline(ByVal pPl_old As IPolyline6, ByVal pPl_new As IPolyline6) As IPointCollection2

        Dim pdummypoint As IPoint = Nothing
        Dim dummy_double As Double
        Dim dummy_boolean As Boolean

        'find all points where the two centerlines cross
        Dim pTopoOptr As ITopologicalOperator5
        pTopoOptr = pPl_old
        Dim pPtColl As IPointCollection
        pPtColl = pTopoOptr.Intersect(pPl_new, ESRI.ArcGIS.Geometry.esriGeometryDimension.esriGeometry0Dimension)

        'sort points by down-channel distance
        Dim pReplacePoints As IPointCollection = New Polyline
        Dim switched As Boolean = False
        Dim dummypoint As IPoint = New Point
        Dim j As Integer
        For j = 0 To pPtColl.PointCount - 1
            pPl_new.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPtColl.Point(j), False, dummypoint, pPtColl.Point(j).M, dummy_double, dummy_boolean)
        Next
        Do
            switched = False
            For j = 1 To pPtColl.PointCount - 1
                If pPtColl.Point(j).M < pPtColl.Point(j - 1).M Then
                    pReplacePoints.AddPoint(pPtColl.Point(j))
                    pReplacePoints.AddPoint(pPtColl.Point(j - 1))
                    pPtColl.RemovePoints(j - 1, 2)
                    pPtColl.InsertPointCollection(j - 1, pReplacePoints)
                    switched = True
                    pReplacePoints.RemovePoints(0, 2)
                End If
            Next j
        Loop Until switched = False
        'end sort

        'find the midpoints and write them into a new pointcollection
        Dim pPtCollMidpointsOld As IPointCollection = New Polyline
        Dim pPtCollMidpointsNew As IPointCollection = New Polyline
        Dim startFraction As Double
        Dim endFraction As Double
        Dim pDummyPoint3 As IPoint = New Point
        Dim NewMidPoint As IPoint
        pPtCollMidpointsOld.AddPoint(pPl_old.FromPoint)
        pPtCollMidpointsNew.AddPoint(pPl_new.FromPoint)

        For j = 1 To pPtColl.PointCount - 1
            pPl_old.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPtColl.Point(j - 1), True, pDummyPoint3, startFraction, dummy_double, dummy_boolean)
            pPl_old.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPtColl.Point(j), True, pDummyPoint3, endFraction, dummy_double, dummy_boolean)
            NewMidPoint = New Point
            pPl_old.QueryPoint(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, startFraction / 2 + endFraction / 2, True, NewMidPoint)
            pPtCollMidpointsOld.AddPoint(NewMidPoint)

            pPl_new.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPtColl.Point(j - 1), True, pDummyPoint3, startFraction, dummy_double, dummy_boolean)
            pPl_new.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPtColl.Point(j), True, pDummyPoint3, endFraction, dummy_double, dummy_boolean)
            NewMidPoint = New Point
            pPl_new.QueryPoint(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, startFraction / 2 + endFraction / 2, True, NewMidPoint)
            pPtCollMidpointsNew.AddPoint(NewMidPoint)

        Next
        pPtCollMidpointsOld.AddPoint(pPl_old.ToPoint)
        pPtCollMidpointsNew.AddPoint(pPl_new.ToPoint)
        'end writing midpoints into a new pointcollection
        '*********************************************
        'Interpolate the centerline 

        Dim pPolylineSeg_old As IPolyline6 = Nothing
        Dim pPolylineSeg_new As IPolyline6 = Nothing

        Dim startFractionOld As Double
        Dim startFractionNew As Double
        Dim endFractionOld As Double
        Dim endFractionNew As Double
        Dim IntermediatePlineCenter As IPointCollection2 = New Polyline 'represents the entire centerline at 50% of the time interval
        Dim pPolyCL As IPolyline6 = Nothing             'represents the centerline segment j at 50% of the time between old and new centerlines

        For j = 1 To pPtCollMidpointsOld.PointCount - 1 'pPtColl.PointCount - 1
            pPl_old.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPtCollMidpointsOld.Point(j - 1), True, pdummypoint, startFractionOld, dummy_double, dummy_boolean)
            pPl_old.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPtCollMidpointsOld.Point(j), True, pdummypoint, endFractionOld, dummy_double, dummy_boolean)
            pPl_old.GetSubcurve(startFractionOld, endFractionOld, True, pPolylineSeg_old)
            pPl_new.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPtCollMidpointsNew.Point(j - 1), True, pdummypoint, startFractionNew, dummy_double, dummy_boolean)
            pPl_new.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPtCollMidpointsNew.Point(j), True, pdummypoint, endFractionNew, dummy_double, dummy_boolean)
            pPl_new.GetSubcurve(startFractionNew, endFractionNew, True, pPolylineSeg_new)

            pPolyCL = PolyDivide(pPolylineSeg_old, pPolylineSeg_new)
            IntermediatePlineCenter.AddPointCollection(pPolyCL)

        Next j
        FindMidCenterline = IntermediatePlineCenter
    End Function
    Private Sub AdjustCenterlineNearApexTrajectory(ByRef CenterlineToAdjust As IPolyline6, ByVal OldestCenterline As IPolyline6, ByVal NewestCenterline As IPolyline6, ByVal ApexLine As IPolyline6, ByVal pActiveView As IActiveView)
        'Interpolate the line mid-way between the new and old centerline near the bend apex that is translating
        Dim pBendMidSegment As IPolyline6
        Dim PointsOfIntersection As IPointCollection2 = New Multipoint
        Dim IntermediateApexLine As IPointCollection2 = New Polyline

        Dim Multiplier As Double
        'Dim pIntCenterPline As IPolyline6 = IntermediatePlineCenter
        Dim counter As Integer
        Dim pTopoOptr As ITopologicalOperator5
        Multiplier = 2
        Do   ' make sure that there are only two points of intersection between the interpolated bend and the uninterpolated centerline
            pBendMidSegment = CreateBendTranslationSegment(OldestCenterline, NewestCenterline, Multiplier, ApexLine, pActiveView)
            'CreateDrawingElement(pBendMidSegment, pActiveView)
            pTopoOptr = CenterlineToAdjust
            PointsOfIntersection = pTopoOptr.Intersect(pBendMidSegment, ESRI.ArcGIS.Geometry.esriGeometryDimension.esriGeometry0Dimension)
            If PointsOfIntersection.PointCount > 2 Then Multiplier = 0.9 * Multiplier
            If PointsOfIntersection.PointCount < 2 Then Multiplier = 1.1 * Multiplier
            counter = counter + 1
            If counter > 20 Then
                MsgBox("no solution found for one of the translating bends")
                Return
            End If
        Loop Until PointsOfIntersection.PointCount = 2
        'Patch together the translated bend centerline and the basic interpolated centerline
        Dim pPath As IPath
        Dim pGeomColl As IGeometryCollection = pBendMidSegment
        pPath = pGeomColl.Geometry(0)
        Dim dummy_boolean As Boolean
        Dim dummypoint As IPoint = New Point
        Dim dummy_double As Double
        dummy_boolean = CenterlineToAdjust.Reshape(pPath)
        'Smooth the patched centerline
        Dim pPartToSmooth As IPolyline6 = New Polyline
        Dim startdist As Double
        Dim enddist As Double
        CenterlineToAdjust.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, ApexLine.FromPoint, False, dummypoint, startdist, dummy_double, dummy_boolean)
        startdist = startdist - ApexLine.Length * 2
        CenterlineToAdjust.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, ApexLine.ToPoint, False, dummypoint, enddist, dummy_double, dummy_boolean)
        enddist = enddist + ApexLine.Length * 2
        CenterlineToAdjust.GetSubcurve(startdist, enddist, False, pPartToSmooth)
        'pPartToSmooth.Smooth(ApexLine.Length / 20)
        pGeomColl = pPartToSmooth
        pPath = pGeomColl.Geometry(0)
        dummy_boolean = CenterlineToAdjust.Reshape(pPath)
        CreateDrawingElement(CenterlineToAdjust, pActiveView)
        MsgBox("test")
    End Sub

    Private Function CreateBendTranslationSegment(ByVal pCL_Old As IPolyline6, ByVal pCL_New As IPolyline6, ByVal LengthMultiplier As Double, ByVal pApexLine As IPolyline6, ByRef pActiveView As IActiveView) As IPolyline6
        Dim i As Integer
        'Dim pPointStartOld As IPoint = New Point
        'Dim pPointStartNew As IPoint = New Point
        Dim BendHalfLength As Double
        Dim StartMOld As Double
        Dim ApexMOld As Double
        Dim EndMOld As Double
        Dim StartMNew As Double
        Dim ApexMNew As Double
        Dim EndMNew As Double
        '       Dim DummyPoint As IPoint
        Dim DummyDouble As Double
        Dim DummyBoolean As Boolean
        Dim TranslationSegment As IPointCollection2 = New Polyline
        Dim PointsInNewCL As Integer
        Dim SubCurveNew As IPointCollection = New Polyline
        Dim SubCurveOld As IPointCollection = New Polyline
        Dim NewPoint As IPoint
        Dim OldPoint As IPoint
        Dim CenterPoint As IPoint
        BendHalfLength = pApexLine.Length

        'Find distance into curve of staring point, ending point, and apex

        pCL_New.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pApexLine.ToPoint, False, Nothing, ApexMNew, DummyDouble, DummyBoolean)
        StartMNew = ApexMNew - BendHalfLength * LengthMultiplier
        EndMNew = ApexMNew + BendHalfLength * LengthMultiplier

        pCL_Old.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pApexLine.FromPoint, False, Nothing, ApexMOld, DummyDouble, DummyBoolean)
        StartMOld = ApexMOld - BendHalfLength * LengthMultiplier
        EndMOld = ApexMOld + BendHalfLength * LengthMultiplier

        'Split up centerlines into sub-curves representing this bend only

        Dim SubCurveNew_1stHalf As IPointCollection = New Polyline
        Dim SubCurveNew_2ndHalf As IPointCollection = New Polyline

        Dim SubCurveOld_1stHalf As IPointCollection = New Polyline
        Dim SubCurveOld_2ndHalf As IPointCollection = New Polyline


        pCL_New.GetSubcurve(StartMNew, EndMNew, False, SubCurveNew)
        pCL_Old.GetSubcurve(StartMOld, EndMOld, False, SubCurveOld)




        pCL_New.GetSubcurve(StartMNew, ApexMNew, False, SubCurveNew_1stHalf)
        pCL_New.GetSubcurve(ApexMNew, EndMNew, False, SubCurveNew_2ndHalf)


        pCL_Old.GetSubcurve(StartMOld, ApexMOld, False, SubCurveOld_1stHalf)
        pCL_Old.GetSubcurve(ApexMOld, EndMOld, False, SubCurveOld_2ndHalf)


        PointsInNewCL = 100 'SubCurveNew.PointCount * 5
        'Dim DistanceToPoint As Double
        Dim SubCurveNewAsPolyline As IPolyline6 = SubCurveNew
        Dim SubCurveOldAsPolyline As IPolyline6 = SubCurveOld
        Dim SubCurveNew_1stHalfAsPolyline As IPolyline6 = SubCurveNew_1stHalf
        Dim SubCurveNew_2ndHalfAsPolyline As IPolyline6 = SubCurveNew_2ndHalf
        Dim SubCurveOld_1stHalfAsPolyline As IPolyline6 = SubCurveOld_1stHalf
        Dim SubCurveOld_2ndHalfAsPolyline As IPolyline6 = SubCurveOld_2ndHalf


        Dim Fraction As Double

        'Sample a point on the new interpolated centerline segment for each point on the new centerline for this apex
        'Points are the midpoint between the vertices on the new centerline and a corresponding point the same fraction into the old
        'centerline segment.

        '        For i = 0 To PointsInNewCL - 1 'SubCurveNew.PointCount - 1
        ' NewPoint = New Point
        ' OldPoint = New Point
        ' CenterPoint = New Point
        ' 'SubCurveNew.QueryPoint(i, NewPoint)
        ' 'SubCurveNewAsPolyline.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, NewPoint, True, Nothing, DistanceToPoint, DummyDouble, DummyBoolean)
        ' Fraction = (i + 1) / (PointsInNewCL + 1)
        ' SubCurveNewAsPolyline.QueryPoint(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, Fraction, True, NewPoint)
        ' SubCurveOldAsPolyline.QueryPoint(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, Fraction, True, OldPoint)
        ' CenterPoint.PutCoords(NewPoint.X / 2 + OldPoint.X / 2, NewPoint.Y / 2 + OldPoint.Y / 2)
        ' TranslationSegment.AddPoint(CenterPoint)
        ' Next

        For i = 0 To PointsInNewCL - 1 'SubCurveNew.PointCount - 1
            NewPoint = New Point
            OldPoint = New Point
            CenterPoint = New Point
            'SubCurveNew.QueryPoint(i, NewPoint)
            'SubCurveNewAsPolyline.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, NewPoint, True, Nothing, DistanceToPoint, DummyDouble, DummyBoolean)
            Fraction = (i + 1) / (PointsInNewCL + 1)
            SubCurveNew_1stHalfAsPolyline.QueryPoint(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, Fraction, True, NewPoint)
            SubCurveOld_1stHalfAsPolyline.QueryPoint(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, Fraction, True, OldPoint)
            CenterPoint.PutCoords(NewPoint.X / 2 + OldPoint.X / 2, NewPoint.Y / 2 + OldPoint.Y / 2)
            TranslationSegment.AddPoint(CenterPoint)
        Next

        CenterPoint = New Point
        CenterPoint.PutCoords(pApexLine.FromPoint.X / 2 + pApexLine.ToPoint.X / 2, pApexLine.FromPoint.Y / 2 + pApexLine.ToPoint.Y / 2)
        TranslationSegment.AddPoint(CenterPoint)

        For i = 0 To PointsInNewCL - 1 'SubCurveNew.PointCount - 1
            NewPoint = New Point
            OldPoint = New Point
            CenterPoint = New Point
            'SubCurveNew.QueryPoint(i, NewPoint)
            'SubCurveNewAsPolyline.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, NewPoint, True, Nothing, DistanceToPoint, DummyDouble, DummyBoolean)
            Fraction = (i + 1) / (PointsInNewCL + 1)
            SubCurveNew_2ndHalfAsPolyline.QueryPoint(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, Fraction, True, NewPoint)
            SubCurveOld_2ndHalfAsPolyline.QueryPoint(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, Fraction, True, OldPoint)
            CenterPoint.PutCoords(NewPoint.X / 2 + OldPoint.X / 2, NewPoint.Y / 2 + OldPoint.Y / 2)
            TranslationSegment.AddPoint(CenterPoint)
        Next





        Dim dummypolly As IPolyline6 = TranslationSegment

        CreateDrawingElement(dummypolly, pActiveView)

        dummypolly = SubCurveOld
        'CreateDrawingElement(dummypolly, pActiveView)
        CreateBendTranslationSegment = TranslationSegment
        MsgBox("Translation Segment Added")
    End Function

    Private Sub DrawBezier(ByRef pStartPt As IPoint, ByRef pEndPt As IPoint, ByRef pPl_start As IPolyline, ByRef pPl_end As IPolyline, ByRef pBezier As IBezierCurve3, ByRef pActiveView As IActiveView)
        'Creates a bezier curve between two lines such that the curve is on the start point and end point,
        'normal to both lines, and has interior bezier points set away from the respective lines at a distance
        'equal to half the straight-line distance between the start and end points.

        'Dim pNormalPoint1 As IPoint
        'Dim pNormalPoint2 As IPoint
        Dim distance_along_curve As Double
        Dim distance_from_curve As Double
        Dim dummy_side As Boolean
        Dim end_on_right_side As Boolean
        Dim pBezierPoints(3) As IPoint
        Dim pLine As ILine
        Dim Length As Double
        Dim pDummyPoint As IPoint
        Dim pDummyBezier As IBezierCurve3

        pDummyPoint = New Point
        pLine = New Line
        pBezierPoints(0) = New Point
        pBezierPoints(1) = New Point
        pBezierPoints(2) = New Point
        pBezierPoints(3) = New Point

        pLine.PutCoords(pStartPt, pEndPt)
        'Call CreateDrawingElement(pLine, pActiveView)

        Length = pLine.Length / 1.5
        'Determine if the end point is on the right or left side of the starting line
        pPl_start.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pEndPt, False, pDummyPoint, distance_along_curve, distance_from_curve, end_on_right_side)

        If Not end_on_right_side Then Length = -Length

        pPl_start.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pStartPt, False, pBezierPoints(0), distance_along_curve, distance_from_curve, dummy_side)
        pPl_start.QueryNormal(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, distance_along_curve, False, Length, pLine)
        pLine.QueryToPoint(pBezierPoints(1))
        'pPl_start.QueryNormal(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, distance_along_curve, False, Length, pLine)

        pPl_end.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pEndPt, False, pBezierPoints(3), distance_along_curve, distance_from_curve, dummy_side)
        pPl_end.QueryNormal(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, distance_along_curve, False, -Length / 4, pLine)
        pLine.QueryToPoint(pBezierPoints(2))

        pDummyBezier = New BezierCurve

        pDummyBezier.PutCoord(0, pBezierPoints(0))
        pDummyBezier.PutCoord(1, pBezierPoints(1))
        pDummyBezier.PutCoord(2, pBezierPoints(2))
        pDummyBezier.PutCoord(3, pBezierPoints(3))

        'Dim pScreenDisplay As IScreenDisplay = pActiveView.ScreenDisplay
        'Dim pSymbol As ISimpleMarkerSymbol
        'pSymbol = New SimpleMarkerSymbol
        'pSymbol.Size = 4
        pDummyBezier.PutCoord(0, pBezierPoints(0))
        pDummyBezier.PutCoord(1, pBezierPoints(1))
        pDummyBezier.PutCoord(2, pBezierPoints(2))
        pDummyBezier.PutCoord(3, pBezierPoints(3))

        'With pScreenDisplay
        ' .StartDrawing(pScreenDisplay.hDC, -1)
        ' .SetSymbol(pSymbol)
        ' .DrawPoint(pBezierPoints(0))
        ' .DrawPoint(pBezierPoints(1))
        ' .DrawPoint(pBezierPoints(2))
        ' .DrawPoint(pBezierPoints(3))
        ' .FinishDrawing()
        ' End With


        'Call CreateDrawingElement(pBezier, pActiveView)
        pBezier = pDummyBezier

        '
        'Call CreateDrawingElement(pBezier, pActiveView)


    End Sub
    Private Function PolyDivide(ByVal Pline1 As IPolyline6, ByVal Pline2 As IPolyline6) As IPolyline6
        Dim i As Integer
        'Dim Densified_Pline1 As IPointCollection = Nothing
        'Dim Densified_Pline2 As IPointCollection = Nothing
        'Dim RayCollection As IPointCollection4 = Nothing
        Dim Centerline As IPointCollection4 = New Polyline
        Dim Pline1Col As IPointCollection4 = Pline1
        Dim Pline2Col As IPointCollection4 = Pline2
        Dim pPoint1 As IPoint = Nothing
        Dim pPoint2 As IPoint = Nothing
        Dim pPoint3 As IPoint = Nothing
        Dim pSeg1 As IPolyline6 = Nothing
        Dim pSeg2 As IPolyline6 = Nothing
        'Dim StartPoint As IPoint = New Point
        'Dim EndPoint As IPoint = New Point

        Dim CenterlinePoint As IPoint = Nothing
        Dim Ray As ILine2 = Nothing
        Dim Fraction As Double
        Dim maxPoints As Integer
        Dim x As Double
        Dim y As Double
        PolyDivide = Nothing

        maxPoints = Math.Max(Pline1Col.PointCount, Pline2Col.PointCount)
        If Pline1 Is Nothing Then
            MsgBox("Pline 1 is invalid")
            Return Nothing
        Else
            pPoint3 = New Point
            pPoint3.PutCoords(Pline1.FromPoint.X / 2 + Pline2.FromPoint.X / 2, Pline1.FromPoint.Y / 2 + Pline2.FromPoint.Y / 2)
            Centerline.AddPoint(pPoint3)
            For i = 1 To maxPoints
                Fraction = i / (maxPoints + 1)
                '            RayCollection = Nothing
                Pline1.GetSubcurve(0, Fraction, True, pSeg1)
                pPoint1 = pSeg1.ToPoint
                Pline2.GetSubcurve(0, Fraction, True, pSeg2)
                pPoint2 = pSeg2.ToPoint
                x = pPoint1.X / 2 + pPoint2.X / 2
                y = pPoint1.Y / 2 + pPoint2.Y / 2

                pPoint3 = New Point
                pPoint3.PutCoords(x, y)
                Centerline.AddPoint(pPoint3)

                
            Next i
            pPoint3 = New Point
            pPoint3.PutCoords(Pline1.ToPoint.X / 2 + Pline2.ToPoint.X / 2, Pline1.ToPoint.Y / 2 + Pline2.ToPoint.Y / 2)
            Centerline.AddPoint(pPoint3)
            End If
        PolyDivide = Centerline


        'Dim pScreenDisplay As IScreenDisplay = pActiveView.ScreenDisplay
        'Dim pSymbol As ISimpleLineSymbol
        'pSymbol = New SimpleLineSymbol
        'pSymbol.Size = 4

        'With pScreenDisplay
        ' .StartDrawing(pScreenDisplay.hDC, -1)
        ' .SetSymbol(pSymbol)
        ' .DrawPolyline(PolyDivide)
        '.DrawPoint(pBezierPoints(0))
        ' .DrawPoint(pBezierPoints(1))
        ' .DrawPoint(pBezierPoints(2))
        ' .DrawPoint(pBezierPoints(3))
        '.FinishDrawing()
        'End With

    End Function
    Private Sub FindBestBezier(ByRef pBezier As IBezierCurve3, ByVal pStartPt As IPoint, ByVal pPl_old As IPolyline, ByVal pPl_new As IPolyline, ByRef last_distance As Double, ByRef pActiveView As IActiveView) ', distance_increment As Double)
        ' Finds the shortest bezier curve on the input point and normal to both polylines with interior
        ' bezier points set at a distance equal to half the straight-line distance between end-points of the bezier
        ' curve.

        Dim i As Integer
        Dim j As Integer
        Dim pEndPoint As IPoint
        Dim pBestEndPoint As IPoint
        Dim best_distance As Double
        Dim best_length As Double
        Dim loop_starting_distance As Double
        Dim start_on_rightside As Boolean
        Dim end_on_rightside As Boolean
        Dim dummy_double1 As Double
        Dim dummy_double2 As Double
        Dim distance_increment As Double
        Dim nearest_streamwise_distance As Double
        Dim nearest_normal_distance As Double

        pBezier = New BezierCurve
        pEndPoint = New Point
        pBestEndPoint = New Point

        best_length = 9999
        loop_starting_distance = last_distance

        'determine if the starting point is on the right side of the old centerline.  Also determine the distance
        'along the old centerline to the nearest point to the starting point (which is on the new centerline).

        pPl_old.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pStartPt, False, Nothing, nearest_streamwise_distance, nearest_normal_distance, start_on_rightside)

        'set the initial distance increment to be 1/5 the distance along the old centerline to the nearest point
        'to the start point minus the distance along the centerline to the last point.  This ensures that in the first loop
        'wherein the program attempts to find short bezier paths, it will try the path between the start point
        'and the nearest point on the old centerline to the start point.

        distance_increment = 0.2 * (nearest_streamwise_distance - last_distance)

        If distance_increment < 0.2 * nearest_normal_distance Then
            distance_increment = 0.2 * nearest_normal_distance
        End If

        For j = 1 To 4
            For i = 0 To 10
                pPl_old.QueryPoint(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, distance_increment * i + loop_starting_distance, False, pEndPoint)

                'check which side of the new centerline the end point is on
                pPl_new.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pEndPoint, False, Nothing, dummy_double1, dummy_double2, end_on_rightside)

                Call DrawBezier(pStartPt, pEndPoint, pPl_new, pPl_old, pBezier, pActiveView)
                If pBezier.Length < best_length And start_on_rightside <> end_on_rightside Then
                    best_length = pBezier.Length
                    best_distance = distance_increment * i + loop_starting_distance
                    'MsgBox("Best Length " & best_length & "  :  Bezier Length " & pBezier.Length)
                    'Call CreateDrawingElement(pBezier, pActiveView)
                End If
            Next i

            loop_starting_distance = best_distance - distance_increment
            If loop_starting_distance < last_distance Then loop_starting_distance = last_distance
            distance_increment = 2 * distance_increment / 10

        Next j


        'See if the best end point falls past the point where the apex line intersects the old
        'centerline, but not more than twice the distance along the centerline from this point
        'to the new apex point. Apices are defined by the apex features, which should be lines
        'connecting apices and are to be digitized by hand prior to using this program.

        pPl_old.QueryPoint(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, best_distance, False, pBestEndPoint)
        last_distance = best_distance
        Call DrawBezier(pStartPt, pBestEndPoint, pPl_new, pPl_old, pBezier, pActiveView)
        'MsgBox "Best Length " & best_length & "  :  Bezier Length " & pBezier.Length & " Best EP " & pBestEndPoint.x & " " & pBestEndPoint.y
        'Call CreateDrawingElement(pBezier, pActiveView)

    End Sub


    Private Sub CreateDrawingElement(ByVal pGeomLine As IGeometry5, ByRef pAV As IActiveView)


        'Takes an IGeometry and IActiveView and creates a Line element in the ActiveView's BasicGraphicsLayer
        Dim pElemLine As ILineElement
        Dim pElem As IElement
        Dim pGraCont As IGraphicsContainer
        Dim pSLnSym As ISimpleLineSymbol
        Dim pRGB As IRgbColor
        Dim pPoly As IPolyline6

        Dim pSegColl As ISegmentCollection
        pSegColl = New Polyline
        'Dim pSegment As ISegment
        'pSegment = pGeomLine

        'pSegColl.AddSegment(pGeomLine)
        'pSegColl.AddSegment(pSegment)
        'pPoly = pSegColl
        pPoly = pGeomLine
        ' Use the IElement interface to set the LineElement's Geometry
        pElem = New LineElement
        pElem.Geometry = pPoly

        ' QI for the IFillShapeElement interface so that the Symbol property can be set
        pElemLine = pElem

        ' Create a new RGBColor
        pRGB = New RgbColor
        With pRGB
            .Red = 100
            .Green = 200
            .Blue = 200
        End With

        ' Create a new SimpleFillSymbol and set its Color and Style
        pSLnSym = New SimpleLineSymbol
        pSLnSym.Color = pRGB
        pSLnSym.Style = ESRI.ArcGIS.Display.esriSimpleLineStyle.esriSLSSolid
        pElemLine.Symbol = pSLnSym

        ' QI for the IGraphicsContainer interface from the IActiveView, allows access to the BasicGraphicsLayer
        pGraCont = pAV
        'Add the element at Z order zero
        pGraCont.AddElement(pElemLine, 0)
        pAV.PartialRefresh(esriViewDrawPhase.esriViewAll, Nothing, Nothing)
    End Sub

    Private Sub CreateMultiplePolys(ByVal pPolyline As IPolyline, ByVal migrated_distance() As Double, ByVal m() As Double, ByVal m_old() As Double, ByVal mig1() As Double, ByVal mig2() As Double, ByVal mig3() As Double, ByVal mig4() As Double)
        Dim i As Long
        Dim pi As Double
        '       Dim pGxFile As IGxFile
        Dim pFields As IFields
        Dim path As String = Nothing
        Dim name As String = Nothing

        pi = 3.141592654

        '       Const strShapeFieldName As String = "Shape"

        pFields = New Fields
        Call SetupFields(pFields, DefaultSpatialReference)

        m_pGxDialog = New GxDialog
        Dim pFeatClass As IFeatureClass
        pFeatClass = CreateNewShapefile(pFields, path, name)

        '  Now, create the Line data and add it to the new FeatureClass along with the
        '  specified attributes.

        If pFeatClass Is Nothing Then Exit Sub

        Dim FCA_index_1 As Long
        Dim FCA_index_2 As Long
        Dim FCA_index_3 As Long
        Dim FCA_index_4 As Long
        Dim FCA_index_5 As Long
        Dim FCA_index_6 As Long
        Dim FCA_index_7 As Long
        Dim FCA_index_8 As Long

        FCA_index_1 = pFeatClass.FindField("Mig_dist")
        FCA_index_2 = pFeatClass.FindField("i")
        FCA_index_3 = pFeatClass.FindField("m")
        FCA_index_4 = pFeatClass.FindField("old_m")
        FCA_index_5 = pFeatClass.FindField("Mig_1")
        FCA_index_6 = pFeatClass.FindField("Mig_2")
        FCA_index_7 = pFeatClass.FindField("Mig_3")
        FCA_index_8 = pFeatClass.FindField("Mig_4")


        Dim pGeomcoll As IPointCollection
        '       Dim pSegColl As ISegmentCollection
        '      Dim pLine As ILine
        Dim pPolygon As IPolygon
        Dim pFeat As IFeature
        '        Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
        '        Dim x3 As Double, y3 As Double, x4 As Double, y4 As Double
        Dim pPt1 As IPoint
        Dim pPt2 As IPoint
        Dim pPt3 As IPoint
        Dim pPt4 As IPoint
        '       Dim theta1 As Double
        '       Dim theta2 As Double
        Dim dist As Double
        Dim pMultipoint As IPointCollection
        Dim pClPC As IPointCollection
        Dim pTopologicalOperator As ITopologicalOperator
        Dim pBag As IGeometryCollection

        pBag = New GeometryBag

        pTopologicalOperator = pPolyline
        'MsgBox "Simple? " & pTopologicalOperator.IsSimple
        pPolyline = pTopologicalOperator

        pClPC = pPolyline
        'MsgBox pClPC.PointCount

        dist = InputBox("enter the distance lateral to channel centerline to draw polygons")

        Dim pCLine As ILine
        pCLine = New Line
        Dim pNLine As ILine
        pNLine = New Line

        Dim pStatusBar As IStatusBar
        Dim pProgbar As IStepProgressor
        pStatusBar = My.ArcMap.Application.StatusBar
        pProgbar = pStatusBar.ProgressBar

        pProgbar.Position = 0
        pStatusBar.ShowProgressBar("Working...", 0, pClPC.PointCount - 2, 1, True)


        For i = 1 To pClPC.PointCount - 2

            pStatusBar.StepProgressBar()
            pStatusBar.Message(0) = "i = " & i

            pMultipoint = New Multipoint
            pPt1 = New Point
            pPt2 = New Point
            pPt3 = New Point
            pPt4 = New Point

            pCLine.PutCoords(pClPC.Point(i - 1), pClPC.Point(i))
            pCLine.QueryNormal(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, 0.5, True, dist, pNLine)
            pNLine.QueryToPoint(pPt1)
            pCLine.QueryNormal(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, 0.5, True, -dist, pNLine)
            pNLine.QueryToPoint(pPt2)

            pCLine.PutCoords(pClPC.Point(i), pClPC.Point(i + 1))
            pCLine.QueryNormal(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, 0.5, True, -dist, pNLine)
            pNLine.QueryToPoint(pPt3)
            pCLine.QueryNormal(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, 0.5, True, dist, pNLine)
            pNLine.QueryToPoint(pPt4)

            'Make sure that the two edge lines don't intersect
            'do this later

            pGeomcoll = New Polygon
            pMultipoint.AddPoint(pPt1)
            pMultipoint.AddPoint(pPt2)
            pMultipoint.AddPoint(pPt3)
            pMultipoint.AddPoint(pPt4)
            pMultipoint.AddPoint(pPt1)

            pGeomcoll.AddPointCollection(pMultipoint)
            pPolygon = pGeomcoll
            pPolygon.SpatialReference = DefaultSpatialReference

            'Create the new feature, set the Feature's Shape, and set the migrated distance attribute value
            pFeat = pFeatClass.CreateFeature
            pFeat.Shape = pPolygon
            pFeat.Value(FCA_index_1) = migrated_distance(i)
            pFeat.Value(FCA_index_2) = i
            pFeat.Value(FCA_index_3) = m(i)
            pFeat.Value(FCA_index_4) = m_old(i)
            pFeat.Value(FCA_index_5) = mig1(i)
            pFeat.Value(FCA_index_6) = mig2(i)
            pFeat.Value(FCA_index_7) = mig3(i)
            pFeat.Value(FCA_index_8) = mig4(i)

            pFeat.Store()

        Next i
        pProgbar.Hide()
    End Sub

  
    Public Sub New()

    End Sub



    Protected Overrides Sub OnUpdate()
        Enabled = My.ArcMap.Application IsNot Nothing
    End Sub


End Class
Module utilities_for_migration
    Sub CleanGeometry(ByRef pGeom As IGeometry)
        ' This ensures that a geometry does not contain duplicate points due to
        ' the presence of several shapes making up the geometry
        Dim pPc As IPointCollection
        Dim pPc_new As IPointCollection
        Dim pGc As IGeometryCollection
        Dim iCount As Long
        Dim jCount As Long
        Dim i As Long
        Dim j As Long

        pGc = pGeom
        pPc_new = pGc.Geometry(0)

        iCount = pGc.GeometryCount - 1
        For i = 1 To iCount
            pPc = pGc.Geometry(i)
            jCount = pPc.PointCount - 1
            For j = 1 To jCount
                pPc_new.AddPoint(pPc.Point(j))
            Next j
        Next i
        pGc = New Polyline
        pGc.AddGeometry(pPc_new)
        pGeom = pGc
    End Sub
    Public Sub SetupFields(ByRef pFields As IFields, ByVal spatialreference As ISpatialReference3) '(StrFolder As String, StrName As String)

        Const strShapeFieldName As String = "Shape"
        ' Do not include .shp extension in strName

        ' Set up a simple fields collection
        Dim pFieldsEdit As IFieldsEdit
        pFields = New ESRI.ArcGIS.Geodatabase.Fields
        pFieldsEdit = pFields

        Dim pField As IField
        Dim pFieldEdit As IFieldEdit

        ' Make the shape field
        ' it will need a geometry definition, with a spatial reference
        pField = New ESRI.ArcGIS.Geodatabase.Field
        pFieldEdit = pField
        pFieldEdit.Name_2 = strShapeFieldName
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeGeometry

        Dim pGeomDef As IGeometryDef
        Dim pGeomDefEdit As IGeometryDefEdit
        pGeomDef = New GeometryDef
        pGeomDefEdit = pGeomDef
        With pGeomDefEdit
            .GeometryType_2 = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon
            .SpatialReference_2 = spatialreference
        End With
        pFieldEdit.GeometryDef_2 = pGeomDef
        pFieldsEdit.AddField(pField)

        ' Add another double field
        pField = New ESRI.ArcGIS.Geodatabase.Field
        pFieldEdit = pField
        With pFieldEdit
            .Length_2 = 10
            .Name_2 = "Mig_dist"
            .Type_2 = esriFieldType.esriFieldTypeDouble
            .Precision_2 = 10
            .Scale_2 = 3
        End With
        pFieldsEdit.AddField(pField)

        ' Add another double field
        pField = New ESRI.ArcGIS.Geodatabase.Field
        pFieldEdit = pField
        With pFieldEdit
            .Length_2 = 10
            .Name_2 = "i"
            .Type_2 = esriFieldType.esriFieldTypeDouble
            .Precision_2 = 10
            .Scale_2 = 2
        End With
        pFieldsEdit.AddField(pField)

        ' Add another double field
        pField = New ESRI.ArcGIS.Geodatabase.Field
        pFieldEdit = pField
        With pFieldEdit
            .Length_2 = 10
            .Name_2 = "m"
            .Type_2 = esriFieldType.esriFieldTypeDouble
            .Precision_2 = 10
            .Scale_2 = 3
        End With
        pFieldsEdit.AddField(pField)


        ' Add another double field
        pField = New ESRI.ArcGIS.Geodatabase.Field
        pFieldEdit = pField
        With pFieldEdit
            .Length_2 = 10
            .Name_2 = "old_m"
            .Type_2 = esriFieldType.esriFieldTypeDouble
            .Precision_2 = 10
            .Scale_2 = 3
        End With
        pFieldsEdit.AddField(pField)

        ' Add another double field
        pField = New ESRI.ArcGIS.Geodatabase.Field
        pFieldEdit = pField
        With pFieldEdit
            .Length_2 = 10
            .Name_2 = "Mig_1"
            .Type_2 = esriFieldType.esriFieldTypeDouble
            .Precision_2 = 10
            .Scale_2 = 3
        End With
        pFieldsEdit.AddField(pField)

        ' Add another double field
        pField = New ESRI.ArcGIS.Geodatabase.Field
        pFieldEdit = pField
        With pFieldEdit
            .Length_2 = 10
            .Name_2 = "Mig_2"
            .Type_2 = esriFieldType.esriFieldTypeDouble
            .Precision_2 = 10
            .Scale_2 = 3
        End With
        pFieldsEdit.AddField(pField)

        ' Add another double field
        pField = New ESRI.ArcGIS.Geodatabase.Field
        pFieldEdit = pField
        With pFieldEdit
            .Length_2 = 10
            .Name_2 = "Mig_3"
            .Type_2 = esriFieldType.esriFieldTypeDouble
            .Precision_2 = 10
            .Scale_2 = 3
        End With
        pFieldsEdit.AddField(pField)

        ' Add another double field
        pField = New ESRI.ArcGIS.Geodatabase.Field
        pFieldEdit = pField
        With pFieldEdit
            .Length_2 = 10
            .Name_2 = "Mig_4"
            .Type_2 = esriFieldType.esriFieldTypeDouble
            .Precision_2 = 10
            .Scale_2 = 3
        End With
        pFieldsEdit.AddField(pField)

        ' Add another double field
        pField = New ESRI.ArcGIS.Geodatabase.Field
        pFieldEdit = pField
        With pFieldEdit
            .Length_2 = 10
            .Name_2 = "LB_len"
            .Type_2 = esriFieldType.esriFieldTypeDouble
            .Precision_2 = 10
            .Scale_2 = 3
        End With
        pFieldsEdit.AddField(pField)

        ' Add another double field
        pField = New ESRI.ArcGIS.Geodatabase.Field
        pFieldEdit = pField
        With pFieldEdit
            .Length_2 = 10
            .Name_2 = "RB_len"
            .Type_2 = esriFieldType.esriFieldTypeDouble
            .Precision_2 = 10
            .Scale_2 = 3
        End With
        pFieldsEdit.AddField(pField)

        ' Add another double field
        pField = New ESRI.ArcGIS.Geodatabase.Field
        pFieldEdit = pField
        With pFieldEdit
            .Length_2 = 10
            .Name_2 = "LB_buf_ln"
            .Type_2 = esriFieldType.esriFieldTypeDouble
            .Precision_2 = 10
            .Scale_2 = 3
        End With
        pFieldsEdit.AddField(pField)

        ' Add another double field
        pField = New ESRI.ArcGIS.Geodatabase.Field
        pFieldEdit = pField
        With pFieldEdit
            .Length_2 = 10
            .Name_2 = "RB_buf_ln"
            .Type_2 = esriFieldType.esriFieldTypeDouble
            .Precision_2 = 10
            .Scale_2 = 3
        End With
        pFieldsEdit.AddField(pField)
        ' Create the shapefile
        ' (some parameters apply to geodata base options and can be defaulted as Nothing)


    End Sub
End Module