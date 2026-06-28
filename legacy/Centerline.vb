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




Class Tool1
    Inherits ESRI.ArcGIS.Desktop.AddIns.Tool

    Dim DefaultSpatialReference As ISpatialReference3 = Nothing
    'Private m_pGxDialog As IGxDialog

    Dim pPcLeft As IPointCollection
    Dim pPcRight As IPointCollection
    Dim pPcCL As IPointCollection
    Dim pMxApp As IMxApplication
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pActiveView As IActiveView
    Dim pEnvelope As IEnvelope
    Dim pEnumFeat As IEnumFeature
    Dim pGeom As IGeometry

    Dim xy() As Double  'size of array xy will be set by the size of the line object
    Dim xy_centerline() As Double
    Dim i As Long
    Dim steps As Integer
    Dim R0 As Double
    Dim pOutputCenterline As IPolyline
    Dim width() As Double
    Dim m() As Double
    Dim theta() As Double
    Dim dtheta() As Double
    Dim r_curve() As Double


    Protected Overrides Sub OnActivate()
        MyBase.OnActivate()
        MsgBox("select left bank line")
        steps = 0

    End Sub

    Protected Overrides Sub OnMouseDown(ByVal arg As ESRI.ArcGIS.Desktop.AddIns.Tool.MouseEventArgs)
        MyBase.OnMouseDown(arg)
        'MsgBox(steps)
        'steps = steps + 1
        'Written by Wes Lauer
        'University of Minnesota
        'Saint Anthony Falls Laboratory
        '2 3rd Avenue, SE, Minneapolis, MN 55414

        'April 6, 2004
        'Updated for VB.NET July 28, 2011

        'This tool creates a line shapefile that represents the center of two roughly parallel lines.

        pMxApp = My.ArcMap.Application
        pMxDoc = My.ArcMap.Application.Document
        pMap = pMxDoc.FocusMap
        pActiveView = pMap
        Select Case steps
            Case 0
                'GET THE COORDINATES FOR THE LEFT BANK
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

                'pEnumFeat.Reset()
                pPcLeft = New Polyline
                pGeom = pEnumFeat.Next.ShapeCopy
                'pGeom.SpatialReference = pEnumFeat.Next.ShapeCopy.SpatialReference

                Dim pTopologicalOperator As ITopologicalOperator = pGeom

                'MsgBox(pTopologicalOperator.IsSimple)

                pPcLeft.AddPointCollection(pGeom)
                MsgBox("Select the right bank line")
                DefaultSpatialReference = pGeom.SpatialReference

                steps = 1

            Case 1
                'GET THE COORDINATES FOR THE RIGHT BANK

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
                pGeom = pEnumFeat.Next.ShapeCopy

                pPcRight = New Polyline
                pPcRight.AddPointCollection(pGeom)

                pPcCL = findCL(pPcLeft, pPcRight, pActiveView, width, m, theta, dtheta, r_curve, DefaultSpatialReference)

                'FIND THE POINTS ON THE LEFT AND RIGHT BANK THAT CORRESPOND WITH EACH POINT ON THE CENTERLINE
                Dim pPcRightFound As IPointCollection
                pPcRightFound = New Polyline
                Dim pPcLeftFound As IPointCollection
                pPcLeftFound = New Polyline
                Dim pLeftPline As IPolyline
                Dim pLeftPoint As IPoint
                pLeftPoint = New Point
                Dim pRightPline As IPolyline
                Dim pRightPoint As IPoint
                pRightPoint = New Point
                Dim DummyBoolean As Boolean
                Dim DummyBoolean2 As Boolean
                Dim DummyDouble As Double
                Dim DummyDouble2 As Double
                Dim i As Integer

                pLeftPline = pPcLeft
                pRightPline = pPcRight

                For i = 0 To pPcCL.PointCount - 1
                    pLeftPline.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPcCL.Point(i), DummyBoolean, pLeftPoint, DummyDouble, DummyDouble2, DummyBoolean2)
                    pPcLeftFound.AddPoint(pLeftPoint)
                    pRightPline.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPcCL.Point(i), DummyBoolean, pRightPoint, DummyDouble, DummyDouble2, DummyBoolean2)
                    pPcRightFound.AddPoint(pRightPoint)
                Next i

                'EXPORT THE CENTERLINE AS A NEW POLYLINE M SHAPEFILE
                Dim path As String = Nothing
                Dim name As String = Nothing

                Call CreatePolylineMShapefile(pPcCL, path, name, DefaultSpatialReference)
                '
                'EXPORT THE WIDTH DATA AS A TEXTFILE

                'Call ExportWidthText(m, width, theta, dtheta, r_curve, path, name)
                Call ExportWidthText(m, width, theta, dtheta, r_curve, path, name, pPcCL, pPcLeftFound, pPcRightFound)

                steps = 2

        End Select
        ' End Sub
        If steps = 2 Then My.ArcMap.Application.CurrentTool = Nothing
    End Sub


    Public Sub New()

    End Sub



    Protected Overrides Sub OnUpdate()
        Enabled = My.ArcMap.Application IsNot Nothing
    End Sub
End Class

Module program_utilities
    Public Function findCL(ByVal pLeft As IPolyline, ByVal pRight As IPolyline, ByVal pActiveView As IActiveView, ByRef width() As Double, ByRef m() As Double, ByRef theta() As Double, ByRef dtheta() As Double, ByRef r_curve() As Double, ByVal spatialreference As ISpatialReference3) As IPolyline
        Const pi = 3.141592654
        Dim i As Long
        Dim j As Long
        Dim k As Long
        Dim max_fraction As Double
        Dim min_fraction As Double
        Dim best_error As Double
        Dim theta_local As Double
        Dim dist As Double
        Dim increment As Double
        Dim fraction As Double
        Dim dist_along_left_bank As Double
        Dim dist_from_left_bank As Double
        Dim right_of_left_bank As Boolean
        Dim dist_along_right_bank As Double
        Dim dist_from_right_bank As Double
        Dim DistanceFromCurve As Double
        Dim bRightside As Boolean
        Dim right_of_right_bank As Boolean
        Dim dist_error As Double
        Dim best_fraction As Double
        Dim found As Boolean
        Dim pPt As IPoint
        Dim pLeftArcInt As IPoint
        Dim pRightArcInt As IPoint
        Dim pLeftPt As IPoint
        Dim pRightPt As IPoint
        Dim pPc As IPointCollection
        Dim pCircularArc As ICircularArc
        Dim pScreenDisplay As IScreenDisplay
        Dim pSymbol As ISimpleMarkerSymbol
        Dim pLine As ILine
        Dim pStatusBar As IStatusBar
        Dim pProgbar As IStepProgressor
        Dim approximate_final_index As Long
        Dim max_pts As Long
        Dim ResultingWidth() As Double
        Dim ResultingM() As Double
        Dim ResultingTheta() As Double
        Dim ResultingDtheta() As Double
        Dim ResultingRCurve() As Double
        Dim pSegColl As ISegmentCollection

        pStatusBar = My.ArcMap.Application.StatusBar
        pProgbar = pStatusBar.ProgressBar

        pPt = New Point
        pLeftPt = New Point
        pRightPt = New Point
        pPc = New Polyline
        pScreenDisplay = pActiveView.ScreenDisplay
        pSymbol = New SimpleMarkerSymbol

        pSymbol.Size = 4

        pLine = New Line
        pLine.PutCoords(pRight.FromPoint, pLeft.FromPoint)
        pLine.QueryPoint(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, 0.5, True, pPt)

        theta_local = pLine.Angle - pi
        ReDim ResultingTheta(0)
        ResultingTheta(0) = theta_local

        pPc.AddPoint(pPt)

        dist = InputBox("Enter distance between points: ")
        max_pts = InputBox("Enter maximum number of points to find: ")

        i = 0

        approximate_final_index = Int((pLeft.Length / 2 + pRight.Length / 2) / dist)

        pProgbar.Position = 0
        pStatusBar.ShowProgressBar("Working...", 0, approximate_final_index, 1, True)

        Do
            pStatusBar.StepProgressBar()
            pStatusBar.Message(0) = "i = " & i
            i = i + 1
            pCircularArc = New CircularArc
            pCircularArc.PutCoordsByAngle(pPt, theta_local, pi, dist)
            'Call CreateDrawingElement(pCircularArc, pActiveView)
            'If i = 1550 Then Stop
            'Set the maximum and minimum fractional distance along the curve using the points
            'where the curve intersects the bank lines (if any).  If there are no intersections,
            'the center point could be anywhere on the curve.  If there is one or more intersections
            'with either bank line, then use the intersection that occurs closest along the bank to
            'the last centerline point.

            pSegColl = New Polyline
            pSegColl.AddSegment(pCircularArc)

            pLeftArcInt = GetNearestIntersectionPoint(pSegColl, pLeft, pPt)
            pRightArcInt = GetNearestIntersectionPoint(pSegColl, pRight, pPt)

            'With pScreenDisplay
            '    .StartDrawing pScreenDisplay.hDC, esriNoScreenCache
            '    .SetSymbol pSymbol
            '    .DrawPoint pLeftArcInt
            '    .DrawPoint pRightArcInt
            '    .FinishDrawing
            'End With

            If pLeftArcInt Is Nothing Then
                max_fraction = 1
            Else
                pCircularArc.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pLeftArcInt, True, Nothing, max_fraction, DistanceFromCurve, bRightside)
            End If

            If pRightArcInt Is Nothing Then
                min_fraction = 0
            Else
                pCircularArc.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pRightArcInt, True, Nothing, min_fraction, DistanceFromCurve, bRightside)
            End If

            best_error = 10000
            For k = 1 To 5
                increment = (max_fraction - min_fraction) / 4
                For j = 0 To 4
                    fraction = min_fraction + increment * j
                    pCircularArc.QueryPoint(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, fraction, True, pPt)
                    pLeft.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPt, False, pLeftPt, dist_along_left_bank, _
                        dist_from_left_bank, right_of_left_bank)
                    pRight.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPt, False, pRightPt, dist_along_right_bank, _
                        dist_from_right_bank, right_of_right_bank)

                    dist_error = dist_from_left_bank - dist_from_right_bank

                    If Math.Abs(dist_error) < Math.Abs(best_error) And right_of_left_bank And Not right_of_right_bank Then
                        best_fraction = fraction
                        best_error = dist_error
                    End If
                Next j
                min_fraction = best_fraction - increment
                max_fraction = best_fraction + increment
            Next k

            pCircularArc.QueryPoint(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, best_fraction, True, pPt)
            pLeft.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPt, True, Nothing, dist_along_left_bank, _
                dist_from_left_bank, right_of_left_bank)
            pRight.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPt, True, Nothing, dist_along_right_bank, _
                dist_from_right_bank, right_of_right_bank)

            If dist_along_left_bank = 1 Or dist_along_right_bank = 1 Or right_of_right_bank Or Not right_of_left_bank Then
                found = False
                pPc.AddPoint(pPt)
                With pScreenDisplay
                    .StartDrawing(pScreenDisplay.hDC, -1)
                    .SetSymbol(pSymbol)
                    .DrawPoint(pPt)
                    .FinishDrawing()
                End With
                pLine.PutCoords(pPc.Point(i - 1), pPc.Point(i))
                theta_local = pLine.Angle - pi / 2
                'Call CreateDrawingElement(pLine, pActiveView)
            Else
                found = True
                pPc.AddPoint(pPt)
                If i Mod 10 = 0 Then
                    With pScreenDisplay
                        .StartDrawing(pScreenDisplay.hDC, -1)
                        .SetSymbol(pSymbol)
                        .DrawPoint(pPt)
                        .FinishDrawing()
                    End With
                End If
                pLine.PutCoords(pPc.Point(i - 1), pPc.Point(i))
                theta_local = pLine.Angle - pi / 2
                'Call CreateDrawingElement(pLine, pActiveView)
            End If

            ReDim Preserve ResultingWidth(i)
            ResultingWidth(i) = dist_from_right_bank + dist_from_left_bank
            ReDim Preserve ResultingM(i)
            ResultingM(i) = dist * i
            ReDim Preserve ResultingTheta(i)
            ResultingTheta(i - 1) = theta_local + pi / 2
            If ResultingTheta(i - 1) < 0 Then ResultingTheta(i - 1) = ResultingTheta(i - 1) + 2 * pi
            ReDim Preserve ResultingDtheta(i)
            If i > 1 Then
                ResultingDtheta(i - 1) = ResultingTheta(i - 1) - ResultingTheta(i - 2)

                If ResultingDtheta(i - 1) >= pi Then ResultingDtheta(i - 1) = ResultingDtheta(i - 1) - 2 * pi
                If ResultingDtheta(i - 1) <= -pi Then ResultingDtheta(i - 1) = ResultingDtheta(i - 1) + 2 * pi
            Else
                ResultingDtheta(i - 1) = -99999
            End If

            ReDim Preserve ResultingRCurve(i)
            If ResultingDtheta(i - 1) = 0 Then
                ResultingRCurve(i - 1) = -99999
            Else
                ResultingRCurve(i - 1) = dist / ResultingDtheta(i - 1)
            End If

        Loop Until found = False Or i = max_pts
        pStatusBar.HideProgressBar()

        ResultingTheta(i) = -99999
        ResultingDtheta(i) = -99999
        ResultingRCurve(i) = -99999


        width = ResultingWidth
        m = ResultingM
        theta = ResultingTheta
        dtheta = ResultingDtheta
        r_curve = ResultingRCurve

        findCL = pPc
        findCL.SpatialReference = spatialreference

    End Function
    Function GetNearestIntersectionPoint(ByVal pCirc As IPolyline, ByVal pBank As IPolyline, ByVal pStartPt As IPoint) As IPoint

        'Do the intersection

        'Intersect the two polylines creating a multipoint

        Dim pTopoOptr As ITopologicalOperator
        Dim pGeomcoll As IGeometryCollection
        Dim OldDistanceAlongBank As Double
        Dim DistanceAlongBank As Double
        Dim DistanceFromBank As Double
        Dim DistanceFromStart As Double
        Dim ShortestDistanceFromStart As Double
        Dim bRightside As Boolean
        Dim count As Integer
        Dim pPt As IPoint

        GetNearestIntersectionPoint = Nothing
        pTopoOptr = pCirc
        pGeomcoll = pTopoOptr.Intersect(pBank, ESRI.ArcGIS.Geometry.esriGeometryDimension.esriGeometry0Dimension)

        pBank.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pStartPt, True, Nothing, DistanceAlongBank, DistanceFromBank, bRightside)
        ShortestDistanceFromStart = 1 - OldDistanceAlongBank

        For count = 0 To pGeomcoll.GeometryCount - 1
            pPt = pGeomcoll.Geometry(count)
            pBank.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPt, True, Nothing, DistanceAlongBank, DistanceFromBank, bRightside)
            DistanceFromStart = DistanceAlongBank - OldDistanceAlongBank
            If DistanceFromStart < ShortestDistanceFromStart Then
                ShortestDistanceFromStart = DistanceFromStart
                GetNearestIntersectionPoint = pPt
            End If
        Next count

        'If no intersection points, exit
        If pGeomcoll.GeometryCount = 0 Then
            GetNearestIntersectionPoint = Nothing
            Exit Function
        End If

    End Function
    Sub CreatePolylineMShapefile(ByVal pPolyline As IPolyline, ByRef path As String, ByRef name As String, ByRef spatialreference As ISpatialReference3)
        ' Set up a simple fields collection
        Dim pFields As IFields
        pFields = New Fields
        Dim pFieldsEdit As IFieldsEdit
        pFields = New ESRI.ArcGIS.Geodatabase.Fields 'esriGeoDatabase.Fields
        pFieldsEdit = pFields

        Dim pField As IField
        Dim pFieldEdit As IFieldEdit

        ' Make the shape field
        ' it will need a geometry definition, with a spatial reference
        pField = New ESRI.ArcGIS.Geodatabase.Field
        pFieldEdit = pField
        'pFieldEdit.name = "Shape"
        pFieldEdit.Name_2 = "Shape"
        'pFieldEdit.Type = esriFieldTypeGeometry
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeGeometry

        Dim pGeomDef As IGeometryDef
        Dim pGeomDefEdit As IGeometryDefEdit
        pGeomDef = New GeometryDef
        pGeomDefEdit = pGeomDef
        With pGeomDefEdit
            .GeometryType_2 = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline
            .HasM_2 = True
            .SpatialReference_2 = spatialreference
        End With
        pFieldEdit.GeometryDef_2 = pGeomDef
        pFieldsEdit.AddField(pField)

        Dim pFeatClass As IFeatureClass
        'Dim pGxObject As IGxObject
        pFeatClass = CreateNewShapefile(pFields, path, name)
        Dim pFeat As IFeature
        pFeat = pFeatClass.CreateFeature
        Dim pMaware As IMAware
        pMaware = pPolyline
        pMaware.MAware = True

        Dim pMsegmentation As IMSegmentation
        pMsegmentation = pPolyline
        pMsegmentation.SetMsAsDistance(False)

        pFeat.Shape = pPolyline
        pFeat.Store()

    End Sub
    Function CreateNewShapefile(ByVal pFields As IFields, ByRef path As String, ByRef name As String) As IFeatureClass
        'Dim pClone As IClone
        Dim pFeatureWorkspace As IFeatureWorkspace
        Dim m_pGxDialog As IGxDialog

        'Dim pFields As IFields
        Dim pGxFile As IGxFile

        Dim pNewFeatureClass As IFeatureClass
        Dim pWorkspaceFactory As IWorkspaceFactory

        On Error GoTo ErrorHandler
        m_pGxDialog = New GxDialog
        m_pGxDialog.Title = "Enter New Output Shapefile:"
        If m_pGxDialog.DoModalSave(0) Then
            pGxFile = m_pGxDialog.FinalLocation
        Else
            CreateNewShapefile = Nothing
            Exit Function
        End If

        pWorkspaceFactory = New ShapefileWorkspaceFactory
        pFeatureWorkspace = pWorkspaceFactory.OpenFromFile(pGxFile.Path, 0)
        'Set pClone = pInFeatureClass.Fields
        'Set pFields = pClone.Clone

        pNewFeatureClass = pFeatureWorkspace.CreateFeatureClass(m_pGxDialog.Name, pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, "shape", "")

        CreateNewShapefile = pNewFeatureClass

        path = pGxFile.Path
        name = m_pGxDialog.Name

        Exit Function

ErrorHandler:
        CreateNewShapefile = Nothing

    End Function

    Sub ExportWidthText(ByVal m() As Double, ByVal width() As Double, ByVal theta() As Double, ByVal dtheta() As Double, ByVal r_curve() As Double, ByVal location As String, ByVal name As String, ByVal pPcCenter As IPointCollection, ByVal pPcLeft As IPointCollection, ByVal pPcRight As IPointCollection)

        Dim i As Integer
        Dim a, fs
        fs = CreateObject("Scripting.FileSystemObject")
        a = fs.CreateTextFile(location & "\" & name & ".txt", True)

        a.writeline("OID,m,width,theta,dtheta,r_curve,cl_x,cl_y,left_x,left_y,right_x,right_y")

        For i = 1 To UBound(m)
            a.writeline("0," & m(i) & "," & width(i) & "," & theta(i) & "," & dtheta(i) & "," & r_curve(i) & "," & pPcCenter.Point(i).X & "," & pPcCenter.Point(i).Y & "," & pPcLeft.Point(i).X & "," & pPcLeft.Point(i).Y & "," & pPcRight.Point(i).X & "," & pPcRight.Point(i).Y)
        Next i

        a.Close()
    End Sub

    End Module