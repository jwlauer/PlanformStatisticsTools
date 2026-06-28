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


Class BankPolys
    Inherits ESRI.ArcGIS.Desktop.AddIns.Tool
    Dim pPcLeft As IPointCollection
    Dim pPcRight As IPointCollection
    Dim pPcCL As IPointCollection
    Dim pPcLeftPhoto As IPointCollection
    Dim pPcRightPhoto As IPointCollection
    Dim steps As Integer
    Dim BufferWidth As Double
    Dim DefaultSpatialReference As ISpatialReference3

    Protected Overrides Sub OnActivate()
        MyBase.OnActivate()
        MsgBox("Select the center line")
        steps = 0

    End Sub

    Protected Overrides Sub OnMouseDown(ByVal arg As ESRI.ArcGIS.Desktop.AddIns.Tool.MouseEventArgs)
        MyBase.OnMouseDown(arg)
        'Written by Wes Lauer
        'University of Minnesota
        'Saint Anthony Falls Laboratory
        '2 3rd Avenue, SE, Minneapolis, MN 55414
        'Updated August 4, 2011

        Dim pMxApp As IMxApplication
        Dim pMxDoc As IMxDocument
        Dim pMap As IMap
        Dim pActiveView As IActiveView
        Dim pEnvelope As IEnvelope
        Dim pEnumFeat As IEnumFeature
        Dim pGeom As IGeometry
        'Dim xy() As Double  'size of array xy will be set by the size of the line object
        'Dim xy_centerline() As Double
        'Dim i As Long
        'Dim R0 As Double
        'Dim pOutputCenterline As IPolyline
        'Dim width() As Double
        'Dim m() As Double

        pMxApp = My.ArcMap.Application
        pMxDoc = My.ArcMap.Application.Document
        pMap = pMxDoc.FocusMap
        pActiveView = pMap

        Select Case steps
            Case 0
                'GET THE COORDINATES FOR THE CENTERLINE
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
                pPcCL = New Polyline
                pGeom = pEnumFeat.Next.ShapeCopy
                DefaultSpatialReference = pGeom.SpatialReference

                pPcCL.AddPointCollection(pGeom)


                BufferWidth = InputBox("Input the buffer width")

                MsgBox("Select a right bank line to be used for bank length measurements")

                steps = steps + 1

            Case 1
                'GET THE COORDINATES FOR THE RIGHT BANK AS DEFINED BY PHOTO
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

                pPcRightPhoto = New Polyline
                pPcRightPhoto.AddPointCollection(pGeom)

                MsgBox("Select a left bank to be used for bank length measurements")
                steps = steps + 1
            Case 2
                'GET THE COORDINATES FOR THE LEFT BANK AS DEFINED BY PHOTO
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

                pPcLeftPhoto = New Polyline
                pPcLeftPhoto.AddPointCollection(pGeom)

                MsgBox("Select the right bank line to be used to create buffers")
                steps = steps + 1

            Case 3
                'GET THE COORDINATES FOR THE RIGHT BANK AS DEFINED BY LIDAR
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

                MsgBox("select the left bank line to be used to create buffers")
                steps = steps + 1

            Case 4
                'GET THE COORDINATES FOR THE LEFT BANK AS DEFINED BY LIDAR
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

                pPcLeft = New Polyline
                pPcLeft.AddPointCollection(pGeom)

                Call CreateBankBufferPolygons(pPcLeft, pPcCL, pPcRight, pPcLeftPhoto, pPcRightPhoto, pActiveView, BufferWidth)
                steps = 5
        End Select
        If steps = 5 Then My.ArcMap.Application.CurrentTool = Nothing
    End Sub


    Public Sub New()

    End Sub
    Sub CreateBankBufferPolygons(ByVal pPlLeft As IPolyline, ByVal pPlCl As IPolyline, ByVal pPlRight As IPolyline, ByVal pPlLeftPhoto As IPolyline, ByVal pPlRightPhoto As IPolyline, ByVal pActiveView As IActiveView, ByVal BufferWidth As Double)

        Dim pPlLeftOffset As IPolyline
        Dim pPlRightOffset As IPolyline
        Dim i As Long
        'Dim m() As Double
        Dim pi As Double
        'Dim pGxFile As IGxFile
        Dim pFields As IFields
        Dim path As String = Nothing
        Dim name As String = Nothing
        Dim dist As Double
        'Dim pLtangent2 As ILine
        dist = InputBox("Enter maximum distance that bank buffer is from channel")

        pi = 4 * Math.Atan(1)

        pPlRightOffset = ConstructOffset(pPlRight, BufferWidth)
        pPlLeftOffset = ConstructOffset(pPlLeft, -BufferWidth)

        'Const strShapeFieldName As String = "Shape"

        pFields = New Fields
        Call SetupFields(pFields, DefaultSpatialReference)

        '    Set m_pGxDialog = New GxDialog

        'create main polygon shapefile
        Dim pFeatClass As IFeatureClass
        pFeatClass = CreateNewShapefile(pFields, path, name)

        'create left bank polygons shapefile
        Dim pWorkspaceFactory As IWorkspaceFactory
        pWorkspaceFactory = New ShapefileWorkspaceFactory

        Dim pFeatureWorkspace As IFeatureWorkspace
        pFeatureWorkspace = pWorkspaceFactory.OpenFromFile(path, 0)

        Dim pLeftFeatureClass As IFeatureClass
        pLeftFeatureClass = pFeatureWorkspace.CreateFeatureClass(name & "_left", pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, "shape", "")

        'create right bank polygon shapefile
        Dim pRightFeatureClass As IFeatureClass
        pRightFeatureClass = pFeatureWorkspace.CreateFeatureClass(name & "_right", pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, "shape", "")

        Dim pLeftFeat As IFeature
        Dim pRightFeat As IFeature

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

        FCA_index_1 = pFeatClass.FindField("Mig_dist")
        FCA_index_2 = pFeatClass.FindField("i")
        FCA_index_3 = pFeatClass.FindField("old_m")
        FCA_index_4 = pFeatClass.FindField("LB_len")
        FCA_index_5 = pFeatClass.FindField("RB_len")
        FCA_index_6 = pFeatClass.FindField("LB_buf_ln")
        FCA_index_7 = pFeatClass.FindField("RB_buf_ln")

        Dim pGeomcoll As IPointCollection
        Dim pGeomcoll2 As IPointCollection
        'Dim pSegColl As ISegmentCollection
        Dim pLine As ILine
        Dim pPolygon As IPolygon
        Dim pPolygon2 As IPolygon
        Dim pFeat As IFeature
        'Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
        'Dim x3 As Double, y3 As Double, x4 As Double, y4 As Double
        Dim pPt1 As IPoint
        Dim pPt2 As IPoint
        Dim pPt3 As IPoint
        Dim pPt4 As IPoint
        Dim pClPtStart As IPoint
        Dim pClPtUpstream As IPoint
        Dim dummy_double1 As Double
        Dim dummy_double2 As Double
        Dim dummy_boolean As Boolean
        'Dim DistanceAlongCurve As Double

        'Dim theta1 As Double
        'Dim theta2 As Double
        'Dim dist As Double
        Dim pMultipoint As IPointCollection
        Dim pClPC As IPointCollection
        'Dim pTopologicalOperator As ITopologicalOperator
        Dim pBag As IGeometryCollection

        pBag = New GeometryBag

        pClPC = New Polyline

        '    dist = InputBox("enter the distance lateral to channel centerline to draw polygons")

        'Dim pCLine As ILine
        'Set pCLine = New Line
        'Dim pNLine As ILine
        'Set pNLine = New Line

        Dim pStatusBar As IStatusBar
        Dim pProgbar As IStepProgressor
        pStatusBar = My.ArcMap.Application.StatusBar
        pProgbar = pStatusBar.ProgressBar

        pProgbar.Position = 0
        pStatusBar.ShowProgressBar("Working...", 0, pClPC.PointCount - 2, 1, True)

        Dim pClpt1 As IPoint
        Dim pClpt2 As IPoint
        Dim pClpt3 As IPoint
        Dim pClpt4 As IPoint
        'Dim theta As Double
        'Dim theta_old As Double
        Dim dtheta As Double
        Dim pSegment As ILine
        Dim dx As Double
        Dim angle_threshold As Double
        Dim theta_upstream As Double
        Dim theta_downstream As Double
        pSegment = New Line
        pClpt1 = New Point
        pClpt2 = New Point
        pClpt3 = New Point
        pClpt4 = New Point
        pPt1 = New Point
        pPt2 = New Point
        pPt3 = New Point
        pPt4 = New Point
        pLine = New Line
        Dim pLine2 As ILine
        Dim pPLine2 As ISegmentCollection
        'Dim pPLine3 As IPolyline
        Dim pTopoOptr As ITopologicalOperator

        ' Dim pGeomCollection As IGeometryCollection
        Dim pPlLeftSub As IPointCollection
        Dim pPlRightSub As IPointCollection
        Dim pPlLeftBankSub As IPointCollection
        Dim pPlRightBankSub As IPointCollection
        Dim pPlLeftBankSubPl As IPolyline
        Dim pPlRightBankSubPl As IPolyline
        Dim StartBankDistance As Double
        Dim EndBankDistance As Double
        Dim theta_local_segment As Double

        pLine2 = New Line
        pPlLeftSub = New Polyline
        pPlRightSub = New Polyline
        pClPC = pPlCl
        pClPtStart = New Point

        ' define starting points

        pSegment.PutCoords(pClPC.Point(0), pClPC.Point(1))
        pSegment.QueryPoint(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, 0.5, True, pClPtStart)
        dx = pSegment.Length
        pPlCl.QueryNormal(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, dx / 2, False, -dist, pLine)

        pPLine2 = New Polyline
        pPLine2.AddSegment(pLine)
        pTopoOptr = pPLine2
        pGeomcoll = pTopoOptr.Intersect(pPlLeftOffset, ESRI.ArcGIS.Geometry.esriGeometryDimension.esriGeometry0Dimension)
        pPt3 = pGeomcoll.Point(0)


        pPlCl.QueryNormal(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, dx / 2, False, dist, pLine)

        pPLine2 = New Polyline
        pPLine2.AddSegment(pLine)
        pTopoOptr = pPLine2
        pGeomcoll = pTopoOptr.Intersect(pPlRightOffset, ESRI.ArcGIS.Geometry.esriGeometryDimension.esriGeometry0Dimension)
        pPt4 = pGeomcoll.Point(0)


        angle_threshold = InputBox("Enter threshold angle (°) above which normal line is drawn")


        'define remaining points

        For i = 1 To pClPC.PointCount - 3

            pStatusBar.StepProgressBar()
            pStatusBar.Message(0) = "i = " & i

            pMultipoint = New Multipoint
            pClpt1 = pClPC.Point(i - 1) '2 points upstream on centerline
            pClpt2 = pClPC.Point(i) '1 point upstream on centerline
            pClpt3 = pClPC.Point(i + 1) '1 point downstream centerline point
            pClpt4 = pClPC.Point(i + 2) '2 point downstream on centerline
            pClPtUpstream = pClPtStart
            pPt1 = pPt3 'pPt1 is upstream point on left
            pPt2 = pPt4 'pPt2 is upstream point on right
            pPt3 = New Point 'pPt3 is current point on left
            pPt4 = New Point 'pPt4 is current point on right
            pClPtStart = New Point

            pSegment.PutCoords(pClPC.Point(i - 1), pClPC.Point(i))
            theta_upstream = pSegment.Angle

            pLine2.PutCoords(pClPC.Point(i), pClPC.Point(i + 1))
            theta_local_segment = pSegment.Angle

            pSegment.PutCoords(pClPC.Point(i + 1), pClPC.Point(i + 2))
            theta_downstream = pSegment.Angle

            pLine2.QueryPoint(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, 0.5, True, pClPtStart)

            'Centered Difference approximation for dtheta
            dtheta = 0.5 * (theta_downstream - theta_upstream)
            If dtheta < -pi Then dtheta = dtheta + 2 * pi
            If dtheta > pi Then dtheta = dtheta - 2 * pi

            If dtheta > angle_threshold * pi / 180 Then
                'Channel is curving sharply to left, so extend normal line to right and look for near point on left
                pLine2.QueryNormal(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, 0.5, True, dist, pLine)
                pPLine2 = New Polyline
                pPLine2.AddSegment(pLine)

                'find the new intersection point

                'Call CreateDrawingElement(pLine, pActiveView)
                pTopoOptr = pPLine2
                pGeomcoll = pTopoOptr.Intersect(pPlRightOffset, ESRI.ArcGIS.Geometry.esriGeometryDimension.esriGeometry0Dimension)
                If pGeomcoll.PointCount > 0 Then
                    pPt4 = ClosestPt(pClPtStart, pGeomcoll)
                Else
                    pPt4 = pPt2
                End If
                'pLine.QueryToPoint pPt4

                pPlLeftOffset.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pClPtStart, False, pPt3, dummy_double1, dummy_double2, dummy_boolean)

            ElseIf dtheta < -angle_threshold * pi / 180 Then
                'channel is curving sharply to right, so extend normal line to left and look for near point on right
                pLine2.QueryNormal(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, 0.5, True, -dist, pLine)
                'Call CreateDrawingElement(pLine2, pActiveView)
                pPLine2 = New Polyline
                pPLine2.AddSegment(pLine)

                'find the new intersection point

                'Call CreateDrawingElement(pLine, pActiveView)
                pTopoOptr = pPLine2
                pGeomcoll = pTopoOptr.Intersect(pPlLeftOffset, ESRI.ArcGIS.Geometry.esriGeometryDimension.esriGeometry0Dimension)
                If pGeomcoll.PointCount > 0 Then
                    pPt3 = ClosestPt(pClPtStart, pGeomcoll)
                Else
                    pPt3 = pPt1
                End If
                pPlRightOffset.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pClPtStart, False, pPt4, dummy_double1, dummy_double2, dummy_boolean)

            Else
                'channel is not curving much.  Look for points along normal lines on both sides
                pLine2.QueryNormal(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, 0.5, True, dist, pLine)
                pPLine2 = New Polyline
                pPLine2.AddSegment(pLine)

                'find the new intersection point

                'Call CreateDrawingElement(pLine, pActiveView)
                pTopoOptr = pPLine2
                pGeomcoll = pTopoOptr.Intersect(pPlRightOffset, ESRI.ArcGIS.Geometry.esriGeometryDimension.esriGeometry0Dimension)
                If pGeomcoll.PointCount > 0 Then
                    pPt4 = ClosestPt(pClPtStart, pGeomcoll)
                Else
                    pPt4 = pPt2
                End If

                pLine2.QueryNormal(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, 0.5, True, -dist, pLine)
                pPLine2 = New Polyline
                pPLine2.AddSegment(pLine)

                'find the new intersection point

                'Call CreateDrawingElement(pLine, pActiveView)
                pTopoOptr = pPLine2
                pGeomcoll = pTopoOptr.Intersect(pPlLeftOffset, ESRI.ArcGIS.Geometry.esriGeometryDimension.esriGeometry0Dimension)
                If pGeomcoll.PointCount > 0 Then
                    pPt3 = ClosestPt(pClPtStart, pGeomcoll)
                Else
                    pPt3 = pPt1
                End If

            End If


            'Check that the new point on the offset bank is not upstream of the old point on the offset bank

            pPlRightOffset.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPt2, False, Nothing, StartBankDistance, dummy_double2, dummy_boolean)
            pPlRightOffset.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPt4, False, Nothing, EndBankDistance, dummy_double2, dummy_boolean)
            If StartBankDistance > EndBankDistance Then
                pPt4 = pPt2
            End If

            pPlLeftOffset.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPt1, False, Nothing, StartBankDistance, dummy_double2, dummy_boolean)
            pPlLeftOffset.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPt3, False, Nothing, EndBankDistance, dummy_double2, dummy_boolean)
            If StartBankDistance > EndBankDistance Then
                pPt3 = pPt1
            End If

            'create point collections to go between the intersection points
            pPlLeftOffset.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPt1, True, Nothing, StartBankDistance, dummy_double2, dummy_boolean)
            pPlLeftOffset.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPt3, True, Nothing, EndBankDistance, dummy_double2, dummy_boolean)
            pPlLeftOffset.GetSubcurve(StartBankDistance, EndBankDistance, True, pPlLeftSub)

            pPlRightOffset.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPt4, True, Nothing, StartBankDistance, dummy_double2, dummy_boolean)
            pPlRightOffset.QueryPointAndDistance(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, pPt2, True, Nothing, EndBankDistance, dummy_double2, dummy_boolean)
            pPlRightOffset.GetSubcurve(StartBankDistance, EndBankDistance, True, pPlRightSub)

            pGeomcoll = New Polygon

            'add the point collection to the polygon
            pMultipoint.addpoint(pClPtUpstream)
            pMultipoint.addpoint(pPt1)
            pMultipoint.AddPointCollection(pPlLeftSub)
            pMultipoint.addpoint(pPt3)
            pMultipoint.addpoint(pClPtStart)
            pMultipoint.addpoint(pPt4)
            pMultipoint.AddPointCollection(pPlRightSub)
            pMultipoint.addpoint(pPt2)
            pMultipoint.addpoint(pClPtUpstream)

            pGeomcoll.AddPointCollection(pMultipoint)
            pPolygon = pGeomcoll
            pTopoOptr = pPolygon
            pTopoOptr.Simplify()
            dummy_boolean = pTopoOptr.IsSimple

            'MsgBox pPlRightBankSubPl.Length
            'MsgBox pPlLeftBankSubPl.Length

            'Create the new feature, set the Feature's Shape, and set the index value
            pFeat = pFeatClass.CreateFeature
            pFeat.Shape = pPolygon

            pFeat.Value(FCA_index_2) = i

            'Intersect the new polygon with the actual bank line (defined by aerial photo)

            pTopoOptr = pPlLeftPhoto
            pGeomcoll2 = pTopoOptr.Intersect(pPolygon, ESRI.ArcGIS.Geometry.esriGeometryDimension.esriGeometry1Dimension)
            pPlLeftBankSub = pGeomcoll2
            pPlLeftBankSubPl = pPlLeftBankSub

            pTopoOptr = pPlRightPhoto
            pGeomcoll2 = pTopoOptr.Intersect(pPolygon, ESRI.ArcGIS.Geometry.esriGeometryDimension.esriGeometry1Dimension)
            pPlRightBankSub = pGeomcoll2
            pPlRightBankSubPl = pPlRightBankSub

            pFeat.Value(FCA_index_4) = pPlLeftBankSubPl.Length
            pFeat.Value(FCA_index_5) = pPlRightBankSubPl.Length

            'Intersect the new polygon with the actual bank line (defined by lidar)

            pTopoOptr = pPlLeft
            pGeomcoll2 = pTopoOptr.Intersect(pPolygon, ESRI.ArcGIS.Geometry.esriGeometryDimension.esriGeometry1Dimension)
            pPlLeftBankSub = pGeomcoll2
            pPlLeftBankSubPl = pPlLeftBankSub

            pTopoOptr = pPlRight
            pGeomcoll2 = pTopoOptr.Intersect(pPolygon, ESRI.ArcGIS.Geometry.esriGeometryDimension.esriGeometry1Dimension)
            pPlRightBankSub = pGeomcoll2
            pPlRightBankSubPl = pPlRightBankSub

            pFeat.Value(FCA_index_6) = pPlLeftBankSubPl.Length
            pFeat.Value(FCA_index_7) = pPlRightBankSubPl.Length

            pFeat.Store()

            pLeftFeat = pLeftFeatureClass.CreateFeature
            pTopoOptr = pPolygon
            pPolygon2 = New Polygon
            pTopoOptr.Cut(pPlLeft, pPolygon2, Nothing)
            pLeftFeat.Shape = pPolygon2
            pLeftFeat.Value(FCA_index_2) = i
            pLeftFeat.Store()

            pRightFeat = pRightFeatureClass.CreateFeature
            pTopoOptr = pPolygon
            pPolygon2 = New Polygon
            pTopoOptr.Cut(pPlRight, Nothing, pPolygon2)
            pRightFeat.Shape = pPolygon2
            pRightFeat.Value(FCA_index_2) = i
            pRightFeat.Store()


        Next i
        pProgbar.Hide()


    End Sub

    Private Function ConstructOffset(ByVal pInPolyline As IPolyline, ByVal dOffset As Double) As IPolyline
        Dim pConstructCurve As IConstructCurve

        On Error GoTo ErrorHandler

        If pInPolyline Is Nothing Or pInPolyline.IsEmpty Then
            ConstructOffset = Nothing
            Exit Function
        End If

        pConstructCurve = New Polyline
        pConstructCurve.ConstructOffset(pInPolyline, dOffset, 8)

        ConstructOffset = pConstructCurve
        Exit Function

ErrorHandler:
        ConstructOffset = Nothing
    End Function

    Private Function ClosestPt(ByVal pInPoint As IPoint, ByVal pInPointCollection As IPointCollection) As IPoint
        Dim shortestlength As Double
        Dim bestindex As Integer
        Dim j As Integer
        Dim pTestline As ILine

        pTestline = New Line

        pTestline.PutCoords(pInPoint, pInPointCollection.Point(0))
        shortestlength = pTestline.Length
        bestindex = 0
        For j = 0 To pInPointCollection.PointCount - 1
            pTestline.PutCoords(pInPoint, pInPointCollection.Point(j))
            If pTestline.Length < shortestlength Then
                bestindex = j
                shortestlength = pTestline.Length
            End If
        Next j

        ClosestPt = pInPointCollection.Point(bestindex)

    End Function



    Protected Overrides Sub OnUpdate()
        Enabled = My.ArcMap.Application IsNot Nothing
    End Sub
End Class