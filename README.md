# PlanformStatisticsTools
Code for compiling the PlanformStatisticsToolbox, a custom ArcGIS add-in for ArcMap.

The code utilizes the ESRI ArcObjects libraries, which require an appropriate ESRI license. The code was compiled in Visual Studio 2008 with an installation of the ESRI ArcObjects SDK from around that era.  It has not been maintained since around 2012, so it is an open question what it will take to compile using a more recent version of ArcObjects.

There is code for three tools in the repository.  The first, centerline.vb, is for interpolating a centerline at evenly spaced intervals from a set of two bank lines.  The second, ArcGISAddin1.vb (sorry--that's a default name from Visual Studio), is the tool for estimating migration distances between two centerlines.  The third, BankPolys.vb, is for creating polygons along banks that correspond to a given centerline point.
