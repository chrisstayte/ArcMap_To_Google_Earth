//The MIT License (MIT)

//Copyright (c) 2016 Chris Stayte

//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:

//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.

//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.

using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Geometry;
using System;
using System.Globalization;
using System.IO;

namespace SyncArcMapToGoogleEarth
{
    public class AM2GE : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        #region Properties

        // Default Location Is Batman Building In Japan
        private string _latitude = "26.357896";
        private string _longitude = "127.783809";
        private string _altitude = "100";
        private string _flyToView = "1";


        // Lcoation of save fiel and names
        private static string _saveDirectory;
        private static string _currentViewFileName = "AM2GE_CurrentView.kml";
        private static string _networkLinkFileName = "AM2GE_NetworkLink.kml";

        // Application Specifics
        private IApplication _application;
        private IMxDocument _mxdocument;
        private IMap _map;

        #endregion

        #region Constructor(s)

        public AM2GE()
        {
            _saveDirectory = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"AM2GE\");
        }

        #endregion

        #region Event Handler(s)

        private ESRI.ArcGIS.Carto.IActiveViewEvents_ViewRefreshedEventHandler _activeViewEventsViewRefreshed;

        #endregion

        #region Event(s)

        protected override void OnClick()
        {
            //
            //  TODO: Sample code showing how to access button host
            //
            ArcMap.Application.CurrentTool = null;
            Initiliaze();

            if (this.Checked)
                WriteNetworkLink();
        }

        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }

        private void OnActiveViewEventsViewRefreshed(ESRI.ArcGIS.Carto.IActiveView view, ESRI.ArcGIS.Carto.esriViewDrawPhase phase, System.Object data, ESRI.ArcGIS.Geometry.IEnvelope envelope)
        {
            IPoint point = new Point();
            IPoint lowerLeftPoint = new Point();
            IPoint upperRightPoint = new Point();


            double latXmin, latXmax, longYmin, longYmax, diagonal;

            lowerLeftPoint.X = view.Extent.XMin;
            lowerLeftPoint.Y = view.Extent.YMin;

            upperRightPoint.X = view.Extent.XMax;
            upperRightPoint.Y = view.Extent.YMax;

            PointToLatLong(lowerLeftPoint, out latXmin, out longYmin);
            PointToLatLong(upperRightPoint, out latXmax, out longYmax);

            diagonal = Distance(latXmin, longYmin, latXmax, longYmax, 'K') * 1000; // "1KM * 1000
            diagonal = Math.Round(diagonal, 2);

            _altitude = Convert.ToString(0.5 * Math.Sqrt(3) * diagonal, CultureInfo.InvariantCulture);

            point.X = (view.Extent.XMax + view.Extent.XMin) / 2;
            point.Y = (view.Extent.YMax + view.Extent.YMin) / 2;

            Double lat;
            Double lon;

            PointToLatLong(point, out lat, out lon);

            lat = Math.Round(lat, 5);
            lon = Math.Round(lon, 5);

            _latitude = Convert.ToString(lat, CultureInfo.InvariantCulture);
            _longitude = Convert.ToString(lon, CultureInfo.InvariantCulture);

            CreateTrackingKML();
        }

        #endregion

        #region Method(s)

        private void Initiliaze()
        {
            _application = this.Hook as IApplication;
            if (_application != null) _mxdocument = (IMxDocument)_application.Document;
            _map = _mxdocument.FocusMap;
            ActiveViewEventTracking();
        }

        private void WriteNetworkLink()
        {

            if (!System.IO.Directory.Exists(_saveDirectory))
            {
                System.IO.Directory.CreateDirectory(_saveDirectory);
            }

            var networkLinkFile = System.IO.Path.Combine(_saveDirectory, _networkLinkFileName);

            using (TextWriter tw = new StreamWriter(networkLinkFile))
            {
                tw.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                tw.WriteLine("<kml xmlns=\"http://www.opengis.net/kml/2.2\">");
                tw.WriteLine("<Folder>");
                tw.WriteLine("<open>1</open>");
                tw.WriteLine("<name>ArcMap to Google Earth Sync</name>");
                tw.WriteLine("<NetworkLink>");
                tw.WriteLine("<name>Camera</name>");
                tw.WriteLine("<flyToView>" + _flyToView + "</flyToView>");
                tw.WriteLine("<Link>");
                tw.WriteLine("<href>" + _currentViewFileName + "</href>");
                tw.WriteLine("<refreshMode>onInterval</refreshMode>");
                tw.WriteLine("<refreshInterval>0.500000</refreshInterval>");
                tw.WriteLine("</Link>");
                tw.WriteLine("</NetworkLink>");
                tw.WriteLine("</Folder>");
                tw.WriteLine("</kml>");
            }

            try
            {
                System.Diagnostics.Process.Start(networkLinkFile);
            } catch (Exception ex)
            {
                ex.ToString();
            }
        }

        private void CreateTrackingKML()
        {
            if (!System.IO.Directory.Exists(_saveDirectory))
            {
                System.IO.Directory.CreateDirectory(_saveDirectory);
            }

            var currentviewfile = System.IO.Path.Combine(_saveDirectory, _currentViewFileName);

            using (TextWriter tw = new StreamWriter(currentviewfile))
            {
                tw.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                tw.WriteLine("<kml xmlns=\"http://www.opengis.net/kml/2.2\">");
                tw.WriteLine("<NetworkLinkControl>");
                tw.WriteLine("<LookAt>");
                tw.WriteLine("<longitude>" + _longitude + "</longitude>");
                tw.WriteLine("<latitude>" + _latitude + "</latitude>");
                tw.WriteLine("<altitudeMode>relativeToGround</altitudeMode>");
                //tw.WriteLine("<altitude> + " + _gpsAltitude + "</altitude>");
                tw.WriteLine("<heading>0</heading>");
                tw.WriteLine("<tilt>0</tilt>");
                tw.WriteLine("<range>" + _altitude + "</range>");
                tw.WriteLine("</LookAt>");
                tw.WriteLine("</NetworkLinkControl>");
                tw.WriteLine("</kml>");
            }
        }

        private void PointToLatLong(IPoint Point, out double Latitude, out double Longitude)
        {
            Latitude = 39.759444;
            Longitude = -84.191667;

            try
            {
                var spatialReferenceEnvironment = new SpatialReferenceEnvironment();

                var geographicCoordinateSystem = spatialReferenceEnvironment.CreateGeographicCoordinateSystem((int)esriSRGeoCSType.esriSRGeoCS_WGS1984);
                var spatialReferenceOutput = geographicCoordinateSystem;
                spatialReferenceOutput.SetFalseOriginAndUnits(-180, -90, 1000000);

                var geometry2 = (IGeometry2)Point;

                var spatialReference = _map.SpatialReference;
                if (spatialReference == null)
                    return; 

                geometry2.SpatialReference = spatialReference;
                geometry2.Project(spatialReferenceOutput);

                IPoint newPoint = (Point)geometry2;

                Latitude = newPoint.Y;
                Longitude = newPoint.X;

                return;
            }
            catch (Exception)
            {
                // ignored
            }
            return;
        }


        private double Distance(double lat1, double lon1, double lat2, double lon2, char unit)
        {
            //'M' is statute miles
            //'K' is kilometers (default)
            //'N' is nautical miles  
            var theta = lon1 - lon2;
            var dist = Math.Sin(Degrees2Radians(lat1)) * Math.Sin(Degrees2Radians(lat2)) + Math.Cos(Degrees2Radians(lat1)) * Math.Cos(Degrees2Radians(lat2)) * Math.Cos(Degrees2Radians(theta));
            dist = Math.Acos(dist);
            dist = Radians2Degrees(dist);
            dist = dist * 60 * 1.1515;
            switch (unit)
            {
                case 'K':
                    dist = dist * 1.609344;
                    break;
                case 'N':
                    dist = dist * 0.8684;
                    break;
            }
            return (dist);
        }

        private static double Degrees2Radians(Double deg)
        {
            return (deg * Math.PI / 180.0);
        }

        private static double Radians2Degrees(Double rad)
        {
            return (rad / Math.PI * 180.0);
        }

        private void ActiveViewEventTracking()
        {
            var activeViewEvents = _map as IActiveViewEvents_Event;
            _activeViewEventsViewRefreshed = OnActiveViewEventsViewRefreshed;
            if (!this.Checked)
            {
                this.Checked = true;
                _flyToView = "1";
                if (activeViewEvents != null) activeViewEvents.ViewRefreshed += _activeViewEventsViewRefreshed;
                WriteNetworkLink();
            }
            else
            {
                this.Checked = false;
                _flyToView = "0";
                if (activeViewEvents != null) activeViewEvents.ViewRefreshed -= _activeViewEventsViewRefreshed;
                WriteNetworkLink();
            }

        }
        #endregion
    }

}
