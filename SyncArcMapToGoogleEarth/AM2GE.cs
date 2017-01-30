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
using System.IO;

namespace SyncArcMapToGoogleEarth
{
    public class AM2GE : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        #region Enums

        private enum enumLongLat
        {
            Latitude = 1,
            Longitude = 2
        };

        private enum enumReturnFormat
        {
            WithSigns = 0,
            NMEA = 1
        };

        #endregion

        #region Properties

        private String _gpsString = String.Empty;

        private String _arcMapLat = "39.715620";
        private String _arcMapLong = "-84.103466";
        private String _arcMapAlt = "510";
        private String _gpsAltitude = String.Empty;

        private StreamWriter _sw;
        static private String _trackingFileNameEnd = "_AM2GE.kml";
        static private String _trackingFileName = "TrackingFile" + _trackingFileNameEnd;
        static private String _networkFileName = "NetworkLink" + _trackingFileNameEnd;
        static private String _FilePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Arcmap2GoogleEarth\\";
        static private String _networkLinkFilePath = _FilePath + _networkFileName;
        static private String _trackingFilePath = _FilePath + _trackingFileName;
        private Double dblEGA = 1000;

        private IApplication _application;
        private IMxDocument _mxdocument;
        private IMap _map;

        #endregion

        #region Constructor(s)

        public AM2GE()
        {
        }

        #endregion

        #region Event Handler(s)

        private ESRI.ArcGIS.Carto.IActiveViewEvents_ViewRefreshedEventHandler _ActiveViewEventsViewRefreshed;

        #endregion

        #region Event(s)

        protected override void OnClick()
        {
            //
            //  TODO: Sample code showing how to access button host
            //
            ArcMap.Application.CurrentTool = null;
            initiliaze();

            if (this.Checked)
                initializeTracking();
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


            Double latXmin;
            Double latXmax;

            Double longYmin;
            Double longYmax;

            Double diagonal;
            Double fakeGPSaltitude;

            lowerLeftPoint.X = view.Extent.XMin;
            lowerLeftPoint.Y = view.Extent.YMin;

            upperRightPoint.X = view.Extent.XMax;
            upperRightPoint.Y = view.Extent.YMax;

            PointToLatLong(lowerLeftPoint, out latXmin, out longYmin);
            PointToLatLong(upperRightPoint, out latXmax, out longYmax);

            diagonal = distance(latXmin, longYmin, latXmax, longYmax, 'K') * 1000; // "1KM * 1000
            diagonal = Math.Round(diagonal, 2);

            fakeGPSaltitude = dblEGA + (0.5 * Math.Sqrt(3) * diagonal);
            fakeGPSaltitude = Math.Round(fakeGPSaltitude, 2);

            _arcMapAlt = Convert.ToString(0.5 * Math.Sqrt(3) * diagonal);
            _gpsAltitude = Convert.ToString(dblEGA + (0.5 * Math.Sqrt(3) * diagonal));

            point.X = (view.Extent.XMax + view.Extent.XMin) / 2;
            point.Y = (view.Extent.YMax + view.Extent.YMin) / 2;

            Double lat;
            Double lon;

            PointToLatLong(point, out lat, out lon);

            Double n = 1000;

            if (_map.MapScale > 0)
            {
                if (_map.MapScale > 5000)
                    n = Math.Round(_map.MapScale * 0.45);
                else if (_map.MapScale > 4000 && _map.MapScale <= 5000)
                    n = Math.Round(_map.MapScale * 0.5);
                else if (_map.MapScale > 3000 && _map.MapScale <= 5000)
                    n = Math.Round(_map.MapScale * 0.55);
                else if (_map.MapScale > 2000 && _map.MapScale <= 5000)
                    n = Math.Round(_map.MapScale * 0.6);
                else if (_map.MapScale > 1000 && _map.MapScale <= 5000)
                    n = Math.Round(_map.MapScale * 0.65);
                else
                {
                    Double factor = (1001 - _map.MapScale) / 1000;
                    factor = Math.Pow(factor, 2);
                    factor = 0.65 + factor;
                    if (factor > 1)
                    {
                        factor = Math.Pow(factor, 2);
                        if (factor > 1.7)
                            factor = factor * 1.5;
                    }
                    n = Math.Round(_map.MapScale * factor);
                }
            }

            String convLat = DecimalPosToDegrees(lat, enumLongLat.Latitude, enumReturnFormat.NMEA);
            String convLong = DecimalPosToDegrees(lon, enumLongLat.Longitude, enumReturnFormat.NMEA);

            lat = Math.Round(lat, 5);
            lon = Math.Round(lon, 5);

            _arcMapLat = Convert.ToString(lat);
            _arcMapLong = Convert.ToString(lon);

            String nowTime = DateTime.Now.ToString("HHmmss.ss");
            String nowDate = DateTime.Now.ToString("ddMMyy");

            CreateTrackingKML();
        }

        #endregion

        #region Method(s)

        private void initiliaze()
        {
            _application = this.Hook as IApplication;
            _mxdocument = (IMxDocument)_application.Document;
            _map = _mxdocument.FocusMap;
            ActiveViewEventTracking();
        }

        private void initializeTracking()
        {
            if (!System.IO.Directory.Exists(_FilePath))
                System.IO.Directory.CreateDirectory(_FilePath);

            _sw = File.CreateText(_networkLinkFilePath);

            _sw.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            _sw.WriteLine("<kml xmlns=\"http://www.opengis.net/kml/2.2\">");
            _sw.WriteLine("<Folder>");
            _sw.WriteLine("<open>1</open>");
            _sw.WriteLine("<name>ArcMap to Google Earth Sync</name>");
            _sw.WriteLine("<NetworkLink>");
            _sw.WriteLine("<name>Camera</name>");
            _sw.WriteLine("<flyToView>1</flyToView>");
            _sw.WriteLine("<Link>");
            _sw.WriteLine("<href>" + _trackingFileName + "</href>");
            _sw.WriteLine("<refreshMode>onInterval</refreshMode>");
            _sw.WriteLine("<refreshInterval>0.300000</refreshInterval>");
            _sw.WriteLine("</Link>");
            _sw.WriteLine("</NetworkLink>");
            _sw.WriteLine("</Folder>");
            _sw.WriteLine("</kml>");

            _sw.Close();

            System.Diagnostics.Process.Start(_networkLinkFilePath);
        }

        private void CreateTrackingKML()
        {
            if (!System.IO.Directory.Exists(_FilePath))
                System.IO.Directory.CreateDirectory(_FilePath);

            _sw = File.CreateText(_trackingFilePath);

            _sw.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            _sw.WriteLine("<kml xmlns=\"http://www.opengis.net/kml/2.2\">");
            _sw.WriteLine("<NetworkLinkControl>");
            _sw.WriteLine("<LookAt>");
            _sw.WriteLine("<longitude>" + _arcMapLong + "</longitude>");
            _sw.WriteLine("<latitude>" + _arcMapLat + "</latitude>");
            _sw.WriteLine("<altitudeMode>relativeToGround</altitudeMode>");
            //_sw.WriteLine("<altitude> + " + _gpsAltitude + "</altitude>");
            _sw.WriteLine("<heading>0</heading>");
            _sw.WriteLine("<tilt>0</tilt>");
            _sw.WriteLine("<range>" + _arcMapAlt + "</range>");
            _sw.WriteLine("</LookAt>");
            _sw.WriteLine("</NetworkLinkControl>");
            _sw.WriteLine("</kml>");

            _sw.Close();
        }

        private void PointToLatLong(IPoint Point, out double Latitude, out double Longitude)
        {
            Latitude = 39.759444;
            Longitude = -84.191667;

            try
            {
                SpatialReferenceEnvironment SpRFc = new SpatialReferenceEnvironment();

                IGeographicCoordinateSystem GCS = SpRFc.CreateGeographicCoordinateSystem((int)esriSRGeoCSType.esriSRGeoCS_WGS1984);
                ISpatialReference SpRefOutput = GCS;
                SpRefOutput.SetFalseOriginAndUnits(-180, -90, 1000000);

                IGeometry2 Geometry2 = (IGeometry2)Point;

                ISpatialReference SpRefInput = _map.SpatialReference;
                if (SpRefInput == null)
                    return;

                Geometry2.SpatialReference = SpRefInput;
                Geometry2.Project(SpRefOutput);

                IPoint newPoint = (Point)Geometry2;

                Latitude = newPoint.Y;
                Longitude = newPoint.X;

                return;
            }
            catch (Exception) { }
            return;
        }


        private double distance(double lat1, double lon1, double lat2, double lon2, char unit)
        {
            //'M' is statute miles
            //'K' is kilometers (default)
            //'N' is nautical miles  
            double theta = lon1 - lon2;
            double dist = Math.Sin(deg2rad(lat1)) * Math.Sin(deg2rad(lat2)) + Math.Cos(deg2rad(lat1)) * Math.Cos(deg2rad(lat2)) * Math.Cos(deg2rad(theta));
            dist = Math.Acos(dist);
            dist = rad2deg(dist);
            dist = dist * 60 * 1.1515;
            if (unit == 'K')
            {
                dist = dist * 1.609344;
            }
            else if (unit == 'N')
            {
                dist = dist * 0.8684;
            }
            return (dist);
        }

        private Double deg2rad(Double deg)
        {
            return (deg * Math.PI / 180.0);
        }

        private Double rad2deg(Double rad)
        {
            return (rad / Math.PI * 180.0);
        }

        private void ActiveViewEventTracking()
        {
            ESRI.ArcGIS.Carto.IActiveViewEvents_Event activeViewEvents = _map as ESRI.ArcGIS.Carto.IActiveViewEvents_Event;
            _ActiveViewEventsViewRefreshed = new ESRI.ArcGIS.Carto.IActiveViewEvents_ViewRefreshedEventHandler(OnActiveViewEventsViewRefreshed);
            if (!this.Checked)
            {
                this.Checked = true;
                activeViewEvents.ViewRefreshed += _ActiveViewEventsViewRefreshed;
            }
            else
            {
                this.Checked = false;
                activeViewEvents.ViewRefreshed -= _ActiveViewEventsViewRefreshed;
            }

        }

        private String DecimalPosToDegrees(double Decimalpos, enumLongLat Type, enumReturnFormat OutputFormat, int SecondResolution = 2)
        {
            Int32 Deg = 0;
            Double Min = 0, Sec = 0;
            String Dir = "";
            Double tmpPos = Decimalpos;
            if (tmpPos < 0)
                tmpPos = Decimalpos * -1;

            Deg = (int)Math.Floor(tmpPos);
            Min = (tmpPos - Deg) * 60;

            switch (Type)
            {
                case enumLongLat.Latitude:
                    if (Decimalpos < 0)
                        Dir = "S";
                    else
                        Dir = "N";
                    break;

                case enumLongLat.Longitude:
                    if (Decimalpos < 0)
                        Dir = "W";
                    else
                        Dir = "E";
                    break;
            }

            Min = Math.Round(Min, 5);

            if (Dir == "W" || Dir == "E")
                return AddZeros(Deg, 3) + AddZeros(Min, 2) + Sec + "," + Dir;
            else
                return AddZeros(Deg, 2) + AddZeros(Min, 2) + Sec + "," + Dir;

        }

        private String AddZeros(double Value, int Zeros)
        {
            if (Math.Floor(Value).ToString().Length < Zeros)
                return Value.ToString().PadLeft(Zeros, (char)'0');
            return Value.ToString();
        }

        #endregion
    }

}
