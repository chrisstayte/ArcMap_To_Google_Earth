using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ESRI.ArcGIS.ADF.CATIDs;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Geometry;
using System.Globalization;

namespace AM2GE
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

        Boolean StartedUP = false;
        private SerialPort _serialPort = new SerialPort();
        private String NL = System.Environment.NewLine;
        private String _gpsPort = String.Empty;
        private Boolean _gpsRunning = false;
        private String _gpsString = String.Empty;
        private String _gpsGLLString = String.Empty;
        private String _gpsGGAString = String.Empty;
        private String _gpsRMCString = String.Empty;
        private String _gpsVTGString = String.Empty;

        private String _arcMapLat = "39.715620";
        private String _arcMapLong = "-84.103466";
        private String _arcMapAlt = "510";
        private String _gpsAltitude = String.Empty;

        private Boolean _portOpen = false;
        private Boolean _programRunning = true;
        private String[] _Ports = new String[] { };
        private Boolean _noPorts = false;
        private Int32 _portCount = 0;
        private StreamWriter _sw;
        static private String _trackingFileName = "AM2GE_Tracking.kml";
        static private String _trackingFilePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Arcmap2GoogleEarth\\";
        static private String _trackingFileLocation = _trackingFilePath + _trackingFileName;
        static private String _trackingFileLocation2 = _trackingFilePath + "2" + _trackingFileName;
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
        }

        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }

        #endregion

        #region Method(s)

        private void initiliaze()
        {
            _application = this.Hook as IApplication;
            _mxdocument = (IMxDocument)_application.Document;
            _map = _mxdocument.FocusMap;
            ActiveViewEventTracking(true);


            if (!System.IO.Directory.Exists(_trackingFilePath))
                System.IO.Directory.CreateDirectory(_trackingFilePath);

            _sw = File.CreateText(_trackingFileLocation);

            _sw.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            _sw.WriteLine("<kml xmlns=\"http://www.opengis.net/kml/2.2\">");
            _sw.WriteLine("<NetworkLink>");
            _sw.WriteLine("<name>ArcMap to Google Earth Sync</name>");
            _sw.WriteLine("<visibility>1</visibility>");
            _sw.WriteLine("<LookAt>");
            _sw.WriteLine("<longitude>-84.103446</longitude>");
            _sw.WriteLine("<latitude>39.715620</latitude>");
            _sw.WriteLine("<altitude>0</altitude>");
            _sw.WriteLine("<heading>0</heading>");
            _sw.WriteLine("<tilt>0</tilt>");
            _sw.WriteLine("<range>510</range>");
            _sw.WriteLine("</LookAt>");
            _sw.WriteLine("<flyToView>1</flyToView>");
            _sw.WriteLine("<Link>");
            _sw.WriteLine("<href>" + _trackingFileLocation2 + "</href>");
            _sw.WriteLine("<refreshMode>onInterval</refreshMode>");
            _sw.WriteLine("<refreshInterval>2</refreshInterval>");
            _sw.WriteLine("</Link>");
            _sw.WriteLine("</NetworkLink>");
            _sw.WriteLine("</kml>");

            _sw.Close();

            System.Diagnostics.Process.Start(_trackingFileLocation);
            
        }

        private void CreateKML()
        {
            if (!System.IO.Directory.Exists(_trackingFilePath))
                System.IO.Directory.CreateDirectory(_trackingFilePath);

            _sw = File.CreateText(_trackingFileLocation2);

            _sw.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            _sw.WriteLine("<kml xmlns=\"http://www.opengis.net/kml/2.2\">");
            _sw.WriteLine("<Placemark>");
            _sw.WriteLine("<name>ArcMap Location</name>");
            _sw.WriteLine("<visibility>1</visibility>");
            _sw.WriteLine("<LookAt>");
            _sw.WriteLine("<longitude>" + _arcMapLong + "</longitude>");
            _sw.WriteLine("<latitude>" + _arcMapLat + "</latitude>");
            //_sw.WriteLine("<altitudeMode>clampToGround</altitudemode>")
            //_sw.WriteLine("<altitude>" + arcMapAlt + "</altitude>")
            _sw.WriteLine("<heading>0</heading>");
            _sw.WriteLine("<tilt>0</tilt>");
            _sw.WriteLine("<range>" + _arcMapAlt + "</range>");
            _sw.WriteLine("</LookAt>");
            _sw.WriteLine("<Point>");
            _sw.WriteLine("<coordinates>" + _arcMapLat + "," + _arcMapLong + "</coordinates>");
            _sw.WriteLine("</Point>");
            _sw.WriteLine("</Placemark>");
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

            String checkSumString, checkSumRMC, checkSumGLL, checkSumGGA, checkSumVTG;

            checkSumString = "$GPRMC," + nowTime + ",A," + convLat + ","+ convLong + ",0.00,0.00," + nowDate + ",0.0,E";
            checkSumRMC = GetCheckSum(ref checkSumString);

            _gpsString = "$GPRMC," + nowTime + ",A," + convLat + "," + convLong + ",0.00,0.00," + nowDate + ",0.0,E*" + checkSumRMC;
            _gpsRMCString = _gpsString;

            checkSumString = "$GPGLL," + convLat + "," + convLong + "," + nowTime + ",A";
            checkSumGLL = GetCheckSum(ref checkSumString);

            _gpsString = "$GPGLL," + convLat + "," + convLong + "," + nowTime + ",A*" + checkSumGLL;
            _gpsGLLString = _gpsString;

            checkSumString = "$GPGGA," + nowTime + "," + convLat + "," + convLong + ",1,5,0.0," + fakeGPSaltitude + ",M," + fakeGPSaltitude + ",M,,";
            checkSumGGA = GetCheckSum(ref checkSumString);

            _gpsString = "$GPGGA," + nowTime + "," + convLat + "," + convLong + ",1,5,0.0," + fakeGPSaltitude + ",M," + fakeGPSaltitude + ",M,,*" + checkSumGGA;
            _gpsGGAString = _gpsString;

            checkSumString = "$GPVTG,0.0,T,0.0,M,0.0,N,00.00,K";
            checkSumVTG = GetCheckSum(ref checkSumString);

            _gpsString = "$GPVTG,0.0,T,0.0,M,0.0,N,00.00,K*" + checkSumVTG;
            _gpsVTGString = _gpsString;


            FakeGPS();
            CreateKML();

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

        private void ActiveViewEventTracking(bool turnOn)
        {
            if (turnOn)
            {
                ESRI.ArcGIS.Carto.IActiveViewEvents_Event activeViewEvents = _map as ESRI.ArcGIS.Carto.IActiveViewEvents_Event;
                _ActiveViewEventsViewRefreshed = new ESRI.ArcGIS.Carto.IActiveViewEvents_ViewRefreshedEventHandler(OnActiveViewEventsViewRefreshed);
                activeViewEvents.ViewRefreshed += _ActiveViewEventsViewRefreshed;
            }
            else
            {

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

        private String GetCheckSum(ref String InString)
        {
            long Current;
            long Last;

            try
            {
                if (InString.Substring(1, 1) == "$")
                    InString = InString.Substring(2);

                Last = (int)InString.Substring(1, 1).ToCharArray()[0];

                for (Current = 2; Current <= InString.Length; Current++)
                    Last = Last ^ (int)InString.Substring((int)Current, 1).ToCharArray()[0];
                return String.Format("{0:X}", Last);
                

            }
            catch (Exception) { }
            return "0"; // this might cause a crash >....
        }

        private void FakeGPS()
        {
            if (_gpsRunning)
            {
                _serialPort.Write(_gpsRMCString + NL);
                System.Threading.Thread.Sleep(200);
                _serialPort.Write(_gpsGLLString + NL);
                System.Threading.Thread.Sleep(200);
                _serialPort.Write(_gpsGGAString + NL);
                System.Threading.Thread.Sleep(100);
                _serialPort.Write(_gpsVTGString + NL);
                System.Threading.Thread.Sleep(100);
                _serialPort.Write("$GPGSA,A,3,05,24,17,30,02,,,,,,,,5.6,3.3,4.5*34" + NL);
                System.Threading.Thread.Sleep(100);
                _serialPort.Write("$GPGSV,3,1,12,30,72,254,30,05,70,125,39,24,37,083,43,02,36,113,45*7B" + NL);
                System.Threading.Thread.Sleep(100);
                _serialPort.Write("$GPGSV,3,2,12,04,32,059,34,01,27,307,00,14,26,256,00,06,24,219,00*7F" + NL);
                System.Threading.Thread.Sleep(100);
                _serialPort.Write("$GPGSV,3,3,12,17,22,135,40,25,20,311,31,09,19,159,25,20,08,346,34*7C" + NL);
                System.Threading.Thread.Sleep(100);
                _serialPort.Write(_gpsRMCString + NL);
                System.Threading.Thread.Sleep(100);
                _serialPort.Write(_gpsGGAString + NL);
                System.Threading.Thread.Sleep(100);
                _serialPort.Write(_gpsGLLString + NL);
            }
        }
        
        #endregion
    }

}
