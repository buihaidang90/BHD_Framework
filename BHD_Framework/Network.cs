using System;
using System.Collections.Generic;
using System.Net;
using System.Text;

namespace BHD_Framework
{
    public static class Network
    {
        public static bool IsIpAddress(string IpString)
        {
            System.Net.IPAddress ip;
            bool validIP = System.Net.IPAddress.TryParse(IpString, out ip);
            return validIP;
        }

        public struct LocationDetails
        {
            public string status { get; set; }
            public string message { get; set; }
            public string query { get; set; }
            public string country { get; set; }
            public string countryCode { get; set; }
            public string region { get; set; }
            public string regionName { get; set; }
            public string city { get; set; }
            public string isp { get; set; }
            public string org { get; set; }
            public string timezone { get; set; }
            public double latitude { get; set; }
            public double longtitude { get; set; }
            public string zip { get; set; }
        }
        public static LocationDetails GetLocationDetail(string IpAddress, int? TimeOutSpace)
        {
            LocationDetails geoInfo = new LocationDetails();
            geoInfo.status = "null";
            geoInfo.city = "";
            geoInfo.country = "";
            geoInfo.countryCode = "";
            geoInfo.region = "";
            geoInfo.regionName = "";
            geoInfo.isp = "";
            geoInfo.org = "";
            geoInfo.timezone = "";
            if (!IsIpAddress(IpAddress))
            {
                geoInfo.status = "fail";
                geoInfo.message = "invalid query";
                geoInfo.query = IpAddress;
                return geoInfo;
            }
            string IpApiUrl = string.Concat("http://ip-api.com/json/", IpAddress);
            //
            using (ExtendedWebClient client = new ExtendedWebClient(IpApiUrl, TimeOutSpace))
            {
                //client.Headers.Add(HttpRequestHeader.ContentType, "text/xml");
                client.Headers[HttpRequestHeader.ContentType] = "application/json"; // push parameters follow json type
                client.Encoding = Encoding.UTF8; // encode string has sign
                string strJson = client.DownloadString(IpApiUrl);
                //Console.WriteLine(strJson);
                strJson = strJson.Replace("{", "").Replace("}", "").Replace("\r\n", "").Replace(@"""", "");
                string[] arr = strJson.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                string splitSign = ":"
                    , keyStatus = "status", keyMessage = "message", keyCountry = "country", keyCountryCode = "countryCode"
                    , keyRegion = "region", keyRegionName = "regionName", keyCity = "city"
                    , keyZip = "zip", keyLatitude = "lat", keyLongtitude = "lon", keyTimezone = "timezone"
                    , keyIsp = "isp", keyOrg = "org", keyQuery = "query";

                foreach (string item in arr)
                {
                    string s = "";
                    s = Utility.GetValueOf(keyStatus, splitSign, item); if (s != "")
                    {
                        geoInfo.status = s;
                        continue;
                    }
                    s = Utility.GetValueOf(keyMessage, splitSign, item); if (s != "")
                    {
                        geoInfo.message = s;
                        continue;
                    }
                    s = Utility.GetValueOf(keyCountry, splitSign, item); if (s != "")
                    {
                        geoInfo.country = s;
                        continue;
                    }
                    s = Utility.GetValueOf(keyCountryCode, splitSign, item); if (s != "")
                    {
                        geoInfo.countryCode = s;
                        continue;
                    }
                    s = Utility.GetValueOf(keyRegion, splitSign, item); if (s != "")
                    {
                        geoInfo.region = s;
                        continue;
                    }
                    s = Utility.GetValueOf(keyRegionName, splitSign, item); if (s != "")
                    {
                        geoInfo.regionName = s;
                        continue;
                    }
                    s = Utility.GetValueOf(keyCity, splitSign, item); if (s != "")
                    {
                        geoInfo.city = s;
                        continue;
                    }
                    s = Utility.GetValueOf(keyZip, splitSign, item); if (s != "")
                    {
                        geoInfo.zip = s;
                        continue;
                    }
                    s = Utility.GetValueOf(keyLatitude, splitSign, item); if (s != "")
                    {
                        double d = 0;
                        if (double.TryParse(s, out d)) geoInfo.latitude = d;
                        continue;
                    }
                    s = Utility.GetValueOf(keyLongtitude, splitSign, item); if (s != "")
                    {
                        double d = 0;
                        if (double.TryParse(s, out d)) geoInfo.longtitude = d;
                        continue;
                    }
                    s = Utility.GetValueOf(keyTimezone, splitSign, item); if (s != "")
                    {
                        geoInfo.timezone = s;
                        continue;
                    }
                    s = Utility.GetValueOf(keyIsp, splitSign, item); if (s != "")
                    {
                        geoInfo.isp = s;
                        continue;
                    }
                    s = Utility.GetValueOf(keyOrg, splitSign, item); if (s != "")
                    {
                        geoInfo.org = s;
                        continue;
                    }
                    s = Utility.GetValueOf(keyQuery, splitSign, item); if (s != "")
                    {
                        geoInfo.query = s;
                        continue;
                    }
                }
            }
            return geoInfo;
            /*
            {
                "status": "success",
                "country": "Vietnam",
                "countryCode": "VN",
                "region": "SG",
                "regionName": "Ho Chi Minh",
                "city": "Ho Chi Minh City",
                "zip": "",
                "lat": 10.822,
                "lon": 106.6257,
                "timezone": "Asia/Ho_Chi_Minh",
                "isp": "VNPT",
                "org": "Vietnam Posts and Telecommunications Group",
                "as": "AS45899 VNPT Corp",
                "query": "14.161.31.0"
            }
             */
        }


    }


    public class ExtendedWebClient : WebClient
    {
        private const int DefaultTimeoutSpace = 10000; /// millisecond = 10 second
        private const int MinTimeoutSpace = 1; /// millisecond
        private const int MaxTimeoutSpace = 300000; /// millisecond = 10 minute
        private int _Timeout = DefaultTimeoutSpace;
        public int Timeout
        {
            get
            {
                return _Timeout;
            }
            set
            {
                if (MinTimeoutSpace <= value && value <= MaxTimeoutSpace)
                    _Timeout = value;
                else
                    _Timeout = DefaultTimeoutSpace;
            }
        }
        public ExtendedWebClient(Uri address, int timeout)
        {
            this._Timeout = timeout;//In Milli seconds
            var objWebClient = GetWebRequest(address);
        }
        public ExtendedWebClient(string stringUriAddress, int? timeout)
        {
            int timeoutValue = (timeout == null ? DefaultTimeoutSpace : timeout.Value);
            this._Timeout = timeoutValue;//In Milli seconds
            if (!Uri.IsWellFormedUriString(stringUriAddress, UriKind.Absolute))
            {
                throw new Exception("Your uri address is not valid.");
            }
            var objWebClient = GetWebRequest(new Uri(stringUriAddress, UriKind.Absolute));
        }

        protected override WebRequest GetWebRequest(Uri address)
        {
            WebRequest wr = base.GetWebRequest(address);
            wr.Timeout = this._Timeout;
            ((HttpWebRequest)wr).ReadWriteTimeout = this._Timeout; // dont remove
            return wr;
        }
    }



}
