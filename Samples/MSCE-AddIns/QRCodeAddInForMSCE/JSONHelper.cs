using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization.Json;
using System.IO;

namespace QRCodeAddInForMSCE
{
    /// <summary>
    /// Helper class to serialize and deserialize JSON without a dependency on JSON.NET.  That class is great, but I've run into a lot of version conflicts when trying to deploy it
    /// Usage example:
    /// 
    /// using (var client = new WebClient())
    /// {
    ///     Shortener urlShortener = new Shortener() { longUrl = sLongURL };
    ///     var body = JSONHelper.SerializeJSon<Shortener>(urlShortener);
    ///     client.Headers[HttpRequestHeader.Accept] = "application/json";
    ///     client.Headers[HttpRequestHeader.ContentType] = "application/json";
    ///     byte[] response = client.UploadData(new Uri(url), "POST", Encoding.UTF8.GetBytes(body));
    ///     ShortenedURL shortUrl = JSONHelper.DeserializeJSon<ShortenedURL>(response);
    ///     System.Diagnostics.Debug.WriteLine(shortUrl.id);
    ///  }
    /// </summary>
    public class JSONHelper
    {
        public static T DeserializeJSon<T>(string jsonString)
        {
            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(T));
            MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(jsonString));
            T obj = (T)ser.ReadObject(stream);
            return obj;
        }

        public static T DeserializeJSon<T>(byte[] jsonBytes)
        {
            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(T));
            MemoryStream stream = new MemoryStream(jsonBytes);
            T obj = (T)ser.ReadObject(stream);
            return obj;
        }
        public static string SerializeJSon<T>(T t)
        {
            MemoryStream stream = new MemoryStream();
            DataContractJsonSerializer ds = new DataContractJsonSerializer(typeof(T));
            DataContractJsonSerializerSettings s = new DataContractJsonSerializerSettings();
            ds.WriteObject(stream, t);
            string jsonString = Encoding.UTF8.GetString(stream.ToArray());
            stream.Close();
            return jsonString;
        }
    }
}
