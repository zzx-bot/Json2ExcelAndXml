using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace DrillingBuildLibrary
{
    class ToXML
    {
        public static void SerialzeToXML<T>(string filePath,T data )
        {
            if (string.IsNullOrEmpty(filePath)) return;
            if (data == null) return;
            try
            {
                using (FileStream fileStream =new FileStream(filePath,FileMode.Create,FileAccess.Write))
                {
                    StreamWriter streamWriter = new StreamWriter(fileStream);
                    XmlSerializer xs = new XmlSerializer(typeof(T));
                    xs.Serialize(streamWriter, data);
                }
            }
            catch (Exception ex)
            {

                throw new Exception($"XML序列化失败:{ex.Message}");
            }
        }
    }
}
