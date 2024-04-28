using System.IO;
using System.Text;
using System.Xml.Serialization;

namespace ECDIS_eGloebe___RouteConverter.Utilities
{

    public class XmlHelper
    {
        public T Deserialize<T>(string inputXml, string rootName)
        {
            using (StringReader reader = new StringReader(inputXml))
            {
                XmlRootAttribute xmlRoot = new XmlRootAttribute(rootName);

                XmlSerializer serializer = new XmlSerializer(typeof(T), xmlRoot);

                return (T)serializer
                    .Deserialize(reader);
            }
        }

        public string Serialize<T>(T obj, string rootName)
        {
            StringBuilder sb = new StringBuilder();

            XmlRootAttribute xmlRoot =
                new XmlRootAttribute(rootName);

            XmlSerializerNamespaces xmlNamespace =
                new XmlSerializerNamespaces();
            xmlNamespace
                .Add(string.Empty, string.Empty);

            StringWriter writer =
                new StringWriter(sb);

            XmlSerializer serializer =
                new XmlSerializer(typeof(T), xmlRoot);

            serializer
                .Serialize(writer, obj, xmlNamespace);

            return sb.ToString().TrimEnd();
        }
    }
}