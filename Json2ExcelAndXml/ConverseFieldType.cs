using OSGeo.OGR;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Json2ExcelAndXml
{
    class ConverseFieldType
    {
        public ConverseFieldType()
        {

        }

        public static FieldType GetFieldType(string type)
        {
            switch (type)
            {
                case "L6":
                case "VARCHAR":
                    return FieldType.OFTString;
                case "NUMERIC":
                case "D10.1":
                    return FieldType.OFTReal;
                case "INTEGER":
                    return FieldType.OFTInteger;
                case "DATE":
                    return FieldType.OFTDate;
                case "BLOB":
                    return FieldType.OFTBinary;

                default:
                    return FieldType.OFTString;
            }
        }
    }
}
