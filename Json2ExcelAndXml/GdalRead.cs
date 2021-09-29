using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//
using System.IO;
using OSGeo.GDAL;
using OSGeo.OGR;
using OSGeo.OSR;
using System.Collections;
using Json2ExcelAndXml;

namespace GdalReadSHP
{
    /// <summary>
    /// 定义SHP解析类
    /// </summary>
    public class ShpRead
    {
        /// 保存SHP属性字段
        public OSGeo.OGR.Driver oDerive;
        public List<string> m_FeildList;
        public Layer readLayer;
        public string sCoordiantes;
        public ShpRead()
        {

            GdalConfiguration.ConfigureGdal();
            GdalConfiguration.ConfigureOgr();

            m_FeildList = new List<string>();
            readLayer = null;
            sCoordiantes = null;
        }

        /// <summary>
        /// 初始化Gdal
        /// </summary>
        public void InitinalGdal()
        {
            // 为了支持中文路径
            Gdal.SetConfigOption("GDAL_FILENAME_IS_UTF8", "YES");
            // 为了使属性表字段支持中文
            Gdal.SetConfigOption("SHAPE_ENCODING", "");
            Gdal.AllRegister();
            Ogr.RegisterAll();

            oDerive = Ogr.GetDriverByName("ESRI Shapefile");
            if (oDerive == null)
            {
                MessageBox.Show("文件不能打开，请检查");
            }
        }

        /// <summary>
        /// 获取SHP文件的层
        /// </summary>
        /// <param name="sfilename"></param>
        /// <param name="oLayer"></param>
        /// <returns></returns>
        public bool GetShpLayer(string sfilename)
        {
            if (null == sfilename || sfilename.Length <= 3)
            {
                readLayer = null;
                return false;
            }
            if (oDerive == null)
            {
                MessageBox.Show("文件不能打开，请检查");
            }
            DataSource ds = oDerive.Open(sfilename, 1);
            if (null == ds)
            {
                readLayer = null;
                return false;
            }
            int iPosition = sfilename.LastIndexOf("\\");
            string sTempName = sfilename.Substring(iPosition + 1, sfilename.Length - iPosition - 4 - 1);
            readLayer = ds.GetLayerByName(sTempName);
            if (readLayer == null)
            {
                ds.Dispose();
                return false;
            }
            return true;
        }
        /// <summary>
        /// 获取所有的属性字段
        /// </summary>
        /// <returns></returns>
        public bool GetFeilds()
        {
            if (null == readLayer)
            {
                return false;
            }
            m_FeildList.Clear();
            wkbGeometryType oTempGeometryType = readLayer.GetGeomType();
            List<string> TempstringList = new List<string>();

            //
            FeatureDefn oDefn = readLayer.GetLayerDefn();
            int iFieldCount = oDefn.GetFieldCount();
            for (int iAttr = 0; iAttr < iFieldCount; iAttr++)
            {
                FieldDefn oField = oDefn.GetFieldDefn(iAttr);
                if (null != oField)
                {
                    m_FeildList.Add(oField.GetNameRef());
                }
            }
            return true;
        }
        /// <summary>
        ///  获取某条数据的字段内容
        /// </summary>
        /// <param name="iIndex"></param>
        /// <param name="FeildStringList"></param>
        /// <returns></returns>
        public bool GetFeildContent(int iIndex, out List<string> FeildStringList)
        {
            FeildStringList = new List<string>();
            Feature oFeature = null;
            if ((oFeature = readLayer.GetFeature(iIndex)) != null)
            {

                FeatureDefn oDefn = readLayer.GetLayerDefn();
                int iFieldCount = oDefn.GetFieldCount();
                // 查找字段属性
                for (int iAttr = 0; iAttr < iFieldCount; iAttr++)
                {
                    FieldDefn oField = oDefn.GetFieldDefn(iAttr);
                    string sFeildName = oField.GetNameRef();

                    #region 获取属性字段
                    FieldType Ftype = oFeature.GetFieldType(sFeildName);
                    switch (Ftype)
                    {
                        case FieldType.OFTString:
                            string sFValue = oFeature.GetFieldAsString(sFeildName);
                            string sTempType = "string";
                            FeildStringList.Add(sFValue);
                            break;
                        case FieldType.OFTReal:
                            double dFValue = oFeature.GetFieldAsDouble(sFeildName);
                            sTempType = "float";
                            FeildStringList.Add(dFValue.ToString());
                            break;
                        case FieldType.OFTInteger:
                            int iFValue = oFeature.GetFieldAsInteger(sFeildName);
                            sTempType = "int";
                            FeildStringList.Add(iFValue.ToString());
                            break;
                        default:
                            //sFValue = oFeature.GetFieldAsString(ChosenFeildIndex[iFeildIndex]);
                            sTempType = "string";
                            break;
                    }
                    #endregion
                }
            }
            return true;
        }
        /// <summary>
        /// 获取数据
        /// </summary>
        /// <returns></returns>
        public bool GetGeometry(int iIndex)
        {
            if (null == readLayer)
            {
                return false;
            }
            long iFeatureCout = readLayer.GetFeatureCount(0);
            Feature oFeature = null;
            oFeature = readLayer.GetFeature(iIndex);
            //  Geometry
            Geometry oGeometry = oFeature.GetGeometryRef();
            wkbGeometryType oGeometryType = oGeometry.GetGeometryType();
            switch (oGeometryType)
            {
                case wkbGeometryType.wkbPoint:
                    oGeometry.ExportToWkt(out sCoordiantes);
                    sCoordiantes = sCoordiantes.ToUpper().Replace("POINT (", "").Replace(")", "");
                    break;
                case wkbGeometryType.wkbLineString:
                case wkbGeometryType.wkbLinearRing:
                    oGeometry.ExportToWkt(out sCoordiantes);
                    sCoordiantes = sCoordiantes.ToUpper().Replace("LINESTRING (", "").Replace(")", "");
                    break;
                default:
                    break;
            }
            return false;
        }

    }//END class
}