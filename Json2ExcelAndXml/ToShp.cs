using DrillingBuildLibrary.Model;
using OSGeo.OGR;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Json2ExcelAndXml
{
    public static class  ToShp
    {
        [DllImport("gdal202.dll", EntryPoint = "OGR_F_GetFieldAsString", CallingConvention = CallingConvention.Cdecl)]
        public extern static System.IntPtr OGR_F_GetFieldAsString(HandleRef handle, int index);

        public static void WriteVectorFileShp(String strVectorFile, Entity entity)  //创建算法生产的边界矢量图
        {

            // 注册GDAL
            // 为了支持中文路径，请添加下面这句代码
            OSGeo.GDAL.Gdal.SetConfigOption("GDAL_FILENAME_IS_UTF8", "YES");
            // 为了使属性表字段支持中文，请添加下面这句
            OSGeo.GDAL.Gdal.SetConfigOption("SHAPE_ENCODING", "");
            OSGeo.GDAL.Gdal.AllRegister();
            ////注册所有的驱动
            Ogr.RegisterAll();

            //创建数据，创建ESRI的shp文件
            string strDriverName = "ESRI Shapefile";
            Driver oDriver = Ogr.GetDriverByName(strDriverName);
            if (oDriver == null)
            {
                Debug.WriteLine("%s 驱动不可用！\n", strVectorFile);
                return;
            }

            // 步骤1、创建数据源
            DataSource oDS = oDriver.CreateDataSource(strVectorFile, null);
            if (oDS == null)
            {
                Debug.WriteLine("创建矢量文件【%s】失败！", strVectorFile);
                return;
            }
            //步骤2、创建空间坐标系
            OSGeo.OSR.SpatialReference oSRS = new OSGeo.OSR.SpatialReference("");
            //设置为西安80
            oSRS.SetWellKnownGeogCS("EPSG: 4610");
            //步骤3、创建图层，并添加坐标系，创建一个多边形图层(wkbGeometryType.wkbUnknown,存放任意几何特征)
            Layer oLayer = oDS.CreateLayer("TestPolygon", oSRS, wkbGeometryType.wkbUnknown, null);
            if (oLayer == null)
            {
                Debug.WriteLine("图层创建失败！");
                return;
            }

            try
            {
                if (entity != null)
                {



                    int columnCount = entity.fields.Count;//列数


                    for (int c = 0; c < columnCount; c++)
                    {

                        FieldDefn oFieldC = new FieldDefn($"{entity.fields[c].defName}", GetFieldType(entity.fields[c].type));
                        oFieldC.SetWidth(Convert.ToInt32(entity.fields[c].len));
                        oLayer.CreateField(oFieldC, 1);

                    }



                }
            }
            catch (Exception e)
            {

                throw e;
            }
            FeatureDefn oDefn = oLayer.GetLayerDefn();
            Feature oFeature = new Feature(oDefn);
            oLayer.CreateFeature(oFeature);
            oDS.Dispose();


            //// 步骤4、下面创建属性表
            //FieldDefn oFieldPlotArea = new FieldDefn("PlotArea", FieldType.OFTString);          // 先创建一个叫PlotArea的属性
            //oFieldPlotArea.SetWidth(100);
            //// 步骤5、将创建的属性表添加到图层中
            //oLayer.CreateField(oFieldPlotArea, 1);
            ////步骤6、定义一个特征要素oFeature(特征要素包含两个方面1.属性特征2.几何特征)
            //FeatureDefn oDefn = oLayer.GetLayerDefn();
            //Feature oFeature = new Feature(oDefn);    //建立了一个特征要素并将指向图层oLayer的属性表
            ////*****先不设置******/
            ////步骤7、设置属性特征的值 
            ////oFeature.SetField(0, area.ToString());

            ///******不要设置坐标信息******/
            ////OSGeo.OGR.Geometry geomTriangle = OSGeo.OGR.Geometry.CreateFromWkt(wkt);//创建一个几何特征
            //////步骤8、设置几何特征
            ////oFeature.SetGeometry(geomTriangle);
            ////步骤9、将特征要素添加到图层中
            //oLayer.CreateFeature(oFeature);
            //oDS.Dispose();
            //Debug.WriteLine("数据集创建完成！");
            MessageBox.Show("shp图层创建完成！");
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
