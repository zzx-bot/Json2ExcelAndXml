
using DrillingBuildLibrary.Model;
using GdalReadSHP;
using OSGeo.OGR;
using OSGeo.OSR;
using SmartGeological.Database.Core;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace Json2ExcelAndXml
{

    public class GdalUtilities
    {

        public GdalUtilities()
        {
            GdalConfiguration.ConfigureGdal();
            GdalConfiguration.ConfigureOgr();
        }

        /// <summary>
        /// 旧
        /// </summary>
        /// <param name="entity"></param>
        /// <param name="strVectorFile"></param>
        /// <param name="_element"></param>
        /// <returns></returns>
        public bool convertJsonToShapeFile(Entity entity, string strVectorFile, Entity _element)
        {
            string elementCode = "";
            string elementType = "";
            foreach (var field in _element.fields)
            {
                if(entity.defKey.Contains(field.defKey))
                {
                    elementType = field.type;
                    elementCode = field.defKey;
                }
            }
           

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
                return false ;
            }

            // 步骤1、创建数据源
            DataSource oDS = oDriver.CreateDataSource(strVectorFile, null);
            if (oDS == null)
            {
                Debug.WriteLine("创建矢量文件【%s】失败！", strVectorFile);
                return  false;
            }
            //步骤2、创建空间坐标系
            OSGeo.OSR.SpatialReference oSRS = new OSGeo.OSR.SpatialReference("");

            ////设置为西安80
            ////oSRS.SetProjCS("UTM 49(WGS84) in northern hemisphere.");
            //oSRS.SetWellKnownGeogCS("GCSXian1980");
            ////oSRS.SetUTM(49, 0);
            ////打印
            //string pszWKT;
            //oSRS.ExportToWkt(out pszWKT);
            //Console.WriteLine(pszWKT);

            //读取一个图层
            ShpRead shpRead = new ShpRead();
            shpRead.InitinalGdal();
            shpRead.GetShpLayer(@"C:\Users\210320\Desktop\touying\DLBHLU01.shp");

            //步骤3、创建图层，并添加坐标系，创建一个多边形图层(wkbGeometryType.wkbUnknown, 存放任意几何特征)

            oSRS = shpRead.readLayer.GetSpatialRef();

            if (oDS.GetLayerByName($"{elementCode}")!=null)
                oDS.DeleteLayer(0) ;

            Layer oLayer = oDS.CreateLayer($"{elementCode}", oSRS, GetGeometryType(elementType), null);


            if (oLayer == null)
            {
                Debug.WriteLine("图层创建失败！");
                return  false;
            }

            try
            {
                if (entity != null)
                {



                    int columnCount = entity.fields.Count;//列数


                    for (int c = 0; c < columnCount; c++)
                    {

                        FieldDefn oFieldC = new FieldDefn($"{entity.fields[c].defKey}", GetFieldType(entity.fields[c].type));
                        int len = 20;
                        if (entity.fields[c].len!=null&&entity.fields[c].len.ToString() != "")
                            len = Convert.ToInt32(entity.fields[c].len);
                        oFieldC.SetWidth(len);

                        FeatureDefn oDefn1 = oLayer.GetLayerDefn();
                        int FieldCount1 = oDefn1.GetFieldCount();

                        oLayer.CreateField(oFieldC, FieldCount1+1);
                        oDefn1.Dispose();
                        oFieldC.Dispose();
                    }



                }
            }
            catch (Exception e)
            {

                throw e;
            }

            //FeatureDefn oDefn = oLayer.GetLayerDefn();
            //Feature oFeature = new Feature(oDefn);
            //oLayer.CreateFeature(oFeature);
            
            //将资源都关闭
            oSRS.Dispose();
            oLayer.Dispose();
            oDS.Dispose();
            oDriver.Dispose();






            #region   json 区域


            //Driver jsonFileDriver = Ogr.GetDriverByName("GeoJSON");
            //DataSource jsonFile = Ogr.Open(jsonFilePath, 0);
            //if (jsonFile == null)
            //{
            //    return false;
            //}

            //string filesPathName = shapeFilePath.Substring(0, shapeFilePath.Length - 4);
            //removeShapeFileIfExists(filesPathName);

            //Layer jsonLayer = jsonFile.GetLayerByIndex(0);

            //Driver esriShapeFileDriver = Ogr.GetDriverByName("ESRI Shapefile");
            //DataSource shapeFile = Ogr.Open(shapeFilePath, 0);

            //shapeFile = esriShapeFileDriver.CreateDataSource(shapeFilePath, new string[] { });
            //Layer shplayer = shapeFile.CreateLayer(jsonLayer.GetName(), jsonLayer.GetSpatialRef(), jsonLayer.GetGeomType(), new string[] { });


            //// create fields (properties) in new layer
            //Feature jsonFeature = jsonLayer.GetNextFeature();
            //for (int i = 0; i < jsonFeature.GetFieldCount(); i++)
            //{
            //    FieldDefn fieldDefn = new FieldDefn(getValidFieldName(jsonFeature.GetFieldDefnRef(i)), jsonFeature.GetFieldDefnRef(i).GetFieldType());
            //    shplayer.CreateField(fieldDefn, 1);
            //}

            //while (jsonFeature != null)
            //{
            //    Geometry geometry = jsonFeature.GetGeometryRef();
            //    Feature shpFeature = createGeometryFromGeometry(geometry, shplayer, jsonLayer.GetSpatialRef());

            //    // copy values for each field
            //    for (int i = 0; i < jsonFeature.GetFieldCount(); i++)
            //    {
            //        if (FieldType.OFTInteger == jsonFeature.GetFieldDefnRef(i).GetFieldType())
            //        {
            //            shpFeature.SetField(getValidFieldName(jsonFeature.GetFieldDefnRef(i)), jsonFeature.GetFieldAsInteger(i));
            //        }
            //        else if (FieldType.OFTReal == jsonFeature.GetFieldDefnRef(i).GetFieldType())
            //        {
            //            shpFeature.SetField(getValidFieldName(jsonFeature.GetFieldDefnRef(i)), jsonFeature.GetFieldAsDouble(i));
            //        }
            //        else
            //        {
            //            shpFeature.SetField(getValidFieldName(jsonFeature.GetFieldDefnRef(i)), jsonFeature.GetFieldAsString(i));
            //        }
            //    }
            //    shplayer.SetFeature(shpFeature);

            //    jsonFeature = jsonLayer.GetNextFeature();
            //}


            //shapeFile.Dispose();

            //// if you want to generate zip of generated files
            //string zipName = filesPathName + ".zip";
            ////CompressToZipFile(new List<string>() { shapeFilePath, filesPathName + ".dbf", filesPathName + ".prj", filesPathName + ".shx" }, zipName);

            #endregion


            return true;
        }


        public bool convertJsonToShapeFile(Entity entity, string strVectorFile)
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
                return false;
            }

            // 步骤1、创建数据源
            DataSource oDS = oDriver.CreateDataSource(strVectorFile, null);
            if (oDS == null)
            {
                Debug.WriteLine("创建矢量文件【%s】失败！", strVectorFile);
                return false;
            }
            //步骤2、创建空间坐标系
            OSGeo.OSR.SpatialReference oSRS = new OSGeo.OSR.SpatialReference("");

            ////设置为西安80
            ////oSRS.SetProjCS("UTM 49(WGS84) in northern hemisphere.");
            //oSRS.SetWellKnownGeogCS("GCSXian1980");
            ////oSRS.SetUTM(49, 0);
            ////打印
            //string pszWKT;
            //oSRS.ExportToWkt(out pszWKT);
            //Console.WriteLine(pszWKT);

            //读取一个图层
            ShpRead shpRead = new ShpRead();
            shpRead.InitinalGdal();
            shpRead.GetShpLayer(@"C:\Users\210320\Desktop\touying\DLBHLU01.shp");

            //步骤3、创建图层，并添加坐标系，创建一个多边形图层(wkbGeometryType.wkbUnknown, 存放任意几何特征)

            oSRS = shpRead.readLayer.GetSpatialRef();

            if (oDS.GetLayerByName($"{entity.defKey}") != null)
                oDS.DeleteLayer(0);

            Layer oLayer = oDS.CreateLayer($"{entity.defKey}", oSRS, GetGeometryType(entity.comment), null);


            if (oLayer == null)
            {
                Debug.WriteLine("图层创建失败！");
                return false;
            }

            try
            {
                if (entity != null)
                {



                    int columnCount = entity.fields.Count;//列数


                    for (int c = 0; c < columnCount; c++)
                    {

                        FieldDefn oFieldC = new FieldDefn($"{entity.fields[c].defKey}", GetFieldType(entity.fields[c].type));
                        //oFieldC.SetName(entity.fields[c].defName);

                        int len = 20;
                        if (entity.fields[c].len != null && entity.fields[c].len.ToString() != "")
                            len = Convert.ToInt32(entity.fields[c].len);

                        if (entity.fields[c].type == "Char")
                            oFieldC.SetWidth(len);
                        else if (entity.fields[c].type == "Float")
                        {
                            oFieldC.SetWidth(len+1);
                            oFieldC.SetPrecision(Convert.ToInt32(entity.fields[c].scale));
                        }


                        FeatureDefn oDefn1 = oLayer.GetLayerDefn();
                        int FieldCount1 = oDefn1.GetFieldCount();

                        oLayer.CreateField(oFieldC, FieldCount1 + 1);
                        oDefn1.Dispose();
                        oFieldC.Dispose();
                    }



                }
            }
            catch (Exception e)
            {

                throw e;
            }

            //FeatureDefn oDefn = oLayer.GetLayerDefn();
            //Feature oFeature = new Feature(oDefn);
            //oLayer.CreateFeature(oFeature);

            //将资源都关闭
            oSRS.Dispose();
            oLayer.Dispose();
            oDS.Dispose();
            oDriver.Dispose();






            #region   json 区域


            //Driver jsonFileDriver = Ogr.GetDriverByName("GeoJSON");
            //DataSource jsonFile = Ogr.Open(jsonFilePath, 0);
            //if (jsonFile == null)
            //{
            //    return false;
            //}

            //string filesPathName = shapeFilePath.Substring(0, shapeFilePath.Length - 4);
            //removeShapeFileIfExists(filesPathName);

            //Layer jsonLayer = jsonFile.GetLayerByIndex(0);

            //Driver esriShapeFileDriver = Ogr.GetDriverByName("ESRI Shapefile");
            //DataSource shapeFile = Ogr.Open(shapeFilePath, 0);

            //shapeFile = esriShapeFileDriver.CreateDataSource(shapeFilePath, new string[] { });
            //Layer shplayer = shapeFile.CreateLayer(jsonLayer.GetName(), jsonLayer.GetSpatialRef(), jsonLayer.GetGeomType(), new string[] { });


            //// create fields (properties) in new layer
            //Feature jsonFeature = jsonLayer.GetNextFeature();
            //for (int i = 0; i < jsonFeature.GetFieldCount(); i++)
            //{
            //    FieldDefn fieldDefn = new FieldDefn(getValidFieldName(jsonFeature.GetFieldDefnRef(i)), jsonFeature.GetFieldDefnRef(i).GetFieldType());
            //    shplayer.CreateField(fieldDefn, 1);
            //}

            //while (jsonFeature != null)
            //{
            //    Geometry geometry = jsonFeature.GetGeometryRef();
            //    Feature shpFeature = createGeometryFromGeometry(geometry, shplayer, jsonLayer.GetSpatialRef());

            //    // copy values for each field
            //    for (int i = 0; i < jsonFeature.GetFieldCount(); i++)
            //    {
            //        if (FieldType.OFTInteger == jsonFeature.GetFieldDefnRef(i).GetFieldType())
            //        {
            //            shpFeature.SetField(getValidFieldName(jsonFeature.GetFieldDefnRef(i)), jsonFeature.GetFieldAsInteger(i));
            //        }
            //        else if (FieldType.OFTReal == jsonFeature.GetFieldDefnRef(i).GetFieldType())
            //        {
            //            shpFeature.SetField(getValidFieldName(jsonFeature.GetFieldDefnRef(i)), jsonFeature.GetFieldAsDouble(i));
            //        }
            //        else
            //        {
            //            shpFeature.SetField(getValidFieldName(jsonFeature.GetFieldDefnRef(i)), jsonFeature.GetFieldAsString(i));
            //        }
            //    }
            //    shplayer.SetFeature(shpFeature);

            //    jsonFeature = jsonLayer.GetNextFeature();
            //}


            //shapeFile.Dispose();

            //// if you want to generate zip of generated files
            //string zipName = filesPathName + ".zip";
            ////CompressToZipFile(new List<string>() { shapeFilePath, filesPathName + ".dbf", filesPathName + ".prj", filesPathName + ".shx" }, zipName);

            #endregion


            return true;
        }
        public static FieldType GetFieldType(string type)
        {
            switch (type)
            {
                case "Char":
                    return FieldType.OFTString;
                case "Float":
                    return FieldType.OFTReal;
                case "Int":
                    return FieldType.OFTInteger;
                default:
                    return FieldType.OFTString;
            }
        }


        public static wkbGeometryType GetGeometryType(string type)
        {
            switch (type)
            {
                case "点":
                    return wkbGeometryType.wkbMultiPoint;

                case "线":
                    return wkbGeometryType.wkbLineString;
                case "面":
                    return wkbGeometryType.wkbMultiPolygon;

                default:
                    return wkbGeometryType.wkbMultiPoint;
            }
        }

        private void removeShapeFileIfExists(string filesPathName)
        {
            removeFileIfExists(filesPathName + ".shp");
            removeFileIfExists(filesPathName + ".shx");
            removeFileIfExists(filesPathName + ".prj");
            removeFileIfExists(filesPathName + ".zip");
        }

        public static bool removeFileIfExists(string filePath)
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
                return true;
            }
            return false;
        }

        // the field names in shapefile have limit of 10 characteres
        private string getValidFieldName(FieldDefn fieldDefn)
        {
            string fieldName = fieldDefn.GetName();
            return fieldName.Length > 10 ? fieldName.Substring(0, 10) : fieldName;
        }

        private Feature createGeometryFromGeometry(Geometry geometry, Layer layer, SpatialReference reference)
        {
            Feature feature = new Feature(layer.GetLayerDefn());

            string wktgeometry = "";
            geometry.ExportToWkt(out wktgeometry);
            Geometry newGeometry = Geometry.CreateFromWkt(wktgeometry);
            newGeometry.AssignSpatialReference(reference);
            newGeometry.SetPoint(0, geometry.GetX(0), geometry.GetY(0), 0);

            feature.SetGeometry(newGeometry);
            layer.CreateFeature(feature);

            return feature;
        }

        //public static void CompressToZipFile(List<string> files, string zipPath)
        //{
        //    using (ZipFile zip = new ZipFile())
        //    {
        //        foreach (string file in files)
        //        {
        //            zip.AddFile(file, "");
        //        }
        //        zip.Save(zipPath);
        //    }
        //}
    }
}