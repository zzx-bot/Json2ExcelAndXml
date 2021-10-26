using DrillingBuildLibrary.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Json2ExcelAndXml.SQL
{
    public class InputPgDB
    {
        public static void CreatSpatialTable(Entity entity)
        {

            if (entity != null)
            {
                try
                {
                    String sqlcom = "create table " + $"{entity.defKey}" + " (";
                    int columnCount = entity.fields.Count;//列数


                    string attribute = "";
                    string comment = "";

                    //                  CREATE TABLE "public"."table_pgtable"(
                    //                    "id" int4 NOT NULL DEFAULT nextval('auto_id'::regclass),
                    //                    "gid" varchar(36) DEFAULT uuid_generate_v4(),
                    //                     "name" varchar(255) COLLATE "pg_catalog"."default",
                    //  "geometry" "public"."geometry",
                    //  "remarks" varchar(255) COLLATE "pg_catalog"."default",
                    //  "time" timestamp(6) DEFAULT CURRENT_TIMESTAMP
                    //);
                    //                  COMMENT ON COLUMN "public"."table_pgtable"."id" IS '自增长id';
                    //                  COMMENT ON COLUMN "public"."table_pgtable"."gid" IS 'uuid';
                    //                  COMMENT ON COLUMN "public"."table_pgtable"."name" IS '名称';
                    //                  COMMENT ON COLUMN "public"."table_pgtable"."geometry" IS '几何信息';
                    //                  COMMENT ON COLUMN "public"."table_pgtable"."remarks" IS '备注';
                    //                  COMMENT ON COLUMN "public"."table_pgtable"."time" IS '时间';


                    for (int c = 0; c < columnCount; c++)
                    {
                        int len = 20, scale = 0;

                        if (entity.fields[c].len != null && entity.fields[c].len.ToString() != "")
                            len = Convert.ToInt32(entity.fields[c].len);

                        if (entity.fields[c].scale != null && entity.fields[c].scale.ToString() != "")
                            scale = Convert.ToInt32(entity.fields[c].scale);

                        //字段名字及类型
                        if (entity.fields[c].type.ToUpper() == "NUMERIC")
                       
                            attribute += entity.fields[c].defKey + " NUMERIC(" + $"{len.ToString()}" + "," + $"{scale.ToString()}" + ")";
                       
                        else if (entity.fields[c].type.ToUpper() == "INT")

                            attribute += entity.fields[c].defKey + " int4";

                        else if (entity.fields[c].type.ToUpper() == "VARCHAR")

                            attribute += entity.fields[c].defKey + " varchar(" + $"{len.ToString()}" + ")";

                        else if (entity.fields[c].type.ToUpper() == "BLOB")

                            attribute += entity.fields[c].defKey + " varchar(" + $"{len.ToString()}" + ")";

                        else if (entity.fields[c].type.ToUpper() == "DATE")

                            attribute += entity.fields[c].defKey + " varchar(" + $"{len.ToString()}" + ")";
                        //字段是否为空
                        if (entity.fields[c].notNull)
                            attribute += " not null, ";
                        else
                            attribute += ",";
                        //添加字段描述
                        comment += "COMMENT ON COLUMN " + entity.defKey + "." + entity.fields[c].defKey + " IS " + "'" + entity.fields[c].comment + "';";

                    }

                    //添加几何要素
                    //attribute += "the_geog" + " geography(" + GetGeometryType(entity.comment) + "," + " 4610" + "));";

                    ////添加字段描述
                    //comment += "COMMENT ON COLUMN " + entity.defKey + ".the_geog" + " IS " + "'" + entity.comment + "';";

                    AddComment addComment = new AddComment();
                    addComment.ExecuteNonQuery(sqlcom + attribute + comment);

                    //INSERT INTO public.table_pgtable(name, geometry, remarks) VALUES('点数据', st_GeomFromText('POINT (120.163629 30.2592077)'),'点数据');


                }
                catch (Exception e)
                {

                    throw e;
                }
            }

        }


        public static void InsertSpatialTable(Entity entity)
        {

            if (entity != null)
            {
                try
                {
                    String sqlcom = "INSERT INTO " + $"{entity.defKey}" + " (";
                    int columnCount = entity.fields.Count;//列数


                    string attribute = "";
                    string value = " VALUES (";

                    //        INSERT INTO cggccdbg ( chfcac, gcicad, swclae, swnda, the_geog )
                    //
                    //        VALUES
                    //        (
                    //                'Tom',
                    //                'a',
                    //                '1.2',
                    //                'Jerry',
                    //        ST_GeographyFromText ('SRID=4610 ;LINESTRING ( 120.163629 30.2592077, 120.163599 30.259802, 120.163617 30.259803, 120.163647 30.259207 )' ) );



                    for (int c = 0; c < columnCount; c++)
                    {

                        attribute += entity.fields[c].defKey+" , ";

                        //添加字段描述
                        value += "'1',";

                    }

                    attribute += "the_geog )";

                    if(GetGeometryType(entity.comment)== "POINT")
                    //添加字段描述
                        value += "ST_GeographyFromText ('SRID=4610 ; POINT (120.163629 30.2592077)' ) ); ";
                    else if (GetGeometryType(entity.comment) == "LINESTRING")
                        //添加字段描述
                        value += "ST_GeographyFromText ('SRID=4610 ;LINESTRING ( 120.163629 30.2592077, 120.163599 30.259802, 120.163617 30.259803, 120.163647 30.259207 )' ) ); ";
                    else if (GetGeometryType(entity.comment) == "POLYGON")
                        //添加字段描述
                        value += "ST_GeographyFromText ('SRID=4610 ;POLYGON ((120.163629 30.259207,120.163599 30.259802,120.163617 30.259803,120.163647 30.259207,120.163629 30.259207))') ) ; ";

                    AddComment addComment = new AddComment();
                    addComment.ExecuteNonQuery(sqlcom + attribute + value);

                    //INSERT INTO public.table_pgtable(name, geometry, remarks) VALUES('点数据', st_GeomFromText('POINT (120.163629 30.2592077)'),'点数据');


                }
                catch (Exception e)
                {

                    throw e;
                }
            }

        }

    
        public static string GetGeometryType(string type)
        {
            switch (type)
            {
                case "点":
                    return "POINT";

                case "线":
                    return "LINESTRING";
                case "面":
                    return "POLYGON";

                default:
                    return "POINT";
            }
        }
    }
}
