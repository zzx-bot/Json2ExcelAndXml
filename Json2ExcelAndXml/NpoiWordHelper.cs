using DrillingBuildLibrary.Model;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DrillingBuildLibrary
{
    public class NpoiWordHelper
    {
        public static List<Dict> GetDictsFromWrod(string wordFile)
        {
            List<Dict> dicts = new List<Dict>();

            try
            {
                Stream stream = File.OpenRead(wordFile);

                XWPFDocument doc = new XWPFDocument(stream);

                var paragraphs = doc.Paragraphs.ToList().Where(r => r.ParagraphText != "").Select(r => r.ParagraphText).ToList();
                for (int i = 0; i < paragraphs.Count; i++)
                {

                    string str = paragraphs[i].Split(' ').Where(r=>r!="").ToList()[1];
                    var array = str.Split(new char[4] { '(', ')', '（', '）' });

                    Dict dict = dicts.Where(r => r.defKey == array[1]).FirstOrDefault();
                    if (dict == null)
                    {
                        dict = new Dict();
                        dict.items = new List<Item>();
                        dict.defKey = array[1];
                        dict.defName = array[0];
                        dicts.Add(dict);
                    }




                    var table = doc.Tables[i];

                    for (int j = 1; j < table.Rows.Count; j++)
                    {
                        Item item = new Item();

                        var row = table.Rows[j];
                        int count = row.GetTableCells().Count;

                        if (count == 3)
                        {
                            item.defKey = row.GetCell(0).GetText();
                            item.defName = row.GetCell(1).GetText();
                            item.intro = row.GetCell(2).GetText();
                        }
                        else if (count == 4)
                        {
                            item.defKey = row.GetCell(2).GetText();
                            item.defName = row.GetCell(1).GetText();
                            item.sort = row.GetCell(0).GetText();

                            item.intro = row.GetCell(3).GetText();
                            item.enabled = true;
                        }

                        dict.items.Add(item);

                    }

                }

                stream.Dispose();
            }
            catch (Exception ex)
            {
                throw ex;
            }
             
            
            return dicts;
        }

        /// <summary>
        /// 原位试验字典获取
        /// </summary>
        /// <param name="wordFile"></param>
        /// <returns></returns>
        public static Dict GetInSituTestDictsFromWrod(string wordFile)
        {

            Stream stream = File.OpenRead(wordFile);

            XWPFDocument doc = new XWPFDocument(stream);

            Dict dict = new Dict();
            dict.items = new List<Item>();

            dict.defKey = "室内试验指标类型";
            dict.defName = "室内试验指标类型";
            var paragraphs = doc.Paragraphs.Where(r => r.Style == "3").ToList();
            List<string> excludeFiled = new List<string>() { "统一编号", "试验编号", "样品编号", "环境温度", "环境湿度", "试验方法","备注", "送样单位", "试验单位", "试验日期", "试验人" };
            for (int i = 0; i < paragraphs.Count; i++)
            {
               

                var table = doc.Tables[i];


                Item parentItem = new Item();
                parentItem.defKey = (i+1).ToString().PadLeft(2,'0');
                parentItem.defName  = paragraphs[i].Text.Replace("结果表", "");
                parentItem.enabled = true;
                dict.items.Add(parentItem);
                int childNum = 1;
                for (int j = 1; j < table.Rows.Count; j++)
                {
                    Item item = new Item();

                    var row = table.Rows[j];
                    string fieldName = row.GetCell(1).GetText();
                    if (excludeFiled.Contains(fieldName))
                    {
                        continue;
                    }
                    string symbol = row.GetCell(7).GetText();

                    item.defKey = parentItem.defKey+ childNum.ToString().PadLeft(2,'0');
                    childNum++;
                    item.defName = fieldName;
                    item.attr1 = symbol;
                    item.parentKey = parentItem.defKey;
                    item.enabled = true;

                    dict.items.Add(item);

                }


            }

            stream.Dispose();

            return dict;
        }
        
        public static  Entity GetEntitiesFromWrod(string wordFile)
        {

            List<Entity> entities = new List<Entity>();
            Entity entity = new Entity();
            entity.fields = new List<Field>();

            Stream stream = File.OpenRead(wordFile);
            XWPFDocument doc = new XWPFDocument(stream);

            List<XWPFParagraph> paragraphs = new List<XWPFParagraph>();
            //var paragraphs = doc.Paragraphs.Select(r => r?.ParagraphText[0]=='表').ToList();
            string[] nowKeys = {"cs", "QXAZFG02", "DLAKZD01", "DLBHLU01", "DLBHPO01", "DLCCZQ01", "DLCJMD01", "DLCSSH01", "DLDJTO01", "DLFXZQ01", "DLFXZX01", "ZYEDYJ01", "DLGDMQ01", "QYCHSY1", "QYCQRY01", "QYDJDS01", "QYDJFB01", "QYDCHF01", "QYDCDW01", "QYDSWC01", "QYDMDC01", "QYDCXC01", "QYDSJC01", "QYDDJX01", "QYDTCJ01", "QYDSSC01", "QYDPMX01", "QYDGZX01", "QYFDCX01", "QYFZZX01", "QYFGZD01", "QYGGQH01", "QYGGDM01", "QYGGHD01", "GCAQYW01", "GCACQW01", "GCADJW01", "GCAJAW01", "GCAGCQ01", "GCADJT01", "GCARTF01", "GCASTF01", "GCARGT01", "GCATYT01", "GCAYMX01", "GCACFD01", "GCACHD01", "GCAYTG01", "GCAYJG01", "GCATGL01", "GCATJL01", "GCADXJ01", "GCAJZJ01", "GCATDJ01", "GCARTP01",
"GCAJSY01", "SWAKCL01", "SWAYSZ01", "SWAYSG01", "SWAXTH01", "SWASYD01", "SWALXH01", "SWAFSH01", "SWASHX01", "SWAHSZ01", "SWADBG01", "SWADBX01", "SWAQSW01", "SWACYS01", "SWAQYZ01", "SWACYZ01", "SWAQDZ01", "SWASDZ01", "SWATZX01", "SWAKHD01", "SWAJSR01", "SWAJSZ01", "SWAGSH01", "SWAHQS01", "SWASTX01", "SWAYLX01", "SWACSX01", "SWAGSD01", "SWAXST01", "SWAWDS01", "SWABZM01", "SWAKZM01", "SWAXKM01", "SWAFKL01", "SWAQLM01", "SWAQLX01", "SWASKL01", "SWAZLQ01", "SWAFSX01", "SWAWRC01", "SWAFWX01", "SWAWZP01", "SWACRX01", "SWADGC01", "SWAHSC01", "SWAGNP01", "SWASWD01", "HJADCQ01", "HJAYZH01", "HJATDW01", "HJAHLW01", "HJAHBW01", "HJAWFX01", "HJAWFZ01", "HJAWZP01", "HJBDZL01", "HJBDCJ01", "HJBLCJ01", "HJBCJD01", "HJBYCD01", "HJBFXD01", "HJBLFY01", "HJBXWQ01", "HJBEFX01", "HJBHYC01", "HJBFZG01", "HJCSTL01", "HJCSMH01", "HJCTLX01", "HJCSDQ01", "HJCDFB01", "HJCHQF01", "HJCHAX01", "HJCBSY01", "HJCDGH01", "HJCSZP01", "WLBHJX01", "WLBZTS01", "WLBZJX01", "WLBZSF01", "WLDDFY01", "WLGDZL01", "WLGTDS01", "HXACMP01", "HXAYHX01", "HXAYYC01", "HXAYZY01", "HXBQTX01", "HXBQTY01", "HXBQZY01", "HXBBCX01", "HXBBCY01", "HXBBZY01", "HXBSYX01", "HXBSYC01", "HXBSZY01", "HXBTWP01", "HXBTSP01", "HXBTZL01", "HXBTYP01", "HXBTFQ01", "HXBTDH01", "HXBTZP01", "HXBFSX01", "HXBSHY01", "HXCHHX01", "HXCHYC01", "HXCHZY01", "HXCHWP01",
"HXEZHX01", "HXEZHY01", "HXEZHZ01", "HXENCP01", "HXENZW01", "HXFSHD01", "HXFSHY01", "HXFSHZ01", "HXFHSF01", "HXFSZL01", "YGBTGZ01", "YGBCFB01", "ZYAKKC01", "ZYAGKQ01", "ZYAGKD01", "ZYAJSK01", "ZYAMTF01", "ZYBSZY01", "ZYBDXS01", "ZYBKQS01", "ZYDRZY01", "ZYDRSW01", "ZYDEKQ01", "ZYDZQL01", "ZYDRZL01", "ZYDQRX01", "ZYDRYC01", "ZYDRTD01", "ZYDRGH01", "ZYDBSY01", "ZYDGSY01", "ZYDBZY01", "ZYDGZY01", "ZYDBPJ01", "ZYDGPJ01", "ZYEDGY01", "ZYFLXZ01", "ZYFSYX01", "ZYFKFG01"};

            List<string> lostKeys=new List<string> ();
            foreach (var para in doc.Paragraphs)
            {
                if (para.ParagraphText.Length > 0 && para.ParagraphText[0] == '表')
                {
                    paragraphs.Add(para);
                    foreach (var key in nowKeys)
                    {
                        if(para.ParagraphText.Contains(key))
                            lostKeys.Add(key);
                    }
                   
                }


            }

            List<string> L = new List<string>();
            foreach (var key in nowKeys)
            {
                if (!lostKeys.Contains(key))
                    L.Add(key);
            }
            


            var tables = doc.Tables;
            var package = doc.Package;

            for (int i = 0; i < tables.Count; i++)
            {
                //var paragraph = paragraphs[i];
                var table = tables[i];
                //if(nowKeys.Contains(table)

                //entity.defKey = paragraph.Text;
                //entity.defName = paragraph.Text;


                for (int j = 1; j < table.Rows.Count; j++)
                {
                    var row = table.Rows[j];
                    Field field = new Field();
                    field.defName = row.GetCell(2).GetText();
                    field.defKey = row.GetCell(3).GetText()=="" ? field.defName: row.GetCell(3).GetText();
                    field.comment = row.GetCell(1).GetText() == "" ? entity.fields[entity.fields.Count-1].comment : row.GetCell(1).GetText();
                    field.type=row.GetCell(5).GetText();
                    entity.fields.Add(field);
                }
                
            }
            stream.Dispose();

            return entity;
        }
        public static List<Entity> GetEntitieTablesFromWrod(string wordFile)
        {
            List<Entity> entities = new List<Entity>();
            try
            {
                Stream stream = File.OpenRead(wordFile);
                XWPFDocument doc = new XWPFDocument(stream);

                //表10 (QYAYDD02)地质调查点基本情况表（续）
                var paragraphs = doc.Paragraphs.Where(r => RegexHelper.IsMatch(r.ParagraphText, @"^表\s{0,2}\d{1,10}")).Select(r => r.ParagraphText).ToList();
                var tables = doc.Tables;

                for (int i = 0; i < paragraphs.Count; i++)
                {
                    var paragraph = paragraphs[i].Trim(); 
                    var array = paragraph.Split(new char[4] { '(', ')', '（', '）' });

                    var table = tables[i];
                    Entity entity = entities.Where(r => r.defKey ==  array[1]).FirstOrDefault();
                    if (entity == null)
                    {
                        entity = new Entity();
                        entity.defKey = array[1];
                        entity.defName = string.Join("", array.ToList().GetRange(2, array.ToList().Count - 2));// ;
                    
                        entity.fields = new List<Field>();
                        entities.Add(entity);

                    }


                    for (int j = 1; j < table.Rows.Count; j++)
                    {

                        var row = table.Rows[j];
                        if (row.GetTableCells().Count == 1)
                        {
                            continue;
                        }

                        Field field = new Field();
                        field.defName = row.GetCell(1).GetText();
                        field.defKey = row.GetCell(2).GetText() == "" ? field.defName : row.GetCell(2).GetText();
                        field.comment = row.GetCell(3).GetText();
                        field.notNull = row.GetCell(5).GetText() == "M" ? true : false;
                        var units = row.GetCell(7).GetText();
                        if (units != "/" && units != "")
                        {
                            field.comment += ";单位：" + units;
                        }
                        string filedType = row.GetCell(4).GetText();

                        if (filedType.StartsWith("C"))
                        {
                            field.type = "VARCHAR";
                            field.len = filedType.Replace("C", "");
                        }
                        else if (filedType.StartsWith("Date"))
                        {
                            field.type = "DATE";
                        }
                        else if (filedType.StartsWith("F"))
                        {
                            field.type = "NUMERIC";
                            field.len = filedType.Replace("F", "").Split('.')[0];
                            field.scale = filedType.Replace("F", "").Split('.')[1];
                        }
                        else if (filedType.StartsWith("N"))
                        {
                            field.type = "INTEGER";
                            field.len = filedType.Replace("N", "");
                        }
                        else
                        {
                            field.type = filedType;
                        }



                        entity.fields.Add(field);
                    }
                }
                stream.Dispose();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            

            return entities;
        }
        public static List<Dict> GetWordDict(string wordFile)
        {

            Stream stream = File.OpenRead(wordFile);

            XWPFDocument doc = new XWPFDocument(stream);

            var paragraphs = doc.Paragraphs.Select(r=>r.ParagraphText).ToList();


            foreach (var item in paragraphs)
            {
                if (item=="")
                {
                    continue;
                }
            }

            stream.Dispose();

            return null;
        }

    }



}
