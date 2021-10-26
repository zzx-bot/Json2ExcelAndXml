using DrillingBuildLibrary;
using DrillingBuildLibrary.Model;
using Json2ExcelAndXml.SQL;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Json2ExcelAndXml
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private DatabaseModel _databaseModel;

        private Entity _element;

        private string _Path;

        public MainWindow()
        {
            InitializeComponent();
        }


        #region 导入

        /// <summary>
        /// 读取建模文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ReadbuildlibraryData(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Title = "选择数据库建模文件";
                //openFileDialog.Filter = "建模文件|*.chnr.json";
                openFileDialog.Multiselect = false;
                if (openFileDialog.ShowDialog() == true)
                {
                    _Path = openFileDialog.FileName;
                    _databaseModel = JsonUtility.FromJsonFile<DatabaseModel>(openFileDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ImportFeatureClassFromWord(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Title = "实体表word文件";
                openFileDialog.Filter = "实体表word文件|*.docx";
                openFileDialog.Multiselect = false;
                if (openFileDialog.ShowDialog() == true)
                {
                    _element = NpoiWordHelper.GetEntitiesFromWrod(openFileDialog.FileName);
                    //foreach (var item in entitys)
                    //{
                    //    item.headers = _databaseModel.entities[0].headers;
                    //    item.properties = _databaseModel.entities[0].properties;
                    //    item.nameTemplate = _databaseModel.entities[0].nameTemplate;
                    //    item.correlations = _databaseModel.entities[0].correlations;
                    //    item.indexes = _databaseModel.entities[0].indexes;
                    //    if (_databaseModel.entities.Contains(item))
                    //    {
                    //        int index = _databaseModel.entities.IndexOf(item);
                    //        _databaseModel.entities[index] = item;
                    //    }
                    //    else
                    //    {
                    //        _databaseModel.entities.Add(item);
                    //    }
                    //}
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        #endregion

        #region 导出
        private void ExportModelToExcel(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog openFileDialog = new System.Windows.Forms.FolderBrowserDialog();  //选择文件夹


            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                foreach (var group in _databaseModel.viewGroups)
                {
                    //新建文件夹
                    string path = openFileDialog.SelectedPath;
                    path = System.IO.Path.Combine(path, group.defKey + "模板");
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                }

                string pathRoot = openFileDialog.SelectedPath;


                foreach (var entity in _databaseModel.entities)
                {

                    DataTable dataTable = new DataTable();

                    //前缀名；excel文件路径
                    string preName = "", filePath;
                    //找到每个表的前缀分类
                    foreach (var group in _databaseModel.viewGroups)
                        if (group.refEntities.Contains(entity.defKey))
                            preName = group.defKey + "模板";

                    if (entity.defName.Contains("/"))
                        entity.defName = entity.defName.Replace("/", "");

                    //表名
                    dataTable.TableName = preName + "-" + entity.defName;

                    if (entity.defName == "地热钻孔地层岩性表")
                        ;
                    if (entity.defName.Contains("反射地震法"))
                        ;
                    if (entity.defName.Contains("生物化学分析表"))
                        ;
                    filePath = System.IO.Path.Combine(pathRoot, preName, dataTable.TableName + ".xlsx");

                    //下拉框字典
                    Dictionary<string, string[]> dropDownLists = new Dictionary<string, string[]>();
                    foreach (var field in entity.fields)
                    {
                        if (!dataTable.Columns.Contains(field.defKey + "-" + field.defName))
                        {
                            //dataTable.Columns.Contains();

                            dataTable.Columns.Add(new DataColumn(field.defKey + "-" + field.defName));

                            if (_databaseModel.dicts.Exists(d => d.defKey == field.defKey))
                            {
                                Dict dict;//
                                dict = _databaseModel.dicts.Find(d => d.defKey == field.defKey);

                                string[] vs = new string[dict.items.Count];
                                //读取所有的下拉框值
                                for (int i = 0; i < dict.items.Count; i++)
                                {
                                    vs[i] = (dict.items[i].defKey + dict.items[i].defName);
                                }
                                if (!dropDownLists.ContainsKey(dict.defKey))
                                    dropDownLists.Add(dict.defKey, vs);
                            }
                        }
                    }
                    NPOIHelper.DataModelTableToExcel(dataTable, filePath, dropDownLists);
                }
            }

        }



        #endregion

        private void ExportModelToXml(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog openFolder = new System.Windows.Forms.FolderBrowserDialog();

            //为每个模板新建文件夹
            foreach (var group in _databaseModel.viewGroups)
            {
                //新建文件夹
                string path = openFolder.SelectedPath;
                path = System.IO.Path.Combine(path, group.defKey + "模板");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
            }


            if (openFolder.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string path = openFolder.SelectedPath;
                foreach (Entity entity in _databaseModel.entities)
                {
                    //xml文件名前缀
                    string preName = "";
                    //找到每个表的前缀分类
                    foreach (var group in _databaseModel.viewGroups)
                        if (group.refEntities.Contains(entity.defKey))
                            preName = group.defKey + "模板";

                    if (entity.defName.Contains("/"))
                        entity.defName = entity.defName.Replace("/", "");
                    //表名
                    string xmlPath = System.IO.Path.Combine(path, preName, preName + "-" + entity.defName + ".xml");

                    //实例化模板
                    DrillTemplate template = new DrillTemplate();
                    template.TemplateKey = entity.defKey;
                    template.TemplateName = entity.defName;
                    template.TemplateType = preName;
                    template.IsInfoTable = true;

                    template.TemplateContents = new List<DrillTemplateTemplateContent>();

                    for (int i = 0; i < entity.fields.Count; i++)
                    {

                        Field field = entity.fields[i];

                        DrillTemplateTemplateContent item = new DrillTemplateTemplateContent();

                        string dropDownComment = "";
                        if (_databaseModel.dicts.Exists(d => d.defKey == field.defKey))
                        {
                            Dict dict;//
                            dict = _databaseModel.dicts.Find(d => d.defKey == field.defKey);

                            if (dict.items.Count > 0)
                            {
                                dropDownComment = dict.items[0].defName;
                                //读取所有的下拉框值
                                for (int itemsIndex = 1; itemsIndex < dict.items.Count; itemsIndex++)
                                    dropDownComment = dropDownComment + "、" + dict.items[itemsIndex].defName;
                            }

                            item.FiledDescription = "根据实际情况，填写：" + dropDownComment + "等";
                        }
                        else
                        {
                            item.FiledDescription = field.comment;
                        }

                        item.FieldKey = entity.fields[i].defKey;

                        item.FieldName = entity.fields[i].defName;

                        char col = (char)('A' + i);
                        item.FieldAddress = col + "2";

                        item.ValueAddress = template.TemplateName + "!$" + $"{col}" + "$3:$" + $"{col}" + "$200";

                        item.IsDynamicRange = true;



                        if (item.FieldName != "统一编号")
                            item.RelationField = "统一编号";

                        item.DataVerifyRuleGroups = new List<TemplateContentDataVerifyRules>();


                        if (field.notNull)
                        {
                            TemplateContentDataVerifyRules rule1 = new TemplateContentDataVerifyRules();
                            rule1.RuleName = "非空";
                            rule1.Description = "字段值非空";
                            item.DataVerifyRuleGroups.Add(rule1);
                        }

                        if (field.type.Equals("DATE"))
                        {
                            TemplateContentDataVerifyRules rule1 = new TemplateContentDataVerifyRules();
                            rule1.RuleName = "日期时间类型";
                            rule1.Description = "字段值类型为日期时间类型。格式为xxxx-xx-xx或xxxx/xx/xx";
                            item.DataVerifyRuleGroups.Add(rule1);
                        }

                        else if (field.type.Equals("VARCHAR"))
                        {
                            TemplateContentDataVerifyRules rule1 = new TemplateContentDataVerifyRules();
                            rule1.RuleName = "自由文本类型";
                            rule1.Description = "字段值类型为自由文本类型";
                            //rule1.ParamExpression = Convert.ToInt32(field.len.ToString());
                            item.DataVerifyRuleGroups.Add(rule1);
                        }

                        else if (field.type.Equals("NUMERIC"))
                        {
                            TemplateContentDataVerifyRules rule1 = new TemplateContentDataVerifyRules();
                            rule1.RuleName = "小数位数";
                            rule1.Description = "字段值类型为数值型，且小数点默认保留2位";
                            rule1.ParamExpression = new DrillTemplateTemplateContentDataVerifyRulesRuleParams();
                            rule1.ParamExpression.ParamExpression = field.scale.ToString();
                            item.DataVerifyRuleGroups.Add(rule1);
                        }

                        else if (field.type.Equals("BLOB"))
                        {
                            TemplateContentDataVerifyRules rule1 = new TemplateContentDataVerifyRules();
                            rule1.RuleName = "二进制大对象类型";
                            rule1.Description = "字段值类型为二进制文件类型";
                            item.DataVerifyRuleGroups.Add(rule1);
                        }

                        else if (field.type.Equals("INTEGER"))
                        {
                            TemplateContentDataVerifyRules rule1 = new TemplateContentDataVerifyRules();
                            rule1.RuleName = "整型类型";
                            rule1.Description = "字段值类型为整型类型";
                            item.DataVerifyRuleGroups.Add(rule1);
                        }

                        else
                        {
                            ;
                        }

                        template.TemplateContents.Add(item);



                    }
                    ToXML.SerialzeToXML(xmlPath, template);
                }
            }

        }

        #region  导出shp

        private void ExportModelToShp(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog openFileDialog = new System.Windows.Forms.FolderBrowserDialog();  //选择文件夹



            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string shpRootPath = openFileDialog.SelectedPath;
                //GdalUtilities gdalUtilities = new GdalUtilities();
                //gdalUtilities.convertJsonToShapeFile(_Path, shpPath);

                //string fDpath;
                //foreach (var field in _element.fields)
                //{
                //    //新建文件夹
                //    fDpath = openFileDialog.SelectedPath;
                //    fDpath = System.IO.Path.Combine(fDpath, field.comment);
                //    if (!Directory.Exists(fDpath))
                //    {
                //        Directory.CreateDirectory(fDpath);
                //    }
                //}

                List<string> allTypes = new List<string>();
                foreach (var entity in _databaseModel.entities)
                {

                    DataTable dataTable = new DataTable();

                    //前缀名；excel文件路径
                    string shpFilePath, fDpath;

                    // 去除"/"
                    //if (entity.defName.Contains("表"))
                    //    entity.defName = entity.defName.Replace("表", "");



                    /* 读取缺失的shp */
                    //flds.Remove(flds.Find(t => t.defKey == entity.defKey));


                    string rootClass="";
                    foreach (var viewGroup in _databaseModel.viewGroups)
                    {
                        foreach (var item in viewGroup.refEntities)
                        {
                            if (item.Equals(entity.defKey))
                                rootClass = viewGroup.defKey;
                        }

                    }

                    fDpath = openFileDialog.SelectedPath;

                    fDpath = System.IO.Path.Combine(fDpath, rootClass, entity.defName);
                    if (!Directory.Exists(fDpath))
                    {
                        Directory.CreateDirectory(fDpath);

                    }


                    allTypes.Add("");
                    GdalUtilities gdal = new GdalUtilities();

                    for ( int index=0;index< entity.fields.Count;index++)
                    {
                        int c;
                        for (c=0;c< allTypes.Count;c++)
                        {
                            if (entity.fields[index].type.Equals(allTypes[c]))
                            {
                                break;
                            }

                        }
                        if (c == allTypes.Count)
                            allTypes.Add(entity.fields[index].type);
                    }

                    //gdal.convertJsonToShapeFile(entity, fDpath);

                    //if( GetGeometryType(entity.comment)!=null)
                    //allTypes.Add(entity.comment);


                    InputPgDB.CreatSpatialTable(entity);
                    //InputPgDB.InsertSpatialTable(entity);
                    //ToShp.WriteVectorFileShp(shpRootPath, entity);
                }

                ///* 读取缺失的shp */
                //FileInfo myFile = new FileInfo(@"C:\Users\210320\Desktop\1.txt");
                //StreamWriter sW = myFile.CreateText();

                //foreach (var s in flds)
                //{
                //    sW.WriteLine(s.comment + " " + s.defKey + " " + s.defName);
                //}
                //sW.Close();

            }





        }
        #endregion

        public static string GetGeometryType(string type)
        {
            switch (type)
            {
                case "点":
                    return null;

                case "线":
                    return null;
                case "面":
                    return null;

                default:
                    return type;
            }
        }

    }
}
