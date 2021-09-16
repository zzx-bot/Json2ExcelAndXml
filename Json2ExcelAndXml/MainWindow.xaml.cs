using DrillingBuildLibrary;
using DrillingBuildLibrary.Model;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
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
                openFileDialog.Filter = "建模文件|*.chnr.json";
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
        #endregion

        #region 导出
        private void ExportModelToExcel(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog openFileDialog = new System.Windows.Forms.FolderBrowserDialog();  //选择文件夹


            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string path = openFileDialog.SelectedPath;
                foreach (var entity in _databaseModel.entities)
                {
                    
                    DataTable dataTable = new DataTable();

                    //找到每个表的前缀分类
                    foreach (var group in _databaseModel.viewGroups)
                        if (group.refEntities.Contains(entity.defKey))
                            dataTable.TableName = group.defKey;
                    //表名
                    dataTable.TableName = dataTable.TableName + "-" + entity.defName;

                    string filePath = System.IO.Path.Combine(path, dataTable.TableName + ".xlsx");

                    //下拉框字典
                    Dictionary<string, string[]> dropDownLists=new Dictionary<string, string[]>();
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
                                if(!dropDownLists.ContainsKey(dict.defKey))
                                    dropDownLists.Add(dict.defKey, vs);
                            }
                        }
                    }
                    NPOIHelper.DataModelTableToExcel(dataTable, filePath, dropDownLists);
                }
            }

        }

        #endregion

        private void ToXml(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog openFolder = new System.Windows.Forms.FolderBrowserDialog();

            if (openFolder.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string path = openFolder.SelectedPath;
                foreach (Entity entity in _databaseModel.entities)
                {
                    //xml文件名前缀
                    string preName="";
                    //找到每个表的前缀分类
                    foreach (var group in _databaseModel.viewGroups)
                        if (group.refEntities.Contains(entity.defKey))
                            preName = group.defKey;

                    if (entity.defName.Contains("/"))
                        entity.defName = entity.defName.Replace("/", "");
                    //表名
                    string xmlPath = System.IO.Path.Combine(path, preName + "-" + entity.defName + ".xml");

                    //实例化模板
                    DrillTemplate template = new DrillTemplate();
                    template.TemplateName=entity.defName;
                    template.TemplateType = preName;
                    template.IsInfoTable = true;

                    template.TemplateContents = new List<DrillTemplateTemplateContent>();

                    for (int i=0;i<entity.fields.Count;i++ )
                    {

                        Field field = entity.fields[i];

                        DrillTemplateTemplateContent item = new DrillTemplateTemplateContent();

                        string dropDownComment="";
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
                        
                        item.FieldName = entity.fields[i].defName;

                        char col = (char)('A' + i);
                        item.FieldAddress = col+"2";

                        item.ValueAddress = template.TemplateName + "!$"+ $"{col}"+"$3:$" +$"{col}"+"$200";

                        item.IsDynamicRange = true;


                        

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
    }
}
