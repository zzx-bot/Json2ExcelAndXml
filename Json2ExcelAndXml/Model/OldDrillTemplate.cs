///************************************************************************
//* * Copyright  2020 南方智能. All rights reserved.
//* * --------------------------Origin------------------------------------
//* Depiction: 钻孔模板
//* Author: 叶宜博
//* CDT: 2020/11/4 19:06:09
//* CLR: 4.0.30319.42000
//* Version: 1.0.0.0
//* Note:
//* * --------------------------Refactoring-------------------------------
//* ReWriter:
//* UDT:
//* Desc:
//************************************************************************/

//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Text.RegularExpressions;
//using System.Threading.Tasks;
//using System.Xml.Serialization;
//using SmartGeoLogical.DrillDataTool.DrillTemplateRules;
//using SmartGeoLogical.DrillDataTool.DrillTemplateRules.RulesStruct;
//using SmartGeoLogical.ExcelOperation.Helps;

//namespace SmartGeoLogical.DrillDataTool.TemplateStruct
//{
//    /// <summary>
//    /// 钻孔模板
//    /// </summary>
//    public class DrillTemplate
//    {
//        #region 字段属性

//        /// <summary>
//        /// 模板名称
//        /// </summary>
//        [XmlAttribute]
//        public string TemplateName { get; set; }

//        /// <summary>
//        /// 模板Id字段。非试验信息模板另存名称
//        /// </summary>
//        [XmlAttribute]
//        public string TemplateIdField { get; set; }

//        /// <summary>
//        /// 模板路径
//        /// </summary>
//        [XmlIgnore]
//        public string TemplatePath { get; set; }

//        /// <summary>
//        /// 模板类型
//        /// </summary>
//        [XmlAttribute]
//        public string TemplateType { get; set; }

//        /// <summary>
//        /// 模板描述
//        /// </summary>
//        [XmlElement("Description")]
//        public string TemplateDescription { get; set; }

//        /// <summary>
//        /// 备注
//        /// </summary>
//        [XmlElement("Remarks")]
//        public string TemplateRemarks { get; set; }

//        /// <summary>
//        /// 模板对应的Excel路径
//        /// </summary>
//        [XmlIgnore]
//        public string ExcelTemplatePath => TemplatePath != null ? Path.ChangeExtension(TemplatePath, ".xlsx") : null;

//        /// <summary>
//        /// 是否为信息表
//        /// </summary>
//        [XmlElement("IsInfoTable")]
//        public bool IsInfoTable { get; set; }

//        /// <summary>
//        /// 钻孔模板内容
//        /// </summary>
//        public List<TemplateContent> TemplateContents { get; set; }

//        public List<ContentRelation> ContentRelationShips { get; set; }

//        #endregion

//        #region 构造函数

//        public DrillTemplate()
//        {

//        }

//        public DrillTemplate(string templateName, string templatePath, string templateType)
//        {
//            TemplateName = templateName;
//            TemplatePath = templatePath;
//            TemplateType = templateType;
//        }

//        #endregion
//    }

//    /// <summary>
//    /// 模板内容
//    /// </summary>
//    public class TemplateContent
//    {
//        #region 字段属性

//        /// <summary>
//        /// 钻孔字段名
//        /// </summary>
//        [XmlAttribute]
//        public string FieldName { get; set; }

//        /// <summary>
//        /// 钻孔字段名
//        /// </summary>
//        [XmlIgnore]
//        public string FieldAlasName => _fieldAlasName ?? (_fieldAlasName = GetFieldAlasName());
//        private string _fieldAlasName;

//        /// <summary>
//        /// 钻孔字段名在Excel中对应单元格区域
//        /// </summary>
//        [XmlAttribute]
//        public string FieldAddress { get; set; }

//        /// <summary>
//        /// 钻孔字段值在Excel中对应单元格区域
//        /// </summary>
//        [XmlAttribute]
//        public string ValueAddress { get; set; }

//        /// <summary>
//        /// 关联字段
//        /// </summary>
//        public string RelationField { get; set; }

//        /// <summary>
//        /// 值域地址引用
//        /// </summary>
//        [XmlIgnore]
//        public string ValueRangeReference { get; set; }

//        /// <summary>
//        /// 字段描述
//        /// </summary>
//        public string FiledDescription { get; set; }

//        /// <summary>
//        /// 钻孔字段值
//        /// </summary>
//        [XmlIgnore]
//        public string Value { get; set; }

//        /// <summary>
//        /// 是否为唯一值。判断在一个工程中该值是否唯一，用于减少数据录入工作量
//        /// </summary>
//        [XmlAttribute]
//        public bool IsOnlyValue { get; set; }

//        /// <summary>
//        /// 是否为统计字段
//        /// </summary>
//        [XmlAttribute]
//        public bool IsDynamicRange { get; set; }

//        /// <summary>
//        /// 钻孔字段值的计算规则
//        /// </summary>
//        [XmlAttribute]
//        public string ValueCalculateRule { get; set; }

//        /// <summary>
//        /// 数据验证规则组集合。组间是或运算，组内是与运算
//        /// </summary>
//        [XmlElement]
//        public List<DataVerifyRuleGroup> DataVerifyRuleGroups;

//        /// <summary>
//        /// 备注
//        /// </summary>
//        [XmlElement("Remarks")]
//        public string ContentRemarks { get; set; }
//        #endregion

//        #region 构造函数

//        public TemplateContent()
//        {

//        }

//        public TemplateContent(string fieldName, string fieldAddress = null, string valueAddress = null)
//        {
//            FieldName = fieldName;
//            FieldAddress = fieldAddress;
//            ValueAddress = valueAddress;
//        }

//        #endregion

//        #region 方法

//        /// <summary>
//        /// 获取验证信息
//        /// </summary>
//        /// <param name="verifyValue">待数据验证</param>
//        /// <param name="verifyValueCanNull">待验证的值能否为空</param>
//        /// <returns></returns>
//        public List<KeyValuePair<string, string>> GetVerifyRuleGroupsExpression(object verifyValue, ref bool verifyValueCanNull)
//        {
//            if (DataVerifyRuleGroups == null || DataVerifyRuleGroups.Count <= 0) return null;
//            List<KeyValuePair<string, string>> groupsRuleVerifyResult = new List<KeyValuePair<string, string>>();
//            foreach (DataVerifyRuleGroup verifyRuleGroup in DataVerifyRuleGroups)
//            {
//                string groupExpression = verifyRuleGroup.GetVerifyRuleExpression(verifyValue, ref verifyValueCanNull);
//                if (groupExpression == null) continue;
//                KeyValuePair<string, string> groupRuleVerifyResult = new KeyValuePair<string, string>(groupExpression, verifyRuleGroup.GroupRulesDescription);
//                groupsRuleVerifyResult.Add(groupRuleVerifyResult);
//            }

//            return groupsRuleVerifyResult;
//        }

//        /// <summary>
//        /// 获取字段别名。
//        /// Todo:主要用来生成ExcelNameRange的Name。
//        /// </summary>
//        /// <returns></returns>
//        private string GetFieldAlasName()
//        {
//            #region 重构前
//            //string[] rangeNameArray = FieldName?.Split('（', '(', '【');
//            //if (rangeNameArray == null || rangeNameArray.Length <= 0) return null;
//            //string rangeName = rangeNameArray[0];
//            //for (int i = 1; i < rangeNameArray.Length; i++)
//            //{
//            //    if (rangeNameArray[i][0] > 127)
//            //    {
//            //        rangeName += rangeNameArray[i];
//            //    }
//            //}
//            ////NamedRange名字不能有非法字符。使用正则表达式获取只有 字母，数字，下划线，汉字组成的字符串
//            //MatchCollection matchCollection = Regex.Matches(rangeName, "[a-zA-Z0-9_\u4e00-\u9fa5]+");
//            //string handelName = matchCollection.Cast<Match>().Aggregate("", (current, match) => current + match.Value);

//            ////防止名称用单元格地址命名，加上下划线
//            //bool hasChinese = Regex.Matches(handelName, "[\u4e00-\u9fa5]+").Count > 0;
//            //string namedRangeName = hasChinese ? handelName : $"_{handelName}";
//            //return namedRangeName; 
//            #endregion

//            #region 重构后。

//            //包含单位的字段要剔除单位
//            string prepareFieldName = RemoveUnit(FieldName);

//            return ExcelRangeHelper.ConvertStringToRangeName(prepareFieldName);
//            #endregion
//        }

//        /// <summary>
//        /// 字段中是否包含单位
//        /// </summary>
//        /// <param name="field"></param>
//        /// <returns></returns>
//        private bool HasUnit(string field)
//        {
//            //包含单位的标志为字符串中有一对圆括号()、（），并且以) ）结尾
//            if ((field.Contains("(") && field.Contains(")")) || (field.Contains("（") && field.Contains("）")))
//            {
//                if (field.EndsWith(")") || field.EndsWith("）")) return true;
//            }

//            return false;
//        }

//        /// <summary>
//        /// 移除单位
//        /// </summary>
//        /// <param name="field"></param>
//        /// <returns></returns>
//        private string RemoveUnit(string field)
//        {
//            if (!HasUnit(field)) return field;
//            int unitFlagIndex = field.LastIndexOfAny(new[] { '（', '(' });
//            return field.Substring(0, unitFlagIndex);
//        }
//        #endregion
//    }

//    /// <summary>
//    /// 
//    /// </summary>
//    public class ContentRelation
//    {
//        [XmlAttribute]
//        public string Name { get; set; }
//        [XmlAttribute]
//        public string Key { get; set; }
//        [XmlAttribute]
//        public string Value { get; set; }
//        public ContentRelation() { }
//        public ContentRelation(string name, string key, string value)
//        {
//            this.Name = name;
//            this.Key = key;
//            this.Value = value;
//        }
//    }
//    public enum enumContentRelation
//    {
//        TextMatching
//    }

//}
