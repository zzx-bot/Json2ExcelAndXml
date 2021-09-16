using System;
using System.Collections.Generic;


namespace DrillingBuildLibrary.Model
{


    // 注意: 生成的代码可能至少需要 .NET Framework 4.5 或 .NET Core/Standard 2.0。
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "", IsNullable = false)]
    public partial class DrillTemplate
    {

        private bool isInfoTableField;

        private List<DrillTemplateTemplateContent> templateContentsField;

        private object contentRelationShipsField;

        private string templateNameField;

        private string templateTypeField;

        /// <remarks/>
        public bool IsInfoTable
        {
            get
            {
                return this.isInfoTableField;
            }
            set
            {
                this.isInfoTableField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("TemplateContent", IsNullable = false)]
        public List<DrillTemplateTemplateContent> TemplateContents
        {
            get
            {
                return this.templateContentsField;
            }
            set
            {
                this.templateContentsField = value;
            }
        }

        /// <remarks/>
        public object ContentRelationShips
        {
            get
            {
                return this.contentRelationShipsField;
            }
            set
            {
                this.contentRelationShipsField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string TemplateName
        {
            get
            {
                return this.templateNameField;
            }
            set
            {
                this.templateNameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string TemplateType
        {
            get
            {
                return this.templateTypeField;
            }
            set
            {
                this.templateTypeField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class DrillTemplateTemplateContent
    {

        //private DataVerifyRuleGroups dataVerifyRuleGroupsField;

        private List<TemplateContentDataVerifyRules> dataVerifyRuleGroupsField;

        private string relationFieldField;

        private string filedDescriptionField;

        private string fieldNameField;

        private string fieldAddressField;

        private string valueAddressField;

        private bool isOnlyValueField;

        private bool isDynamicRangeField;

        /// <remarks/>
        //public DataVerifyRuleGroups DataVerifyRuleGroups2
        //{
        //    get
        //    {
        //        return this.dataVerifyRuleGroupsField;
        //    }
        //    set
        //    {
        //        this.dataVerifyRuleGroupsField = value;
        //    }
        //}


        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("DataVerifyRules", IsNullable = false)]
        public List<TemplateContentDataVerifyRules> DataVerifyRuleGroups
        {
            get
            {
                return this.dataVerifyRuleGroupsField;
            }
            set
            {
                this.dataVerifyRuleGroupsField = value;
            }
        }


        /// <remarks/>
        public string RelationField
        {
            get
            {
                return this.relationFieldField;
            }
            set
            {
                this.relationFieldField = value;
            }
        }

        /// <remarks/>
        public string FiledDescription
        {
            get
            {
                return this.filedDescriptionField;
            }
            set
            {
                this.filedDescriptionField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string FieldName
        {
            get
            {
                return this.fieldNameField;
            }
            set
            {
                this.fieldNameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string FieldAddress
        {
            get
            {
                return this.fieldAddressField;
            }
            set
            {
                this.fieldAddressField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string ValueAddress
        {
            get
            {
                return this.valueAddressField;
            }
            set
            {
                this.valueAddressField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool IsOnlyValue
        {
            get
            {
                return this.isOnlyValueField;
            }
            set
            {
                this.isOnlyValueField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool IsDynamicRange
        {
            get
            {
                return this.isDynamicRangeField;
            }
            set
            {
                this.isDynamicRangeField = value;
            }
        }
    }

    public class TemplateContentDataVerifyRules
    {
        private DrillTemplateTemplateContentDataVerifyRulesRuleParams ruleParamsField;

        private string ruleNameField;

        private string descriptionField;

        /// <remarks/>
        public DrillTemplateTemplateContentDataVerifyRulesRuleParams ParamExpression
        {
            get
            {
                return this.ruleParamsField;
            }
            set
            {
                this.ruleParamsField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string RuleName
        {
            get
            {
                return this.ruleNameField;
            }
            set
            {
                this.ruleNameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string Description
        {
            get
            {
                return this.descriptionField;
            }
            set
            {
                this.descriptionField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class DrillTemplateTemplateContentDataVerifyRulesRuleParams
    {

        private string paramExpressionField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string ParamExpression
        {
            get
            {
                return this.paramExpressionField;
            }
            set
            {
                this.paramExpressionField = value;
            }
        }
    }


}
