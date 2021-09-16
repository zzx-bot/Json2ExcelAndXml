using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DrillingBuildLibrary.Model
{

    public class DatabaseModel
    {
        public string name { get; set; }
        public string describe { get; set; }
        public string avatar { get; set; }
        public string version { get; set; }
        public string createdTime { get; set; }
        public string updatedTime { get; set; }
        public object[] dbConns { get; set; }
        public Profile profile { get; set; }
        public List<Entity> entities { get; set; }
        public object[] views { get; set; }
        public List<Dict> dicts { get; set; }
        public Viewgroup[] viewGroups { get; set; }
        public Datatypemapping dataTypeMapping { get; set; }
        public Domain[] domains { get; set; }
        public Diagram[] diagrams { get; set; }
        public Dbconn[] dbConn { get; set; }
        public object[] standardFields { get; set; }
    }

    public class Profile
    {
        public defaultDB _default { get; set; }
        public string javaHome { get; set; }
        public Sql sql { get; set; }
        public string[] dataTypeSupports { get; set; }
        public Codetemplate[] codeTemplates { get; set; }
        public Generatordoc generatorDoc { get; set; }
        public int relationFieldSize { get; set; }
        public Uihint[] uiHint { get; set; }
        public string modelType { get; set; }
    }

    public class Sql
    {
        public string delimiter { get; set; }
    }

    public class Generatordoc
    {
        public string docTemplate { get; set; }
    }

    public class Codetemplate
    {
        public string type { get; set; }
        public string applyFor { get; set; }
        public string createTable { get; set; }
        public string createIndex { get; set; }
        public string content { get; set; }
    }

    public class Uihint
    {
        public string defKey { get; set; }
        public string defName { get; set; }
        public string[] data { get; set; }
    }

    public class Datatypemapping
    {
        public string referURL { get; set; }
        public Mapping[] mappings { get; set; }
    }

    public class Mapping
    {
        public string defKey { get; set; }
        public string defName { get; set; }
        public string JAVA { get; set; }
        public string ORACLE { get; set; }
        public string MYSQL { get; set; }
        public string SQLServer { get; set; }
        public string PostgreSQL { get; set; }
        public string DB2 { get; set; }
    }

    public class Entity
    {
        public string defKey { get; set; }
        public string defName { get; set; }
        public string comment { get; set; }
        public Properties properties { get; set; }
        public string nameTemplate { get; set; }
        public Header[] headers { get; set; }
        public List<Field> fields { get; set; }
        public Correlation[] correlations { get; set; }
        public Index[] indexes { get; set; }
    }

    public class Properties
    {
    }

    public class Header
    {
        public bool freeze { get; set; }
        public string refKey { get; set; }
        public bool hideInGraph { get; set; }
    }

    public class Field
    {
        public string defKey { get; set; }
        public string defName { get; set; }
        public string comment { get; set; }
        public string type { get; set; }
        public object len { get; set; }
        public object scale { get; set; }
        public bool primaryKey { get; set; }
        public bool notNull { get; set; }
        public bool autoIncrement { get; set; }
        public string defaultValue { get; set; }
        public bool hideInGraph { get; set; }
        public string domain { get; set; }
        public string refDict { get; set; }
    }

    public class Correlation
    {
        public string myField { get; set; }
        public string refEntity { get; set; }
        public string refField { get; set; }
        public string myRows { get; set; }
        public string refRows { get; set; }
        public string innerType { get; set; }
    }

    public class Index
    {
        public string defKey { get; set; }
        public object defName { get; set; }
        public bool unique { get; set; }
        public string comment { get; set; }
        public object[] fields { get; set; }
    }

    public class Dict
    {
        public string defKey { get; set; }
        public string defName { get; set; }
        public string sort { get; set; }
        public string intro { get; set; }
        public List<Item> items { get; set; }
    }

    public class Item
    {
        public string defKey { get; set; }
        public string defName { get; set; }
        public string sort { get; set; }
        public string parentKey { get; set; }
        public string intro { get; set; }
        public bool enabled { get; set; }
        public string attr1 { get; set; }
        public string attr2 { get; set; }
        public string attr3 { get; set; }
    }

    public class Viewgroup
    {
        public string defKey { get; set; }
        public string defName { get; set; }
        public string[] refEntities { get; set; }
        public object[] refViews { get; set; }
        public string[] refDiagrams { get; set; }
        public string[] refDicts { get; set; }
    }

    public class Domain
    {
        public string defKey { get; set; }
        public string defName { get; set; }
        public string applyFor { get; set; }
        public object len { get; set; }
        public object scale { get; set; }
        public Uihint1 uiHint { get; set; }
    }

    public class Uihint1
    {
    }

    public class Diagram
    {
        public string defKey { get; set; }
        public string defName { get; set; }
        public Canvasdata canvasData { get; set; }
    }

    public class Canvasdata
    {
        public Cell[] cells { get; set; }
    }

    public class Cell
    {
        public string id { get; set; }
        public string shape { get; set; }
        public Position position { get; set; }
        public int count { get; set; }
        public string originKey { get; set; }
        public Source source { get; set; }
        public Target target { get; set; }
        public string relation { get; set; }
        public string fillColor { get; set; }
        public Vertex[] vertices { get; set; }
        public string label { get; set; }
        public Size size { get; set; }
        public string[] children { get; set; }
        public string parent { get; set; }
        public Label[] labels { get; set; }
    }

    public class Position
    {
        public float x { get; set; }
        public float y { get; set; }
    }

    public class Source
    {
        public string cell { get; set; }
        public string port { get; set; }
    }

    public class Target
    {
        public string cell { get; set; }
        public string port { get; set; }
    }

    public class Size
    {
        public int width { get; set; }
        public int height { get; set; }
    }

    public class Vertex
    {
        public float x { get; set; }
        public float y { get; set; }
    }

    public class Label
    {
        public Attrs attrs { get; set; }
    }

    public class Attrs
    {
        public Text text { get; set; }
        public Rect rect { get; set; }
    }

    public class Text
    {
        public string text { get; set; }
    }

    public class Rect
    {
    }

    public class Dbconn
    {
        public string defKey { get; set; }
        public string defName { get; set; }
        public string type { get; set; }
        public Properties1 properties { get; set; }
    }

    public class Properties1
    {
        public string driver_class_name { get; set; }
        public string url { get; set; }
        public string password { get; set; }
        public string username { get; set; }
    }

    public class defaultDB
    {
        public string db { get; set; }
        public string dbConn { get; set; }
        public Entityinitfield[] entityInitFields { get; set; }
        public Entityinitproperties entityInitProperties { get; set; }
    }

    public class Entityinitproperties
    {
    }

    public class Entityinitfield
    {
        public string defKey { get; set; }
        public string defName { get; set; }
        public string comment { get; set; }
        public string type { get; set; }
        public object len { get; set; }
        public string scale { get; set; }
        public bool primaryKey { get; set; }
        public bool notNull { get; set; }
        public bool autoIncrement { get; set; }
        public string defaultValue { get; set; }
        public bool hideInGraph { get; set; }
        public string domain { get; set; }
    }

}
