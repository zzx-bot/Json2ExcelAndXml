using FreeSql;
using System;
using System.Collections.Generic;
using System.Text;

namespace SmartGeological.Database.Core
{


    /// <summary>
    /// 数据库实体操作方法
    /// </summary>
    public class FreeSqlHelper
    {
        public static IFreeSql _FreeSqlInstance = null;
        static FreeSqlHelper()
        {
            //初始化实例对象
            _FreeSqlInstance = new FreeSqlBuilder()
                //指定数据库类型以及数据库连接
                .UseConnectionString(DataType.PostgreSQL, "Host=localhost;Port=5432;Username=postgres;Password=PgSQL2021; Database=IODP;Pooling=true;Minimum Pool Size=1")
                //aop监听sql
                .UseMonitorCommand(cmd =>//执行前
                {
                    Console.WriteLine("--------------------------------------------------执行前begin--------------------------------------------------");
                    Console.WriteLine(cmd.CommandText);
                    Console.WriteLine("--------------------------------------------------执行前end--------------------------------------------------");
                }, (cmd, valueString) =>//执行后
                {
                    Console.WriteLine("--------------------------------------------------执行后begin--------------------------------------------------");
                    Console.WriteLine(cmd.CommandText);
                    Console.WriteLine(valueString);
                    Console.WriteLine("--------------------------------------------------执行后end--------------------------------------------------");
                })
                .UseAutoSyncStructure(true)//CodeFirst自动同步将实体同步到数据库结构（开发阶段必备），默认是true，正式环境请改为false
                .Build();//创建实例（官方建议使用单例模式）
        }
    }
}
