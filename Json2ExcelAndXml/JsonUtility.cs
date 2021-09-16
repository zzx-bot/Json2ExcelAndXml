/**********************************************************************
** Copyright © 2020 广州南方智能技术有限公司 All rights reserved.
***********************************************************************
** CLR 版本:4.0.30319.42000
** 创建时间:2020/10/7 14:28:27
** 作    者:LiJing
** 说    明:
** ============================== 修改 ============================== **
** 修改时间:
** 作    者:
** 说    明:
**********************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Newtonsoft.Json;

namespace DrillingBuildLibrary
{
    public static class JsonUtility
    {
        public static void ToJsonFile(object data,string jsonPath)
        {
            try
            {
                string str = JsonConvert.SerializeObject(data, Formatting.Indented);
                File.WriteAllText(jsonPath, str);
            }
            catch(Exception ex)
            {
                throw new Exception($"写入文件出错：{ex.Message}");
            }
        }

        public static T FromJsonFile<T>(string jsonPath)
        {
            if (!File.Exists(jsonPath))
            {
                throw new Exception($"读取的文件不存在：{jsonPath}");
            }

            try
            {
                using (StreamReader reader = new StreamReader(jsonPath))
                {
                    JsonTextReader jsonReader = new JsonTextReader(reader);
                    JsonSerializer jsonSerializer = JsonSerializer.CreateDefault();
                    return jsonSerializer.Deserialize<T>(jsonReader);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"读取文件出错：{ex.Message}");
            }
        }

        public static string JsonEscape(string inputStr)
        {
            string outStr;
            StringBuilder stringBuilder = new StringBuilder();
            for (int i = 0; i < inputStr.Length; i++)
            {
                char c = inputStr[i];
                if (c=='"')
                {
                    stringBuilder.Append('\\');
                }

                stringBuilder.Append(c);

            }
            outStr = stringBuilder.ToString();
            return outStr;
        }

        /// <summary>
        /// 解决读取json文件中文乱码问题
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="jsonPath"></param>
        /// <returns></returns>
        public static T FromJsonFile1<T>(string jsonPath)
        {
            if (!File.Exists(jsonPath))
            {
                throw new Exception($"读取的文件不存在：{jsonPath}");
            }

            try
            {
                using (StreamReader reader = new StreamReader(jsonPath, Encoding.Default))
                {
                    JsonTextReader jsonReader = new JsonTextReader(reader);
                    JsonSerializer jsonSerializer = JsonSerializer.CreateDefault();
                    return jsonSerializer.Deserialize<T>(jsonReader);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"读取文件出错：{ex.Message}");
            }
        }

    }
}
