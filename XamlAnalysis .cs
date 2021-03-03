/*
    create: 2021.3.3 huyu
    info: xaml 文件审核工具
*/

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

// 输出库的命令：csc /t:library /out:"XamlAnalysis.dll" "XamlAnalysis .cs"

namespace RPA.HyCheck
{
    public class XamlAnalysis
    {
        // 用来存放标签列表
        private Hashtable label_set = new Hashtable();
        // 用来存放组件列表['组件名称', '使用数量']
        public Hashtable sub_set = new Hashtable();

        // 用来存放变量列表['变量名称', '变量类型']
        public List<string> variable_list = new List<string>();
        // 用来存放网址、系统'操作对象', '参数']
        public List<string> sys_list = new List<string>();
        // 用来存放实际操作过程
        public List<string> step_list = new List<string>();

        // 用来存放标签队列
        private Queue analyse_queue = new Queue();

        public Dictionary<string, object> start_analyse(string xamlFile, string[][] sublist)
        {
            for (int i = 0; i < sublist.Length; i++) {
                label_set.Add(sublist[i][0], sublist[i][1]);
            }
            analyse_access(read_Xaml(xamlFile));
            Dictionary<string, Object> result = new Dictionary<string, Object>();
            result.Add("variables", variable_list);
            result.Add("subs", sub_set);
            result.Add("sys", sys_list);
            result.Add("steps", step_list);
            return result; 
         }

        private XmlNode read_Xaml(string xamlFile)
        {
            XmlDocument xDoc = new XmlDocument();
            xDoc.LoadXml(xamlFile);
            return xDoc.DocumentElement.ChildNodes[2];
        }

        private void analyse_access(XmlNode xmlNode)
        {
            foreach (XmlNode xn in xmlNode.ChildNodes) {
                string tagName = xn.Name;
                if (tagName.Contains("Variables"))
                {
                    variables_analyse(xn);
                }
                else if (label_set.ContainsKey(tagName)){ 
                    sub_set[tagName] = (sub_set.ContainsKey(tagName))? int.Parse(sub_set[tagName].ToString()) + 1 : 1;

                    if (xn.Attributes["DisplayName"] != null)
                    {
                        step_list.Add(xn.Attributes["DisplayName"].Value);
                    }

                    switch (tagName) {
                        case "rsw:OpenBrowser":
                            sys_list.Add("浏览器：" + xn.Attributes["Url"].Value);
                            break;
                        case "rsp:StartProcess":
                            sys_list.Add("客户端：" + xn.Attributes["ApplicationName"].Value);
                            break;
                        case "rsp:OpenApplication":
                            sys_list.Add("带参程序：" +xn.Attributes["FileName"].Value+",参数：" +xn.Attributes["Arguments"].Value);
                            break;
                        case "rsc1:InvokeCode":
                            sys_list.Add("VB代码行数：" + (Regex.Matches(xn.Attributes["Code"].Value, "\r").Count + 1 )); 
                            break;
                        case "rse:EPExcelApplicationScope":
                            sys_list.Add("EXCEL文档：" + xn.Attributes["ExcelPath"].Value);
                            break;
                        case "rsn:NPOIExcelApplicationScope":
                            sys_list.Add("EXCEL文档：" + xn.Attributes["ExcelPath"].Value);
                            break;
                        default:
                            break;                  
                    }
                }
                analyse_queue.Enqueue(xn);
            }
            analyse_process();
        }

        private void analyse_process() {
            while (analyse_queue.Count>0) {
                XmlNode element = (XmlNode)analyse_queue.Dequeue();
                if ((element.Name == "Sequence" || element.Name == "Flowchart") && element.Attributes["DisplayName"] != null)
                    {
                        step_list.Add("【"+element.Attributes["DisplayName"].Value+ "】");
                    }
                analyse_access(element);
            }
        }

        private void variables_analyse(XmlNode xmlnode)
        {
            foreach (XmlNode xn in xmlnode.ChildNodes)
            {
                if (xn.Name.Contains("Variable")) {
                    string a = xn.Attributes["x:TypeArguments"].Value;
                    string b = xn.Attributes["Name"].Value;
                    variable_list.Add("(" + xn.Attributes["x:TypeArguments"].Value + ")" + xn.Attributes["Name"].Value);
                }
            }
        }
    }
}
