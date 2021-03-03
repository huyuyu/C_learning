## C_learning


#[using RPA.HyCheck.XamlAnalysis] : 传入要分析的Xaml文件，以及指定标签序列，分析其组件、变量、系统使用，以及操作流。

// 读取组件名称列表 string[][] sub_list;
string sub_path = "D:\\coding\\xnewrpa-release3.4.0\\RPA.Robot.Core\\Common\\SubList.txt";
var sub_list = File.ReadAllLines(sub_path).Select(x => x.Split('|')).ToArray();
// 调用 XmalAnalysis 分析流程源文件
XamlAnalysis xamlInfo = new XamlAnalysis();
Dictionary<string, object> analysis = xamlInfo.start_analyse(str,sub_list);
Hashtable a = xamlInfo.sub_set;
