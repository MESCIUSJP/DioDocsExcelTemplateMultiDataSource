// See https://aka.ms/new-console-template for more information
using GrapeCity.Documents.Excel;
using System.Data;

Console.WriteLine("テンプレート構文で複数のデータソースを使用して売上管理表を作成する");

// 新規ワークブックの作成
Workbook workbook = new();

// 帳票テンプレートを読み込む
workbook.Open("sales-management.xlsx");

// データソースを追加
#region 注文データ
{
    var datasource = new DataTable();
    datasource.Columns.Add(new DataColumn("oid", typeof(string)));
    datasource.Columns.Add(new DataColumn("cid", typeof(string)));
    datasource.Columns.Add(new DataColumn("pid", typeof(string)));
    datasource.Columns.Add(new DataColumn("count", typeof(double)));
    datasource.Rows.Add("OD00001", "C001", "P001", 3);
    datasource.Rows.Add("OD00002", "C001", "P002", 6);
    datasource.Rows.Add("OD00003", "C001", "P003", 9);
    datasource.Rows.Add("OD00004", "C001", "P004", 2);
    datasource.Rows.Add("OD00005", "C001", "P005", 4);
    datasource.Rows.Add("OD00006", "C002", "P009", 7);
    datasource.Rows.Add("OD00007", "C002", "P004", 1);
    datasource.Rows.Add("OD00008", "C002", "P008", 8);
    datasource.Rows.Add("OD00009", "C002", "P007", 5);
    datasource.Rows.Add("OD00010", "C002", "P006", 4);
    datasource.Rows.Add("OD00011", "C003", "P009", 5);
    datasource.Rows.Add("OD00012", "C003", "P003", 2);
    datasource.Rows.Add("OD00013", "C003", "P002", 1);
    datasource.Rows.Add("OD00014", "C003", "P006", 3);
    datasource.Rows.Add("OD00015", "C004", "P010", 10);
    datasource.Rows.Add("OD00016", "C005", "P008", 9);
    datasource.Rows.Add("OD00017", "C006", "P007", 8);
    datasource.Rows.Add("OD00018", "C007", "P011", 7);
    datasource.Rows.Add("OD00019", "C007", "P012", 4);
    datasource.Rows.Add("OD00020", "C008", "P013", 6);
    datasource.Rows.Add("OD00021", "C008", "P014", 5);
    datasource.Rows.Add("OD00022", "C008", "P015", 2);
    workbook.AddDataSource("order", datasource);
}
#endregion

#region 顧客データ
{
    var datasource = new DataTable();
    datasource.Columns.Add(new DataColumn("cid", typeof(string)));
    datasource.Columns.Add(new DataColumn("name", typeof(string)));
    datasource.Rows.Add("C001", "田中太郎");
    datasource.Rows.Add("C002", "鈴木花子");
    datasource.Rows.Add("C003", "佐藤健一");
    datasource.Rows.Add("C004", "高橋和子");
    datasource.Rows.Add("C005", "伊藤誠二");
    datasource.Rows.Add("C006", "渡辺美奈");
    datasource.Rows.Add("C007", "山田一郎");
    datasource.Rows.Add("C008", "加藤由美");
    workbook.AddDataSource("customer", datasource);
}
#endregion

#region 商品データ
{
    var datasource = new DataTable();
    datasource.Columns.Add(new DataColumn("pid", typeof(string)));
    datasource.Columns.Add(new DataColumn("name", typeof(string)));
    datasource.Columns.Add(new DataColumn("unitprice", typeof(double)));
    datasource.Rows.Add("P001", "りんご", 120);
    datasource.Rows.Add("P002", "バナナ", 150);
    datasource.Rows.Add("P003", "みかん", 80);
    datasource.Rows.Add("P004", "いちご", 300);
    datasource.Rows.Add("P005", "ぶどう", 400);
    datasource.Rows.Add("P006", "トマト", 200);
    datasource.Rows.Add("P007", "レタス", 180);
    datasource.Rows.Add("P008", "キャベツ", 160);
    datasource.Rows.Add("P009", "にんじん", 100);
    datasource.Rows.Add("P010", "ピーマン", 90);
    datasource.Rows.Add("P011", "ほうれん草", 130);
    datasource.Rows.Add("P012", "さつまいも", 250);
    datasource.Rows.Add("P013", "キウイ", 140);
    datasource.Rows.Add("P014", "オレンジ", 200);
    datasource.Rows.Add("P015", "アスパラガス", 350);
    workbook.AddDataSource("product", datasource);
}
#endregion

// 売上管理表を作成
workbook.ProcessTemplate();

// 列幅を自動調整
workbook.Worksheets[0].Range["A:F"].AutoFit();

// Excelファイルに保存
workbook.Save("result.xlsx");