using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

var xls = MyExcelBook.Create(@"c:\temp\example.xlsx");
xls.CreateSheet("mySheet");
xls.Save();


public sealed class MyExcelBook
{
	private XSSFWorkbook _xssFWorkbook;
	private ISheet _sheet;
	private string _filepath;

	private MyExcelBook()
	{
		_xssFWorkbook = new XSSFWorkbook();//これは単なる初期化という意味しかない。。？ ★★２
	}

	public static MyExcelBook Create(string filepath)//生成する前にCALLできるものなのか。。。
	{
		var obj = new MyExcelBook();
		obj._filepath = filepath;
		obj._xssFWorkbook = new XSSFWorkbook();//コンストラクタでいれてるのでこれがなくても動くのでは。。。
		return obj;
	}

	public void CreateSheet(string name) =>　//ここはラムダ式になっているので注意？⇒これはラムダ式なのか。。。{}が省略されているだけ？？。。。
		_sheet = _xssFWorkbook.CreateSheet(name);

	public void Save()
	{
		using var stream = new FileStream(_filepath, FileMode.Create);//createとは2という意味だけ（列挙型）
																	　//なにが設定されるんだろうか　下の奴がその引数の型をとるからそのように作成しているということ
		_xssFWorkbook.Write(stream);
	}
}

