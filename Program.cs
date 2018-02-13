/*
 * Сделано в SharpDevelop.
 * Пользователь: artem279 (Артём Агаев)
 * Дата: 14.04.2016
 * Время: 10:05
 * 
 * Для изменения этого шаблона используйте Сервис | Настройка | Кодирование | Правка стандартных заголовков.
 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using banks_dictonary.CreditInfoWS;
using System.Xml.Linq;
using System.Net;
using System.Net.Cache;
using System.Threading;

namespace banks_dictonary
{
	class Program
	{
		//Создаём объект (веб-ссылку), сервис данных о КО ЦБРФ
		static CreditOrgInfo ws = new CreditOrgInfo();
		
		
		public struct Form101
		{
			public string IndID;
			public string IndCode;
			public string name;
			public string IndType;
			public string IndChapter;
			public string RegNum;
			public string Date;
			public string pln;
			public string ap;
			public string vr;
			public string vv;
			public string vitg;
			public string ora;
			public string ova;
			public string oitga;
			public string orp;
			public string ovp;
			public string oitgp;
			public string ir;
			public string iv;
			public string iitg;
			public string value;
		}
		
		public struct Form102
		{
			public string symid;
			public string symsort;
			public string symbol;
			public string symtype;
			public string name;
			public string symset;
			public string RegNum;
			public string Date;
			public string value;
		}
		
		public struct Form123
		{
			public string IndCode;
			public string name;
			public string RegNum;
			public string Date;
			public string value;
		}
		
		public struct license
		{
			public string RegNum;
			public string LCode;
			public string name;
			public string Date;
		}
		
		//Структура таблицы-справочника КО
		public struct bank
		{
			public double IntCode;
			public string RegNum;
			public string Bic;
			public string ogrn;
			public string Name;
			public string OrgFullName;
			public string MainDateReg;
			public string DateKGRRegistration;
			public string SSV_Date;
			public string UstavAdr;
			public string FactAdr;
			public string Director;
			public string UstMoney;
			public string OrgStatus;
			public string RegCode;
			public string phones;
			public List<license> licenses;
		}
		
		public struct bank_childs
		{
			public double IntCode;
			public string RegNum;
			public string Bic;
			public string ogrn;
			public string Name;
			public string OrgFullName;
			public string cregnum;
			public string cname;
			public string cndate;
			public string straddrmn;
			public string RegId;
		}
		
		
		public class SymbolComparer : IEqualityComparer<Form102>
		{

    		public bool Equals(Form102 b1, Form102 b2)
    		{
        		if (b1.symbol == b2.symbol)
        		{
            		return true;
        		}
        		return false;
    		}


    		public int GetHashCode(Form102 bx)
    		{
        		return bx.symbol.GetHashCode();
    		}
		
		}
		
		
		public static string DownloadCO(string url)
       	{
           	// открываем соединение
           	
           	HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

			request.UserAgent = "User-Agent	YaBrowser/16.3.0.7146";
			
			request.Referer = url;
			
			request.AllowAutoRedirect = true;
			
			request.UseDefaultCredentials = true;
			
			HttpRequestCachePolicy noCache = new HttpRequestCachePolicy(HttpRequestCacheLevel.NoCacheNoStore);
			
		   	request.CachePolicy = noCache;
            
            // Задаём тип содержимого
            request.ContentType = "text";
           
           	// Получаем ответ
           	HttpWebResponse Response = (HttpWebResponse)request.GetResponse();
           	Stream dataStream = Response.GetResponseStream();
           	
           	Console.WriteLine (((HttpWebResponse)Response).StatusDescription);

           	StreamReader Reader = new StreamReader(dataStream,Encoding.Default);

           	// читаем полностью поток
           	string PageContent = Reader.ReadToEnd();

       		
           	// убираем за собой!
           	Reader.Close();
           	dataStream.Close();
           	Response.Close();
			
           	return PageContent;
       	}
		
		public static void Main(string[] args)
		{
			
			
			//указываем урл сервиса
			ws.Url = "http://www.cbr.ru/CreditInfoWebServ/CreditOrgInfo.asmx?WSDL";
			
			//урл справочника БИК КО
			string url = "http://cbr.ru/scripts/XML_bic2.asp";
			
			//Грузим справочники по формам 101 и 102
			XDocument cls101 = XDocument.Parse(ws.Form101IndicatorsEnumXML().OuterXml);
			//cls101.Save("cls101.xml");
			XDocument cls102 = XDocument.Parse(ws.Form102IndicatorsEnumXML().OuterXml);
			
			Thread.Sleep(1500);
			//Грузим список КО
			XDocument doc = XDocument.Parse(DownloadCO(url));
			
			List<bank> banks = new List<bank>();
			List<bank_childs> cOffices = new List<bank_childs>();
			List<Form101> form101 = new List<Form101>();
			List<Form102> form102 = new List<Form102>();
			List<Form123> form123 = new List<Form123>();
			
			//"Поехали!" © Юрий гагарин
			foreach(XElement e in doc.Root.Elements("Record"))
			{
				Console.WriteLine("БИК: {0} ОГРН: {1}",e.Element("Bic").Value,e.Element("RegNum").Value);
				
				bank BankCard = new bank();
				
				try { BankCard.Bic = e.Element("Bic").Value; } catch { BankCard.Bic = ""; }
				try { BankCard.ogrn = e.Element("RegNum").Value; } catch { BankCard.ogrn = ""; }
				try { BankCard.IntCode = ws.BicToIntCode(BankCard.Bic); } catch { BankCard.IntCode = 0; }
				
				Thread.Sleep(1500);
				//грузим карточку КО
				XDocument CreditInfo = null;
				while (CreditInfo == null)
				{
					try
					{
						CreditInfo = XDocument.Parse(ws.CreditInfoByIntCodeXML(BankCard.IntCode).OuterXml);
					}
					catch
					{
						Console.WriteLine("Trying connection...");
					}
				}
				try
				{
					CreditInfo = XDocument.Parse(ws.CreditInfoByIntCodeXML(BankCard.IntCode).OuterXml);
				}
				catch
				{
					CreditInfo = XDocument.Parse(ws.CreditInfoByIntCodeXML(BankCard.IntCode).OuterXml);
				}
				
				
				//Заполняем остальные поля
				try { BankCard.RegNum = CreditInfo.Root.Element("CO").Element("RegNumber").Value; } catch { BankCard.RegNum = ""; }
				try { BankCard.Name = CreditInfo.Root.Element("CO").Element("OrgName").Value; } catch { BankCard.Name = ""; }
				try { BankCard.OrgFullName = CreditInfo.Root.Element("CO").Element("OrgFullName").Value; } catch { BankCard.OrgFullName = ""; }
				try { BankCard.phones = CreditInfo.Root.Element("CO").Element("phones").Value; } catch { BankCard.phones = ""; }
				try { BankCard.DateKGRRegistration = CreditInfo.Root.Element("CO").Element("DateKGRRegistration").Value; } catch { BankCard.DateKGRRegistration = ""; }
				try { BankCard.MainDateReg = CreditInfo.Root.Element("CO").Element("MainDateReg").Value; } catch { BankCard.MainDateReg = ""; }
				try { BankCard.UstavAdr = CreditInfo.Root.Element("CO").Element("UstavAdr").Value; } catch { BankCard.UstavAdr = ""; }
				try { BankCard.FactAdr = CreditInfo.Root.Element("CO").Element("FactAdr").Value; } catch { BankCard.FactAdr = ""; }
				try { BankCard.Director = CreditInfo.Root.Element("CO").Element("Director").Value; } catch { BankCard.Director = ""; }
				try { BankCard.UstMoney = CreditInfo.Root.Element("CO").Element("UstMoney").Value; } catch { BankCard.UstMoney = ""; }
				try { BankCard.OrgStatus = CreditInfo.Root.Element("CO").Element("OrgStatus").Value; } catch { BankCard.OrgStatus = ""; }
				try { BankCard.RegCode = CreditInfo.Root.Element("CO").Element("RegCode").Value; } catch { BankCard.RegCode = ""; }
				try { BankCard.SSV_Date = CreditInfo.Root.Element("CO").Element("SSV_Date").Value; } catch { BankCard.SSV_Date = ""; }
				
				//офисы/филиалы
				XDocument offices = XDocument.Parse(ws.GetOfficesXML(BankCard.IntCode).OuterXml);
				
				foreach(XElement o in offices.Root.Elements("Offices"))
				{
					bank_childs cOffice = new bank_childs();
					try { cOffice.IntCode = BankCard.IntCode; } catch { cOffice.IntCode = 0; }
					try { cOffice.RegNum = BankCard.RegNum; } catch { cOffice.RegNum = ""; }
					try { cOffice.Bic = BankCard.Bic; } catch { cOffice.Bic = ""; }
					try { cOffice.ogrn = BankCard.ogrn; } catch { cOffice.ogrn = ""; }
					try { cOffice.Name = BankCard.Name; } catch { cOffice.Name = ""; }
					try { cOffice.OrgFullName = BankCard.OrgFullName; } catch { cOffice.OrgFullName = ""; }
					try { cOffice.cregnum = o.Element("cregnum").Value; } catch { cOffice.cregnum = ""; }
					try { cOffice.cname = o.Element("cname").Value; } catch { cOffice.cname = ""; }
					try { cOffice.cndate = o.Element("cndate").Value; } catch { cOffice.cndate = ""; }
					try { cOffice.straddrmn = o.Element("straddrmn").Value; } catch { cOffice.straddrmn = ""; }
					try { cOffice.RegId = o.Element("RegId").Value; } catch { cOffice.RegId = ""; }
					
					cOffices.Add(cOffice);
					
				}
				
				
				//заполняем Licenses
				foreach (XElement l in CreditInfo.Root.Elements("LIC"))
				{
					license lic = new license();
					BankCard.licenses = new List<license>();
					try { lic.RegNum = BankCard.RegNum; } catch { lic.RegNum = ""; }
					try { lic.LCode = l.Element("LCode").Value; } catch { lic.LCode = ""; }
					try { lic.name = l.Element("LT").Value; } catch { lic.name = ""; }
					try { lic.Date = l.Element("LDate").Value; } catch { lic.Date = ""; }
					BankCard.licenses.Add(lic);
				}
				
				
				banks.Add(BankCard);
				
				
			}
			
			//заполняем массив рег. номеров КО, для получения общей отчетности по всем КО
			object [] CredOrgNumbers = new object[banks.Count];
			for(int i = 0; i < banks.Count; i++)
			{
				Console.WriteLine("{0} {1} {2} {3}", banks[i].RegNum,banks[i].IntCode,banks[i].Bic,banks[i].Name);
				CredOrgNumbers[i] = banks[i].RegNum;
			}
			
			//Сбрасываем в файл справочник КО
			StreamWriter writer = new StreamWriter("cls_CreditOrgsInfo.csv",false,Encoding.Default);
			StreamWriter writer_licenses = new StreamWriter("cls_CreditOrgsInfo_licenses.csv",false,Encoding.Default);
			StreamWriter writer_offices = new StreamWriter("cls_CreditOrgsInfo_offices.csv",false,Encoding.Default);
			foreach (bank b in banks)
			{
				Console.WriteLine("{0} {1} {2} {3}", b.RegNum,b.IntCode,b.Bic,b.Name);
				writer.WriteLine(b.IntCode+";"+b.RegNum.Trim()+";"+b.Bic.Trim()+";"+b.ogrn.Trim()+";"+b.Name.Trim()+";"+b.OrgFullName.Trim()+";"+b.MainDateReg.Trim()+";"+
				                 b.DateKGRRegistration.Trim()+";"+b.SSV_Date.Trim()+";"+b.UstavAdr.Trim()+";"+b.FactAdr.Trim()+";"+b.Director.Trim()+";"+b.UstMoney.Trim()+";"+
				                 b.OrgStatus.Trim()+";"+b.RegCode.Trim()+";"+b.phones.Trim()
				                );
				
				try
				{
					foreach(bank_childs c in cOffices)
					{
						writer_offices.WriteLine(c.IntCode+";"+c.RegNum.Trim()+";"+c.Bic.Trim()+";"+c.ogrn.Trim()+";"+c.Name.Trim()+";"+c.OrgFullName.Trim()+";"+
						                         c.cregnum.Trim()+";"+c.cname.Trim()+";"+c.cndate.Trim()+";"+
						                         c.straddrmn.Replace("\n","").Replace("\r","").Replace(";",",").Trim()+";"+c.RegId.Trim()
						                        );
						writer_offices.Flush();
					}
				} catch { Console.WriteLine("Нет офисов?!"); }
				
				try
				{
					foreach (license l in b.licenses)
					{
						writer_licenses.WriteLine(l.RegNum.Trim()+";"+l.LCode.Trim()+";"+l.name.Trim()+";"+l.Date.Trim());
						writer_licenses.Flush();
					}
					writer.Flush();
				} catch { Console.WriteLine("Нет лицензий?!"); }
			}
			writer.Close();
			writer_licenses.Close();
			writer.Dispose();
			writer_licenses.Dispose();
			writer_offices.Close();
			writer_offices.Dispose();
			
			//Собираем отчетность по 123 форме для всех КО
			writer = new StreamWriter("cls_CreditOrgsInfo_form123.csv",false,Encoding.Default);
			foreach(bank b in banks)
			{
				int RegNum = 0;
				try { RegNum = Convert.ToInt16(b.RegNum); } catch { RegNum = 0; }
				DateTime[] dates = null;
				while (dates == null)
				{
					try
					{
						dates = ws.GetDatesForF123(RegNum);
					}
					catch { Console.WriteLine("Trying connection..."); }
					
				}
				
				foreach(DateTime date in dates.Where(d=>DateTime.Compare(d, DateTime.Parse("01.10.2017")) > 0))
				{
					XDocument Data123Full = null;
					Thread.Sleep(1500);
					try 
					{
						Data123Full = XDocument.Parse(ws.Data123FormFullXML(RegNum, date).OuterXml);
					}
					catch
					{
						File.AppendAllText("log123.txt", date.ToString()+" | " + b.RegNum + Environment.NewLine);
						Console.WriteLine("oops");
						break;
					}
					//form123
					Form123 f = new Form123();
					
					foreach (XElement e in Data123Full.Root.Elements("F123"))
					{
					
						try { f.RegNum = b.RegNum; } catch { f.RegNum = ""; } //Регистрацонный номер КО
						
						try { f.Date = date.ToString("yyyy-MM-dd"); } catch { f.Date = ""; } //Дата (значения показателей привязаны к дате)
						
						try { f.IndCode = e.Element("CODE").Value; } catch { f.IndCode = ""; } //Номер счета индикатора/код показателя
						
						try { f.name = e.Element("NAME").Value; } catch { f.RegNum = ""; } //Название счетов баланса индикатора
						
						try { f.value = e.Element("VALUE").Value; } catch { f.value = ""; } //значение показателя
						
						Console.WriteLine("{0} {1} {2}", "Форма 123", f.RegNum, f.IndCode);
						
						writer.WriteLine(f.IndCode+";"+f.name.Trim().Replace(";","").Replace("\r"," ").Replace("\n","")+";"+f.Date+";"+f.RegNum+";"+f.value);
						writer.Flush();
						
					}
					
				}
			}
			writer.Close();
			//Собираем отчетность по 101 форме для всех КО
			//Сбрасываем в файл отчетность КО 101 формы
			writer = new StreamWriter("cls_CreditOrgsInfo_form101.csv",false,Encoding.Default);
			foreach (XElement e in cls101.Root.Elements("EIND").Where(a=>a.Element("IndCode") != null))
			{
				Console.WriteLine("Грузим данные по 101 форме...");
				string IndCode = "";
				try {IndCode = e.Element("IndCode").Value; } catch {IndCode = "";}
				XDocument Data101Full = null;
				while(Data101Full == null)
				{
					try
					{
						Data101Full = XDocument.Parse(ws.Data101FullExV2XML(CredOrgNumbers,IndCode,DateTime.Parse("02.10.2017"),DateTime.Now).OuterXml);
					}
					catch
					{
						Console.WriteLine("Trying connection...");
					}
				}
				

				try
				{
					foreach (XElement d in Data101Full.Root.Elements("FD"))
					{
						Form101 f = new Form101();
						
						try { f.RegNum = d.Element("regn").Value; } catch { f.RegNum = ""; } //Регистрацонный номер КО
						
						try { f.Date = d.Element("DT").Value; } catch { try { f.Date = d.Element("dt").Value; } catch { f.Date = ""; } } //Дата (значения показателей привязаны к дате)
						
						try { f.value = d.Element("val").Value; } catch { f.value = ""; } //Значение символа - cумма
						
						try { f.IndID = e.Element("IndID").Value; } catch { f.IndID = ""; } //Внутренний код индикатора
						
						try { f.IndCode = e.Element("IndCode").Value; } catch { f.IndCode = ""; } //Номер счета индикатора
						
						try { f.name = e.Element("name").Value; } catch { f.name = ""; } //Название счетов баланса индикатора
						
						try { f.IndType = e.Element("IndType").Value; } catch { f.IndType = ""; } //Код счета индикатора
						
						try { f.IndChapter = e.Element("IndChapter").Value; } catch { f.IndChapter = ""; } //Код раздела индикатора
						//
						try { f.pln = d.Element("pln").Value; } catch { f.pln = ""; } //Балансовые счета
						
						try { f.ap = d.Element("ap").Value; } catch { f.ap = ""; } //Актив/Пассив
						
						try { f.vr = d.Element("vr").Value; } catch { f.vr = ""; } //Входящие остатки в рублях
						
						try { f.vv = d.Element("vv").Value; } catch { f.vv = ""; } //Входящие остатки в ин. вал., драг. металлы
						
						try { f.vitg = d.Element("vitg").Value; } catch { f.vitg = ""; } //Входящие остатки итого
						
						try { f.ora = d.Element("ora").Value; } catch { f.ora = ""; } //Обороты за отчетный период по дебету в рублях
						
						try { f.ova = d.Element("ova").Value; } catch { f.ova = ""; } //Обороты за отчетный период по дебету в ин. вал., драг. металлы
						
						try { f.oitga = d.Element("oitga").Value; } catch { f.oitga = ""; } //Обороты за отчетный период по дебету итого
						
						try { f.orp = d.Element("orp").Value; } catch { f.orp = ""; } //Обороты за отчетный период по кредиту в рублях
						
						try { f.ovp = d.Element("ovp").Value; } catch { f.ovp = ""; } //Обороты за отчетный период по кредиту в ин. вал., драг. металлы
						
						try { f.oitgp = d.Element("oitgp").Value; } catch { f.oitgp = ""; } //Обороты за отчетный период по кредиту итого
						
						try { f.ir = d.Element("ir").Value; } catch { f.ir = ""; } //Исходящие остатки в рублях
						
						try { f.iv = d.Element("iv").Value; } catch { f.iv = ""; } //Исходящие остатки в ин. вал., драг. металлы
						
						try { f.iitg = d.Element("iitg").Value; } catch { f.iitg = ""; } //Исходящие остатки итого
						
						
						Console.WriteLine("{0} {1} {2}", "Форма 101", f.RegNum, f.IndCode);
						
						writer.WriteLine(f.IndID+";"+f.IndCode+";"+f.name.Trim().Replace(";","").Replace("\r"," ").Replace("\n","")+";"+f.IndType+";"+f.IndChapter+";"+f.RegNum+";"+f.Date+";"+
						                 f.value+";"+f.pln+";"+f.ap+";"+f.vr+";"+f.vv+";"+f.vitg+";"+f.ora+";"+f.ova+";"+f.oitga+";"+f.orp+";"+f.ovp+";"+f.oitgp+";"+
						                 f.ir+";"+f.iv+";"+f.iitg
						                );
						writer.Flush();
						
						
					}
					
				} catch { Console.WriteLine("Нет такого элемента!"); }
				
				try
				{
					foreach (XElement d in Data101Full.Root.Elements("FDF"))
					{
						Form101 f = new Form101();
						
						try { f.RegNum = d.Element("regn").Value; } catch { f.RegNum = ""; } //Регистрацонный номер КО
						
						try { f.Date = d.Element("DT").Value; } catch { try { f.Date = d.Element("dt").Value; } catch { f.Date = ""; } } //Дата (значения показателей привязаны к дате)
						
						try { f.value = d.Element("val").Value; } catch { f.value = ""; } //Значение символа - cумма
						
						try { f.IndID = e.Element("IndID").Value; } catch { f.IndID = ""; } //Внутренний код индикатора
						
						try { f.IndCode = e.Element("IndCode").Value; } catch { f.IndCode = ""; } //Номер счета индикатора
						
						try { f.name = e.Element("name").Value; } catch { f.name = ""; } //Название счетов баланса индикатора
						
						try { f.IndType = e.Element("IndType").Value; } catch { f.IndType = ""; } //Код счета индикатора
						
						try { f.IndChapter = e.Element("IndChapter").Value; } catch { f.IndChapter = ""; } //Код раздела индикатора
						//
						try { f.pln = d.Element("pln").Value; } catch { f.pln = ""; } //Балансовые счета
						
						try { f.ap = d.Element("ap").Value; } catch { f.ap = ""; } //Актив/Пассив
						
						try { f.vr = d.Element("vr").Value; } catch { f.vr = ""; } //Входящие остатки в рублях
						
						try { f.vv = d.Element("vv").Value; } catch { f.vv = ""; } //Входящие остатки в ин. вал., драг. металлы
						
						try { f.vitg = d.Element("vitg").Value; } catch { f.vitg = ""; } //Входящие остатки итого
						
						try { f.ora = d.Element("ora").Value; } catch { f.ora = ""; } //Обороты за отчетный период по дебету в рублях
						
						try { f.ova = d.Element("ova").Value; } catch { f.ova = ""; } //Обороты за отчетный период по дебету в ин. вал., драг. металлы
						
						try { f.oitga = d.Element("oitga").Value; } catch { f.oitga = ""; } //Обороты за отчетный период по дебету итого
						
						try { f.orp = d.Element("orp").Value; } catch { f.orp = ""; } //Обороты за отчетный период по кредиту в рублях
						
						try { f.ovp = d.Element("ovp").Value; } catch { f.ovp = ""; } //Обороты за отчетный период по кредиту в ин. вал., драг. металлы
						
						try { f.oitgp = d.Element("oitgp").Value; } catch { f.oitgp = ""; } //Обороты за отчетный период по кредиту итого
						
						try { f.ir = d.Element("ir").Value; } catch { f.ir = ""; } //Исходящие остатки в рублях
						
						try { f.iv = d.Element("iv").Value; } catch { f.iv = ""; } //Исходящие остатки в ин. вал., драг. металлы
						
						try { f.iitg = d.Element("iitg").Value; } catch { f.iitg = ""; } //Исходящие остатки итого
						
						
						Console.WriteLine("{0} {1} {2}", "Форма 101", f.RegNum, f.IndCode);
						
						writer.WriteLine(f.IndID+";"+f.IndCode+";"+f.name.Trim().Replace(";","").Replace("\r"," ").Replace("\n","")+";"+f.IndType+";"+f.IndChapter+";"+f.RegNum+";"+f.Date+";"+
						                 f.value+";"+f.pln+";"+f.ap+";"+f.vr+";"+f.vv+";"+f.vitg+";"+f.ora+";"+f.ova+";"+f.oitga+";"+f.orp+";"+f.ovp+";"+f.oitgp+";"+
						                 f.ir+";"+f.iv+";"+f.iitg
						                );
						writer.Flush();
						
					}
				} catch { Console.WriteLine("Нет такого элемента!"); }
				
				
			}
			writer.Close();
			writer.Dispose();
			
			
			
			//------------------------------------------------------------------------------
			

			List<Form102> clsForm102 = new List<Form102>();
			foreach (XElement e in cls102.Root.Elements("SIND").Where(a=>a.Element("symbol") != null))
			{
				try {
				

					
						Form102 f = new Form102();
						
						try { f.symid = e.Element("symid").Value; } catch { f.symid = ""; } //внутренний код
						
						try { f.symsort = e.Element("symsort").Value; } catch { f.symsort = ""; } //порядок сортировки
						
						try { f.symbol = e.Element("symbol").Value; } catch { f.symbol = ""; } //Код показателя
						
						try { f.symtype = e.Element("symtype").Value; } catch { f.symtype = ""; }
						
						try { f.name = e.Element("name").Value.Trim(); } catch { f.name = ""; } //Название показателя
						
						try { f.symset = e.Element("symset").Value; } catch { f.symset = ""; }
						
						clsForm102.Add(f);
					
					
				} catch {}
				
			}
			
			
			//------------------------------------------------------------------------
			
			
			//Собираем отчетность по 102 форме для всех КО
			//Сбрасываем в файл отчетность КО 102 формы
			writer = new StreamWriter("cls_CreditOrgsInfo_form102.csv",false,Encoding.Default);
			
			foreach (Form102 e in clsForm102)
			{
				try 
				{
					Console.WriteLine("Грузим данные по 102 форме...");
					XDocument Data102Full = null;
					while(Data102Full == null)
					{
						try
						{
							Data102Full = XDocument.Parse(ws.Data102FormExXML(CredOrgNumbers,Convert.ToInt32(e.symbol),DateTime.Parse("02.10.2017"),DateTime.Now).OuterXml);
						}
						catch
						{
							Console.WriteLine("Trying connection...");
						}
					}
					
					foreach (XElement d in Data102Full.Root.Elements("FD"))
					{
						Form102 f = new Form102();
						
						try { f.RegNum = d.Element("regnum").Value; } catch { f.RegNum = ""; } //Рег.номер КО
						
						try { f.Date = d.Element("DT").Value; } catch { f.Date = ""; } //Дата
						
						try { f.value = d.Element("val").Value; } catch { f.value = ""; } //Значение символа - cумма
						
						try { f.symid = e.symid; } catch { f.symid = ""; } //внутренний код
						
						try { f.symsort = e.symsort; } catch { f.symsort = ""; } //порядок сортировки
						
						try { f.symbol = e.symbol; } catch { f.symbol = ""; } //Код показателя
						
						try { f.symtype = e.symtype; } catch { f.symtype = ""; }
						
						try { f.name = e.name.Trim(); } catch { f.name = ""; } //Название показателя
						
						try { f.symset = e.symset; } catch { f.symset = ""; }
						
						Console.WriteLine("{0} {1} {2}", "102 форма", f.RegNum, f.symbol);
							
						writer.WriteLine(f.symid+";"+f.symsort+";"+f.symbol+";"+f.symtype+";"+f.name.Trim().Replace(";",",").Replace("\r"," ").Replace("\n","")+";"+f.symset+";"+f.RegNum+";"+f.Date+";"+f.value);
						writer.Flush();
											
					}
						
				} catch {}
				
			}
			
			writer.Close();
			writer.Dispose();
			
			Console.Write("Press any key to continue . . . ");
			Console.ReadKey(true);
		}
	}
}