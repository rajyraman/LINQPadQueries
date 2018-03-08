<Query Kind="Program">
  <Reference>&lt;RuntimeDirectory&gt;\Microsoft.Build.Framework.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\Microsoft.Build.Tasks.v4.0.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\Microsoft.Build.Utilities.v4.0.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.ComponentModel.DataAnnotations.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.Configuration.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.Design.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.DirectoryServices.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.DirectoryServices.Protocols.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.EnterpriseServices.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.Runtime.Caching.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.Security.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.ServiceProcess.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.Web.ApplicationServices.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.Web.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.Web.RegularExpressions.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.Web.Services.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.Windows.Forms.dll</Reference>
  <NuGetReference>AngleSharp</NuGetReference>
  <Namespace>AngleSharp</Namespace>
  <Namespace>AngleSharp.Dom</Namespace>
  <Namespace>AngleSharp.Dom.Html</Namespace>
  <Namespace>System</Namespace>
  <Namespace>System.Net</Namespace>
  <Namespace>System.Runtime.InteropServices</Namespace>
  <Namespace>System.Web</Namespace>
</Query>

void Main()
{
	var awardCategories = new string[]{
	  "Access",
	  "AI",
	  "Business Solutions",
	  "Cloud and Datacenter Management",
	  "Data Platform",
	  "Enterprise Mobility",
	  "Excel",
	  "Microsoft Azure",
	  "Office Development",
	  "Office Servers and Services",
	  "OneNote",
	  "Outlook",
	  "PowerPoint",
	  "Visio",
	  "Visual Studio and Development Technologies",
	  "Windows and Devices for IT",
	  "Windows Development",
	  "Word"
	};
	var concatCategories = string.Join("\n", awardCategories.Select((c, i) => $"{i + 1}: {c}"));
	int.TryParse(Util.ReadLine($"Please choose an award category:\n{concatCategories}"), out int choice);
	if (choice == 0)
	{
		"Invalid choice".Dump();
		return;
	};
	var selectedCategory = awardCategories[choice - 1];
	var pageNumber = 1;
	var webClient = new WebClient();
	IHtmlDocument doc;
	IHtmlCollection<IElement> mvpProfiles = null;
	List<MVPInfo> mvps = new List<MVPInfo>();
	var dc = new DumpContainer().Dump("MVPs");
	var htmlParser = new AngleSharp.Parser.Html.HtmlParser(new Configuration().WithDefaultLoader());
	do
	{
		var data = webClient.DownloadData(new Uri($"https://mvp.microsoft.com/en-us/MvpSearch?ex={HttpUtility.UrlEncode(selectedCategory)}&sc=s&pn={pageNumber++}&ps=48"));
		doc = htmlParser.Parse(Encoding.UTF8.GetString(data));

		mvpProfiles = doc.QuerySelectorAll(".profileListItem");
		var currentPage = mvpProfiles
		.Select(d =>
		{
			var nameAndLocation = d.QuerySelector(".profileListItemFullName a");
			return new MVPInfo
			{
				Name = nameAndLocation.TextContent,
				Url = $"https://mvp.microsoft.com{nameAndLocation.GetAttribute("href")}",
				Country = d.QuerySelector(".profileListItemLocation .subItemContent").TextContent
			};
		})
		.ToList();
		mvps.AddRange(currentPage);
		dc.Content = $"Page {pageNumber} -> {currentPage.Count} MVPs";
	} while (mvpProfiles.Any());
	dc.Content = "Finished retrieving all MVPs";
	Util.WriteCsv(mvps, $@"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\{selectedCategory} MVPs.csv");
	new Hyperlinq($@"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\{selectedCategory} MVPs.csv", $"Open {selectedCategory} MVPs.csv").Dump();
	mvps.Dump();
	webClient.Dispose();
}

public class MVPInfo
{
	public string Name { get; set; }
	public string Url { get; set; }
	public string Country { get; set; }
}
