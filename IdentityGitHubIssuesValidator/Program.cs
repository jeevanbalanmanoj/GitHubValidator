using IdentityGitHubIssuesValidator.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
//using Octokit;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
//using ProductHeaderValue = Octokit.ProductHeaderValue;
using Issue = IdentityGitHubIssuesValidator.Models.Issue;

namespace IdentityGitHubIssuesValidator
{
    class Program
    {
        static HttpClient client = new HttpClient();
        static int stateCol = 25;
        static int repoUrlCol = 2;
        static int apiUrl = 26;


        static string GetGitHubIssueApiUrl(string owner, string repoName, int issueId)
        {
            return "https://api.github.com/repos/" + owner + "/" + repoName + "/issues/" + issueId;
        }
        static IssueRequest GetIssueRequest(string repoUrl)
        {
            string[] words = repoUrl.Split('/');
            return new IssueRequest
            {
                Owner = words[3],
                RepoName = words[4],
                IssueId = int.Parse(words[6])

            };
        }

        static async Task<Issue> GetIssueAsync(IssueRequest request)
        {
            Issue issue = null;
            //var byteArray = Encoding.ASCII.GetBytes("username:password");
            //client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));           
            string path = GetGitHubIssueApiUrl(request.Owner, request.RepoName, request.IssueId);
            HttpResponseMessage response = await client.GetAsync(path);
            if (response.IsSuccessStatusCode)
            {
                issue = await response.Content.ReadAsAsync<Issue>();
            }
            else
                Console.WriteLine(response);

              return issue;
        }
        static async Task Main(string[] args)
        {
            try
            {
                string errors = "";
                string dataAnomalies = "";
                client.DefaultRequestHeaders.Add("User-Agent", "request");
                client.DefaultRequestHeaders.Add("Authorization", "<<from postman>>");
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(@"C:\data\AllData1110.xlsx")))
                {
                    var firstSheet = package.Workbook.Worksheets["Sheet1"];
                    int colCount = firstSheet.Dimension.End.Column;  //get Column Count
                    int rowCount = firstSheet.Dimension.End.Row;     //get row count
                    int totalIssues = rowCount - 4;
                    int mistakeCount = 0;
                    if (colCount != stateCol)
                    {
                        firstSheet.InsertColumn(stateCol, 1);
                        firstSheet.Cells[3, stateCol].Value = "Issue state";
                    }
                        // firstSheet.Cells["A1:B1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    // firstSheet.Cells[3, colCount + 1].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    Issue issue = null;
                    IssueRequest request = null;
                    HashSet<string> myhash1 = new HashSet<string>();

                    for (int row = 4; row <= rowCount; row++)
                    {
                        if (row % 100 == 0)
                            Console.WriteLine("At Row " +row);
                        request  = GetIssueRequest(firstSheet.Cells[row, repoUrlCol].Value?.ToString().Trim());
                        issue = await GetIssueAsync(request);
                        if (issue == null)
                        {
                            firstSheet.Cells[row, stateCol].Value ="ERROR";
                            Console.WriteLine("Error  at row  " + row);
                            errors += "\n" + request.Owner + "    " + request.RepoName + "  NA  " + firstSheet.Cells[row, repoUrlCol].Value?.ToString().Trim();
                            Console.WriteLine(firstSheet.Cells[row, repoUrlCol].Value?.ToString().Trim());
                            continue;
                        }
                        else
                            firstSheet.Cells[row, stateCol].Value = issue.State;
                        if (!issue.State.Equals("open")) 
                        {
                            dataAnomalies+="\n"+ request.Owner+"    "+request.RepoName+"    "+issue.Id+"     " + firstSheet.Cells[row, repoUrlCol].Value?.ToString().Trim();
                            Console.WriteLine(firstSheet.Cells[row, repoUrlCol].Value?.ToString().Trim());
                            mistakeCount++;
                        }
                
                    }
                    package.Save();
                    using (StreamWriter writer = new StreamWriter(@"C:\data\errors.txt", false))
                    {
                        writer.Write(errors);
                    }
                    using (StreamWriter writer = new StreamWriter(@"C:\data\Data_Anomalies.txt", false))
                    {
                        writer.Write(dataAnomalies);
                    }
                    Console.WriteLine(" Total " + totalIssues + " Mistakes " + mistakeCount + "Percentage" + (mistakeCount / totalIssues) * 100 + "%");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }
    }
}
