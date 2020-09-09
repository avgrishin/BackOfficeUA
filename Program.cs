using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.IO.Compression;

namespace BackOfficeUA
{
  class Program
  {
    private static string LastWorkDate { get; set; }
    private static StreamWriter log;
    static async System.Threading.Tasks.Task Main(string[] args)
    {
      using (log = new StreamWriter("BackOfficeUA.log", true, Encoding.GetEncoding(1251)))
      {
        try
        {
          WebClient wc = new WebClient
          {
            Credentials = new NetworkCredential("techadmin", "Bpgjldsgjldthnf77!", "AM"),
          };
          LastWorkDate = wc.DownloadString("http://assetsmgr/home/GetLastWorkDate");
          ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
          service.UseDefaultCredentials = false;
          service.Credentials = new NetworkCredential("BackOfficeUA", "QWESZ123@");
          service.AutodiscoverUrl("BackOfficeUA@am-uralsib.ru", RedirectionUrlValidationCallback);
          JavaScriptSerializer jss = new JavaScriptSerializer();
          var cfg = new EFolder[] { };

          var FileName = "BackOfficeUA.dat";

          if (File.Exists(FileName))
            using (StreamReader sr = new StreamReader(FileName, Encoding.GetEncoding(1251)))
            {
              cfg = jss.Deserialize<EFolder[]>(sr.ReadToEnd());
            }
          foreach (var ef in cfg)
          //for (int i= cfg.Length-1; i>=0;i--)
          {
            Console.WriteLine(ef.DisplayName);
            FolderView view = new FolderView(1);
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName);
            SearchFilter searchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, ef.DisplayName);
            FindFoldersResults findResults = service.FindFolders(new FolderId(WellKnownFolderName.Inbox), searchFilter, view);
            if (findResults.TotalCount == 1)
            {
              FolderId id1 = findResults.Folders[0].Id;
              bool moreChangesAvailable = false;
              do
              {
                ChangeCollection<ItemChange> icc = service.SyncFolderItems(id1, PropertySet.IdOnly, null, 512, SyncFolderItemsScope.NormalItems, ef.SyncState);
                if (icc.Count > 0)
                {
                  foreach (ItemChange ic in icc)
                  {
                    if (ic.ChangeType == ChangeType.Create)
                    {
                      EmailMessage e = ic.Item as EmailMessage;
                      e.Load(new PropertySet(EmailMessageSchema.Attachments, EmailMessageSchema.From, EmailMessageSchema.Subject, EmailMessageSchema.IsRead));
                      foreach (Attachment attachment in e.Attachments)
                      {
                        if (attachment is FileAttachment)
                        {
                          FileAttachment fileAttachment = attachment as FileAttachment;
                          Console.WriteLine("File attachment name: " + fileAttachment.Name);
                          if (e.From.Address == "broker3@open.ru")
                          {
                            OpenFc(fileAttachment);
                          }
                          else if (new string[] { "webreport@uralsibweb.ru", "romanovskijam@uralsib.ru", "GordinaTS@uralsib.ru", "novikovanv@uralsib.ru" , "MaltsevMY@uralsib.ru" }.Contains(e.From.Address))
                          {
                            UralsibWeb(fileAttachment);
                          }
                          else if (e.From.Address == "back-office@open.ru")
                          {
                            OpenBr(fileAttachment);
                          }
                          else if (e.From.Address.EndsWith("@bloomberg.net"))
                          {
                            Bloomberg(fileAttachment);
                          }
                          else if (e.From.Address == "noreply@bcs.ru" && e.Subject.StartsWith("Broker report"))
                          {
                            await Bcs(fileAttachment);
                          }
                          else if (e.From.Address == "mivanova@msk.bcs.ru" || e.From.Address == "ik@bcs.ru")
                          {
                            await BcsM(fileAttachment);
                          }
                        }
                      }
                    }
                  }
                }
                ef.SyncState = icc.SyncState;
                moreChangesAvailable = icc.MoreChangesAvailable;
              }
              while (moreChangesAvailable);
            }
          }
          using (var sw = new StreamWriter(FileName, false, Encoding.GetEncoding(1251)))
          {
            sw.WriteLine(jss.Serialize(cfg));
          }
          //log.Write(DateTime.Now);
          //log.Write(" ");
          //log.WriteLine(jss.Serialize(cfg));
        }
        catch (Exception ex)
        {
          if (!string.IsNullOrEmpty(ex.Message))
          {
            log.Write(DateTime.Now);
            log.Write(" ");
            log.WriteLine(ex.Message);
          }
        }
      }
    }

    private static void UralsibWeb(FileAttachment fileAttachment)
    {
      if (LastWorkDate.Length == 6)
      {
        if (Regex.IsMatch(fileAttachment.Name, "^(?:(?:brok_rpt_.+_(?:12351|13085|16282|16283|31446|31447|31443|90203|90210|90212)[EBXN]{0,1}_\\d.+_final\\.xls\\.zip)|(?:(?:12351|13085|16282|16283|31446|31447|31443|37878|38480|38484)-.+?-(\\d{2}\\.\\d{2}\\.\\d{4})-\\1\\.xls))$"))
        {
          var dir = @"V:\VOL1\ASSETS\OUK\Реестры биржевых сделок_УК УралСиб\";
          dir = Path.Combine(dir, "20" + LastWorkDate.Substring(4, 2));
          if (!Directory.Exists(dir))
            Directory.CreateDirectory(dir);
          dir = Path.Combine(dir, LastWorkDate.Substring(2, 2));
          if (!Directory.Exists(dir))
            Directory.CreateDirectory(dir);
          dir = Path.Combine(dir, LastWorkDate.Substring(0, 2));
          if (!Directory.Exists(dir))
            Directory.CreateDirectory(dir);
          fileAttachment.Load(Path.Combine(dir, fileAttachment.Name));

          var fn = Path.Combine(@"V:\VOL1\ASSETS\EDO\NSD\Cheremisina\OUT\OUT\", fileAttachment.Name);
          if (!File.Exists(fn))
            fileAttachment.Load(fn);
        }
      }
    }

    private static void OpenFc(FileAttachment fileAttachment)
    {
      if (LastWorkDate.Length == 6)
      {
        var d1 = new Regex("((\\d{2})\\.(\\d{2})\\.(\\d{4})).*?\\1|(\\d{4})(\\d{2})(\\d{2})");
        if (d1.IsMatch(fileAttachment.Name))
        {
          if (new string[] { ".html", ".xml" }.Contains(Path.GetExtension(fileAttachment.Name)))
          {
            var fn = Path.Combine(@"V:\VOL1\ASSETS\EDO\NSD\Cheremisina\OUT\OUT\", fileAttachment.Name);
            if (!File.Exists(fn))
              fileAttachment.Load(fn);
            var dir = @"V:\VOL1\ASSETS\OUK\Реестры биржевых сделок_УК УралСиб\";
            dir = Path.Combine(dir, "20" + LastWorkDate.Substring(4, 2));
            if (!Directory.Exists(dir))
              Directory.CreateDirectory(dir);
            dir = Path.Combine(dir, LastWorkDate.Substring(2, 2));
            if (!Directory.Exists(dir))
              Directory.CreateDirectory(dir);
            dir = Path.Combine(dir, LastWorkDate.Substring(0, 2));
            if (!Directory.Exists(dir))
              Directory.CreateDirectory(dir);
            fileAttachment.Load(Path.Combine(dir, fileAttachment.Name));

            if (Path.GetExtension(fileAttachment.Name) == ".xml")
            {
              fileAttachment.Load(Path.Combine(@"V:\VOL1\ASSETS\Reports_BROKER\BANK FK OTKRYTIE\", fileAttachment.Name));
            }
          }
        }
      }
    }

    private static void OpenBr(FileAttachment fileAttachment)
    {
      var d1 = new Regex("^2005_20\\d{2}-\\d{2}-\\d{2}_day_rus_MMVB_TP.zip$");
      if (d1.IsMatch(fileAttachment.Name))
      {
        fileAttachment.Load(Path.Combine(@"V:\VOL1\ASSETS\Reports_BROKER\OTKRYTIE\Отчет_брокера\", fileAttachment.Name));
      }
    }
    private static async System.Threading.Tasks.Task Bcs(FileAttachment fileAttachment)
    {
      var d1 = new Regex("^BackOfficeUA@am-uralsib.ru.B.k.+?\\.(zip|xls)$");
      if (d1.IsMatch(fileAttachment.Name))
      {
        var fn = Path.Combine(@"V:\VOL1\ASSETS\Reports_BROKER\BROKER CREDIT\", fileAttachment.Name);
        fileAttachment.Load(fn);

        var d2 = new Regex("^BackOfficeUA@am-uralsib.ru.B.k-(1154407|1154399|1154387|1145284|1175291|1175007|1175175|1174873|1175067|1175340).+?\\.zip$");
        if (d2.IsMatch(fileAttachment.Name))
        {
          using (ZipArchive archive = new ZipArchive(File.OpenRead(fn), ZipArchiveMode.Read))
          {
            foreach (var e in archive.Entries)
            {
              var fnb = Path.Combine(@"c:\tmp", $"{Path.GetFileName(e.Name)}");
              //            var fnd = Path.Combine(@"\\am-uralsib.ru\uralsib\MSK\COMMON\VOL1\ASSETS\EDO\NSD\Cheremisina\OUT\OUT\", Path.GetFileName(e.Name));
              using (Stream zipStream = e.Open())
              {
                using (FileStream fileStream = new FileStream(fnb, FileMode.Create))
                {
                  await zipStream.CopyToAsync(fileStream);
                }
              }
              //              File.Copy(fnb, fnd, true);
              //await log.WriteLineAsync($"{DateTime.Now} {fnd}");

              var dir = @"V:\VOL1\ASSETS\OUK\Реестры биржевых сделок_УК УралСиб\";
              dir = Path.Combine(dir, "20" + LastWorkDate.Substring(4, 2));
              if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
              dir = Path.Combine(dir, LastWorkDate.Substring(2, 2));
              if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
              dir = Path.Combine(dir, LastWorkDate.Substring(0, 2));
              if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
              var fnd = Path.Combine(dir, Path.GetFileName(e.Name));
              File.Copy(fnb, fnd, true);
              await log.WriteLineAsync($"{DateTime.Now} {fnd}");
              File.Delete(fnb);
            }
          }
        }
      }
    }
    private static async System.Threading.Tasks.Task BcsM(FileAttachment fileAttachment)
    {
      if (new string[] { ".jpg", ".png", ".pdf" }.Contains(Path.GetExtension(fileAttachment.Name)))
        return;
      var fn = Path.Combine(@"\\am-uralsib.ru\uralsib\MSK\COMMON\VOL1\ASSETS\EDO\NSD\Cheremisina\OUT\OUT\", fileAttachment.Name.Replace("?", ""));
      fileAttachment.Load(fn);
      await log.WriteLineAsync($"{DateTime.Now} {fn}");

      var dir = @"V:\VOL1\ASSETS\OUK\Реестры биржевых сделок_УК УралСиб\";
      dir = Path.Combine(dir, "20" + LastWorkDate.Substring(4, 2));
      if (!Directory.Exists(dir))
        Directory.CreateDirectory(dir);
      dir = Path.Combine(dir, LastWorkDate.Substring(2, 2));
      if (!Directory.Exists(dir))
        Directory.CreateDirectory(dir);
      dir = Path.Combine(dir, LastWorkDate.Substring(0, 2));
      if (!Directory.Exists(dir))
        Directory.CreateDirectory(dir);
      fn = Path.Combine(dir, fileAttachment.Name.Replace("?", ""));
      fileAttachment.Load(fn);
      await log.WriteLineAsync($"{DateTime.Now} {fn}");

      fn = Path.Combine(@"V:\VOL1\ASSETS\4All\BCS_BROKER\", fileAttachment.Name.Replace("?", ""));
      fileAttachment.Load(fn);

    }
    private static void Bloomberg(FileAttachment fileAttachment)
    {
      var d1 = new Regex("^.+?\\.gif$");
      if (d1.IsMatch(fileAttachment.Name))
      {
        var fn = Path.Combine(@"V:\VOL1\ASSETS\EDO\NSD\Cheremisina\OUT\OUT\", fileAttachment.Name);
        if (!File.Exists(fn))
          fileAttachment.Load(fn);
      }
    }
    private static bool RedirectionUrlValidationCallback(string redirectionUrl)
    {
      // The default for the validation callback is to reject the URL.
      bool result = false;
      Uri redirectionUri = new Uri(redirectionUrl);
      // Validate the contents of the redirection URL. In this simple validation
      // callback, the redirection URL is considered valid if it is using HTTPS
      // to encrypt the authentication credentials. 
      if (redirectionUri.Scheme == "https")
      {
        result = true;
      }
      return result;
    }
  }

  public class EFolder
  {
    public string DisplayName { get; set; }
    public string SyncState { get; set; }
  }
}
