using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO.Compression;
using Ionic.Zip;
using System.Threading;

namespace TMC
{
    class Program
    {
        static HistoryWeight _hw = new HistoryWeight();
        static LastActivityWeight _lw = new LastActivityWeight();

        static void Main(string[] args)
        {
            #region Open file
            IWorkbook book = null;
            try {
                using (FileStream file = new FileStream(@"Record.xlsx", FileMode.Open, FileAccess.Read)) {
                    book = new XSSFWorkbook(file);
                }
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Press enter to exit.");
                Console.ReadLine();
                return;
            }
            #endregion

            LoadWeight();

            #region Backup
            Console.Write("Start to backup.");
            using (ZipFile zip = new ZipFile()) {
                zip.AddFile(@".\Record.xlsx");
                var output = @".\Record.zip";
                zip.Save(output);
            }
            Console.WriteLine("\rBackup finished.");
            #endregion

            #region Excel style setting
            XSSFCellStyle style = (XSSFCellStyle)book.CreateCellStyle();
            XSSFFont font = (XSSFFont)book.CreateFont();
            font.FontHeightInPoints = 10;
            font.FontName = "Times New Roman";
            style.SetFont(font);
            #endregion

            #region History analysis
            Console.WriteLine("History record analysis start.");
            ISheet activitySheet = book.GetSheet("Activity");
            ISheet journeySheet = book.GetSheet("Journey");
            ISheet myIeRecordSheet = book.GetSheet("MyIERecord");
            List<RecentActivity> raColl = new List<RecentActivity>();
            List<Journey> journeyColl = new List<Journey>();
            List<MyIERecord> myIeRecordColl = new List<MyIERecord>();
            for (int i = 1; i < activitySheet.LastRowNum + 1; i++) {
                if (activitySheet.GetRow(i) != null && activitySheet.GetRow(i).GetCell(0) != null) {
                    RecentActivity ra = new RecentActivity
                    {
                        #region prop setting
                        Name = activitySheet.GetRow(i).GetCell(0).StringCellValue.Trim(),
                        TME = (int)activitySheet.GetRow(i).GetCell(1).NumericCellValue,
                        Timer = (int)activitySheet.GetRow(i).GetCell(2).NumericCellValue,
                        AhCounter = (int)activitySheet.GetRow(i).GetCell(3).NumericCellValue,
                        Variety = (int)activitySheet.GetRow(i).GetCell(4).NumericCellValue,
                        TableTopics = (int)activitySheet.GetRow(i).GetCell(5).NumericCellValue,
                        GE = (int)activitySheet.GetRow(i).GetCell(6).NumericCellValue,
                        IE = (int)activitySheet.GetRow(i).GetCell(7).NumericCellValue,
                        LE = (int)activitySheet.GetRow(i).GetCell(8).NumericCellValue,
                        #endregion
                    };
                    raColl.Add(ra);

                    Journey sh = new Journey
                    {
                        Name = activitySheet.GetRow(i).GetCell(0).StringCellValue.Trim(),
                        Achievements = new List<Role>()
                    };
                    journeyColl.Add(sh);

                    MyIERecord record = new MyIERecord
                    {
                        Name = activitySheet.GetRow(i).GetCell(0).StringCellValue.Trim(),
                        Records = new List<IERecord>()
                    };
                    myIeRecordColl.Add(record);
                }
                Console.Write($"\rProgress {i - 1}/{activitySheet.LastRowNum - 1}");
            }
            Console.WriteLine();
            Console.WriteLine("History analysis finished.");
            #endregion

            #region Meeting analysis
            Console.WriteLine("Meeting record analysis start.");
            ISheet meetingSheet = book.GetSheet("Meeting");
            List<Meeting> meetings = new List<Meeting>();
            for (int i = 1; i <= meetingSheet.LastRowNum; i++) {
                IRow row = meetingSheet.GetRow(i);
                if (row.Cells.All(d => d.CellType == CellType.Blank))
                    break;
                if (DateTime.Now < DateTime.FromOADate(row.GetCell(1).NumericCellValue))
                    continue;
                if (row != null) {
                    Meeting m = new Meeting
                    {
                        #region prop setting
                        Id = (int)row.GetCell(0)?.NumericCellValue,
                        DATE = DateTime.FromOADate(row.GetCell(1).NumericCellValue),
                        TME = row.GetCell(2)?.StringCellValue.Trim(),
                        Timer = row.GetCell(3)?.StringCellValue.Trim(),
                        AhCounter = row.GetCell(4)?.StringCellValue.Trim(),
                        Variety = row.GetCell(5)?.StringCellValue.Trim(),
                        EventHost = row.GetCell(6)?.StringCellValue.Trim(),
                        EventTitle = row.GetCell(7)?.StringCellValue.Trim(),
                        Speaker1 = row.GetCell(8)?.StringCellValue.Trim(),
                        Prj1 = row.GetCell(9)?.StringCellValue.Trim(),
                        Title1 = row.GetCell(10)?.StringCellValue.Trim(),
                        Speaker2 = row.GetCell(11)?.StringCellValue.Trim(),
                        Prj2 = row.GetCell(12)?.StringCellValue.Trim(),
                        Title2 = row.GetCell(13)?.StringCellValue.Trim(),
                        Speaker3 = row.GetCell(14)?.StringCellValue.Trim(),
                        Prj3 = row.GetCell(15)?.StringCellValue.Trim(),
                        Title3 = row.GetCell(16)?.StringCellValue.Trim(),
                        Speaker4 = row.GetCell(17)?.StringCellValue.Trim(),
                        Prj4 = row.GetCell(18)?.StringCellValue.Trim(),
                        Title4 = row.GetCell(19)?.StringCellValue.Trim(),
                        TableTopics = row.GetCell(20)?.StringCellValue.Trim(),
                        GE = row.GetCell(21)?.StringCellValue.Trim(),
                        IE1 = row.GetCell(22)?.StringCellValue.Trim(),
                        IE2 = row.GetCell(23)?.StringCellValue.Trim(),
                        IE3 = row.GetCell(24)?.StringCellValue.Trim(),
                        IE4 = row.GetCell(25)?.StringCellValue.Trim(),
                        LE = row.GetCell(26)?.StringCellValue.Trim(),
                        RowIndex = row.RowNum
                        #endregion
                    };
                    meetings.Add(m);
                }
                Console.Write($"\rProgress {i}");
            }
            Console.WriteLine();
            Console.WriteLine("Meeting analysis finished.");
            #endregion

            #region Journey analysis
            Console.WriteLine("Journey analysis start.");
            foreach (var meeting in meetings) {
                #region Add role to journey
                Role done = new Role
                {
                    Name = "AhCounter",
                    Date = meeting.DATE
                };
                journeyColl.FirstOrDefault(sh => sh.Name == meeting.AhCounter)?.Achievements?.Add(done);
                done = new Role
                {
                    Name = "GE",
                    Date = meeting.DATE
                };
                journeyColl.FirstOrDefault(sh => sh.Name == meeting.GE)?.Achievements?.Add(done);
                done = new Role
                {
                    Name = "IE",
                    Date = meeting.DATE
                };
                var IEs = journeyColl.Where(sh => sh.Name == meeting.IE1 || sh.Name == meeting.IE2 || sh.Name == meeting.IE3 || sh.Name == meeting.IE4);
                foreach (var ie in IEs) {
                    ie?.Achievements?.Add(done);
                }
                done = new Role
                {
                    Name = "LE",
                    Date = meeting.DATE
                };
                journeyColl.FirstOrDefault(sh => sh.Name == meeting.LE)?.Achievements?.Add(done);
                done = new Role
                {
                    Name = "Speaker",
                    Date = meeting.DATE
                };
                var speakers = journeyColl.Where(sh => sh.Name == meeting.Speaker1 || sh.Name == meeting.Speaker2 || sh.Name == meeting.Speaker3 || sh.Name == meeting.Speaker4);
                foreach (var speaker in speakers) {
                    speaker?.Achievements?.Add(done);
                }
                done = new Role
                {
                    Name = "TopicsMaster",
                    Date = meeting.DATE
                };
                journeyColl.FirstOrDefault(sh => sh.Name == meeting.TableTopics)?.Achievements?.Add(done);
                done = new Role
                {
                    Name = "Timer",
                    Date = meeting.DATE
                };
                journeyColl.FirstOrDefault(sh => sh.Name == meeting.Timer)?.Achievements?.Add(done);
                done = new Role
                {
                    Name = "Variety",
                    Date = meeting.DATE
                };
                journeyColl.FirstOrDefault(sh => sh.Name == meeting.Variety)?.Achievements?.Add(done);
                done = new Role
                {
                    Name = "TME",
                    Date = meeting.DATE
                };
                journeyColl.FirstOrDefault(sh => sh.Name == meeting.TME)?.Achievements?.Add(done);
                Console.Write($"\rProgress {meeting.DATE}");
                #endregion
            }
            Console.WriteLine();
            Console.WriteLine("Finished Journey analysis.");

            int bound = 1;
            Dictionary<string, int> indexOfName = new Dictionary<string, int>();
            raColl.Select(ra => ra.Name).ToList().ForEach(name => { indexOfName[name] = bound++; });

            foreach (var key in indexOfName.Keys) {
                journeySheet.GetRow(0).CreateCell(indexOfName[key]).CellStyle = style;
                journeySheet.GetRow(0).GetCell(indexOfName[key]).SetCellValue(key);
            }

            #region Journey writing
            Console.WriteLine("Start Journey sheet writing.");
            foreach (var journey in journeyColl) {
                int index = indexOfName[journey.Name];
                for (int i = 1; i < journeySheet.LastRowNum + 1; i++) {
                    DateTime date = DateTime.FromOADate(journeySheet.GetRow(i).GetCell(0).NumericCellValue);
                    Role done = journey.Achievements.FirstOrDefault(d => d.Date == date);
                    if (done != null) {
                        journeySheet.GetRow(i).CreateCell(index).CellStyle = style;
                        journeySheet.GetRow(i).GetCell(index).SetCellValue(done.Name);
                    }
                    Console.Write($"\rProgress {journey.Name}: {i}/{journeySheet.LastRowNum}");
                }
                Console.WriteLine();
            }
            Console.WriteLine("Finished Journey sheet writing.");
            #endregion
            #endregion

            Console.WriteLine("IE Record analysis start.");
            foreach (var meeting in meetings) {
                #region Add IE record
                var speaker1 = myIeRecordColl.FirstOrDefault(me => me.Name == meeting.Speaker1);
                var speaker2 = myIeRecordColl.FirstOrDefault(me => me.Name == meeting.Speaker2);
                var speaker3 = myIeRecordColl.FirstOrDefault(me => me.Name == meeting.Speaker3);
                var speaker4 = myIeRecordColl.FirstOrDefault(me => me.Name == meeting.Speaker4);
                if (speaker1 != null) {
                    IERecord record = new IERecord
                    {
                        Name = meeting.IE1,
                        ProjLevel = meeting.Prj1,
                        MeetingDate = meeting.DATE
                    };
                    speaker1.Records.Add(record);
                }
                if (speaker2 != null) {
                    IERecord record = new IERecord
                    {
                        Name = meeting.IE2,
                        ProjLevel = meeting.Prj2,
                        MeetingDate = meeting.DATE
                    };
                    speaker2.Records.Add(record);
                }
                if (speaker3 != null) {
                    IERecord record = new IERecord
                    {
                        Name = meeting.IE3,
                        ProjLevel = meeting.Prj3,
                        MeetingDate = meeting.DATE
                    };
                    speaker3.Records.Add(record);
                }
                if (speaker4 != null) {
                    IERecord record = new IERecord
                    {
                        Name = meeting.IE4,
                        ProjLevel = meeting.Prj4,
                        MeetingDate = meeting.DATE
                    };
                    speaker4.Records.Add(record);
                }
                Console.Write($"\rProgress {meeting.DATE}");
                #endregion
            }
            Console.WriteLine();
            Console.WriteLine("Finished IE Record analysis.");

            #region IE record writing
            myIeRecordSheet.CreateRow(0);
            foreach (var key in indexOfName.Keys) {
                myIeRecordSheet.GetRow(0).CreateCell(indexOfName[key]).CellStyle = style;
                myIeRecordSheet.GetRow(0).GetCell(indexOfName[key]).SetCellValue(key);
            }

            Console.WriteLine("Start MyIERecord sheet writing.");
            foreach (var person in myIeRecordColl) {
                int index = indexOfName[person.Name];
                for (int i = 0; i < person.Records.Count(); i++) {
                    IERecord record = person.Records[i];
                    if (myIeRecordSheet.GetRow(i + 1) == null)
                        myIeRecordSheet.CreateRow(i + 1);
                    myIeRecordSheet.GetRow(i + 1).CreateCell(index).CellStyle = style;
                    myIeRecordSheet.GetRow(i + 1).GetCell(index).SetCellValue($"{record.ProjLevel} - {record.Name}");
                    Console.Write($"\rProgress {person.Name}: {i + 1}/{person.Records.Count()}");
                }
                Console.WriteLine();
            }
            Console.WriteLine("Finished Journey sheet writing.");
            #endregion

            #region Calculating score
            meetings.Reverse();
            Console.WriteLine("Start weight calculating.");
            #region Extract avg
            double avgTME = raColl.Average(ra => ra.TME);
            double avgTimer = raColl.Average(ra => ra.Timer);
            double avgAhCounter = raColl.Average(ra => ra.AhCounter);
            double avgVariety = raColl.Average(ra => ra.Variety);
            double avgTableTopics = raColl.Average(ra => ra.TableTopics);
            double avgGE = raColl.Average(ra => ra.GE);
            double avgIE = raColl.Average(ra => ra.IE);
            double avgLE = raColl.Average(ra => ra.LE);
            #endregion

            for (int i = 0; i < raColl.Count(); i++) {
                #region History credit
                Meeting m = meetings.FirstOrDefault(row => Match(raColl[i].Name, row.TME, row.Timer, row.AhCounter,
                    row.Variety, row.TableTopics, row.GE, row.IE1, row.IE2, row.IE3, row.LE));
                List<string> jobTitles = new List<string>(){"TME", "Timer", "AhCounter",
                    "Variety", "TableTopics", "GE", "IE1", "IE2", "IE3", "LE" };
                raColl[i].LastAssignment = WhichYouWere(raColl[i].Name, m, jobTitles) ?? "NONE";
                raColl[i].LastAssignAt = m?.DATE ?? DateTime.Now.AddYears(-2);
                m = meetings.FirstOrDefault(row => Match(raColl[i].Name, row.Speaker1, row.Speaker2, row.Speaker3, row.Speaker4));
                List<string> speechOrders = new List<string>() { "Speaker1", "Speaker2", "Speaker3", "Speaker4" };
                string order = WhichYouWere(raColl[i].Name, m, speechOrders)?.Replace("Speaker", "");
                if (order != null)
                    raColl[i].LastSpeech = m.GetType().GetProperty($"Prj{order}").GetValue(m, null).ToString();
                else
                    raColl[i].LastSpeech = "NONE";
                raColl[i].LastSpeechAt = m?.DATE ?? DateTime.Now.AddYears(-2);

                raColl[i].HistoryCredit = _hw.Base * (_hw.TME * raColl[i].TME / avgTME) +
                             _hw.Base * (_hw.Timer * raColl[i].Timer / avgTimer) +
                             _hw.Base * (_hw.AhCounter * raColl[i].AhCounter / avgAhCounter) +
                             _hw.Base * (_hw.Variety * raColl[i].Variety / avgVariety) +
                             _hw.Base * (_hw.TableTopics * raColl[i].TableTopics / avgTableTopics) +
                             _hw.Base * (_hw.GE * raColl[i].GE / avgGE) +
                             _hw.Base * (_hw.IE * raColl[i].IE / avgIE) +
                             _hw.Base * (_hw.LE * raColl[i].LE / avgLE);
                #endregion

                #region Switch lapse rate
                double lapseRate = 1;
                switch (raColl[i].LastAssignment) {
                    case "TME":
                        lapseRate = _lw.TME;
                        break;
                    case "GE":
                        lapseRate = _lw.GE;
                        break;
                    case "Variety":
                        lapseRate = _lw.Variety;
                        break;
                    case "TableTopics":
                        lapseRate = _lw.TableTopics;
                        break;
                    case "LE":
                        lapseRate = _lw.LE;
                        break;
                    case "IE1":
                    case "IE2":
                    case "IE3":
                        lapseRate = _lw.IE;
                        break;
                    case "Timer":
                        lapseRate = _lw.Timer;
                        break;
                    case "AhCounter":
                        lapseRate = _lw.AhCounter;
                        break;
                    default:
                        break;
                }
                #endregion

                int assignDays = new TimeSpan(DateTime.Now.Ticks - raColl[i].LastAssignAt.Ticks).Days;
                raColl[i].CurrAssignCredit = _lw.Base * (_lw.BufferDays - (assignDays * lapseRate)) / _lw.BufferDays;
                if (raColl[i].CurrAssignCredit < -100)
                    raColl[i].CurrAssignCredit = -100;

                int speechDays = new TimeSpan(DateTime.Now.Ticks - raColl[i].LastSpeechAt.Ticks).Days;
                raColl[i].CurrSpeechCredit = _lw.Base * (_lw.BufferDays - (speechDays * _lw.Speech)) / _lw.BufferDays;
                if (raColl[i].CurrSpeechCredit < -100)
                    raColl[i].CurrSpeechCredit = -100;
                Console.Write($"\rProgress {i}/{raColl.Count() - 1}");
            }
            Console.WriteLine();
            #endregion

            Console.WriteLine("Start Activity sheet writing.");
            for (int i = 0; i < raColl.Count(); i++) {
                #region Cells set val
                activitySheet.GetRow(i + 1).CreateCell(9).CellStyle = style;
                activitySheet.GetRow(i + 1).GetCell(9).SetCellValue(raColl[i].LastAssignment);
                activitySheet.GetRow(i + 1).CreateCell(10).CellStyle = style;
                activitySheet.GetRow(i + 1).GetCell(10).SetCellValue(raColl[i].LastAssignAt.ToString("yyyy-MM-dd"));
                activitySheet.GetRow(i + 1).CreateCell(11).CellStyle = style;
                activitySheet.GetRow(i + 1).GetCell(11).SetCellValue(raColl[i].LastSpeech);
                activitySheet.GetRow(i + 1).CreateCell(12).CellStyle = style;
                activitySheet.GetRow(i + 1).GetCell(12).SetCellValue(raColl[i].LastSpeechAt.ToString("yyyy-MM-dd"));
                activitySheet.GetRow(i + 1).CreateCell(13).CellStyle = style;
                activitySheet.GetRow(i + 1).GetCell(13).SetCellValue(raColl[i].HistoryCredit);
                activitySheet.GetRow(i + 1).CreateCell(14).CellStyle = style;
                activitySheet.GetRow(i + 1).GetCell(14).SetCellValue(raColl[i].CurrAssignCredit);
                activitySheet.GetRow(i + 1).CreateCell(15).CellStyle = style;
                activitySheet.GetRow(i + 1).GetCell(15).SetCellValue(raColl[i].CurrSpeechCredit);
                Console.Write($"\rProgress {i}/{raColl.Count() - 1}");
                #endregion
            }
            Console.WriteLine();
            Console.WriteLine("Finished Activity sheet writing.");

            using (FileStream file = new FileStream(@"Record.xlsx", FileMode.Create, FileAccess.Write)) {
                book.Write(file);
            }

            Console.WriteLine("All finished. Prepare self close.");
            Thread.Sleep(1500);
        }

        private static void LoadWeight()
        {
            if (File.Exists(@".\Weight.ini")) {
                INI ini = new INI(@".\Weight.ini");
                _lw.Base = double.Parse(ini.Read("LastActivity", "Base"));
                _lw.BufferDays = double.Parse(ini.Read("LastActivity", "BufferDays"));
                _lw.TME = double.Parse(ini.Read("LastActivity", "TME"));
                _lw.Timer = double.Parse(ini.Read("LastActivity", "Timer"));
                _lw.AhCounter = double.Parse(ini.Read("LastActivity", "AhCounter"));
                _lw.Variety = double.Parse(ini.Read("LastActivity", "Variety"));
                _lw.TableTopics = double.Parse(ini.Read("LastActivity", "TableTopics"));
                _lw.GE = double.Parse(ini.Read("LastActivity", "GE"));
                _lw.IE = double.Parse(ini.Read("LastActivity", "IE"));
                _lw.LE = double.Parse(ini.Read("LastActivity", "LE"));
                _lw.Speech = double.Parse(ini.Read("LastActivity", "Speech"));

                _hw.Base = double.Parse(ini.Read("History", "Base"));
                _hw.TME = double.Parse(ini.Read("History", "TME"));
                _hw.Timer = double.Parse(ini.Read("History", "Timer"));
                _hw.AhCounter = double.Parse(ini.Read("History", "AhCounter"));
                _hw.Variety = double.Parse(ini.Read("History", "Variety"));
                _hw.TableTopics = double.Parse(ini.Read("History", "TableTopics"));
                _hw.GE = double.Parse(ini.Read("History", "GE"));
                _hw.IE = double.Parse(ini.Read("History", "IE"));
                _hw.LE = double.Parse(ini.Read("History", "LE"));
            }
        }

        private static bool Match(string name, params string[] takers)
        {
            return takers.Any(taker => taker == name);
        }

        private static string WhichYouWere(string name, Meeting m, List<string> things)
        {
            string ret = null;

            var propNames = typeof(Meeting).GetProperties().Where(
                prop => things.Contains(prop.Name)).Select(prop => prop.Name).ToList();

            propNames.ForEach(pn => {
                string taker = m?.GetType().GetProperty(pn)?.GetValue(m, null)?.ToString();
                if (name == taker)
                    ret = pn;
            });

            return ret;
        }
    }
}
