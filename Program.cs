using System;
using System.IO;
using System.Web;
using System.Net;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Reflection;
using FluentFTP;

namespace MailTo {
    class Program {
        static int Main(string[] args) {
            TextWriter oldOut = Console.Out;

            try {
                string exeDir = new FileInfo(Assembly.GetEntryAssembly().Location).Directory.ToString();
                using (var ostrm = new FileStream(exeDir + @"\\mailto.log", FileMode.Create, FileAccess.Write)) {
                    using (var writer = new StreamWriter(ostrm)) {
                        Console.SetOut(writer);

                        try {

                            if (args.Length != 2) {
                                throw new ArgumentException("Ivalid aguments count.");
                            }

                            string file = args[0];
                            string company = args[1].ToLower();

                            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                            var lines = File.ReadAllLines(file, Encoding.GetEncoding(1250));
                            if (lines.Length < 4) {
                                throw new Exception("Špatný formát souboru mailto.txt.");
                            }

                            Console.WriteLine("---------------- Passed file ----------------");
                            Console.WriteLine(string.Join(Environment.NewLine, lines));

                            var mailto = new Mailto(lines, company);
                            var mailtoUrl = string.Format(
                                "mailto://{0}",
                                mailto.GetMailtoUrl()
                            );

                            var process = new Process() {
                                StartInfo = new ProcessStartInfo {
                                    FileName = "cmd.exe",
                                    Arguments = String.Format(
                                        "/c START \"{0}\"  \"{1}\"",
                                        "Mailto",
                                        mailtoUrl
                                    ),
                                    UseShellExecute = true,
                                    RedirectStandardOutput = false,
                                    WindowStyle = ProcessWindowStyle.Hidden
                                }
                            };

                            Console.WriteLine(Environment.NewLine + "---------------- Process ----------------");
                            Console.WriteLine(process.StartInfo.ToString());
                            process.Start();

                        } catch (Exception e) {
                            Console.WriteLine("Error: " + e.Message + "\nPodrobnosti:" + e.ToString());
                        }
                    }
                }
                return 0;

            } catch (Exception e) {
                Console.SetOut(oldOut);
                Console.WriteLine(e.Message);
                return -1;

            } finally {
                Console.SetOut(oldOut);
            }
        }
    }

    public static class UploadFtp {
        private const string host = "pemalbc.savana-hosting.cz";

        private static string getHostUrl(string company) {
            if (company == "pema") {
                return "https://velkoobchoddrogerie.cz";
            } else if (company == "lipa") {
                return "https://velkoobchodpapirem.cz";
            }
            throw new ArgumentException("Allowed values pema and lipa");
        }

        private static string randomString(int length) {
            Random random = new Random();
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        public static string uploadFile(string localFile, string company) {
            var newFileName = string.Format(
                "{0}-{1}{2}",
                Path.GetFileNameWithoutExtension(localFile),
                UploadFtp.randomString(12),
                Path.GetExtension(localFile)
            );

            string userName = company + "docs";
            using (var client = new FtpClient()) {
                var serverFile = string.Format("/{0}", newFileName);
                client.Host = UploadFtp.host;
                client.Credentials = new NetworkCredential(userName, Secret.get());
                client.DataConnectionType = FtpDataConnectionType.AutoActive;
                client.Connect();
                var status = client.UploadFile(localFile, serverFile, FtpRemoteExists.Overwrite, false, FtpVerify.Throw);
                Console.WriteLine("\nUpload file status:" + status.ToString());
            }

            return string.Format(
                "{0}/docs/{1}",
                UploadFtp.getHostUrl(company),
                newFileName
            );
        }
    }

    public class Mailto {
        protected string receiver;
        protected string subject;
        protected string attachment;
        protected string body = "";
        protected string company;

        public Mailto(string[] lines, string company) {
            this.company = company;
            receiver = lines[0].Trim();
            subject = lines[1].Trim();
            attachment = lines[2].Trim();
            for (var i = 3; i < lines.Length; i++) {
                body += lines[i] + Environment.NewLine;
            }
        }

        public string GetMailtoUrl() {
            if (File.Exists(this.attachment)) {
                var url = UploadFtp.uploadFile(this.attachment, this.company);
                this.insertAttachmentUrl(url);
            }

            return String.Format(
                "{0}?subject={1}&body={2}",
                this.receiver,
                HttpUtility.UrlEncode(this.subject),
                HttpUtility.UrlEncode(this.body)
            );
        }

        private void insertAttachmentUrl(string url) {
            var body = this.body;
            this.body = body.Replace("<attachment_url>", url);
        }
    }
}
