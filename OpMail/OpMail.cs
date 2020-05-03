using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;

namespace OscarSoft
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("OpMail")]
    public class OpMail
    {

        private readonly string _version = "1.0.0";
        public string from = "";
        public string user = "";
        public string password = "";
        public string subjet = "";
        public string body = "";
        public bool bodyHtml = true;
        public bool clear = false;
        public bool fileSize = false;
        public bool notification = false;
        public int timeout = 100000;
        private readonly List<string> _emailTo = new List<string>();
        private readonly List<string> _emailCc = new List<string>();
        private readonly List<string> _emailBcc = new List<string>();
        private readonly List<string> _filesList = new List<string>();
        public string server = "";
        public int port = 587;
        public bool ssl = true;
        public int priority = 0;
        private MailMessage mailMessage = new MailMessage();
        private Attachment fileAdd;


        //Main Methods
        public bool Send()
        {
            bool flag = false;
            Error = "";
            mailMessage = new MailMessage(from, To, subjet, body)
            {
                IsBodyHtml = bodyHtml,
                BodyEncoding = Encoding.UTF8,
                SubjectEncoding = Encoding.UTF8,
                Priority = setPriority(priority)


            };
            if (notification)
            {
                mailMessage.Headers.Add("Disposition-Notification-To", user);
                mailMessage.Headers.Add("Return-Receipt-To", user);
            }
            if (_emailCc.Count > 0)
                _emailCc.ForEach((Action<string>)(email => mailMessage.CC.Add(email)));

            if (_emailBcc.Count > 0)
                _emailBcc.ForEach((Action<string>)(email => mailMessage.Bcc.Add(email)));

            if (_filesList.Count > 0)

                _filesList.ForEach((Action<string>)(file =>
                {
                    fileAdd = new Attachment(file, "application/octet-stream");
                    mailMessage.Attachments.Add(fileAdd);
                }));


            SmtpClient smtpClient = new SmtpClient(server)
            {
                EnableSsl = ssl,
                Host = server,
                Port = port,
                UseDefaultCredentials = false,
                Timeout = timeout,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                Credentials = new NetworkCredential(user, password)

            };

            try
            {
                smtpClient.Send(mailMessage);
                if (clear)
                    Clear();
            }
            catch (Exception ex)
            {
                Error = ex.ToString();
            }

            smtpClient.Dispose();
            if (_filesList.Count > 0)
                fileAdd.Dispose();

            mailMessage.Attachments.Clear();
            if (Error.Length <= 0)
                flag = true;
            return flag;
        }


        public string Version
        {
            get
            {
                return _version;
            }
        }

        public string Error { get; private set; } = "";

        public Decimal Size { get; private set; } = new Decimal();

        public void Clear()
        {
            from = "";
            subjet = "";
            body = "";
            Error = "";
            Size = decimal.Zero;
            ClearTo();
            ClearCc();
            ClearBcc();
            ClearAttachments();
        }

        private MailPriority setPriority(int num)
        {
            switch (num)
            {
                case 0:
                    return MailPriority.Normal;
                case 1:
                    return MailPriority.Low;
                case 2:
                    return MailPriority.High;
                default:
                    return MailPriority.Normal;
            }
        }

        public string Attachments
        {
            get
            {
                return setListToString(_filesList);
            }
        }

        public decimal AddAttachments(string pathFile)
        {
            _filesList.Add(pathFile);
            if (fileSize)
                Size += getFileSize(pathFile);
            else
                Size = decimal.Zero;
            return Size;
        }

        public void ClearAttachments()
        {
            _filesList.Clear();
        }

        public int CountAttachments
        {
            get
            {
                return _filesList.Count;
            }
        }

        public string Cc
        {
            get
            {
                return setListToString(_emailCc);
            }
        }

        public void AddCc(string email)
        {
            _emailCc.Add(email);
        }

        public void ClearCc()
        {
            _emailCc.Clear();
        }

        public int CountCc
        {
            get
            {
                return _emailCc.Count;
            }
        }

        public string Bcc
        {
            get
            {
                return setListToString(_emailBcc);
            }
        }

        public void AddBcc(string email)
        {
            _emailBcc.Add(email);
        }

        public void ClearBcc()
        {
            _emailBcc.Clear();
        }

        public int CountBcc
        {
            get
            {
                return _emailBcc.Count;
            }
        }

        public string To
        {
            get
            {
                return setListToString(_emailTo);
            }
        }

        public void AddTo(string email)
        {
            _emailTo.Add(email);
        }

        public void ClearTo()
        {
            _emailTo.Clear();
        }

        public int CountTo
        {
            get
            {
                return _emailTo.Count;
            }
        }

        private string setListToString(List<string> list)
        {
            string listToString = "";
            int index = 0;
            list.ForEach(item =>
            {
                ++index;
                if (index > 1)
                    listToString += ", ";
                listToString += item;
            });
            return listToString;
        }

        private decimal getFileSize(string pathFile)
        {
            return new FileInfo(pathFile).Length / 1024L;
        }


    }
}
