using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using ImapX;
using ImapX.Enums;
using System.Timers;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;

namespace S2W_Mail_Checker
{
    class Program
    {
        /* BLOCK: connections */
        public static string connectionString = ConfigurationManager.ConnectionStrings["baseConnectionString"].ConnectionString;
        public static SqlConnection connection = new SqlConnection(connectionString);
        public static ImapClient client;

        /* BLOCK: application values */
        public static string adminMail;
        public static int maxAttachmentSize, maxTextSize;
        public static int inboxCheckTimeout, adminCheckTimeout, timeoutIfCheckingInbox, timeoutIfCheckingAdmin;
        public static List<string> mailKeyWords = new List<string>();

        public static string imapServerName, serverPassword, serverLogin;
        public static int imapServerPort;

        public static bool markAdminMessagesAsSeen;

        /* BLOCK: other */
        public static bool checking = false;
        static void Main(string[] args)
        {
            getInfoFromConfig(Directory.GetCurrentDirectory() + "//config.cfg");

            Console.SetWindowSize(120, 40);

            System.Timers.Timer inboxMailChecker = new System.Timers.Timer(inboxCheckTimeout);
            System.Timers.Timer adminMailChecker = new System.Timers.Timer(adminCheckTimeout);

            client = new ImapX.ImapClient(imapServerName, imapServerPort, true);

            WriteLine("- Подключаюсь к IMAP-серверу: ", false, false);
            WriteLine("- Сервер - " + imapServerName + ":" + imapServerPort.ToString() + "\n", false, false);
            WriteLine("- Входные данные: ", false, false);
            WriteLine("- Login: " + serverLogin + "   Password: " + serverPassword + "\n", false, false);
            client.Connect();
            WriteLine("- Вошел на почту", false, false);
            client.Login(serverLogin, serverPassword);
            

            checkInboxMail(inboxMailChecker, null);
            inboxMailChecker.Elapsed += new ElapsedEventHandler(checkInboxMail);

            checkForAdminMail(adminMailChecker, null);
            adminMailChecker.Elapsed += new ElapsedEventHandler(checkForAdminMail);

            Console.ReadLine();
        }
        public static void checkInboxMail(object sender, ElapsedEventArgs e)
        {
            (sender as Timer).Stop();

            if (!checking)
            {
                checking = true;
                (sender as Timer).Interval = inboxCheckTimeout;

                WriteLine("- Запускаю поиск непрочитанных сообщений во входящих сообщениях..", false, false);
                var messages = client.Folders.Inbox.Search("UNSEEN").ToList();
                WriteLine("- Количество непрочитанных сообщений: " + messages.Count.ToString() + (messages.Count == 0 ? "\n" : ""), false, (messages.Count == 0 ? false : true));
                foreach (var message in messages)
                {
                    try
                    {
                        Packet author = new Packet(null, null, null);

                        connection.Open();
                        SqlCommand checkSender = new SqlCommand("EXEC CheckSender N'" + message.From.Address.ToString() + "'", connection);
                        SqlDataReader checkSenderReader = checkSender.ExecuteReader();
                        if (checkSenderReader.Read())
                        {
                            author = new Packet(checkSenderReader.GetInt32(0), checkSenderReader.GetString(1), checkSenderReader.GetString(2));
                            connection.Close();

                            if (message.Subject.IndexOf("[#") != -1 && message.Subject.IndexOf(']', message.Subject.IndexOf("[#")) != -1)
                            {
                                string groupIdStringTemp = message.Subject.Substring(message.Subject.IndexOf("[#") + 2, message.Subject.IndexOf("]", message.Subject.IndexOf("[#") + 2) - message.Subject.IndexOf("[#") - 2);
                                int groupIdIntTemp = 0;
                                string groupId = "";

                                if (int.TryParse(groupIdStringTemp, out groupIdIntTemp))
                                {
                                    groupId = groupIdIntTemp.ToString();

                                    connection.Open();
                                    SqlCommand checkUserRelation = new SqlCommand("EXEC CheckUserRelation '" + groupId + "', '" + author.userId.ToString() + "'", connection);
                                    SqlDataReader checkUserRelationReader = checkUserRelation.ExecuteReader();
                                    if (checkUserRelationReader.Read())
                                    {
                                        connection.Close();
                                        string mailSender = message.From.Address;
                                        string mailText = formatMessage(message.Body.Text);

                                        if (mailText.Length <= maxTextSize)
                                        {
                                            if (mailText == "" || mailText == " ")
                                            {
                                                // Ошибка - пустое сообщение
                                                WriteLine("Найдено пустое сообщение. Сообщение перемещено в папку для Администратора", false, true);
                                                newMailMessage("empty message error", author, null, message);
                                                message.MoveTo(client.Folders["For Admin"]);
                                            }
                                            else
                                            {
                                                List<string> attachments = new List<string>();
                                                List<string> attachmentSizeError = new List<string>();
                                                foreach (ImapX.Attachment attach in message.Attachments.Concat(message.EmbeddedResources))
                                                {
                                                    if (Convert.ToInt32(attach.FileSize * 0.73 / 1024) > maxAttachmentSize)
                                                        attachmentSizeError.Add(attach.FileName);
                                                    else
                                                    {
                                                        attach.Download();
                                                        string folderPath = System.Environment.CurrentDirectory.ToString() + "\\" + mailSender;
                                                        if (!System.IO.Directory.Exists(folderPath))
                                                            System.IO.Directory.CreateDirectory(folderPath);
                                                        attach.Save(folderPath);
                                                        attachments.Add(folderPath + "\\" + attach.FileName);
                                                    }
                                                }
                                                if (attachmentSizeError.Count != 0)
                                                {
                                                    // Ошибка - размер вложения более 3 Мб
                                                    WriteLine("Найдено как минимум 1 сообщение с размером более 3 Мб", false, true);
                                                    newMailMessage("attachment size error", author, attachmentSizeError, message);
                                                }
                                                newMailMessage(new Packet(mailSender, mailText, message.Date.ToString()), author, groupId, attachments);
                                                message.Remove();
                                            }
                                        }
                                        else
                                        {
                                            WriteLine("Сообщение больше допустимого размера. Сообщение перемещено в папку для Администратора", false, true);
                                            newMailMessage("text size error", author, null, message);
                                            message.MoveTo(client.Folders["For Admin"]);
                                        }
                                    }
                                    else
                                    {
                                        connection.Close();
                                        // Ошибка - у пользователя нет доступа к беседе
                                        WriteLine("У пользователя " + author.userId.ToString() + " нет доступа к беседе " + groupId + ". Сообщение перемещено в папку для Администратора", false, true);
                                        newMailMessage("access error", author, null, message);
                                        message.MoveTo(client.Folders["For Admin"]);
                                    }
                                }
                                else
                                {
                                    // Ошибка - косяк в теме беседы
                                    WriteLine("Не найдена тема беседы. Сообщение перемещено в папку для Администратора", false, true);
                                    newMailMessage("subject error", author, null, message);
                                    message.MoveTo(client.Folders["For Admin"]);
                                }
                            }
                            else
                            {
                                // Ошибка - тема не содержит номер беседы
                                WriteLine("Не найдена тема беседы. Сообщение перемещено в папку для Администратора", false, true);
                                newMailMessage("subject error", author, null, message);
                                message.MoveTo(client.Folders["For Admin"]);
                            }
                        }
                        else
                        {
                            connection.Close();
                            message.Remove();
                        }
                        Console.WriteLine();
                    }
                    catch (Exception ex)
                    {
                        WriteLine(ex.ToString() + "\n", true, true);
                    }
                }
            }
            else
                (sender as Timer).Interval = timeoutIfCheckingAdmin;
            checking = false;
            (sender as Timer).Start();
        }
        public static void checkForAdminMail(object sender, ElapsedEventArgs e)
        {
            (sender as Timer).Stop();
            if (!checking)
            {
                checking = true;
                (sender as Timer).Interval = adminCheckTimeout;
                WriteLine("- Запускаю поиск непрочитанных сообщений в сообщениях для Администратора..", false, false);
                var messages = client.Folders["For Admin"].Search("UNSEEN").ToList();
                WriteLine("- Количество непрочитанных сообщений: " + messages.Count.ToString() + (messages.Count == 0 ? "\n" : ""), false, (messages.Count == 0 ? false : true));
                if (messages.Count > 0)
                {
                    newMailMessage(messages.Count);
                    if (markAdminMessagesAsSeen)
                        foreach (var message in messages)
                            message.Seen = true;
                }
            }
            else
                (sender as Timer).Interval = timeoutIfCheckingInbox;
            checking = false;
            (sender as Timer).Start();
        }

        /* METHOD: new comment message */
        public static void newMailMessage(Packet message, Packet author, string groupId, List<string> attachmentsPaths)
        {
            string attachmentsConcatenated = "";
            for (int i = 0; i < attachmentsPaths.Count; i++)
                attachmentsConcatenated += i != attachmentsPaths.Count - 1 ? attachmentsPaths[i] + "•" : attachmentsPaths[i];
            string messageId = "";

            string labName = "";

            connection.Open();
            SqlCommand getLabName = new SqlCommand("SELECT lbs.[name] " +
            "FROM [dbo].[groups] grps " +
            "JOIN [dbo].[labs] lbs " +
            "ON grps.[labId] = lbs.[id] " +
            "WHERE grps.[id] = '" + groupId + "'", connection);
            SqlDataReader getLabNameReader = getLabName.ExecuteReader();
            if (getLabNameReader.Read())
                labName = getLabNameReader.GetString(0);
            connection.Close();

            string receivers = getReceivers(groupId, author.userId.ToString());
            connection.Open();
            SqlCommand insertMailMessage = new SqlCommand("EXEC AddNewMailMessageAndShow '0', 'new comment', '" + message.mailSender + "', N'" + message.mailText + "', '" + message.mailDateSend + "', '" + receivers + "', ' ', N'" + author.userName + "', '" + author.userMail + "', '" + groupId + "', N'" + labName + "'", connection);
            SqlDataReader messageIdReader = insertMailMessage.ExecuteReader();
            if (messageIdReader.Read())
                messageId = messageIdReader.GetInt32(0).ToString();
            connection.Close();

            if (attachmentsConcatenated != "")
            {
                connection.Open();
                SqlCommand insertAttach = new SqlCommand("EXEC AddNewAttachment '" + messageId + "', N'" + attachmentsConcatenated + "'", connection);
                insertAttach.ExecuteNonQuery();
                connection.Close();
                WriteLine("- Добавил в очередь на отправку " + attachmentsPaths.Count.ToString() + " вложений", true, true);
            }

            WriteLine("- Добавил в очередь сообщение для " + author.userMail + " о новом уведомлении", false, true);
            WriteLine("- Тип: new comment", true, true);
        }

        /* METHOD: error message */
        public static void newMailMessage(string type, Packet author, List<string> objects, ImapX.Message message)
        {
            string attachmentsErrors = "";
            if (objects != null)
                for (int i = 0; i < objects.Count; i++)
                    attachmentsErrors += i != objects.Count - 1 ? objects[i] + ", " : objects[i];

            connection.Open();
            SqlCommand insertError = new SqlCommand("EXEC AddNewMailMessageAndShow '1', '" + type + "', '" + author.userMail + "', " 
                + (type != "text size error" ? "N'" + formatMessage(message.Body.Text) + "'": "''") + ", '" + message.Date.ToString() + "', '"
                + author.userMail + "', N'" + attachmentsErrors + "', N'" + author.userName + "', '" + author.userMail + "', NULL, ''", connection);
            insertError.ExecuteNonQuery();
            connection.Close();

            WriteLine("- Добавил в очередь сообщение для " + author.userMail + " об ошибке отправки.", false, true);
            WriteLine("- Тип: " + type, true, true);
        }

        /* METHOD: message for admin */
        public static void newMailMessage(int countOfMessages)
        {
            connection.Open();
            SqlCommand insertMailMessage = new SqlCommand("EXEC AddNewMailMessageAndShow '0', 'for admin', '', '" + countOfMessages.ToString() + "', '', '" + adminMail + "', '', '', '', NULL, ''", connection);
            insertMailMessage.ExecuteNonQuery();
            connection.Close();
            
            WriteLine("- Добавил в очередь сообщение для Администратора о " + countOfMessages.ToString() + " непрочитанных сообщений в папке For Admin", false, true);
            WriteLine("- Тип: for admin", true, true);
        }

        /* BLOCK: service methods */
        public static void WriteLine(string text, bool makeNewLine, bool overriden)
        {
            if (overriden)
            {
                Console.BackgroundColor = ConsoleColor.White;
                Console.ForegroundColor = ConsoleColor.Black;
            }
            Console.WriteLine(DateTime.Now.ToString() + " " + text);
            if (overriden)
            {
                Console.BackgroundColor = ConsoleColor.Black;
                Console.ForegroundColor = ConsoleColor.White;
            }
            if (makeNewLine)
                Console.WriteLine();
        }
        public static string getReceivers(string groupId, string authorId)
        {
            string result = "";

            connection.Open();
            SqlCommand getReceivers = new SqlCommand("EXEC GetReceivers '" + groupId + "', '" + authorId + "'", connection);
            SqlDataReader getReceiversReader = getReceivers.ExecuteReader();
            while (getReceiversReader.Read())
                result += getReceiversReader.GetString(0) + " ";
            connection.Close();

            result = result.Remove(result.Length - 1, 1);
            return result;
        }
        public static string formatMessage(string message)
        {
            if (message.IndexOf("Уведомитель Study 2 Work") != -1)
                message = message.Substring(0, message.IndexOf("Уведомитель Study 2 Work"));

            foreach (string keyWord in mailKeyWords)
                if (message.LastIndexOf(keyWord) != -1)
                    message = message.Substring(0, message.LastIndexOf(keyWord));

            if (message.LastIndexOf('\n') != -1)
                message = message.Substring(0, message.LastIndexOf('\n') + 1);

            message = message.Replace("\r\n", " ").Replace("\t", "").Replace("\n", " ");
            for (int i = 0; i < message.Length - 1; i++)
                if (message[i] == ' ' && message[i + 1] == ' ')
                {
                    message = message.Remove(i + 1, 1);
                    i--;
                }

            if (message.LastIndexOf(' ') == message.Length - 1 && message != "")
                message = message.Remove(message.Length - 1);
            return message;
        }

        public static void getInfoFromConfig(string configPath)
        {
            List<Setting> configData = new List<Setting>();
            FileStream config = new FileStream(configPath, FileMode.Open, FileAccess.Read);
            StreamReader configReader = new StreamReader(config, Encoding.UTF8);
            foreach (string line in configReader.ReadToEnd().Replace("\r\n", "•").Split('•'))
                if (formatMessage(line) != "" && formatMessage(line) != " ")
                {
                    string[] splittedLine = line.Split('=');
                    if (splittedLine.Length != 1)
                        configData.Add(new Setting(splittedLine[0], splittedLine[1]));
                    else
                        configData.Add(new Setting(line, null));
                }
            string objectOptionSwitcher = "";

            foreach (Setting set in configData)
            {
                if (set.option.IndexOf('[') != -1)
                    objectOptionSwitcher = set.option;
                else
                    switch (objectOptionSwitcher)
                    {
                        case "[IMAP SERVER]":
                            switch (set.option)
                            {
                                case "name":
                                    imapServerName = set.value;
                                    break;
                                case "port":
                                    imapServerPort = Convert.ToInt32(set.value);
                                    break;
                                case "login":
                                    serverLogin = set.value;
                                    break;
                                case "password":
                                    serverPassword = set.value;
                                    break;
                            }
                            break;
                        case "[ADMIN]":
                            switch (set.option)
                            {
                                case "post":
                                    adminMail = set.value;
                                    break;
                                case "checkTimeout":
                                    adminCheckTimeout = Convert.ToInt32(set.value);
                                    break;
                            }
                            break;
                        case "[APP]":
                            switch (set.option)
                            {
                                case "timeout":
                                    inboxCheckTimeout = Convert.ToInt32(set.value);
                                    break;
                                case "timeoutIfCheckingInbox":
                                    timeoutIfCheckingInbox = Convert.ToInt32(set.value);
                                    break;
                                case "timeoutIfCheckingAdmin":
                                    timeoutIfCheckingAdmin = Convert.ToInt32(set.value);
                                    break;
                                case "markAdminMessagesAsSeen":
                                    markAdminMessagesAsSeen = set.value == "true" ? true : false;
                                    break;  
                                case "maxAttachmentSize":
                                    maxAttachmentSize = Convert.ToInt32(set.value);
                                    break;
                                case "maxTextSize":
                                    maxTextSize = Convert.ToInt32(set.value);
                                    break;
                                case "mailKeyWords":
                                    mailKeyWords.AddRange(set.value.Split('|').ToList());
                                    break;
                            }
                            break;
                    }
            }
        }
    }
}
