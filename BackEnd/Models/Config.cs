using System;
using System.Collections.Generic;

namespace NBS.MailBox.BackEnd.Models
{
    public class Config
    {
        public static Type MailBoxApp { get; internal set; }
        public List<MailBoxApp> MailBoxApps { get; set; }
    }

    public class MailBoxApp
    {
        public string Name { get; set; }
        public string Address { get; set; }
        public string AppAddress { get; set; }
        public string SpWebBaseUrl { get; set; }
        public string SpDocLibId { get; set; }
        public string SpListId { get; set; }
        public List<User> Users { get; set; }
    }

    public class User
    {
        public string Upn { get; set; }
        public string Role { get; set; }
    }
}
