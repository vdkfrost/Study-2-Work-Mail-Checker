using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace S2W_Mail_Checker
{
    public class Packet
    {
        public string mailSender, mailText, mailDateSend;

        public int userId;
        public string userName, userMail;

        public int labId, labOwnerId;
        public string labName, labOwnerName, labOwnerMail;
        public List<Packet> labGroups;

        public int groupId, groupLabId;
        public string groupName;
        public List<Packet> groupUsers;
        public Packet(string mailSender, string mailText, string mailDateSend)
        {
            this.mailSender = mailSender;
            this.mailText = mailText;
            this.mailDateSend = mailDateSend;
        }
        public Packet(int userId, string userName, string userMail)
        {
            this.userId = userId;
            this.userName = userName;
            this.userMail = userMail;
        }

        public Packet(int labId, string labName, int labOwnerId, string labOwnerName, string labOwnerMail, List<Packet> labGroups)
        {
            this.labId = labId;
            this.labName = labName;
            this.labOwnerId = labOwnerId;
            this.labOwnerName = labOwnerName;
            this.labOwnerMail = labOwnerMail;
            this.labGroups = labGroups;
        }
        public Packet(int labId, string labName, int labOwnerId, string labOwnerName, string labOwnerMail)
        {
            this.labId = labId;
            this.labName = labName;
            this.labOwnerId = labOwnerId;
            this.labOwnerName = labOwnerName;
            this.labOwnerMail = labOwnerMail;
        }
        public Packet(int groupId, string groupName, int groupLabId, List<Packet> groupUsers)
        {
            this.groupId = groupId;
            this.groupName = groupName;
            this.groupLabId = groupLabId;
            this.groupUsers = groupUsers;
        }
    }
}
