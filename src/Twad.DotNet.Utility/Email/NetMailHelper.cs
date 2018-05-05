using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Text;

namespace Twad.DotNet.Utility.Email
{
    public class NetMailHelper
    {
        public static bool Send(SendInfo entity)
        {
            try
            {
                MailMessage email = new MailMessage();
                //添加收件人
                foreach (var item in entity?.SendAddress)
                {
                    email.To.Add(item);
                }
                //添加抄送
                foreach (var item in entity?.CopyAddress)
                {
                    email.CC.Add(item);
                }
                email.From = new MailAddress(entity.FromAddress, entity.FromName, Encoding.UTF8);// 发件人地址，姓名，编码  
                email.Subject = entity.Title;
                email.SubjectEncoding = Encoding.UTF8;//邮件标题编码  
                email.Body = entity.Body;//邮件内容    
                email.BodyEncoding = Encoding.UTF8;//邮件内容编码    
                email.IsBodyHtml = true;//是否是HTML邮件    
                email.Priority = MailPriority.Normal;//邮件优先级    
                SmtpClient client = new SmtpClient();
                client.Credentials = new System.Net.NetworkCredential(entity.EmailAccount.Account, entity.EmailAccount.PassWord);
                client.Host = entity.EmailAccount.Host;
                client.Send(email);
                return true;
            }
            catch (Exception ex)
            {
                throw ex;

            }

        }
    }


    public class SendInfo
    {
        public SendInfo() { EmailAccount = new EmailAccount(); }
        /// <summary>
        /// 发送到地址
        /// </summary>
        public List<string> SendAddress { set; get; }
        /// <summary>
        /// 抄送地址
        /// </summary>
        public List<string> CopyAddress { set; get; }
        /// <summary>
        /// 发送者的地址
        /// </summary>
        public string FromAddress { set; get; }
        /// <summary>
        /// 发送者名称
        /// </summary>
        public string FromName { set; get; }

        /// <summary>
        /// 邮件标题
        /// </summary>
        public string Title { set; get; }
        /// <summary>
        /// 邮件正文
        /// </summary>
        public string Body { set; get; }
        /// <summary>
        /// 邮件账户
        /// </summary>
        public EmailAccount EmailAccount { set; get; }

    }
    /// <summary>
    /// 邮件账户
    /// </summary>
    public class EmailAccount
    {
        /// <summary>
        /// 邮件服务地址
        /// </summary>
        public string Host { set; get; }
        /// <summary>
        /// 邮件账户
        /// </summary>
        public string Account { set; get; }
        /// <summary>
        /// 邮件密码
        /// </summary>
        public string PassWord { set; get; }
    }
}

