using NBS.MailBox.BackEnd.Models;
using System.Threading.Tasks;

namespace NBS.MailBox.BackEnd.Services
{
    public interface IMailService
    {
        Task SendEmailAsync(Message message);
    }
}
