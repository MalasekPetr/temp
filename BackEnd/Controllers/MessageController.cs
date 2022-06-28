using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using MimeKit;
using NBS.MailBox.BackEnd.Models;
using NBS.MailBox.BackEnd.Services;
using System;
using System.IO;
using System.Threading.Tasks;
using MimeKit.Utils;

namespace NBS.MailBox.BackEnd.Controllers
{
    //[Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class MessageController : ControllerBase
    {
        private readonly ILogger<MessageController> _logger;
        private readonly IMailService mailService;

        public MessageController(
            ILogger<MessageController> logger,
            IMailService mailService
            )
        {
            _logger = logger;
            this.mailService = mailService;
        }

        // POST api/<ConfigController>/parse
        [EnableCors(PolicyName = "AllowedOrigins")]
        [RequestFormLimits(MultipartBodyLengthLimit = 26214400)]
        [HttpPost("parse")]
        public IActionResult ParseMessage([FromBody] EncodedMessage encodedmessage)
        {
            try
            {
                MemoryStream ms = new(Convert.FromBase64String(encodedmessage.Msgcontent));
                var msg = new Message(MimeMessage.Load(ms));
                return new JsonResult(msg);
            }
            catch (Exception ex)
            {
                _logger.LogError("Unexpected error!\r\nMessage: {0}\r\nStackTrace: {1}", ex.Message, ex.StackTrace);
                return StatusCode(500);
            }
        }

        // POST api/<ConfigController>/send
        [EnableCors(PolicyName = "AllowedOrigins")]
        [RequestFormLimits(MultipartBodyLengthLimit = 26214400)]
        [HttpPost("send")]
        public async Task<IActionResult> SendMail([FromBody] Message message)
        {
            try
            {
                if (string.IsNullOrEmpty(message.MessageId))
                {
                    int atIndex = message.Sender.LastIndexOf("@") + 1;
                    string domain = message.Sender.Substring(atIndex, message.Sender.Length - atIndex - 1);
                    message.MessageId = MimeUtils.GenerateMessageId(domain);
                }
                await mailService.SendEmailAsync(message);
                return new JsonResult(message);
            }
            catch (Exception ex)
            {
                _logger.LogError("Unexpected error!\r\nMessage: {0}\r\nStackTrace: {1}", ex.Message, ex.StackTrace);
                return StatusCode(500);
            }
        }
    }
}
