using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using NBS.MailBox.BackEnd.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace NBS.MailBox.BackEnd.Controllers
{
    //[Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class ConfigController : ControllerBase
    {
        private readonly ILogger<MessageController> _logger;
        private readonly ConfigStore _configstore;

        public ConfigController(
            ILogger<MessageController> logger, 
            IOptions<ConfigStore> configstore
            )
        {
            _logger = logger;
            _configstore = configstore.Value;
        }

        // GET: api/<ConfigController>
        [EnableCors(PolicyName = "AllowedOrigins")]
        [HttpGet]
        public IActionResult Get()
        {
            try
            {
                ReadOnlySpan<byte> jsonReadOnlySpan = System.IO.File.ReadAllBytes(_configstore.FilePath);
                var reader = new Utf8JsonReader(jsonReadOnlySpan);
                var config = JsonSerializer.Deserialize<Config>(ref reader);
                return new JsonResult(config);
            }
            catch (Exception ex)
            {
                _logger.LogError("Unexpected error!\r\nMessage: {0}\r\nStackTrace: {1}", ex.Message, ex.StackTrace);
                return StatusCode(500);
            }
        }

        // GET api/<ConfigController>/address
        [EnableCors(PolicyName = "AllowedOrigins")]
        [HttpGet("{address}")]
        public IActionResult Get(string address)
        {
            try
            {
                ReadOnlySpan<byte> jsonReadOnlySpan = System.IO.File.ReadAllBytes(_configstore.FilePath);
                var reader = new Utf8JsonReader(jsonReadOnlySpan);
                var config = JsonSerializer.Deserialize<Config>(ref reader);
                var mailBoxApp = config.MailBoxApps.FirstOrDefault<MailBoxApp>(a => a.Address == address);
                if (mailBoxApp != null)
                {
                    return new JsonResult(mailBoxApp);
                }
                else
                {
                    _logger.LogWarning("Address not found in config store.");
                    return StatusCode(500);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError("Unexpected error!\r\nMessage: {0}\r\nStackTrace: {1}", ex.Message, ex.StackTrace);
                return StatusCode(500);
            }
        }

        // POST api/<ConfigController>
        [EnableCors(PolicyName = "AllowedOrigins")]
        [HttpPost]
        public IActionResult Post([FromBody] MailBoxApp value)
        {
            try
            {
                ReadOnlySpan<byte> jsonReadOnlySpan = System.IO.File.ReadAllBytes(_configstore.FilePath);
                var reader = new Utf8JsonReader(jsonReadOnlySpan);
                var config = JsonSerializer.Deserialize<Config>(ref reader);
                if (config.MailBoxApps != null)
                {
                    var mailboxapp = config.MailBoxApps.FirstOrDefault<MailBoxApp>(a => a.Address == value.Address);
                    if (mailboxapp != null)
                    {
                        _logger.LogWarning("Duplicate address in config store.");
                        return StatusCode(500);
                    }
                }
                else
                {
                    config.MailBoxApps = new List<MailBoxApp>();
                }

                config.MailBoxApps.Add(value);

                using FileStream fs = System.IO.File.Create(_configstore.FilePath);
                var writerOptions = new JsonWriterOptions
                {
                    Indented = true
                };

                var documentOptions = new JsonDocumentOptions
                {
                    CommentHandling = JsonCommentHandling.Skip
                };
                using var writer = new Utf8JsonWriter(fs, options: writerOptions);
                using JsonDocument document = JsonDocument.Parse(JsonSerializer.Serialize<Config>(config), documentOptions);

                JsonElement root = document.RootElement;

                if (root.ValueKind == JsonValueKind.Object)
                {
                    writer.WriteStartObject();
                }

                foreach (JsonProperty property in root.EnumerateObject())
                {
                    property.WriteTo(writer);
                }

                writer.WriteEndObject();
                writer.Flush();

                return new JsonResult(config);
            }
            catch (Exception ex)
            {
                _logger.LogError("Unexpected error!\r\nMessage: {0}\r\nStackTrace: {1}", ex.Message, ex.StackTrace);
                return StatusCode(500);
            }
        }

        // PUT api/<ConfigController>
        [EnableCors(PolicyName = "AllowedOrigins")]
        [HttpPut]
        public IActionResult Put([FromBody] MailBoxApp value)
        {
            try
            {
                ReadOnlySpan<byte> jsonReadOnlySpan = System.IO.File.ReadAllBytes(_configstore.FilePath);
                var reader = new Utf8JsonReader(jsonReadOnlySpan);
                var config = JsonSerializer.Deserialize<Config>(ref reader);
                var mailboxapp = config.MailBoxApps.FirstOrDefault<MailBoxApp>(a => a.Address == value.Address);
                if (mailboxapp != null)
                {
                    mailboxapp.Name = value.Name;
                    mailboxapp.Address = value.Address;
                    mailboxapp.AppAddress = value.AppAddress;
                    mailboxapp.SpDocLibId = value.SpDocLibId;
                    mailboxapp.SpListId = value.SpListId;
                    mailboxapp.SpWebBaseUrl = value.SpWebBaseUrl;
                    mailboxapp.Users = value.Users;
                }
                else
                {
                    _logger.LogWarning("Address not found in config store.");
                    return StatusCode(500);
                }

                var writerOptions = new JsonWriterOptions
                {
                    Indented = true
                };

                var documentOptions = new JsonDocumentOptions
                {
                    CommentHandling = JsonCommentHandling.Skip
                };

                using FileStream fs = System.IO.File.Create(_configstore.FilePath);
                using var writer = new Utf8JsonWriter(fs, options: writerOptions);
                using JsonDocument document = JsonDocument.Parse(JsonSerializer.Serialize<Config>(config), documentOptions);

                JsonElement root = document.RootElement;

                if (root.ValueKind == JsonValueKind.Object)
                {
                    writer.WriteStartObject();
                }

                foreach (JsonProperty property in root.EnumerateObject())
                {
                    property.WriteTo(writer);
                }

                writer.WriteEndObject();
                writer.Flush();

                return new JsonResult(config);
            }
            catch (Exception ex)
            {
                _logger.LogError("Unexpected error!\r\nMessage: {0}\r\nStackTrace: {1}", ex.Message, ex.StackTrace);
                return StatusCode(500);
            }
        }

        // DELETE api/<ConfigController>/address
        [EnableCors(PolicyName = "AllowedOrigins")]
        [HttpDelete("{address}")]
        public IActionResult Delete(string address)
        {
            try
            {
                ReadOnlySpan<byte> jsonReadOnlySpan = System.IO.File.ReadAllBytes(_configstore.FilePath);
                var reader = new Utf8JsonReader(jsonReadOnlySpan);
                var config = JsonSerializer.Deserialize<Config>(ref reader);
                var mailboxapp = config.MailBoxApps.FirstOrDefault<MailBoxApp>(a => a.Address == address);
                if (mailboxapp != null)
                {
                    config.MailBoxApps.Remove(mailboxapp);
                }
                else
                {
                    _logger.LogWarning("Address not found in config store.");
                    return StatusCode(500);
                }

                var writerOptions = new JsonWriterOptions
                {
                    Indented = true
                };

                var documentOptions = new JsonDocumentOptions
                {
                    CommentHandling = JsonCommentHandling.Skip,
                };

                using FileStream fs = System.IO.File.Create(_configstore.FilePath);
                using var writer = new Utf8JsonWriter(fs, options: writerOptions);
                using JsonDocument document = JsonDocument.Parse(JsonSerializer.Serialize<Config>(config), documentOptions);

                JsonElement root = document.RootElement;

                if (root.ValueKind == JsonValueKind.Object)
                {
                    writer.WriteStartObject();
                }

                foreach (JsonProperty property in root.EnumerateObject())
                {
                    property.WriteTo(writer);
                }

                writer.WriteEndObject();
                writer.Flush();

                return new JsonResult(config);
            }
            catch (Exception ex)
            {
                _logger.LogError("Unexpected error!\r\nMessage: {0}\r\nStackTrace: {1}", ex.Message, ex.StackTrace);
                return StatusCode(500);
            }
        }
    }
}
