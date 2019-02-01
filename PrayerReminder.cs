#r "Newtonsoft.Json"

using System.Net;
using Newtonsoft.Json;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Net.Mail;
using Newtonsoft.Json;

public static void Run(TimerInfo daily, TraceWriter log)
{
    log.Info($"PrayerReminder function executed at: {DateTime.Now}");

    // Get application configuration
    var app_id = ConfigurationManager.AppSettings["planningCenterAppId"]; 
    var secret = ConfigurationManager.AppSettings["planningCenterSecret"];
    var outgoingEmailAddress = ConfigurationManager.AppSettings["outgoingEmailAddress"];
    var outgoingEmailPassword = ConfigurationManager.AppSettings["outgoingEmailPassword"];

    // Set variable/instantiate objects
            var allReturned = false;
            var offset = 0; 
            List<PersonAttributes> people = new List<PersonAttributes>();

            //Iteratate through api until all members are returned
            while (allReturned == false)
            {
                // Set the offset and URL params
                var url = $"https://api.planningcenteronline.com/people/v2/people?order=last_name&where[membership]=Member&offset={offset}";
                //Use the planning center api to return persons records
                Task<string> peopleResults = PlanningCenterApi(url, app_id, secret);
                //Convert the result from json to a contract object
                var peopleObj = JsonConvert.DeserializeObject<PersonContract>(peopleResults.Result);

                //Get notes 
                var notesUrl = $"https://api.planningcenteronline.com/people/v2/notes?where[note_category_id]=4001";
                Task<string> notesResults = PlanningCenterApi(notesUrl, app_id, secret);
                var noteObj = JsonConvert.DeserializeObject<NoteContract>(notesResults.Result);

                // Loop through return object assigning members to list
                foreach (var person in peopleObj.data)
                {
                    if(person.attributes.status == "active")
                    {
                        people.Add(new PersonAttributes()
                        {
                            avatar = person.attributes.avatar,
                            first_name = person.attributes.first_name,
                            last_name = person.attributes.last_name,
                            notes = noteObj.data.Where(n => n.attributes.person_id == person.id && n.attributes.created_at >= DateTime.Now.AddMonths(-3)).ToList()
                        });
                    }
                }

                // Check to see if the current people list returned was a full set
                allReturned = peopleObj.data.Count == 25 ? false : true;

                // If the previour list was not a full set, increment the offset by 25
                if (allReturned == false)
                {
                    //Get the next 25 records
                    offset += 25;
                }

            }

            // Get the total number of days per month
            var daysInMonth = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month);
            // Calculate the number of people to put in each day
            double peoplePerDay = ((people.Count * 1.0) / (daysInMonth * 1.0));
            // Round up when people per day is not a whole number
            var peoplePerDayPrecise = (int)Math.Ceiling(peoplePerDay);

            var prayer = people.GetRange(((peoplePerDayPrecise - 1) * DateTime.Now.Day), peoplePerDayPrecise);

            var html = ReturnEmail(prayer);

            SmtpClient client = new SmtpClient() {
                Host = "smtp.gmail.com",
                Port = 587,
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(outgoingEmailAddress, outgoingEmailPassword)
            };

            var message = new MailMessage(outgoingEmailAddress, "", $"Elder Daily Prayer - {DateTime.Now.ToLongDateString()}", html);
            message.IsBodyHtml = true;
            client.Send(message);
}

private static async Task<string> PlanningCenterApi(string url, string app_id, string secret)
{
    var client = new HttpClient();

    var byteArray = new UTF8Encoding().GetBytes($"{app_id}:{secret}");
    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));

    client.BaseAddress = new Uri(url);
    var request = new HttpRequestMessage(HttpMethod.Get, url);

    var response = await client.SendAsync(request);
    var content = await response.Content.ReadAsStringAsync();

    return content;

}

private static string ReturnEmail(List<PersonAttributes> todaysPrayer)
{

    var html = "<html>";
    html += $"<h5>{DateTime.Now.ToLongDateString()}</h5>";

    foreach (var person in todaysPrayer)
    {
        var name = $"{person.first_name} {person.last_name}";
        var avatar = person.avatar;
        var notes = ""; 

        if(person.notes.Count > 0)
        {
            var noteRows = "";

            foreach (var note in person.notes.OrderByDescending(d => d.attributes.created_at))
            {
                noteRows += $"<tr><td style=\"vertical-align: text-top; padding:2px;\">{note.attributes.created_at.ToShortDateString()}: </td><td style=\"padding:2px;\">{note.attributes.note}</td></tr>";
            };

            notes += $"<td style=\"padding:2px\"> Prayer Notes:<table>{noteRows}</table></td>";
        }

        html += $"<table><tr><td style=\"vertical-align: text-top; padding:2px;\"><img height=\"100\" width=\"100\" src=\"{avatar}\"></td><td><table><tr><td style=\"vertical-align: text-top; padding:2px;\"><h5>{name}</h5></td></tr><tr>{notes}</tr></table></td></tr></table>";
    }

    html += "</html>";

    return html;
}

public class PersonData
{
    public string type { get; set; }
    public int id { get; set; }
    public PersonAttributes attributes { get; set; }
}

public class NoteData
{
    public string type { get; set; }
    public int id { get; set; }
    public NoteAttributes attributes { get; set; }
}

public class Links
{
    public string self { get; set; }
    public string next { get; set; }
}

public class PersonAttributes
{
    public string avatar { get; set; }
    public string first_name { get; set; }
    public string last_name { get; set; }
    public string status { get; set; }
    public string birthdate { get; set; }

    public List<NoteData> notes { get; set; }

}

public class NoteAttributes
{
    public string note { get; set; }
    public int person_id { get; set; }
    public DateTime created_at { get; set; }
    public int created_by_id { get; set; }

}

public class PersonContract
{
    public Links links { get; set; }
    public List<PersonData> data { get; set; }
}

public class NoteContract
{
    public Links links { get; set; }
    public List<NoteData> data { get; set; }
}