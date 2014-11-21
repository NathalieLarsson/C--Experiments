using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace HelloWorld
{
	class Program
	{
		static void Main(string[] args)
		{
			ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
			service.UseDefaultCredentials = true;
			service.AutodiscoverUrl("nathalie@mirror.se", RedirectionUrlValidationCallback);
			//service.Credentials = new WebCredentials("konferensrummet@mirror.se", "abc123");

			//service.AutodiscoverUrl("konferensrummet@mirror.se", RedirectionUrlValidationCallback);
			service.TraceEnabled = true;
			service.TraceFlags = TraceFlags.All;

			DateTime startDate = DateTime.Now;
			DateTime endDate = DateTime.Now.AddDays(10);
			const int maxItems = 10;

			CalendarFolder calendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());
			CalendarView cView = new CalendarView(startDate, endDate, maxItems);
			cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End);
			FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);

			Console.WriteLine("\nThe first " + maxItems + " appointments on your calendar from " + startDate.Date.ToShortDateString() +
							  " to " + endDate.Date.ToShortDateString() + " are: \n");

			foreach (Appointment a in appointments)
			{
				Console.Write("Subject: " + a.Subject.ToString() + " ");
				Console.Write("Start: " + a.Start.ToString() + " ");
				Console.Write("End: " + a.End.ToString());
				Console.WriteLine();
			}

			//EmailMessage email = new EmailMessage(service);
			//email.ToRecipients.Add("mattias.festin@mirror.se");

			//email.Subject = "HelloWorld";
			//email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API");

			////email.Send();
			Console.ReadLine();
		}

		private static bool RedirectionUrlValidationCallback(string redirectionUrl)
		{
			// The default for the validation callback is to reject the URL.
			bool result = false;

			Uri redirectionUri = new Uri(redirectionUrl);

			// Validate the contents of the redirection URL. In this simple validation
			// callback, the redirection URL is considered valid if it is using HTTPS
			// to encrypt the authentication credentials. 
			if (redirectionUri.Scheme == "https")
			{
				result = true;
			}
			return result;
		}
	}
}
