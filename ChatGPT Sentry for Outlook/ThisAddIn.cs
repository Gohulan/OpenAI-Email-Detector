﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace ChatGPT_Sentry_for_Outlook
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.ItemLoad += Application_ItemLoad;
        }

        private async void Application_ItemLoad(object item)
        {
            if (item is MailItem mailItem)
            {
                if (await IsGeneratedByChatGPT(mailItem.Body))
                {
                    MessageBox.Show("This email was generated by ChatGPT.");
                }
            }
        }

        private async Task<bool> IsGeneratedByChatGPT(string emailBody)
        {
            using (var client = new HttpClient())
            {
                // Set the authorization header with your OpenAI API key
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", "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");

                // Prepare the HTTP request body with the email content
                var requestBody = new JObject
                {
                    { "model", "text-davinci-002" }, // Use the Davinci model for the best results
                    { "prompt", emailBody },
                    { "temperature", 0.5 } // Adjust the temperature for more or less creative responses
                };

                // Send the request to the OpenAI API
                var response = await client.PostAsync("https://api.openai.com/v1/completions", new StringContent(requestBody.ToString()));

                // Parse the response and check if it contains the ChatGPT prompt
                var responseBody = await response.Content.ReadAsStringAsync();
                var jsonObject = JObject.Parse(responseBody);
                var completions = jsonObject["choices"];
                foreach (var completion in completions)
                {
                    if (completion["text"].ToString().Contains("ChatGPT"))
                    {
                        return true;
                    }
                }

                return false;
            }
        }
    
private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
