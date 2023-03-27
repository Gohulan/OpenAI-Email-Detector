=================
Project Name: ChatGPT-Detector
=================
ChatGPT-Detector is an Outlook plugin that detects emails created using OpenAI's GPT language model API. This plugin uses Microsoft's Office Interop Outlook library and .NET Framework to scan incoming emails and determine whether they were composed using the ChatGPT language model.

=================
Installation
=================
To install the ChatGPT-Detector plugin, follow these steps:

Clone this repository to your local machine.
Open the solution file in Visual Studio.
Build the solution to generate the .dll file.
Open Outlook and navigate to File > Options > Add-ins.
Click on "Manage COM Add-ins" and then "Add".
Browse to the location of the .dll file and select it.
Click "OK" to add the plugin to Outlook.
The ChatGPT-Detector plugin is now installed and active in your Outlook account.

=================
Usage
=================
Once installed, the ChatGPT-Detector plugin automatically scans incoming emails and identifies those that were composed using the ChatGPT language model. When an email is detected, the plugin adds a label to the email's subject line and displays a notification to the user.

=================
Dependencies
=================
This plugin requires the following libraries and frameworks:

Microsoft.Office.Interop.Outlook
Microsoft.Office.Core
System.Net.Http
System.Net.Http.Headers
Newtonsoft.Json.Linq
System.Windows.Forms
System.Threading.Tasks

=================
Contributing
=================
If you would like to contribute to this project, please open a pull request with your proposed changes. We welcome all contributions!

=================
License
=================
This project is licensed under the MIT License - see the LICENSE.md file for details.
