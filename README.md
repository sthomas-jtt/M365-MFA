Script Information
------------------
This script was developed with the assistance of Microsoft Copilot, an AI companion that provided 
technical insights, optimization suggestions, and structural enhancements to improve functionality 
and efficiency.

It was designed to be user friendly, assist with installing the required modules, and offer methods 
for filtering the results to retrieve specific information. It will offer options to export the 
results to a CSV file (saved to the directory where the script was executed from), and an option
to see a summary with statistics based on your filter criteria. 

For MFA, we're concerned about having MFA configured, specifically with Strong authentication methods. 
We consider Email Authentication and Phone Authentication (SMS) to be legacy methods (Weak MFA). We
especially need to verify MFA for Global Administrators, as these accounts have Global access to all
aspects of Azure/M365. This is the key to the Kingdom!

Now that legacy (per-user MFA) authentication is being retired and MFA management moving to Security 
Defaults or Conditional Access, we should be setting all per-user MFA settings to disabled!
https://learn.microsoft.com/en-us/microsoft-365/admin/security-and-compliance/multi-factor-authentication-microsoft-365 

------------------
First, do you need to check for missing Powershell modules?

Welcome to the Stew's fancy MFA script!

The results will always be displayed in a separate Grid View Powershell window. (Like a spreadsheet)

You can choose to filter the results. Filter mode options are:
  1) Including all users in the report
  2) Filter where any condition can be true. Flexible = More Results
  3) Filter where ALL conditions must be true. Strict = Fewer Results
  4) Choose Default Filter = Include any Global Admins or any Licensed accounts

If you choose to filter, you will be asked which filter mode, and if you want to include:
  - Global Admins
  - Licensed users
  - Users without MFA methods
  - Users NOT set to disabled in Per-user MFA
  - Users with Sign-in allowed

You can choose to export the results to a CSV file that will be saved to the same location as the script.

You can choose to display a summary of the results in a separate window.

