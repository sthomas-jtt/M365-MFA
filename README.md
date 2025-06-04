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

