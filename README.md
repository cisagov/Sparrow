# Sparrow.ps1 #

Sparrow.ps1 was created by CISA's Cloud Forensics team to help detect possible compromised accounts and applications in the Azure/m365 environment. The tool is intended for use by incident responders, and focuses on the narrow scope of user and application activity endemic to identity and authentication based attacks seen recently in multiple sectors. It is neither comprehensive nor exhaustive of available data, and is intended to narrow a larger set of available investigation modules and telemetry to those specific to recent attacks on federated identity sources and applications.
 
Sparrow.ps1 will check and install the required PowerShell modules on the analysis machine, check the unified audit log in Azure/M365 for certain indicators of compromise (IoC's), list Azure AD domains, and check Azure service principals and their Microsoft Graph API permissions to identify potential malicious activity. The tool then outputs the data into multiple CSV files in a default directory.

## Requirements ##

The following AzureAD/m365 permissions are required to run Sparrow.ps1, and provide it read-only access to the Tenant.

   - Azure Active Directory:
     - Security Reader
   - Security and Compliance Center:
     - Compliance Adminstrator
   - Exchange Online Admin Center: Utilize a custom group for these specific permissions:
     - Mail Recipients
     - Security Group Creation and Membership
     - User options
     - View-Only Audit log
     - View-Only Configuration
     - View-Only Recipients

To check for the MailItemsAccessed Operation, your tenant organization requires an Office 365 or Microsoft 365 E5/G5 license.

## Installation ##

Sparrow.ps1 does not require any extra steps for installation once the permissions detailed in Requirements are satisfied.

The function, Check-PSModules, will check to see if the three required PowerShell modules are installed on the system and if not, it will use the default PowerShell repository on the system to reach out and install. If the modules are present but not imported, the script will also import the missing modules so that they are ready for use.

The required PowerShell modules:

  - CloudConnect (https://www.powershellgallery.com/packages/CloudConnect/1.1.2)
  - AzureAD (https://www.powershellgallery.com/packages/AzureAD/2.0.2.128)
  - MSOnline (https://www.powershellgallery.com/packages/MSOnline/1.1.183.57)

## Usage ##

To use Sparrow.ps1, type the following command into a PowerShell window (assuming file is in your working directory):

`.\Sparrow.ps1`

## Issues ##

If you have issues using the code, open an issue on the repository!

You can do this by clicking "Issues" at the top and clicking "New Issue" on the following page.

## Contributing ##

We welcome contributions!  Please see [here](CONTRIBUTING.md) for details.

## License ##

This project is in the worldwide [public domain](LICENSE).

This project is in the public domain within the United States, and copyright and related rights in the work worldwide are waived through the [CC0 1.0 Universal public domain dedication](https://creativecommons.org/publicdomain/zero/1.0/).

All contributions to this project will be released under the CC0 dedication. By submitting a pull request, you are agreeing to comply with this waiver of copyright interest.

## Legal Disclaimer ##

NOTICE

This software package (“software” or “code”) was created by the United States Government and is not subject to copyright. You may use, modify, or redistribute the code in any manner. However, you may not subsequently copyright the code as it is distributed. The United States Government makes no claim of copyright on the changes you effect, nor will it will it restrict your distribution of bona fide changes to the software. If you decide to update or redistribute the code, please include this notice with the code. Where relevant, we ask that you credit the Cybersecurity and Infrastructure Security Agency with the following statement: “Original code developed by the Cybersecurity and Infrastructure Security Agency (CISA), U.S. Department of Homeland Security.”

USE THIS SOFTWARE AT YOUR OWN RISK. THIS SOFTWARE COMES WITH NO WARRANTY, EITHER EXPRESS OR IMPLIED. THE UNITED STATES GOVERNMENT ASSUMES NO LIABILITY FOR THE USE OR MISUSE OF THIS SOFTWARE OR ITS DERIVATIVES.

THIS SOFTWARE IS OFFERED “AS-IS.” THE UNITED STATES GOVERNMENT WILL NOT INSTALL, REMOVE, OPERATE OR SUPPORT THIS SOFTWARE AT YOUR REQUEST. IF YOU ARE UNSURE OF HOW THIS SOFTWARE WILL INTERACT WITH YOUR SYSTEM, DO NOT USE IT.