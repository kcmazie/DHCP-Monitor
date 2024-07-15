<#==============================================================================
         File Name : DHCP-Manager.ps1
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr.com)
                   : 
       Description : A multi-purpose tool for managing and monitoring DHCP.
                   : 
             Notes : Normal operation is with no command line options.  
                   : Commandline options intentionally left out to avoid accidents.
                   :
      Requirements : Requires the PowerShell DHCPServer extensions.
                   : 
                   : 
          Warnings : 
                   :   
             Legal : Public Domain. Modify and redistribute freely. No rights reserved.
                   : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF 
                   : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED.
                   :
           Credits : Code snippets and/or ideas came from many sources including but 
                   : not limited to the following:
                   : https://github.com/n2501r/spiderzebra/blob/master/PowerShell/DHCP_Scope_Report.ps1
                   : 
    Last Update by : Kenneth C. Mazie                                           
   Version History : v1.00 - 08-16-22 - Original 
    Change History : v2.00 - 09-00-23 - Numerous operational & bug fixes
                   : v2.10 - 12-15-23 - Adjusted email options, report format, other minor bugs.
                   : v3.00 - 12-25-23 - Relocated private settings out to external config for publishing. 
                   : v3.10 - 01-25-24 - Altered email send so it always goes out if over 80 or 95 %
                   : v3.20 - 05-21-24 - Added color gradiations for % used column.  Added generalized 
                   :                    grey background.
                   : v4.00 - 07-12-24 - Fixed detection of failed purges.  Added new messaging on report
                   :                    for failed purges as well as attching log files to the email.  
                   :                    Altered detection of console verses IDE.  Shuffled order of 
                   :                    operations to better report stats.  Tweaked HTML report coloring.
                   : v4.10 - 07-15-24 - Added gradiated % colors back.  Accidentaly removed after v3.0
                   :                  
==============================================================================#>
