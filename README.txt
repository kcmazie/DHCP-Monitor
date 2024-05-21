<#==============================================================================
         File Name : DHCP-Monitor.ps1
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr.com)
                   : 
       Description : A tool for monitoring DHCP scope statistics.
                   : 
             Notes : Normal operation is with no command line options.  
                   : Commandline options intentionally left out to avoid accidents.
                   :
      Requirements : Requires the PowerShell DHCPServer extensions.  Must be run ON a DHCP server.
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
   Version History : v1.0 - 08-16-22 - Original.  Forked from DHCP manager script. 
    Change History : v2.0 - 09-00-23 - Numerous operational & bug fixes
                   : v2.1 - 12-15-23 - Adjusted email options, report format, other minor bugs.
                   : v3.0 - 12-25-23 - Relocated private settings out to external config for publishing. 
                   : v3.1 - 01-25-24 - Altered email send so it always goes out if over 80 or 95 %
                   : v3.2 - 05-21-24 - Added color gradiations for % used column.  Added generalized 
                   : grey background.
                   :                  
==============================================================================#>
