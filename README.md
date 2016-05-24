# Power-Refresh
Refresher of Excel files with Power Query and PowerPivot model as Visual Basic Script (VBS).

This is an open source refresher that is aimed to optimize self-service Excel solutions on Windows workstations
  without administator rights
  without buying additional software - only Excel is needed

Typically, reporting specialists are interested in automated way of reports preparation. 
When reports are done as Excel files with 
  Power Query queries pulling data from external source + 
  Data Model (aka PowerPivot) to digest data + 
  Pivot Tables, Pivot Charts etc. based on Data Model data
developer needs a way to refresh all this content in a workbook.

Imagine situation when reports developer has 50 Excel models or more. Would be great if they can be refreshed during night, once per day, every hour - in other words - follow schedule. Several models can be refreshed in same time in separate Excel applications on same computer.

Basic idea of refresh is pretty simple. 
Refresher must
  Create new instance of Excel application
  Open workbook provided in parameter
  Run ThisWorkbook.RefreshAll
  Save workbook
  Kill Excel process
that's all.

But this is only basics. However, basics are not always match specific needs of someone. Some want to run macro before RefreshAll, or instead of RefreshAll they want to refresh chain of PQ queries in defined order, or run macro after RefreshAll, or something else. In Self-Service BI area we can find endless number of scenarios.

Workstation with Windows is considered because then it is possible to use Task Scheduler without buying any additional software. Purchase of software can be a problem for many employees.

Nothing should stop Self-Service BI developers :-). That's why this project done in Visual Basic Script.
VBS is very similar to VBA, which is usually 'must-have' skill for reporting specialist. There are a lot of samples on the Internet how to convert VBA to VBS, how to run VBA from VBS and vice versa.
Therefore, basic script from this project can be adapted to particular needs easily by those who are familiar with VBA/VBS.

How to use
Create folder 'Power Refresh' on C:\ drive

Download all project files to this folder

Read Instruction in 'Refresher.vbs' to understand logic, Report vs Data Transfer, and Scopes concept

Launch Refresher.vbs from command line or via scheduled task in Task Scheduler
