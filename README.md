# Scheduler and Refresher for Excel files

Reports Controller helps to schedule udpdating of Excel reports / data models / workbooks. In general, any Excel file.

It is an open-source solution that is focused on optimization of Excel-based reports

    - No administrator rights needed
    - No purchasing of additional software needed

What you need to successfully use it: Excel 2016 or later vresion, and possibility to run VBA macros on the workstation where you plan to have it launched.

# How to use this solution

1. Download latest release [Power Refresh.zip](https://github.com/IvanBond/Power-Refresh/releases)
2. Unpack it on C:\ drive (so you have C:\Power Refresh\ folder)
3. Open 'Reports Controller' and test how it works on test files (read description in each line to understand the scenario behing it).
4. Configure your own schedule and parameters for your workbooks
5. Enjoy your coffee while Reports Controller does the work for you :-)

# Additional scenarios

If you don't like idea of using Reports Controller, you are still able to use functionality of Refresher.xlsb, which handles refreshing process taking into consideraion all provided parameters.

Just call it from VBSciprt or .BAT file, see sample [Starter.vbs](https://github.com/IvanBond/Power-Refresh/blob/master/Starter.vbs).

How to schedule .vbs or .bat you may see on video [here](https://www.youtube.com/watch?v=oC_i1Cf9O2w).

# History of the idea

Typically, reporting specialists are interested in automation of standard reports refreshing.

When development of Excel report is finished, file may contain

    - Power Query (Get & Transform) queries, which are pulling data from multiple sources
    - Data Model (aka PowerPivot) to digest data and calculate various measures with DAX 
    - connections to enterprise sources, such as SAP BI (e.g. BW4HANA), SSAS, Azure Data Lake etc.
    - ordinary Excel formulas
    - Pivot Tables, Pivot Charts, usual Charts, shapes and so on
    - etc.
  
Developer needs a way to refresh content of the workbook with zero or minimum manual effort.

Imagine a situation when reports developer has 50 Excel models or even more than that. Would be great if such 'farm of reports' could be refreshed over night automatically, once per day, per month, or every hour - in other words - each file at scheduled time. And, in addition, reporting specialist would have a simple solution to control configuration for all those reports - kind of Control Panel (Mission Control Centre).
    
Basic idea of the refresh process is very simple.
'Refresher' must be able

    - Create new instance of Excel application (since Excel is not the most robust application, 
    so best practice is to use new Excel process each time)
    - Open target workbook provided as parameter for that specific Excel process
    - Run ThisWorkbook.RefreshAll (all queries and connections must be configured in a proper way, obviously)
    - Save workbook
    - Quit / Kill Excel process

that's all.

But this is only a basic scenario. Some reports require to run a macro before RefreshAll, or instead of RefreshAll they might need to refresh several Power Query queries in the pre-defined order, or run a macro after RefreshAll, or something else. In Self-Service BI area we can find endless number of scenarios.

Provided solution is flexible enough to manage many scenarios out-of-the-box. You just need to tweak parameters, not coding required.
However, having open-source refresher, analysts can adjust it for their own needs if they are confident with VBA programming.

# What additional requirements can we expect?

    - opportunity to refresh several workbooks simultanenously on the same computer (parallel)
    - different ways of saving the result - xlsx/xlsm/xlsb/csv/pdf etc.
    - saving resulting file(s) to local/network drive, or upload to SharePoint
    - opportunity to send resulting file via email (Outlook, CDO, Gmail etc.)
    - run report only on working days (using Business / Factory / Country specific calendar of working days)
    - skip refresh on the days when one of report's data sources is not available due to maintenance (if you know in advance, fill in a special table)
    - etc.
    
For example, if your enterprise data source is SAP BI - BO or BW, you can integrate Power Refresh solution with another one - [SAP BOA Automation](https://github.com/IvanBond/SAP-BOA-Automation)

# Why not Windows Task Scheduler?

To name a few issues I faced with it: it requires admin rights, hard to manage many reports, no control over used resourses or number of running Excel sessions, no simple log of execution process, hard to transfer tasks from one workstation to another (comparing to Copy/Paste-ing Power Refresh Excel file).
