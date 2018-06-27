# Scheduler and Refresher for Excel files

Reports Controller helps to schedule udpdate of Excel reports / data models / just files. In general, any Excel file.

It is an open-source solution that is focused on optimization of Excel-based reports

    - No administator rights
    - No purchasing additional software

only Excel is needed (Excel 2016 and later), and possibility to run VBA macros on workstation.

# How to use this solution

1. Download latest release [Power Refresh.zip](https://github.com/IvanBond/Power-Refresh/releases)
2. Unpack it on C:\ drive
3. Open 'Reports Controller' and test how it works on test files
4. Configure your own schedule and parameters for your workbooks
5. Enjoy!

# Additional scenarios

If you don't like idea to use Reports Controller, you still able to use functionality of Refresher.xlsb, which handles refresh process taking into consideraion all provided parameters.

Just call it from VBSciprt or .BAT file, see sample [Starter.vbs](https://github.com/IvanBond/Power-Refresh/blob/master/Starter.vbs).

How to schedule .vbs or .bat you may see on video [here](https://www.youtube.com/watch?v=oC_i1Cf9O2w).

# History of idea

Typically, reporting specialists are interested in automation of standard reports refresh.

When development of Excel report is finished, file may contain

    - Power Query (Get & Transform) queries, which pulling data from multiple sources
    - Data Model (aka PowerPivot) to digest data and calculate measure with DAX 
    - connections to enterprise sources, such as SAP BI, SSAS, Azure Data Lake etc.
    - ordinary Excel formulas
    - Pivot Tables, Pivot Charts, usual Charts, shapes etc.
    - etc.
  
Developer needs a way to refresh content of workbook with zero or minimum of manual effort.

Imagine situation when reports developer has 50 Excel models or more. Would be great if such 'farm of reports' could be refreshed over night, once per day, per month, or every hour - in other words - each file at planned time. And, in addition, reporting specialist would have a simple solution to control configuration for all those reports - kind of Control Panel (Mission Control Centre).
    
Basic idea of refresh is very simple. 
'Refresher' must be able

    - Create new instance of Excel application. Excel is not the most robust application, 
    so best practice is to use new Excel process each time.
    - Open workbook provided in parameter
    - Run ThisWorkbook.RefreshAll (queries must be configured in a proper way)
    - Save workbook
    - Quit / Kill Excel process

that's all.

But this is only basic scenario. Some reports require to run macro before RefreshAll, or instead of RefreshAll they want to refresh several PQ queries in defined order, or run macro after RefreshAll, or something else. In Self-Service BI area we can find endless number of scenarios.

Provided solution is flexible enough to manage many scenarios out-of-the-box. You just need to tweak parameters, not coding required.
However, having open-source refresher, analysts can adjust it for their needs as they usually know VBA.

# What additional requirements can we expect?

    - opportunity to refresh several models in the same time on one computer
    - different options to save result - xlsx/xlsm/xlsb/csv/pdf etc.
    - save resulting file(s) to local/network drive, or upload to SharePoint
    - opportunity to send resulting file via email (Outlook, CDO, Gmail etc.)
    - run report only on working days
    - skip refresh no day when one of report's data sources is not available due to maintenance
    - etc.
    
For example, if your enterprise data source is SAP BI - BO or BW, you can integrate this solution with another one - [SAP BOA Automation](https://github.com/IvanBond/SAP-BOA-Automation)

Nothing should stop Self-Service BI developers from achieving business goals!

Purchase and installation of software can be a problem for many specialists in large organizations due to strict IT policies.

Therefore, basic script from this project can be adapted to particular needs easily by those who are familiar with VBA/VBScipt.

# Why not Windows Task Scheduler

It requires admin rights, hard to manage many reports, no control over used resourses or number of running Excel sessions, no simple log, hard to transfer tasks from one workstation to another (comparing to Copy/Paste of Excel file) and so on.
