History
=======

- **RangeChecker v1.0** (19-Feb-2005): hello world! (linear model implemented)
- RangeChecker v1.1 (24-Feb-2005): added data controlling & comfort functions to increase usability (ratio_percent)
- RangeChecker v1.2 (26-Feb-2005): added "counting values in range" algorithm, added Tools to check data integrity before using main functions (see addon)...
- RangeChecker v1.3 (08-Dec-2005): new data integrity check v1.1
- RangeChecker v1.4 (07-Jan-2006): added time lowerthanrange, time higherthanrange, dev lowerthanrange, dev higherthanrange, selection mechanism and workaround for actual setting (all basing on linear model), integrated debug_checkrange
- RangeChecker v1.5 (17-Jan-2007): added maxTimeInterval cutoff possibility for datasets in which data entries are missing over a long time period
- **RangeChecker v2.0** (12-Dec-2010): moved from VBA for Excel to Visual Basic 2008 (.NET v3.5) added bridging management. Main work was to implement a proper data import from tab spaced spreadsheets.
- RangeChecker v2.02 (18-Jan-2011): moved to Visual Basic 2010 (.NET v4.0) test validation with data
- RangeChecker v2.1 (06-Feb-2011 Super Bowl Edition): handle multiple files (including mainLog and individualLogs); calc mean / median of delta t between measurements
- RangeChecker v2.1.1 (01-Mar-2011): create also a summary result file.
- RangeChecker v2.2 (26-Oct-2011): check for empty UPN; output table legend only once in summary result file; check occurence of UPN (multiples/missing)
- RangeChecker v2.2.1 (18-Jan-2012): another quality check before running on real data of the second PS-OAK study (to be published)