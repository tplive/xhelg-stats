# X-Helg geocaching Advent calendar statistics
Statistics for a geocaching Advent calendar. Made for the geocaching community in Helgeland, Norway for the X-Helg Advent calendar, but probably useful elsewhere too. First iteration written in PowerShell.

Website here: https://tplive.github.io/xhelg-stats/
# Versions
PowerShell 5.0
GSAK 8.6.0
SQLite for .NET 1.0.103.0 (Tested version for .NET 4.6 64-bit: sqlite-netFx46-setup-bundle-x64-2015-1.0.103.0.exe)

# Usage
The script is written in PowerShell, and connects to a SQLite database using a SQLite for .Net library. Database is generated using GSAK.

First create an empty GSAK database, and import all caches from the Advent calendar. Using a geocaching bookmark-list can help with this. Once in GSAK, download all logs for the caches. This makes the database ready for use.

The top section of the script does a dataprep of the database:

1. Create (with drop if existes) a points table, create a row for each nickname.
2. Create (with drop if exists) a first-to-find (ftf) table, selects and adds all FTF logs, based on logs containing the "well-known" FTF code "FTF in asterisks and curly braces": {\*FTF\*}.
3. Updates each cache in the calendar: sets the correct placeddate (Dec. 1-24) and uses the User4 field to store a short-code for each cache; [#n] where n is the day of Advent 1-24.

Then we get all logs and give points based on these rules:
- Logged find in December equals 1 point.
- Logged find on the same date as the cache was published is 2 points.
- First to find is 3 points.
- "Co-FTF", ie more than one cacher logs a FTF, is 2 points each.
- Publishing a listing of your own is 3 points.

The points-table is generated with short-codes and poinsts and stored in a CSV file. You can then use Excel to create a "pretty" report, see example.

TODO:
- Want HTML report to publish online.
- Want dynamic, sortable report, linked to each cache and geocacher.
- Want to see if Project GC has anything, or can make anything for this purpose.
