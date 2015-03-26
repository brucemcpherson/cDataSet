# VBA Project: cDataSet
This cross reference list for repo (cDataSet) was automatically created on 26/03/2015 10:03:40 by VBAGit.For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")
You can see [library and dependency information here](dependencies.md)

###Below is a cross reference showing which modules and procedures reference which others
*module*|*proc*|*referenced by module*|*proc*
---|---|---|---
cCell||cDataRow|create
cDataColumn||cDataSet|create
cDataRow||cDataSet|filterOk
cDataRow||cDataSet|create
cDataSets||cDataSet|populateData
cHeadingRow||cDataSet|Class_Initialize
cJobject||cDataSet|jObject
cJobject||cDataSet|populateJSON
cJobject||cDataSet|populateGoogleWire
cregXLib||regXLib|rxMakeRxLib
cStringChunker||cJobject|recurseSerialize
cStringChunker||cJobject|unSplitToString
cStringChunker||cJobject|serialize
regXLib|rxReplace|cDataSet|populateGoogleWire
usefulcJobject|toISODateTime|cDataSet|jObject
usefulSheetStuff|firstCell|cDataSet|rePopulate
usefulSheetStuff|getLikelyColumnRange|cDataSet|populateData
usefulSheetStuff|toEmptyRow|cDataSet|create
usefulSheetStuff|wholeSheet|cDataSet|load
UsefulStuff|makeKey|cDataSet|create
UsefulStuff|makeKey|cDataSet|populateData
UsefulStuff|makeKey|cDataSet|bigCommit
UsefulStuff|q|cDataSet|populateGoogleWire
