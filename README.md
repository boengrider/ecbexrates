# ecbexrates
 https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml
 
 Script downloads and parses the ECB published XML file containing exchange rates valid for next working day.
 ECB publishes the new rates @16:00 CET except TCD*
 
 *TCD - Target Closing Day is day on which ECB does not publish new exchange rates that would normally be published this day and
 would be valid for next working day
 

 This script produces output file suitable for uploading to SAP. See the sample output file.
 
 Output files are named like this: YYYYMMDD.txt e.g 20200521.txt
 This file would contain exchange rates valid for 21st May 2020
 
 Output files are saved in C:\ExRate directory
 Log is in C:\ExRate\log.txt
