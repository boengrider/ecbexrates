# ecbexrates
 https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml
 
 Script downloads and parses the ECB published XML file containing exchange rates valid for next working day.
 ECB publishes the new rates @16:00 CET. 
 
 *TCD - Target Closing Day is day on which ECB does not publish new exchange rates that would normally be published this day and
 would be valid for next working day
 
 Script is configured to run between 17:00 - 23:59. In this time window it downloads XML and parses it searching
 for desired currencies.  
 
 Following schema applies the day on which script runs IS NOT a TCD (Target Closing Day*):
 
 Rate published day         Rate valid for day
 Monday                     Tuesday
 Tuesday                    Wednesday
 Wednesday                  Thursday
 Thursday                   Friday,Saturday,Sunday
 Friday                     Monday
 
 TCD example
 Mo     Tu      We      Th       Fr      Sa      Su      Mo
                     Apr 30th  May 1st                 May 4th
                                 
                                
 On May 1st, which is a TCD, ECB will not publish new echange rates. On Friday the script will determine that
 today is a TCD and will use a historical rates provided by ECB. It will use exchange rates from the closest non TCD weekday
 which in this case is April 30 - Thursday
                        
 
