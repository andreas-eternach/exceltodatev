This programm reads excel-tables and stores the content in the
datev-format, so that the single accounting-records can be imported into datev.

Following accounting-information is supported:
  * Haben / Soll (the amount of money)
  * Konto
  * Gegenkonto
  * Kost1
  * Kost2
  * Datum
  * Text
  * Beleg1
  * Beleg2

The interface towards office is implemented via DDE. This way only the currently opened excel-files may be converted.