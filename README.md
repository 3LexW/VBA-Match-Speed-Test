# VBA-Match-Speed-Test
## Test Goal
Using Match and Index function, we would like to test which of the following has the fastest running speed
Since the test is under , the test might updated in other environment.

## Test Process
Using the following Excel file
| Num A   | Num B | Row No  | Found Num A at Num B's row |
|-----|-------|-----|--------|
| 1   | 60000 | R1  | R60000 |
| 2   | 59999 | R2  | R59999 |
| 3   | 59998 | R3  | R59998 |
| ... | ...   | ... | ...    |

The goal is to complete 60,000 rows of matching, using the following methods as parameter
+ Set Range
+ Set Variant
+ Use Dictionary Object

## Test Environment (a - Remote Control Server)
- Processor : i3-6100 @ 3.70GHz
- RAM: 4.00 GB
- Windows 10 Pro
