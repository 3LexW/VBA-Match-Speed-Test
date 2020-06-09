# VBA-Match-Speed-Test
## Test Goal
Using Match and Index function, we would like to test which of the following has the fastest running speed.
The test might updated in other environment.

## Test Process
Using the following Excel file
| Num A   | Num B | Row No  | Found Num A at Num B's row |
|-----|-------|-----|--------|
| 1   | 60000 | R1  | R60000 |
| 2   | 59999 | R2  | R59999 |
| 3   | 59998 | R3  | R59998 |
| ... | ...   | ... | ...    |

The goal is to complete 10,000 rows of matching, using the following methods as parameter
+ Set Range
+ Set Variant
+ Use Dictionary Object

## Test Environment (a - Remote Control Server)
- Processor : i3-6100 @ 3.70GHz
- RAM: 4.00 GB
- Windows 10 Pro
### Result
| Method     | Test1 | Test2  | Test3  | Test4  | Test5  | Mean           |
|------------|-------|--------|--------|--------|--------|----------------|
| Range      | 20s   | 20.15s | 19.98s | 20.59s | 20.18s | 20.18s         |
| Variant    | #N/A  | #N/A   | #N/A   | #N/A   | #N/A   | Not measurable |
| Dictionary | 2.43s | 2.12s  | 2.25s  | 2.45s  | 2.08s  | 2.266s         |
