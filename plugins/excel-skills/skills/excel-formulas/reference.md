# Formula patterns

Assume data on a sheet `Data` with headers in row 1: A Date, B Customer, C Invoice No,
D Amount, E Status. Adjust letters to the user's layout.

## Lookups

| Need | Excel 365 / 2021 / Sheets | Excel 2016 / 2019 |
|---|---|---|
| Find a value by key | `=XLOOKUP(H2, Data!C:C, Data!D:D, "Not found")` | `=IFNA(INDEX(Data!D:D, MATCH(H2, Data!C:C, 0)), "Not found")` |
| Lookup to the left | same XLOOKUP | INDEX/MATCH (VLOOKUP can't look left) |
| Two keys | `=XLOOKUP(1, (Data!B:B=H2)*(Data!A:A=I2), Data!D:D)` | `=INDEX(Data!D:D, MATCH(1, INDEX((Data!B:B=H2)*(Data!A:A=I2),0), 0))` |
| Last match | `=XLOOKUP(H2, Data!B:B, Data!D:D, , 0, -1)` | `=LOOKUP(2, 1/(Data!B:B=H2), Data!D:D)` |
| Slab / rate table (tax, commission) | `=XLOOKUP(H2, Slabs!A:A, Slabs!C:C, , -1)` | `=VLOOKUP(H2, Slabs!A:D, 3, TRUE)` (table sorted ascending) |

VLOOKUP pitfalls: the 4th argument must be `FALSE` for exact match; inserting a column breaks
the column number; keys stored as text vs numbers never match (`=VALUE()` or clean the data).

## Totals with conditions

- Sum for a customer: `=SUMIFS(Data!D:D, Data!B:B, H2)`
- Sum for a month (month start date in H1): `=SUMIFS(Data!D:D, Data!A:A, ">="&H1, Data!A:A, "<"&EDATE(H1,1))`
- Sum between two dates: `=SUMIFS(Data!D:D, Data!A:A, ">="&H1, Data!A:A, "<="&H2)`
- Count unpaid: `=COUNTIFS(Data!E:E, "Due")`
- Sum where text contains a word: `=SUMIFS(Data!D:D, Data!B:B, "*traders*")`
- Exact match on text containing `*` or `?`: escape with `~`: `"Usman~*Store"`
- Sum with OR (Lahore or Karachi): `=SUM(SUMIFS(Data!D:D, Data!F:F, {"Lahore","Karachi"}))`
- Unique list (365): `=SORT(UNIQUE(FILTER(Data!B2:B10000, Data!B2:B10000<>"")))`

## Running balance and ledgers

- Running balance (Debit in D, Credit in E, row 2 first): `=SUM($D$2:D2)-SUM($E$2:E2)`
- Opening balance in F1: `=F1+D2-E2` filled down
- Debit/credit from one signed column: Dr `=MAX(D2,0)`, Cr `=MAX(-D2,0)`

## Receivable ageing (invoice date in A, report date in $K$1)

- Days outstanding: `=$K$1-A2`
- Bucket: `=IFS(K2<=30,"0-30", K2<=60,"31-60", K2<=90,"61-90", TRUE,"90+")`
  (2016/2019: `=IF(K2<=30,"0-30",IF(K2<=60,"31-60",IF(K2<=90,"61-90","90+")))`)

## Text and dates

- Trim extra spaces, including non-breaking ones: `=TRIM(SUBSTITUTE(B2, CHAR(160), " "))`
- Text number to number: `=VALUE(SUBSTITUTE(SUBSTITUTE(D2,",",""),"Rs",""))`
- DD/MM text date to date, in any regional setting:
  `=DATE(RIGHT(A2,4), MID(A2,4,2), LEFT(A2,2))` (for `05/11/2026`)
- Month name: `=TEXT(A2, "mmm-yyyy")`; fiscal year (July-June):
  `=IF(MONTH(A2)>=7, YEAR(A2)&"-"&RIGHT(YEAR(A2)+1,2), YEAR(A2)-1&"-"&RIGHT(YEAR(A2),2))`
- First name: `=TEXTBEFORE(B2, " ")` (365) or `=LEFT(B2, FIND(" ", B2&" ")-1)`
- PK mobile digits only, as 03XX-XXXXXXX:
  `=TEXT(VALUE(RIGHT(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(G2,"-","")," ",""),"+",""),10)),"0000-0000000")`
  (for numbers like +92 300 1234567 / 0300-1234567; returns 0300-1234567)

## Pakistani number formats

Format Cells > Number > Custom:

- Lakh/crore grouping (1,50,000 and 1,25,00,000):
  `[>=10000000]##\,##\,##\,##0;[>=100000]##\,##\,##0;##,##0`
- With "Rs": `[>=10000000]"Rs "##\,##\,##\,##0;[>=100000]"Rs "##\,##\,##0;"Rs "##,##0`
- Show in lakhs: number formats can only divide by thousands, so use a helper column
  `=D2/100000` with the format `0.00 "lakh"`.

Custom formats only handle two conditions, so negative amounts use normal grouping. For
negatives in lakh style, put the sign in a separate column or use a formula with TEXT.

## Amount in words

Excel has no built-in function. Options: a VBA function (needs a macro-enabled .xlsm), a
LAMBDA in Excel 365, or generate the words outside Excel. For Pakistani invoices the wording
uses lakh and crore: 1,50,000 is "One Lakh Fifty Thousand Rupees Only".

## Error guide

| Error | Usual cause | Fix |
|---|---|---|
| `#N/A` | Lookup value not found: extra space, text vs number, wrong range | TRIM/VALUE the key; check the lookup range; IFNA for a friendly message |
| `#VALUE!` | Text in a calculation, or a text date | Clean the values; DATE() instead of text dates |
| `#REF!` | A referenced column/row/sheet was deleted | Rebuild the reference; use Tables or INDEX/MATCH to avoid it |
| `#SPILL!` | Dynamic array has no room to spill | Clear the cells below/right of the formula |
| `#DIV/0!` | Dividing by zero or blank | `=IF(B2=0, 0, A2/B2)` or IFERROR |
| `#NAME?` | Misspelled function, or a 365 function in an older Excel | Check spelling; use the classic version |
| Total is 0 or too low | Numbers stored as text | Clean the column (Data > Text to Columns > Finish, or the clean-data skill) |
| Formula shows as text | Cell formatted as Text, or a leading space/apostrophe | Format as General, re-enter |
| Values don't update | Calculation set to Manual | Formulas > Calculation Options > Automatic |
