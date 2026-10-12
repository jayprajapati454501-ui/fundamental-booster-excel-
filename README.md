# Fundamental Booster — Excel Project

## Project Overview
The **Fundamental Booster** project is an Excel practice workbook designed to strengthen spreadsheet skills using student, sales, product, and employee datasets. It covers data entry, formatting, calculations, lookups, text manipulation, date functions, and dynamic data analysis.

## Objectives
- Practise relative and absolute cell references.
- Classify student grades using `IF`, nested `IF`, `AND`, and `OR`.
- Summarize data with `COUNTIFS`, `SUMIFS`, and `AVERAGEIFS`.
- Retrieve information using `VLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`, and `XMATCH`.
- Manipulate text with `LEFT`, `FIND`, `UPPER`, `LOWER`, and `TEXT`.
- Use `INDIRECT` and `OFFSET` for dynamic references.
- Perform date calculations and financial rounding.
- Extract matching records with `FILTER`.

## Workbook Sheets

| Worksheet | Purpose |
|---|---|
| `Read Me` | Project overview and instructions |
| `Student Data` | Student scores, totals, averages, grades, and performance analysis |
| `Sales Data` | Sales totals, discounts, regional/product summaries, and dynamic calculations |
| `Product Data` | Product catalogue and product-code lookup examples |
| `Employee Data` | Employee salaries, departments, joining dates, ages, and lookup examples |
| `Formula Tasks` | Checklist of Excel skills with example formulas |

## Key Excel Functions

### Cell References and Formatting
- Relative reference: `=A2*B2`
- Absolute reference: `=A2*$B$1`
- Apply suitable date, currency, percentage, and number formats.

### Conditional Formulas
- Grade classification: `=IF(G5>=90,"A",IF(G5>=80,"B",IF(G5>=70,"C",IF(G5>=60,"D","F"))))`
- Check whether both Math and Science exceed 80: `=IF(AND(C5>80,D5>80),"Yes","No")`
- Discount eligibility example: `=IF(OR(H5>=500,C5="North"),10%,0%)`

### Conditional Summary Functions
- Count averages above 60: `=COUNTIFS(G5:G16,">60")`
- Sum sales by region and product: `=SUMIFS(H5:H16,C5:C16,"North",E5:E16,"Notebook")`
- Average scores above 60: `=AVERAGEIFS(G5:G16,G5:G16,">60")`

### Lookup and Matching
- Product price: `=VLOOKUP(I5,A5:F12,4,FALSE)`
- Employee salary: `=XLOOKUP(M5,A5:A12,D5:D12,"Not found")`
- Employee department: `=INDEX(C5:C12,MATCH(M5,A5:A12,0))`
- Product position: `=XMATCH(I5,A5:A12,0)`

### Text, Dates, and Dynamic References
- First name: `=LEFT(B5,FIND(" ",B5&" ")-1)`
- Uppercase text: `=UPPER(B5)`
- Joining month: `=TEXT(E5,"mmmm")`
- Age in years: `=DATEDIF(F5,TODAY(),"Y")`
- Days since joining: `=TODAY()-E5`
- Dynamic sum: `=SUM(OFFSET(H5,0,0,6,1))`
- Dynamic reference: `=SUM(INDIRECT("H5:H10"))`

### Rounding and Filtering
- Round to a whole number: `=ROUND(D5,0)`
- Round up to a multiple of five: `=CEILING(I6,5)`
- Round down to a multiple of five: `=FLOOR(I6,5)`
- Filter students whose average is above 80: `=FILTER(A5:L16,G5:G16>80,"No matching students")`

## How to Use the Workbook
1. Open `Fundamental_Booster_Excel_Project.xlsx` in Microsoft Excel.
2. Start with the `Read Me` worksheet.
3. Select formula cells to inspect their formulas in the formula bar.
4. Change sample inputs, such as a product code or employee ID, to test lookup formulas.
5. Practise additional formulas in blank cells and save your completed copy.

## Requirements
- Microsoft Excel with formula support.
- `XLOOKUP`, `XMATCH`, and `FILTER` require a recent compatible Excel version.
- Automatic calculation should be enabled so formulas update when inputs change.

## Learning Outcomes
After completing this project, learners should be able to analyse student and sales data, retrieve values from lookup tables, work with text and dates, use dynamic references, and build formulas with conditional logic.

## Notes
- The workbook uses fictional sample data for practice.
- If data is moved, adjust formula references as needed.
- `FILTER` returns a dynamic array and requires empty cells for its results.
- The multi-condition `INDEX/MATCH` example may require a recent Excel version or a helper column in older versions.

## Author
Fundamental Booster Excel Practice Project
