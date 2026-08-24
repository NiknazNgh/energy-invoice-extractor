# Natural Gas Output Schema

The Atmos natural-gas extractor writes a multi-sheet Excel workbook.

## Gas Invoice Summary

The summary sheet uses these columns in order:

1. Month
2. Service Account
3. Customer Name
4. Facility ID
5. Address
6. City
7. State
8. Zip
9. Current Charges
10. Billed MCF
11. Billed MMBtu
12. 1st 0 to 1,500 MMBtu Rate
13. Next 3,500 MMBtu Rate
14. All MMBtu over 5,000 MMBtu Rate
15. 1st 0 to 1,500 MMBtu Amount
16. Next 3,500 MMBtu Amount
17. All MMBtu over 5,000 MMBtu Amount
18. Total Cost Based on MMBtu
19. Customer Charge
20. Plant Protection Fee
21. Reimbursement of MGRT 1
22. Reimbursement of MGRT 2
23. Reimbursement of MGRT 3
24. Reimbursement of MGRT 4
25. Total Reimbursement of MGRT
26. Street & Alley Fees 1
27. Street & Alley Fees 2
28. Street & Alley Fees 3
29. Street & Alley Fees 4
30. Total Street & Alley Fees
31. FIN 48 Tax Refund
32. Pipeline Safety Fee
33. GCR - Industrial Sales ($/MMBtu)
34. GCR - Transportation ($/MMBtu)
35. Billed CCF
36. Calculated Rate ($/MMBTU)
37. From Billing Date
38. To Billing Date
39. Extra Charge Details
40. Extra Charge Total
41. Review Status
42. Review Notes

## Raw Line Items

Keeps the parsed invoice rows used to construct the summary, including source PDF, line number, production month, service account, description, MCF, MMBtu, rate, amount, and review flag.

## Validation

Compares `Summary Current Charges` with the sum of parsed line items. The sheet records the difference, math status, review status, and review notes.

## Review behavior

Known standard descriptions are summarized into fixed columns. Unexpected or different charge descriptions are retained in `Extra Charge Details` and marked `NEEDS REVIEW` rather than silently discarded.

The parser is rule-based and tailored to the Atmos invoice layouts used by this project. Always review flagged invoices before relying on the workbook for downstream reporting.
