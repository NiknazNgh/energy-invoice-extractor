# Output schema

The primary Excel/CSV table uses these invoice fields in this exact order:

1. Production Month
2. From
3. To
4. Invoice Date
5. Power Factor
6. Load Factor
7. Actual Demand (KW)
8. Billing Demand (KW)
9. 4CP Charges Qty (KW)
10. 4CP Charges Rate ($/KW)
11. 4CP Charges ($)
12. Usage - Actual KWH
13. UOM
14. Energy Charge
15. Nodal Congestion Charge
16. Market Securitization (Debt) Financing - Default Charge
17. Prior Period Pass Through Charge
18. ERCOT Cont Reserve Serv (ECRS)
19. Firm Fuel Supply Service
20. Firm Fuel Supply Service - Backbill
21. Market Securitization - Uplift Charge
22. TX-ERCOT Admin Fees - CIL
23. Transmission Charges
24. Taxes & PUC Assessment Charge
25. Ancilliary Service Obligation Adjustment
26. Other Taxes
27. Bill Total

Two review fields are appended after the invoice schema:

- Review Status
- Review Notes

## Validation behavior

`Review Status` is `OK` when the parsed bill total reconciles within the configured tolerance and no other validation issue was detected.

`NEEDS REVIEW` is used when a required field is missing, 4CP quantity × rate does not match the parsed 4CP charge, or the bill total does not reconcile to the non-overlapping parsed summary buckets.

The reconciliation intentionally avoids double-counting nested values:

- 4CP is not added separately when it is already part of Transmission Charges.
- Nodal Congestion and Default Charge are not added again because they are included in Energy Charge.
- Prior-period detail fields are used only as a fallback when the prior-period total is unavailable.
