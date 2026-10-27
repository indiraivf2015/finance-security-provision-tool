# Test data (not in git)

Copy these files here before running the tests. They contain finance data, so they are git-ignored.

| File name here | What it is |
|---|---|
| extract_Sep-26.xlsx | Oracle ERP "IIVF Account Analysis Report" for Sep-26 (accounts 51701-51704) |
| statement_Aug-26.xlsx | Finance's Aug-26 Security Provision statement (opening for Sep-26) |
| provision_Sep-26.xlsx | Rajeev's provision file for Sep-26 ("HK & Security Sep-27 Provision.xlsx") |
| finance_statement_Sep-26.xlsx | Finance's own Sep-26 statement, for comparison |

| golden/security_sep26.xlsx | Finance's approved Sep-26 Security output (downloaded from the live page); the golden-output test compares against it |
| hk-expected.json | Figures printed in the Housekeeping SOPs (Part D rows, worked examples, Sep-26 totals), used by the SOP-figure tests |
