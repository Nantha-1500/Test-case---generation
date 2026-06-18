# Clarifications Required

1. Detailed field-listing attachments referenced in SRS for many screens (Customer Details, Proposed Exposure, Security, Pricing, etc.) are not present in the repository; exact field-level constraints (lengths, formats, regex, ranges) need confirmation.
2. OTP expiry duration and max retry limits are referenced as configurable; exact values are required for precise boundary test data.
3. NTB co-applicant/guarantor flow is marked as on-hold; confirm expected UI behavior (hide/disable/informational message).
4. Security valuation and title search integrations depend on external API specs (Velocity/Adverse Risk) and mail templates not attached.
5. Exposure Norms and Value of Account API integration readiness from bank is pending confirmation.
6. Approved external rating agency list, rating categories, and instruments are pending confirmation.
7. Score Card inputs are pending from MSME department; current test cases use placeholder assumptions.
8. DOA matrix details (exact min/max slabs, hierarchy combinations) should be shared to finalize all authority-boundary test values.
9. Committee workflow and separate approval screen behavior are mentioned but not fully specified.
10. CGTMSE integration request/response contract and success/failure message catalog are not available.
