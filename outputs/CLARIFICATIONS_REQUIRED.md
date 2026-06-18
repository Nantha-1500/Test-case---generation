1. Field listing attachments are referenced but not provided for multiple screens (Customer Details, Proposed Exposure, Security Detail, Exposure Norms, Pricing/Business Approval, etc.); confirm exact field names, mandatory flags, and validation rules.
2. Password policy mentions “minimum 8, alphanumeric, one uppercase,” while other places mention special characters; confirm final policy baseline.
3. OTP expiry duration, resend limits, and lockout threshold are not specified; confirm exact values.
4. NTB co-applicant/guarantor flow is marked on hold; confirm expected system behavior (hide, disable, or show informational message).
5. Deviation and sanction constraints contain limitation notes; confirm final sanction gating behavior when any deviation is rejected.
6. Proposal recommendation lower/upper sanction boundaries are ambiguous (“can’t be greater than recommended amount and less than loan amount”); confirm definitive rule.
7. API contracts and failure codes for CBS dedupe, liability fetch, Customer 360/CIC, external rating, CGTMSE push are not provided; confirm request/response and retry behavior.
8. Value of Account API availability is pending bank confirmation; confirm whether manual-only in current release.
9. Workflow matrix (all stages, role mappings, and reversible transitions) is partially described; provide full stage map for complete deterministic coverage.
10. Post Sanction Documentation detailed flow is referenced via annexure/manual but not attached; provide final flow and validations.
