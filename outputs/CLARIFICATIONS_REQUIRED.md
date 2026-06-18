# Clarifications Required

1. Detailed field listings are referenced as attached in multiple sections (Customer Details, Exposure pages, Security, Assessment, Verification, etc.) but are not available as separate readable attachments in the repository.
2. API payload/response definitions and error-code catalogs are pending for integrations: CBS liability fetch, dedupe, Customer 360 bureaus, valuation/title search vendors, ICON rating, and CGTMSE push.
3. Exact OTP expiry duration and retry/resend threshold are not explicitly specified.
4. Precise validation rules (min/max lengths, regex, value ranges) for many form fields are not explicitly specified.
5. Final business logic confirmation is pending for some calculations and limitations (e.g., security coverage logic confirmation, deviation limitations noted in SRS).
6. Committee approval screen details and screen-level fields are mentioned but not fully defined.
7. UI label-level alert/error message master for all screens is not fully specified.

## Assumptions Used While Generating Test Cases

- Standard banking LOS validations were assumed where exact field constraints were not provided.
- Mandatory validations were created for fields explicitly marked mandatory and for critical flow actions.
- Integration test cases include retry/error handling based on expected API behavior in enterprise systems.
- Role and permission coverage was derived from described user groups (Processing Officer, Approver, Supervisory roles, Admin-configured access).
