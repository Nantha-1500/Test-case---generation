# Clarifications Required - MSME BASE SRS V1.0

The following items require confirmation because detailed artifacts are referenced but not attached in the current package:

1. Detailed field-list annexures referenced for multiple screens (Customer Details, Exposure, Security, Documents, Exposure Norms, Guarantees, Evaluation, etc.).
2. Exact error-message catalog for all authentication/API failure states.
3. OTP controls: validity period, resend cooldown, max retries, lock duration.
4. Full Undo Sanction workflow and authority matrix.
5. Final approved formulas for Security Coverage and related derived values.
6. Integration specs pending for ICON, Velocity, Adverse Risk, CGTMSE, and some CBS touchpoints.
7. Approved external rating agency/category/instrument master list.
8. Score Card input model pending MSME department input.
9. Committee approval specific screen/routing details.
10. Detailed post-sanction documentation annexure mapping.

Assumptions used in test generation:

- Referenced screens/modules exist and are role-configured through setup.
- External APIs provide standard success/failure responses with retry support.
- Workflow events are timestamped and auditable.
- Setup masters control dropdown values, templates, and formula configurations.
