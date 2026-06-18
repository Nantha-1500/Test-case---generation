# CLARIFICATIONS REQUIRED

1. **Field listing attachments referenced but missing in repository**
   - Multiple sections state "please find the attached field listing" (Customer Details, Proposed Exposure, Document Details, Security Details, Exposure Norms, Personal/Corporate Guarantee, Due Diligence, etc.) but those annexure files are not present in `requirements/`.
   - Please share all referenced field-level annexures to finalize exact control-level validations and error messages.

2. **Conflicting rule for Requested Loan Amount editability**
   - SRS states requested amount cannot be changed once initiated, but also states it can be changed within ES-note limit.
   - Please confirm final business rule and exact stage(s) where edit is permitted.

3. **NTB co-applicant/guarantor onboarding decision pending**
   - SRS marks NTB onboarding as on hold.
   - Please confirm whether NTB flow should be disabled in current release or enabled with conditions.

4. **Security Coverage calculation logic pending business confirmation**
   - Formula references are partially provided but SRS also says business must confirm logic.
   - Please provide final approved formulas for primary, collateral, customer coverage, and total coverage computation.

5. **External integrations pending API specifications**
   - API specs are missing for Valuation (Velocity), Title Search (Adverse Risk), External Rating (ICON/agency mapping), Account Handoff/CGTMSE payload contracts.
   - Please share request/response schema, timeout/retry/error code standards, and idempotency rules.

6. **Role and permission matrix not fully defined**
   - SRS mentions examples (processing officer, supervisor, special officer, MLP officer) but full role-action matrix is not attached.
   - Please share complete role-wise screen/action permissions and delegation mappings.
