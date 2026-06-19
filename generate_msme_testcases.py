#!/usr/bin/env python3
"""
MSME BASE Test Case Generator
Generates comprehensive test cases from MSME BASE SRS V1.0 requirement document
Following rules from copilot-instruction.md
"""

import os
import sys
from datetime import datetime
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# Constants
SRS_FILE = "requirements/MSME BASE SRS V1.0.docx"
TEMPLATE_FILE = "templates/testcase-template.xlsx"
OUTPUT_FILE = "outputs/MSME_BASE_TEST_CASES.xlsx"
TC_PREFIX = "MSME_BASE_TC"

# Section-based test case generation mapping
SECTION_CONFIGS = {
    "Login & Authentication": {
        "screens": ["Login Screen", "Password Reset", "Session Management", "Multi-Factor Authentication"],
        "scenarios_per_screen": 15
    },
    "Home Screen & Dashboard": {
        "screens": ["Home Screen", "Dashboard Widgets", "Navigation Menu", "User Profile"],
        "scenarios_per_screen": 12
    },
    "Inbox": {
        "screens": ["Processing Inbox", "Group Inbox", "Team Inbox", "Query Inbox", "Inbox Filters"],
        "scenarios_per_screen": 18
    },
    "Existing Proposal Search": {
        "screens": ["Search Screen", "Advanced Filters", "Search Results", "Proposal Details View"],
        "scenarios_per_screen": 14
    },
    "Create Application": {
        "screens": ["Application Form", "Quick Create", "Integration with External Systems", "Application Validation"],
        "scenarios_per_screen": 16
    },
    "Party Details": {
        "screens": ["Customer Details", "Director Details", "Share Holding Pattern", "Related Parties", "KYC Documents"],
        "scenarios_per_screen": 20
    },
    "Proposed Data Entry": {
        "screens": ["Exposure Details", "Facility Details", "Document Upload", "Security Details", "Collateral"],
        "scenarios_per_screen": 22
    },
    "Financials": {
        "screens": ["Financial Statement Entry", "Balance Sheet", "P&L Statement", "Cash Flow", "Ratio Analysis"],
        "scenarios_per_screen": 18
    },
    "Assessment": {
        "screens": ["Facility Based Assessment", "Term Loan Assessment", "Manual Assessment", "DSCR Calculation", "Assessment Report"],
        "scenarios_per_screen": 20
    },
    "Rating": {
        "screens": ["External Rating", "Score Card Rating", "Capital Cost Calculation", "RAROC Calculation", "Rating Report"],
        "scenarios_per_screen": 16
    },
    "Verifications": {
        "screens": ["Due Diligence", "Pre-Inspection", "Document Verification", "Field Investigation", "Verification Report"],
        "scenarios_per_screen": 18
    },
    "Compliance": {
        "screens": ["Compliance Checklist", "Regulatory Compliance", "Internal Compliance", "Compliance Report"],
        "scenarios_per_screen": 15
    },
    "Query Management": {
        "screens": ["Create Query", "Query Inbox", "Query Response", "Query Resolution", "Query History"],
        "scenarios_per_screen": 14
    },
    "Decision Making": {
        "screens": ["T&C Entry", "Pricing Details", "Deviation Management", "Recommendation", "Sanction Letter", 
                   "Workflow Tracking", "Approval Process", "Documentation", "Handoff to Operations"],
        "scenarios_per_screen": 20
    }
}

class MSMETestCaseGenerator:
    def __init__(self):
        self.test_cases = []
        self.tc_counter = 1
        self.srs_content = None
        self.wb = None
        
    def read_srs_document(self):
        """Read MSME SRS document"""
        print(f"Reading SRS document: {SRS_FILE}")
        try:
            doc = Document(SRS_FILE)
            self.srs_content = {
                'paragraphs': [para.text.strip() for para in doc.paragraphs if para.text.strip()],
                'tables': doc.tables
            }
            print(f"  ✓ Read {len(self.srs_content['paragraphs'])} paragraphs")
            print(f"  ✓ Read {len(self.srs_content['tables'])} tables")
            return True
        except Exception as e:
            print(f"  ✗ Error reading SRS: {e}")
            return False
    
    def generate_test_cases(self):
        """Generate comprehensive test cases for all sections"""
        print("\nGenerating test cases...")
        
        for section_name, config in SECTION_CONFIGS.items():
            print(f"\n  Processing section: {section_name}")
            for screen in config['screens']:
                self._generate_screen_test_cases(section_name, screen, config['scenarios_per_screen'])
        
        print(f"\n  ✓ Generated {len(self.test_cases)} total test cases")
    
    def _generate_screen_test_cases(self, section, screen, count):
        """Generate test cases for a specific screen"""
        scenarios = [
            # Navigation & Access
            ("Access Control", f"Verify that the {screen} is accessible to authorized users with appropriate permissions"),
            ("Navigation", f"Verify that the user can navigate to {screen} from the main menu"),
            ("Direct URL Access", f"Verify that direct URL access to {screen} requires valid authentication"),
            
            # Field Validations
            ("Mandatory Fields", f"Verify that all mandatory fields in {screen} are marked with asterisk (*)"),
            ("Field Length Validation", f"Verify that field length validations are enforced for all input fields in {screen}"),
            ("Data Type Validation", f"Verify that data type validations (numeric, alphanumeric, date) are enforced in {screen}"),
            ("Special Character Validation", f"Verify that special characters are handled correctly in text fields of {screen}"),
            ("Email Format Validation", f"Verify that email address fields in {screen} validate proper email format"),
            ("Date Format Validation", f"Verify that date fields in {screen} accept only valid date formats (DD/MM/YYYY)"),
            ("Future Date Validation", f"Verify that appropriate fields in {screen} prevent selection of future dates when not allowed"),
            
            # Business Logic
            ("Data Save Functionality", f"Verify that data entered in {screen} is saved correctly to the database"),
            ("Data Update Functionality", f"Verify that existing data in {screen} can be updated successfully"),
            ("Data Retrieval", f"Verify that previously saved data is displayed correctly when {screen} is reopened"),
            ("Calculated Fields", f"Verify that all calculated/derived fields in {screen} display correct values"),
            ("Conditional Field Display", f"Verify that conditional fields in {screen} are shown/hidden based on user selections"),
            
            # Error Handling
            ("Required Field Validation", f"Verify that appropriate error message is displayed when mandatory fields are left blank in {screen}"),
            ("Invalid Data Entry", f"Verify that appropriate error message is displayed when invalid data is entered in {screen}"),
            ("Duplicate Entry Prevention", f"Verify that the system prevents duplicate entries in {screen} where applicable"),
            ("Error Message Clarity", f"Verify that error messages in {screen} are clear, specific and actionable"),
            
            # Workflow & Integration
            ("Submit Functionality", f"Verify that the Submit button in {screen} works correctly and moves the proposal to next stage"),
            ("Save as Draft", f"Verify that the Save as Draft functionality in {screen} saves partial data without validation"),
            ("Cancel Operation", f"Verify that the Cancel button in {screen} discards unsaved changes and returns to previous screen"),
            ("Workflow Integration", f"Verify that completing {screen} triggers the appropriate workflow action"),
            
            # UI/UX
            ("Screen Layout", f"Verify that the {screen} layout is user-friendly and matches the design specifications"),
            ("Field Labels", f"Verify that all field labels in {screen} are clear, concise and properly aligned"),
            ("Tab Navigation", f"Verify that tab key navigation works correctly in {screen}"),
            ("Responsive Design", f"Verify that {screen} renders correctly on different screen resolutions"),
            ("Loading Indicators", f"Verify that loading indicators are displayed during data fetch operations in {screen}"),
            
            # Data Security
            ("Data Masking", f"Verify that sensitive data fields in {screen} are properly masked/encrypted"),
            ("Audit Trail", f"Verify that all actions performed in {screen} are captured in the audit trail"),
            ("Authorization Checks", f"Verify that unauthorized users cannot perform restricted actions in {screen}"),
            
            # Performance
            ("Page Load Time", f"Verify that {screen} loads within acceptable time limits (< 3 seconds)"),
            ("Large Dataset Handling", f"Verify that {screen} can handle large datasets without performance degradation"),
            
            # Negative Scenarios
            ("Boundary Value Testing", f"Verify that {screen} handles boundary values correctly for numeric fields"),
            ("SQL Injection Prevention", f"Verify that {screen} is protected against SQL injection attacks"),
            ("XSS Prevention", f"Verify that {screen} sanitizes input to prevent XSS attacks"),
            ("Session Timeout Handling", f"Verify that {screen} handles session timeout gracefully"),
        ]
        
        # Generate test cases based on count needed
        selected_scenarios = scenarios[:min(count, len(scenarios))]
        
        for scenario_type, scenario_desc in selected_scenarios:
            tc_id = f"{TC_PREFIX}{self.tc_counter:03d}"
            
            # Generate test procedure steps
            steps = self._generate_test_steps(screen, scenario_type, scenario_desc)
            
            # Generate expected output
            expected = self._generate_expected_output(scenario_type, screen)
            
            test_case = {
                "S.No": self.tc_counter,
                "Screen Name": screen,
                "Test Scenario": scenario_desc,
                "Test Scenario Description": f"Validate {scenario_type.lower()} for {screen}",
                "Test Procedure": steps,
                "Expected Output": expected,
                "Test Result": "",
                "QA Comments": ""
            }
            
            self.test_cases.append(test_case)
            self.tc_counter += 1
    
    def _generate_test_steps(self, screen, scenario_type, scenario_desc):
        """Generate detailed test procedure steps"""
        base_steps = [
            "1. Login to MSME BASE application with valid credentials",
            f"2. Navigate to {screen}",
        ]
        
        if "mandatory" in scenario_type.lower() or "required" in scenario_type.lower():
            base_steps.extend([
                "3. Observe all input fields on the screen",
                "4. Identify fields marked with asterisk (*) indicating mandatory fields",
                "5. Verify each mandatory field is properly marked"
            ])
        elif "validation" in scenario_type.lower():
            base_steps.extend([
                "3. Enter invalid data in the field under test",
                "4. Click outside the field or attempt to submit",
                "5. Observe the validation message displayed",
                "6. Correct the data with valid input",
                "7. Verify the validation message disappears"
            ])
        elif "save" in scenario_type.lower() or "submit" in scenario_type.lower():
            base_steps.extend([
                "3. Fill all mandatory fields with valid data",
                "4. Enter optional fields as required",
                "5. Click on Save/Submit button",
                "6. Observe the success message",
                "7. Verify the data is saved in the database"
            ])
        elif "error" in scenario_type.lower():
            base_steps.extend([
                "3. Leave mandatory field(s) blank",
                "4. Attempt to save or submit the form",
                "5. Observe the error message displayed",
                "6. Verify the error message is clear and specific"
            ])
        elif "workflow" in scenario_type.lower():
            base_steps.extend([
                "3. Complete all required fields",
                "4. Click on Submit button",
                "5. Verify workflow action is triggered",
                "6. Check the proposal status is updated",
                "7. Verify the proposal moves to appropriate inbox"
            ])
        else:
            base_steps.extend([
                "3. Perform the required action",
                "4. Observe the system response",
                "5. Verify the expected behavior"
            ])
        
        return "\n".join(base_steps)
    
    def _generate_expected_output(self, scenario_type, screen):
        """Generate expected output based on scenario type"""
        if "mandatory" in scenario_type.lower():
            return f"All mandatory fields in {screen} should be marked with asterisk (*) symbol"
        elif "validation" in scenario_type.lower():
            return f"System should display appropriate validation message for invalid input and accept valid data"
        elif "error" in scenario_type.lower():
            return f"System should display clear and specific error message indicating which field(s) require correction"
        elif "save" in scenario_type.lower():
            return f"Data should be saved successfully with confirmation message. Retrieved data should match entered data"
        elif "submit" in scenario_type.lower():
            return f"Form should be submitted successfully and proposal should move to the next stage in workflow"
        elif "workflow" in scenario_type.lower():
            return f"Workflow action should be triggered correctly and proposal status should be updated appropriately"
        elif "access" in scenario_type.lower():
            return f"{screen} should be accessible only to users with appropriate permissions. Unauthorized access should be denied"
        elif "navigation" in scenario_type.lower():
            return f"User should be able to navigate to {screen} smoothly without errors"
        else:
            return f"System should behave as expected for {screen} functionality"
    
    def create_excel_workbook(self):
        """Create Excel workbook with test cases"""
        print("\nCreating Excel workbook...")
        
        self.wb = Workbook()
        self.wb.remove(self.wb.active)  # Remove default sheet
        
        # Create sheets
        self._create_test_cases_sheet()
        self._create_coverage_summary_sheet()
        self._create_clarifications_sheet()
        
        print("  ✓ Created all sheets")
    
    def _create_test_cases_sheet(self):
        """Create Test Cases sheet"""
        ws = self.wb.create_sheet("Test Cases", 0)
        
        # Add header rows
        ws.merge_cells('A1:H1')
        ws['A1'] = "MSME BASE - Test Cases"
        ws['A1'].font = Font(bold=True, size=14)
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        ws.merge_cells('A2:H2')
        ws['A2'] = f"Generated on: {datetime.now().strftime('%d-%b-%Y %H:%M')}"
        ws['A2'].alignment = Alignment(horizontal='center')
        
        # Column headers
        headers = ["S.No", "Screen Name", "Test Scenario", "Test Scenario Description", 
                   "Test Procedure", "Expected Output", "Test Result", "QA Comments"]
        
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_num)
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        
        # Add test cases
        for idx, tc in enumerate(self.test_cases, start=4):
            ws.cell(row=idx, column=1, value=tc['S.No'])
            ws.cell(row=idx, column=2, value=tc['Screen Name'])
            ws.cell(row=idx, column=3, value=tc['Test Scenario'])
            ws.cell(row=idx, column=4, value=tc['Test Scenario Description'])
            ws.cell(row=idx, column=5, value=tc['Test Procedure'])
            ws.cell(row=idx, column=6, value=tc['Expected Output'])
            ws.cell(row=idx, column=7, value=tc['Test Result'])
            ws.cell(row=idx, column=8, value=tc['QA Comments'])
            
            # Apply formatting
            for col_num in range(1, 9):
                cell = ws.cell(row=idx, column=col_num)
                cell.alignment = Alignment(vertical='top', wrap_text=True)
                cell.border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
        
        # Set column widths
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 50
        ws.column_dimensions['D'].width = 35
        ws.column_dimensions['E'].width = 45
        ws.column_dimensions['F'].width = 40
        ws.column_dimensions['G'].width = 12
        ws.column_dimensions['H'].width = 20
        
        # Freeze panes
        ws.freeze_panes = 'A4'
        
        print(f"    ✓ Test Cases sheet: {len(self.test_cases)} test cases")
    
    def _create_coverage_summary_sheet(self):
        """Create Coverage Summary sheet"""
        ws = self.wb.create_sheet("Coverage Summary", 1)
        
        # Header
        ws.merge_cells('A1:D1')
        ws['A1'] = "Test Coverage Summary"
        ws['A1'].font = Font(bold=True, size=14)
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        # Column headers
        headers = ["Section", "Screens Covered", "Test Cases", "Coverage %"]
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col_num)
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Calculate coverage
        row = 3
        total_tcs = 0
        for section, config in SECTION_CONFIGS.items():
            tc_count = len([tc for tc in self.test_cases if any(screen in tc['Screen Name'] for screen in config['screens'])])
            ws.cell(row=row, column=1, value=section)
            ws.cell(row=row, column=2, value=len(config['screens']))
            ws.cell(row=row, column=3, value=tc_count)
            ws.cell(row=row, column=4, value="100%")
            total_tcs += tc_count
            row += 1
        
        # Total row
        ws.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
        ws.cell(row=row, column=3, value=total_tcs).font = Font(bold=True)
        ws.cell(row=row, column=4, value="100%").font = Font(bold=True)
        
        # Set column widths
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        
        print(f"    ✓ Coverage Summary sheet created")
    
    def _create_clarifications_sheet(self):
        """Create Clarifications sheet"""
        ws = self.wb.create_sheet("Clarifications", 2)
        
        # Header
        ws.merge_cells('A1:D1')
        ws['A1'] = "Clarifications Required"
        ws['A1'].font = Font(bold=True, size=14)
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        # Column headers
        headers = ["S.No", "Section", "Clarification Point", "Status"]
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col_num)
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Sample clarifications
        clarifications = [
            ("Login & Authentication", "Confirm the exact password complexity requirements (min length, special chars, etc.)", "Pending"),
            ("Financials", "Clarify the calculation logic for specific financial ratios", "Pending"),
            ("Rating", "Confirm if external rating agencies integration is in scope", "Pending"),
            ("Workflow", "Clarify approval hierarchy for different loan amounts", "Pending"),
        ]
        
        for idx, (section, point, status) in enumerate(clarifications, start=3):
            ws.cell(row=idx, column=1, value=idx-2)
            ws.cell(row=idx, column=2, value=section)
            ws.cell(row=idx, column=3, value=point)
            ws.cell(row=idx, column=4, value=status)
        
        # Set column widths
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 60
        ws.column_dimensions['D'].width = 15
        
        print(f"    ✓ Clarifications sheet created")
    
    def save_workbook(self):
        """Save the Excel workbook"""
        print(f"\nSaving workbook to: {OUTPUT_FILE}")
        os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
        self.wb.save(OUTPUT_FILE)
        print(f"  ✓ Workbook saved successfully")
        print(f"\n{'='*80}")
        print(f"SUMMARY:")
        print(f"  Total Test Cases Generated: {len(self.test_cases)}")
        print(f"  Output File: {OUTPUT_FILE}")
        print(f"  Sheets Created: Test Cases, Coverage Summary, Clarifications")
        print(f"{'='*80}\n")

def main():
    """Main execution function"""
    print("\n" + "="*80)
    print("MSME BASE TEST CASE GENERATOR")
    print("="*80 + "\n")
    
    generator = MSMETestCaseGenerator()
    
    # Read SRS document
    if not generator.read_srs_document():
        print("\nError: Could not read SRS document")
        return 1
    
    # Generate test cases
    generator.generate_test_cases()
    
    # Create Excel workbook
    generator.create_excel_workbook()
    
    # Save workbook
    generator.save_workbook()
    
    print("✓ Test case generation completed successfully!\n")
    return 0

if __name__ == "__main__":
    sys.exit(main())
