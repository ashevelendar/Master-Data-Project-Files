SUPPLIER CHECKING TOOL - COMPANIES HOUSE REPAIR

Application: Supplier Checking Tool
Branch: supplier-thing

Repair files are provided as plain text so they can be viewed in GitHub without
downloading an archive.

Installation order:
1. Back up the Access database.
2. Export a backup of Module4.
3. Replace the existing Module4 code with Module4_CompaniesHouse_FIXED.txt.
   The file contains VBA source. Copy its contents into the existing Module4
   module, or save a local copy as .bas and import it.
4. Open the GTIN form's code module and replace its entire contents with
   Form_GTIN_CODE_FIXED.txt.
5. In the GTIN form Property Sheet, Event tab, set On Timer to [Event Procedure].
6. Run Debug > Compile.
7. Test using a supplier archive record with a valid CompanyRegistrationNo.

Important:
- Do not store or commit the Companies House API key in GitHub.
- The key should remain in the local SUPP_tblSettings table.
- Test in a copy of the Access database first.
- The lookup is deliberately started by a one-shot Form_Timer after the GTIN
  form opens, rather than during Form_Load.
- Address components use New field values first and corresponding Current
  values when the New field is blank.
- The comparison subform textbox is named txtCompanyRegistrationNumber.

Files in this repair set:
- Module4_CompaniesHouse_FIXED.txt
- Form_GTIN_CODE_FIXED.txt
- CompaniesHouse_Repair_README.txt
