# Importing the openpyxl package
import openpyxl

path_to_acetylome_data = "Acetylome_Results_Contaminants_Removed.xlsx"
path_to_total_proteomics_data = "Total_Proteomics_Contaminants_FALSE_Table.xlsx"

# Loads the Acetylome_Results_Contaminants_Removed Excel workbook
workbook = openpyxl.load_workbook(path_to_acetylome_data, data_only=True)

# Sets the active sheet as the first sheet in the Acetylome_Results_Contaminants_Removed.xlsx Excel workbook
active_sheet = workbook.active


# Accepts the acetylated peptide accession as a parameter and finds and returns the average counts of the matching
# accession in the Total_Proteomics_Contaminants_FALSE_Table Excel workbook
def find_matching_accession(accession):
    workbook = openpyxl.load_workbook(path_to_total_proteomics_data, data_only=True)
    active_sheet = workbook.active
    p53_OFF_WCL_average = None
    p53_OFF_nuclear_average = None
    p53_OFF_cytoplasm_average = None
    p53_ON_WCL_average = None
    p53_ON_nuclear_average = None
    p53_ON_cytoplasm_average = None

    for row in active_sheet.iter_rows(min_row=2, min_col=1, max_col=1):
        for cell in row:
            if cell.value == accession:
                row_num = cell.row
                p53_OFF_WCL_average = active_sheet.cell(row=row_num, column=33).value
                p53_OFF_nuclear_average = active_sheet.cell(row=row_num, column=34).value
                p53_OFF_cytoplasm_average = active_sheet.cell(row=row_num, column=35).value
                p53_ON_WCL_average = active_sheet.cell(row=row_num, column=36).value
                p53_ON_nuclear_average = active_sheet.cell(row=row_num, column=37).value
                p53_ON_cytoplasm_average = active_sheet.cell(row=row_num, column=38).value
                return_list = [p53_OFF_WCL_average, p53_OFF_nuclear_average, p53_OFF_cytoplasm_average,
                               p53_ON_WCL_average, p53_ON_nuclear_average, p53_ON_cytoplasm_average]
                workbook = openpyxl.load_workbook(path_to_acetylome_data, data_only=True)
                active_sheet = workbook.active
                return return_list


# Loops through every accession in the Acetylome_Results_Contaminants_Removed.xlsx Excel workbook and passes the
# accession to the find_matching_accession(accession) method. If no match is found and there is more than one accession
# assigned to the given peptide, each subsequent accessions is passed to find_matching_accession(accession) until a
# match is found
for row in active_sheet.iter_rows(min_row=2, min_col=1, max_col=6):
    matched_accession_found = False
    if matched_accession_found is False:
        for cell in row:
            row_number = cell.row
            acetylome_accession = cell.value
            matching_total_proteome_averages = find_matching_accession(acetylome_accession)
            if matching_total_proteome_averages is not None:
                matched_accession_found = True
                break
    print(row_number, matching_total_proteome_averages)
