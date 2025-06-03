
import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import os
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

# Set page title and layout
st.set_page_config(page_title="NTSL Data Processor", layout="wide")

def main():
    st.title("NTSL Data Processing Tool")
    
    # File upload section
    st.header("Step 1: Upload ZIP File")
    uploaded_file = st.file_uploader("Upload your NTSL ZIP file", type="zip")
    
    if uploaded_file is not None:
        with st.spinner("Processing files..."):
            # Process the uploaded file through all steps
            process_all_steps(uploaded_file)

def process_all_steps(uploaded_file):
    # Step 1: Process the ZIP file
    combined_data_path = "combined_data.xlsx"
    filter_zip_excel_data(uploaded_file, combined_data_path)
    
    # Step 2: Clean descriptions
    output_file_path = "output_file.xlsx"
    process_excel_file(combined_data_path, output_file_path)
    
    # Step 3: Process the cleaned data
    combined_output_path = "combined_output.xlsx"
    process_combined_output(output_file_path, combined_output_path)
    
    # Step 4: Aggregate the data
    combined_aggregated_path = "combined_aggregated_output.xlsx"
    process_aggregated_output(output_file_path, combined_aggregated_path)
    
    # Display download buttons
    st.success("Processing complete!")
    
    col1, col2 = st.columns(2)
    with col1:
        with open(combined_output_path, "rb") as f:
            st.download_button(
                label="Download Combined Output",
                data=f,
                file_name="combined_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    with col2:
        with open(combined_aggregated_path, "rb") as f:
            st.download_button(
                label="Download Aggregated Output",
                data=f,
                file_name="combined_aggregated_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

def filter_zip_excel_data(zip_file, output_file):
    # Create a new Excel writer to save filtered data into different sheets
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Open the ZIP file
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            # List all files in the zip archive
            extracted_files = [f for f in zip_ref.namelist() if f.endswith('.xls')]

            # List to store the dataframes
            dfs = []
            progress_bar = st.progress(0)
            total_files = len(extracted_files)

            # Process each file
            for i, file in enumerate(extracted_files, start=1):
                progress_bar.progress(i / total_files)
                with zip_ref.open(file) as file_data:
                    # Read the excel file from the ZIP stream
                    df = pd.read_excel(file_data, header=None)

                    # Find the row index where the required headers exist
                    header_row_index = None
                    for row_index, row in df.iterrows():
                        if 'Description' in row.values and 'No of Txns' in row.values and 'Debit' in row.values and 'Credit' in row.values:
                            header_row_index = row_index
                            break

                    if header_row_index is None:
                        st.warning(f"Headers not found in {file}. Skipping.")
                        continue

                    # Set the correct header row
                    df.columns = df.iloc[header_row_index]
                    df = df[header_row_index + 1:]

                    # Reset index
                    df.reset_index(drop=True, inplace=True)

                    # Skip empty dataframes
                    if df.empty:
                        st.warning(f"No valid data in {file}. Skipping.")
                        continue

                    # Add the dataframe to the list
                    dfs.append(df)

            # Move the first sheet to the last position
            if dfs:
                first_sheet = dfs.pop(0)
                dfs.append(first_sheet)

                # Write to the Excel file with adjusted sheet order
                for i, df in enumerate(dfs, start=1):
                    sheet_name = f"sheet{i}"
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

    st.success(f"Data from all Excel files has been saved to {output_file}")

# Define the suffixes you want to remove
suffixes = [" - CC", " - CC -Paid", " - CC -Received"]

def clean_description(description):
    
    # Ensure that description is a string before processing
    if isinstance(description, str):
        # Remove any leading/trailing spaces before processing
        description = description.strip()

        original_description = description  # For debugging
        # Loop through the suffixes and remove them if they exist at the end
        for suffix in suffixes:
            if description.endswith(suffix):
                description = description[:-len(suffix)]  # Remove the suffix
                break  # Remove only the first matched suffix (as it's at the end)

        # After removing the suffix, strip any extra spaces that may remain
        description = description.strip()

        # Print out the transformation for debugging (optional)
        # if original_description != description:
        #     # st.write(f"Changed: {original_description} -> {description}")
    else:
        description = ''  # If it's not a string, clear it (you can modify this if needed)

    return description

def process_excel_file(input_file, output_file):
    # Check if the input file exists
    if not os.path.exists(input_file):
        st.error(f"Error: The file {input_file} does not exist.")
        return

    # Load the Excel file
    xl = pd.ExcelFile(input_file)

    # Prepare to write to output file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Iterate through each sheet
        sheet_names = xl.sheet_names
        progress_bar = st.progress(0)
        total_sheets = len(sheet_names)

        for i, sheet_name in enumerate(sheet_names):
            progress_bar.progress((i + 1) / total_sheets)
            # st.write(f"Processing sheet: {sheet_name}")

            # Load the sheet into a DataFrame
            df = xl.parse(sheet_name)

            # Ensure that the column you're processing is named 'Description' (case-sensitive)
            if 'Description' in df.columns:
                # Apply the clean_description function to each row in the 'Description' column
                df['Description'] = df['Description'].apply(clean_description)

            # Save the updated DataFrame back to the same sheet in the output Excel file
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    st.success(f"Processing complete. The modified file is saved as {output_file}.")

# def process_combined_output(file_path, output_file):
#     # Create an Excel writer object
#     with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
#         # Code 1 - Part 1
#         results_1 = []
#         excel_sheets = pd.ExcelFile(file_path).sheet_names

#         progress_bar = st.progress(0)
#         total_sheets = len(excel_sheets)

#         for i, sheet_name in enumerate(excel_sheets):
#             progress_bar.progress((i + 1) / total_sheets)
#             df = pd.read_excel(file_path, sheet_name=sheet_name)
#             filtered_rows = df[df['Description'].str.startswith('Beneficiary', na=False) &
#                                (df['Description'].str.endswith('Approved Transaction Amount', na=False) |
#                                 df['Description'].str.endswith('U3 RB Approved Transaction Amount', na=False))]
#             summed_row = filtered_rows[['No of Txns', 'Debit', 'Credit']].sum()
#             new_row = {
#                 'sheetname': sheet_name,
#                 'Description': 'Beneficiary Approved Transaction Amount U3',
#                 'No of Txns': summed_row['No of Txns'],
#                 'Debit': summed_row['Debit'],
#                 'Credit': summed_row['Credit']
#             }
#             results_1.append(new_row)

#             output_df_1 = pd.DataFrame(results_1)
#             output_df_1.to_excel(writer, index=False, sheet_name="Combined", startrow=0)

def process_combined_output(file_path, output_file):
    # Create an Excel writer object
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Code 1 - Part 1
        results_1 = []
        excel_sheets = pd.ExcelFile(file_path).sheet_names

        progress_bar = st.progress(0)
        total_sheets = len(excel_sheets)

        for i, sheet_name in enumerate(excel_sheets):
            progress_bar.progress((i + 1) / total_sheets)
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            filtered_rows = df[df['Description'].str.startswith('Beneficiary', na=False) &
                               (df['Description'].str.endswith('Approved Transaction Amount', na=False) |
                                df['Description'].str.endswith('U3 RB Approved Transaction Amount', na=False))]
            summed_row = filtered_rows[['No of Txns', 'Debit', 'Credit']].sum()
            new_row = {
                'Cycle': sheet_name,
                'Description': 'Beneficiary Approved Transaction Amount U3',
                'No of Txns': summed_row['No of Txns'],
                'Debit': summed_row['Debit'],
                'Credit': summed_row['Credit']
            }
            results_1.append(new_row)

        output_df_1 = pd.DataFrame(results_1)
        output_df_1.to_excel(writer, index=False, sheet_name="Combined", startrow=0)
        
        # Get the sheet object to append more data
        workbook = writer.book
        sheet = workbook["Combined"]
        
        # Add blank rows (6 lines)
        for _ in range(6):
            sheet.append([])
            
        # Beneficiary U2 Approved Transaction Amount
        results_21 = []
        for sheet_name in excel_sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            if not all(col in df.columns for col in ['Description', 'No of Txns', 'Debit', 'Credit']):
                continue
            filtered_df = df[df['Description'].str.startswith('Beneficiary', na=False) &
                             df['Description'].str.endswith('U2 Approved Transaction Amount', na=False)]
            if not filtered_df.empty:
                results_21.append({
                    'Cycle': sheet_name,
                    'Description': "Beneficiary U2 Approved Transaction Amount",
                    'No of Txns': filtered_df['No of Txns'].sum(),
                    'Debit': filtered_df['Debit'].sum(),
                    'Credit': filtered_df['Credit'].sum()
                })
        output_df_21 = pd.DataFrame(results_21)
        for r in dataframe_to_rows(output_df_21, index=False, header=True):
            sheet.append(r)

        # Add blank rows (6 lines)
        for _ in range(6):
            sheet.append([])

        # Beneficiary U2 RB Approved Transaction Amount
        results_22 = []
        for sheet_name in excel_sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            if not all(col in df.columns for col in ['Description', 'No of Txns', 'Debit', 'Credit']):
                continue
            filtered_df = df[df['Description'].str.startswith('Beneficiary', na=False) &
                             df['Description'].str.endswith('U2 RB Approved Transaction Amount', na=False)]
            if not filtered_df.empty:
                results_22.append({
                    'Cycle': sheet_name,
                    'Description': "Beneficiary U2 RB Approved Transaction Amount",
                    'No of Txns': filtered_df['No of Txns'].sum(),
                    'Debit': filtered_df['Debit'].sum(),
                    'Credit': filtered_df['Credit'].sum()
                })
        output_df_22 = pd.DataFrame(results_22)
        for r in dataframe_to_rows(output_df_22, index=False, header=True):
            sheet.append(r)

        # Add blank rows (6 lines)
        for _ in range(6):
            sheet.append([])

        # Beneficiary U3 Approved Transaction Amount
        results_23 = []
        for sheet_name in excel_sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            if not all(col in df.columns for col in ['Description', 'No of Txns', 'Debit', 'Credit']):
                continue
            filtered_df = df[df['Description'].str.startswith('Beneficiary', na=False) &
                             df['Description'].str.endswith('U3 Approved Transaction Amount', na=False)]
            if not filtered_df.empty:
                results_23.append({
                    'Cycle': sheet_name,
                    'Description': "Beneficiary U3 Approved Transaction Amount",
                    'No of Txns': filtered_df['No of Txns'].sum(),
                    'Debit': filtered_df['Debit'].sum(),
                    'Credit': filtered_df['Credit'].sum()
                })
        output_df_23 = pd.DataFrame(results_23)
        for r in dataframe_to_rows(output_df_23, index=False, header=True):
            sheet.append(r)

        # Add blank rows (6 lines)
        for _ in range(6):
            sheet.append([])

        # Beneficiary U3 RB Approved Transaction Amount
        results_24 = []
        for sheet_name in excel_sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            if not all(col in df.columns for col in ['Description', 'No of Txns', 'Debit', 'Credit']):
                continue
            filtered_df = df[df['Description'].str.startswith('Beneficiary', na=False) &
                             df['Description'].str.endswith('U3 RB Approved Transaction Amount', na=False)]
            if not filtered_df.empty:
                results_24.append({
                    'Cycle': sheet_name,
                    'Description': "Beneficiary U3 RB Approved Transaction Amount",
                    'No of Txns': filtered_df['No of Txns'].sum(),
                    'Debit': filtered_df['Debit'].sum(),
                    'Credit': filtered_df['Credit'].sum()
                })
        output_df_24 = pd.DataFrame(results_24)
        for r in dataframe_to_rows(output_df_24, index=False, header=True):
            sheet.append(r)

        # Add blank rows (6 lines)
        for _ in range(6):
            sheet.append([])

        # Remitter Approved Transaction Amount
        results_2 = []
        for sheet_name in excel_sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            if not all(col in df.columns for col in ['Description', 'No of Txns', 'Debit', 'Credit']):
                continue
            filtered_df = df[df['Description'].str.startswith('Remitter', na=False) &
                             df['Description'].str.endswith('Transaction Amount', na=False)]
            if not filtered_df.empty:
                results_2.append({
                    'Cycle': sheet_name,
                    'Description': "Remitter Approved Transaction Amount",
                    'No of Txns': filtered_df['No of Txns'].sum(),
                    'Debit': filtered_df['Debit'].sum(),
                    'Credit': filtered_df['Credit'].sum()
                })
        output_df_2 = pd.DataFrame(results_2)
        for r in dataframe_to_rows(output_df_2, index=False, header=True):
            sheet.append(r)

        # Add blank rows (6 lines)
        for _ in range(6):
            sheet.append([])

        # Remitter U2 Approved Transaction Amount
        results_28 = []
        for sheet_name in excel_sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            if not all(col in df.columns for col in ['Description', 'No of Txns', 'Debit', 'Credit']):
                continue
            filtered_df = df[df['Description'].str.startswith('Remitter', na=False) &
                             df['Description'].str.endswith('U2 Approved Transaction Amount', na=False)]
            if not filtered_df.empty:
                results_28.append({
                    'Cycle': sheet_name,
                    'Description': "Remitter U2 Approved Transaction Amount",
                    'No of Txns': filtered_df['No of Txns'].sum(),
                    'Debit': filtered_df['Debit'].sum(),
                    'Credit': filtered_df['Credit'].sum()
                })
        output_df_28 = pd.DataFrame(results_28)
        for r in dataframe_to_rows(output_df_28, index=False, header=True):
            sheet.append(r)

        # Add blank rows (6 lines)
        for _ in range(6):
            sheet.append([])

        # Remitter U2 RB Approved Transaction Amount
        results_29 = []
        for sheet_name in excel_sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            if not all(col in df.columns for col in ['Description', 'No of Txns', 'Debit', 'Credit']):
                continue
            filtered_df = df[df['Description'].str.startswith('Remitter', na=False) &
                             df['Description'].str.endswith('U2 RB Approved Transaction Amount', na=False)]
            if not filtered_df.empty:
                results_29.append({
                    'Cycle': sheet_name,
                    'Description': "Remitter U2 RB Approved Transaction Amount",
                    'No of Txns': filtered_df['No of Txns'].sum(),
                    'Debit': filtered_df['Debit'].sum(),
                    'Credit': filtered_df['Credit'].sum()
                })
        output_df_29 = pd.DataFrame(results_29)
        for r in dataframe_to_rows(output_df_29, index=False, header=True):
            sheet.append(r)

        # Add blank rows (6 lines)
        for _ in range(6):
            sheet.append([])

        # Remitter U3 Approved Transaction Amount
        results_30 = []
        for sheet_name in excel_sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            if not all(col in df.columns for col in ['Description', 'No of Txns', 'Debit', 'Credit']):
                continue
            filtered_df = df[df['Description'].str.startswith('Remitter', na=False) &
                             df['Description'].str.endswith('U3 Approved Transaction Amount', na=False)]
            if not filtered_df.empty:
                results_30.append({
                    'Cycle': sheet_name,
                    'Description': "Remitter U3 Approved Transaction Amount",
                    'No of Txns': filtered_df['No of Txns'].sum(),
                    'Debit': filtered_df['Debit'].sum(),
                    'Credit': filtered_df['Credit'].sum()
                })
        output_df_30 = pd.DataFrame(results_30)
        for r in dataframe_to_rows(output_df_30, index=False, header=True):
            sheet.append(r)

        # Add blank rows (6 lines)
        for _ in range(6):
            sheet.append([])

        # Remitter U3 RB Approved Transaction Amount
        results_31 = []
        for sheet_name in excel_sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            if not all(col in df.columns for col in ['Description', 'No of Txns', 'Debit', 'Credit']):
                continue
            filtered_df = df[df['Description'].str.startswith('Remitter', na=False) &
                             df['Description'].str.endswith('U3 RB Approved Transaction Amount', na=False)]
            if not filtered_df.empty:
                results_31.append({
                    'Cycle': sheet_name,
                    'Description': "Remitter U3 RB Approved Transaction Amount",
                    'No of Txns': filtered_df['No of Txns'].sum(),
                    'Debit': filtered_df['Debit'].sum(),
                    'Credit': filtered_df['Credit'].sum()
                })
        output_df_31 = pd.DataFrame(results_31)
        for r in dataframe_to_rows(output_df_31, index=False, header=True):
            sheet.append(r)

        # Add blank rows (6 lines)
        for _ in range(6):
            sheet.append([])

        # Net Adjusted Amount with difference calculation
        results_3 = []
        difference_dict = {}  # To store differences for later use
        for sheet_name in excel_sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            # Filter for 'Net Adjusted Amount' rows
            filtered_row = df[df['Description'] == 'Net Adjusted Amount'].copy()

            if not filtered_row.empty:
                filtered_row.loc[:, 'Cycle'] = sheet_name
                filtered_row.loc[:, 'difference_debit_credit'] = filtered_row['Credit'] - filtered_row['Debit']
                results_3.append(filtered_row)
                difference_dict[sheet_name] = filtered_row.iloc[0]['difference_debit_credit']

        if results_3:
            output_df_3 = pd.concat(results_3, ignore_index=True)
            output_df_3 = output_df_3[['Cycle', 'Description', 'No of Txns', 'Debit', 'Credit', 'difference_debit_credit']]
            for r in dataframe_to_rows(output_df_3, index=False, header=True):
                sheet.append(r)

        # Add blank rows (6 lines)
        for _ in range(6):
            sheet.append([])

        # Beneficiary/Remitter Sub Totals and Settlement Amount
        results_2 = []
        for sheet_name in excel_sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            df.columns = [col.strip() for col in df.columns]

            filtered_row = df[df['Description'] == 'Beneficiary / Remitter Sub Totals']
            if not filtered_row.empty:
                debit = filtered_row['Debit'].sum()
                credit = filtered_row['Credit'].sum()
            else:
                debit = 0
                credit = 0

            settlement_row = df[df['Description'] == 'Settlement Amount']
            if not settlement_row.empty:
                settlement_amount = settlement_row['Debit'].sum() + settlement_row['Credit'].sum()
            else:
                settlement_amount = 0

            results_2.append([sheet_name, debit, credit, settlement_amount])

        output_df_code2 = pd.DataFrame(results_2, columns=['Beneficiary / Remitter Sub Totals', 'DR (Amount)', 'CR (Amount)', 'NTSL Settlement Amount'])
        for r in dataframe_to_rows(output_df_code2, index=False, header=True):
            sheet.append(r)

        # Add blank rows (6 lines)
        for _ in range(6):
            sheet.append([])

        # Final Settlement Amount with difference calculation
        results_code3 = pd.DataFrame(columns=['Final Settlement Amount', 'DR(Amount)', 'CR(Amount)', 'Difference In Settlement'])
        for sheet_name in excel_sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            df.columns = df.columns.str.strip()

            # Extract relevant rows
            final_settlement_row = df[df['Description'].str.contains('Final Settlement Amount', case=False, na=False)]
            settlement_row = df[df['Description'] == 'Settlement Amount']

            if not final_settlement_row.empty and not settlement_row.empty:
                final_debit = final_settlement_row.iloc[0]['Debit'] if 'Debit' in final_settlement_row.columns else 0
                final_credit = final_settlement_row.iloc[0]['Credit'] if 'Credit' in final_settlement_row.columns else 0
                final_settlement_amount = final_debit - final_credit

                settlement_debit = settlement_row.iloc[0]['Debit'] if 'Debit' in settlement_row.columns else 0
                settlement_credit = settlement_row.iloc[0]['Credit'] if 'Credit' in settlement_row.columns else 0
                ntsl_settlement_amount = settlement_debit - settlement_credit

                # Fetch previously calculated difference_debit_credit
                difference_debit_credit = difference_dict.get(sheet_name, 0)

                # Calculate difference
                difference = (final_settlement_amount - ntsl_settlement_amount ) + difference_debit_credit

                results_code3 = pd.concat([results_code3, pd.DataFrame({
                    'Final Settlement Amount': [sheet_name],
                    'DR(Amount)': [final_debit],
                    'CR(Amount)': [final_credit],
                    'Difference In Settlement': [round(difference)]
                })], ignore_index=True)

        # Write to Excel
        for r in dataframe_to_rows(results_code3, index=False, header=True):
            sheet.append(r)


        st.success(f"Final combined output saved to {output_file}")

def process_aggregated_output(file_path, output_file_path):
    # Define conditions for Code 1 (Remitter) and Code 2 (Beneficiary)
    remitter_conditions = [
        ("Remitter", "U2 Approved Fee"),
        ("Remitter", "U2 Approved Fee Gst"),
        ("Remitter", "U2 Approved NPCI Switching Fee"),
        ("Remitter", "U2 Approved NPCI Switching Fee Gst"),
        ("Remitter", "U2 RB Approved NPCI Switching Fee"),
        ("Remitter", "U2 RB Approved NPCI Switching Fee Gst"),
        ("Remitter", "U3 RB Approved NPCI Switching Fee"),
        ("Remitter", "U3 RB Approved NPCI Switching Fee Gst"),
        (None, "U2 RB Approved Payer PSP Fee"),
        (None, "U2 RB Approved Payer PSP Fee Gst"),
        ("Remitter", "U3 RB Approved Payer PSP Fee"),
        ("Remitter", "U3 RB Approved Payer PSP Fee Gst"),
        (None, "U2 Approved Payer PSP Fee"),
        (None, "U2 Approved Payer PSP Fee Gst"),
        (None, "U3 RB Approved Fee"),
        (None, "U3 RB Approved Fee Gst"),
        (None, "U3 Approved Fee"),
        ("Remitter", "U3 Approved Fee Gst"),
        ("Remitter", "U3 Approved NPCI Switching Fee"),
        ("Remitter", "U3 Approved NPCI Switching Fee Gst"),
        ("Remitter", "U3 Approved Payer PSP Fee"),
        ("Remitter", "U3 Approved Payer PSP Fee Gst"),
        ("Remitter", "U2 RB Approved Fee"),
        ("Remitter", "U2 RB Approved Fee Gst"),
        ("Remitter", "U2 RB Approved Surcharge Fee"),
        ("Remitter", "U2 RB Approved Surcharge Fee Gst"),
        ("Remitter", "U2 Approved Surcharge Fee"),
        ("Remitter", "U2 Approved Surcharge Fee Gst"),
        ("Remitter", "U3 Approved Surcharge Fee"),
        ("Remitter", "U3 Approved Surcharge Fee Gst"),
    ]

    beneficiary_conditions = [
        (None, "U2 Approved Fee"),
        (None, "U2 Approved Fee Gst"),
        (None, "U2 Approved NPCI Switching Fee"),
        ("Beneficiary", "U2 Approved NPCI Switching Fee Gst"),
        ("Beneficiary", "U2 RB Approved NPCI Switching Fee"),
        ("Beneficiary", "U2 RB Approved NPCI Switching Fee Gst"),
        ("Beneficiary", "U3 RB Approved NPCI Switching Fee"),
        ("Beneficiary", "U3 RB Approved NPCI Switching Fee Gst"),
        (None, "U2 RB Approved Payer PSP Fee"),
        ("Beneficiary", "U2 RB Approved Payer PSP Fee Gst"),
        ("Beneficiary", "U3 RB Approved Payer PSP Fee"),
        ("Beneficiary", "U3 RB Approved Payer PSP Fee Gst"),
        (None, "U2 Approved Payer PSP Fee"),
        (None, "U2 Approved Payer PSP Fee Gst"),
        (None, "U3 RB Approved Fee"),
        (None, "U3 RB Approved Fee Gst"),
        (None, "U3 Approved Fee"),
        ("Beneficiary", "U3 Approved Fee Gst"),
        ("Beneficiary", "U3 Approved NPCI Switching Fee"),
        ("Beneficiary", "U3 Approved NPCI Switching Fee Gst"),
        ("Beneficiary", "U3 Approved Payer PSP Fee"),
        ("Beneficiary", "U3 Approved Payer PSP Fee Gst"),
        (None, "U2 RB Approved Fee"),
        (None, "U2 RB Approved Fee Gst"),
        ("Beneficiary", "U2 RB Approved Surcharge Fee"),
        ("Beneficiary", "U2 RB Approved Surcharge Fee GST"),
        ("Beneficiary", "U2 Approved Surcharge Fee"),
        ("Beneficiary", "U2 Approved Surcharge Fee Gst"),
        ("Beneficiary", "U3 Approved Surcharge Fee"),
        ("Beneficiary", "U3 Approved Surcharge Fee Gst"),
    ]

    # Load all sheets into a dictionary of DataFrames
    sheets = pd.read_excel(file_path, sheet_name=None)

    # Process the conditions for remitter and beneficiary
    remitter_data = process_conditions(sheets, remitter_conditions, "Debit")
    beneficiary_data = process_conditions(sheets, beneficiary_conditions, "Credit")

    # Aggregate all cycles for remitter and beneficiary
    remitter_aggregated = aggregate_all_cycles(sheets, "Debit")
    beneficiary_aggregated = aggregate_all_cycles(sheets, "Credit")

    # Add "Beneficiary / Remitter Sub Totals" to the aggregated data
    remitter_sub_totals = aggregate_sub_totals(sheets, "Debit")
    beneficiary_sub_totals = aggregate_sub_totals(sheets, "Credit")

    remitter_aggregated = pd.concat([remitter_aggregated, remitter_sub_totals], ignore_index=True)
    beneficiary_aggregated = pd.concat([beneficiary_aggregated, beneficiary_sub_totals], ignore_index=True)

    # Write both results to the same Excel file with appropriate gaps
    with pd.ExcelWriter(output_file_path, engine="openpyxl") as writer:
        # Add "Remitter" heading and data
        pd.DataFrame(["Remitter"]).to_excel(writer, sheet_name="Combined Data", index=False, header=False, startrow=0)
        remitter_data.to_excel(writer, sheet_name="Combined Data", index=False, startrow=1)

        # Add a gap and "Beneficiary" heading
        gap_row = len(remitter_data) + 3
        pd.DataFrame(["Beneficiary"]).to_excel(writer, sheet_name="Combined Data", index=False, header=False, startrow=gap_row)
        beneficiary_data.to_excel(writer, sheet_name="Combined Data", index=False, startrow=gap_row + 1)

        # Add a gap of 5 rows and write aggregated data
        aggregated_start_row = gap_row + len(beneficiary_data) + 6
        pd.DataFrame(["Remitter Aggregated Data"]).to_excel(writer, sheet_name="Combined Data", index=False, header=False, startrow=aggregated_start_row)
        remitter_aggregated.to_excel(writer, sheet_name="Combined Data", index=False, startrow=aggregated_start_row + 1)

        aggregated_beneficiary_start_row = aggregated_start_row + len(remitter_aggregated) + 5
        pd.DataFrame(["Beneficiary Aggregated Data"]).to_excel(writer, sheet_name="Combined Data", index=False, header=False, startrow=aggregated_beneficiary_start_row)
        beneficiary_aggregated.to_excel(writer, sheet_name="Combined Data", index=False, startrow=aggregated_beneficiary_start_row + 1)

    st.success(f"Combined output saved to: {output_file_path}")

def process_conditions(sheets, conditions, data_type):
    final_data = []

    for sheet_name, df in sheets.items():
        # Initialize a dictionary to store aggregated results for the current sheet
        data = {"Cycle": [sheet_name]}

        # Process each condition
        for start_condition, end_condition in conditions:
            # Initialize the variables for each condition
            total_txn = 0
            total_value = 0

            # Filter rows based on the condition
            if start_condition is not None:
                filtered_df = df[
                    df['Description'].notna() &  # Exclude NaN values
                    df['Description'].str.startswith(start_condition) &
                    df['Description'].str.endswith(end_condition)
                ]
            else:
                filtered_df = df[
                    df['Description'].notna() &  # Exclude NaN values
                    df['Description'].str.endswith(end_condition)
                ]

            # Aggregate the data
            total_txn += filtered_df["No of Txns"].sum()
            total_value += filtered_df[data_type].sum()

            # Add aggregated data to the dictionary
            data[f"{end_condition} No of Txns"] = [total_txn]
            data[f"{end_condition} {data_type}"] = [total_value]

        # Append the results for the current sheet to the final data
        final_data.append(pd.DataFrame(data))

    # Combine all sheet data into a single DataFrame
    final_df = pd.concat(final_data, ignore_index=True)

    # Add a total row at the end
    total_row = {"Cycle": "Total"}
    for column in final_df.columns:
        if column not in ["Cycle"]:  # Sum only numeric columns
            total_row[column] = final_df[column].sum()

    # Append the total row to the DataFrame
    final_df = pd.concat([final_df, pd.DataFrame([total_row])], ignore_index=True)

    # Remove columns ending with "No of Txns"
    final_df = final_df.loc[:, ~final_df.columns.str.endswith("No of Txns")]

    return final_df

def aggregate_sub_totals(sheets, data_type):
    total_txn = 0
    total_value = 0

    for sheet_name, df in sheets.items():
        filtered_df = df[
            df['Description'].notna() &
            df['Description'].str.contains("Beneficiary / Remitter Sub Totals")
        ]

        total_txn += filtered_df["No of Txns"].sum()
        total_value += filtered_df[data_type].sum()

    return pd.DataFrame({
        "Description": ["Beneficiary / Remitter Sub Totals"],
        "Total Txns": [total_txn],
        f"Total {data_type}": [total_value]
    })

def aggregate_all_cycles(sheets, data_type):
    aggregated_results = {}

    conditions = [
       "U2 Approved Fee",
    "U2 Approved Fee Gst",
    "U2 Approved NPCI Switching Fee",
    "U2 Approved NPCI Switching Fee Gst",
    "U2 RB Approved NPCI Switching Fee",
     "U2 RB Approved NPCI Switching Fee Gst",
     "U3 RB Approved NPCI Switching Fee",
     "U3 RB Approved NPCI Switching Fee Gst",
     "U2 RB Approved Payer PSP Fee",
     "U2 RB Approved Payer PSP Fee Gst",
     "U3 RB Approved Payer PSP Fee",
    "U3 RB Approved Payer PSP Fee Gst",
   "U2 Approved Payer PSP Fee",
     "U2 Approved Payer PSP Fee Gst",
     "U3 RB Approved Fee",
     "U3 RB Approved Fee Gst",
     "U3 Approved Fee",
     "U3 Approved Fee Gst",
     "U3 Approved NPCI Switching Fee",
   "U3 Approved NPCI Switching Fee Gst",
    "U3 Approved Payer PSP Fee",
     "U3 Approved Payer PSP Fee Gst",
   "U2 RB Approved Fee",
    "U2 RB Approved Fee Gst",
       "U2 RB Approved Surcharge Fee",
    "U2 Approved Surcharge Fee",
    "U2 Approved Surcharge Fee Gst",
     "U3 Approved Surcharge Fee",
     "U3 Approved Surcharge Fee Gst"
    ]

    for condition in conditions:
        total_txn = 0
        total_value = 0

        for sheet_name, df in sheets.items():
            filtered_df = df[
                df['Description'].notna() &
                df['Description'].str.endswith(condition)
            ]

            total_txn += filtered_df["No of Txns"].sum()
            total_value += filtered_df[data_type].sum()

        # Store the aggregated results
        aggregated_results[condition] = {"Total Txns": total_txn, f"Total {data_type}": total_value}

    # Convert results to a DataFrame
    aggregated_df = pd.DataFrame(aggregated_results).T.reset_index()
    aggregated_df.columns = ["Description", "Total Txns", f"Total {data_type}"]

    return aggregated_df

if __name__ == "__main__":
    main()

