import streamlit as st

import pandas as pd
import numpy as np

from io import BytesIO

import zipfile
import tempfile
import os

from openpyxl import load_workbook, Workbook

import copy

from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

class ExcelProcessor:
    def __init__(self, bca_path = '', assets_path = '', po_details_path = ''):
        self.bca_path = bca_path
        self.assets_path = assets_path
        self.po_details_path = po_details_path
        self.df_bca = pd.DataFrame()
        self.df_assets = pd.DataFrame()
        self.df_po = pd.DataFrame()
        self.df_sorted = pd.DataFrame()
        self.balances = False
        self.required_columns = ['Transaction Number', 'Balance Type', 'Transaction Amount', 
                    'Liquidation Transaction Number', 'Cluster', 
                    'Commitment Nr', 'Obligation Nr', 'Expenditure Nr']
        
        self.sheet_names = ['BCA']
        self.file_paths = []

        self.CFB_flag = False

    def extract_data_from_excel(self, file_path, file_type='bca', verbose=False):
        print('Extracting')
        workbook = load_workbook(filename=file_path)

        found_target = False        

        # Set target value based on file type
        target_value = ''
        if file_type == 'bca':
            target_value='Budget Account'
        elif file_type == 'PO':
            target_value='Procurement Business Unit'

        # Look to find the target value in the first column of the first sheet        
        print('Loop')
        loop_count = 0
        while not found_target and loop_count < 5:
            loop_count += 1
            sheet = workbook.worksheets[0]
            num_rows = sheet.max_row
            target_row_idx = -1
            target_col_idx = 1 

            for row_idx in range(1, num_rows + 1): 
                cell_value = sheet.cell(row=row_idx, column=target_col_idx).value
                if cell_value == target_value:
                    found_target = True
                    target_row_idx = row_idx
                    break

            if found_target:
                headers = [cell.value for cell in sheet[target_row_idx]]

                start_row_idx = target_row_idx + 1
                data = []
                for row in sheet.iter_rows(min_row=start_row_idx, max_row=num_rows, values_only=True):
                    data.append(row)

                df = pd.DataFrame(data, columns=headers)

            else:
                # print(f"'{target_value}' not found in the first column.")
                if file_type == 'PO':
                    target_value='Procurement Business Unit Name'
        
        print('Loop Done')
        # Check if the DataFrame is empty
        if found_target:
            all_none = df.isnull().all().all()
        else:
            all_none = True
            st.warning(f"**_Warning:_** No data found in {file_path.name}.")

        if all_none:
            return None
        else:
            if file_type == 'bca':
                # print('Reading BCA data\n')
                df['Cluster'] = ''
                df['Commitment Nr'] = ''
                df['Obligation Nr'] = ''
                df['Expenditure Nr'] = ''
                df['Requester Name'] = ''
                df['Supplier Name'] = ''
                df['Item Description'] = ''
                df['Project Code'] = ''
                df['Item Category Description'] = ''
                # print(f'Shape of DataFrame:', df.shape)
                
                try:
                    df['Transaction Amount'] = df['Transaction Amount'].astype(str)
                    df['Transaction Amount'] = df['Transaction Amount'].str.replace(',', '')
                    df['Transaction Amount'] = df['Transaction Amount'].astype(float)

                    df['Transaction Number'] = df['Transaction Number'].astype(str)
                except:
                    print('Error converting Transaction Amount to float')

                df_obligations = df[df['Balance Type'] == 'Obligation']
                df_commitments = df[df['Balance Type'] == 'Commitment']
                df_budget = df[df['Balance Type'] == 'Budget']
                df_expenditures = df[df['Balance Type'] == 'Expenditure']

                # Check if liquidation transaction number is in the transaction number column
                self.missing_columns = [col for col in self.required_columns if col not in df.columns]

                if not self.missing_columns:
                    # st.markdown("All required columns are present in the dataframe.")
                    pass
                else:
                    st.markdown("The following columns are missing:", self.missing_columns)
                

                if verbose:

                    print(f'Before operations:\n\n')
                    print(f'Obligations: {df_obligations.shape}')
                    print(f'Commitments: {df_commitments.shape}')
                    print(f'Budget: {df_budget.shape}')
                    print(f'Expenditures: {df_expenditures.shape}')

                    print(f'Total: {df_obligations.shape[0] + df_commitments.shape[0] + df_budget.shape[0] + df_expenditures.shape[0]}\n')

                    print('The sum of all the transactions in the different categories are:')
                    print(f"Obligation Sum: {df_obligations['Transaction Amount'].sum()}\nCommitment Sum: {df_commitments['Transaction Amount'].sum()}\nBudget Sum: {df_budget['Transaction Amount'].sum()}\nExpenditure Sum: {df_expenditures['Transaction Amount'].sum()}")

                return df
            else:
                return df

    def add_descriptions(self, df, numbers):
        for i in range(len(numbers)):
            comm_nr = [numbers[i]]
            obl_nr = df.loc[
                (df['Liquidation Transaction Number'] == comm_nr[0]),
                'Transaction Number'
            ].unique()

            if len(obl_nr) == 0:
                df.loc[
                    (df['Transaction Number'] == comm_nr[0]),
                    'Commitment Nr'
                ] = comm_nr[0]
            else:
                for obl_number in obl_nr:
                    exp_nr = df.loc[
                        (df['Liquidation Transaction Number'] == obl_number) & 
                        (df['Balance Type'] != 'Commitment'),
                        'Transaction Number'
                    ].unique()
                    
                    if len(exp_nr) == 0:
                        df.loc[
                            (df['Transaction Number'] == obl_number),
                            'Commitment Nr'
                        ] = comm_nr[0]

                        df.loc[
                            (df['Transaction Number'] == obl_number),
                            'Obligation Nr'
                        ] = obl_number
                    else:
                        for exp_number in exp_nr:
                            transaction_number = np.concatenate((comm_nr, [obl_number], [exp_number]))

                            df.loc[
                                (df['Transaction Number'].isin(transaction_number)) & 
                                (df['Balance Type'] == 'Expenditure'),
                                'Commitment Nr'
                            ] = comm_nr[0]

                            df.loc[
                                (df['Transaction Number'].isin(transaction_number)) & 
                                (df['Balance Type'] == 'Expenditure'),
                                'Obligation Nr'
                            ] = obl_number

                            df.loc[
                                (df['Transaction Number'].isin(transaction_number)) & 
                                (df['Balance Type'] == 'Expenditure'),
                                'Expenditure Nr'
                            ] = exp_number

        return df

    def group_and_sort(self, df, numbers, threshold=10, offset=0):

        # Add cluster numbers
        for i in range(len(numbers)):
            trans_nrs = [numbers[i]]
            trans_nrs = df.loc[df['Liquidation Transaction Number'].isin(trans_nrs) | 
                df['Transaction Number'].isin(trans_nrs), 'Transaction Number'].unique()
            
            trans_nrs = df.loc[df['Liquidation Transaction Number'].isin(trans_nrs) | 
                df['Transaction Number'].isin(trans_nrs),'Transaction Number'].unique()

            df.loc[df['Transaction Number'].isin(trans_nrs), 'Cluster'] = offset + i

        # Create the transaction amounts based on cluster
        for i in range(len(numbers)):
            df.loc[
                (df['Cluster'] == i) & 
                (df['Balance Type'] == 'Commitment'), 
                'Amount'
            ] = df.loc[
                (df['Cluster'] == i) & 
                (df['Balance Type'] == 'Commitment'), 
                'Transaction Amount'
            ].sum()

            df.loc[
                (df['Cluster'] == i) & 
                (df['Balance Type'] == 'Obligation'), 
                'Amount'
            ] = df.loc[
                (df['Cluster'] == i) & 
                (df['Balance Type'] == 'Obligation'), 
                'Transaction Amount'
            ].sum()

        # Assign project codes based on amount vs threshold
        df.loc[
            (df['Amount'] < threshold) & 
            (df['Amount'] > -threshold), 
            'Project Code'
        ] = 'Ignore'

        return df

    def concatenate_rows_by_po_number(self, df, group_by='Purchase Order Number'):
        return df.groupby(group_by).agg(lambda x: ' | '.join(x.astype(str))).reset_index()  

    def create_file_name(self, filename):        
        formatted_filename = " - ".join(filename.split(" - ")[:2]) + " - Output.xlsx"

        return formatted_filename
    
    def create_output_file(self, df, file_paths):
        new_workbook = Workbook()
        new_workbook.remove(new_workbook.active)  

        for i, path in enumerate(file_paths):
            workbook = load_workbook(path)
            
            sheet = workbook.active
            
            new_sheet = new_workbook.create_sheet(title=self.sheet_names[i])
            
            for row in sheet:
                for cell in row:
                    new_sheet[cell.coordinate].value = cell.value
                    new_sheet[cell.coordinate].font = copy.copy(cell.font)
                    new_sheet[cell.coordinate].border = copy.copy(cell.border)
                    new_sheet[cell.coordinate].fill = copy.copy(cell.fill)
                    new_sheet[cell.coordinate].number_format = copy.copy(cell.number_format)
                    new_sheet[cell.coordinate].protection = copy.copy(cell.protection)
                    new_sheet[cell.coordinate].alignment = copy.copy(cell.alignment)

            for merged_range in sheet.merged_cells.ranges:
                new_sheet.merge_cells(str(merged_range))

        progress_placeholder.markdown(f"Current processing: {70}% complete...")

        ##### Add new data

        new_sheet = new_workbook.create_sheet(title='Processed')

        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                new_cell = new_sheet.cell(row=r_idx, column=c_idx, value=value)
                if r_idx == 1:
                    new_cell.font = Font(bold=True)
                
                if c_idx == 24: 
                    new_cell.number_format = 'R#,##0.00'
        
        # Set column width based on content
        for col in new_sheet.columns:
            max_length = 0
            column = col[0].column_letter  
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)  # Adding a little extra space
            new_sheet.column_dimensions[column].width = adjusted_width

        # Set column width of first 16 columns
        for col_idx in range(1, 17): 
            col_letter = get_column_letter(col_idx)
            new_sheet.column_dimensions[col_letter].width = 12.75 

        # # Ignore rows with project code 'Ignore'
        # project_code_col = 21
        # for row in new_sheet.iter_rows(min_row=2, max_col=project_code_col, max_row=new_sheet.max_row):
        #     project_code_value = row[project_code_col - 1].value
        #     if project_code_value == "Ignore":
        #         new_sheet.row_dimensions[row[0].row].hidden = True

        new_sheet.auto_filter.ref = new_sheet.dimensions
        # new_sheet.auto_filter.add_filter_column(20, ['Ignore'], blank=False) 

        new_sheet.freeze_panes = 'A2'

        progress_placeholder.markdown(f"Current processing: {80}% complete...")

        # Add balances sheet

        df_balances = self.get_balances()

        new_sheet = new_workbook.create_sheet(title='Balances')

        for r_idx, row in enumerate(dataframe_to_rows(df_balances, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                new_cell = new_sheet.cell(row=r_idx, column=c_idx, value=value)
                if r_idx == 1:
                    new_cell.font = Font(bold=True)

                if c_idx == 2:  
                    new_cell.number_format = 'R#,##0.00'
        
        for col in new_sheet.columns:
            max_length = 0
            column = col[0].column_letter  # Get the column name
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)  # Adding a little extra space
            new_sheet.column_dimensions[column].width = adjusted_width

        progress_placeholder.markdown(f"Current processing: {90}% complete...")

        # Save workbook to BytesIO stream

        excel_stream = BytesIO()
        new_workbook.save(excel_stream)

        excel_stream.seek(0)

        return excel_stream

    def get_balances(self):
        data = {
            'Description': [
                'Opening balance:',
                'Closing balance:',
                'Commitments during period:',
                'Obligations during period:',
                'Expenses during period:',
                'Total consumption during period:',
                'Total income during period:'
            ],
            'BCA': [None] * 7 ,
            'BCA Assets': [None] * 7,
            'Total': [None] * 7}
        df = pd.DataFrame(data)

        if self.CFB_flag:
            formulas_bca = [
                "=Processed!X2",
                "=BCA!K4",
                "=BCA!AA23",
                "=BCA!AD23",
                "=BCA!AI23",
                "=BCA!G4",
                "=B3-(B2-B7)"
            ]
        else:
            formulas_bca = [
                "=BCA!D4",
                "=BCA!K4",
                "=BCA!AA23",
                "=BCA!AD23",
                "=BCA!AI23",
                "=BCA!G4",
                "=B3-(B2-B7)"
            ]

        if self.balances:
            formulas_assets = [
                "==IF(ISNUMBER('BCA Assets'!D4),'BCA Assets'!D4,0)",
                "==IF(ISNUMBER('BCA Assets'!K4),'BCA Assets'!K4,0)",
                "==IF(ISNUMBER('BCA Assets'!AA23),'BCA Assets'!AA23,0)",
                "==IF(ISNUMBER('BCA Assets'!AD23),'BCA Assets'!AD23,0)",
                "==IF(ISNUMBER('BCA Assets'!AI23),'BCA Assets'!AI23,0)",
                "==IF(ISNUMBER('BCA Assets'!G4),'BCA Assets'!G4,0)",
                "=C3-(C2-C7)"
            ]

            formulas_total = [
                "=B2+C2",
                "=B3+C3",
                "=B4+C4",
                "=B5+C5",
                "=B6+C6",
                "=B7+C7",
                "=B8+C8"
            ]

        # Update the DataFrame with the formulas
        for i, formula in enumerate(formulas_bca):
            df.at[i, 'BCA'] = formula

        if self.balances:
            for i, formula in enumerate(formulas_assets):
                df.at[i, 'BCA Assets'] = formula

            for i, formula in enumerate(formulas_total):
                df.at[i, 'Total'] = formula

        return df

    def extract_and_check_data(self):
        # Extract data from the BCA file
        if self.bca_path is not None:
            self.df_bca = self.extract_data_from_excel(self.bca_path, verbose=False)
        else:
            self.df_bca = None

        # Extract data from the Assets file if provided
        if self.assets_path is not None:
            self.df_assets = self.extract_data_from_excel(self.assets_path, verbose=False)
        else:
            self.df_assets = None

        # Extract data from the PO Details file if provided
        if self.po_details_path is not None:
            self.df_po = self.extract_data_from_excel(self.po_details_path, file_type='PO', verbose=False)
        else:
            self.df_po = None

        # Check if the extracted data is None
        bca_exists = self.df_bca is not None
        assets_exists = self.df_assets is not None
        po_exists = self.df_po is not None

        # Provide feedback based on the existence of the data
        if not bca_exists:
            st.warning('BCA file is empty. Please ensure that the SUNFIN Sheet has transactions.')
        if (not assets_exists) & (self.assets_path is not None):
            st.warning('BCA Assets file is empty.')
        if (self.po_details_path is not None) & (not po_exists):
            st.warning('PO Details file is empty.')

        # Return a tuple of True/False for each file
        return bca_exists, assets_exists, po_exists

    # 10% Completion
    def process_bca(self):
        self.file_paths.append(self.bca_path)

        # Debugging line
        # self.df_bca = self.df_bca.loc[(self.df_bca['Balance Type'] != 'Budget'), ['Transaction Number', 'Balance Type', 'Transaction Amount', 'Liquidation Transaction Number', 
        #                                                                          'Cluster', 'Commitment Nr', 'Obligation Nr', 'Expenditure Nr']]
        
        # Actual line
        self.df_bca.loc[
            (self.df_bca['Balance Type'] == 'Budget') &
            self.df_bca['Transaction Number'].str.contains("Carry forward balance", na=False), 'Balance Type'] = 'Budget CFB'
        self.df_bca = self.df_bca.loc[(self.df_bca['Balance Type'] != 'Budget')]
        self.df_bca = self.df_bca.drop(columns=[col for col in self.df_bca.columns if col is None])

    # 20% Completion
    def process_assets(self):
        self.sheet_names.append('BCA Assets')
        self.file_paths.append(self.assets_path)

        self.df_assets = self.df_assets.loc[(self.df_assets['Balance Type'] != 'Budget')]
        self.df_assets = self.df_assets.drop(columns=[col for col in self.df_assets.columns if col is None])
        self.df_bca = pd.concat([self.df_bca, self.df_assets])

        self.balances = True

    # 30% Completion
    def process_transactions(self):
        comm_numbers = self.df_bca.loc[(self.df_bca['Transaction Amount'] > 0) & (self.df_bca['Balance Type'] == 'Commitment'), 'Transaction Number'].unique()

        self.df_bca = self.add_descriptions(self.df_bca, comm_numbers)
        self.df_reduced = self.df_bca.loc[(self.df_bca['Balance Type'] != 'Budget')]

        if comm_numbers.size == 0:
            print(f"**_Note:_** No new commitments have been ")
            self.df_sorted = self.df_reduced
        else:
            self.df_sorted = self.group_and_sort(self.df_reduced, comm_numbers)

        ob_numbers = self.df_sorted.loc[(self.df_sorted['Cluster'] == '') &
                                        (self.df_sorted['Balance Type'] == 'Obligation') &
                                        (self.df_sorted['Transaction Amount'] > 0), 'Transaction Number'].unique()
        
        if ob_numbers.size == 0:
            if comm_numbers.size == 0:
                print(f"**_Note:_** No possitive obligations found")
        else:
            self.df_sorted =self.group_and_sort(self.df_sorted, ob_numbers, offset=len(comm_numbers))

    # 40% Completion
    def process_po(self):
        self.sheet_names.append('PO Details')
        self.file_paths.append(self.po_details_path)

        columns = self.df_po.iloc[:, [1, 2, 6, 15, 16]].columns

        self.df_po = self.df_po[[columns[0], columns[1], columns[2], columns[3], columns[4]]]

        self.df_po_concatenated = self.concatenate_rows_by_po_number(self.df_po, group_by=columns[0])

        po_numbers = self.df_po_concatenated[columns[0]].unique()

        for po_number in po_numbers:
            self.df_sorted.loc[(self.df_sorted['Obligation Nr'] == po_number) | (self.df_sorted['Transaction Number'] == po_number), 'Item Description'] = self.df_po_concatenated.loc[self.df_po_concatenated[columns[0]] == po_number, columns[3]].values[0]
            self.df_sorted.loc[(self.df_sorted['Obligation Nr'] == po_number) | (self.df_sorted['Transaction Number'] == po_number), 'Item Category Description'] = self.df_po_concatenated.loc[self.df_po_concatenated[columns[0]] == po_number, columns[4]].values[0]
            self.df_sorted.loc[(self.df_sorted['Obligation Nr'] == po_number) | (self.df_sorted['Transaction Number'] == po_number), 'Requester Name'] = self.df_po_concatenated.loc[self.df_po_concatenated[columns[0]] == po_number, columns[1]].values[0]
            self.df_sorted.loc[(self.df_sorted['Obligation Nr'] == po_number) | (self.df_sorted['Transaction Number'] == po_number), 'Supplier Name'] = self.df_po_concatenated.loc[self.df_po_concatenated[columns[0]] == po_number, columns[2]].values[0] 

        progress_placeholder.markdown(f"Processing: {40}% complete...")

    # 50% Completion
    def process_accounts(self):
        acc_num_desc = pd.read_excel('Account.numbers.table.xlsx')
        unique_accounts = self.df_sorted['Transaction Account'].unique()
        for trans_account in unique_accounts:
            if trans_account is not None:
                try:
                    self.df_sorted.loc[self.df_sorted['Transaction Account'] == trans_account, 'Temp'] = acc_num_desc.loc[acc_num_desc['Acc No'] == int(trans_account.split('-')[2]), 'Account description'].item()
                except:
                    pass

        if 'Temp' in self.df_sorted.columns:
            self.df_sorted['Item Description'] =  self.df_sorted['Temp'] + ' | ' + self.df_sorted['Item Description'].astype(str) 
            self.df_sorted['Item Category Description'] =  self.df_sorted['Temp'] + ' | ' + self.df_sorted['Item Category Description'].astype(str) 
            self.df_sorted = self.df_sorted.drop(columns=['Temp'])

    # 60% Completion to 100% Completion
    def process_output(self):        
        column_order = ['Budget Account', 'Cost Center Segment Description', 'Account Description', 
                                   'Transaction Type', 'Transaction SubType', 'Transaction Action', 'Transaction Number', 
                                   'Expense Report Owner', 'Transaction Account', 'Transaction ID', 'Transaction Currency', 
                                   'Activity Type', 'Reservation Amount', 'Liquidation Transaction Type', 'Liquidation Transaction Number', 
                                   'Liquidation Amount', 'Commitment Nr', 'Obligation Nr', 'Expenditure Nr', 
                                   'Cluster', 'Project Code', 'Budget Date', 
                                   'Balance Type', 'Transaction Amount', 'Item Description', 
                                   'Requester Name', 'Supplier Name', 'Item Category Description']
        
        # Ensure all columns are in df_sorted

        for col in column_order:
            if col not in self.df_sorted.columns:
                self.df_sorted[col] = None
        
        existing_columns = [col for col in column_order if col in self.df_sorted.columns]
        
        self.df_sorted = self.df_sorted.reindex(columns=existing_columns)

        progress_placeholder.markdown(f"Current processing: {60}% complete...")

        # Order the transactions to make sense

        if "Budget CFB" in self.df_sorted['Balance Type'].unique():
            balance_type_order = pd.CategoricalDtype(
                categories=["Budget CFB","Commitment", "Obligation", "Expenditure"],
                ordered=True
            )
            self.CFB_flag = True
        else:
            balance_type_order = pd.CategoricalDtype(
                categories=["Commitment", "Obligation", "Expenditure"],
                ordered=True
            )

        self.df_sorted["Balance Type"] = self.df_sorted["Balance Type"].astype(balance_type_order)

        self.df_sorted = self.df_sorted.sort_values(by=["Balance Type"], ascending=[True])

        print(self.df_sorted.iloc[0])
        output_file = self.create_output_file(self.df_sorted, self.file_paths)
        progress_placeholder.markdown(f"Current processing: {100}% complete...")

        return output_file
        
    def auto_process(self):
        bca_, asts_, po_ = self.extract_and_check_data()

        print('1')
        if bca_:
            self.process_bca()

            progress_placeholder.markdown(f"Current processing: {10}% complete...")

            if asts_:
                self.process_assets()

            progress_placeholder.markdown(f"Current processing: {20}% complete...")

            self.process_transactions()

            progress_placeholder.markdown(f"Current processing: {30}% complete...")

            if po_:
                self.process_po()

            progress_placeholder.markdown(f"Current processing: {40}% complete...")

            self.process_accounts()

            progress_placeholder.markdown(f"Current processing: {50}% complete...")

            output_file = self.process_output()

            return output_file

        else:
            print('2')
            return None

# Streamlit App

## Title

st.title('Making Sense of SUNFIN')

st.markdown('---')

st.markdown('''Version: 1.6''')

st.markdown('''Use the app at your own risk, and please don’t blame us if it does not work or gives the wrong information. 
            You are welcome to improve it by accessing the source code here: [Github](https://github.com/Divanvdb/SUNFIN_app)
''')

## Download Guide to SUNFIN
with open('Guide_to_Making_Sense_of_SunFin.pdf', 'rb') as file:
    pdf_data = file.read()

st.download_button(
    label="Download User Guide",
    data=pdf_data,
    file_name='Guide_to_Making_Sense_of_SunFin.pdf',
    mime='application/pdf'
)

st.markdown('---')

## Sidebar with file upload options

progress_placeholder = st.empty()

st.sidebar.header('Upload Files from Folder')

bca_file, assets_file, po_file = None, None, None

uploaded_files = st.sidebar.file_uploader("Upload Files from Folder", type=["xlsx"], accept_multiple_files=True)

st.sidebar.header('Upload Files Individually')

bca_file_individual = st.sidebar.file_uploader("Upload BCA File", type=["xlsx"])
assets_file_individual = st.sidebar.file_uploader("Upload Assets File", type=["xlsx"])
po_file_individual = st.sidebar.file_uploader("Upload PO Details File", type=["xlsx"])

unique_ids = []
if uploaded_files:
    
    for uploaded_file in uploaded_files:
        extracted_value = uploaded_file.name.split(' - ')[0]
        
        if extracted_value not in unique_ids:
            if '.xlsx' not in extracted_value:
                unique_ids.append(extracted_value)
else:
    bca_file = bca_file_individual
    assets_file = assets_file_individual
    po_file = po_file_individual

## Processing function

if st.sidebar.button('Process'):

    if (bca_file is not None) | (len(unique_ids) >= 1):

        if len(unique_ids) >= 1:
            output_files = []
            output_names = []
            st.write("Multiple files uploaded successfully. Processing will start...")

            st.write(unique_ids)
            for i, unique_id in enumerate(unique_ids):
                bca_file, assets_file, po_file = None, None, None
                for uploaded_file in uploaded_files:
                    if f"{unique_id} -" in uploaded_file.name:
                        if "Assets" in uploaded_file.name or "asset" in uploaded_file.name or 'Assets' in uploaded_file.name:
                            assets_file = uploaded_file
                        elif "4 - BudgetaryControlAnalysis" in uploaded_file.name or "BCA" in uploaded_file.name:
                            bca_file = uploaded_file
                        elif "PO" in uploaded_file.name or "PODetails" in uploaded_file.name:
                            po_file = uploaded_file

                st.write(f"**{i + 1} / {len(unique_ids)}** - Files uploaded successfully for **{unique_id}**:")
                if bca_file is not None:
                    st.write(f"- BCA: {bca_file.name}")
                if assets_file is not None:
                    st.write(f"- BCA Assets: {assets_file.name}")
                if po_file is not None:
                    st.write(f"- PO File: {po_file.name}")

                if bca_file:
                    processor = ExcelProcessor(bca_file, assets_file, po_file)
                    progress_placeholder = st.empty()

                    try:
                        output = processor.auto_process()
                        if output is not None:
                            output_files.append(output)
                            
                            output_name = processor.create_file_name(bca_file.name)
                            output_names.append(output_name)

                            st.success(f"**Completed processing for** {unique_id}")
                        else:
                            st.warning(f"**No output for** {unique_id}")
                    except:
                        st.error(f"**Error processing** {unique_id}")
                else:
                    st.error(f"**No BCA file found for** {unique_id}")

            if len(output_files) == 0:
                st.error("No output files were generated.")

            else:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                    for output_file, output_name in zip(output_files, output_names):

                        if isinstance(output_file, BytesIO):
                            # Create a temporary file
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
                                temp_file.write(output_file.getvalue())  # Write the BytesIO content to the temp file
                                temp_file_path = temp_file.name  # Store the temporary file path

                            # Add the temporary file to the ZIP
                            zip_file.write(temp_file_path, arcname=output_name)

                            # Clean up: Delete the temporary file after adding to the ZIP
                            os.remove(temp_file_path)
                        else:
                            # If output_file is already a file path, just add it directly
                            zip_file.write(output_file, arcname=output_name)

                zip_buffer.seek(0)

                st.download_button(
                    label="Download All Updated Excel Files",
                    data=zip_buffer,
                    file_name="output_files.zip",
                    mime="application/zip"
                )

        else:
            if bca_file is not None:
                st.write(f"Files uploaded successfully:")
                st.write(f"- BCA: {bca_file.name}")

                if assets_file is not None:
                    st.write(f"- BCA Assets: {assets_file.name}")
                if po_file is not None:
                    st.write(f"- PO File: {po_file.name}")

                processor = ExcelProcessor(bca_file, assets_file, po_file)

                output_file = processor.auto_process()

                if output_file is not None:

                    st.write("**Completed processing.**")

                    st.download_button(
                                label="Download Updated Excel",
                                data=output_file,
                                file_name=processor.create_file_name(bca_file.name),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                else:
                    st.error("No output file was generated.")
            else:
                st.error("Please upload the BCA file.")

    else:
        st.error("Please upload all required files.")
