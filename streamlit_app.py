import streamlit as st

import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import load_workbook, Workbook
import copy
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

class ExcelProcessor:
    def __init__(self, bca_path = '', assets_path = '', po_details_path = ''):
        self.bca_path = bca_path
        self.assets_path = assets_path
        self.po_details_path = po_details_path
        self.df_bca = pd.DataFrame()
        self.df_assets = pd.DataFrame()
        self.df_po = pd.DataFrame()
        self.df_sorted = pd.DataFrame()
        self.po_order_number = 'Purchase Order Number'
        self.balances = True
        self.required_columns = ['Transaction Number', 'Balance Type', 'Transaction Amount', 
                    'Liquidation Transaction Number', 'Cluster', 
                    'Commitment Nr', 'Obligation Nr', 'Expenditure Nr']
        
        self.sheet_names = ['BCA']

    def extract_data_from_excel(self, file_path, file_type='bca', verbose=False):
        workbook = load_workbook(filename=file_path)

        found_target = False
        

        target_value = ''

        if file_type == 'bca':
            target_value='Budget Account'
        elif file_type == 'PO':
            target_value='Procurement Business Unit'
        
        while not found_target:
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
                print(f"'{target_value}' not found in the first column.")
                if file_type == 'PO':
                    target_value='Procurement Business Unit Name'
                    self.po_order_number = 'Order Number'

        if file_type == 'bca':
            print('Reading BCA data\n')
            df['Cluster'] = ''
            df['Commitment Nr'] = ''
            df['Obligation Nr'] = ''
            df['Expenditure Nr'] = ''
            df['Requester Name'] = ''
            df['Supplier Name'] = ''
            df['Item Description'] = ''
            df['Item Category Description'] = ''
            print(f'Shape of DataFrame:', df.shape)
            
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
                st.markdown("All required columns are present in the dataframe.")
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

        # Order the transactions to make sense
        balance_type_order = pd.CategoricalDtype(
            categories=["Commitment", "Obligation", "Expenditure"],
            ordered=True
        )

        df["Balance Type"] = df["Balance Type"].astype(balance_type_order)

        df = df.sort_values(by=["Cluster", "Balance Type"], ascending=[True, True])

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

        progress_placeholder.markdown(f"Processing: {70}% complete...")

        # Add new data

        new_sheet = new_workbook.create_sheet(title='Processed')

        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                new_cell = new_sheet.cell(row=r_idx, column=c_idx, value=value)
                if r_idx == 1:
                    new_cell.font = Font(bold=True)
                
                if c_idx == 24: 
                    new_cell.number_format = 'R#,##0.00'
        
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

        new_sheet.auto_filter.ref = new_sheet.dimensions

        new_sheet.freeze_panes = 'A2'

        progress_placeholder.markdown(f"Processing: {80}% complete...")

        # Add balances sheet
        if self.balances:

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

        progress_placeholder.markdown(f"Processing: {90}% complete...")

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
            'Value': [None] * 7 }
        df = pd.DataFrame(data)

        formulas = [
            "=BCA!D4+'BCA Assets'!D4",
            "=BCA!K4+'BCA Assets'!K4",
            "=BCA!AA23+'BCA Assets'!AA23",
            "=BCA!AD23+'BCA Assets'!AD23",
            "=BCA!AI23+'BCA Assets'!AI23",
            "=BCA!G4+'BCA Assets'!G4",
            "=B3-(B2-B7)"
        ]

        # Update the DataFrame with the formulas
        for i, formula in enumerate(formulas):
            df.at[i, 'Value'] = formula

        return df

    def process(self):
        # BCA setup 
        file_paths = []
        self.df_bca = self.extract_data_from_excel(self.bca_path, verbose=False)

        file_paths.append(self.bca_path)

        # Debugging line
        # self.df_bca = self.df_bca.loc[(self.df_bca['Balance Type'] != 'Budget'), ['Transaction Number', 'Balance Type', 'Transaction Amount', 'Liquidation Transaction Number', 
        #                                                                          'Cluster', 'Commitment Nr', 'Obligation Nr', 'Expenditure Nr']]
        
        # Actual line
        self.df_bca = self.df_bca.loc[(self.df_bca['Balance Type'] != 'Budget')]
        self.df_bca = self.df_bca.drop(columns=[col for col in self.df_bca.columns if col is None])

        progress_placeholder.markdown(f"Processing: {10}% complete...")

        # Assets setup
        if self.assets_path is not None:
            self.sheet_names.append('BCA Assets')
            file_paths.append(self.assets_path)
            self.df_assets = self.extract_data_from_excel(self.assets_path, verbose=False)

            self.df_assets = self.df_assets.loc[(self.df_assets['Balance Type'] != 'Budget')]
            self.df_assets = self.df_assets.drop(columns=[col for col in self.df_assets.columns if col is None])

            all_none = (self.df_assets['Transaction Amount'] == 'None').all()

            if all_none:
                st.markdown(f"**_Note:_** Assets are None")
            else:
                self.df_bca = pd.concat([self.df_bca, self.df_assets])
        else:
            self.balances = False

        # BCA and Assets processing

        progress_placeholder.markdown(f"Processing: {20}% complete...")

        comm_numbers = self.df_bca.loc[(self.df_bca['Transaction Amount'] > 0) & (self.df_bca['Balance Type'] == 'Commitment'), 'Transaction Number'].unique()

        self.df_bca = self.add_descriptions(self.df_bca, comm_numbers)
        self.df_reduced = self.df_bca.loc[(self.df_bca['Balance Type'] != 'Budget')]

        if comm_numbers.size == 0:
            st.markdown(f"**_Note:_** No possitive commitments found")
            self.df_sorted = self.df_reduced
        else:
            self.df_sorted = self.group_and_sort(self.df_reduced, comm_numbers)

        ob_numbers = self.df_sorted.loc[(self.df_sorted['Cluster'] == '') &
                                        (self.df_sorted['Balance Type'] == 'Obligation') &
                                        (self.df_sorted['Transaction Amount'] > 0), 'Transaction Number'].unique()
        
        if ob_numbers.size == 0:
            if comm_numbers.size == 0:
                st.markdown(f"**_Note:_** No possitive obligations found")
        else:
            self.df_sorted =self.group_and_sort(self.df_sorted, ob_numbers, offset=len(comm_numbers))

        progress_placeholder.markdown(f"Processing: {30}% complete...")

        # PO Details setup
        if self.po_details_path is not None:
            self.sheet_names.append('PO Details')
            file_paths.append(self.po_details_path)
            self.df_po = self.extract_data_from_excel(self.po_details_path, file_type='PO', verbose=True)

            progress_placeholder.markdown(f"Processing: {40}% complete...")

            all_columns_none = self.df_po.isna().all().all()

            if not all_columns_none:
                self.df_po = self.df_po[[self.po_order_number, 'Requester Name', 'Supplier Name', 'Item Description', 'Item Category Description']]
                self.df_po_concatenated = self.concatenate_rows_by_po_number(self.df_po, group_by=self.po_order_number)

                po_numbers = self.df_po_concatenated[self.po_order_number].unique()

                for po_number in po_numbers:
                    self.df_sorted.loc[(self.df_sorted['Obligation Nr'] == po_number) | (self.df_sorted['Transaction Number'] == po_number), 'Item Description'] = self.df_po_concatenated.loc[self.df_po_concatenated[self.po_order_number] == po_number, 'Item Description'].values[0]
                    self.df_sorted.loc[(self.df_sorted['Obligation Nr'] == po_number) | (self.df_sorted['Transaction Number'] == po_number), 'Item Category Description'] = self.df_po_concatenated.loc[self.df_po_concatenated[self.po_order_number] == po_number, 'Item Category Description'].values[0]
                    self.df_sorted.loc[(self.df_sorted['Obligation Nr'] == po_number) | (self.df_sorted['Transaction Number'] == po_number), 'Requester Name'] = self.df_po_concatenated.loc[self.df_po_concatenated[self.po_order_number] == po_number, 'Requester Name'].values[0]
                    self.df_sorted.loc[(self.df_sorted['Obligation Nr'] == po_number) | (self.df_sorted['Transaction Number'] == po_number), 'Supplier Name'] = self.df_po_concatenated.loc[self.df_po_concatenated[self.po_order_number] == po_number, 'Supplier Name'].values[0] 
            else:
                st.markdown(f"**_Note:_** PO Details are None")

        progress_placeholder.markdown(f"Processing: {50}% complete...")

        column_order = ['Budget Account', 'Cost Center Segment Description', 'Account Description', 
                                   'Transaction Type', 'Transaction SubType', 'Transaction Action', 'Transaction Number', 
                                   'Expense Report Owner', 'Transaction Account', 'Transaction ID', 'Transaction Currency', 
                                   'Activity Type', 'Reservation Amount', 'Liquidation Transaction Type', 'Liquidation Transaction Number', 
                                   'Liquidation Amount', 'Commitment Nr', 'Obligation Nr', 'Expenditure Nr', 
                                   'Cluster', 'Project Code', 'Budget Date', 
                                   'Balance Type', 'Transaction Amount', 'Item Description', 
                                   'Requester Name', 'Supplier Name', 'Item Category Description']
        
        existing_columns = [col for col in column_order if col in self.df_sorted.columns]

        
        self.df_sorted = self.df_sorted.reindex(columns=existing_columns)

        progress_placeholder.markdown(f"Processing: {60}% complete...")
        output_file = self.create_output_file(self.df_sorted, file_paths)
        progress_placeholder.markdown(f"Processing: {100}% complete...")

        return output_file
        


# Streamlit App

st.title('Making Sense of SUNFIN')

st.markdown('---')

st.markdown('''Version: 1.3''')

st.markdown('''This app is dedicated to all the engineers out there that understand the importance of first engaging with customers towards clearly defining their basic requirements, 
before designing a system for them to use. May this understanding spread widely and make apps like this one unnecessary...  
''')

st.markdown('''Use the app at your own risk, and please don’t blame us if it does not work or gives the wrong information. 
            You are welcome to improve it by accessing the source code here: [Github](https://github.com/Divanvdb/SUNFIN_app)
''')

st.markdown('---')

st.markdown('This app processes BCA, Assets and PO Details files and returns an updated Excel file.')

st.markdown('''**_Updates to V1.3:_**   
- Formatting based on Change Request #1

''')

st.markdown('''**_TODO:_**  

- Add threshold for transasction cancelling 
''')

st.markdown('---')

progress_placeholder = st.empty()

st.sidebar.header('Upload Files')
bca_file = st.sidebar.file_uploader("Upload BCA File", type=["xlsx"])
assets_file = st.sidebar.file_uploader("Upload Assets File", type=["xlsx"])
po_file = st.sidebar.file_uploader("Upload PO Details File", type=["xlsx"])

if st.sidebar.button('Process'):
    if bca_file:
        st.write("Files uploaded successfully. Processing will start...")
        processor = ExcelProcessor(bca_file, assets_file, po_file)
        output_file = processor.process()
        st.write("Processing complete!")
        # st.markdown(f"[Download the output file]")
        st.download_button(
            label="Download Updated Excel",
            data=output_file,
            file_name=processor.create_file_name(bca_file.name),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Please upload all required files.")
