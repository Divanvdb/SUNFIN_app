import streamlit as st

import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

class ExcelProcessor:
    def __init__(self, bca_path = '', assets_path = '', po_details_path = '', acc_analyses_path = '', po_target = 1):
        self.bca_path = bca_path
        self.assets_path = assets_path
        self.po_details_path = po_details_path
        self.acc_analyses_path = acc_analyses_path
        self.df_bca = pd.DataFrame()
        self.df_assets = pd.DataFrame()
        self.df_po = pd.DataFrame()
        self.df_acc = pd.DataFrame()
        self.df_sorted = pd.DataFrame()

    def extract_data_from_excel(self, file_path, target_value='Budget Account', file_type='bca', verbose=False):
        workbook = load_workbook(filename=file_path)
        sheet = workbook.worksheets[0]
        num_rows = sheet.max_row

        target_col_idx = 1 

        found_target = False
        target_row_idx = -1

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
            if file_type == 'bca':
                print(f"'{target_value}' not found in the first column.")

        if file_type == 'bca':
            print('Reading BCA data\n')
            df['Transaction Description'] = ''
            df['Commitment Nr'] = ''
            df['Obligation Nr'] = ''
            df['Expenditure Nr'] = ''
            df['Requester Name'] = ''
            df['Supplier Name'] = ''
            df['Item Description'] = ''
            df['Item Category Description'] = ''
            print(f'Shape of DataFrame:', df.shape)

            df['Transaction Amount'] = df['Transaction Amount'].astype(str)
            df['Transaction Amount'] = df['Transaction Amount'].str.replace(',', '')
            df['Transaction Amount'] = df['Transaction Amount'].astype(float)

            df['Transaction Number'] = df['Transaction Number'].astype(str)

            df_obligations = df[df['Balance Type'] == 'Obligation']
            df_commitments = df[df['Balance Type'] == 'Commitment']
            df_budget = df[df['Balance Type'] == 'Budget']
            df_expenditures = df[df['Balance Type'] == 'Expenditure']

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
        elif file_type == 'acc':
            wb = load_workbook(file_path)
            ws = wb.active

            data = []
            for row in ws.iter_rows(values_only=True):
                data.append(row)

            # Convert to DataFrame if needed
            df = pd.DataFrame(data)

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

    def group_and_sort(self, df, numbers, threshold=10):
        for i in range(len(numbers)):
            trans_nrs = [numbers[i]]
            trans_nrs = df.loc[
                df['Liquidation Transaction Number'].isin(trans_nrs) | 
                df['Transaction Number'].isin(trans_nrs), 
                'Transaction Number'
            ].unique()
            trans_nrs = df.loc[
                df['Liquidation Transaction Number'].isin(trans_nrs) | 
                df['Transaction Number'].isin(trans_nrs), 
                'Transaction Number'
            ].unique()

            df.loc[df['Transaction Number'].isin(trans_nrs), 'Transaction Description'] = i

        for i in range(len(numbers)):
            df.loc[
                (df['Transaction Description'] == i) & 
                (df['Balance Type'] == 'Commitment'), 
                'Amount'
            ] = df.loc[
                (df['Transaction Description'] == i) & 
                (df['Balance Type'] == 'Commitment'), 
                'Transaction Amount'
            ].sum()

            df.loc[
                (df['Transaction Description'] == i) & 
                (df['Balance Type'] == 'Obligation'), 
                'Amount'
            ] = df.loc[
                (df['Transaction Description'] == i) & 
                (df['Balance Type'] == 'Obligation'), 
                'Transaction Amount'
            ].sum()

        df.loc[
            (df['Amount'] < threshold) & 
            (df['Amount'] > -threshold), 
            'Drop'
        ] = True

        balance_type_order = pd.CategoricalDtype(
            categories=["Commitment", "Obligation", "Expenditure"],
            ordered=True
        )

        df["Balance Type"] = df["Balance Type"].astype(balance_type_order)

        df = df.sort_values(by=["Transaction Description", "Balance Type"], ascending=[True, True])

        return df

    def concatenate_rows_by_po_number(self, df, group_by='Purchase Order Number'):
        return df.groupby(group_by).agg(lambda x: ' | '.join(x.astype(str))).reset_index()

    def process(self):
        self.df_bca = self.extract_data_from_excel(self.bca_path, verbose=False)
        self.df_bca = self.df_bca.loc[(self.df_bca['Balance Type'] != 'Budget')]
        if self.assets_path is not None:
            self.df_assets = self.extract_data_from_excel(self.assets_path, verbose=False)
            self.df_assets = self.df_assets.loc[(self.df_assets['Balance Type'] != 'Budget')]

            self.df_bca = pd.concat([self.df_bca, self.df_assets])

        comm_numbers = self.df_bca.loc[(self.df_bca['Transaction Amount'] > 0) & (self.df_bca['Balance Type'] == 'Commitment'), 'Transaction Number'].unique()

        self.df_bca = self.add_descriptions(self.df_bca, comm_numbers)
        self.df_reduced = self.df_bca.loc[(self.df_bca['Balance Type'] != 'Budget')]

        self.df_sorted = self.group_and_sort(self.df_reduced, comm_numbers)

        if self.po_details_path is not None:

            self.df_po = self.extract_data_from_excel(self.po_details_path, target_value='Procurement Business Unit Name', file_type='PO', verbose=True)
            self.df_po = self.df_po[['Order Number', 'Requester Name', 'Supplier Name', 'Item Description', 'Item Category Description']]
            self.df_po_concatenated = self.concatenate_rows_by_po_number(self.df_po, group_by='Order Number')

            po_numbers = self.df_po_concatenated['Order Number'].unique()

            for po_number in po_numbers:
                self.df_sorted.loc[(self.df_sorted['Obligation Nr'] == po_number) | (self.df_sorted['Transaction Number'] == po_number), 'Item Description'] = self.df_po_concatenated.loc[self.df_po_concatenated['Order Number'] == po_number, 'Item Description'].values[0]
                self.df_sorted.loc[(self.df_sorted['Obligation Nr'] == po_number) | (self.df_sorted['Transaction Number'] == po_number), 'Item Category Description'] = self.df_po_concatenated.loc[self.df_po_concatenated['Order Number'] == po_number, 'Item Category Description'].values[0]
                self.df_sorted.loc[(self.df_sorted['Obligation Nr'] == po_number) | (self.df_sorted['Transaction Number'] == po_number), 'Requester Name'] = self.df_po_concatenated.loc[self.df_po_concatenated['Order Number'] == po_number, 'Requester Name'].values[0]
                self.df_sorted.loc[(self.df_sorted['Obligation Nr'] == po_number) | (self.df_sorted['Transaction Number'] == po_number), 'Supplier Name'] = self.df_po_concatenated.loc[self.df_po_concatenated['Order Number'] == po_number, 'Supplier Name'].values[0]
        
        if self.acc_analyses_path is not None:
            self.df_acc_ = self.extract_data_from_excel(self.acc_analyses_path, file_type='acc', verbose=False)
            self.df_acc = pd.DataFrame()
            self.df_acc['Transaction Number'] = self.df_acc_[8]
            self.df_acc['Line Description'] = self.df_acc_[10]

            self.df_acc = self.df_acc.dropna() 
            self.df_acc = self.df_acc[self.df_acc['Transaction Number'] != 'Transaction Number']
            self.df_acc_concatenated = self.concatenate_rows_by_po_number(self.df_acc, group_by='Transaction Number')

            acc_numbers = self.df_acc_concatenated['Transaction Number'].unique()

            for acc_number in acc_numbers:
                self.df_sorted.loc[(self.df_sorted['Obligation Nr'] == acc_number) | (self.df_sorted['Transaction Number'] == acc_number), 'Item Description'] = self.df_acc_concatenated.loc[self.df_acc_concatenated['Transaction Number'] == acc_number, 'Line Description'].values[0]

        
        workbook = load_workbook(filename=self.bca_path)

        # Add a new sheet
        new_sheet = workbook.create_sheet(title='Processed')

        # Optionally, write data to the new sheet
        for r_idx, row in enumerate(dataframe_to_rows(self.df_sorted, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                new_sheet.cell(row=r_idx, column=c_idx, value=value)

        # Save the workbook to a variable (use BytesIO to store in-memory)
        from io import BytesIO

        # Save workbook to a BytesIO object
        excel_stream = BytesIO()
        workbook.save(excel_stream)

        # Rewind the stream to the beginning
        excel_stream.seek(0)

        return excel_stream
       

# Streamlit App
st.title('Excel Data Processor')

st.sidebar.header('Upload Files')
bca_file = st.sidebar.file_uploader("Upload BCA File", type=["xlsx"])
assets_file = st.sidebar.file_uploader("Upload Assets File", type=["xlsx"])
po_file = st.sidebar.file_uploader("Upload PO Details File", type=["xlsx"])
acc_file = st.sidebar.file_uploader("Upload ACC Analyses File", type=["xlsx"])

if st.sidebar.button('Process'):
    if bca_file:
        st.write("Files uploaded successfully. Processing will start...")
        processor = ExcelProcessor(bca_file, assets_file, po_file, acc_file)
        output_file = processor.process()
        st.write("Processing complete!")
        st.markdown(f"[Download the output file]")
        st.download_button(
            label="Download Updated Excel",
            data=output_file,
            file_name="Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Please upload all required files.")