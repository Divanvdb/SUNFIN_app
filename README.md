# 🎈 SUNFIN APP

Application to format Income Statements

Please find the online application here: [Application](https://blank-app-awq6au0oktk.streamlit.app/)

**_Dedication:_**  
The fact that this app is required highlights the importance of first engaging with customers towards clearly defining their basic requirements, before designing a system for them to use. May this understanding spread more widely...

**_Disclaimer:_**  
Use the app at your own risk, and please don’t blame us if it does not work or gives the wrong information.  
You are welcome to improve it by contacting divanvdb@sun.ac.za 

_Note:_ The 'Guide to Making Sense of SunFin.pdf' instructions will explain how the input files are generated.

### How to run it on your own machine

1. Install the requirements

   ```
   $ pip install -r requirements.txt
   ```

2. Run the app

   ```
   $ streamlit run streamlit_app.py
   ```

### App updates 
**_Updates to V1.1:_**  
- Removed Account Analyses  
- Added functionality to extract_data_from_excel function  
- Dropped NaN columns  
- Ordered the columns differently with formatting  
- Checking if assets and PO are None  
- Fixed the PO Number heading requirements  
- Rename Transaction Description to Cluster  
- Added a balances sheet

**_Updates to V1.2:_**  
- Bug fixes regarding BCA files with no possitive commitments
- No balances sheet if there is no assets file
- Added Obligation grouping round

**_Updates to V1.3:_**
- Format changes
- Added Account_number_changes.xlsx

**_Updates to V1.6:_**
- Added multi file processing
- Error catching for files during automated process.

