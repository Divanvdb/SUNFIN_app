# 🎈 SUNFIN APP

Application to format Income Statements

**_Dedication:_**  
This app is dedicated to all the engineers out there that understand the importance of first engaging with customers towards clearly defining their basic requirements, 
before designing a system for them to use. May this understanding spread widely and make apps like this one unnecessary...  

**_Disclaimer:_**  
Use the app at your own risk, and please don’t blame us if it does not work or gives the wrong information.  
You are welcome to improve it by contacting divanvdb@sun.ac.za 

### How to run it on your own machine

1. Install the requirements

   ```
   $ pip install -r requirements.txt
   ```

2. Run the app

   ```
   $ streamlit run streamlit_app.py
   ```

3. App updates 
- Updates:

   - Removed acc_analyses_path
   - Added while loop into extract_data_from_excel function
   - Dropped nan columns
   - Orderded the columns differently with formatting
   - Checking if assets and PO are None
   - Fixed the PO Number heading requirements

4. TODO:

   - Obligation grouping without commitment numbers
   - Auto download
