# 🎈 SUNFIN APP

Application to format Income Statements

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://blank-app-template.streamlit.app/)

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
