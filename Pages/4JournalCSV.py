import streamlit as st
import pandas as pd
import os
from datetime import date

# Page Configuration
st.set_page_config(
    page_title="Journal CSV Exporter",
    page_icon="📄",
    layout="wide",
)

st.title("Journal CSV Exporter")

# File Upload
uploaded_file = st.file_uploader("Upload the file with DDI and Plan Journals:", type=["xlsx"])
row_limit = 4000  # Netsuite limit

# Calendar widget for journal month
journal_date = st.date_input("Select the journal date:")
selected_month = journal_date.strftime("%B")

if uploaded_file and journal_date:
    try:
        # Load the Excel file
        data = pd.ExcelFile(uploaded_file)

        # Check for required tabs
        if "DDI Journal" not in data.sheet_names or "Plan Journals" not in data.sheet_names:
            st.error("The uploaded file must contain 'DDI Journal' and 'Plan Journals' tabs.")
        else:
            # Load tabs
            ddi_df = pd.read_excel(data, sheet_name="DDI Journal")
            plan_df = pd.read_excel(data, sheet_name="Plan Journals")

            # Split Logic
            def split_and_export(df, journal_type):
                total_debit, total_credit = 0, 0
                split_count = 0
                output_files = []
                rows_processed = 0

                while not df.empty:
                    split_count += 1

                    # Create a split
                    split_df = pd.DataFrame()
                    while not df.empty and len(split_df) + 2 <= row_limit:
                        entry_group = df.iloc[:2]
                        split_df = pd.concat([split_df, entry_group], ignore_index=True)
                        df = df.iloc[2:]  # Remove processed rows

                    # Calculate totals for this split
                    debit_total = split_df["Debit"].sum()
                    credit_total = split_df["Credit"].sum()
                    total_debit += debit_total
                    total_credit += credit_total

                    # Save this split
                    file_name = f"SMS {journal_type.upper()} JOURNALS {selected_month} SPLIT {split_count}.csv"
                    split_df.to_csv(file_name, index=False)
                    output_files.append(file_name)

                    # Log progress
                    st.write(
                        f"Split {split_count}: Processed Debit: ${debit_total:,.2f}, Credit: ${credit_total:,.2f}"
                    )

                    rows_processed += len(split_df)

                # Log final totals
                st.write(
                    f"Total {journal_type.capitalize()} Journal: Debit: ${total_debit:,.2f}, Credit: ${total_credit:,.2f}"
                )

                return output_files

            # Process DDI Journal
            st.header("Processing DDI Journal...")
            ddi_files = split_and_export(ddi_df, "DDI")

            # Process Plan Journals
            st.header("Processing Plan Journals...")
            plan_files = split_and_export(plan_df, "Plan")

            # Display download links
            def zip_files(file_list, output_zip):
                import zipfile

                with zipfile.ZipFile(output_zip, "w") as zipf:
                    for file in file_list:
                        zipf.write(file)
                        os.remove(file)  # Clean up after zipping

            all_files = ddi_files + plan_files
            zip_name = f"Journal_Splits_{selected_month}.zip"
            zip_files(all_files, zip_name)

            st.success("Splits created successfully!")
            with open(zip_name, "rb") as zip_file:
                st.download_button(
                    label="Download All Splits (ZIP)",
                    data=zip_file,
                    file_name=zip_name,
                    mime="application/zip",
                )

    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
else:
    st.info("Please upload the file and select the journal date to proceed.")
