import re
import pandas as pd
import streamlit as st
import io
from io import StringIO

# Set Streamlit page configuration
st.set_page_config(
    page_title="LIC Agent Credit Calculator",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="expanded",
)

# Function to process agents' credit amounts
def process_agents(data, agent_codes):
    agent_codes = set(agent_codes)

    # Define patterns to extract required information
    agent_pattern = re.compile(r"Agency Code/Name\s*:\s*([A-Za-z0-9]+)\s*\((.*?)\)")

    # Define patterns for each Dr. Amount field
    patterns = {
        "First_Comm_Participating": re.compile(r"\|\s*11272100\s*\|\s*First Comm Participating\s*\|\s*(\d+\.\d+)\s*\|"),
        "First_Year_Comm_Participating": re.compile(r"\|\s*11270600\s*\|\s*First Year Comm Participating\s*\|\s*(\d+\.\d+)\s*\|"),
        "Bonus_Comm_Participating": re.compile(r"\|\s*11271200\s*\|\s*Bonus Comm Participating\s*\|\s*(\d+\.\d+)\s*\|"),
        "Renewal_Comm_Participating": re.compile(r"\|\s*11273100\s*\|\s*Renewal Comm Participating\s*\|\s*(\d+\.\d+)\s*\|"),
        "Comm_Other_FY_Prem": re.compile(r"\|\s*96270600\s*\|\s*Comm. on Other FY Prem. to Agents\(873\)\s*\|\s*(\d+\.\d+)\s*\|"),
        "Bonus_Comm_Agents": re.compile(r"\|\s*96271100\s*\|\s*Bonus Comm. to Agents\(873\)\s*\|\s*(\d+\.\d+)\s*\|"),
        "Income_Tax": re.compile(r"\|\s*11110300\s*\|\s*Income Tax\s*\|\s*(\d+\.\d+)\s*\|"),
    }

    # Pattern to extract Cr. Amount values
    cr_amount_pattern = re.compile(r"\|\s*\d+\s*\|\s*.*?\s*\|\s*\d+\.\d+\s*\|\s*(\d+\.\d+)\s*\|")

    # Split the data by vouchers
    vouchers = data.split("****Voucher ")

    # Initialize a dictionary to store results
    results = {}
    for voucher in vouchers:
        # Extract agent code and name
        agent_match = agent_pattern.search(voucher)
        if not agent_match:
            continue

        agent_code, agent_name = agent_match.groups()
        if agent_code not in agent_codes:
            continue

        # Extract Dr. Amounts using patterns
        extracted_values = {key: pattern.search(voucher) for key, pattern in patterns.items()}
        dr_amounts = {key: (float(match.group(1)) if match else "-") for key, match in extracted_values.items()}

        # Calculate the sum of Dr. Amounts (only numeric values)
        sum_dr_amount = sum(value for value in dr_amounts.values() if isinstance(value, float)) if any(isinstance(value, float) for value in dr_amounts.values()) else "-"

        # Extract Cr. Amounts
        cr_amounts = cr_amount_pattern.findall(voucher)
        total_cr_amount = sum(map(float, cr_amounts))

        # Add to results
        if agent_code in results:
            for key in dr_amounts:
                if isinstance(dr_amounts[key], float):
                    results[agent_code][key] += dr_amounts[key]
            results[agent_code]["Sum_Dr_Amount"] += sum_dr_amount if isinstance(sum_dr_amount, float) else 0
            results[agent_code]["Cr_Amount"] += total_cr_amount
        else:
            results[agent_code] = {
                "Name": agent_name,
                **dr_amounts,
                "Sum_Dr_Amount": sum_dr_amount,
                "Cr_Amount": total_cr_amount,
            }

    return results

# Streamlit App
st.title("📋 LIC Agent Credit Amount Calculator")
st.markdown("### Calculate and Summarize Credit Amounts for LIC Agents")

# File Upload
uploaded_file = st.file_uploader("Upload your LIC data file (.txt/.prt)", type=["txt", "prt"])
if uploaded_file:
    # Read the uploaded file
    data = StringIO(uploaded_file.getvalue().decode("utf-8")).read()
    st.success("File uploaded successfully!")

    # Agent Codes Input
    st.markdown("### Provide Agent Codes")
    input_type = st.radio(
        "Select input method for agent codes:",
        ("Single Code", "Comma-Separated Codes", "Upload Excel File")
    )

    agent_codes = []
    if input_type == "Single Code":
        code = st.text_input("Enter a single agent code:")
        if code:
            agent_codes = [code.strip()]
    elif input_type == "Comma-Separated Codes":
        codes = st.text_area("Enter comma-separated agent codes:")
        if codes:
            agent_codes = [code.strip() for code in codes.split(",")]
    elif input_type == "Upload Excel File":
        excel_file = st.file_uploader("Upload Excel file with agent codes", type=["xlsx"])
        if excel_file:
            df = pd.read_excel(excel_file)
            if not df.empty:
                agent_codes = df.iloc[1:, 0].dropna().astype(str).tolist()
            st.success("Agent codes extracted from Excel!")

    # Process and Output Results
    if agent_codes:
        st.write("Processing with the following agent codes:", agent_codes)
        results = process_agents(data, agent_codes)

        if results:
            st.success("Processing complete!")
            df = pd.DataFrame.from_dict(results, orient='index')
            df.index.name = 'Agent_Code'
            df.reset_index(inplace=True)
            st.dataframe(df)

            # Create an in-memory buffer
            output = io.BytesIO()

            # Save DataFrame to Excel file in memory
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name="Summary")
                writer.close()
            output.seek(0)  # Rewind the buffer to the beginning

            # Provide download link
            st.download_button(
                label="Download Results as Excel",
                data=output,
                file_name="agent_credit_totals.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No matching agent codes found in the data.")
