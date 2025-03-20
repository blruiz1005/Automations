import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import zipfile
import re
from datetime import datetime, timedelta
from io import BytesIO

# Add custom CSS to change sidebar color
st.markdown(
    """
    <style>
    /* Sidebar background color */
    .css-1d391kg {
        background-color: #000;  /* Example: dark background */
    }

    /* Sidebar header text color */
    .css-1d391kg .sidebar .sidebar-header h2 {
        color: #ff6347;  /* Tomato color for header */
    }

    /* Sidebar text color */
    .css-1d391kg .sidebar .sidebar-content {
        color: white;  /* White text color */
    }

    /* Sidebar button background and text color */
    .css-1d391kg .sidebar .sidebar-content .stButton button {
        background-color: #ff6347;  /* Tomato color for buttons */
        color: white;
    }

    /* Adjusting file uploader button style */
    .css-1d391kg .block-container .stFileUploader > div > button {
        background-color: #ff6347;
        color: white;
    }
    </style>
    """, unsafe_allow_html=True
)



# Define the regex pattern for matching the filename
FILENAME_PATTERN = r"BPI CARDS XDAYS Bcrm Upload (\d+) as of \d{4}-\d{2}-\d{2}\.xls"

# List of required columns
REQUIRED_COLUMNS = [
    "CYCLE", "Phone Number", "Rang", "System Disposition", "LEN", "FINANCIER ID", "APPLICATION ID",
    "CUSTOMER ID", "USER ID", "ACTION DATE", "ACTION TIME", "ACTION CODE", "CONTACT MODE",
    "PERSON CONTACTED", "PLACE CONTACTED", "CURRENCY", "ACTION AMOUNT", "NEXT ACTION DATE",
    "NEXT ACTION TIME", "REMINDER MODE", "CONTACTED BY", "REMARKS"
]

# Function for Call Logs Filtering
def call_logs_filtering():
    st.title("Call Logs (Filter)")
    st.markdown('<p style="color:#ff6347; font-size:24px; text-align:center;"><strong>Filter Your Call Logs</strong></p>', unsafe_allow_html=True)

    uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx", "xls"], accept_multiple_files=True)

    if uploaded_files:
        try:
            dfs = []
            for uploaded_file in uploaded_files:
                df = pd.read_excel(uploaded_file)
                dfs.append(df)

            df = pd.concat(dfs, ignore_index=True)

            if "Call Made Date" in df.columns:
                df["ACTION DATE"] = df["Call Made Date"]
            if "Time of Call" in df.columns:
                df["ACTION TIME"] = df["Time of Call"]
            if "Acct Number" in df.columns:
                df["CUSTOMER ID"] = df["Acct Number"].apply(
                    lambda x: f"0{str(int(x))}" if pd.notna(x) and str(x).replace('.0', '').isdigit() else f"0{x}" if pd.notna(x) else ""
                )

            if "Ch Code" in df.columns:
                df["CYCLE"] = df["Ch Code"].apply(
                    lambda x: str(x)[:2] if pd.notna(x) and len(str(x)) >= 2 else ""
                )
                ch_code_index = df.columns.get_loc("Ch Code")
                df.insert(ch_code_index, "CYCLE", df.pop("CYCLE"))

            st.subheader("Original Combined Data")
            st.write(df)

            for col in REQUIRED_COLUMNS:
                if col not in df.columns:
                    df[col] = ""

            try:
                if "Duration of the Call" in df.columns:
                    df = df[df["Duration of the Call"] == "00:00:00"]

                if "ACTION DATE" in df.columns:
                    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
                    df = df[pd.to_datetime(df["ACTION DATE"], errors='coerce').dt.strftime("%Y-%m-%d") == yesterday]

                if "Acct Number" in df.columns:
                    df = df[df["Acct Number"].notna()]

                filtered_df = df[REQUIRED_COLUMNS]

                st.subheader("Filtered Data")
                st.write(filtered_df)

                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    filtered_df.to_excel(writer, index=False, sheet_name="Filtered Data")

                st.download_button(
                    label="Download Filtered Data",
                    data=output.getvalue(),
                    file_name="filtered_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Error during filtering: {e}")

        except Exception as e:
            st.error(f"Error reading the files: {e}")
    else:
        st.warning("Please upload at least one Excel file.")

# Function for Extracting Endo ZIP
def extracting_endo_zip():
    def extract_matching_xls(zip_file):
        matching_files = []
        with zipfile.ZipFile(zip_file) as z:
            for file_name in z.namelist():
                if re.fullmatch(FILENAME_PATTERN, file_name):
                    extracted_file = z.open(file_name)
                    matching_files.append((file_name, extracted_file.read()))
        return matching_files

    def get_cycle_from_xls(file_content):
        try:
            df = pd.read_excel(BytesIO(file_content))
            if "Ch Code" not in df.columns:
                return {"error": "Ch Code column not found"}

            df["Cycle"] = df["Ch Code"].astype(str).str[:2]
            cycle_counts = df["Cycle"].value_counts().to_dict()

            return cycle_counts
        except Exception as e:
            return {"error": f"Error reading file: {e}"}

    st.title("Extract .xls Files from Endo ZIP File")
    st.markdown('<p style="color:#ff6347; font-size:24px; text-align:center;"><strong>Extract and Organize Endo Files</strong></p>', unsafe_allow_html=True)

    st.write(
        "Upload one or more `.zip` files. The app will extract `.xls` files with filenames matching the pattern:\n\n"
        "**`BPI CARDS XDAYS Bcrm Upload <number> as of YYYY-MM-DD.xls`**\n\n"
        "It will also extract cycle information based on the first two digits of the 'Ch Code' column."
    )

    uploaded_files = st.file_uploader("Upload .zip files", type=["zip"], accept_multiple_files=True)

    if uploaded_files:
        extracted_data = []
        cycle_counts = {}
        file_numbers = {}
        cycle_grouped_files = {}

        for uploaded_file in uploaded_files:
            extracted_files = extract_matching_xls(uploaded_file)
            if extracted_files:
                for file_name, file_content in extracted_files:
                    cycle_info = get_cycle_from_xls(file_content)

                    match = re.search(FILENAME_PATTERN, file_name)
                    if match:
                        file_number = match.group(1)
                        file_numbers[file_number] = file_numbers.get(file_number, 0) + 1

                    extracted_data.append((file_name, file_content, cycle_info))

                    if "error" not in cycle_info:
                        for cycle, count in cycle_info.items():
                            cycle_counts[cycle] = cycle_counts.get(cycle, 0) + count

                            if cycle not in cycle_grouped_files:
                                cycle_grouped_files[cycle] = []
                            cycle_grouped_files[cycle].append((file_name, file_content))
            else:
                st.warning(f"No matching files found in {uploaded_file.name}")

        if extracted_data:
            st.subheader("Extracted Files and Cycle Information")
            cycle_info_list = []
            for file_name, _, cycle_info in extracted_data:
                if "error" in cycle_info:
                    cycle_info_list.append([file_name, cycle_info["error"]])
                else:
                    cycle_info_str = ", ".join([f"Cycle {key}: {value}" for key, value in cycle_info.items()])
                    cycle_info_list.append([file_name, cycle_info_str])

            cycle_info_df = pd.DataFrame(cycle_info_list, columns=["File Name", "Cycle Information"])
            st.dataframe(cycle_info_df)

            st.subheader("Total Cycle Counts")
            cycle_counts_df = pd.DataFrame(list(cycle_counts.items()), columns=["Cycle", "Count"]).sort_values(by="Count", ascending=False)
            st.dataframe(cycle_counts_df)

            for cycle, files in cycle_grouped_files.items():
                with BytesIO() as buffer:
                    with zipfile.ZipFile(buffer, "w") as zip_buffer:
                        for file_name, file_content in files:
                            zip_buffer.writestr(file_name, file_content)

                    buffer.seek(0)
                    st.download_button(
                        label=f"Download Cycle {cycle} Files",
                        data=buffer,
                        file_name=f"Cycle_{cycle}_BPI_CARDS_Files.zip",
                        mime="application/zip",
                    )

        else:
            st.info("No matching files were extracted.")

# Function for Cycle Count
def cycle_count():
    st.title("Cycle Count")
    st.markdown('<p style="color:#ff6347; font-size:24px; text-align:center;"><strong>Count Specific Cycles in Your Data</strong></p>', unsafe_allow_html=True)

    uploaded_file = st.file_uploader("Upload an Excel file", type=['xls', 'xlsx'])

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, engine='openpyxl')

            st.write("### Preview of the uploaded file:")
            st.dataframe(df.head())

            if 'CH CODE' not in df.columns:
                st.error("The uploaded file does not contain a 'CH CODE' column.")
            else:
                df['CH CODE'] = df['CH CODE'].astype(str).str[:2]  # Take the first two characters

                target_cycles = ['05', '14', '15', '27']
                cycle_counts = {cycle: (df['CH CODE'] == cycle).sum() for cycle in target_cycles}

                st.write("### Cycle Counts:")
                for cycle, count in cycle_counts.items():
                    color = "green" if count > 0 else "red"
                    st.markdown(f"<p style='color:{color}; font-size:20px;'><strong>Cycle {cycle}: {count}</strong></p>", unsafe_allow_html=True)

                cycle_count_df = pd.DataFrame(list(cycle_counts.items()), columns=["Cycle", "Count"])
                st.write("### Cycle Count Table:")
                st.dataframe(cycle_count_df)

        except Exception as e:
            st.error(f"An error occurred while processing the file: {e}")
    else:
        st.info("Please upload an Excel file to begin.")

# Main app logic
st.sidebar.title("Automation Selector")
option = st.sidebar.radio("Select an automation:", ["Call Logs Filtering", "Extracting Endo Zip", "Cycle Count"])

if option == "Call Logs Filtering":
    call_logs_filtering()
elif option == "Extracting Endo Zip":
    extracting_endo_zip()
elif option == "Cycle Count":
    cycle_count()
