import streamlit as st
import pandas as pd
import re
from tableaudocumentapi import Workbook
from io import BytesIO
from pyvis.network import Network
import streamlit.components.v1 as components


# Function to clean strings
def remove_sp_char_leave_underscore_square_brackets(string_to_convert):
    return re.sub(r'[^a-zA-Z0-9\s_\[\]]', '', string_to_convert).replace(' ', "_")

# Function to process uploaded Tableau workbook
def process_workbook(file):
    try:
        TWBX_Workbook = Workbook(file)
        collator = []

        # Extract data from the Tableau workbook
        c = 0
        for datasource in TWBX_Workbook.datasources:
            datasource_name = datasource.name
            datasource_caption = datasource.caption if datasource.caption else datasource_name

            for count, field in enumerate(datasource.fields.values()):
                dict_temp = {
                    'counter': c,
                    'datasource_name': datasource_name,
                    'datasource_caption': datasource_caption,
                    'alias': field.alias,
                    'field_calculation': field.calculation,
                    'field_calculation_bk': field.calculation,
                    'field_caption': field.caption,
                    'field_datatype': field.datatype,
                    'field_def_agg': field.default_aggregation,
                    'field_desc': field.description,
                    'field_hidden': field.hidden,
                    'field_id': field.id,
                    'field_is_nominal': field.is_nominal,
                    'field_is_ordinal': field.is_ordinal,
                    'field_is_quantitative': field.is_quantitative,
                    'field_name': field.name,
                    'field_role': field.role,
                    'field_type': field.type,
                    'field_worksheets': field.worksheets
                }
                c += 1
                collator.append(dict_temp)

        df = pd.DataFrame(collator)
        return df

    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
        return None

# Streamlit App
st.set_page_config(page_title="Tableau Workbook Analyzer", layout="wide")
st.title("Tableau Workbook Analyzer")
st.write("Upload your `.twbx` files to analyze their fields and view the results.")

# Add footer
st.markdown(
    """
    <style>
    .footer {
        position: fixed;
        bottom: 0;
        left: 0;
        right: 0;
        background-color: #f0f0f0;
        text-align: center;
        padding: 10px 0;
        font-size: 12px;
        color: #333;
    }
    </style>
    <div class="footer">
        Code inspiration taken from: <a href="https://github.com/scinana/tableauCalculationExport" target="_blank">https://github.com/scinana/tableauCalculationExport</a><br>
        App developed by: Ali Iqbal
    </div>
    """,
    unsafe_allow_html=True
)

# Sidebar for customization
with st.sidebar:
    # File Upload Section for Multiple Files
    uploaded_files = st.file_uploader(
        "Upload Tableau Workbooks (.twbx)", type=["twbx"], accept_multiple_files=True
    )
    
    st.header("App Settings")
    show_visualizations = st.checkbox("Enable Field Dependencies Graph", value=True)
    export_format = st.selectbox("Choose Export Format", ["CSV", "JSON"])




if uploaded_files:    
    all_dfs = []
    for uploaded_file in uploaded_files:
        with st.spinner(f"Processing {uploaded_file.name}..."):
            # Process workbook from the uploaded file (in-memory)
            df = process_workbook(BytesIO(uploaded_file.getbuffer()))

            if df is not None:
                st.success(f"{uploaded_file.name} processed successfully!")
                all_dfs.append((uploaded_file.name, df))

    # Combine results across all files
    combined_df = pd.concat([df.assign(Source=file_name) for file_name, df in all_dfs], ignore_index=True)
    
    # Display combined DataFrame
    st.write("### Combined Data Across All Files")
    st.dataframe(combined_df, use_container_width=True)

    # Export options
    st.write("### Export Options")
    # if export_format == "Excel":
        # output_file = combined_df.to_excel(index=False)
        # st.download_button("Download as Excel", data=output_file, file_name="combined_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if export_format == "CSV":
        csv_output = combined_df.to_csv(index=False)
        st.download_button("Download as CSV", data=csv_output, file_name="combined_data.csv", mime="text/csv")
    elif export_format == "JSON":
        json_output = combined_df.to_json(orient="records")
        st.download_button("Download as JSON", data=json_output, file_name="combined_data.json", mime="application/json")

    # Visualize Field Dependencies Graph
    if show_visualizations:
        st.write("### Interactive Field Dependencies Graph")

        # Create a Pyvis Network
        net = Network(height="750px", width="100%", directed=True, notebook=False)

        # Track added nodes to avoid duplicates
        added_nodes = set()

        # Add nodes for each field
        for field in combined_df['field_name']:
            if field not in added_nodes:
                net.add_node(field, label=field)
                added_nodes.add(field)

        # Add edges based on dependencies
        for _, row in combined_df.iterrows():
            if row['field_calculation']:
                dependent_fields = re.findall(r'\[.*?\]', row['field_calculation'])  # Extract fields in calculation
                for dep in dependent_fields:
                    dep_clean = dep.strip("[]")  # Clean brackets

                    # Add dependent node if it doesn't already exist
                    if dep_clean not in added_nodes:
                        net.add_node(dep_clean, label=dep_clean)
                        added_nodes.add(dep_clean)

                    # Add the edge
                    net.add_edge(dep_clean, row['field_name'])

        # Generate the interactive graph HTML
        net_html_path = "graph.html"
        net.write_html(net_html_path)

        # Display the graph in Streamlit
        with open(net_html_path, "r", encoding="utf-8") as f:
            graph_html = f.read()
        components.html(graph_html, height=800)
