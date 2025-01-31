import streamlit as st
import pandas as pd
import json
from itertools import product

def validate_data(prompts, values, models, parameters):
    errors = []
    
    if "prompt_id" not in prompts.columns:
        errors.append("The Prompts sheet must have a 'prompt_id' column.")
    if "prompt_id" not in values.columns:
        errors.append("The Values sheet must have a 'prompt_id' column.")
    if "Name" not in models.columns or "Version" not in models.columns:
        errors.append("The Models sheet must have 'Name' and 'Version' columns.")
    if parameters.empty or parameters.isnull().any().any():
        errors.append("Each parameter in the Parameters sheet must have a name and values.")
    
    return "\n".join(errors) if errors else None

def generate_combinations(prompts, values, models, parameters):
    results = []
    parameter_names = parameters.iloc[:, 0].tolist()
    parameter_values = [str(v).split(',') for v in parameters.iloc[:, 1]]
    param_combinations = list(product(*parameter_values))
    
    for _, prompt_row in prompts.iterrows():
        prompt_id = prompt_row["prompt_id"]
        prompt_template = str(prompt_row[1])  # Ensure this is correctly extracted
        
        matched_values = values[values["prompt_id"] == prompt_id]
        for _, value_row in matched_values.iterrows():
            filled_prompt = prompt_template
            for col in values.columns[1:]:
                filled_prompt = filled_prompt.replace(f"{{{{{col}}}}}", str(value_row[col]))
            
            for _, model_row in models.iterrows():
                for param_set in param_combinations:
                    results.append([
                        prompt_id,
                        model_row["Name"],
                        model_row["Version"],
                        *param_set,
                        filled_prompt
                    ])
    
    columns = ["prompt_id", "model_name", "model_version", *parameter_names, "filled_prompt"]
    return pd.DataFrame(results, columns=columns)

st.title("Spreadsheet Validator and Combination Generator")

sample_file_path = "example.xlsx"
st.download_button("Download Sample Spreadsheet", data=open(sample_file_path, "rb"), file_name="example.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

uploaded_file = st.file_uploader("Upload your spreadsheet", type=["xls", "xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    required_sheets = ["Prompts", "Values", "Models", "Parameters"]
    
    if not all(sheet in xls.sheet_names for sheet in required_sheets):
        st.error("Missing one or more required sheets: Prompts, Values, Models, Parameters")
    else:
        prompts = pd.read_excel(xls, sheet_name="Prompts")
        values = pd.read_excel(xls, sheet_name="Values")
        models = pd.read_excel(xls, sheet_name="Models")
        parameters = pd.read_excel(xls, sheet_name="Parameters", header=None)
        
        validation_message = validate_data(prompts, values, models, parameters)
        
        if validation_message:
            st.error(f"Validation Error:\n{validation_message}")
        else:
            st.success("Validation Passed!")
            
            with st.expander("Preview of Uploaded Data"):
                st.write("#### Prompts Sheet")
                st.dataframe(prompts.head())
                st.write("#### Values Sheet")
                st.dataframe(values.head())
                st.write("#### Models Sheet")
                st.dataframe(models.head())
                st.write("#### Parameters Sheet")
                st.dataframe(parameters.head())
            
            combinations = generate_combinations(prompts, values, models, parameters)
            st.write(f"### Generated {len(combinations)} Rows")
            st.dataframe(combinations)
            
            csv = combinations.to_csv(index=False).encode('utf-8')
            st.download_button("Download CSV", csv, "combinations.csv", "text/csv")
            
            json_data = combinations.to_dict(orient='records')
            json_str = json.dumps(json_data, indent=4)
            st.download_button("Download JSON", json_str, "combinations.json", "application/json")
