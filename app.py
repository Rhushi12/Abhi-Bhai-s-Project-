import streamlit as st
import pandas as pd
import os
import json
from automation import automate_transfer, FIELD_MAPPINGS

# Page Config
st.set_page_config(page_title="Excel Automation", page_icon="📊", layout="wide")

# Theme-Aware CSS
st.markdown("""
<style>
    /* Global Font */
    html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif, "Apple Color Emoji", "Segoe UI Emoji", "Segoe UI Symbol";
    }

    /* Cards */
    div[data-testid="stFileUploader"], div.stDataFrame {
        background-color: var(--secondary-background-color);
        padding: 2rem;
        border-radius: 15px;
        border: 1px solid var(--border-color-primary);
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        transition: transform 0.2s, box-shadow 0.2s;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: var(--primary-color);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    }

    /* Primary Button Styling */
    .stButton>button {
        background-color: var(--primary-color);
        color: white;
        border-radius: 20px;
        padding: 0.6rem 2rem;
        border: none;
        font-weight: 600;
        font-size: 1rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        transition: all 0.2s ease-in-out;
        width: 100%;
    }
    .stButton>button:hover {
        filter: brightness(110%);
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.15);
    }
    .stButton>button:active {
        transform: translateY(0);
    }

    /* Success Box (Adaptive) */
    .success-box {
        padding: 1.5rem;
        background-color: rgba(76, 175, 80, 0.15); /* Transparent green */
        color: var(--text-color);
        border: 1px solid #4CAF50;
        border-radius: 12px;
        margin-top: 2rem;
        text-align: center;
        font-weight: 500;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 20px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        border-radius: 10px 10px 0 0;
        padding: 0 20px;
    }
    
    /* Headings */
    h1, h2, h3 {
        color: var(--text-color);
    }
    h1 {
        text-align: center;
        margin-bottom: 2rem;
        font-weight: 800;
        color: var(--primary-color);
    }
    
    /* Info Box styling */
    .stAlert {
        border-radius: 10px;
    }
</style>
""", unsafe_allow_html=True)

st.title("Excel Automation Tool")

# Create Tabs
tab1, tab2 = st.tabs(["▶ Run Automation", "⚙️ Configure Mappings"])

# ==========================================
# TAB 1: RUN AUTOMATION
# ==========================================
with tab1:
    col_center = st.columns([1, 8, 1])
    with col_center[1]:
        st.markdown("### 📄 Process Excel Files")
        st.caption("Upload your file to automatically transfer data from *Base data* to *Factsheet*. The magic happens securely on your machine.")
        
        uploaded_file = st.file_uploader("Drop your Excel file here", type=['xlsx', 'xls'])

        if uploaded_file is not None:
            st.divider()
            
            # Using columns to center the button
            c1, c2, c3 = st.columns([1, 2, 1])
            with c2:
                if st.button("▶ Start Processing", use_container_width=True):
                    with st.spinner("Analyzing spreadsheet structure..."):
                        try:
                            # Save uploaded file temporarily
                            with open("temp_input.xlsx", "wb") as f:
                                f.write(uploaded_file.getbuffer())
                            
                            # Run the automation
                            output_path = automate_transfer("temp_input.xlsx", "temp_output.xlsx")
                            
                            if output_path and os.path.exists(output_path):
                                st.balloons()
                                
                                st.markdown(f"""
                                    <div class="success-box">
                                        <div style="font-size: 2rem; margin-bottom: 0.5rem;">🎉</div>
                                        Processing Complete!<br>
                                        <span style="font-size: 0.9em; opacity: 0.8">Your file has been updated with the latest data.</span>
                                    </div>
                                """, unsafe_allow_html=True)
                                
                                st.write("") # Spacer
                                
                                with open(output_path, "rb") as f:
                                    st.download_button(
                                        label="⬇ Download Formatted File",
                                        data=f,
                                        file_name=f"Formatted_{uploaded_file.name}",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        use_container_width=True
                                    )
                                
                                # Cleanup
                                os.remove("temp_input.xlsx")
                                os.remove("temp_output.xlsx")
                            else:
                                st.error("Processing failed. Please check the logs.")
                                
                        except Exception as e:
                            st.error(f"Error during processing: {e}")

# ==========================================
# TAB 2: CONFIGURE MAPPINGS
# ==========================================
with tab2:
    col_center = st.columns([1, 10, 1])
    with col_center[1]:
        st.markdown("### ⚙️ Edit Field Mappings")
        st.info("Define which keywords in *Base data* map to which rows in the *Factsheet*. This makes your automation smarter!")
        
        # Load current mappings
        json_path = 'mappings.json'
        
        # Try loading from file first, else use imported defaults
        if os.path.exists(json_path):
            try:
                with open(json_path, 'r') as f:
                    mappings = json.load(f)
            except Exception:
                st.error("⚠ Mappings file is corrupt. Loading defaults.")
                mappings = FIELD_MAPPINGS
        else:
            mappings = FIELD_MAPPINGS

        # ---------------------------------------------------------
        # POWER TOOLS SECTION
        # ---------------------------------------------------------
        with st.expander("🛠️ Advanced Tools: Add, Swap, Delete", expanded=False):
            t1, t2, t3 = st.tabs(["Add New Field", "Swap Rules", "Delete Fields"])
            
            # --- 1. ADD NEW FIELD ---
            with t1:
                st.write("Add a new mapping rule.")
                c1, c2 = st.columns([1, 2])
                new_key = c1.text_input("Factsheet Field Name (Key)", placeholder="e.g. Total Revenue")
                new_vars = c2.text_input("Matching Variations (comma separated)", placeholder="revenue, sales, total")
                
                if st.button("➕ Add Rule"):
                    if not new_key:
                        st.error("Field Name is required.")
                    elif new_key in mappings:
                        st.error("This field already exists!")
                    else:
                        vars_list = [v.strip().lower() for v in new_vars.split(',') if v.strip()]
                        mappings[new_key] = vars_list
                        # Save
                        with open(json_path, 'w') as f:
                            json.dump(mappings, f, indent=4)
                        st.success(f"Added '{new_key}'!")
                        st.rerun()

            # --- 2. SWAP RULES ---
            with t2:
                st.write("Swap the variations between two fields (e.g., if Sales and Cost are mixed up).")
                keys = list(mappings.keys())
                sc1, sc2 = st.columns(2)
                key1 = sc1.selectbox("Field A", keys, index=0)
                # Try to pick a different default for key2
                default_idx_2 = 1 if len(keys) > 1 else 0
                key2 = sc2.selectbox("Field B", keys, index=default_idx_2)
                
                if st.button("🔄 Swap Rules"):
                    if key1 == key2:
                        st.warning("Please select two different fields.")
                    else:
                        # Swap
                        mappings[key1], mappings[key2] = mappings[key2], mappings[key1]
                        # Save
                        with open(json_path, 'w') as f:
                            json.dump(mappings, f, indent=4)
                        st.success(f"Swapped rules for '{key1}' and '{key2}'!")
                        st.rerun()

            # --- 3. DELETE FIELDS ---
            with t3:
                st.write("Remove fields you don't need.")
                to_delete = st.multiselect("Select fields to remove", list(mappings.keys()))
                
                if st.button("🗑️ Delete Selected"):
                    if not to_delete:
                        st.warning("Select at least one field.")
                    else:
                        for k in to_delete:
                            if k in mappings:
                                del mappings[k]
                        # Save
                        with open(json_path, 'w') as f:
                            json.dump(mappings, f, indent=4)
                        st.success(f"Deleted {len(to_delete)} fields!")
                        st.rerun()

        st.divider()

        # ---------------------------------------------------------
        # MAIN EDITOR TABLE
        # ---------------------------------------------------------
        # Convert to DataFrame for editing
        current_data = []
        for field, variations in mappings.items():
            list_vars = variations if isinstance(variations, list) else []
            current_data.append({
                "Factsheet Field Name": field,
                "Base Data Variations (comma separated)": ", ".join(list_vars)
            })
        
        df = pd.DataFrame(current_data)
        
        # Editable Table
        edited_df = st.data_editor(
            df, 
            num_rows="dynamic", 
            use_container_width=True,
            hide_index=True,
            column_config={
                "Factsheet Field Name": st.column_config.TextColumn("Factsheet Field (Key)", required=True, width="medium", help="The exact row name in the Factsheet tab"),
                "Base Data Variations (comma separated)": st.column_config.TextColumn("Matching Variations", width="large", help="Names to look for in Base data (separate by comma)")
            }
        )
        
        st.caption("💡 Tip: You can also edit directly in the table above. Don't forget to click Save!")
        st.divider()
        
        # Save Button
        c1, c2, c3 = st.columns([2, 1, 2])
        with c2:
            if st.button("💾 Save Table Changes"):
                new_mappings = {}
                seen_keys = set()
                has_errors = False
                
                for index, row in edited_df.iterrows():
                    key = row["Factsheet Field Name"]
                    vars_str = row["Base Data Variations (comma separated)"]
                    
                    if not key and not vars_str:
                        continue
                        
                    if not key or not str(key).strip():
                        st.warning(f"⚠ Row {index + 1}: Factsheet Field Name cannot be empty. This row was skipped.")
                        continue
                    
                    key = str(key).strip()
                    
                    if key in seen_keys:
                        st.error(f"❌ Duplicate Field Name: '{key}'. Keys must be unique.")
                        has_errors = True
                        continue
                    seen_keys.add(key)
                    
                    if vars_str and str(vars_str).strip():
                        vars_list = [v.strip().lower() for v in str(vars_str).split(',') if v.strip()]
                    else:
                        vars_list = [] 
                        
                    new_mappings[key] = vars_list
                
                if not has_errors:
                    try:
                        with open(json_path, 'w') as f:
                            json.dump(new_mappings, f, indent=4)
                        st.toast("Configuration saved successfully!", icon="✅")
                        st.balloons()
                        import time
                        time.sleep(1)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Failed to save: {e}")
