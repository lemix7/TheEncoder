"""
SPSS Prep Tool - Streamlit Application
Automates preparing Google Forms Excel exports for IBM SPSS analysis.

TODO LIST (In-Code):
[P0] Basic file upload and preview - DONE
[P0] Detect unique values per column - DONE
[P0] Column configuration cards - DONE
[P0] Reorder options UI (Up/Down buttons) - DONE
[P0] Generate encoded Excel file - DONE
[P0] Generate SPSS .sps syntax - DONE
[P1] Auto-updating TODO sidebar - DONE
[P1] Variable name sanitization - DONE
[P1] Multi-response detection & warning - DONE
[P2] Preview SPSS syntax in UI - DONE
[P2] Download buttons for files - DONE
"""

import pathlib
import streamlit as st
import pandas as pd
import os
import tempfile
from typing import Dict, List, Optional
import logging

from spss_prep.encoder import detect_columns, ColumnConfig, apply_encoding, save_encoded_excel
from spss_prep.sps_generator import generate_sps_syntax, save_sps_file
from spss_prep.utils import sanitize_variable_name, generate_unique_var_names, is_likely_likert, is_multi_response


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def load_css(file_path):
    try:
        with open(file_path, 'r') as f:
            css_content = f.read()
        st.markdown(f"<style>{css_content}</style>", unsafe_allow_html=True)
    except Exception as e:
        st.error(f"Error loading CSS: {e}")


# Page config
st.set_page_config(
    page_title="SPSS Prep Tool",
    page_icon="📊",
    layout="wide"
)

# Load CSS
css_path = pathlib.Path('src/styles/style.css')
load_css(str(css_path))


def render_navbar():
    """Render the navigation bar."""
    navbar_html = """
    <nav class="navbar">
        <h3 class="navbar-brand">
            <span class="navbar-brand-icon">📊 </span>
            SPSS Prep Tool
        </h3>
    </nav>
    """
    st.markdown(navbar_html, unsafe_allow_html=True)


# Initialize session state
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None
if 'df' not in st.session_state:
    st.session_state.df = None
if 'column_info' not in st.session_state:
    st.session_state.column_info = {}
if 'column_configs' not in st.session_state:
    st.session_state.column_configs = {}
if 'column_orders' not in st.session_state:
    st.session_state.column_orders = {}
if 'encoded_df' not in st.session_state:
    st.session_state.encoded_df = None
if 'sps_syntax' not in st.session_state:
    st.session_state.sps_syntax = None
if 'encoded_path' not in st.session_state:
    st.session_state.encoded_path = None
if 'sps_path' not in st.session_state:
    st.session_state.sps_path = None
if 'unique_var_names' not in st.session_state:
    st.session_state.unique_var_names = {}


def get_todo_status() -> Dict[str, bool]:
    """Calculate TODO completion status based on app state."""
    status = {
        'upload': st.session_state.uploaded_file is not None,
        'detect': st.session_state.column_info != {},
        'configure': len(st.session_state.column_configs) > 0,
        'apply': st.session_state.encoded_df is not None,
        'preview': st.session_state.sps_syntax is not None,
        'download': st.session_state.encoded_path is not None
    }
    return status


def render_sidebar():
    """Render sidebar with TODO checklist and settings"""

    # Progress
    st.sidebar.subheader("✅ Progress")
    todo_status = get_todo_status()

    todos = [
        ('upload', '1. Upload Excel file'),
        ('detect', '2. Detect columns & options'),
        ('configure', '3. Configure encodings'),
        ('apply', '4. Apply encoding and generate files'),
        ('preview', '5. Preview SPSS script'),
        ('download', '6. Download files')
    ]

    completed_count = sum(1 for key, _ in todos if todo_status.get(key, False))
    total_count = len(todos)
    progress_fraction = completed_count / total_count if total_count else 0
    st.sidebar.progress(progress_fraction)
    st.sidebar.caption(f"{completed_count} of {total_count} steps completed")

    # Build styled checklist (non-interactive, reflects state)
    first_incomplete_index = next(
        (i for i, (k, _) in enumerate(todos) if not todo_status[k]), None)
    checklist_items = []
    for idx, (key, label) in enumerate(todos):
        is_completed = bool(todo_status[key])
        is_active = (first_incomplete_index == idx)
        item_class = "check-item completed" if is_completed else (
            "check-item active" if is_active else "check-item")
        icon = "<span class=\"ci-icon\">✓</span>" if is_completed else "<span class=\"ci-box\"></span>"
        checklist_items.append(
            f"<div class=\"{item_class}\">{icon}<span class=\"ci-text\">{label}</span></div>")

    st.sidebar.markdown(
        """
        <div class="sidebar-card">
          <div class="sidebar-card-title"><span class="dot"></span> Progress Checklist</div>
          <div class="checklist">
            {items}
          </div>
        </div>
        """.format(items="\n".join(checklist_items)),
        unsafe_allow_html=True,
    )

    # Settings card
    st.sidebar.markdown(
        """
        <div class="sidebar-card">
          <div class="sidebar-card-title"><span class="gear">⚙️</span> Settings</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    include_save = st.sidebar.toggle(
        "Include SAVE OUTFILE",
        value=False,
        help="Add SAVE OUTFILE command to .sps to also create a .sav file",
    )

    write_same_folder = st.sidebar.toggle(
        "Write .sps to same folder as encoded file",
        value=True,
        help="Place the .sps file next to the encoded Excel file",
    )

    sanitize_names = st.sidebar.toggle(
        "Sanitize variable names for SPSS",
        value=True,
        help="Convert column names to SPSS-compatible names and ensure uniqueness",
    )

    st.sidebar.markdown("---")

    # Actions
    if st.sidebar.button("🔄 Reset UI / Clear State", type="secondary", help="Clear all app state and start over"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

    return include_save, write_same_folder, sanitize_names


def move_option_up(column: str, index: int):
    """Move an option up in the order."""
    if index > 0:
        order = st.session_state.column_orders[column]
        order[index], order[index-1] = order[index-1], order[index]
        st.session_state.column_orders[column] = order


def move_option_down(column: str, index: int):
    """Move an option down in the order."""
    order = st.session_state.column_orders[column]
    if index < len(order) - 1:
        order[index], order[index+1] = order[index+1], order[index]
        st.session_state.column_orders[column] = order


def render_column_card(col_name: str, col_info: Dict, sanitize_names: bool):
    """Render configuration card for a single column."""

    # Initialize order if not exists
    if col_name not in st.session_state.column_orders:
        st.session_state.column_orders[col_name] = col_info['unique_values'].copy(
        )

    # Initialize config if not exists
    if col_name not in st.session_state.column_configs:
        # Detect if numeric, ordinal, or nominal
        if col_info['is_numeric']:
            default_type = 'Scale'
        elif is_likely_likert(col_info['unique_values']):
            default_type = 'Ordinal'
        else:
            default_type = 'Nominal'
        # Use pre-generated unique name or fallback to simple sanitization
        if sanitize_names and col_name in st.session_state.unique_var_names:
            sanitized = st.session_state.unique_var_names[col_name]
        else:
            sanitized = sanitize_variable_name(
                col_name) if sanitize_names else col_name

        st.session_state.column_configs[col_name] = {
            'encoding_type': default_type,
            'start_value': 1,
            'direction': 'Ascending',
            'treat_missing': True,
            'sanitized_name': sanitized
        }

    config = st.session_state.column_configs[col_name]

    with st.expander(f"📋 **{col_name}**", expanded=False):
        col1, col2 = st.columns([2, 1])

        with col1:
            # Show metadata
            n_unique = col_info['n_unique']
            n_missing = col_info['n_missing']

            if col_info['is_numeric']:
                st.info(
                    f"ℹ️ Detected as numeric column ({n_unique} unique values)")
            elif col_info['has_multi_response']:
                st.warning(
                    f"⚠️ Multi-response detected ({n_unique} unique combinations)")
                st.caption(
                    "Contains comma/semicolon separators. Currently treating as atomic strings.")
            else:
                if col_info['is_numeric']:
                    type_hint = "Scale (numeric)"
                elif is_likely_likert(col_info['unique_values']):
                    type_hint = "Ordinal (ordered categories)"
                else:
                    type_hint = "Nominal (categories)"
                st.caption(
                    f"📊 {n_unique} unique values | {n_missing} missing | {type_hint}")

        with col2:
            # Sanitized variable name
            if sanitize_names:
                sanitized = st.text_input(
                    "SPSS Variable Name",
                    value=config['sanitized_name'],
                    key=f"sanitized_{col_name}",
                    help="SPSS-compatible variable name"
                )
                config['sanitized_name'] = sanitized

        st.markdown("---")

        # Encoding controls
        col_a, col_b, col_c = st.columns(3)

        with col_a:
            measure_type = st.selectbox(
                "Measure (SPSS Variable Level)",
                options=['Ordinal', 'Nominal', 'Scale', 'Ignore'],
                index=['Ordinal', 'Nominal', 'Scale', 'Ignore'].index(config['encoding_type']) if config['encoding_type'] in [
                    'Ordinal', 'Nominal', 'Scale', 'Ignore'] else 0,
                key=f"type_{col_name}",
                help="Ordinal: ordered categories (e.g., Likert scales) | Nominal: unordered categories | Scale: continuous numeric | Ignore: don't encode"
            )
            # Update config with new value (backward compatible: Likert -> Ordinal)
            if measure_type == 'Likert':
                measure_type = 'Ordinal'
            config['encoding_type'] = measure_type

        with col_b:
            start_value = st.number_input(
                "Start Value",
                min_value=0,
                value=config['start_value'],
                key=f"start_{col_name}",
                help="Starting numeric code"
            )
            config['start_value'] = int(start_value)

        with col_c:
            direction = st.selectbox(
                "Direction",
                options=['Ascending', 'Descending'],
                index=['Ascending', 'Descending'].index(config['direction']),
                key=f"dir_{col_name}",
                help="Ascending: first→smallest | Descending: first→largest"
            )
            config['direction'] = direction

        # Missing value handling
        treat_missing = st.checkbox(
            "Treat missing/blank as system-missing (leave blank)",
            value=config['treat_missing'],
            key=f"missing_{col_name}"
        )
        config['treat_missing'] = treat_missing

        if measure_type != 'Ignore':
            st.markdown("---")
            st.markdown("**Reorder Options** (drag with ↑ ↓ buttons)")

            # Reordering UI
            current_order = st.session_state.column_orders[col_name]

            for idx, value in enumerate(current_order):
                col_left, col_mid, col_right = st.columns([1, 6, 1])

                with col_left:
                    if st.button("↑", key=f"up_{col_name}_{idx}", disabled=idx == 0):
                        move_option_up(col_name, idx)
                        st.rerun()

                with col_mid:
                    st.text(f"{idx+1}. {value}")

                with col_right:
                    if st.button("↓", key=f"down_{col_name}_{idx}", disabled=idx == len(current_order)-1):
                        move_option_down(col_name, idx)
                        st.rerun()

            # Show preview of mapping
            st.markdown("**Mapping Preview:**")
            preview_config = ColumnConfig(
                column_name=col_name,
                unique_values=current_order,
                encoding_type=measure_type,
                start_value=config['start_value'],
                direction=config['direction']
            )
            mapping = preview_config.get_mapping()

            preview_lines = [
                f"  {value} → **{code}**" for value, code in mapping.items()]
            st.markdown('\n'.join(preview_lines[:5]))
            if len(preview_lines) > 5:
                st.caption(f"... and {len(preview_lines) - 5} more")


def main():
    """Main application logic."""

    # Render navbar
    render_navbar()

    # Render sidebar and get settings
    include_save, write_same_folder, sanitize_names = render_sidebar()

    # Main content

    # Step 1: Upload
    st.header("Step 1: Upload Excel File")

    uploaded_file = st.file_uploader(
        "Upload your Google Forms export (.xlsx)",
        type=['xlsx'],
        help="Upload an Excel file exported from Google Forms"
    )

    if uploaded_file is not None:
        st.session_state.uploaded_file = uploaded_file

        try:
            # Read Excel
            df = pd.read_excel(uploaded_file, dtype=object)
            st.session_state.df = df

            st.success(f"✅ Loaded {len(df)} rows × {len(df.columns)} columns")

            # Show preview
            st.subheader("Data Preview (first 5 rows)")
            st.dataframe(df.head(), use_container_width=True)

            # Detect columns
            if not st.session_state.column_info:
                with st.spinner("Detecting columns..."):
                    column_info = detect_columns(df)
                    st.session_state.column_info = column_info

                    # Generate unique variable names (handles Arabic and duplicates)
                    if sanitize_names:
                        unique_names = generate_unique_var_names(
                            list(df.columns))
                        st.session_state.unique_var_names = unique_names

            st.markdown("---")

            # Step 2: Configure columns
            st.header("Step 2: Configure Column Encodings")

            for col_name in df.columns:
                col_info = st.session_state.column_info[col_name]
                render_column_card(col_name, col_info, sanitize_names)

            st.markdown("---")

            # Step 3: Apply encoding
            st.header("Step 3: Apply Encoding & Generate Files")

            # Check if at least one column is configured
            non_ignored = [
                c for c, cfg in st.session_state.column_configs.items()
                if cfg.get('encoding_type') != 'Ignore'
            ]

            if st.button("🚀 Apply Encoding & Generate Files",
                         type="primary",
                         disabled=len(non_ignored) == 0):

                with st.spinner("Encoding data and generating files..."):
                    # Build ColumnConfig objects
                    configs = {}
                    for col_name, cfg in st.session_state.column_configs.items():
                        configs[col_name] = ColumnConfig(
                            column_name=col_name,
                            unique_values=st.session_state.column_orders[col_name],
                            encoding_type=cfg['encoding_type'],
                            start_value=cfg['start_value'],
                            direction=cfg['direction'],
                            treat_missing=cfg['treat_missing'],
                            sanitized_name=cfg['sanitized_name']
                        )

                    # Apply encoding
                    encoded_df, mappings = apply_encoding(df, configs)
                    st.session_state.encoded_df = encoded_df

                    # Rename columns if sanitized
                    if sanitize_names:
                        rename_map = {
                            col: configs[col].sanitized_name
                            for col in df.columns
                            if col in configs
                        }
                        encoded_df = encoded_df.rename(columns=rename_map)
                        # Update mappings keys to sanitized names
                        mappings = {
                            configs[k].sanitized_name: v
                            for k, v in mappings.items()
                        }

                    # Save encoded Excel
                    temp_dir = tempfile.gettempdir()
                    encoded_path = os.path.join(temp_dir, 'encoded_data.xlsx')
                    save_encoded_excel(encoded_df, encoded_path)
                    st.session_state.encoded_path = encoded_path

                    # Build original names mapping for VARIABLE LABELS
                    original_names = {
                        configs[col].sanitized_name: col
                        for col in df.columns
                        if col in configs
                    }

                    # Build measure types mapping for VARIABLE LEVEL
                    measure_types = {
                        configs[col].sanitized_name: configs[col].encoding_type
                        for col in df.columns
                        if col in configs
                    }

                    # Generate SPSS syntax
                    save_path = encoded_path.replace(
                        '.xlsx', '.sav') if include_save else None
                    sps_syntax = generate_sps_syntax(
                        excel_path=encoded_path,
                        mappings=mappings,
                        original_names=original_names,
                        sheet_name='Sheet1',
                        include_save=include_save,
                        save_path=save_path,
                        use_relative_path=True,  # Use relative path for downloaded files
                        measure_types=measure_types  # Set SPSS variable levels
                    )
                    st.session_state.sps_syntax = sps_syntax

                    # Save .sps file
                    if write_same_folder:
                        sps_path = encoded_path.replace('.xlsx', '.sps')
                    else:
                        sps_path = os.path.join(temp_dir, 'auto_import.sps')

                    save_sps_file(sps_syntax, sps_path)
                    st.session_state.sps_path = sps_path

                st.success("✅ Files generated successfully!")

            # Step 4: Preview and Download
            if st.session_state.sps_syntax:
                st.markdown("---")
                st.header("Step 4: Preview & Download")

                col1, col2 = st.columns(2)

                with col1:
                    st.subheader("📄 Encoded Data Preview")
                    st.dataframe(st.session_state.encoded_df.head(),
                                 use_container_width=True)

                    # Download encoded Excel
                    with open(st.session_state.encoded_path, 'rb') as f:
                        st.download_button(
                            label="⬇️ Download Encoded Excel",
                            data=f.read(),
                            file_name="encoded_data.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                with col2:
                    st.subheader("📜 SPSS Syntax Preview")
                    st.code(st.session_state.sps_syntax, language='sql')

                    # Download .sps file with UTF-8 BOM for Arabic text support
                    sps_bytes = st.session_state.sps_syntax.encode('utf-8-sig')
                    st.download_button(
                        label="⬇️ Download SPSS Syntax (.sps)",
                        data=sps_bytes,
                        file_name="auto_import.sps",
                        mime="text/plain"
                    )

                st.success("📥 **Files ready for download!**")
                st.markdown("### 📋 Next Steps:")
                st.markdown("""
                1. **Download both files** using the buttons above
                2. **Save them to the SAME folder** on your computer
                3. **Open the .sps file in a text editor** (Notepad, etc.)
                4. **Edit the CD command** to point to your folder:
                   ```
                   CD 'C:\\Users\\YourName\\Documents\\MySurvey'.
                   ```
                   Replace with your actual folder path!
                5. **Save the .sps file**
                6. **Open the .sps file in IBM SPSS**
                7. **Run the script** (Ctrl+A to select all, then Ctrl+R to run)
                
                ⚠️ **Important:** The CD command tells SPSS where to find your Excel file. If you skip step 4, you'll get Error 2052!
                """)

                st.info(
                    "💡 **Tip:** Keep both files together and remember which folder you saved them in.")

        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            logger.error(f"Error: {str(e)}", exc_info=True)

    else:
        st.info("👆 Upload an Excel file to get started")


if __name__ == "__main__":
    main()
