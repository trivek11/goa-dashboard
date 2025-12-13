# -------------------------------------------------
# 0. IMPORTS
# -------------------------------------------------
import pandas as pd
import plotly.graph_objects as go
from dash import Dash, html, dcc, Input, Output, ALL, dash_table, State
import dash
from datetime import datetime
import re
import json

# -------------------------------------------------
# 1. MASTER COUNCIL LIST (North Goa static)
# -------------------------------------------------
north_goa_councils = [
    {"District": "North-Goa", "Council": "CCP", "Total Wards": 23, "Existing Properties": 30991},
    {"District": "North-Goa", "Council": "Mapusa", "Total Wards": 20, "Existing Properties": 29654},
    {"District": "North-Goa", "Council": "Bicholim", "Total Wards": 14, "Existing Properties": 8821},
    {"District": "North-Goa", "Council": "Pernem", "Total Wards": 10, "Existing Properties": 1794},
    {"District": "North-Goa", "Council": "Valpoi", "Total Wards": 10, "Existing Properties": 4552},
    {"District": "North-Goa", "Council": "Sanquelim", "Total Wards": 12, "Existing Properties": 6337},
]


# -------------------------------------------------
# 2. WARD DATA FACTORY FOR NORTH GOA
# -------------------------------------------------
def generate_ward_data(council_name, total_wards, existing_properties):
    surveyed_tot = int(existing_properties * 0.92)
    suspected_tot = int(existing_properties * 0.18)
    old_gsuda_tot = int(existing_properties * 0.75)
    ext_area_tot = int(existing_properties * 0.05)
    objections_tot = max(1, total_wards // 2)

    def split(total, wards):
        base, rem = divmod(total, wards)
        return [base + 1 if i < rem else base for i in range(wards)]

    upic_lst, survey_lst = split(existing_properties, total_wards), split(surveyed_tot, total_wards)
    suspected_lst, old_lst = split(suspected_tot, total_wards), split(old_gsuda_tot, total_wards)
    ext_lst, obj_lst = split(ext_area_tot, total_wards), split(objections_tot, total_wards)

    wards = []
    for i in range(1, total_wards + 1):
        wards.append({
            "Sr No": i,
            "Ward No": i,
            "UPIC": upic_lst[i - 1],
            "Survey": survey_lst[i - 1],
            "Old GSUDA Properties": old_lst[i - 1],
            "Suspected New Properties": suspected_lst[i - 1],
            "Extended Area Properties": ext_lst[i - 1],
            "Objections Received": obj_lst[i - 1],
            "Council": council_name,
            "District": "North-Goa"
        })
    return pd.DataFrame(wards)


# Generate North Goa ward data
north_goa_wards = pd.concat([
    generate_ward_data(c["Council"], c["Total Wards"], c["Existing Properties"])
    for c in north_goa_councils
], ignore_index=True)

# -------------------------------------------------
# 3. LOAD SOUTH GOA DATA FROM EXCEL
# -------------------------------------------------
import os

# For local development and deployment
file_path = os.path.join(os.path.dirname(__file__), "data", "South Goa.xlsx")

# Or if you want to check multiple locations:
if not os.path.exists(file_path):
    # Try alternative location
    file_path = "South Goa.xlsx"
    if not os.path.exists(file_path):
        file_path = os.path.join("data", "South Goa.xlsx")

try:
    # Load Excel
    xls = pd.ExcelFile(file_path)
    print("Sheets in Excel:", xls.sheet_names)

    # Load summary sheet (first sheet)
    south_goa_summary = pd.read_excel(file_path, sheet_name=xls.sheet_names[0])
    print("Summary sheet columns:", south_goa_summary.columns.tolist())


    # Function to clean and standardize column names
    def clean_column_name(col_name):
        if pd.isna(col_name):
            return "Unnamed"

        if not isinstance(col_name, str):
            col_name = str(col_name)

        col_name = col_name.strip()

        # Standardize Sr No variations
        if re.search(r'^sr\.?\s*no\.?$', col_name.lower()):
            return "Sr No"

        # Standardize Ward No variations
        if re.search(r'^ward\.?\s*no\.?$', col_name.lower()):
            return "Ward No"

        # Standardize specific column names exactly as they appear in your data
        col_lower = col_name.lower()

        # Handle exact matches first
        exact_matches = {
            "sr. no": "Sr No",
            "ward no": "Ward No",
            "ward boundary mapping": "Ward Boundary Mapping",
            "digitlisation of polygon": "Digitlisation Of Polygon",
            "generation of upic number": "Generation of UPIC Number",
            "upic": "UPIC",
            "nic property": "NIC Property",
            "suspected new": "Suspected New",
            "total survey": "Total Survey",
            "remark": "Remark"
        }

        for pattern, standard_name in exact_matches.items():
            if pattern == col_lower:
                return standard_name

        # Handle partial matches
        if 'upic' in col_lower:
            return "UPIC"
        elif 'survey' in col_lower and 'total' in col_lower:
            return "Total Survey"
        elif 'survey' in col_lower:
            return "Survey"
        elif 'nic' in col_lower and 'property' in col_lower:
            return "NIC Property"
        elif 'suspected' in col_lower and 'new' in col_lower:
            return "Suspected New"
        elif 'remark' in col_lower:
            return "Remark"
        elif 'ward boundary' in col_lower:
            return "Ward Boundary Mapping"
        elif 'digitlisation' in col_lower or 'digitization' in col_lower:
            return "Digitlisation Of Polygon"
        elif 'generation' in col_lower and 'upic' in col_lower:
            return "Generation of UPIC Number"
        elif 'old' in col_lower and 'gsuda' in col_lower:
            return "Old GSUDA Properties"
        elif 'extended' in col_lower and 'area' in col_lower:
            return "Extended Area Properties"
        elif 'objections' in col_lower:
            return "Objections Received"

        # Remove any "Unnamed" columns
        if 'unnamed' in col_lower:
            return None

        return col_name


    # Dictionary to store council-wise column information
    council_columns_info = {}

    # Load ward sheets from all other sheets
    ward_sheets = []
    sheet_council_names = {}

    for sheet_name in xls.sheet_names[1:]:  # Skip first summary sheet
        try:
            print(f"\n{'=' * 60}")
            print(f"Processing sheet: {sheet_name}")
            print(f"{'=' * 60}")

            # Read the sheet
            df_sheet = pd.read_excel(file_path, sheet_name=sheet_name)

            # Store original columns
            original_columns = df_sheet.columns.tolist()
            print(f"Original columns ({len(original_columns)}): {original_columns}")

            # Clean column names
            cleaned_columns = [clean_column_name(col) for col in original_columns]

            # Apply cleaned column names
            df_sheet.columns = cleaned_columns

            # Remove None columns (Unnamed columns)
            columns_to_keep = [col for col in cleaned_columns if col is not None]
            df_sheet = df_sheet[columns_to_keep]

            # Remove duplicate columns
            df_sheet = df_sheet.loc[:, ~df_sheet.columns.duplicated()]

            # Store column info for this council
            council_columns_info[sheet_name] = {
                'original': original_columns,
                'cleaned': df_sheet.columns.tolist()
            }

            print(f"Cleaned columns ({len(df_sheet.columns)}): {df_sheet.columns.tolist()}")
            print(f"First few rows:")
            print(df_sheet.head(3))

            # Add council and district columns
            df_sheet["Council"] = sheet_name
            df_sheet["District"] = "South-Goa"

            # Store the sheet name
            sheet_council_names[sheet_name] = sheet_name

            ward_sheets.append(df_sheet)

        except Exception as e:
            print(f"Error loading sheet {sheet_name}: {e}")
            import traceback

            traceback.print_exc()

    # Combine all ward data
    if ward_sheets:
        south_goa_wards = pd.concat(ward_sheets, ignore_index=True)
    else:
        print("No ward sheets found")
        south_goa_wards = pd.DataFrame()

    # -------------------------------------------------
    # CREATE SOUTH GOA COUNCIL LIST FROM SUMMARY SHEET
    # -------------------------------------------------
    south_goa_councils = []

    # Clean column names in summary
    south_goa_summary.columns = [str(col).strip() for col in south_goa_summary.columns]

    # Find columns
    council_col = "Council" if "Council" in south_goa_summary.columns else south_goa_summary.columns[0]

    # Find NIC Record Property column
    nic_property_col = None
    order_date_col = None

    for col in south_goa_summary.columns:
        col_lower = str(col).lower()
        if "nic record property" in col_lower:
            nic_property_col = col
        elif "order issue date" in col_lower:
            order_date_col = col

    # Format Order Issue Date to date only (without time)
    if order_date_col and order_date_col in south_goa_summary.columns:
        def format_date(x):
            if pd.isna(x):
                return ""
            try:
                if isinstance(x, datetime):
                    return x.strftime('%d-%m-%Y')
                else:
                    # Try to parse the date
                    parsed_date = pd.to_datetime(x, errors='coerce')
                    if pd.isna(parsed_date):
                        # Try different date formats
                        for fmt in ['%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d', '%Y/%m/%d']:
                            try:
                                parsed_date = datetime.strptime(str(x), fmt)
                                break
                            except:
                                continue
                    return parsed_date.strftime('%d-%m-%Y') if not pd.isna(parsed_date) else str(x)
            except Exception as e:
                print(f"Error formatting date {x}: {e}")
                return str(x)


        south_goa_summary[order_date_col] = south_goa_summary[order_date_col].apply(format_date)

    for idx, row in south_goa_summary.iterrows():
        council_name = str(row[council_col]).strip() if pd.notna(row[council_col]) else ""

        # Skip TOTAL row and empty rows
        if not council_name or council_name.upper() == "TOTAL":
            continue

        # Get NIC Record Property count
        nic_property_count = 0
        if nic_property_col and nic_property_col in row:
            try:
                nic_property_count = int(float(row[nic_property_col]))
            except:
                nic_property_count = 0

        # Get Order Issue Date
        order_issue_date = ""
        if order_date_col and order_date_col in row and pd.notna(row[order_date_col]):
            order_issue_date = str(row[order_date_col])

        south_goa_councils.append({
            "District": "South-Goa",
            "Council": council_name,
            "NIC Record Properties": nic_property_count,
            "Order Issue Date": order_issue_date
        })

    print(f"\nCreated {len(south_goa_councils)} councils for South Goa")

except Exception as e:
    print(f"Error loading Excel file: {e}")
    import traceback

    traceback.print_exc()
    # Create empty DataFrames if file loading fails
    south_goa_summary = pd.DataFrame()
    south_goa_wards = pd.DataFrame()
    south_goa_councils = []

# -------------------------------------------------
# 4. CREATE DASH APP
# -------------------------------------------------
app = Dash(__name__, suppress_callback_exceptions=True)


# -------------------------------------------------
# 5. PAGE LAYOUTS
# -------------------------------------------------

def create_pie_chart(district, councils):
    if district == "North-Goa":
        df = pd.DataFrame(councils)
        prop_col = "Existing Properties"
        label_col = "Council"
        title = f"{district} – Property Distribution"
    elif district == "South-Goa":
        df = pd.DataFrame(councils)
        if df.empty:
            return html.Div([
                html.H3("South-Goa", style={"textAlign": "center"}),
                html.P("No council data available")
            ])

        prop_col = "NIC Record Properties"
        label_col = "Council"
        title = f"{district} – NIC Record Properties Distribution"
    else:
        return html.Div("No data available")

    # Calculate percentages
    total = df[prop_col].sum()

    fig = go.Figure()
    fig.add_trace(go.Pie(
        labels=df[label_col],
        values=df[prop_col],
        hole=0.55,
        textinfo="label+percent",
        hovertemplate="<b>%{label}</b><br>Properties: %{value:,}<br>%{percent}",
        marker=dict(line=dict(color="white", width=2))
    ))
    fig.update_layout(
        title=title,
        margin=dict(l=20, r=20, t=40, b=20),
        height=350,
        showlegend=False
    )

    return html.Div([
        html.H3(district, style={"textAlign": "center"}),
        dcc.Graph(
            figure=fig,
            id={"type": "pie", "index": district},
            style={"height": "380px", "overflow": "hidden"}
        )
    ])


def create_column_chart(district, councils):
    if district == "North-Goa":
        df = pd.DataFrame(councils)
        prop_col = "Existing Properties"
        title = f"{district} – Council-wise Properties"
    elif district == "South-Goa":
        df = pd.DataFrame(councils)
        if df.empty:
            return html.Div()
        prop_col = "NIC Record Properties"
        title = f"{district} – Council-wise NIC Properties"
    else:
        return html.Div()

    # Sort by properties
    df = df.sort_values(prop_col, ascending=False)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df["Council"],
        y=df[prop_col],
        text=df[prop_col],
        textposition='auto',
        marker_color='#3498db',
        hovertemplate="<b>%{x}</b><br>Properties: %{y:,}<extra></extra>"
    ))

    fig.update_layout(
        title=title,
        xaxis_title="Council",
        yaxis_title="Number of Properties",
        height=400,
        margin=dict(l=20, r=20, t=40, b=40),
        xaxis={'categoryorder': 'total descending'}
    )

    return html.Div([
        dcc.Graph(figure=fig)
    ])


def dashboard_page():
    return html.Div([
        html.H1("Dashboard Overview", style={"textAlign": "center", "color": "#2c3e50", "marginBottom": "30px"}),

        # Pie Charts Row
        html.Div([
            html.Div([
                create_pie_chart("North-Goa", north_goa_councils)
            ], style={"width": "48%", "display": "inline-block", "verticalAlign": "top"}),

            html.Div([
                create_pie_chart("South-Goa", south_goa_councils)
            ], style={"width": "48%", "display": "inline-block", "verticalAlign": "top"})
        ], style={"marginBottom": "40px"}),

        # Column Charts Row
        html.Div([
            html.Div([
                create_column_chart("North-Goa", north_goa_councils)
            ], style={"width": "48%", "display": "inline-block", "verticalAlign": "top"}),

            html.Div([
                create_column_chart("South-Goa", south_goa_councils)
            ], style={"width": "48%", "display": "inline-block", "verticalAlign": "top"})
        ])
    ])


def north_goa_page():
    # Calculate statistics
    total_properties = sum(c["Existing Properties"] for c in north_goa_councils)
    total_wards = sum(c["Total Wards"] for c in north_goa_councils)

    return html.Div([
        html.H1("North Goa Councils", style={"textAlign": "center", "color": "#2c3e50", "marginBottom": "20px"}),

        # Statistics
        html.Div([
            html.Div([
                html.H4(f"{total_properties:,}", style={"color": "#3498db", "margin": "0"}),
                html.P("Total Properties", style={"color": "#7f8c8d", "margin": "0"})
            ], style={"textAlign": "center", "padding": "20px", "backgroundColor": "white", "borderRadius": "10px",
                      "boxShadow": "0 2px 5px rgba(0,0,0,0.1)", "margin": "10px"}),

            html.Div([
                html.H4(f"{total_wards}", style={"color": "#2ecc71", "margin": "0"}),
                html.P("Total Wards", style={"color": "#7f8c8d", "margin": "0"})
            ], style={"textAlign": "center", "padding": "20px", "backgroundColor": "white", "borderRadius": "10px",
                      "boxShadow": "0 2px 5px rgba(0,0,0,0.1)", "margin": "10px"}),

            html.Div([
                html.H4(f"{len(north_goa_councils)}", style={"color": "#e74c3c", "margin": "0"}),
                html.P("Total Councils", style={"color": "#7f8c8d", "margin": "0"})
            ], style={"textAlign": "center", "padding": "20px", "backgroundColor": "white", "borderRadius": "10px",
                      "boxShadow": "0 2px 5px rgba(0,0,0,0.1)", "margin": "10px"})
        ], style={"display": "flex", "justifyContent": "center", "marginBottom": "30px"}),

        # Council Cards
        html.Div([
            html.Div([
                html.H4(council["Council"], style={"color": "#2c3e50", "marginBottom": "10px"}),
                html.P(f"Wards: {council['Total Wards']}", style={"color": "#7f8c8d", "margin": "5px 0"}),
                html.P(f"Properties: {council['Existing Properties']:,}",
                       style={"color": "#7f8c8d", "margin": "5px 0"}),
                html.Button("View Ward Details",
                            id={"type": "council-button", "index": council["Council"], "district": "North-Goa"},
                            n_clicks=0,
                            style={
                                "backgroundColor": "#3498db",
                                "color": "white",
                                "border": "none",
                                "padding": "10px 20px",
                                "borderRadius": "5px",
                                "cursor": "pointer",
                                "marginTop": "10px",
                                "width": "100%"
                            })
            ], style={
                "backgroundColor": "white",
                "padding": "20px",
                "borderRadius": "10px",
                "boxShadow": "0 2px 10px rgba(0,0,0,0.1)",
                "margin": "10px",
                "width": "calc(33.333% - 20px)",
                "minWidth": "250px"
            })
            for council in north_goa_councils
        ], style={"display": "flex", "flexWrap": "wrap", "justifyContent": "center"}),

        # Ward Details Area (hidden initially)
        html.Div(id="north-goa-ward-details", style={"marginTop": "40px"})
    ])


def south_goa_page():
    # Calculate statistics
    total_properties = sum(c.get("NIC Record Properties", 0) for c in south_goa_councils)

    return html.Div([
        html.H1("South Goa Councils", style={"textAlign": "center", "color": "#2c3e50", "marginBottom": "20px"}),

        # Statistics
        html.Div([
            html.Div([
                html.H4(f"{total_properties:,}", style={"color": "#3498db", "margin": "0"}),
                html.P("Total NIC Properties", style={"color": "#7f8c8d", "margin": "0"})
            ], style={"textAlign": "center", "padding": "20px", "backgroundColor": "white", "borderRadius": "10px",
                      "boxShadow": "0 2px 5px rgba(0,0,0,0.1)", "margin": "10px"}),

            html.Div([
                html.H4(f"{len(south_goa_councils)}", style={"color": "#e74c3c", "margin": "0"}),
                html.P("Total Councils", style={"color": "#7f8c8d", "margin": "0"})
            ], style={"textAlign": "center", "padding": "20px", "backgroundColor": "white", "borderRadius": "10px",
                      "boxShadow": "0 2px 5px rgba(0,0,0,0.1)", "margin": "10px"}),
        ], style={"display": "flex", "justifyContent": "center", "marginBottom": "30px"}),

        # Council Cards
        html.Div([
            html.Div([
                html.H4(council["Council"], style={"color": "#2c3e50", "marginBottom": "10px"}),
                html.P(f"NIC Properties: {council.get('NIC Record Properties', 0):,}",
                       style={"color": "#7f8c8d", "margin": "5px 0"}),
                html.P(f"Order Date: {council.get('Order Issue Date', 'N/A')}",
                       style={"color": "#7f8c8d", "margin": "5px 0"}),
                html.Button("View Ward Details",
                            id={"type": "council-button", "index": council["Council"], "district": "South-Goa"},
                            n_clicks=0,
                            style={
                                "backgroundColor": "#3498db",
                                "color": "white",
                                "border": "none",
                                "padding": "10px 20px",
                                "borderRadius": "5px",
                                "cursor": "pointer",
                                "marginTop": "10px",
                                "width": "100%"
                            })
            ], style={
                "backgroundColor": "white",
                "padding": "20px",
                "borderRadius": "10px",
                "boxShadow": "0 2px 10px rgba(0,0,0,0.1)",
                "margin": "10px",
                "width": "calc(33.333% - 20px)",
                "minWidth": "250px"
            })
            for council in south_goa_councils
        ], style={"display": "flex", "flexWrap": "wrap", "justifyContent": "center"}),

        # Ward Details Area (hidden initially)
        html.Div(id="south-goa-ward-details", style={"marginTop": "40px"})
    ])


def summary_page():
    return html.Div([
        html.H1("South Goa Summary Report", style={"textAlign": "center", "color": "#2c3e50", "marginBottom": "30px"}),

        dash_table.DataTable(
            data=south_goa_summary.to_dict("records") if not south_goa_summary.empty else [],
            columns=[{"name": str(col), "id": str(col)} for col in
                     south_goa_summary.columns] if not south_goa_summary.empty else [],
            page_size=20,
            style_table={"overflowX": "auto", "borderRadius": "10px", "boxShadow": "0 2px 10px rgba(0,0,0,0.1)"},
            style_cell={"textAlign": "left", "padding": "10px", "border": "1px solid #eee"},
            style_header={'backgroundColor': '#3498db', 'fontWeight': 'bold', 'color': 'white'},
            style_data_conditional=[
                {'if': {'row_index': 'odd'}, 'backgroundColor': '#f8f9fa'}
            ],
        )
    ])


def reports_page():
    return html.Div([
        html.H1("Reports", style={"textAlign": "center", "color": "#2c3e50", "marginBottom": "30px"}),
        html.P("Reports section will be implemented soon.", style={"textAlign": "center", "color": "#7f8c8d"})
    ])


def settings_page():
    return html.Div([
        html.H1("Settings", style={"textAlign": "center", "color": "#2c3e50", "marginBottom": "30px"}),
        html.P("Settings section will be implemented soon.", style={"textAlign": "center", "color": "#7f8c8d"})
    ])


# -------------------------------------------------
# 6. WARD DETAILS FUNCTION (Reusable)
# -------------------------------------------------
def get_ward_details(council_name, district):
    """Get ward details for a specific council"""
    if district == "North-Goa":
        council_name = council_name

        if council_name not in north_goa_wards["Council"].unique():
            return None

        ward_data = north_goa_wards[north_goa_wards["Council"] == council_name].copy()

        existing_properties = 0
        for council in north_goa_councils:
            if council["Council"] == council_name:
                existing_properties = council.get("Existing Properties", 0)
                break

    elif district == "South-Goa":
        council_name = council_name

        if not south_goa_wards.empty and "Council" in south_goa_wards.columns:
            if council_name not in south_goa_wards["Council"].unique():
                return None

            ward_data = south_goa_wards[south_goa_wards["Council"] == council_name].copy()

            # Remove "TOTAL" rows and any rows that might be summaries
            total_mask = pd.Series([False] * len(ward_data))

            for col in ward_data.columns:
                if col and isinstance(col, str):
                    col_vals = ward_data[col].astype(str).str.upper().str.strip()
                    total_mask = total_mask | (col_vals == "TOTAL")
                    total_mask = total_mask | (col_vals == "TOTAL:")
                    total_mask = total_mask | (col_vals == "GRAND TOTAL")

            if total_mask.any():
                ward_data = ward_data[~total_mask].copy()

        else:
            return None

        # Get NIC Record Property and other details
        nic_property_count = 0
        order_issue_date = ""
        for council in south_goa_councils:
            if council["Council"] == council_name:
                nic_property_count = council.get("NIC Record Properties", 0)
                order_issue_date = council.get("Order Issue Date", "")
                break

    # Check if ward_data is empty
    if ward_data.empty:
        return None

    # Clean column names and reset index
    ward_data = ward_data.loc[:, ~ward_data.columns.duplicated()].reset_index(drop=True)

    # Ensure Sr No and Ward No are properly set
    if "Sr No" in ward_data.columns:
        ward_data["Sr No"] = range(1, len(ward_data) + 1)

    if "Ward No" in ward_data.columns:
        try:
            ward_data["Ward No"] = pd.to_numeric(ward_data["Ward No"], errors='coerce')
        except:
            pass

    # Get available columns
    available_cols = [col for col in ward_data.columns if col not in ["Council", "District"]]

    # Define preferred order for columns
    if district == "North-Goa":
        preferred_order = [
            "Sr No", "Ward No", "UPIC", "Survey",
            "Old GSUDA Properties", "Suspected New Properties",
            "Extended Area Properties", "Objections Received"
        ]
        numeric_cols = ["UPIC", "Survey", "Old GSUDA Properties",
                        "Suspected New Properties", "Extended Area Properties",
                        "Objections Received"]
    else:
        preferred_order = [
            "Sr No", "Ward No",
            "Ward Boundary Mapping", "Digitlisation Of Polygon", "Generation of UPIC Number",
            "UPIC", "NIC Property", "Suspected New", "Total Survey", "Survey",
            "Old GSUDA Properties", "Extended Area Properties", "Objections Received",
            "Remark"
        ]
        numeric_cols = ["UPIC", "NIC Property", "Suspected New", "Total Survey", "Survey",
                        "Old GSUDA Properties", "Extended Area Properties", "Objections Received"]

    # Sort columns
    sorted_cols = []
    for col in preferred_order:
        if col in available_cols:
            sorted_cols.append(col)

    for col in available_cols:
        if col not in sorted_cols:
            sorted_cols.append(col)

    # Calculate total wards
    total_wards = len(ward_data)

    # Prepare data for table
    table_data = ward_data[sorted_cols].to_dict("records")

    # Add Total Row
    if len(table_data) > 0:
        total_row = {}
        for col in sorted_cols:
            if col in numeric_cols and col in ward_data.columns:
                try:
                    numeric_values = pd.to_numeric(ward_data[col], errors='coerce')
                    total_sum = numeric_values.sum(skipna=True)
                    if pd.notna(total_sum):
                        if total_sum % 1 == 0:
                            total_row[col] = int(total_sum)
                        else:
                            total_row[col] = round(total_sum, 2)
                    else:
                        total_row[col] = 0
                except Exception as e:
                    total_row[col] = ""
            elif col == "Sr No":
                total_row[col] = "TOTAL"
            elif col == "Ward No":
                total_row[col] = ""
            elif col == "Remark":
                total_row[col] = ""
            elif col in ["Ward Boundary Mapping", "Digitlisation Of Polygon", "Generation of UPIC Number"]:
                if district == "South-Goa" and col in ward_data.columns:
                    non_empty = ward_data[col].astype(str).str.strip().ne('').sum()
                    total_row[col] = f"{non_empty}/{total_wards}"
                else:
                    total_row[col] = ""
            else:
                total_row[col] = ""

        table_data.append(total_row)

    # Create table
    table = dash_table.DataTable(
        columns=[{"name": col, "id": col} for col in sorted_cols],
        data=table_data,
        page_size=len(table_data),
        style_table={
            "overflowX": "scroll",
            "overflowY": "auto",
            "border": "2px solid #bdc3c7",
            "borderRadius": "5px",
            "minWidth": "100%",
            "width": "100%",
            "maxHeight": "600px",
        },
        style_cell={
            "textAlign": "center",
            "padding": "10px",
            "border": "1px solid #ecf0f1",
            "minWidth": "120px",
            "maxWidth": "200px",
            "whiteSpace": "normal",
            "fontFamily": "Arial, sans-serif",
            "overflow": "hidden",
            "textOverflow": "ellipsis",
        },
        style_header={
            'backgroundColor': '#3498db',
            'fontWeight': 'bold',
            'border': '1px solid #2980b9',
            'color': 'white',
            'fontSize': '14px',
            'position': 'sticky',
            'top': 0,
            'zIndex': 1,
        },
        style_data={
            'whiteSpace': 'normal',
            'height': 'auto',
            'fontSize': '13px',
            'minWidth': '120px',
        },
        style_data_conditional=[
            {
                'if': {'row_index': 'odd'},
                'backgroundColor': '#f8f9fa'
            },
            {
                'if': {'column_id': 'Sr No'},
                'fontWeight': 'bold',
                'backgroundColor': '#e8f4fc',
            },
            {
                'if': {'column_id': 'Ward No'},
                'fontWeight': 'bold',
                'backgroundColor': '#e8f4fc',
            },
            {
                'if': {'row_index': len(table_data) - 1},
                'backgroundColor': '#2ecc71',
                'color': 'white',
                'fontWeight': 'bold',
                'fontSize': '14px',
                'borderTop': '3px solid #27ae60'
            },
        ],
        fixed_rows={'headers': True},
        fixed_columns={'headers': True, 'data': 1},
        virtualization=False,
    )

    # Create header info - WITHOUT BACK BUTTON
    if district == "North-Goa":
        header_info = html.Div([
            html.H3(f"{council_name} - Ward Details ({district})",
                    style={"color": "#2c3e50", "marginBottom": "15px"}),
            html.Div([
                html.Span(f"Total Wards: {total_wards}",
                          style={"fontWeight": "bold", "marginRight": "20px"}),
                html.Span(f"Existing Properties: {existing_properties:,}",
                          style={"fontWeight": "bold", "marginRight": "20px"})
            ], style={"marginBottom": "20px", "fontSize": "16px"})
        ])
    else:
        header_info = html.Div([
            html.H3(f"{council_name} - Ward Details ({district})",
                    style={"color": "#2c3e50", "marginBottom": "15px"}),
            html.Div([
                html.Span(f"Total Wards: {total_wards}",
                          style={"fontWeight": "bold", "marginRight": "20px"}),
                html.Span(f"NIC Record Properties: {nic_property_count:,}",
                          style={"fontWeight": "bold", "marginRight": "20px"}),
                html.Span(f"Order Issue Date: {order_issue_date}",
                          style={"fontWeight": "bold"}) if order_issue_date else ""
            ], style={"marginBottom": "20px", "fontSize": "16px"})
        ])

    return html.Div([
        header_info,
        html.H5(f"Showing {total_wards} wards + Total Row",
                style={"marginBottom": "10px", "color": "#7f8c8d"}),
        html.Div([
            table
        ], style={
            "width": "100%",
            "overflowX": "auto",
            "border": "1px solid #ddd",
            "borderRadius": "5px",
            "padding": "10px",
            "backgroundColor": "#fff"
        })
    ], style={
        "padding": "20px",
        "backgroundColor": "white",
        "borderRadius": "10px",
        "boxShadow": "0 2px 10px rgba(0,0,0,0.1)",
        "width": "100%",
        "maxWidth": "100%",
        "overflow": "hidden"
    })

# -------------------------------------------------
# 7. CREATE SIDEBAR COMPONENT
# -------------------------------------------------
def create_sidebar():
    # Calculate statistics
    total_north_properties = sum(c["Existing Properties"] for c in north_goa_councils)
    total_south_properties = sum(c.get("NIC Record Properties", 0) for c in south_goa_councils)
    total_councils = len(north_goa_councils) + len(south_goa_councils)

    total_north_wards = sum(c["Total Wards"] for c in north_goa_councils)
    if not south_goa_wards.empty:
        total_south_wards = len(south_goa_wards)
    else:
        total_south_wards = 0

    return html.Div([
        # Sidebar Toggle Button
        html.Div([
            html.Button(
                "☰",
                id="sidebar-toggle",
                className="toggle-btn",
                n_clicks=0
            ),
            html.H3("Goa Dashboard", className="sidebar-title")
        ], className="sidebar-header"),

        # Navigation Menu
        html.Ul([
            html.Li([
                html.Button([
                    html.Span("🏠", className="nav-icon"),
                    html.Span("Dashboard", className="nav-text")
                ], id="nav-dashboard", className="nav-link active")
            ], className="nav-item"),

            html.Li([
                html.Button([
                    html.Span("📊", className="nav-icon"),
                    html.Span("North Goa", className="nav-text")
                ], id="nav-north-goa", className="nav-link")
            ], className="nav-item"),

            html.Li([
                html.Button([
                    html.Span("📈", className="nav-icon"),
                    html.Span("South Goa", className="nav-text")
                ], id="nav-south-goa", className="nav-link")
            ], className="nav-item"),

            html.Li([
                html.Button([
                    html.Span("📋", className="nav-icon"),
                    html.Span("Summary", className="nav-text")
                ], id="nav-summary", className="nav-link")
            ], className="nav-item"),

            html.Li([
                html.Button([
                    html.Span("📁", className="nav-icon"),
                    html.Span("Reports", className="nav-text")
                ], id="nav-reports", className="nav-link")
            ], className="nav-item"),

            html.Li([
                html.Button([
                    html.Span("⚙️", className="nav-icon"),
                    html.Span("Settings", className="nav-text")
                ], id="nav-settings", className="nav-link")
            ], className="nav-item"),
        ], className="nav-menu"),

        # Statistics Section
        html.Div([
            html.H4("📊 Quick Stats", className="stats-title"),
            html.Div([
                html.Div([
                    html.Span("Total Properties", className="stat-label"),
                    html.Span(f"{total_north_properties + total_south_properties:,}", className="stat-value")
                ], className="stat-item"),

                html.Div([
                    html.Span("Total Councils", className="stat-label"),
                    html.Span(str(total_councils), className="stat-value")
                ], className="stat-item"),

                html.Div([
                    html.Span("North Goa Wards", className="stat-label"),
                    html.Span(str(total_north_wards), className="stat-value")
                ], className="stat-item"),

                html.Div([
                    html.Span("South Goa Wards", className="stat-label"),
                    html.Span(str(total_south_wards), className="stat-value")
                ], className="stat-item"),
            ])
        ], className="stats-section"),

        # Footer
        html.Div([
            html.P("Goa Property Tax System"),
            html.P("© 2024 Trivek Bisen", style={"fontSize": "10px", "marginTop": "5px"})
        ], className="sidebar-footer")
    ], id="sidebar", className="sidebar")


# -------------------------------------------------
# 8. APP LAYOUT
# -------------------------------------------------
app.layout = html.Div([
    # Main Container
    html.Div([
        # Sidebar
        create_sidebar(),

        # Main Content
        html.Div([
            # Header
            html.Div([
                html.H1("Goa Property Tax Assessment Survey",
                        style={"textAlign": "center", "marginBottom": "10px", "color": "#2c3e50"}),
                html.P("Dashboard Design by: Trivek Bisen",
                       style={"textAlign": "center", "color": "#7f8c8d", "marginBottom": "30px"})
            ], id="page-header"),

            # Main Content Area (will be updated based on navigation)
            html.Div(id="main-content", children=dashboard_page(),
                     style={"height": "calc(100vh - 150px)", "overflowY": "auto", "paddingRight": "10px"})

        ], id="main-content-area", className="main-content expanded",
            style={"height": "100vh", "overflow": "hidden"})
    ], className="container",
        style={"fontFamily": "Arial, sans-serif", "display": "flex", "height": "100vh", "margin": "0", "padding": "0"})
])


# -------------------------------------------------
# 5. PAGE LAYOUTS - SIMPLIFIED VERSIONS
# -------------------------------------------------

def north_goa_page():
    # Calculate statistics
    total_properties = sum(c["Existing Properties"] for c in north_goa_councils)
    total_wards = sum(c["Total Wards"] for c in north_goa_councils)

    return html.Div([
        html.H1("North Goa Councils",
                style={"textAlign": "center", "color": "#2c3e50", "marginBottom": "20px", "paddingTop": "20px"}),

        # Statistics
        html.Div([
            html.Div([
                html.H4(f"{total_properties:,}", style={"color": "#3498db", "margin": "0", "fontSize": "24px"}),
                html.P("Total Properties", style={"color": "#7f8c8d", "margin": "5px 0 0 0", "fontSize": "14px"})
            ], style={"textAlign": "center", "padding": "15px", "backgroundColor": "white", "borderRadius": "10px",
                      "boxShadow": "0 2px 5px rgba(0,0,0,0.1)", "margin": "10px", "flex": "1", "minWidth": "200px"}),

            html.Div([
                html.H4(f"{total_wards}", style={"color": "#2ecc71", "margin": "0", "fontSize": "24px"}),
                html.P("Total Wards", style={"color": "#7f8c8d", "margin": "5px 0 0 0", "fontSize": "14px"})
            ], style={"textAlign": "center", "padding": "15px", "backgroundColor": "white", "borderRadius": "10px",
                      "boxShadow": "0 2px 5px rgba(0,0,0,0.1)", "margin": "10px", "flex": "1", "minWidth": "200px"}),

            html.Div([
                html.H4(f"{len(north_goa_councils)}", style={"color": "#e74c3c", "margin": "0", "fontSize": "24px"}),
                html.P("Total Councils", style={"color": "#7f8c8d", "margin": "5px 0 0 0", "fontSize": "14px"})
            ], style={"textAlign": "center", "padding": "15px", "backgroundColor": "white", "borderRadius": "10px",
                      "boxShadow": "0 2px 5px rgba(0,0,0,0.1)", "margin": "10px", "flex": "1", "minWidth": "200px"})
        ], style={"display": "flex", "flexWrap": "wrap", "justifyContent": "center", "marginBottom": "30px",
                  "marginTop": "20px"}),

        # Council Cards
        html.Div([
            html.Div([
                html.H4(council["Council"], style={"color": "#2c3e50", "marginBottom": "10px", "fontSize": "18px"}),
                html.P(f"Wards: {council['Total Wards']}",
                       style={"color": "#7f8c8d", "margin": "5px 0", "fontSize": "14px"}),
                html.P(f"Properties: {council['Existing Properties']:,}",
                       style={"color": "#7f8c8d", "margin": "5px 0", "fontSize": "14px"}),
                html.Button("View Ward Details",
                            id={"type": "council-button", "index": council["Council"], "district": "North-Goa"},
                            n_clicks=0,
                            style={
                                "backgroundColor": "#3498db",
                                "color": "white",
                                "border": "none",
                                "padding": "10px 20px",
                                "borderRadius": "5px",
                                "cursor": "pointer",
                                "marginTop": "15px",
                                "width": "100%",
                                "fontSize": "14px"
                            })
            ], style={
                "backgroundColor": "white",
                "padding": "20px",
                "borderRadius": "10px",
                "boxShadow": "0 2px 10px rgba(0,0,0,0.1)",
                "margin": "15px",
                "width": "280px",
                "height": "220px",
                "display": "flex",
                "flexDirection": "column"
            })
            for council in north_goa_councils
        ], style={
            "display": "flex",
            "flexWrap": "wrap",
            "justifyContent": "center",
            "gap": "20px",
            "padding": "20px",
            "overflowY": "auto",
            "maxHeight": "70vh"
        }),
    ])


def south_goa_page():
    # Calculate statistics
    total_properties = sum(c.get("NIC Record Properties", 0) for c in south_goa_councils)

    return html.Div([
        html.H1("South Goa Councils",
                style={"textAlign": "center", "color": "#2c3e50", "marginBottom": "20px", "paddingTop": "20px"}),

        # Statistics
        html.Div([
            html.Div([
                html.H4(f"{total_properties:,}", style={"color": "#3498db", "margin": "0", "fontSize": "24px"}),
                html.P("Total NIC Properties", style={"color": "#7f8c8d", "margin": "5px 0 0 0", "fontSize": "14px"})
            ], style={"textAlign": "center", "padding": "15px", "backgroundColor": "white", "borderRadius": "10px",
                      "boxShadow": "0 2px 5px rgba(0,0,0,0.1)", "margin": "10px", "flex": "1", "minWidth": "200px"}),

            html.Div([
                html.H4(f"{len(south_goa_councils)}", style={"color": "#e74c3c", "margin": "0", "fontSize": "24px"}),
                html.P("Total Councils", style={"color": "#7f8c8d", "margin": "5px 0 0 0", "fontSize": "14px"})
            ], style={"textAlign": "center", "padding": "15px", "backgroundColor": "white", "borderRadius": "10px",
                      "boxShadow": "0 2px 5px rgba(0,0,0,0.1)", "margin": "10px", "flex": "1", "minWidth": "200px"}),
        ], style={"display": "flex", "flexWrap": "wrap", "justifyContent": "center", "marginBottom": "30px",
                  "marginTop": "20px"}),

        # Council Cards
        html.Div([
            html.Div([
                html.H4(council["Council"], style={"color": "#2c3e50", "marginBottom": "10px", "fontSize": "18px"}),
                html.P(f"NIC Properties: {council.get('NIC Record Properties', 0):,}",
                       style={"color": "#7f8c8d", "margin": "5px 0", "fontSize": "14px"}),
                html.P(f"Order Date: {council.get('Order Issue Date', 'N/A')}",
                       style={"color": "#7f8c8d", "margin": "5px 0", "fontSize": "14px"}),
                html.Button("View Ward Details",
                            id={"type": "council-button", "index": council["Council"], "district": "South-Goa"},
                            n_clicks=0,
                            style={
                                "backgroundColor": "#3498db",
                                "color": "white",
                                "border": "none",
                                "padding": "10px 20px",
                                "borderRadius": "5px",
                                "cursor": "pointer",
                                "marginTop": "15px",
                                "width": "100%",
                                "fontSize": "14px"
                            })
            ], style={
                "backgroundColor": "white",
                "padding": "20px",
                "borderRadius": "10px",
                "boxShadow": "0 2px 10px rgba(0,0,0,0.1)",
                "margin": "15px",
                "width": "280px",
                "height": "220px",
                "display": "flex",
                "flexDirection": "column"
            })
            for council in south_goa_councils
        ], style={
            "display": "flex",
            "flexWrap": "wrap",
            "justifyContent": "center",
            "gap": "20px",
            "padding": "20px",
            "overflowY": "auto",
            "maxHeight": "70vh"
        }),
    ])


def dashboard_page():
    return html.Div([
        html.H1("Dashboard Overview",
                style={"textAlign": "center", "color": "#2c3e50", "marginBottom": "30px", "paddingTop": "20px"}),

        # Pie Charts Row
        html.Div([
            html.Div([
                create_pie_chart("North-Goa", north_goa_councils)
            ], style={"width": "48%", "display": "inline-block", "verticalAlign": "top"}),

            html.Div([
                create_pie_chart("South-Goa", south_goa_councils)
            ], style={"width": "48%", "display": "inline-block", "verticalAlign": "top"})
        ], style={"marginBottom": "40px"}),

        # Column Charts Row
        html.Div([
            html.Div([
                create_column_chart("North-Goa", north_goa_councils)
            ], style={"width": "48%", "display": "inline-block", "verticalAlign": "top"}),

            html.Div([
                create_column_chart("South-Goa", south_goa_councils)
            ], style={"width": "48%", "display": "inline-block", "verticalAlign": "top"})
        ])
    ])


def summary_page():
    return html.Div([
        html.H1("South Goa Summary Report",
                style={"textAlign": "center", "color": "#2c3e50", "marginBottom": "30px", "paddingTop": "20px"}),

        dash_table.DataTable(
            data=south_goa_summary.to_dict("records") if not south_goa_summary.empty else [],
            columns=[{"name": str(col), "id": str(col)} for col in
                     south_goa_summary.columns] if not south_goa_summary.empty else [],
            page_size=20,
            style_table={
                "overflowX": "auto",
                "borderRadius": "10px",
                "boxShadow": "0 2px 10px rgba(0,0,0,0.1)",
                "margin": "20px"
            },
            style_cell={"textAlign": "left", "padding": "10px", "border": "1px solid #eee"},
            style_header={'backgroundColor': '#3498db', 'fontWeight': 'bold', 'color': 'white'},
            style_data_conditional=[
                {'if': {'row_index': 'odd'}, 'backgroundColor': '#f8f9fa'}
            ],
        )
    ])

def reports_page():
    return html.Div([
        html.H1("Reports", style={"textAlign": "center", "color": "#2c3e50", "marginBottom": "30px"}),
        html.P("Reports section will be implemented soon.", style={"textAlign": "center", "color": "#7f8c8d"})
    ])


def settings_page():
    return html.Div([
        html.H1("Settings", style={"textAlign": "center", "color": "#2c3e50", "marginBottom": "30px"}),
        html.P("Settings section will be implemented soon.", style={"textAlign": "center", "color": "#7f8c8d"})
    ])


# -------------------------------------------------
# 8. APP LAYOUT - SIMPLIFIED VERSION
# -------------------------------------------------
app.layout = html.Div([
    # Main Container
    html.Div([
        # Sidebar
        create_sidebar(),

        # Main Content
        html.Div([
            # Main Content Area (will be updated based on navigation)
            html.Div(id="main-content", children=dashboard_page())

        ], id="main-content-area", className="main-content expanded")
    ], className="container", style={"fontFamily": "Arial, sans-serif", "display": "flex"})
])


# -------------------------------------------------
# 9. CALLBACKS - FIXED VERSION
# -------------------------------------------------

# Navigation Callback - FIXED
@app.callback(
    [Output("main-content", "children"),
     Output("nav-dashboard", "className"),
     Output("nav-north-goa", "className"),
     Output("nav-south-goa", "className"),
     Output("nav-summary", "className"),
     Output("nav-reports", "className"),
     Output("nav-settings", "className")],
    [Input("nav-dashboard", "n_clicks"),
     Input("nav-north-goa", "n_clicks"),
     Input("nav-south-goa", "n_clicks"),
     Input("nav-summary", "n_clicks"),
     Input("nav-reports", "n_clicks"),
     Input("nav-settings", "n_clicks")],
    prevent_initial_call=True
)
def navigate_page(dash_clicks, north_clicks, south_clicks, summary_clicks, reports_clicks, settings_clicks):
    ctx = dash.callback_context
    if not ctx.triggered:
        return dashboard_page(), "nav-link active", "nav-link", "nav-link", "nav-link", "nav-link", "nav-link"

    trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]

    # Reset all nav links
    nav_classes = ["nav-link", "nav-link", "nav-link", "nav-link", "nav-link", "nav-link"]

    # Handle navigation buttons
    if trigger_id == "nav-dashboard":
        nav_classes[0] = "nav-link active"
        return dashboard_page(), nav_classes[0], nav_classes[1], nav_classes[2], nav_classes[3], nav_classes[4], \
        nav_classes[5]

    elif trigger_id == "nav-north-goa":
        nav_classes[1] = "nav-link active"
        return north_goa_page(), nav_classes[0], nav_classes[1], nav_classes[2], nav_classes[3], nav_classes[4], \
        nav_classes[5]

    elif trigger_id == "nav-south-goa":
        nav_classes[2] = "nav-link active"
        return south_goa_page(), nav_classes[0], nav_classes[1], nav_classes[2], nav_classes[3], nav_classes[4], \
        nav_classes[5]

    elif trigger_id == "nav-summary":
        nav_classes[3] = "nav-link active"
        return summary_page(), nav_classes[0], nav_classes[1], nav_classes[2], nav_classes[3], nav_classes[4], \
        nav_classes[5]

    elif trigger_id == "nav-reports":
        nav_classes[4] = "nav-link active"
        return reports_page(), nav_classes[0], nav_classes[1], nav_classes[2], nav_classes[3], nav_classes[4], \
        nav_classes[5]

    elif trigger_id == "nav-settings":
        nav_classes[5] = "nav-link active"
        return settings_page(), nav_classes[0], nav_classes[1], nav_classes[2], nav_classes[3], nav_classes[4], \
        nav_classes[5]

    return dashboard_page(), "nav-link active", "nav-link", "nav-link", "nav-link", "nav-link", "nav-link"


# Separate callback for council buttons (to show ward details)
@app.callback(
    Output("main-content", "children", allow_duplicate=True),
    [Input({"type": "council-button", "index": ALL, "district": ALL}, "n_clicks")],
    [State({"type": "council-button", "index": ALL, "district": ALL}, "id")],
    prevent_initial_call=True
)
def handle_council_button_click(n_clicks_list, button_ids):
    ctx = dash.callback_context
    if not ctx.triggered:
        return dash.no_update

    # Find which button was clicked
    trigger_id = ctx.triggered[0]['prop_id']

    # Extract the button index from the triggered prop_id
    for i, n_clicks in enumerate(n_clicks_list):
        if n_clicks and n_clicks > 0:
            button_info = button_ids[i]
            council_name = button_info["index"]
            district = button_info["district"]

            # Get ward details for this council
            ward_details = get_ward_details(council_name, district)
            if ward_details:
                # Add a back button to return to council list
                back_button = html.Div([
                    html.Button("← Back to Council List",
                                id=f"back-to-{district.lower().replace('-', '')}",
                                n_clicks=0,
                                style={
                                    "backgroundColor": "#95a5a6",
                                    "color": "white",
                                    "border": "none",
                                    "padding": "10px 20px",
                                    "borderRadius": "5px",
                                    "cursor": "pointer",
                                    "marginBottom": "20px"
                                })
                ])

                return html.Div([
                    back_button,
                    ward_details
                ])

    return dash.no_update


# Callback for back buttons in ward details
@app.callback(
    Output("main-content", "children", allow_duplicate=True),
    [Input(f"back-to-northgoa", "n_clicks"),
     Input(f"back-to-southgoa", "n_clicks")],
    prevent_initial_call=True
)
def handle_back_button(north_back, south_back):
    ctx = dash.callback_context
    if not ctx.triggered:
        return dash.no_update

    trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]

    if trigger_id == "back-to-northgoa":
        return north_goa_page()
    elif trigger_id == "back-to-southgoa":
        return south_goa_page()

    return dash.no_update


# Pie Chart Callback (for dashboard) - REMOVE or FIX
# This callback might be interfering with the navigation
# Remove it or modify it to not interfere with the main navigation


# -------------------------------------------------
# 10. ADD CSS - SIMPLIFIED
# -------------------------------------------------

# Create a custom CSS string
custom_css = """
.container {
    display: flex;
    min-height: 100vh;
    margin: 0;
    padding: 0;
    width: 100%;
}

.sidebar {
    width: 250px;
    background: linear-gradient(180deg, #2c3e50 0%, #1a252f 100%);
    color: white;
    padding: 20px;
    transition: all 0.3s ease;
    box-shadow: 3px 0 10px rgba(0,0,0,0.2);
    position: relative;
    z-index: 1000;
    height: 100vh;
    overflow-y: auto;
}

.sidebar.collapsed {
    width: 70px;
}

.sidebar-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 30px;
    padding-bottom: 15px;
    border-bottom: 2px solid #3498db;
}

.sidebar-title {
    font-size: 20px;
    font-weight: bold;
    color: #ecf0f1;
    white-space: nowrap;
    overflow: hidden;
}

.sidebar.collapsed .sidebar-title {
    opacity: 0;
    width: 0;
}

.toggle-btn {
    background: #3498db;
    border: none;
    color: white;
    width: 40px;
    height: 40px;
    border-radius: 50%;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 18px;
    transition: all 0.3s ease;
    box-shadow: 0 2px 5px rgba(0,0,0,0.2);
}

.toggle-btn:hover {
    background: #2980b9;
    transform: scale(1.1);
}

.nav-menu {
    list-style: none;
    padding: 0;
    margin: 0;
}

.nav-item {
    margin: 10px 0;
    border-radius: 8px;
    overflow: hidden;
    transition: all 0.3s ease;
}

.nav-link {
    display: flex;
    align-items: center;
    padding: 12px 15px;
    color: #ecf0f1;
    text-decoration: none;
    border-radius: 8px;
    transition: all 0.3s ease;
    white-space: nowrap;
    background: none;
    border: none;
    width: 100%;
    text-align: left;
    cursor: pointer;
    font-size: 16px;
}

.nav-link:hover {
    background: rgba(52, 152, 219, 0.3);
    color: white;
    padding-left: 20px;
}

.nav-link.active {
    background: #3498db;
    color: white;
    font-weight: bold;
}

.nav-icon {
    font-size: 20px;
    margin-right: 15px;
    min-width: 24px;
    text-align: center;
}

.sidebar.collapsed .nav-icon {
    margin-right: 0;
    margin-left: 8px;
}

.nav-text {
    font-size: 16px;
    transition: opacity 0.3s ease;
}

.sidebar.collapsed .nav-text {
    opacity: 0;
    width: 0;
    height: 0;
    overflow: hidden;
}

.stats-section {
    margin-top: 40px;
    padding: 15px;
    background: rgba(255,255,255,0.1);
    border-radius: 10px;
}

.stats-title {
    font-size: 14px;
    color: #bdc3c7;
    margin-bottom: 10px;
    text-transform: uppercase;
    letter-spacing: 1px;
}

.sidebar.collapsed .stats-title {
    opacity: 0;
}

.stat-item {
    display: flex;
    justify-content: space-between;
    margin: 8px 0;
    font-size: 14px;
}

.stat-value {
    color: #3498db;
    font-weight: bold;
}

.sidebar.collapsed .stat-value,
.sidebar.collapsed .stat-label {
    opacity: 0;
}

.sidebar-footer {
    margin-top: 40px;
    padding-top: 15px;
    border-top: 1px solid rgba(255,255,255,0.1);
    font-size: 12px;
    color: #95a5a6;
    text-align: center;
}

.sidebar.collapsed .sidebar-footer {
    opacity: 0;
}

.main-content {
    flex: 1;
    padding: 0;
    transition: all 0.3s ease;
    background: #f8f9fa;
    width: calc(100% - 250px);
    height: 100vh;
    overflow-y: auto;
}

.main-content.collapsed {
    width: calc(100% - 70px);
}

/* Scrollbar styling */
::-webkit-scrollbar {
    width: 8px;
    height: 8px;
}

::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 4px;
}

::-webkit-scrollbar-thumb {
    background: #c1c1c1;
    border-radius: 4px;
}

::-webkit-scrollbar-thumb:hover {
    background: #a8a8a8;
}
"""

# Add the CSS to the app
app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>Goa Property Tax Dashboard</title>
        {%favicon%}
        {%css%}
        <style>
''' + custom_css + '''
        </style>
    </head>
    <body style="margin: 0; padding: 0; overflow: hidden;">
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>
'''

# -------------------------------------------------
# 11. RUN SERVER
# -------------------------------------------------
import os

server = app.server

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8050))
    app.run(host="0.0.0.0", port=port, debug=False)
