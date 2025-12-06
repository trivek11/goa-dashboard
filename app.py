# -------------------------------------------------
# 0. IMPORTS
# -------------------------------------------------
import pandas as pd
import plotly.graph_objects as go
from dash import Dash, html, dcc, Input, Output, ALL, dash_table
import dash
from datetime import datetime
import re

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

    # Print column information for all councils
    print(f"\n{'=' * 60}")
    print("COUNCIL-WISE COLUMN ANALYSIS")
    print(f"{'=' * 60}")
    for council, info in council_columns_info.items():
        print(f"\n{council}:")
        print(f"  Original columns: {info['original']}")
        print(f"  Cleaned columns:  {info['cleaned']}")

    # Combine all ward data
    if ward_sheets:
        south_goa_wards = pd.concat(ward_sheets, ignore_index=True)
        print(f"\nCombined DataFrame shape: {south_goa_wards.shape}")
        print(f"Combined columns: {south_goa_wards.columns.tolist()}")

        # Check for common columns across all councils
        common_columns = set(south_goa_wards.columns)
        print(f"\nCommon columns across all councils: {common_columns}")
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

    print(f"\nSummary sheet analysis:")
    print(f"  Council column: {council_col}")
    print(f"  NIC Property column: {nic_property_col}")
    print(f"  Order Issue Date column: {order_date_col}")

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
# 5. DISTRICT LAYOUT (PIE CHART)
# -------------------------------------------------
def district_layout(district, councils, wards):
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
            ], style={"width": "48%", "display": "inline-block", "verticalAlign": "top"})

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
        ),
        html.Div(id={"type": "modal", "index": district})
    ], style={"width": "48%", "display": "inline-block", "verticalAlign": "top"})


# -------------------------------------------------
# 6. APP LAYOUT
# -------------------------------------------------
app.layout = html.Div([
    html.H1("Goa Property Tax Dashboard", style={"textAlign": "center"}),

    # District pie charts
    html.Div([
        district_layout("North-Goa", north_goa_councils, north_goa_wards),
        district_layout("South-Goa", south_goa_councils, south_goa_wards)
    ], style={"display": "flex", "justifyContent": "space-around"}),

    # Summary tables section
    html.Div(id="summary-section", children=[
        html.H2("South Goa Summary"),
        dash_table.DataTable(
            data=south_goa_summary.to_dict("records") if not south_goa_summary.empty else [],
            columns=[{"name": str(col), "id": str(col)} for col in
                     south_goa_summary.columns] if not south_goa_summary.empty else [],
            page_size=20,
            style_table={"overflowX": "auto"},
            style_cell={"textAlign": "left", "padding": "5px"},
            style_header={'backgroundColor': 'lightblue', 'fontWeight': 'bold'},
        )
    ]),

    # Modal area for detailed views
    html.Div(id="ward-modal-area", style={"marginTop": "20px"})
])


# -------------------------------------------------
# 7. CALLBACK FOR PIE CLICKS
# -------------------------------------------------
@app.callback(
    Output("ward-modal-area", "children"),
    Input({"type": "pie", "index": ALL}, "clickData")
)
def show_ward_details(pie_clicks):
    ctx = dash.callback_context
    if not ctx.triggered:
        return html.Div("Click on a council in the pie chart to see ward details")

    trigger = ctx.triggered[0]
    prop_id = trigger["prop_id"]

    if not trigger["value"]:
        return html.Div()

    try:
        district = eval(prop_id.split(".")[0])["index"]
        clicked_data = trigger["value"]
        clicked_label = clicked_data["points"][0]["label"]

        print(f"\n{'=' * 60}")
        print(f"Clicked on: District={district}, Council={clicked_label}")
        print(f"{'=' * 60}")

        if district == "North-Goa":
            council_name = clicked_label
            ward_data = north_goa_wards[north_goa_wards["Council"] == council_name]

            # Get NIC Record Property for this council
            nic_property_count = 0
            for council in north_goa_councils:
                if council["Council"] == clicked_label:
                    nic_property_count = council.get("Existing Properties", 0)
                    break

            print(f"North Goa ward data columns: {ward_data.columns.tolist()}")
            print(f"Sample data:\n{ward_data.head()}")

        elif district == "South-Goa":
            council_name = clicked_label

            if not south_goa_wards.empty and "Council" in south_goa_wards.columns:
                ward_data = south_goa_wards[south_goa_wards["Council"] == council_name]
                print(f"South Goa ward data columns for {council_name}: {ward_data.columns.tolist()}")
                print(f"Sample data:\n{ward_data.head()}")
            else:
                return html.Div([html.H3("No ward data available for South Goa")])

            # Get NIC Record Property and other details for this council
            nic_property_count = 0
            order_issue_date = ""
            for council in south_goa_councils:
                if council["Council"] == council_name:
                    nic_property_count = council.get("NIC Record Properties", 0)
                    order_issue_date = council.get("Order Issue Date", "")
                    break

        if ward_data.empty:
            return html.Div([
                html.H3(f"No ward data found for {clicked_label}"),
                html.P(f"District: {district}")
            ])

        # Clean column names (remove duplicates)
        ward_data = ward_data.loc[:, ~ward_data.columns.duplicated()]

        # Get available columns (excluding Council and District)
        available_cols = [col for col in ward_data.columns if col not in ["Council", "District"]]

        # Define preferred order for columns based on common patterns
        preferred_order = [
            "Sr No", "Ward No",
            "Ward Boundary Mapping", "Digitlisation Of Polygon", "Generation of UPIC Number",
            "UPIC", "NIC Property", "Suspected New", "Total Survey", "Survey",
            "Old GSUDA Properties", "Extended Area Properties", "Objections Received",
            "Remark"
        ]

        # Sort columns according to preferred order
        sorted_cols = []
        for col in preferred_order:
            if col in available_cols:
                sorted_cols.append(col)

        # Add any remaining columns
        for col in available_cols:
            if col not in sorted_cols:
                sorted_cols.append(col)

        print(f"Displaying columns ({len(sorted_cols)}): {sorted_cols}")

        # Calculate total wards
        if "Ward No" in ward_data.columns:
            total_wards = ward_data["Ward No"].nunique()
        else:
            total_wards = len(ward_data)

        # Create council details div
        council_details = [
            html.H4(f"{clicked_label} - Ward Details ({district})",
                    style={"color": "#2c3e50", "marginBottom": "15px"}),
            html.Div([
                html.Span(f"Total Wards: {total_wards}",
                          style={"fontWeight": "bold", "marginRight": "20px"}),
                html.Span(f"NIC Record Properties: {nic_property_count:,}",
                          style={"fontWeight": "bold", "marginRight": "20px"}),
                html.Span(f"Order Issue Date: {order_issue_date}",
                          style={"fontWeight": "bold"}) if order_issue_date and district == 'South-Goa' else ""
            ], style={"marginBottom": "20px", "fontSize": "16px"})
        ]

        # Add the table
        council_details.append(
            html.Div([
                html.H5(f"Showing {len(ward_data)} records",
                        style={"marginBottom": "10px", "color": "#7f8c8d"}),
                dash_table.DataTable(
                    columns=[{"name": col, "id": col} for col in sorted_cols],
                    data=ward_data[sorted_cols].to_dict("records"),
                    page_size=len(ward_data),  # Show all rows in single page
                    style_table={
                        "overflowX": "auto",
                        "border": "2px solid #bdc3c7",
                        "maxHeight": "600px",
                        "overflowY": "auto",
                        "borderRadius": "5px"
                    },
                    style_cell={
                        "textAlign": "center",
                        "padding": "10px",
                        "border": "1px solid #ecf0f1",
                        "minWidth": "100px",
                        "maxWidth": "200px",
                        "whiteSpace": "normal",
                        "fontFamily": "Arial, sans-serif"
                    },
                    style_header={
                        'backgroundColor': '#3498db',
                        'fontWeight': 'bold',
                        'border': '1px solid #2980b9',
                        'color': 'white',
                        'fontSize': '14px'
                    },
                    style_data={
                        'whiteSpace': 'normal',
                        'height': 'auto',
                        'fontSize': '13px'
                    },
                    style_data_conditional=[
                        {
                            'if': {'row_index': 'odd'},
                            'backgroundColor': '#f8f9fa'
                        }
                    ],
                    fixed_rows={'headers': True},
                )
            ])
        )

        return html.Div(council_details, style={"padding": "20px", "backgroundColor": "white", "borderRadius": "10px",
                                                "boxShadow": "0 2px 10px rgba(0,0,0,0.1)"})

    except Exception as e:
        print(f"Error in callback: {e}")
        import traceback
        traceback.print_exc()
        return html.Div([html.H3(f"Error loading data", style={"color": "red"})])


# -------------------------------------------------
# 8. RUN SERVER
import os

server = app.server

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8050))
    app.run(host="0.0.0.0", port=port, debug=False)
