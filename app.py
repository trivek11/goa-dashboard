# -------------------------------------------------
# 0.  IMPORTS  (keep yours)
# -------------------------------------------------
import pandas as pd, plotly.express as px, plotly.graph_objects as go, dash, math, os
from dash import Dash, dcc, html, Input, Output, State, ALL, MATCH
import dash.dash_table as dt
# -------------------------------------------------
# 1.  MASTER COUNCIL LIST
# -------------------------------------------------
councils = [
    # ---------- North Goa ----------
    {"District": "North-Goa", "Council": "CCP",                "Total Wards": 23,  "Work Order Issue Date": "2024-04-27", "Existing Properties": 30991},
    {"District": "North-Goa", "Council": "Mapusa",             "Total Wards": 20,  "Work Order Issue Date": "2024-04-27", "Existing Properties": 29654},
    {"District": "North-Goa", "Council": "Bicholim",           "Total Wards": 14,  "Work Order Issue Date": "2024-04-27", "Existing Properties": 8821},
    {"District": "North-Goa", "Council": "Pernem",             "Total Wards": 10,  "Work Order Issue Date": "2024-04-27", "Existing Properties": 1794},
    {"District": "North-Goa", "Council": "Valpoi",             "Total Wards": 10,  "Work Order Issue Date": "2024-04-27", "Existing Properties": 4552},
    {"District": "North-Goa", "Council": "Sanquelim",          "Total Wards": 12,  "Work Order Issue Date": "2024-04-27", "Existing Properties": 6337},
    # ---------- South Goa ----------
    {"District": "South-Goa", "Council": "Mormugao",           "Total Wards": 25,  "Work Order Issue Date": "2024-07-30", "Existing Properties": 33479},
    {"District": "South-Goa", "Council": "Ponda",              "Total Wards": 15,  "Work Order Issue Date": "2024-09-24", "Existing Properties": 13251},
    {"District": "South-Goa", "Council": "Quepem",             "Total Wards": 13,  "Work Order Issue Date": "2024-08-26", "Existing Properties": 5908},
    {"District": "South-Goa", "Council": "Margao",             "Total Wards": 25,  "Work Order Issue Date": "2024-10-11", "Existing Properties": 58975},
    {"District": "South-Goa", "Council": "Sanguem",            "Total Wards": 10,  "Work Order Issue Date": "2024-11-22", "Existing Properties": 1855},
    {"District": "South-Goa", "Council": "Curchorem-Cacora",   "Total Wards": 15,  "Work Order Issue Date": "2024-12-04", "Existing Properties": 7775},
    {"District": "South-Goa", "Council": "Canacona",           "Total Wards": 12,  "Work Order Issue Date": "2024-11-29", "Existing Properties": 5879},
    {"District": "South-Goa", "Council": "Cuncolim",           "Total Wards": 14,  "Work Order Issue Date": "2025-01-07", "Existing Properties": 6671},
]
north_goa_councils = [c for c in councils if c["District"] == "North-Goa"]
south_goa_councils = [c for c in councils if c["District"] == "South-Goa"]

# -------------------------------------------------
# 2.  WARD DATA FACTORY  (unchanged logic)
# -------------------------------------------------
def generate_ward_data(council_name, total_wards):
    council = next(c for c in councils if c["Council"] == council_name)
    upic_total = council["Existing Properties"]
    surveyed_tot = int(upic_total * 0.92)          # 92 % avg
    suspected_tot = int(upic_total * 0.18)
    old_gsuda_tot = int(upic_total * 0.75)
    ext_area_tot = int(upic_total * 0.05)
    objections_tot = max(1, total_wards // 2)

    def split(total, wards):
        base, rem = divmod(total, wards)
        return [base + 1 if i < rem else base for i in range(wards)]

    upic_lst, survey_lst, suspected_lst = split(upic_total, total_wards), split(surveyed_tot, total_wards), split(suspected_tot, total_wards)
    old_lst, ext_lst, obj_lst = split(old_gsuda_tot, total_wards), split(ext_area_tot, total_wards), split(objections_tot, total_wards)

    wards = []
    for i in range(1, total_wards + 1):
        wards.append({
            "Sr No": i, "Ward No": i,
            "Ward Boundary Mapping": "Completed", "Digitlisation Of Polygon": "Completed",
            "Generation of UPIC Number": "Completed",
            "Numbering (UPIC)": upic_lst[i - 1],
            "Survey": survey_lst[i - 1],
            "Old GSUDA Properties": old_lst[i - 1],
            "Suspected New Properties": suspected_lst[i - 1],
            "Extended Area Properties": ext_lst[i - 1],
            "Validation of Assessment Register": "Validated",
            "Objections Received": obj_lst[i - 1],
            "QR Code Installation": "Done", "Remark": "No remarks", "Optional Notes": "Optional notes",
            "Council": council_name
        })
    return pd.DataFrame(wards)

# -------------------------------------------------
# 3.  BUILD THE TWO BIG FRAMES  (now safe)
# -------------------------------------------------
north_goa_wards = pd.concat([generate_ward_data(c["Council"], c["Total Wards"]) for c in north_goa_councils], ignore_index=True)
south_goa_wards = pd.concat([generate_ward_data(c["Council"], c["Total Wards"]) for c in south_goa_councils], ignore_index=True)
for df in (north_goa_wards, south_goa_wards):
    df["Survey_UPIC_Percent"] = df["Survey"] / df["Numbering (UPIC)"] * 100

# -------------------------------------------------
# 4.  DASH APP  (keep everything below identical)
# -------------------------------------------------
app = dash.Dash(__name__, suppress_callback_exceptions=True, serve_locally=True)

ward_cols = ["Sr No", "Ward No", "Ward Boundary Mapping", "Digitlisation Of Polygon",
             "Generation of UPIC Number", "Numbering (UPIC)", "Survey",
             "Old GSUDA Properties", "Suspected New Properties", "Extended Area Properties",
             "Validation of Assessment Register", "Objections Received",
             "QR Code Installation", "Remark", "Optional Notes"]

def district_layout(district: str, councils: list, wards: pd.DataFrame):
    pie_df = wards.groupby("Council")["Survey_UPIC_Percent"].mean().reset_index()
    pie_fig = px.pie(pie_df, names="Council", values="Survey_UPIC_Percent",
                     title=f"{district} – Survey / UPIC %", hole=0.5,
                     color_discrete_sequence=px.colors.sequential.Teal)
    pie_fig.update_traces(textinfo="label+percent", textposition="outside")
    pie_fig.update_layout(transition_duration=500)

    bar_df = wards.groupby("Council")[["Old GSUDA Properties", "Numbering (UPIC)", "Survey"]].sum().reset_index()
    bar_fig = go.Figure()
    bar_fig.add_bar(x=bar_df["Council"], y=bar_df["Old GSUDA Properties"], name="Existing GSUDA", marker_color="#20B2AA")
    bar_fig.add_bar(x=bar_df["Council"], y=bar_df["Numbering (UPIC)"], name="Numbering (UPIC)", marker_color="#00CED1")
    bar_fig.add_bar(x=bar_df["Council"], y=bar_df["Survey"], name="Survey", marker_color="#008080")
    bar_fig.update_layout(title=f"{district} – Property Counts", barmode="group", height=400,
                          clickmode="event+select", transition_duration=500)

    return html.Div([
        html.H3(district, style={"textAlign": "center", "marginTop": 20}),
        dcc.Graph(figure=pie_fig, id={"type": "pie", "index": district}),
        dcc.Graph(figure=bar_fig, id={"type": "bar", "index": district}),
        html.Div(id={"type": "modal", "index": district})
    ], style={"marginBottom": 40})

app.layout = html.Div([
    html.H1("GSUDA Property Assessment Dashboard", style={"textAlign": "center", "color": "#006666"}),
    html.Div(html.Button("Refresh Dashboard", id="refresh-btn", n_clicks=0),
             style={"textAlign": "center", "marginBottom": 20}),
    html.Div([
        district_layout("North-Goa", north_goa_councils, north_goa_wards),
        district_layout("South-Goa", south_goa_councils, south_goa_wards)
    ], style={"display": "flex", "justifyContent": "space-between"}),
    html.Div(id="ward-modal-area")
], style={"fontFamily": "Arial", "margin": 20})

# -------------------------------------------------
# 5.  SINGLE MODAL CALLBACK  (pattern-matching)
# -------------------------------------------------
@app.callback(Output("ward-modal-area", "children"),
              Input({"type": "pie", "index": ALL}, "clickData"),
              Input({"type": "bar", "index": ALL}, "clickData"),
              Input("refresh-btn", "n_clicks"))
def show_modal(pie_clicks, bar_clicks, refresh):
    ctx = dash.callback_context
    if not ctx.triggered or ctx.triggered[0]["prop_id"] == "refresh-btn.n_clicks":
        return html.Div()

    trig = ctx.triggered[0]["prop_id"]
    district = eval(trig.split(".")[0])["index"]  # 'North-Goa' or 'South-Goa'
    wards_df = north_goa_wards if district == "North-Goa" else south_goa_wards
    councils = north_goa_councils if district == "North-Goa" else south_goa_councils

    raw_click = ctx.triggered[0]["value"]
    council_name = raw_click["points"][0].get("label") or raw_click["points"][0]["x"]

    meta = next(c for c in councils if c["Council"] == council_name)
    meta_card = html.Div([
        html.H4(f"{council_name} – Council"),
        html.P(f"District: {meta['District']}"),
        html.P(f"Work Order Issue Date: {meta['Work Order Issue Date']}"),
        html.P(f"Existing GSUDA Properties: {meta['Existing Properties']:,}"),
        html.P(["GIS Link: ", html.A("Open Map",
                                     href="https://www.google.com/maps/d/u/0/edit?mid=1j2yKe4iTj4hZUwS0E0ohXjkbmCp5rQo&usp=sharing",
                                     target="_blank", rel="noopener noreferrer")])
    ], style={"border": "1px solid #008080", "padding": "10px", "borderRadius": "5px",
              "marginBottom": "20px", "backgroundColor": "#E0F7F7"})

    df = wards_df[wards_df["Council"] == council_name]
    numeric_cols = ["Numbering (UPIC)", "Survey", "Old GSUDA Properties",
                    "Suspected New Properties", "Extended Area Properties", "Objections Received"]
    total = {col: df[col].sum() if col in numeric_cols else "" for col in ward_cols}
    total["Ward No"] = "Total"
    df_display = pd.concat([df, pd.DataFrame([total])], ignore_index=True)

    table = dt.DataTable(
        columns=[{"name": c, "id": c} for c in ward_cols],
        data=df_display.to_dict("records"),
        style_cell={"textAlign": "center", "minWidth": "55px", "width": "55px", "maxWidth": "55px"},
        style_data_conditional=[{"if": {"filter_query": "{Ward No} = 'Total'"},
                                 "fontWeight": "bold", "backgroundColor": "#E0F7F7"}],
        page_size=15)

    return html.Div([meta_card, html.H4(f"{council_name} – Ward Details"), table])

# -------------------------------------------------
# bottom of app.py
server = app.server          # <-- add this line
if __name__ == "__main__":   # keep only for local testing
    app.run(host="0.0.0.0", port=8050, debug=False)