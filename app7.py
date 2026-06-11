import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import StringIO, BytesIO

# =========================
# Page settings
# =========================
st.set_page_config(
    page_title="Well Production Analyzer",
    page_icon="🛢️",
    layout="wide"
)

st.title("🛢️ Well Production Test Analyzer")
st.write("Upload Excel Test Sheet, choose parameters, add events, and generate interactive production plots.")

# =========================
# Sidebar
# =========================
uploaded_file = st.sidebar.file_uploader(
    "Upload Excel Test Sheet",
    type=["xlsx", "xls"]
)

st.sidebar.header("Excel Reading Settings")
sheet_index = st.sidebar.number_input(
    "Sheet Number",
    min_value=0,
    value=1,
    help="0 = first sheet, 1 = second sheet"
)

start_row = st.sidebar.number_input(
    "Start Data Row",
    min_value=0,
    value=17,
    help="Row index starts from 0"
)

end_row = st.sidebar.number_input(
    "End Data Row",
    min_value=1,
    value=46,
    help="Increase this if your test sheet has more readings"
)

st.sidebar.header("Select Parameters")

parameter_options = {
    "Gross Rate": st.sidebar.checkbox("Gross Rate", True),
    "Oil Rate": st.sidebar.checkbox("Oil Rate", True),
    "Water Rate": st.sidebar.checkbox("Water Rate", True),
    "Gas Rate": st.sidebar.checkbox("Gas Rate", True),
    "WHP": st.sidebar.checkbox("WHP", True),
    "FLP": st.sidebar.checkbox("FLP", False),
    "Choke": st.sidebar.checkbox("Choke", False),
    "BS&W": st.sidebar.checkbox("BS&W", False),
    "Salinity": st.sidebar.checkbox("Salinity", False),
}

st.sidebar.header("Label & Scale Settings")

label_step = st.sidebar.slider(
    "Show Label Every N Points",
    min_value=1,
    max_value=10,
    value=2,
    help="1 = show all labels, 2 = show one label every 2 points, etc."
)

y_padding_percent = st.sidebar.slider(
    "Y Axis Padding %",
    min_value=5,
    max_value=50,
    value=15,
    help="Extra space above and below each chart."
)

st.sidebar.header("Event Manager")

event_count = st.sidebar.number_input(
    "Number of Events",
    min_value=0,
    max_value=20,
    value=0
)

events = []

for i in range(event_count):
    st.sidebar.markdown(f"### Event {i + 1}")

    event_time = st.sidebar.text_input(
        f"Event Time {i + 1}",
        key=f"event_time_{i}",
        placeholder="Example: 05:00"
    )

    event_parameter = st.sidebar.selectbox(
        f"Attach To Parameter {i + 1}",
        list(parameter_options.keys()),
        key=f"event_parameter_{i}"
    )

    event_comment = st.sidebar.text_area(
        f"Comment {i + 1}",
        key=f"event_comment_{i}",
        placeholder="Example: Generator trip / Well restarted"
    )

    if event_time and event_comment:
        events.append(
            {
                "time": event_time.strip(),
                "parameter": event_parameter,
                "comment": event_comment
            }
        )

# =========================
# Functions
# =========================
def format_time_value(value):
    """Convert Excel time/date values into clean HH:MM text when possible."""
    if pd.isna(value):
        return ""

    try:
        if hasattr(value, "strftime"):
            return value.strftime("%H:%M")
    except Exception:
        pass

    text = str(value).strip()

    if " " in text and ":" in text:
        try:
            return pd.to_datetime(text).strftime("%H:%M")
        except Exception:
            return text

    if ":" in text:
        parts = text.split(":")
        if len(parts) >= 2:
            return f"{parts[0].zfill(2)}:{parts[1].zfill(2)}"

    return text


def load_test_sheet(file, sheet_index, start_row, end_row):
    """
    Reads the user's test sheet format.

    Expected columns based on your uploaded file:
    Time        = Column B
    Choke       = Column C
    WHP         = Column D
    FLP         = Column E
    Gas Rate    = Column M
    Oil Rate    = Column W
    Water Rate  = Column X
    BS&W        = Column Y
    Salinity    = Column Z
    Gross Rate  = Column AA
    """
    raw = pd.read_excel(file, sheet_name=sheet_index, header=None)
    df = raw.iloc[start_row:end_row].copy()

    final = pd.DataFrame()

    final["Date"] = df.iloc[:, 0]
    final["Time"] = df.iloc[:, 1].apply(format_time_value)
    final["Choke"] = pd.to_numeric(df.iloc[:, 2], errors="coerce")
    final["WHP"] = pd.to_numeric(df.iloc[:, 3], errors="coerce")
    final["FLP"] = pd.to_numeric(df.iloc[:, 4], errors="coerce")
    final["Gas Rate"] = pd.to_numeric(df.iloc[:, 12], errors="coerce")
    final["Oil Rate"] = pd.to_numeric(df.iloc[:, 22], errors="coerce")
    final["Water Rate"] = pd.to_numeric(df.iloc[:, 23], errors="coerce")
    final["BS&W"] = pd.to_numeric(df.iloc[:, 24], errors="coerce")
    final["Salinity"] = pd.to_numeric(df.iloc[:, 25], errors="coerce")
    final["Gross Rate"] = pd.to_numeric(df.iloc[:, 26], errors="coerce")

    final = final.dropna(subset=["Time"])
    final = final[final["Time"].astype(str).str.strip() != ""]

    final = final[
        final["Gross Rate"].notna()
        | final["Oil Rate"].notna()
        | final["Water Rate"].notna()
        | final["Gas Rate"].notna()
        | final["WHP"].notna()
    ]

    final = final.reset_index(drop=True)

    return final


def get_well_info(file, sheet_index):
    """Try to read well name and test date from the sheet. If failed, use defaults."""
    raw = pd.read_excel(file, sheet_name=sheet_index, header=None)

    well_name = "Unknown Well"
    test_date = ""

    try:
        possible_well = raw.iloc[2, 2]
        if pd.notna(possible_well):
            well_name = str(possible_well)
    except Exception:
        pass

    try:
        possible_date = raw.iloc[0, 25]
        if pd.notna(possible_date):
            test_date = pd.to_datetime(possible_date).strftime("%d/%m/%Y")
    except Exception:
        pass

    return well_name, test_date


def build_plot(df, selected_parameters, events, well_name, test_date, label_step=2, y_padding_percent=15):
    units = {
        "Gross Rate": "BBL/D",
        "Oil Rate": "STB/D",
        "Water Rate": "BBL/D",
        "Gas Rate": "MMSCF/D",
        "WHP": "PSI",
        "FLP": "PSI",
        "Choke": "%",
        "BS&W": "%",
        "Salinity": "K PPM",
    }

    colors = {
        "Gross Rate": "#0033ff",
        "Oil Rate": "#008000",
        "Water Rate": "#ff0000",
        "Gas Rate": "#ff9900",
        "WHP": "#6f00ff",
        "FLP": "#8b4513",
        "Choke": "#111111",
        "BS&W": "#990000",
        "Salinity": "#008b8b",
    }

    fig = make_subplots(
        rows=len(selected_parameters),
        cols=1,
        shared_xaxes=True,
        vertical_spacing=0.09,
        subplot_titles=[f"{p.upper()} ({units[p]})" for p in selected_parameters],
    )

    for row_number, parameter in enumerate(selected_parameters, start=1):
        # Main line trace
        fig.add_trace(
            go.Scatter(
                x=df["Time"],
                y=df[parameter],
                mode="lines+markers",
                name=parameter,
                line=dict(color=colors[parameter], width=2),
                marker=dict(color=colors[parameter], size=8),
                hovertemplate=(
                    "<b>Time:</b> %{x}<br>"
                    f"<b>{parameter}:</b> %{{y}} {units[parameter]}<br>"
                    "<extra></extra>"
                ),
            ),
            row=row_number,
            col=1,
        )

        fig.update_yaxes(
            title_text=units[parameter],
            row=row_number,
            col=1,
            showgrid=True,
            gridcolor="rgba(0,0,0,0.12)",
        )

        # Add vertical scale padding so labels are not stuck to the chart edges
        y_min = df[parameter].min()
        y_max = df[parameter].max()

        if pd.notna(y_min) and pd.notna(y_max):
            if y_max == y_min:
                padding = abs(y_max) * (y_padding_percent / 100)
                if padding == 0:
                    padding = 1
            else:
                padding = (y_max - y_min) * (y_padding_percent / 100)

            y_axis_min = y_min - padding
            y_axis_max = y_max + padding

            fig.update_yaxes(
                range=[y_axis_min, y_axis_max],
                row=row_number,
                col=1
            )

            # Data labels are drawn as a separate text trace above the line.
            # This prevents values from overlapping the line and uses a neutral color,
            # same color as the graph line.
            label_offset = padding * 0.35
            label_y = df[parameter] + label_offset
            label_y = label_y.clip(upper=y_axis_max - padding * 0.10)

            label_text = [
                (
                    f"{v:.2f}" if pd.notna(v) and abs(v) < 10
                    else f"{v:.0f}"
                )
                if (i % label_step == 0 and pd.notna(v))
                else ""
                for i, v in enumerate(df[parameter])
            ]

            fig.add_trace(
                go.Scatter(
                    x=df["Time"],
                    y=label_y,
                    mode="text",
                    text=label_text,
                    textposition="middle center",
                    textfont=dict(color=colors[parameter], size=12),
                    hoverinfo="skip",
                    showlegend=False,
                ),
                row=row_number,
                col=1,
            )

        # Event annotations
        for event in events:
            if event["parameter"] == parameter:
                event_time = event["time"].strip()

                if event_time in df["Time"].astype(str).values:
                    point_row = df[df["Time"].astype(str) == event_time].iloc[0]
                    y_value = point_row[parameter]

                    if pd.notna(y_value):
                        fig.add_annotation(
                            x=event_time,
                            y=y_value,
                            text=event["comment"],
                            showarrow=True,
                            arrowhead=2,
                            arrowsize=1,
                            arrowwidth=2,
                            arrowcolor="black",
                            ax=80,
                            ay=-80,
                            bordercolor="black",
                            borderwidth=1,
                            borderpad=6,
                            bgcolor="lightyellow",
                            opacity=0.95,
                            row=row_number,
                            col=1,
                        )

    fig.update_layout(
        title=(
            f"<b>Production Test Results vs Time</b><br>"
            f"Well: {well_name} &nbsp;&nbsp; Date: {test_date}"
        ),
        height=max(650, 420 * len(selected_parameters)),
        hovermode="x unified",
        spikedistance=-1,
        hoverdistance=100,
        template="plotly_white",
        showlegend=True,
        margin=dict(l=60, r=40, t=120, b=60),
        font=dict(size=14),
    )

    # Add an external border around every subplot area
    for r in range(1, len(selected_parameters) + 1):
        x_domain = fig.layout[f"xaxis{r if r > 1 else ''}"].domain
        y_domain = fig.layout[f"yaxis{r if r > 1 else ''}"].domain

        fig.add_shape(
            type="rect",
            xref="paper",
            yref="paper",
            x0=x_domain[0],
            x1=x_domain[1],
            y0=y_domain[0],
            y1=y_domain[1],
            line=dict(color="rgba(0,0,0,0.45)", width=1.4),
            fillcolor="rgba(0,0,0,0)",
            layer="below",
        )

    # Show time axis on every chart, not only the last subplot
    for r in range(1, len(selected_parameters) + 1):
        fig.update_xaxes(
            title_text="Time",
            row=r,
            col=1,
            showticklabels=True,
            showgrid=True,
            gridcolor="rgba(0,0,0,0.12)",
            showspikes=True,
            spikemode="across",
            spikesnap="cursor",
            spikecolor="rgba(0,0,0,0.55)",
            spikethickness=1,
        )

        fig.update_yaxes(
            row=r,
            col=1,
            showspikes=True,
            spikemode="across",
            spikesnap="cursor",
            spikecolor="rgba(0,0,0,0.55)",
            spikethickness=1,
        )

    return fig


def summary_table(df, selected_parameters):
    rows = []

    for p in selected_parameters:
        rows.append(
            {
                "Parameter": p,
                "Average": round(df[p].mean(), 2),
                "Maximum": round(df[p].max(), 2),
                "Minimum": round(df[p].min(), 2),
            }
        )

    return pd.DataFrame(rows)


# =========================
# Main App
# =========================
if uploaded_file:
    try:
        df = load_test_sheet(uploaded_file, sheet_index, start_row, end_row)
        well_name, test_date = get_well_info(uploaded_file, sheet_index)

        st.subheader(f"Well: {well_name}")
        if test_date:
            st.write(f"Test Date: {test_date}")

        with st.expander("Show Uploaded Data"):
            st.dataframe(df, use_container_width=True)

        selected_parameters = [
            parameter for parameter, checked in parameter_options.items() if checked
        ]

        if len(selected_parameters) == 0:
            st.warning("اختار Parameter واحد على الأقل من الـ Checkbox.")
        elif df.empty:
            st.error("No valid data found. جرّب تغير Start Data Row / End Data Row.")
        else:
            # Time range selector
            st.sidebar.header("Time Range Filter")
            all_times = df["Time"].astype(str).tolist()

            if len(all_times) > 1:
                start_time, end_time = st.sidebar.select_slider(
                    "Select Time Range",
                    options=all_times,
                    value=(all_times[0], all_times[-1]),
                )

                start_idx = all_times.index(start_time)
                end_idx = all_times.index(end_time)

                if start_idx <= end_idx:
                    df_plot = df.iloc[start_idx : end_idx + 1].copy()
                else:
                    df_plot = df.iloc[end_idx : start_idx + 1].copy()
            else:
                df_plot = df.copy()

            fig = build_plot(
                df_plot,
                selected_parameters,
                events,
                well_name,
                test_date,
                label_step,
                y_padding_percent,
            )

            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Summary")
            st.dataframe(
                summary_table(df_plot, selected_parameters),
                use_container_width=True,
                hide_index=True,
            )

            html_buffer = StringIO()
            fig.write_html(html_buffer)

            st.download_button(
                label="Download Interactive Plot HTML",
                data=html_buffer.getvalue(),
                file_name=f"{well_name}_production_plot.html",
                mime="text/html",
            )

            # PDF report export using ReportLab + Matplotlib.
            # This includes a static chart image and does NOT need Chrome or Kaleido.
            try:
                import matplotlib.pyplot as plt
                from reportlab.lib.pagesizes import A4, landscape
                from reportlab.lib import colors as rl_colors
                from reportlab.lib.styles import getSampleStyleSheet
                from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
                from reportlab.lib.units import cm

                pdf_buffer = BytesIO()

                doc = SimpleDocTemplate(
                    pdf_buffer,
                    pagesize=landscape(A4),
                    rightMargin=1.0 * cm,
                    leftMargin=1.0 * cm,
                    topMargin=1.0 * cm,
                    bottomMargin=1.0 * cm,
                )

                styles = getSampleStyleSheet()
                story = []

                story.append(Paragraph("Well Production Test Report", styles["Title"]))
                story.append(Spacer(1, 8))
                story.append(Paragraph(f"<b>Well:</b> {well_name}", styles["Normal"]))
                story.append(Paragraph(f"<b>Test Date:</b> {test_date}", styles["Normal"]))
                story.append(Spacer(1, 10))

                # Build static chart for PDF using matplotlib
                chart_colors = {
                    "Gross Rate": "#0033ff",
                    "Oil Rate": "#008000",
                    "Water Rate": "#ff0000",
                    "Gas Rate": "#ff9900",
                    "WHP": "#6f00ff",
                    "FLP": "#8b4513",
                    "Choke": "#111111",
                    "BS&W": "#990000",
                    "Salinity": "#008b8b",
                }

                chart_units = {
                    "Gross Rate": "BBL/D",
                    "Oil Rate": "STB/D",
                    "Water Rate": "BBL/D",
                    "Gas Rate": "MMSCF/D",
                    "WHP": "PSI",
                    "FLP": "PSI",
                    "Choke": "%",
                    "BS&W": "%",
                    "Salinity": "K PPM",
                }

                fig_pdf, axes = plt.subplots(
                    len(selected_parameters),
                    1,
                    figsize=(15, max(4, 2.7 * len(selected_parameters))),
                    sharex=False
                )

                if len(selected_parameters) == 1:
                    axes = [axes]

                for ax, p in zip(axes, selected_parameters):
                    x_values = df_plot["Time"].astype(str).tolist()
                    y_values = df_plot[p]

                    ax.plot(
                        x_values,
                        y_values,
                        marker="o",
                        linewidth=2,
                        color=chart_colors.get(p, "#000000")
                    )

                    ax.set_title(f"{p} ({chart_units[p]})", fontsize=11, fontweight="bold")
                    ax.set_ylabel(chart_units[p])
                    ax.set_xlabel("Time")
                    ax.grid(True, alpha=0.3)

                    y_min = y_values.min()
                    y_max = y_values.max()

                    if pd.notna(y_min) and pd.notna(y_max):
                        if y_max == y_min:
                            pad = abs(y_max) * 0.15 if y_max != 0 else 1
                        else:
                            pad = (y_max - y_min) * 0.15

                        ax.set_ylim(y_min - pad, y_max + pad)

                        label_offset = pad * 0.35

                        for i, (x, y) in enumerate(zip(x_values, y_values)):
                            if i % label_step == 0 and pd.notna(y):
                                label = f"{y:.2f}" if abs(y) < 10 else f"{y:.0f}"
                                ax.text(
                                    i,
                                    y + label_offset,
                                    label,
                                    ha="center",
                                    va="bottom",
                                    fontsize=8,
                                    color=chart_colors.get(p, "#000000"),
                                    fontweight="bold"
                                )

                    for e in events:
                        if e["parameter"] == p and e["time"] in x_values:
                            idx = x_values.index(e["time"])
                            y_event = y_values.iloc[idx]

                            if pd.notna(y_event):
                                ax.annotate(
                                    e["comment"],
                                    xy=(idx, y_event),
                                    xytext=(idx + 0.5, y_event + (pad * 1.5 if "pad" in locals() else 1)),
                                    arrowprops=dict(arrowstyle="->", lw=1),
                                    bbox=dict(boxstyle="round,pad=0.3", fc="lightyellow", ec="black"),
                                    fontsize=8
                                )

                    ax.tick_params(axis="x", labelrotation=45)

                fig_pdf.suptitle(
                    f"Production Test Results vs Time - Well {well_name}",
                    fontsize=14,
                    fontweight="bold"
                )
                fig_pdf.tight_layout(rect=[0, 0, 1, 0.97])

                chart_buffer = BytesIO()
                fig_pdf.savefig(chart_buffer, format="png", dpi=180, bbox_inches="tight")
                plt.close(fig_pdf)
                chart_buffer.seek(0)

                story.append(Image(chart_buffer, width=27 * cm, height=16 * cm))
                story.append(PageBreak())

                story.append(Paragraph("Summary", styles["Heading2"]))

                summary_df = summary_table(df_plot, selected_parameters)
                summary_data = [summary_df.columns.tolist()] + summary_df.astype(str).values.tolist()

                summary_tbl = Table(summary_data, repeatRows=1)
                summary_tbl.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), rl_colors.lightgrey),
                    ("GRID", (0, 0), (-1, -1), 0.5, rl_colors.grey),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("ALIGN", (1, 1), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ]))

                story.append(summary_tbl)
                story.append(Spacer(1, 14))

                if events:
                    story.append(Paragraph("Events", styles["Heading2"]))
                    events_data = [["Time", "Parameter", "Comment"]]
                    for e in events:
                        events_data.append([e["time"], e["parameter"], e["comment"]])

                    events_tbl = Table(events_data, repeatRows=1)
                    events_tbl.setStyle(TableStyle([
                        ("BACKGROUND", (0, 0), (-1, 0), rl_colors.lightgrey),
                        ("GRID", (0, 0), (-1, -1), 0.5, rl_colors.grey),
                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ]))
                    story.append(events_tbl)
                    story.append(Spacer(1, 14))

                doc.build(story)
                pdf_buffer.seek(0)

                st.download_button(
                    label="Download PDF Report with Charts",
                    data=pdf_buffer.getvalue(),
                    file_name=f"{well_name}_production_report.pdf",
                    mime="application/pdf",
                )

            except Exception as pdf_error:
                st.warning("PDF report export failed.")
                st.code(str(pdf_error))

    except Exception as e:
        st.error("حصل Error أثناء قراءة الملف.")
        st.write("تفاصيل الخطأ:")
        st.code(str(e))
        st.info("جرّب تغير Sheet Number أو Start/End Row من الشمال.")
else:
    st.info("ارفع Excel Test Sheet من الشمال علشان يبدأ التحليل.")
