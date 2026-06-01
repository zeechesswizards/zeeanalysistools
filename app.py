import streamlit as st
import pandas as pd
import io
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

# Set up browser tab styling
st.set_page_config(page_title="CFO Strategy & Audit Engine", layout="wide")

st.title("💼 CFO Strategy & Audit Engine")
st.write("Upload your QuickBooks ledger file below to instantly generate your Executive Dashboard and Audit Report.")

# --- WEB UI: FILE UPLOADER ---
uploaded_file = st.file_uploader("Choose a QuickBooks Excel file (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        with st.spinner("Processing ledger data and compiling metrics..."):
            # 1. Open file and dynamically find where the data starts
            raw_df = pd.read_excel(uploaded_file)
            skip_rows = 0
            for i, row in raw_df.iterrows():
                if 'Date' in row.values and 'Amount' in row.values:
                    skip_rows = i + 1
                    break

            df = pd.read_excel(uploaded_file, skiprows=skip_rows)
            df.columns = df.columns.str.strip()

            # 2. Flexible Column Mapping
            col_mapping = {}
            for c in df.columns:
                c_lower = c.lower().strip()
                if c_lower in ['name', 'vendor', 'vendor name']:
                    col_mapping[c] = 'Name'
                if c_lower in ['amount', 'amount (usd)', 'total amount']:
                    col_mapping[c] = 'Amount'
                if c_lower in ['date', 'transaction date']:
                    col_mapping[c] = 'Date'
                if c_lower in ['no.', 'num', 'no', 'reference', 'doc number']:
                    col_mapping[c] = 'No.'
                if c_lower in ['account', 'account name', 'distribution account']:
                    col_mapping[c] = 'Account'
                if c_lower in ['split', 'distribution']:
                    col_mapping[c] = 'Split'
                if c_lower in ['type', 'transaction type', 'txn type']:
                    col_mapping[c] = 'Transaction Type'
                
                acct_num_names = [
                    'account #', 'account#', 'acct #', 'acct#', 
                    'account number', 'account no', 'account no.', 
                    'acct no', 'acct no.', 'acc #', 'acc#', 'account num'
                ]
                if c_lower in acct_num_names:
                    col_mapping[c] = 'Account_Num'

            df = df.rename(columns=col_mapping)

            if 'No.' not in df.columns: df['No.'] = 'N/A'
            if 'Account' not in df.columns: df['Account'] = 'N/A'
            if 'Split' not in df.columns: df['Split'] = 'N/A'
            if 'Transaction Type' not in df.columns: df['Transaction Type'] = 'N/A'

            df = df.dropna(subset=['Date', 'Amount'])
            df['Date'] = pd.to_datetime(df['Date'])
            df['Amount'] = pd.to_numeric(df['Amount']).abs()
            
            # GLOBAL ZERO-DOLLAR FILTER
            df = df[df['Amount'] > 0.0].copy()
            
            df['Name'] = df['Name'].fillna('[Unassigned/Blank Vendor]')
            df['Name'] = df['Name'].astype(str).str.strip()
            df.loc[df['Name'] == '', 'Name'] = '[Unassigned/Blank Vendor]'

            # EXCLUDED PREFIXES
            EXCLUDED_PREFIXES = ['1', '2', '3', '4']

            def extract_account_number(val):
                val = str(val).strip()
                if not val or val.lower() in ('nan', 'n/a', ''): return ''
                val = re.sub(r'\.0+$', '', val)
                match = re.match(r'^(\d+)', val)
                return match.group(1) if match else ''

            def is_excluded_code(val):
                code = extract_account_number(val)
                if not code: return False
                return any(code.startswith(p) for p in EXCLUDED_PREFIXES)

            filter_col = 'Account_Num' if 'Account_Num' in df.columns else 'Account'
            df['_filter_col'] = df[filter_col].fillna('').astype(str)
            df_filtered = df[~df['_filter_col'].apply(is_excluded_code)].copy()
            df_filtered = df_filtered.drop(columns=['_filter_col'])

            df_filtered['Month_Name'] = df_filtered['Date'].dt.strftime('%b-%y')
            df_filtered['YearMonth'] = df_filtered['Date'].dt.to_period('M')
            fiscal_map = {8:1, 9:2, 10:3, 11:4, 12:5, 1:6, 2:7, 3:8, 4:9, 5:10, 6:11, 7:12}
            df_filtered['Fiscal_Sort'] = df_filtered['Date'].dt.month.map(fiscal_map)
            df_sorted = df_filtered.sort_values(by=['Fiscal_Sort', 'Date'])

            # --- GL ACCOUNT ---
            gl_summary = df_filtered.groupby('Account')['Amount'].sum().reset_index()
            gl_summary = gl_summary.sort_values(by='Amount', ascending=False).head(5)
            gl_summary.columns = ['Top GL Accounts / Categories', 'Total Spend']

            # --- MoM TOP MOVERS ---
            recent_months = sorted(df_sorted['YearMonth'].drop_duplicates().tolist())[-2:]
            if len(recent_months) == 2:
                prev_m, curr_m = recent_months[0], recent_months[1]
                df_mom = df_filtered[df_filtered['YearMonth'].isin([prev_m, curr_m])]
                mom_pivot = df_mom.pivot_table(index='Name', columns='YearMonth', values='Amount', aggfunc='sum').fillna(0)
                
                if prev_m in mom_pivot.columns and curr_m in mom_pivot.columns:
                    mom_pivot['Variance'] = mom_pivot[curr_m] - mom_pivot[prev_m]
                    if '[Unassigned/Blank Vendor]' in mom_pivot.index:
                        mom_pivot = mom_pivot.drop(index='[Unassigned/Blank Vendor]')
                    top_up = mom_pivot.nlargest(3, 'Variance').reset_index()
                    top_up['Status'] = 'SPIKED UP'
                    top_down = mom_pivot.nsmallest(3, 'Variance').reset_index()
                    top_down['Status'] = 'DROPPED DOWN'
                    movers_df = pd.concat([top_up[['Name', 'Status', 'Variance']], top_down[['Name', 'Status', 'Variance']]])
                    movers_df.columns = ['Vendor', 'MoM Movement', 'Variance ($)']
                else:
                    movers_df = pd.DataFrame({"Vendor": ["Not enough data"], "MoM Movement": ["N/A"], "Variance ($)": [0]})
            else:
                movers_df = pd.DataFrame({"Vendor": ["Not enough data for MoM"], "MoM Movement": ["N/A"], "Variance ($)": [0]})

            # --- NEW VENDOR DETECTION ---
            latest_month = sorted(df_filtered['YearMonth'].drop_duplicates().tolist())[-1]
            first_seen = df_filtered.groupby('Name')['YearMonth'].min()
            new_vendors_list = first_seen[first_seen == latest_month].index.tolist()
            if '[Unassigned/Blank Vendor]' in new_vendors_list:
                new_vendors_list.remove('[Unassigned/Blank Vendor]')
            if new_vendors_list:
                new_vendors_df = df_filtered[(df_filtered['Name'].isin(new_vendors_list)) & (df_filtered['YearMonth'] == latest_month)]
                new_vendors_summary = new_vendors_df.groupby('Name')['Amount'].sum().reset_index().sort_values(by='Amount', ascending=False)
                new_vendors_summary.columns = ['New Vendor (First Payment This Month)', 'Initial Spend']
            else:
                new_vendors_summary = pd.DataFrame({'New Vendor (First Payment This Month)': ['No new vendors detected'], 'Initial Spend': [0.00]})

            # WEEKEND EXPENSES
            df_filtered['Is_Weekend'] = df_filtered['Date'].dt.dayofweek.isin([5, 6])
            weekend_transactions = df_filtered[df_filtered['Is_Weekend']].copy()
            controls_scorecard = pd.DataFrame({
                "WEEKEND EXPENSE CONTROLS": ["Weekend Off-Cycle Outflows", "Weekend Transactions Volume"],
                "EXPOSURE": [f"${weekend_transactions['Amount'].sum():,.2f}", f"{len(weekend_transactions)} Rows Flagged"],
                "THREAT LEVEL": ["MEDIUM RISK" if len(weekend_transactions) > 0 else "LOW RISK", "MONITOR"]
            })
            weekend_transactions['Reason'] = 'Off-Hours Weekend Transaction'
            leakage_report = weekend_transactions[['Date', 'Transaction Type', 'No.', 'Name', 'Account', 'Amount', 'Reason']].sort_values(by='Amount', ascending=False)

            # DUPLICATES
            df_dup_check = df_filtered.sort_values(by=['Name', 'Amount', 'Date'])
            same_prev = ((df_dup_check['Name'] == df_dup_check['Name'].shift(1)) & (df_dup_check['Amount'] == df_dup_check['Amount'].shift(1)))
            same_next = ((df_dup_check['Name'] == df_dup_check['Name'].shift(-1)) & (df_dup_check['Amount'] == df_dup_check['Amount'].shift(-1)))
            diff_prev = (df_dup_check['Date'] - df_dup_check['Date'].shift(1)).dt.days.abs() <= 7
            diff_next = (df_dup_check['Date'] - df_dup_check['Date'].shift(-1)).dt.days.abs() <= 7
            is_dup = (same_prev & diff_prev) | (same_next & diff_next)
            report2_display = df_dup_check[is_dup][['Date', 'Transaction Type', 'No.', 'Name', 'Account', 'Split', 'Amount']].sort_values(by=['Name', 'Date', 'Amount'])

            # MONTHLY SUMMARY & PARETO
            report1 = df_sorted.pivot_table(index='Name', columns=['Fiscal_Sort', 'Month_Name'], values='Amount', aggfunc='sum').fillna(0)
            report1.columns = [col[1] for col in report1.columns]
            report1['Total Vendor Spend'] = report1.sum(axis=1)
            report1 = report1.sort_values(by='Total Vendor Spend', ascending=False)
            
            total_val = df_filtered['Amount'].sum()
            report1['Cum_Pct'] = (report1['Total Vendor Spend'].cumsum() / total_val) * 100
            core_count = max(1, len(report1[report1['Cum_Pct'] <= 80.0]))
            tail_count = max(0, len(report1) - core_count)
            core_amt = report1['Total Vendor Spend'].iloc[:core_count].sum()
            tail_amt = total_val - core_amt

            pareto_df = pd.DataFrame({
                "Strategic Spend Segment (80/20 Rule)": ["Core Leverage Vendors (Top 80%)", "Long-Tail Vendors (Bottom 20%)"],
                "Count": [core_count, tail_count],
                "Total Segment Spend": [core_amt, tail_amt],
                "Budget %": [f"{(core_amt/total_val)*100:.1f}%" if total_val else "0.0%", f"{(tail_amt/total_val)*100:.1f}%" if total_val else "0.0%"]
            })
            report1 = report1.drop(columns=['Cum_Pct'])
            clean_top = report1.drop(index='[Unassigned/Blank Vendor]', errors='ignore').head(5).reset_index()
            top_5_detailed = clean_top.rename(columns={'Name': 'Vendor', 'Total Vendor Spend': 'Total'})

            # SPIKES
            baseline = df_filtered.groupby('Name')['Amount'].transform('median')
            report3 = df_filtered[df_filtered['Amount'] > (baseline * 1.5)].sort_values(by='Amount', ascending=False)
            report3_display = report3[['Date', 'Transaction Type', 'Name', 'Account', 'Split', 'Amount']]

            # RECONCILIATION & KPIs
            kpi_df = pd.DataFrame({
                "EXECUTIVE SCORECARD": ["Total Filtered Spend", "Total Active Vendors", "Potential Duplicates", "Expense Spikes"],
                "VALUE": [f"${total_val:,.2f}", f"{df_filtered['Name'].nunique()}", f"{len(report2_display)} entries", f"{len(report3)} instances"]
            })
            recon_df = pd.DataFrame({
                "Metric": ["Filtered Ledger Spend"],
                "Source": [f"${total_val:,.2f}"],
                "Output": [f"${report1['Total Vendor Spend'].sum():,.2f}"],
                "Status": ["MATCH"]
            })

            # TREND
            monthly_trend = df_sorted.groupby(['Fiscal_Sort', 'Month_Name'])['Amount'].sum().reset_index()
            monthly_trend['Raw_Pct'] = monthly_trend['Amount'].pct_change()
            trend_display = monthly_trend[['Month_Name', 'Amount']].copy()
            trend_display['Trend Vector'] = monthly_trend['Raw_Pct'].apply(lambda x: "Baseline" if pd.isna(x) else (f"UP +{x*100:.1f}%" if x > 0 else f"DOWN {x*100:.1f}%"))
            trend_display.columns = ['Fiscal Month', 'Total Spend', 'Trend Vector']

            # --- DISPLAY ON WEB SCREEN ---
            st.success("✅ Analysis Complete! Previewing core metrics below:")
            col1, col2 = st.columns(2)
            with col1:
                st.dataframe(kpi_df, use_container_width=True, hide_index=True)
                st.dataframe(top_5_detailed, use_container_width=True, hide_index=True)
            with col2:
                st.dataframe(pareto_df, use_container_width=True, hide_index=True)
                st.dataframe(gl_summary, use_container_width=True, hide_index=True)

            # --- COMPILE EXCEL MEMORY BUFFER FOR DOWNLOAD ---
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                kpi_df.to_excel(writer, sheet_name="Executive Dashboard", index=False, startrow=1, startcol=1)
                pareto_df.to_excel(writer, sheet_name="Executive Dashboard", index=False, startrow=1, startcol=5)
                top_5_detailed.to_excel(writer, sheet_name="Executive Dashboard", index=False, startrow=7, startcol=1)
                gl_summary.to_excel(writer, sheet_name="Executive Dashboard", index=False, startrow=7, startcol=5)
                trend_display.to_excel(writer, sheet_name="Executive Dashboard", index=False, startrow=15, startcol=1)
                controls_scorecard.to_excel(writer, sheet_name="Executive Dashboard", index=False, startrow=15, startcol=5)
                movers_df.to_excel(writer, sheet_name="Executive Dashboard", index=False, startrow=24, startcol=1)
                new_vendors_summary.to_excel(writer, sheet_name="Executive Dashboard", index=False, startrow=24, startcol=5)

                recon_df.to_excel(writer, sheet_name="Recon", index=False)
                report1.to_excel(writer, sheet_name="R1-Summary")
                report2_display.to_excel(writer, sheet_name="R2-Duplicates", index=False)
                report3_display.to_excel(writer, sheet_name="R3-Spikes", index=False)
                leakage_report.to_excel(writer, sheet_name="R4-Weekend", index=False)

            # Apply OpenPyXL styles safely inside memory buffer
            output_buffer.seek(0)
            wb = openpyxl.load_workbook(output_buffer)
            navy_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            zebra_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            w_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            t_border = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'), top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))

            for sheet in wb.sheetnames:
                ws = wb[sheet]
                ws.views.sheetView[0].showGridLines = True
                
                if sheet == "Executive Dashboard":
                    def format_block(start_r, start_c, df_ref):
                        for c in range(start_c, start_c + len(df_ref.columns)):
                            cell = ws.cell(row=start_r + 1, column=c)
                            cell.fill = navy_fill
                            cell.font = w_font
                        for r in range(start_r + 2, start_r + 2 + len(df_ref)):
                            for c in range(start_c, start_c + len(df_ref.columns)):
                                val = str(ws.cell(row=start_r+1, column=c).value).lower()
                                if any(k in val for k in ['spend', 'total', 'exposure', 'variance', 'initial']):
                                    ws.cell(row=r, column=c).number_format = '$#,##0.00'
                                    ws.cell(row=r, column=c).alignment = Alignment(horizontal='right')
                    
                    format_block(1, 2, kpi_df)
                    format_block(1, 6, pareto_df)
                    format_block(7, 2, top_5_detailed)
                    format_block(7, 6, gl_summary)
                    format_block(15, 2, trend_display)
                    format_block(15, 6, controls_scorecard)
                    format_block(24, 2, movers_df)
                    format_block(24, 6, new_vendors_summary)

                    chart = BarChart()
                    chart.type, chart.style, chart.title = "col", 10, "Monthly Spend Trajectory"
                    chart.y_axis.title, chart.height, chart.width, chart.legend = "Outflow ($)", 12, 16, None
                    chart.add_data(Reference(ws, min_col=3, min_row=16, max_row=16+len(trend_display)), titles_from_data=True)
                    chart.set_categories(Reference(ws, min_col=2, min_row=17, max_row=16+len(trend_display)))
                    ws.add_chart(chart, "L1")
                else:
                    ws.freeze_panes = "A2"
                    for cell in ws[1]:
                        cell.fill, cell.font = navy_fill, w_font
                    for row in range(2, ws.max_row + 1):
                        if row % 2 == 0:
                            for col in range(1, ws.max_column + 1):
                                if ws.cell(row=row, column=col).fill.fill_type is None:
                                    ws.cell(row=row, column=col).fill = zebra_fill
                        for col in range(1, ws.max_column + 1):
                            cell = ws.cell(row=row, column=col)
                            cell.border = t_border
                            header = str(ws.cell(row=1, column=col).value).lower()
                            if any(k in header for k in ['amount', 'spend', 'total']):
                                cell.number_format = '$#,##0.00'
                            if 'date' in header and isinstance(cell.value, pd.Timestamp):
                                cell.number_format = 'yyyy-mm-dd'

                for col in ws.columns:
                    max_len = 0
                    col_letter = get_column_letter(col[0].column)
                    for cell in col:
                        if cell.value: max_len = max(max_len, len(str(cell.value)))
                    ws.column_dimensions[col_letter].width = max(max_len + 2, 12)

            final_buffer = io.BytesIO()
            wb.save(final_buffer)
            final_buffer.seek(0)

            # --- DOWNLOAD BUTTON ---
            st.write("---")
            st.download_button(
                label="📥 Download Structured Analysis Report (Excel)",
                data=final_buffer,
                file_name="Analysis_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"An unexpected error occurred during processing: {e}")