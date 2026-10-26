"""Launch: streamlit run app.py"""
from pathlib import Path
from datetime import date
import io
import streamlit as st
from engine import read_upload,analyze,build

st.set_page_config(page_title='AP Workload & Performance',page_icon='📊',layout='wide')
st.title('AP Workload & Performance Dashboard')
st.caption('Upload the latest 19-field invoice export to generate the original ten-tab Excel analysis dashboard.')
st.info('Source of truth: Master Data. Paid payment-cycle eligibility requires Paid status and a valid nonnegative date difference. Charts regenerate to the chosen reporting month.')
source=st.file_uploader('Invoice source workbook (.xlsx)',type=['xlsx'])
if source:
 try:
  with st.spinner('Reading invoice records…'):
   rows=read_upload(source.getvalue())
  dates=[r[5] for r in rows if r[5]]
  guessed=max(dates) if dates else date.today()
  st.write(f'**Records loaded:** {len(rows):,}')
  report_date=st.date_input('Reporting as-of date',value=guessed,help='Used for aging, historical comparisons, and latest month. Set it to your official report cutoff date.')
  summary=analyze(rows,report_date)
  c1,c2,c3,c4=st.columns(4)
  c1.metric('Invoices',f"{summary['count']:,}")
  c2.metric('Paid',f"{summary['paid']:,}")
  c3.metric('Eligible payment cycles',f"{summary['eligible']:,}")
  c4.metric('Median paid cycle days',f"{summary['median']:,.1f}")
  if summary['exceptions']:
   st.warning('Data quality exclusions: '+', '.join(f'{k}: {v:,}' for k,v in summary['exceptions'].items()))
  if len(rows)>20000:st.error('The reference template supports 20,000 invoices. Reduce the batch or extend the template capacity.')
  elif st.button('Generate Excel dashboard',type='primary'):
   with st.spinner('Updating workbook formulas, Master Data, and chart series…'):
    payload,stats,ms=build((Path(__file__).parent/'assets'/'reference_dashboard.xlsx').read_bytes(),rows,report_date)
   st.success(f'Excel report generated. Chart month coverage: {ms["periods"][0][0]}-{ms["periods"][0][1]:02d} through {ms["periods"][-1][0]}-{ms["periods"][-1][1]:02d}.')
   st.download_button('Download completed dashboard',data=payload,file_name=f'AP_Workload_Performance_Dashboard_{report_date:%m.%d.%Y}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
 except Exception as ex:
  st.error(f'Could not process file: {ex}')
else:
 st.markdown('**Accepted source fields:** ID, Status, Invoice Number, Vendor, Invoice Date, Date Received, Amount, Aging Received, Urgent, Category, Fiscal Year, PO Number, Department, Fund, Invoice Approver, Invoice Approver 2, AP Payer, PS Entry Date, Voucher Number.')
