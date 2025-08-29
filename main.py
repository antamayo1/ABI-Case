import streamlit as st
import pandas as pd
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import coordinate_to_tuple
from io import BytesIO

fill = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
sec_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")

st.set_page_config(page_title="Design Transformer", layout="wide")

def box(worksheet, start_cell, end_cell, border_style="thin"):
  side = Side(border_style=border_style, color="000000")
  start_row, start_col = coordinate_to_tuple(start_cell)
  end_row, end_col = coordinate_to_tuple(end_cell)
  min_row, max_row = min(start_row, end_row), max(start_row, end_row)
  min_col, max_col = min(start_col, end_col), max(start_col, end_col)
  for row in range(min_row, max_row + 1):
    for col in range(min_col, max_col + 1):
      cell = worksheet.cell(row=row, column=col)
      cell_border = Border(
        left=side if col == min_col else None,
        right=side if col == max_col else None,
        top=side if row == min_row else None,
        bottom=side if row == max_row else None,
      )
      cell.border = Border(
        left=cell_border.left or cell.border.left,
        right=cell_border.right or cell.border.right,
        top=cell_border.top or cell.border.top,
        bottom=cell_border.bottom or cell.border.bottom,
      )
  return worksheet

def box_fill(worksheet, start_cell, end_cell, border_style="thin"):
  side = Side(border_style=border_style, color="000000")
  border = Border(left=side, right=side, top=side, bottom=side)

  start_row, start_col = coordinate_to_tuple(start_cell)
  end_row, end_col = coordinate_to_tuple(end_cell)
  min_row, max_row = min(start_row, end_row), max(start_row, end_row)
  min_col, max_col = min(start_col, end_col), max(start_col, end_col)

  for row in range(min_row, max_row + 1):
    for col in range(min_col, max_col + 1):
      cell = worksheet.cell(row=row, column=col)
      cell.border = border
  return worksheet

def columnRowDimensions(worksheet):
  worksheet.column_dimensions['A'].width = 14.78 * 1.4
  worksheet.column_dimensions['B'].width = 4.56 * 1.4
  worksheet.column_dimensions['C'].width = 20.56 * 1.4
  worksheet.column_dimensions['D'].width = 6.89 * 1.4
  worksheet.column_dimensions['E'].width = 18.11 * 1.4
  worksheet.column_dimensions['F'].width = 25.33 * 1.4
  worksheet.column_dimensions['G'].width = 19.89 * 1.4
  worksheet.column_dimensions['H'].width = 10.11 * 1.4
  worksheet.column_dimensions['I'].width = 9.56 * 1.4
  worksheet.column_dimensions['J'].width = 11.11 * 1.4
  worksheet.column_dimensions['K'].width = 40.11 * 1.4
  worksheet.row_dimensions[1].height = 18
  worksheet.row_dimensions[2].height = 18
  worksheet.row_dimensions[3].height = 18
  worksheet.row_dimensions[4].height = 18
  return worksheet

def headerBorders(worksheet):
  worksheet = box(worksheet, 'A1', 'B4')
  worksheet = box(worksheet, 'C1', 'C4')
  worksheet = box(worksheet, 'D1', 'F2')
  worksheet = box(worksheet, 'D3', 'F4')
  worksheet = box(worksheet, 'G1', 'G4')
  worksheet = box(worksheet, 'H1', 'H2')
  worksheet = box(worksheet, 'H3', 'H4')
  worksheet = box_fill(worksheet, 'I1', 'K4')
  return worksheet

def addHeaderValues(worksheet):
  worksheet.merge_cells('A1:B4')
  worksheet['A1'] = "LOGO"
  worksheet['A1'].font = Font(name='Avenir Book', size=30, bold=True)
  worksheet['A1'].alignment = Alignment(horizontal='center', vertical='center')

  worksheet.merge_cells('E1:F2')
  worksheet['E1'] = "45 CLAPBOARD HILL ROAD"
  worksheet['E1'].font = Font(name='Avenir Book', size=14, bold=True)
  worksheet['E1'].alignment = Alignment(horizontal='center', vertical='center')

  worksheet.merge_cells('E3:F4')
  worksheet['E3'] = "PLUMBING & BATH ACCESSORIES SCHEDULE"
  worksheet['E3'].font = Font(name='Avenir Book', size=12, bold=True)
  worksheet['E3'].alignment = Alignment(horizontal='center', vertical='center')

  worksheet.merge_cells('G2:G4')
  worksheet['G2'] = "S1100"
  worksheet['G2'].font = Font(name='Avenir Book', size=20, bold=True)
  worksheet['G2'].alignment = Alignment(horizontal='center', vertical='center')

  worksheet['C1'] = "63 Wingold Ave, Suite 208"
  worksheet['C1'].font = Font(name='Avenir Book', size=8)
  worksheet['C1'].alignment = Alignment(horizontal='center', vertical='center')

  worksheet['C2'] = "Toronto, ON, M6B 1P8"
  worksheet['C2'].font = Font(name='Avenir Book', size=8)
  worksheet['C2'].alignment = Alignment(horizontal='center', vertical='center')

  worksheet['C4'] = "alibuddinteriors.com"
  worksheet['C4'].font = Font(name='Avenir Book', size=8)
  worksheet['C4'].alignment = Alignment(horizontal='center', vertical='center')

  worksheet['D1'] = "PROJECT"
  worksheet['D1'].font = Font(name='Avenir Book', size=6)
  worksheet['D1'].alignment = Alignment(horizontal='left', vertical='center')

  worksheet['D3'] = "TITLE"
  worksheet['D3'].font = Font(name='Avenir Book', size=6)
  worksheet['D3'].alignment = Alignment(horizontal='left', vertical='center')

  worksheet['G1'] = "NO."
  worksheet['G1'].font = Font(name='Avenir Book', size=6)
  worksheet['G1'].alignment = Alignment(horizontal='left', vertical='center')

  worksheet['H1'] = "BY"
  worksheet['H1'].font = Font(name='Avenir Book', size=6)
  worksheet['H1'].alignment = Alignment(horizontal='left', vertical='center')

  worksheet['H2'] = "--"
  worksheet['H2'].font = Font(name='Avenir Book', size=8)
  worksheet['H2'].alignment = Alignment(horizontal='center', vertical='center')

  worksheet['H3'] = "DATE"
  worksheet['H3'].font = Font(name='Avenir Book', size=6)
  worksheet['H3'].alignment = Alignment(horizontal='left', vertical='center')

  worksheet['H4'] = "2024-06-14"
  worksheet['H4'].font = Font(name='Avenir Book', size=8)
  worksheet['H4'].alignment = Alignment(horizontal='center', vertical='center')

  worksheet['I1'] = "VERSION"
  worksheet['I1'].font = Font(name='Avenir Book', size=8)
  worksheet['I1'].alignment = Alignment(horizontal='left', vertical='center')

  worksheet['J1'] = "DATE"
  worksheet['J1'].font = Font(name='Avenir Book', size=8)
  worksheet['J1'].alignment = Alignment(horizontal='left', vertical='center')

  worksheet['K1'] = "DESCRIPTION"
  worksheet['K1'].font = Font(name='Avenir Book', size=8)
  worksheet['K1'].alignment = Alignment(horizontal='left', vertical='center')

  return worksheet

def mainTableHeaders(worksheet):
  worksheet = box_fill(worksheet, 'A6', 'K6')
  
  worksheet['A6'] = 'LOCATION'
  worksheet['A6'].font = Font(name='Avenir Book', size=8, bold=True)
  worksheet['A6'].alignment = Alignment(horizontal='center', vertical='center')
  worksheet['A6'].fill = fill

  worksheet['B6'] = 'ITEM #'
  worksheet['B6'].font = Font(name='Avenir Book', size=8, bold=True)
  worksheet['B6'].alignment = Alignment(horizontal='center', vertical='center')
  worksheet['B6'].fill = fill

  worksheet['C6'] = 'ITEM'
  worksheet['C6'].font = Font(name='Avenir Book', size=8, bold=True)
  worksheet['C6'].alignment = Alignment(horizontal='center', vertical='center')
  worksheet['C6'].fill = fill

  worksheet['D6'] = 'QTY'
  worksheet['D6'].font = Font(name='Avenir Book', size=8, bold=True)
  worksheet['D6'].alignment = Alignment(horizontal='center', vertical='center')
  worksheet['D6'].fill = fill

  worksheet['E6'] = 'IMAGE'
  worksheet['E6'].font = Font(name='Avenir Book', size=8, bold=True)
  worksheet['E6'].alignment = Alignment(horizontal='center', vertical='center')
  worksheet['E6'].fill = fill

  worksheet['F6'] = 'MODEL'
  worksheet['F6'].font = Font(name='Avenir Book', size=8, bold=True)
  worksheet['F6'].alignment = Alignment(horizontal='center', vertical='center')
  worksheet['F6'].fill = fill

  worksheet['G6'] = 'DIMENSIONS'
  worksheet['G6'].font = Font(name='Avenir Book', size=8, bold=True)
  worksheet['G6'].alignment = Alignment(horizontal='center', vertical='center')
  worksheet['G6'].fill = fill

  worksheet['H6'] = 'FINISH'
  worksheet['H6'].font = Font(name='Avenir Book', size=8, bold=True)
  worksheet['H6'].alignment = Alignment(horizontal='center', vertical='center')
  worksheet['H6'].fill = fill

  worksheet.merge_cells('I6:J6')
  worksheet['I6'] = 'SUPPLIER'
  worksheet['I6'].font = Font(name='Avenir Book', size=8, bold=True)
  worksheet['I6'].alignment = Alignment(horizontal='center', vertical='center')
  worksheet['I6'].fill = fill

  worksheet['K6'] = 'NOTES'
  worksheet['K6'].font = Font(name='Avenir Book', size=8, bold=True)
  worksheet['K6'].alignment = Alignment(horizontal='center', vertical='center')
  worksheet['K6'].fill = fill

  return worksheet

def createHeader(worksheet):
  worksheet = columnRowDimensions(worksheet)
  worksheet = headerBorders(worksheet)
  worksheet = addHeaderValues(worksheet)
  worksheet = mainTableHeaders(worksheet)
  return worksheet

def addMainTable(worksheet, rooms):
  row = 7
  pl = 1
  for main, subs in rooms.items():
    worksheet.merge_cells(f'A{row}:K{row}')
    worksheet[f'A{row}'] = main
    worksheet[f'A{row}'].alignment = Alignment(horizontal='left', vertical='center')
    worksheet[f'A{row}'].font = Font(name='Avenir Book', size=8, bold=True)
    worksheet[f'A{row}'].fill = sec_fill
    row += 1
    for sub in subs:
      start = row
      worksheet[f'A{row}'] = sub
      worksheet[f'A{row}'].alignment = Alignment(horizontal='center', vertical='center')
      worksheet[f'A{row}'].font = Font(name='Avenir Book', size=8)
      for product_type in subs[sub]:

        model_string = f'''BRAND: {st.session_state.details.iloc[pl-1].loc['Brand']}
NAME: {st.session_state.details.iloc[pl-1].loc['Product Name']}
SKU: {st.session_state.details.iloc[pl-1].loc['Product Code #']}'''

        worksheet[f'F{row}'] = model_string
        worksheet[f'F{row}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        worksheet[f'F{row}'].font = Font(name='Avenir Book', size=8)


        worksheet[f'D{row}'] = st.session_state.details.iloc[pl-1].loc['QTY (per Area)']
        worksheet[f'D{row}'].alignment = Alignment(horizontal='center', vertical='center')
        worksheet[f'D{row}'].font = Font(name='Avenir Book', size=8)

        worksheet[f'H{row}'] = st.session_state.details.iloc[pl-1].loc['Finish/Color']
        worksheet[f'H{row}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        worksheet[f'H{row}'].font = Font(name='Avenir Book', size=8)

        worksheet[f'G{row}'] = st.session_state.details.iloc[pl-1].loc['Dimension']
        worksheet[f'G{row}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        worksheet[f'G{row}'].font = Font(name='Avenir Book', size=8)

        worksheet.merge_cells(f'I{row}:J{row}')
        worksheet[f'I{row}'] = st.session_state.details.iloc[pl-1].loc['Supplier']
        worksheet[f'I{row}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        worksheet[f'I{row}'].font = Font(name='Avenir Book', size=8)

        worksheet[f'B{row}'] = f'PL-{pl}'
        worksheet[f'B{row}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        worksheet[f'B{row}'].font = Font(name='Avenir Book', size=8)

        worksheet[f'K{row}'] = 'SEE ID DRAWINGS & SPEC SHEETS FOR MORE INFORMATION'
        worksheet[f'K{row}'].alignment = Alignment(horizontal='center', vertical='center')
        worksheet[f'K{row}'].font = Font(name='Avenir Book', size=8, italic=True)

        worksheet[f'C{row}'] = product_type
        worksheet[f'C{row}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        worksheet[f'C{row}'].font = Font(name='Avenir Book', size=8)
        worksheet.row_dimensions[row].height = 75
        pl += 1
        row += 1
      worksheet.merge_cells(f'A{start}:A{row-1}')
  worksheet = box_fill(worksheet, f'A7', f'K{row-1}')
  return worksheet

if not st.user.is_logged_in:
  with st.container(horizontal_alignment="center"):
    with st.container(border=True, horizontal_alignment="center", width=500):
      st.image("LMC_Logo.jpeg")
      st.title("Export to Schedule", width="content", anchor=False)
      st.caption("Please Login to Continue", width="content")
      login_btn = st.button(
        "**Log in** with **Google**",
        use_container_width=True,
        type="primary"
      )
      if login_btn:
        st.login()
    st.markdown("""
    <div style="text-align: center; color: #6c757d; font-size: 14px; margin-top: 30px;">
      <p style="margin-bottom: 0px;"><strong>Export to Formatted Schedule Application</strong></p>
      <a style="margin-bottom: 0px; text-decoration: none;" href="https://lernmoreconsulting.com">© 2025 LernMore Consulting</a>
      <p>Secure • Fast • Accurate</p>
    </div>
    """, unsafe_allow_html=True)
else:
  user_details = st.user.to_dict()
  with st.container(border=True, horizontal=True, horizontal_alignment="distribute", height=73):
    with st.container(vertical_alignment="center", height="stretch"):
      st.write("**Export to Schedule**")
    logout_btn = st.button(
      "Log out",
      type="primary"
    )
    if logout_btn:
      st.logout()
  with st.expander("Upload Schedule", expanded=True):
    st.session_state.input_file = st.file_uploader("Please input the raw export file", type=["xlsx", "xls"])
  if st.session_state.input_file:
    st.session_state.details = pd.read_excel(st.session_state.input_file, header=9)
    st.session_state.rooms = {}
    for idx, roomName in enumerate(st.session_state.details['Area']):
      [main, sub] = roomName.split(' / ')
      main_key = main[4:]
      sub_key = sub[4:]
      product = st.session_state.details['Product Type'][idx]
      if main_key not in st.session_state.rooms:
        st.session_state.rooms[main_key] = {sub_key: [product]}
      else:
        if sub_key not in st.session_state.rooms[main_key]:
          st.session_state.rooms[main_key][sub_key] = [product]
        else:
          st.session_state.rooms[main_key][sub_key].append(product)
    st.write(st.session_state.details)
    newDataframe = pd.DataFrame()
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
      newDataframe.to_excel(writer, index=False, sheet_name='PLUMBING')
      worksheet = writer.sheets['PLUMBING']
      worksheet.font = Font(name='Avenir Book')
      worksheet = createHeader(worksheet)
      worksheet = addMainTable(worksheet, st.session_state.rooms)
    output.seek(0)
    st.download_button(
      label="Download Formatted Excel File",
      data=output,
      file_name="test.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
  else:
    st.info("Please upload a schedule file to proceed.")