import traceback
from excel_db import get_raw_sheet

try:
    data = get_raw_sheet('../data/SWESch_FSL_2122_HYD_OR007_PLC.xlsx')
    print('SUCCESS - keys:', list(data.keys()))
except Exception:
    traceback.print_exc()
