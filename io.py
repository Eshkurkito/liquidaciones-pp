import pandas as pd
from io import BytesIO

def df_to_excel_bytes(dfs: dict) -> bytes:
    """
    dfs: mapping sheet_name -> dataframe
    returns Excel file bytes
    """
    with BytesIO() as b:
        with pd.ExcelWriter(b, engine="openpyxl") as writer:
            for name, df in dfs.items():
                df.to_excel(writer, sheet_name=str(name)[:31], index=False)
        return b.getvalue()