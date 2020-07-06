import pandas as pd
from tqdm import tqdm


def export(df, export_name, interval=3000):
    with pd.ExcelWriter(f'Spawn/exported_{export_name}.xlsx', engine='openpyxl') as writer:
        for i in tqdm(range(0, len(df), interval)):
            if i == 0:
                if i + interval <= len(df):
                    df.iloc[i:i + interval].to_excel(
                        writer, sheet_name='exported', index=False, header=True)
                else:
                    df.iloc[i:len(df)].to_excel(
                        writer, sheet_name='exported', index=False, header=True)
            else:
                if i + interval <= len(df):
                    df.iloc[i:i + interval].to_excel(
                        writer, sheet_name='exported', index=False, header=False, startrow=i + 1)
                else:
                    df.iloc[i:len(df)].to_excel(
                        writer, sheet_name='exported', index=False, header=False, startrow=i + 1)