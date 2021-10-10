import pandas as pd
from openpyxl import load_workbook
import xlsxwriter

file_path = "VMCR_load_report_100_participant_8_10"
wb = load_workbook(file_path + '.xlsx')
xls = pd.ExcelFile(file_path + '.xlsx')
sheets = wb.sheetnames

dfs = []
for sheet in sheets:
    df = pd.read_excel(xls, sheet)
    if "MFI" and "MFO" in list(df.head()):
        input_list_df = pd.DataFrame(
            [float(x.strip()[:-2]) * 1024 if x.strip()[-2:] == "Mb" else float(x.strip()[:-2]) for x in
             list(df["Input"]) if str(x) != 'nan'])
        output_list_df = pd.DataFrame(
            [float(x.strip()[:-2]) * 1024 if x.strip()[-2:] == "Mb" else float(x.strip()[:-2]) for x in
             list(df["Output"]) if str(x) != 'nan'])
        mfi_list_df = pd.DataFrame([x for x in list(df["MFI"]) if str(x) != 'nan'])
        mfo_list_df = pd.DataFrame([x for x in list(df["MFO"]) if str(x) != 'nan'])

        input_mf = input_list_df * mfi_list_df
        output_mf = output_list_df * mfo_list_df
        total = input_mf + output_mf

        df["Input(MF)"] = list(input_mf[0])
        df["Output(MF)"] = list(output_mf[0])
        df["Total"] = list(total[0])
        dfs.append(df)
    else:
        dfs.append(pd.read_excel(xls, sheet))

        # # write to excel
        # df.to_excel(f"./DATA_SS/{file_path}_processed.xlsx", index=False)
        # writer = pd.ExcelWriter(f'./DATA_SS/{file_path}_processed.xlsx', engine='xlsxwriter')
        # df.to_excel(writer, sheet_name=sheet)
        # writer.save()
        # Create a Pandas Excel writer using XlsxWriter as the engine.
print(len(dfs))
Excelwriter = pd.ExcelWriter("languages_multiple.xlsx", engine="xlsxwriter")
for i, df in enumerate(dfs):
    df.to_excel(Excelwriter, sheet_name=sheets[i], index=False)
Excelwriter.save()
