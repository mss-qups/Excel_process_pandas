import pandas as pd

# load data
file_path = "VMCR_load_report_100_participant_8_10"
file = pd.read_excel(file_path+".xlsx", engine='openpyxl')
data_part = file[["Input", "Output", "MFI", "MFO", "Input(MF)", "Output(MF)"]]

# converts lists to pandas dataframe
input_list_df = pd.DataFrame([float(x.strip()[:-2]) * 1024 if x.strip()[-2:] == "Mb" else float(x.strip()[:-2]) for x in
                              list(data_part["Input"]) if str(x) != 'nan'])
output_list_df = pd.DataFrame(
    [float(x.strip()[:-2]) * 1024 if x.strip()[-2:] == "Mb" else float(x.strip()[:-2]) for x in
     list(data_part["Output"]) if str(x) != 'nan'])
mfi_list_df = pd.DataFrame([x for x in list(data_part["MFI"]) if str(x) != 'nan'])
mfo_list_df = pd.DataFrame([x for x in list(data_part["MFO"]) if str(x) != 'nan'])

input_mf = input_list_df * mfi_list_df
output_mf = output_list_df * mfo_list_df
total = input_mf + output_mf

# write to excel
file["Input(MF)"] = list(input_mf[0])
file["Output(MF)"] = list(output_mf[0])
file["Total"] = list(total[0])
file.to_excel(f"./{file_path}new.xlsx", index=False)
