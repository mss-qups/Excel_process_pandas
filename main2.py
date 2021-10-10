import pandas as pd
import glob

# files = glob.glob('*.xlsx')
# for f in files:
# load data
file_path = "VMCR_4R_70"
# filename = f[:-5]
file = pd.read_excel(file_path + ".xlsx", engine='openpyxl')
data_part = file[["Input", "Output", "MFI", "MFO", "Input(MF)", "Output(MF)"]]

# converts lists to pandas dataframe
input_list_df = pd.DataFrame([float(x.strip()[:-2]) * 1024 if x.strip()[-2:] == "Mb" else float(x.strip()[:-2]) for x in
                              list(data_part["Input"]) if str(x) != 'nan'])
output_list_df = pd.DataFrame(
    [float(x.strip()[:-2]) * 1024 if x.strip()[-2:] == "Mb" else float(x.strip()[:-2]) for x in
     list(data_part["Output"]) if str(x) != 'nan'])
mfi_list_df = pd.DataFrame([x for x in list(data_part["MFI"]) if str(x) != 'nan'])
mfo_list_df = pd.DataFrame([x for x in list(data_part["MFO"]) if str(x) != 'nan'])


# input_mf = input_list_df * mfi_list_df
# output_mf = output_list_df * mfo_list_df
# total = input_mf + output_mf

# file["Input(MF)"] = list(input_mf[0])
# file["Output(MF)"] = list(output_mf[0])
# file["Total"] = list(total[0])


# write to excel

# file.to_excel(f"./DATA_SS/{file_path}_processed.xlsx", index=False)

l = [float(x.strip()[:-2]) * 1024 if x.strip()[-2:] == "Mb" else float(x.strip()[:-2]) for x in
     list(data_part["Input"]) if str(x) != 'nan']
t = [float(x.strip()[:-2]) * 1024 if x.strip()[-2:] == "Mb" else float(x.strip()[:-2]) for x in
     list(data_part["Output"]) if str(x) != 'nan']
mfi = [float(x) for x in list(data_part["MFI"]) if str(x) != 'nan']
mfo = [float(x) for x in list(data_part["MFO"]) if str(x) != 'nan']
# for i in l:
print(len(l))
print(len(t))
print(len(mfi))
print(len(mfo))

print(type(l[0]))

print(all(isinstance(x, float) for x in mfo))
