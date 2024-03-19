import os
import shutil

import pandas as pd
import matplotlib.pyplot as plt
import dropbox
import dropbox.files
import base64
import requests
import json


# INPUT: .DTA DTAFiles within the DTAFiles directory
# OUTPUT: csv with the lowest conductivity extracted from each .DTA file
# Made for retrieving the lowest conductivity point from the gamry echleon test results
def retrieve_lowest_conductivty():


    path = '/Users/Neware3/PycharmProjects/script'
    folder = os.listdir("DTAFiles")
    lowest_values = {"Test_Name": [], "Conductivity": []}
    for file in folder:
        filepath = path + '/DTAFiles/' + file
        if os.path.isfile(filepath):
            file1 = open(filepath, 'r')
            lines = file1.readlines()


            for i in range(len(lines)):
                if 'ZCURVE' in lines[i]:
                    start_line = i + 3
                    break

            zcurve = {'ZReal': [], 'ZImag': []}
            x_value = lines[start_line].split('\t')[4]
            for k in range(31):
                zcurve['ZReal'].append(lines[start_line + k].split('\t')[4])
                zcurve['ZImag'].append(lines[start_line + k].split('\t')[5])

            # zcurve_df = pd.DataFrame(zcurve)
            # zcurve_df.plot(kind='scatter',
            #         x='ZReal',
            #         y='ZImag',
            #         color='red')
            # plt.title('ScatterPlot')
            # plt.show()

            # write the lowest x_values into a dataframe ################
            name = filepath.split('.DTA')[0]
            name = name.split('/')[-1]
            lowest_values["Test_Name"].append(name)
            lowest_values["Conductivity"].append(x_value)

    # Grab thickness values from existing excel and sort the data by Test_Name
    xl = pd.ExcelFile(path + "/WorkDir/input_conductivity_calculation.xlsx")
    df = xl.parse("Sheet1")
    df = df.sort_values(by="Test_Name")

    writer = pd.ExcelWriter(path + '/WorkDir/sorted_conductivity_calculations.xlsx')
    df.to_excel(writer, sheet_name='Sheet1', columns=["Test_Name", "thickness"],
                    index=False)
    writer._save()

    xl = pd.ExcelFile(path + "/WorkDir/sorted_conductivity_calculations.xlsx")
    df = xl.parse("Sheet1")


    # Write calculations into results and calculations dataframe
    print(df)
    new_conductivity_results = {"Test_Name": [], 'thickness': [], "calculation": []}
    new_conductivity_calculations = {"Test_Name": [], 'thickness': [], 'conductivity': [], "calculation": []}
    for i in range(len(df["Test_Name"])):
        new_conductivity_results["Test_Name"].append(df["Test_Name"][i])
        new_conductivity_calculations["Test_Name"].append(df["Test_Name"][i])
        new_conductivity_results['thickness'].append(df['thickness'][i])
        new_conductivity_calculations['thickness'].append(df['thickness'][i])
        new_conductivity_calculations['conductivity'].append(lowest_values['Conductivity'][i])
        calc = 10000000/int(df['thickness'][i])/float(lowest_values["Conductivity"][i])
        new_conductivity_calculations['calculation'].append(calc)
        calc = round(calc,1)
        new_conductivity_results['calculation'].append(calc)

    ##################################################
    # write the data into Excel DTAFiles:
    # new_conductivity_results, new_conductivity_calculations
    results_df = pd.DataFrame(new_conductivity_results)
    writer = pd.ExcelWriter('output/new_conductivity_results.xlsx')
    results_df.to_excel(writer, 'Sheet1', index=False)
    writer._save()

    calculations_df = pd.DataFrame(new_conductivity_calculations)
    writer = pd.ExcelWriter('output/new_conductivity_calculations.xlsx')
    calculations_df.to_excel(writer, 'Sheet1', index=False)
    writer._save()


def download_from_dropbox():
    # code to retrieve DTAFiles from dropbox ########################
    with open('appkey', 'r') as f:
        APP_KEY = f.read()
    with open('appsecret', 'r') as f:
        APP_SECRET = f.read()
    with open('refreshtoken', 'r') as f:
        REFRESH_TOKEN = f.read()
    dbx = dropbox.Dropbox(
        app_key=APP_KEY,
        app_secret =APP_SECRET,
        oauth2_refresh_token =REFRESH_TOKEN
    )



    # remove files from DTAFiles folder before downloading more
    if len(os.listdir('DTAFiles')) != 0:
        folder = 'DTAFiles'
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print('Failed to delete %s. Reason: %s' % (file_path, e))

    for entry in dbx.files_list_folder("").entries:
        if entry.name != "Results":
            if entry.name == "input_conductivity_calculation.xlsx":
                dbx.files_download_to_file(os.path.join("WorkDir", entry.name), f"/{entry.name}")
            else:
                dbx.files_download_to_file(os.path.join("DTAFiles", entry.name), f"/{entry.name}")




def upload_to_dropbox(path):
    with open('appkey', 'r') as f:
        APP_KEY = f.read()
    with open('appsecret', 'r') as f:
        APP_SECRET = f.read()
    with open('refreshtoken', 'r') as f:
        REFRESH_TOKEN = f.read()
    dbx = dropbox.Dropbox(
        app_key=APP_KEY,
        app_secret=APP_SECRET,
        oauth2_refresh_token=REFRESH_TOKEN
    )
    for file in os.listdir(path):
        with open(os.path.join(path, file), "rb") as f:
            data = f.read()
            dbx.files_upload(data, f"/Results/{file}")

def dropbox_refresh_token():
    with open('appkey', 'r') as f:
        APP_KEY = f.read()
    with open('appsecret', 'r') as f:
        APP_SECRET = f.read()
    with open('accesscode', 'r') as f:
        ACCESS_CODE_GENERATED = f.read()


    BASIC_AUTH = base64.b64encode(f'{APP_KEY}:{APP_SECRET}'.encode())

    headers = {
        'Authorization': f"Basic {BASIC_AUTH.decode()}",
        'Content-Type': 'application/x-www-form-urlencoded',
    }

    data = f'code={ACCESS_CODE_GENERATED}&grant_type=authorization_code'

    response = requests.post('https://api.dropboxapi.com/oauth2/token',
                             data=data,
                             auth=(APP_KEY, APP_SECRET))
    print(json.dumps(json.loads(response.text), indent=2))

def main():
    print(len(os.listdir('DTAFiles')))

if __name__ == "__main__":
    #main()
    download_from_dropbox()
    #retrieve_lowest_conductivty()
    #upload_to_dropbox("output")
    #dropbox_refresh_token()
