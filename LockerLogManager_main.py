import os
import time
import requests
import sys
import pandas as pd
import datetime


class LockerLogManager:
    """Class LockerLogManager - automatically downloads google datasheet and parses it for better readability"""

    # Prepare variables
    log_dict = {}
    balance_dict = {}
    x = 1
    temp_dict = {}
    delete_empty_records = []
    record_list = []
    # Declare known items for sorting later on
    guns_list = [
        'Glock 20', 'Glock 19', 'Glock 18 1.0', 'Pistolet MK2', 'Walther P88',
        'Vintage Pistol', 'Ciężki Pistolet', 'Beretta 92FS', 'Combat Pistol',
        'Staccato 2011', 'Glock 19x2', 'CZ-75', 'Beretta M9A3', 'Beretta 98', 'SIG Pistol'
    ]
    mags_item = 'Magazynek do pistoletu '
    # Declare final file name with timestamp
    ct = datetime.datetime.now()
    path_excel = "out/szafka_log_{}_{}_{}.xlsx".format(ct.day, ct.month, ct.year)
    name_csv = "raw_locker_log.csv".format(ct.day, ct.month, ct.year)

    def __init__(self):
        self.download_raw_csv()  # download source document in csv form
        os.makedirs('out/', exist_ok=True)  # create a directory for final file
        self.clean_raw_csv()  # clean downloaded csv from misc characters
        self.split_csv_data()  # split csv's data into dictionaries for better access
        self.calculate_item_IO_balance()  # calculate how many items have been taken/given by the player
        self.clear_null_balance()  # delete records where IO balance was equal to none

        self.generate_exel_log()  # finalize operation by saving a exel log file

    def _getGoogleSheet(self, spreadsheet_id, outDir, outFile):
        url = f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv'
        response = requests.get(url)
        if response.status_code == 200:
            filepath = os.path.join(outDir, outFile)
            with open(filepath, 'wb') as f:
                f.write(response.content)
                print('CSV file saved to: {}'.format(filepath))
        else:
            print(f'Error downloading Google Sheet: {response.status_code}')
            sys.exit(1)

    def download_raw_csv(self, rawDir='tmp/', spreadsheet_id=''):
        os.makedirs(rawDir, exist_ok=True)
        self._getGoogleSheet(spreadsheet_id, rawDir, self.name_csv)

    def clean_raw_csv(self):
        f = open('tmp/raw_locker_log.csv', 'r', encoding='utf-8')
        data_from_csv = f.read()
        f.close()
        data_from_csv = data_from_csv.replace("DUMMY,DUMMY,DUMMY,DUMMY,DUMMY\n", "")
        data_from_csv = data_from_csv.replace("*", "")
        data_from_csv = data_from_csv.replace("\n", "")
        data_from_csv = data_from_csv.replace("\t", "")
        data_from_csv = data_from_csv.replace(",,,,", "|")
        data_from_csv = data_from_csv.replace("\"\"", "-")
        data_from_csv = data_from_csv.replace("\"", "")
        data_from_csv = data_from_csv.replace("               ", "")
        data_from_csv = data_from_csv.replace("    ", "")
        self.record_list = data_from_csv.split("|")

    def split_csv_data(self):
        for record in self.record_list:
            if len(record) < 2:
                break

            amt_item = -1, -1
            name_item = ""
            dodatki = "NULL"
            id_item = ""
            i = 4

            record = record.split(" ")
            imie = record[0]
            nazwisko = record[1]
            id = int(record[2].replace("[", "").replace("]", ""))
            akcja = record[3]

            while 1:
                if record[i].startswith("["):
                    if record[i][1].isalpha():
                        id_item = record[i].replace("[", "").replace("]", "")
                        i += 1
                    elif record[i][1].isdigit():
                        amt_item = int(record[i].replace("[", "").replace("]", ""))  # obsługa błędu jbc do dodania
                        i += 1
                        break
                else:
                    name_item += record[i] + " "
                    i += 1
            if i != len(record):
                dodatki = record[i + 1]

            if len(id_item) > 3:
                amt_item = 1
            # print("id={}\n \timie={} nazwisko={}\n\t\takcja={} item={} id_item={} ilosc={} \n\t\t\tdodatki={}".format(id, imie, nazwisko, akcja,
            #                                                                                          name_item, id_item,
            #                                                                                          amt_item, dodatki
            #                                                                                          ))

            id_dict = str(id) + " " + imie + " " + nazwisko
            if not self.log_dict.get(id_dict):
                self.log_dict[id_dict] = []
            self.log_dict[id_dict].append([name_item, akcja, amt_item])

    def calculate_item_IO_balance(self):
        for person in self.log_dict:
            for action in self.log_dict[person]:
                id = action[0]
                exachnge = action[1]
                amt = action[2]

                if not self.balance_dict.get(id):
                    self.balance_dict[id] = 0  # 200 w szafce za pierwszą akcją itemu

                curr_amt = self.balance_dict[id]
                if exachnge == "Odłożył:":
                    curr_amt += amt
                elif exachnge == "Pobrał:":
                    curr_amt -= amt
                self.balance_dict[id] = curr_amt

            if not self.temp_dict.get(person):
                self.temp_dict[person] = []
            self.temp_dict[person] = self.balance_dict
            self.balance_dict = {}

    def clear_null_balance(self):
        for i in self.temp_dict:
            for j in self.temp_dict[i]:
                if self.temp_dict[i][j] == 0:
                    self.delete_empty_records.append([i, j])

        for nr in self.delete_empty_records:
            del self.temp_dict[nr[0]][nr[1]]

    def generate_exel_log(self):
        writer = pd.ExcelWriter(self.path_excel, engine='xlsxwriter')

        for person, items in self.temp_dict.items():
            # Create a DataFrame for items and values
            df_items = pd.DataFrame(list(items.items()), columns=['Item', 'Value'])

            # Add additional columns for guns and mags
            df_items['Guns'] = df_items.apply(
                lambda row: row['Value'] if any(gun in row['Item'] for gun in self.guns_list) else 0, axis=1)
            df_items['Mags'] = df_items.apply(lambda row: row['Value'] if row['Item'] == self.mags_item else 0, axis=1)

            # Create a DataFrame for user ID and name
            df_user = pd.DataFrame([[person.split()[0], ' '.join(person.split()[1:])]], columns=['User ID', 'Name'])

            # Create a summary DataFrame with counts of guns, mags, and overall balance
            summary_data = {
                'Guns': [df_items['Guns'].sum()],
                'Mags': [df_items['Mags'].sum()]
            }
            df_summary = pd.DataFrame(summary_data)

            # Write the DataFrames to the Excel sheet
            df_user.to_excel(writer, sheet_name=person, index=False)
            df_items[['Item', 'Value']].to_excel(writer, sheet_name=person, index=False,
                                                 startrow=2)  # Keep only 'Item' and 'Value' columns
            df_summary.to_excel(writer, sheet_name=person, index=False,
                                startrow=df_items.shape[0] + 4)  # Start writing summary from the row after items

        writer.close()
        print('EXCEL file saved to: {}'.format(self.path_excel))


LLM = LockerLogManager()
input("Finished. Press Any Key To Continue...")
