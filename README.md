# Game_tracks
import itertools
from openpyxl import Workbook, load_workbook
import os
from openpyxl.styles import Font, Color

def load_game_names(file):
    wb = load_workbook(file, read_only=True, data_only=True)
    ws = wb.active

    home_list = []
    away_list = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        home_name = str(row[3])
        away_name = str(row[4])

        if home_name:
            home_list.append(home_name)
        if away_name:
            away_list.append(away_name)

    print('home_list:', home_list)
    print('Away_list:', away_list)
    return home_list, away_list

def search_names_in_excel(file, home_list, away_list):
    wb = load_workbook(file, data_only=True)
    ws = wb.active
    print(file)

    home_results = {home_name: {'mit1_home': [], 'mit2_home': [], 'fin1_home': [], 'fin2_home': []} for home_name in home_list}
    away_results = {away_name: {'mit1_away': [], 'mit2_away': [], 'fin1_away': [], 'fin2_away': []} for away_name in away_list}
    
    start_row_c= 5
    end_row = ws.max_row
    for sheet in wb.sheetnames:
            ws = wb[sheet]
            
            for row in ws.iter_rows(min_row=2):
                home_name = str(row[3].value)
                away_name = str(row[4].value)
                cell_value_e = row[5].value
                cell_value_f = row[6].value
                cell_value_g = row[7].value
                cell_value_h = row[8].value

                if home_name in home_list:
                    if cell_value_e:
                    	home_results[home_name].update({'mit1_home' : cell_value_e})
                    if cell_value_f:
                        home_results[home_name].update({'mit2_home':cell_value_f})
                    if cell_value_g:
                        home_results[home_name].update({'fin1_home':cell_value_g})
                    if cell_value_h:
                        home_results[home_name].update({'fin2_home' : cell_value_h})
                        print('Home accedé')

                if away_name in away_list:
                    if cell_value_e:
                        away_results[away_name].update({'mit1_away' : cell_value_e})
                    if cell_value_f:
                        away_results[away_name].update({'mit2_away' : cell_value_f})
                    if cell_value_g:
                        away_results[away_name].update({'fin1_away' : cell_value_g})
                    if cell_value_h:
                        away_results[away_name].update({'fin2_away': cell_value_h})

    print('Home_results:', home_results)
    return home_results, away_results

def calculate_probabilities(home_results, away_results):
    print('Début calculate probabilities')
    mit1_home = home_results['mit1_home']
    mit2_home = home_results['mit2_home']
    mit1_away = away_results['mit1_away']
    mit2_away = away_results['mit2_away']
    fin1_home = home_results['fin1_home']
    fin2_home = home_results['fin2_home']
    fin1_away = away_results['fin1_away']
    fin2_away = away_results['fin2_away']

    greaterThan0mit1_home = sum(1 for element in mit1_home if element > 0)
    greaterThan0mit1_away = sum(1 for element in mit1_away if element > 0)
    greaterThan0mit2_home = sum(1 for element in mit2_home if element > 0)
    greaterThan0mit2_away = sum(1 for element in mit2_away if element > 0)
    greaterThan0fin1_home = sum(1 for element in fin1_home if element > 0)
    greaterThan0fin2_home = sum(1 for element in fin2_home if element > 0)
    greaterThan0fin1_away = sum(1 for element in fin1_away if element > 0)
    greaterThan0fin2_away = sum(1 for element in fin2_away if element > 0)
    zero_mit1_home = mit1_home.count(0)
    zero_mit2_away= mit2_away.count(0)

    prh1_home = greaterThan0mit1_home / len(mit1_home)
    prh2_home = greaterThan0mit2_home / len(mit2_home)
    prh1_away = greaterThan0mit1_away / len(mit1_away)
    prh2_away = greaterThan0mit2_away / len(mit2_away)
    prhx_home = zero_mit1_home / len(mit1_home)
    prhx_away = zero_mit2_away / len(mit2_away)
    prs1_home = greaterThan0fin1_home / len(fin1_home)
    prs2_home = greaterThan0fin2_home / len(fin2_home)
    prs1_away = greaterThan0fin1_away / len(fin1_away)
    prs2_away = greaterThan0fin2_away / len(fin2_away)
    prsx_home = zero_mit1_home / len(fin1_home)
    prsx_away = zero_mit2_away / len(fin2_away)
    pr_h1 = [(prh1_home + prh1_away) / 2] * 100
    pr_h2 = [(prh2_home + prh2_away) / 2] * 100
    pr_hx = [(prhx_home + prhx_away) / 2] * 100
    pr_s1 = [(prs1_home + prs1_away) / 2] * 100
    pr_s2 = [(prs2_home + prs2_away) / 2] * 100
    pr_sx = [(prsx_home + prsx_away) / 2] * 100
    print('Probabilities done')
    return {
        'h1_home': pr_h1,
        'h2_home': pr_h2,
        'hx_home': pr_hx,
        'full1_home': pr_s1,
        'full2_home': pr_s2,
        'fullx_home': pr_sx,
    }

def create_propositions_file(home_list, away_list, home_results, away_results, output_file):
    wb = Workbook()
    ws = wb.active

    ws.cell(row=1, column=1, value="Home")
    ws.cell(row=1, column=2, value="Away")
    ws.cell(row=1, column=3, value="H1")
    ws.cell(row=1, column=4, value="H2")
    ws.cell(row=1, column=5, value="HX")
    ws.cell(row=1, column=6, value="Full1")
    ws.cell(row=1, column=7, value="Full2")
    ws.cell(row=1, column=8, value="FullX")
    ws.cell(row=1, column=9, value="FO")
    ws.cell(row=1, column=10, value="M1")
    ws.cell(row=1, column=11, value="M2")
    ws.cell(row=1, column=12, value="MX")

    green_font = Font(color="00FF00")

    for row, home_name in enumerate(home_list):
        ws.cell(row=row + 2, column=1, value=home_name)
    for row, away_name in enumerate(away_list):
        if home_name in home_results and away_name in away_results:
            home_data = home_results[home_name]
            away_data = away_results[away_name]
            
            probabilities = calculate_probabilities(home_data, away_data)
            ws.cell(row=row + 2, column=2, value=away_name)
            ws.cell(row=row + 2, column=3, value=probabilities['h1_home'])
            ws.cell(row=row + 2, column=4, value=probabilities['h2_home'])
            ws.cell(row=row + 2, column=5, value=probabilities['hx_home'])
            ws.cell(row=row + 2, column=6, value=probabilities['full1_home'])
            ws.cell(row=row + 2, column=7, value=probabilities['full2_home'])
            ws.cell(row=row + 2, column=8, value=probabilities['fullx_home'])

    wb.save(output_file)
    print("Le fichier de propositions '{}' a été généré.".format(output_file))

def main():
    today_on_file = '/storage/emulated/0/GOMA (56).xlsx'
    results_directory = '/storage/emulated/0/RÉSULTATS'
    output_file = 'proposition50.xlsx'

    home_list, away_list = load_game_names(today_on_file)
    
    result_files = [f for f in os.listdir(results_directory) if f.endswith(".xlsx")]

    home_results = {}
    away_results = {}
   
    for file in result_files:
        file_path = os.path.join(results_directory, file)
        file_home_results, file_away_results = search_names_in_excel(file_path, home_list, away_list)
        home_results.update(file_home_results)
        away_results.update(file_away_results)
        
        print('Result_files:', result_files)

    create_propositions_file(home_list, away_list, home_results, away_results, output_file)

if __name__ == "__main__":
    main()
