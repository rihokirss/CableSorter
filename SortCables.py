import pandas as pd
from collections import defaultdict
from tkinter import Tk, filedialog
from tkinter.simpledialog import askstring
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def correct_direction(row, direction_counter):
    if direction_counter[row['To']] > direction_counter[row['From']]:
        return True
    return False

def sort_cables(input_file_path, sheet_name, output_file_path):
    # Lae andmed kasutades 'openpyxl', et säilitada formaat
    df = pd.read_excel(input_file_path, sheet_name=sheet_name, engine='openpyxl')
    
    # Suuna tuvastamiseks kasutame sagedusloendurit
    direction_counter = defaultdict(int)
    for _, row in df.iterrows():
        direction_counter[row['From']] += 1
        direction_counter[row['To']] += 1
    
    # Kontrollime ja korrigeerime vajadusel suunda
    for index, row in df.iterrows():
        if correct_direction(row, direction_counter):
            df.at[index, 'From'], df.at[index, 'To'] = row['To'], row['From']
    
    graph = defaultdict(list)
    edges = set()
    
    for _, row in df.iterrows():
        graph[row['From']].append(row['To'])
        edges.add((row['From'], row['To']))
    
    roots = set(graph.keys()) - {to for _, to in edges}
    
    sorted_cables = []
    def dfs(node, path=[]):
        path.append(node)
        for next_node in graph[node]:
            if next_node not in path:
                dfs(next_node, path)
        if node not in graph or not graph[node]:
            sorted_cables.append(path.copy())
            path.pop()
    
    for root in roots:
        dfs(root, [])
    
    from_to_mapping = {(row['From'], row['To']): row for _, row in df.iterrows()}
    sorted_cables_df = pd.DataFrame()
    
    for path in sorted_cables:
        for i in range(len(path) - 1):
            row = from_to_mapping.get((path[i], path[i+1]))
            if row is not None:
                sorted_cables_df = pd.concat([sorted_cables_df, pd.DataFrame([row])], ignore_index=True)
    
    # Salvestame tulpade laiused
    book = load_workbook(input_file_path)
    sheet = book[sheet_name]
    column_widths = []
    for col in sheet.columns:
        max_length = 0
        for cell in col:
            try: # vajalik, et kontrollida tühje rakke
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        column_widths.append(max_length)
    
    # Salvestame sorteeritud andmed, kasutades 'openpyxl' mootorit
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        sorted_cables_df.to_excel(writer, index=False)
        for i, width in enumerate(column_widths, start=1):
            writer.sheets['Sheet1'].column_dimensions[get_column_letter(i)].width = width

# Tkinter dialoogiakende jaoks
root = Tk()
root.withdraw() # Peidame Tkinteri põhiakna

# Valime sisendfaili
input_file_path = filedialog.askopenfilename(title="Vali sisendfail", filetypes=[("Excel files", "*.xlsx *.xls")])
book = load_workbook(input_file_path, read_only=True)
sheet_names = book.sheetnames
sheet_name = askstring("Töölehe valik", "Sisesta töölehe nimi:\n" + "\n".join(sheet_names))

# Valime väljundfaili asukoha
output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

# Kutsu funktsiooni välja
if input_file_path and sheet_name and output_file_path:
    sort_cables(input_file_path, sheet_name, output_file_path)
