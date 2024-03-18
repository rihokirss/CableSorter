import pandas as pd
from collections import defaultdict
from tkinter import Tk, filedialog, Toplevel, Listbox, END
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# Funktsioon kaablite suuna kontrollimiseks ja korrigeerimiseks
def correct_direction(row, direction_counter):
    # Kontrollib, kas "To" sagedus on suurem kui "From" sagedus
    return direction_counter[row['To']] > direction_counter[row['From']]

# Peamine funktsioon kaablite sorteerimiseks
def sort_cables(input_file_path, sheet_name, output_file_path):
    # Laadib andmed DataFrame'i, säilitades algse vormingu
    df = pd.read_excel(input_file_path, sheet_name=sheet_name, engine='openpyxl')
    
    # Loob sagedusloenduri suuna tuvastamiseks
    direction_counter = defaultdict(int)
    for _, row in df.iterrows():
        direction_counter[row['From']] += 1
        direction_counter[row['To']] += 1

    # Korrigeerib kaablite suunda, kui vajalik
    for index, row in df.iterrows():
        if correct_direction(row, direction_counter):
            df.at[index, 'From'], df.at[index, 'To'] = row['To'], row['From']

    # Loob graafi ja servade hulga kaablite marsruutide leidmiseks
    graph = defaultdict(list)
    edges = set()
    for _, row in df.iterrows():
        graph[row['From']].append(row['To'])
        edges.add((row['From'], row['To']))

    # Leiab juured (alguspunktid)
    roots = set(graph.keys()) - {to for _, to in edges}

    sorted_cables = []

    # Süvitsi otsingu (DFS) algoritm graafi läbimiseks
    def dfs(node, path=[]):
        path.append(node)
        for next_node in graph[node]:
            if next_node not in path:
                dfs(next_node, path)
        if not graph[node]:
            sorted_cables.append(path.copy())
            path.pop()

    for root in roots:
        dfs(root, [])

    # Loob sorteeritud kaablite DataFrame'i
    from_to_mapping = {(row['From'], row['To']): row for _, row in df.iterrows()}
    sorted_cables_df = pd.DataFrame()
    for path in sorted_cables:
        for i in range(len(path) - 1):
            row = from_to_mapping.get((path[i], path[i+1]))
            if row is not None:
                sorted_cables_df = pd.concat([sorted_cables_df, pd.DataFrame([row])], ignore_index=True)

    # Salvestab sorteeritud andmed, säilitades veergude laiused
    book = load_workbook(input_file_path)
    sheet = book[sheet_name]
    column_widths = [max(len(str(cell.value)) for cell in col) for col in sheet.columns]

    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        sorted_cables_df.to_excel(writer, index=False)
        for i, width in enumerate(column_widths, start=1):
            writer.sheets['Sheet1'].column_dimensions[get_column_letter(i)].width = width

# Funktsioon töölehe valimiseks dialoogiaknast
def choose_sheet(root, sheet_names):
    def on_select(evt):
        # Valib töölehe, kui kasutaja teeb valiku
        w = evt.widget
        index = int(w.curselection()[0])
        sheet_name = w.get(index)
        root.sheet_name = sheet_name
        top.destroy()

    top = Toplevel(root)
    listbox = Listbox(top)
    listbox.pack(fill="both", expand=True)

    for name in sheet_names:
        listbox.insert(END, name)

    listbox.bind('<<ListboxSelect>>', on_select)
    root.wait_window(top)
    return root.sheet_name

root = Tk()
root.withdraw()  # Peidab Tkinteri põhiakna

# Kasutaja valib sisend- ja väljundfaili
input_file_path = filedialog.askopenfilename(title="Vali sisendfail", filetypes=[("Excel files", "*.xlsx *.xls")])
if input_file_path:
    book = load_workbook(input_file_path, read_only=True)
    sheet_names = book.sheetnames
    root.sheet_name = None  # Lisab atribuudi dünaamiliselt
    sheet_name = choose_sheet(root, sheet_names)

    output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

    # Käivitab sorteerimise, kui on valitud sisendfail, tööleht ja väljundfail
    if sheet_name and output_file_path:
        sort_cables(input_file_path, sheet_name, output_file_path)
