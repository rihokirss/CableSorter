import pandas as pd
from collections import defaultdict
 
# Funktsioon kaablite andmete lugemiseks ja sorteerimiseks
def sort_cables(input_file_path, sheet_name, output_file_path):
    # Lae andmed
    df = pd.read_excel(input_file_path, sheet_name=sheet_name)
    # Loome graafi ja servade komplekti
    graph = defaultdict(list)
    edges = set()
 
    # Täida graaf andmetega
    for _, row in df.iterrows():
        graph[row['From']].append(row['To'])
        edges.add((row['From'], row['To']))
 
    # Leidmaks alguspunktid
    roots = set(graph.keys()) - {to for _, to in edges}
 
    # Sorteeri kaablid
    def dfs(node, path=[]):
        path.append(node)
        for next_node in graph[node]:
            if next_node not in path:
                dfs(next_node, path)
        if node not in graph or not graph[node]:
            sorted_cables.append(path.copy())
            path.pop()
 
    sorted_cables = []
    for root in roots:
        dfs(root, [])
 
    # Koostame sorteeritud kaablite DataFrame'i
    from_to_mapping = {(row['From'], row['To']): row for _, row in df.iterrows()}
    sorted_cables_df = pd.DataFrame()
 
    for path in sorted_cables:
        for i in range(len(path) - 1):
            row = from_to_mapping.get((path[i], path[i+1]))
            if row is not None:
                sorted_cables_df = pd.concat([sorted_cables_df, pd.DataFrame([row])], ignore_index=True)
 
    # Salvestame sorteeritud andmed
    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        sorted_cables_df.to_excel(writer, index=False)
 
# Määra failiteed
input_file_path = 'C:\\Users\\riho\\Documents\\Coding\\CableSorter\\Vihik1.xlsx'
output_file_path = 'C:\\Users\\riho\\Documents\\Coding\\CableSorter\\sorted_cables_correctly.xlsx'
 
# Kutsu funktsiooni välja
sort_cables(input_file_path, 'Leht1', output_file_path)
