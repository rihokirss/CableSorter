import pandas as pd
from collections import defaultdict

def correct_direction(row, direction_counter):
    # Tagastab True, kui rea suund tuleks ümber pöörata
    if direction_counter[row['To']] > direction_counter[row['From']]:
        return True
    return False

def sort_cables(input_file_path, sheet_name, output_file_path):
    df = pd.read_excel(input_file_path, sheet_name=sheet_name)
    
    # Suuna tuvastamiseks kasutame sagedusloendurit
    direction_counter = defaultdict(int)
    for _, row in df.iterrows():
        direction_counter[row['From']] += 1
        direction_counter[row['To']] += 1
    
    # Kontrollime ja korrigeerime vajadusel suunda
    for index, row in df.iterrows():
        if correct_direction(row, direction_counter):
            df.at[index, 'From'], df.at[index, 'To'] = row['To'], row['From']  # Vahetame suunda
    
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
    
    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        sorted_cables_df.to_excel(writer, index=False)
 
# Määra failiteed
input_file_path = 'C:\\Users\\riho\\Documents\\Coding\\CableSorter\\Vihik1.xlsx'
output_file_path = 'C:\\Users\\riho\\Documents\\Coding\\CableSorter\\sorted_cables_correctly.xlsx'
 
# Kutsu funktsiooni välja
sort_cables(input_file_path, 'Leht1', output_file_path)
