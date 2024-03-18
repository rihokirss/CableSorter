import pandas as pd

# Loeme Exceli faili
df = pd.read_excel('Vihik1.xlsx')

# Funktsioon kaablite sorteerimiseks vastavalt vaheühenduste järjekorrale
def sort_cables(df):
    sorted_cables = []
    processed_ids = set()

    # Leiame vaheühendused
    junctions = set(df['To']).intersection(set(df['From']))

    # Alguspunktid, millest alustatakse sorteerimist
    start_points = set(df['From']).difference(junctions)

    # Alguspunktidest alustades jälgime vaheühenduste järjekorda
    for start_point in start_points:
        cable_sequence = []
        current_point = start_point
        while True:
            # Valime järgmise kaabli vastavalt praegusele punktile
            next_cable = df[df['From'] == current_point].iloc[0]
            cable_sequence.append(next_cable)
            # Kui jõuame vaheühenduseni, liigume järgmise alguspunkti juurde
            if next_cable['To'] in junctions:
                current_point = next_cable['To']
            else:
                break
        # Lisame järjestatud kaablid üldisesse listi
        sorted_cables.extend(cable_sequence)
        # Märkme unikaalsete kaablite töötlemiseks
        processed_ids.update(cable_sequence['ID'])

    # Lisame kaablid, mida ei olnud vaja sorteerida
    sorted_cables.extend(df[~df['ID'].isin(processed_ids)])

    # Loome DataFrame sorteeritud kaablitega
    df_sorted = pd.concat(sorted_cables, axis=1).T

    return df_sorted

# Sorteerime kaablid
df_sorted = sort_cables(df)

# Salvestame tulemuse uude Exceli faili
df_sorted.to_excel('Vihik1_sorted.xlsx', index=False)
