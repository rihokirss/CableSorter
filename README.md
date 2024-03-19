
# Kaablite sorteerimise tööriist

See tööriist võimaldab kasutajatel sorteerida kaableid vastavalt nende vahelõikude kulgemisele algus ja lõpppunktide kaudu Exceli failides.

## Tööpõhimõte

1. **Suuna Kontrollimine ja Korrigeerimine**
Selleks, et tagada kaablite õige suund, luuakse sagedusloendur direction_counter. See loendur loeb, mitu korda iga punkt ("From" ja "To") esineb kogu andmestikus.
Kui "To" punkti sagedus on suurem kui "From" punkti sagedus, tähendab see, et suund võib olla vale ja see korrigeeritakse vahetades "From" ja "To" väärtused.
2. **Graafi Loome ja Kaablite Marsruutide Leidmine**
Andmetest luuakse graaf, kus iga kaabli alguspunkt ("From") on seotud selle lõpp-punktiga ("To"). See graaf esindab kaablite marsruute.
Kasutatakse süvitsi otsingut (DFS), et leida kõik võimalikud marsruudid alates juurtest (alguspunktidest, mis ei esine ühegi kaabli lõpp-punktina) kuni lehtedeni (lõpp-punktid, millel ei ole järgnevat kaablit).
Iga leitud marsruut lisatakse sorted_cables listi.
3. **Sorteeritud Kaablite Koostamine**
Pärast kõigi marsruutide leidmist luuakse uus DataFrame sorted_cables_df, kuhu koondatakse kõik marsruutidel leitud kaablid järjestatult.
Kaablite järjestamiseks kasutatakse leitud marsruutide järjekorda, tagades, et iga kaabli järel on kaabel, mis algab eelmise kaabli lõpp-punktist.

## Funktsionaalsus

- **Töölehe Valik**: Kasutajad saavad valida töölehe dialoogiaknast, kasutades hiirt.
- **Kaablite Sorteerimine**: Programm sorteerib kaableid vastavalt nende algus- ja lõpp-punktidele, tagades õige suuna ja järjestuse.
- **Salvestamine**: Sorteeritud andmed salvestatakse uude Exceli faili, säilitades algse vormingu ja veergude laiused.

## Nõuded

Programmi kasutamiseks on vaja järgmisi Pythoni mooduleid:
- pandas
- openpyxl
- Tkinter (tavaliselt on Pythoniga kaasas)

## Installeerimine

Veenduge, et teil on Python 3.6 või uuem versioon. Installeerige vajalikud moodulid käsurealt või terminalist järgmiste käskudega:

```
pip install pandas
pip install openpyxl
```

Tkinter peaks olema Pythoniga vaikimisi kaasas. Kui see puudub, võite vajada Pythoni installi uuesti konfigureerimist või Tkinteri eraldi installimist, mis sõltub teie operatsioonisüsteemist.

## Kasutamine

1. Käivitage programm.
2. Valige sisendfaili jaoks Exceli fail.
3. Valige tööleht dialoogiaknast.
4. Määrake väljundfaili nimi ja asukoht.
5. Programm sorteerib kaableid ja salvestab tulemused määratud väljundfaili.


