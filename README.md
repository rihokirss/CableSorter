
# Kaablite Sorteerimise Tööriist

See tööriist võimaldab kasutajatel sorteerida kaableid Exceli failides, pakkudes lihtsat ja interaktiivset kasutajaliidest töölehtede valimiseks ja andmete töötlemiseks.

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

---

Loodud kasutajaliidese ja automatiseerimise hõlbustamiseks.
