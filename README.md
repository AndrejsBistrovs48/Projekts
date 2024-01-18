# **Projekta apraksts**

Projektu izstrādāja: Andrejs Bistrovs ``` 231RDB020 ``` un Dāvids Bižāns ```231RDB005```

## **Projekta Uzdevums**
Šis projekts ir izstrādāts, lai automatizētu vērtējumu saglabāšanu no ORTUS portāla, izmantojot Python skriptu. Projekta mērķis ir samazināt laiku un piepūli, kas nepieciešams, lai manuāli pārbaudītu un ierakstītu vērtējumus, automātiski saglabājot tos Excel formātā. Skripts izmanto Selenium tīmekļa pārlūka automatizācijas rīku, lai piekļūtu ORTUS portālam, izgūtu atzīmes un saglabātu tās Excel failā, kā arī apstrādātu PDF failus, lai iegūtu nepieciešamo informāciju.

[VIDEO](https://drive.google.com/file/d/1dcDqOc_tHu5O7fKJ3pbMunoFAebhCLHU/view?usp=sharing)

## **Izmantotās Python bibliotēkas un to izmantojuma pamatojums**
**time:** Nodrošina iespēju ieviest pauzes skripta izpildē, kas ir būtiski, lai nodrošinātu tīmekļa lapas pilnīgu ielādi pirms turpmākām darbībām.

**keyboard:** Ļauj simulēt klaviatūras nospiešanas darbības, kas nepieciešamas dažādās interaktīvās situācijās, piemēram, failu saglabāšanai vai drukāšanai.

**selenium:** Galvenā bibliotēka tīmekļa pārlūka automatizācijai, ļauj kontrolēt pārlūkprogrammu un veikt darbības tā, it kā tās veiktu cilvēks.

**chromedriver_autoinstaller:** Nodrošina ChromeDriver instalēšanu un atjaunināšanu, kas ir nepieciešama Selenium darbībai ar Chrome pārlūkprogrammu.

**PyPDF2:** Ļauj lasīt un apstrādāt PDF failus, kas ir svarīgi, lai izgūtu atzīmes no PDF dokumentiem.

**re (regulārās izteiksmes):** Nodrošina iespēju veikt teksta meklēšanu un apstrādi, izmantojot regulārās izteiksmes, kas ir nepieciešams, lai atrastu un apstrādātu specifisku informāciju tekstā.

**openpyxl:** Bibliotēka Excel failu lasīšanai un rakstīšanai, kas ļauj saglabāt izgūtos datus strukturētā veidā.

**dotenv:** Ļauj droši glabāt un izmantot konfidenciālu informāciju, piemēram, lietotāja vārdu un paroli, no .env faila.

## **Koda darbības apraksts**

1. **Izveidojot Selenium draiveri un lietotāja profila konfigurāciju**:
   - Importē nepieciešamās bibliotēkas.
   - Iestata Chrome draivera servisu un norāda lietotāja profila ceļu, lai Selenium varētu piekļūt jau esošam pārlūka profilam.

2. **Lietotāja ievadīto e-pasta skaita jautājums**:
   - Jautā lietotājam par neizlasīto e-pastu skaitu, ko nepieciešams apstrādāt.

Jautājums no konsoles:
```python
Cik neizlasītas vēstules jums ir?
```
3. **Selenium draivera inicializācija**:
   - Inicializē Chrome draiveri ar iepriekš norādītajiem uzstādījumiem.

4. **ORTUS sistēmas autentifikācija**:
   - Ielādē .env failu, kas satur lietotāja ORTUS sistēmas autentifikācijas datus.*
   - Atver ORTUS autentifikācijas lapu.
   - Aizpilda un iesniedz autentifikācijas formas laukus ar lietotāja vārdu un paroli no .env faila.

5. **E-pasta saraksta apstrāde**:
   - Pāriet uz Gmail sarakstu.
   - Meklē e-pastus ar konkrētu frāzi.
   - Katram atbilstošajam e-pastam atver to un izpilda sekojošas darbības:
     - Noklikšķina uz saites, kas ved uz ORTUS uzdevuma iesniegumu.
     - Saglabā atvērto lapu kā PDF failu izmantojot klaviatūras simulācijas komandas.
     - Atver un lasa PDF failu, lai izgūtu nepieciešamo informāciju.

6. **Datu izgūšana no PDF faila**:
   - Lasa PDF failu un izgūst no tā nepieciešamo tekstu.
   - Apstrādā tekstu, lai iegūtu priekšmetu un tēmu nosaukumus, kā arī atzīmi.
   - Ja priekšmeta nosaukumā ir "Datorgra", tas tiek aizvietots, lai izvairītos no problēmām ar īpašajiem simboliem.
```python
    if("Datorgra" in subject):
        subject="Datorgrafikas un attēlu apstrādes pamati(1),23/24-R"
```
7. **Datu saglabāšana Excel failā**:
   - Atver esošo Excel failu un meklē atbilstošo vietu datu ievietošanai.
   - Ievieto izgūto informāciju failā.
   - Saglabā izmaiņas Excel failā un aizver to.

8. **Koda noslēgums**:
   - Aizver PDF failu.
   - Aizver pārlūku, izmantojot Selenium.

*Šajā procesā svarīgi ir, ka pirms e-pasta saites izmantošanas ir jābūt veiktai ielogošanai ORTUS sistēmā. Ja tiek mēģināts izmantot saiti no e-pasta bez iepriekšējas ielogošanās, lietotājs tiks novirzīts uz ORTUS galveno lapu, nevis konkrēto uzdevuma iesniegumu. Tāpēc autentifikācija notiek pirms e-pasta apstrādes.* 

*Papildus jāpiebilst, ka katram lietotājam ir jāizveido savs .env fails ar savām ORTUS sistēmas piekļuves datiem un jānorāda atbilstošs ceļš līdz Chrome lietotāja datiem savā sistēmā.
## .env
```env
ORTUS_LOGIN=Vards.Uzvards
ORTUS_PASSWORD=Parole1234
