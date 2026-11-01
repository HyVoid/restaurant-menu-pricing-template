[ 🌐 عربي ](README.ar.md) | [ 🇩🇪 Deutsch ](README.de.md) | [ 🇫🇷 Français ](README.fr.md) | [ 🇳🇱 Nederlands ](README.nl.md) | [ 🇪🇸 Español ](README.sp.md) | [ 🇬🇧 English ](README.md)

# Restaurantmenu-prijscalculator, modifierconfiguratie & POS-voorbereidings-toolkit (Excel & websjabloon)
<img width="1920" height="1032" alt="image" src="https://github.com/user-attachments/assets/45a7c5cb-ec7e-4744-a38b-e705561fe0d6" />


[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-Web%20App%20%2B%20Excel%20Template-2251FF)
![Tool Type](https://img.shields.io/badge/Tool-Restaurant%20Menu%20Management%20%26%20POS%20Pricing-051C2C)
![Architecture](https://img.shields.io/badge/Architecture-Formula--Driven%20No--VBA-success)

**Een gratis online prijscalculator voor restaurantmenu's en een offline Excel-sjabloon voor POS-configuratie, ontworpen voor zelfstandige restaurants, pizzeria's en QSR-exploitanten. Centraliseer uw menu-stamgegevens, bereken complexe modifiertoeslagen, configureer regels voor beleg op halve en hele pizza's, voorkom dubbele berekening op inbegrepen receptingrediënten en exporteer gevalideerde, POS-gereede menubladen zonder VBA-macro's.**

> **Test uw menuprijsregels direct in uw browser** — geen accountregistratie, software-installatie of geavanceerde spreadsheetformules nodig.
>
> - 🌐 **[Start de gratis interactieve web-app voor restaurantmenuprijzen](https://hyvoid.github.io/restaurant-menu-pricing-template/)** *(Test menuconfiguraties voor één sessie, modifierlogica en klantbestelprijzen online)*
> - 📥 **[Download het offline Excel-sjabloon voor restaurant-POS-menuconfiguratie (.xlsx)](https://theseusworkshop.com/l/dznbyd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=restaurant-menu-pricing-template)** *(Gebouwd voor permanente menu-stamgegevens, prijswijzigingen over meerdere seizoenen en operationele POS-audittrails — ondersteund door een 30-daagse geld-terug-garantie zonder vragen)*

---

## Waar dit sjabloon voor restaurantmenubeheer u bij helpt (van pijnpunt naar oplossing)

| Kernstatistiek voor menugegevens & prijzen | Veelvoorkomend operationeel pijnpunt in restaurants | Hoe deze menuconfiguratie-toolkit dat oplost |
|---|---|---|
| **Menuhiërarchie & beschikbaarheid per maat** | Menucategorieën, itemmaten (Small/Medium/Large) en seizoensgebonden beschikbaarheid zijn verspreid over verouderde pdf's en POS-exports. | Volgt actieve categorieën, subcategorieën, item-SKU's en maatgebonden beschikbaarheid in één relationele menu-stamgegevensdatabase. |
| **Basisprijs & gelaagde modifiertoeslagen** | Personeel gokt extra-toeslagen voor beleg of rekent vaste bedragen die maatverschillen negeren, wat de voedselmarges uitholt. | Berekent exacte basisprijzen plus standaard- en premium-modifiertoeslagen, geschaald per productmaat. |
| **Standaardreceptingrediënten (geen dubbele berekening)** | Klanten die speciale items bestellen (bijv. Hawaiiaanse of Supreme Pizza) betalen dubbel als ze het standaard inbegrepen beleg behouden. | Wijst standaard inbegrepen ingrediënten per item toe, zodat standaardhoeveelheden $0 extra kosten terwijl "Extra"-porties automatisch een toeslag activeren. |
| **Fractionele & halve-pizza belegprijzen** | Handmatige berekening van beleg op de linkerhelft, rechterhelft en hele pizza leidt tot vertraging bij de kassa en inconsistente klantrekeningen. | Past globale fractionele prijsvermenigvuldigers toe om nauwkeurige toeslagen voor halve en hele porties automatisch te berekenen. |
| **Min-/max-selectielimieten per modifiergroep** | Klanten kiezen ongeldige combinatieopties of slaan verplichte keuzes over (zoals korsttype of gaarheid) door foutieve POS-regels. | Handhaaft strikte minimum-, maximum- en per-modifier-hoeveelheidslimieten op modifiergroepniveau vóór live ingebruikname. |
| **Klaarheid voor POS-synchronisatie & QA-status** | Het uploaden van onvolledige itemrecords of ongeverifieerde prijzen rechtstreeks naar een live POS breekt keukenbonnen en online bestellingen. | Markeert ontbrekende prijzen en gebroken modifierkoppelingen op een validatiedashboard en isoleert schone, POS-gereede exportrijen. |

---

## Stapsgewijze handleiding: hoe u restaurantmenuprijzen configureert en POS-imports voorbereidt

Volg deze workflow van 4 stappen voor menu-engineering om uw itemdatabase te standaardiseren, complexe ordercombinaties te testen en een schoon POS-importblad te genereren:

1. **Stap 1: Configureer globale belasting-, valuta- en fractionele prijsparameters voor uw restaurant**
   **Actie:** Open het tabblad `Parameters` en voer uw lokale omzetbelastingtarief, valutasymbool, prijsvermenigvuldiger voor halve porties beleg en toeslagfactor voor premium-modifiers in.
   *Waarom dit belangrijk is:* Door globale prijsvariabelen eenmalig in te stellen, legt u één uniforme rekenbasis vast voor elke menucategorie en elke itemmaat.

2. **Stap 2: Importeer of voer uw menu-stamlijst, maten & modifiergroepen in**
   **Actie:** Vul de gestructureerde invoertabellen met uw menucategorieën, gerechtnamen, maatformen, modifiergroepen en ruwe ingrediëntenprijzen — of plak bestaande rijen uit uw huidige POS-CSV-export.
   *Waarom dit belangrijk is:* Door items en modifiers te koppelen via schone relationele id's voorkomt u dubbele gegevensinvoer wanneer ingrediëntenprijzen veranderen.

3. **Stap 3: Wijs standaardreceptonderdelen toe & simuleer echte klantbestellingen**
   **Actie:** Wijs modifiergroepen toe aan specifieke menu-items, markeer ingebouwde speciale ingrediënten en voer testbestellingen uit in de interactieve **Orderprijssimulator** (test maatswijzigingen, beleg op de linker-/rechterhelft en extra porties).
   *Waarom dit belangrijk is:* Realtime simulatie van bestelprijzen vóór en na belasting legt onderprijsde add-ons en bugs door dubbele berekening bloot voordat ook maar één klant bestelt.

4. **Stap 4: Controleer de configuratiegezondheid & verzeker langetermijnmenuonderhoud met Excel**
   **Actie:** Controleer het automatische QA-dashboard om onvolledige itemmarkeringen op te lossen, exporteer uw gevalideerde POS-voorbereidingstabel en **ga over van de browserdemo voor één sessie naar de herbruikbare Excel-werkmap**.
   *Waarom dit belangrijk is:* Hoewel de [gratis menucalculator in de browser](https://hyvoid.github.io/restaurant-menu-pricing-template/) ideaal is voor snelle eenmalige tests, bewaren browsersessies historische wijzigingen niet permanent. **[Download het volledige offline Excel-sjabloon voor menu-stamgegevens (.xlsx)](https://theseusworkshop.com/l/dznbyd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=restaurant-menu-pricing-template)** om de permanente menudatabase van uw restaurant op te slaan, terugkerende kwartaalprijsupdates uit te voeren, historische auditlogboeken bij te houden en uw configuratie te hergebruiken op meerdere vestigingen zonder opnieuw te beginnen.

**Configureer globale prijsregels eenmalig. Onderhoud één schoon menu-stamgegevensbestand. Simuleer echte klantafrekenprijzen. Exporteer pas naar uw POS na geautomatiseerde validatie.**

---

## Waarom standaard restaurantspreadsheets falen bij modifier- en belegprijzen

Menubeheer voor restaurants wordt ogenschijnlijk complex zodra een gerecht niet langer door één vaste prijs wordt gedefinieerd.

Denk aan een speciale pizza of aanpasbaar combinatiemenu met prijzen per maat, inbegrepen standaardbeleg, optionele premium add-ons en door de klant kiesbare hoeveelheden. Een conventionele menulijst in Excel of Google Sheets kan basisartikelprijzen vastleggen, maar slaagt er niet in **ingrediënten die al in het basisrecept zijn inbegrepen** te scheiden van **extra modifiers waarvoor een toeslag geldt**.

Deze structurele blinde vlek creëert twee kostbare commerciële risico's voor foodbedrijven:

1. **De valkuil van dubbele berekening bij klanten:** Klanten worden per ongeluk extra berekend voor standaardingrediënten die al in de basisprijs van een speciaal gerecht zitten, wat leidt tot geschillen bij de kassa en afgebroken online bestellingen.
2. **Stille erosie van winstmarges:** Kassamedewerkers en POS-beheerders passen inconsistente prijzen toe op extra beleg, maatsupgrades of keuzes voor halve-en-halve pizza omdat prijsregels met de hand worden berekend.

Bijvoorbeeld: een **Hawaiiaanse pizza** bevat normale porties ham en ananas in de basisprijs. Het behouden van standaardhoeveelheden ham en ananas moet **$0.00** aan de rekening toevoegen. Het bestellen van *extra* ananas moet echter een maatgecorrigeerde toeslag activeren — en het plaatsen van premium beleg alleen op de **linkerhelft** moet een gedefinieerde fractionele prijsregel volgen in plaats van giswerk.

Dit zijn geen simpele typefouten. Ze ontstaan omdat vlakke spreadsheets de commerciële regels achter een modern restaurantmenu niet kunnen modelleren.

Deze toolkit vertaalt die menurelaties naar een gestructureerd, formulegestuurd configuratiesysteem. In plaats van voor elk item dubbele beleglijsten bij te houden en prijzen handmatig te controleren, definieert u ingrediënten, modifiergroepen en prijsregels één keer — en past ze consistent toe op uw hele menu.

---

## Handmatige menuspreadsheets vs. gestructureerde toolkit voor POS-menuconfiguratie

| Uitdaging bij restaurantmenuconfiguratie | Handmatige spreadsheets & directe POS-wijzigingen (zonder deze tool) | Gestructureerde menu-stamgegevens & prijstoolkit (met deze tool) |
|---|---|---|
| **Fragmentatie van menu-stamgegevens** | Categorieën, itemmaten en modifierprijzen staan in losse tabbladen of ongeorganiseerde CSV-exports, wat versiebeheerchaos veroorzaakt bij menu-updates. | Relationele hoofdtabellen scheiden categorieën, basisitems en maatspecifieke prijsmatrices met behoud van automatische id-lookups. |
| **Dubbel onderhoud van modifiers & beleg** | Dezelfde kaas, saus of hetzelfde beleg wordt tientallen keren getypt over menu-items, met niet-overeenkomende namen en verouderde toeslagprijzen. | Elke ingrediëntmodifier wordt één keer aangemaakt in een centrale bibliotheek en via herbruikbare modifiergroepen aan meerdere gerechten gekoppeld. |
| **Prijzenbugs bij speciale items & standaardrecepten** | Standaard inbegrepen beleg op speciale pizza's of burgers wordt per ongeluk als betaalde add-on gefactureerd wanneer klanten hun bestelling aanpassen. | Standaardreceptmapping onderscheidt ingebouwde standaardhoeveelheden ($0 toeslag) automatisch van betaalde "Extra"-upgrades. |
| **Inconsistente halve-pizza & fractionele prijzen** | Personeel berekent belegtoeslagen voor de linkerhelft, rechterhelft en hele pizza handmatig, wat factuurfouten en giswerk aan de kassa oplevert. | Globale fractionele prijsfactoren berekenen belegprijzen voor halve porties automatisch op basis van het referentietarief voor de hele pizza. |
| **Foutieve POS-modifierselectieregels** | Minimale verplichte keuzes en maximale beleglimieten worden tijdens de installatie vergeten, waardoor onmogelijke bestellingen het keukenweergavescherm bereiken. | Beperkingsregels op modifiergroepen dwingen expliciete min-/max-selectielimieten en hoeveelheidslimieten per item af tijdens de POS-voorbereiding. |
| **Fouten bij live POS-implementatie & synchronisatie** | Ongeverifieerde prijzen en onvolledige itemrijen gaan rechtstreeks naar live POS-terminals, wat kassacrashes tijdens piekuren veroorzaakt. | Ingebouwde QA-validatie markeert ontbrekende gegevens, simuleert klantkassabonnen en genereert een schone, vooraf gecontroleerde POS-exporttabel. |

---

## Voor wie dit sjabloon voor restaurantmenu & prijzen is (doelrollen & gebruiksscenario's)

Gebouwd voor zelfstandige restaurants, pizzeria's, fastservicerestaurants (QSR), fast-casual-concepten en foodbedrijven met meerdere vestigingen die menu's met meerdere maten, gedeelde ingrediëntmodifiers, voorwaardelijke toeslagen en strikte bestelregels beheren.

### Rolgerichte zoeksenario's & operationele gebruiksscenario's

- **Pizzeria-eigenaren & zelfstandige restaurantexploitanten (`Pizza Menu Pricing Calculator & Topping Spreadsheet`)**
  Stop met geld te verliezen op niet-gefactureerd extra beleg of klanten te veel te berekenen voor speciale pizza's. Modelleer vermenigvuldigers voor hele pizza versus halve pizza, toeslagen per maat en inbegrepen receptitems voordat u uw fysieke borden of online bestelmenu bijwerkt.
- **Operationeel managers bij QSR & fast-casual (`Multi-Size Restaurant Menu Management Template`)**
  Standaardiseer menuhiërarchieën, combinatie-modifiergroepen en seizoensprijsaanpassingen over meerdere productcategorieën in één gecentraliseerd offline hoofdbestand, zonder afhankelijk te zijn van kwetsbare VBA-scripts.
- **POS-implementatiespecialisten & menubeheerders (`Restaurant POS Menu Staging & Import Checklist`)**
  Maak ruwe menugegevens schoon, structureer en valideer ze voordat u databases bouwt of bijwerkt in Toast, Square, Clover, Lightspeed of eigen POS-platforms. Vang ontbrekende basisprijzen en foutieve modifier-min/max-regels op vóór livegang.
- **Restaurantconsultants & menu-engineers (`Menu Engineering & Modifier Architecture Toolkit`)**
  Lever horecacliënten een herhaalbare, auditgereede menuconfiguratiewerkmap. Herstructureer rommelige klantmenu's naar logische modifiergroepen en demonstreer tijdens adviestrajecten de exacte klantafrekenprijzen.

**Voor wie deze toolkit NIET is ontworpen:**
Foodbedrijven die op zoek zijn naar een live kassasysteem, creditcardverwerker, keukenweergavesysteem (KDS), tracker voor rauw voorraadverbruik of automatische calculator voor receptvoedingswaarden. Deze toolkit is een **menu-stamgegevenssysteem vóór de POS, prijssimulator en configuratie-QA-systeem** — het bereidt en valideert uw commerciële menulogica zodat uw POS-platform zonder prijsfouten draait.

---

## De zakelijke logica & methodologie: verborgen prijsblinde vlekken in restaurants oplossen

*(Geschreven voor restauranteigenaren en exploitanten — geen technische of geavanceerde spreadsheetachtergrond nodig)*

Wanneer klanten eten bestellen, ziet menuprijzen er aan de oppervlakte eenvoudig uit: kies een item, voeg beleg toe, betaal het totaal. Achter de toonbank verliezen de meeste zelfstandige restaurants echter geld of frustreren klanten door **drie verborgen commerciële blinde vlekken** die standaardspreadsheets niet opmerken.

Zo beschermt de zakelijke logica in deze toolkit uw winstmarges en klantbeleving, in begrijpelijke taal:

### 1. Het probleem "één update, tien fouten" elimineren (één bron van waarheid)
* **De blinde vlek:** Stel dat de groothandelsprijs van mozzarella of spek met 20% stijgt. In een typische restaurantspreadsheet staat "Extra spek" apart vermeld onder Burgers, Sandwiches, Salades en Pizza's. Als u de prijs op drie plaatsen bijwerkt maar de vierde vergeet, verliest u winstmarge of verwart u klanten met inconsistente prijzen.
* **Hoe de toolkit dit oplost:** U voert elk ingrediënt en elke modifier **één keer** in een centrale lijst in. Of dat ingrediënt nu op 2 items of op 50 items wordt gebruikt: het bijwerken van de prijs in één cel werkt automatisch elk gekoppeld menu-item en elke orderberekening in de hele werkmap bij.

### 2. De valkuil "dubbele berekening bij speciale items" voorkomen (voorwaardelijke receptprijzen)
* **De blinde vlek:** Wanneer een klant een "Meat Lover's Pizza" of een "Deluxe Burger" bestelt, betaalt hij al een hogere basisprijs omdat pepperoni, worst of kaas in het recept zit. Als uw menu-opstelling elk beleg hetzelfde behandelt, kan uw systeem de klant per ongeluk *nogmaals* berekenen alleen al voor het bevestigen van de standaardingrediënten — of, als u al het beleg gratis maakt om klachten te voorkomen, verliest u geld zodra iemand om *dubbel* vlees vraagt.
* **Hoe de toolkit dit oplost:** Het sjabloon scheidt **"standaard inbegrepen"** van **"betaald extra"**. Het weet dat 1x pepperoni op een pepperonipizza $0.00 extra kost, terwijl het verhogen naar 2x (extra pepperoni) automatisch de exacte upgradeprijs berekent. U beschermt klantvertrouwen zonder gratis eten weg te geven.

### 3. Giswerk van kassamedewerkers bij halve porties beleg & maten stoppen (proportionele prijslogica)
* **De blinde vlek:** Klanten vragen vaak beleg alleen op de linkerhelft of rechterhelft van een pizza, of wisselen tussen de maten Small, Medium en Large. Zonder een duidelijke formule berekent personeel ofwel de volledige prijs voor beleg op een hele pizza (waardoor de klant gefrustreerd raakt), ofwel halveert het prijzen handmatig zonder rekening te houden met de voorbereidingsarbeid (wat uw marges schaadt).
* **Hoe de toolkit dit oplost:** U stelt één regel in — bijvoorbeeld "Beleg op een halve pizza kost 60% van beleg op een hele pizza". De toolkit berekent automatisch het exacte, op de cent nauwkeurige bedrag voor plaatsing op de linkerhelft, rechterhelft en hele pizza, voor elke pizzamaat.

### 4. Onmogelijke keukenbonnen blokkeren (min-/max-selectielimieten)
* **De blinde vlek:** Als een klant online een combinatiemenu bestelt en vergeet een verplicht bijgerecht te kiezen — of 12 gratis sauzen selecteert terwijl er maar 2 zijn toegestaan — moet uw keukenpersoneel stoppen met koken, de klant bellen of de kosten op zich nemen.
* **Hoe de toolkit dit oplost:** U definieert eenvoudige grensregels per optiegroep (bijv. *"Kies minstens 1 korst; maximaal 1"* of *"Kies tot 4 standaardbelegsoorten"*). De ingebouwde ordersimulator test uw menu direct tegen deze regels en laat zien of een klant per ongeluk een ongeldige bestelling zou kunnen plaatsen.

### 5. Testen vóór livegang (kwaliteitsborging vóór de POS)
* **De blinde vlek:** De meeste restaurants ontdekken prijsfouten in hun menu pas *nadat* ze hun POS-systeem hebben bijgewerkt — meestal als een klant in de spits bij de kassa klaagt of een online bezorg-app het verkeerde totaal berekent.
* **Hoe de toolkit dit oplost:** De toolkit fungeert als een veilige "crashtestpop" voor uw menu. U kunt echte klantbestellingen simuleren, de exacte kassabon vóór en na belasting zien en een automatisch gezondheidsdashboard raadplegen dat ontbrekende prijzen rood markeert. U exporteert uw menu pas naar uw POS als elk item de inspectie heeft doorstaan.

---

## Over de maker

Ik bouw praktische, formulegestuurde zakelijke tools die rommelige operationele gegevens omzetten in gestructureerde, herhaalbare beslissingsondersteunende workflows.

De meeste zelfstandige restaurants en kleine bedrijven hebben hun prijsgegevens al, maar die zitten opgesloten in losse spreadsheets, rommelige POS-exports en ongedocumenteerde gewoontes van personeel. De echte operationele uitdaging is die records met elkaar te verbinden, consistente prijsregels af te dwingen en configuratiefouten op te sporen voordat ze betalende klanten bereiken.

De **Restaurant Menu Pricing Calculator & POS Configuration Toolkit** biedt horeca-exploitanten menuvalidatie op ondernemingsniveau binnen een transparante, lichtgewicht browser- en Excel-workflow.

- 🌐 **[Probeer nu de gratis menucalculator in de browser](https://hyvoid.github.io/restaurant-menu-pricing-template/)**
- 📥 **[Vraag de volledige herbruikbare Excel-werkmap met menu-stamgegevens aan (.xlsx)](https://theseusworkshop.com/l/dznbyd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=restaurant-menu-pricing-template)**


## Technische details

<details>
<summary>Voor technische beoordelaars, Excel-beoefenaars en ontwikkelaars — vouw uit om de werkmaparchitectuur, prijslogica, formules en validatieregels te inspecteren.</summary>

### Werkmaparchitectuur

De toolkit is ontworpen als een relationele menuconfiguratiedatabase en prijsvalidatie-engine, niet als een verzameling losse menuspreadsheets.

De architectuur volgt drie principes:

1. **Eén bron van waarheid:** Elke categorie, elk product, elke maat, elke modifier en elke prijsregel heeft een aangewezen locatie voor stamgegevens.
2. **Herbruikbare relaties:** Menu-items erven modifiergroepen via toewijzingen op categorie- of itemniveau in plaats van configuratierecords te dupliceren.
3. **Geautomatiseerde validatie en uitvoer:** Dynamische formules verbinden de stamgegevens, berekenen klantprijzen, detecteren configuratiefouten en genereren gestandaardiseerde POS-exportrecords.

De werkmap bevat vijf functionele modules en 13 werkbladen.

#### 1. Globale parameters & prijsregels

| Werkblad | Verantwoordelijkheid |
|---|---|
| `00_Param_Config` | Centraliseert valuta, belastingtarief, prijsfactoren voor hele/halve porties, premium-modifiertoeslagen, operationele drempelwaarden, statusindicatoren en voorwaardelijke prijsregels. |

Dit werkblad is de globale besturingslaag. Prijsaannames en zakelijke statusidentificatoren worden centraal onderhouden in plaats van als vaste waarden ingebed te zijn in afzonderlijke berekeningsformules.

#### 2. Menu-productstamgegevens

| Werkblad | Verantwoordelijkheid |
|---|---|
| `01_Menu_Categories` | Beheert menucategorieën, subcategorieën, weergavevolgorde en categoriebeschikbaarheid. |
| `02_Menu_Items` | Slaat identiteiten van afzonderlijke producten, beschrijvingen, categorierelaties, producttypen en beschikbaarheid op. |
| `03_Item_Sizes` | Slaat maatspecifieke productprijzen op en koppelt elke maat aan het bovenliggende menu-item. |

De scheiding tussen productidentiteit en maatspecifieke prijzen voorkomt dat de productstamgegevens horizontaal uitdijen telkens wanneer een nieuwe maat of prijsoptie wordt geïntroduceerd.

#### 3. Herbruikbare modifiers & multidimensionale prijzen

| Werkblad | Verantwoordelijkheid |
|---|---|
| `04_Modifier_Groups` | Definieert klantgerichte modifiergroepen en hun operationele reikwijdte. |
| `05_Modifiers` | Beheert wereldwijd herbruikbare ingrediënten- en modifieridentiteiten. |
| `06_Modifier_Pricing` | Slaat modifierprijzen op per productmaat en portieniveau, met berekeningen voor plaatsing op de hele, linker en rechter helft. |

Hetzelfde ingrediënt kan over meerdere producten worden hergebruikt zonder afzonderlijke kopieën van zijn stamrecord aan te maken.

#### 4. Relatiemapping & selectiebeperkingen

| Werkblad | Verantwoordelijkheid |
|---|---|
| `07_Default_Modifiers` | Wijst speciale producten toe aan hun inbegrepen ingrediënten en standaardhoeveelheden. |
| `08_Group_Assignment` | Wijst modifiergroepen toe op categorieniveau of op individueel productniveau. |
| `09_Selection_Rules` | Regelt minimale selecties, maximale selecties, verplichte groepen en machtiging voor meervoudige selectie. |

Deze module bepaalt welke modifiergroepen op elk product van toepassing zijn en of de resulterende klantselectieconfiguratie logisch geldig is.

#### 5. Ordersimulatie, POS-export & monitoring

| Werkblad | Verantwoordelijkheid |
|---|---|
| `10_Order_Test` | Simuleert klantselecties en berekent eindprijzen met controle op naleving van de selectieregels. |
| `11_POS_Export` | Genereert een gestandaardiseerde lijst van beschikbare producten, maten, prijzen, modifiersamenvattingen en synchronisatiestatus. |
| `12_Dashboard` | Vat menugezondheid, configuratieanomalieën, productbeschikbaarheid en POS-gereedheid samen. |

#### Gegevensstroom

De volledige gegevensstroom volgt deze volgorde:

```text
GLOBAL PARAMETERS
00_Param_Config
       |
       v
MENU MASTER DATA                 MODIFIER MASTER DATA
01_Menu_Categories              04_Modifier_Groups
       |                        05_Modifiers
02_Menu_Items                   06_Modifier_Pricing
       |                               |
03_Item_Sizes                           |
       |                                |
       +---------------+----------------+
                       |
                       v
             RELATIONSHIP & RULES
             07_Default_Modifiers
             08_Group_Assignment
             09_Selection_Rules
                       |
          +------------+-------------+
          |            |             |
          v            v             v
       10_Order     11_POS        12_Dashboard
         Test       Export
          |            |             |
          v            v             v
       Price QA    POS Update    Configuration
                                  Monitoring
````

### Beslissingskader

De toolkit is gebouwd rond vier operationele beslissingen.

| Beslissing                                 | Vereiste informatie                                                                           | Beslisuitkomst                                                             |
| ------------------------------------------ | ---------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------- |
| Wat kan worden verkocht?                          | Categorie-, product-, maat- en beschikbaarheidsrecords.                                             | Identificeert de geldige menustructuur en beschikbare productconfiguraties.    |
| Wat moeten klanten betalen?                 | Basisprijs, modifierprijzen, inbegrepen ingrediënten, portiehoeveelheden, plaatsing en belasting.    | Bepaalt de berekende klantprijs voor een specifieke orderconfiguratie. |
| Wat kunnen klanten selecteren?                 | Modifiertoewijzingen, groepslidmaatschap, verplichte status en minimale/maximale selectielimieten. | Identificeert geldige en ongeldige bestelconfiguraties.                        |
| Is het menu gereed voor POS-synchronisatie? | Integriteit van stamgegevens, volledigheid van prijzen, modifierrelaties en validatiestatus.    | Scheidt exportgereede records van records die correctie vereisen.        |

Het kader onderscheidt drie verschillende soorten informatie:

- Productfeiten: wat het restaurant verkoopt en welke maten beschikbaar zijn.
- Commerciële regels: hoe ingrediënten, hoeveelheden en plaatsing de verkoopprijs beïnvloeden.
- Operationele beperkingen: welke selecties zijn toegestaan en of de configuratie gereed is voor implementatie.

Deze scheiding voorkomt dat een productbeschrijving, een prijsaanname en een bestelbeperking als onderling uitwisselbare gegevens worden behandeld.

### Drie veelvoorkomende prijs- en configuratievalkuilen

#### Valkuil 1 — Twee keer berekenen voor al inbegrepen ingrediënten

De beslissing: de klantprijs van een speciale pizza vaststellen op basis van de vermelde ingrediënten.

De onopgemerkte fout: een conventionele modifierberekening telt de standaardtoeslag voor beleg op, ongeacht of dat beleg al in het speciale product is inbegrepen.

Denk aan een Hawaiiaanse pizza met ham en ananas inbegrepen in de basisprijs.

| Klantselectie | Foute behandeling                                         | Correcte behandeling                                  |
| ------------------ | ----------------------------------------------------------- | -------------------------------------------------- |
| Normale ham         | Telt een belegtoeslag op                                       | Geen extra toeslag                               |
| Lichte ananas    | Telt een belegtoeslag op                                       | Geen extra toeslag                               |
| Extra ham          | Telt een belegtoeslag op zonder rekening te houden met inclusie         | Past de geconfigureerde prijsregel voor extra hoeveelheid toe |
| Extra ananas    | Gebruikt de standaardmodifierprijs zonder de inclusie te controleren | Past de geconfigureerde prijsregel voor extra hoeveelheid toe |

De foutieve aanpak behandelt elk geselecteerd ingrediënt als een nieuwe aankoop. De gecorrigeerde aanpak erkent dat een ingrediënt al deel kan uitmaken van de commerciële definitie van het product.

De `07_Default_Modifiers`-mapping identificeert de inbegrepen ingrediënten. De voorwaardelijke prijsregels bepalen vervolgens de toepasselijke vermenigvuldiger op basis van inclusiestatus en geselecteerd portietype.

#### Valkuil 2 — Beleg op een halve pizza behandelen als onafhankelijke vaste prijzen

De beslissing: dezelfde modifierprijs berekenen, ongeacht of beleg de hele pizza bedekt of slechts één helft.

De onopgemerkte fout: de hoeveelheid product die het beleg krijgt verandert, maar de prijsberekening weerspiegelt de gekozen plaatsing niet.

Bijvoorbeeld, stel dat beleg op een hele pizza $2.00 kost.

| Plaatsing | Foute prijs | Correcte prijs |
| --------- | --------------- | ------------- |
| Hele      | $2.00           | $2.00         |
| Links      | $2.00           | $1.00         |
| Rechts     | $2.00           | $1.00         |

De foutieve aanpak rekent klanten te veel aan voor halve-pizzakeuzes wanneer het beleid van het restaurant proportionele prijzen is.

De gecorrigeerde aanpak leidt de prijzen voor de linker- en rechterhelft af van de referentieprijs voor de hele pizza met behulp van gecentraliseerde prijsfactoren.

#### Valkuil 3 — POS-menugereedheid behandelen als een gegevensinvoertaak

De beslissing: de nieuwste menuspreadsheet naar de POS-leverancier sturen na het bijwerken van de productprijzen.

De onopgemerkte fout: een voltooide prijsinvoer bewijst niet dat de volledige productconfiguratie geldig is.

Een product kan een basisprijs hebben maar alsnog het volgende bevatten:

- Een ontbrekend maatsrecord.
- Een ongeldige categorierelatie.
- Een niet-beschikbare modifiergroep.
- Ontbrekende toewijzingen van standaardingrediënten.
- Inconsistente minimale en maximale selectieregels.

| Controleaanpak                           | Risico                                             | Gecorrigeerde aanpak                                         |
| ----------------------------------------- | ------------------------------------------------ | ---------------------------------------------------------- |
| Controleren of de productnaam bestaat.    | Onvolledige producten lijken gereed.                | Valideer productidentiteit en categorierelaties.      |
| Controleren of een basisprijs bestaat.        | Sommige beschikbare maten blijven zonder prijs.            | Valideer elke geconfigureerde product-maat-combinatie.         |
| Controleren of modifiergroepen zijn vermeld. | Ongeldige of lege groepen kunnen de POS bereiken.       | Valideer groepslidmaatschap en selectiebeperkingen.       |
| Het hele menu exporteren.                   | Fouten worden overgedragen naar de POS-configuratie. | Exporteer in aanmerking komende records en controleer de synchronisatiestatus. |

De gecorrigeerde aanpak onderscheidt voltooide gegevensinvoer van operationele gereedheid.

### Voorbeeldscenario

Scenario: een Hawaiiaanse pizza in maat Large met extra ananas op de linkerhelft testen

Een restaurant wil bevestigen dat zijn prijzen voor speciale pizza's correct blijven wanneer klanten de hoeveelheid en plaatsing van inbegrepen ingrediënten wijzigen.

Ga uit van de volgende illustratieve configuratie:

| Invoer                | Waarde          |
| -------------------- | -------------- |
| Product              | Hawaiiaanse pizza |
| Maat                 | Large          |
| Basisprijs           | $18.00         |
| Inbegrepen ingrediënten | Ham, ananas |
| Geselecteerde modifier    | Ananas      |
| Geselecteerde hoeveelheid    | Extra          |
| Plaatsing            | Links           |
| Prijs hele modifier | $2.00          |
| Factor linkerhelft     | 0.50           |
| Omzetbelasting            | 8.00%          |

Stap 1 — Identificeer het inbegrepen ingrediënt.

Ananas is gemapt als inbegrepen ingrediënt voor de Hawaiiaanse pizza. De klant wordt daarom niet opnieuw berekend voor de standaard inbegrepen hoeveelheid.

Stap 2 — Bepaal de toepasselijke hoeveelheidsvermenigvuldiger.

De klant kiest Extra in plaats van Normal. Volgens de illustratieve regelmatrix geldt voor een inbegrepen ingrediënt dat als Extra wordt geselecteerd een vermenigvuldiger van 1.00.

Stap 3 — Bereken de plaatsinggecorrigeerde modifierprijs.

```
Whole Modifier Price = $2.00

Left Half Price = $2.00 × 0.50

Raw Portion Price = $1.00

Applicable Multiplier = 1.00

Additional Charge = $1.00 × 1.00

Additional Charge = $1.00
```

Stap 4 — Bereken de klantprijs.

| Berekening                 | Bedrag |
| --------------------------- | ------ |
| Hawaiiaanse pizza Large        | $18.00 |
| Extra ananas — linkerhelft | $1.00  |
| Totaal vóór belasting               | $19.00 |
| Omzetbelasting (8%)              | $1.52  |
| Eindprijs klant        | $20.52 |

Analytische interpretatie

Een conventionele belegcalculator zou mogelijk de standaardananasprijs opnieuw berekenen, zelfs als het ingrediënt al inbegrepen is. Een calculator die de plaatsing negeert, zou daarentegen mogelijk de volledige $2.00 berekenen voor een selectie op de halve pizza.

Het correcte resultaat weerspiegelt beide commerciële omstandigheden: de status van het inbegrepen ingrediënt en de gekozen plaatsing.

Beslissingsimplicatie: het restaurant kan de klantgerichte prijs valideren voordat het zijn live menu wijzigt, waardoor het risico op inconsistente berekeningen en vermijdbare POS-correcties afneemt.

Dit scenario illustreert de geconfigureerde prijslogica. De werkelijke resultaten zijn afhankelijk van de eigen prijsmatrix, modifierprijzen en belastinginstellingen van het restaurant.


### Formuliereferentie

De werkmap gebruikt native Excel-functies voor dynamische matrices om relaties tussen menurecords te onderhouden, prijzen te berekenen en validatie-uitvoer te genereren.

Kernfuncties zijn onder andere:

| Functie | Doel |
|---|---|
| `LET` | Benoemt tussenberekeningen en houdt complexe formules leesbaar. |
| `FILTER` | Extraheert geldige records en genereert dynamisch uitvoerlijsten. |
| `XLOOKUP` | Haalt gerelateerde product-, categorie-, modifier- en prijsinformatie op. |
| `COUNTIF` / `COUNTIFS` | Telt records en controleert selectie- of configuratievoorwaarden. |
| `UNIQUE` | Genereert unieke categorie- en productidentificatoren. |
| `SUM` / `SUMIFS` | Aggregeert modifiertoeslagen en zakelijke waarden. |
| `IF` / `AND` / `OR` | Past prijsvoorwaarden en validatieregels toe. |
| `ROUND` | Beheerst de precisie van geldbedragberekeningen. |
| `AVERAGEIFS` | Berekent gemiddelde startprijzen op categorieniveau. |

#### 1. Globale parameters — `00_Param_Config`

Het parameterwerkblad beheert twee soorten besturingselementen:

**Statische configuratieparameters**

| Parameter | Verwijzing | Doel |
|---|---|---|
| Valutasymbool | `B3` | Regelt de weergave van geldbedragen. |
| Omzetbelastingtarief | `B4` | Berekent de eindprijzen voor klanten. |
| Factor linkerhelft | `B5` | Berekent modifierprijzen voor de linkerhelft. |
| Factor rechterhelft | `B6` | Berekent modifierprijzen voor de rechterhelft. |
| Factor hele portie | `B7` | Regelt de prijsberekening voor hele porties. |
| Premium-modifiertoeslag | `B8` | Levert de geconfigureerde premium-prijscorrectie. |
| Waarschuwingsdrempel modifiergroep | `B9` | Markeert buitensporige toewijzingen van modifiergroepen. |
| Waarschuwingsdrempel hoge prijs | `B10` | Identificeert ongebruikelijk hoge productprijzen. |
| Basislijn beschikbaarheid POS-export | `B11` | Definieert de status die wordt gebruikt voor het filteren van de POS-export. |
| Standaard activestatus | `B12` | Centraliseert de identificator van de activestatus. |
| Inbegrepen status | `B13` | Identificeert standaard inbegrepen ingrediënten. |
| Standaard modifierniveau | `B14` | Identificeert het standaard modifierniveau. |

**Voorwaardelijke prijsmatrix**

Het gebied `E3:J10` bevat de voorwaardelijke prijsregels.

De logica combineert:

- Voorwaarde inbegrepen status
- Voorwaarde selectietype
- Resultaat prijsvermenigvuldiger

Er wordt een samengestelde lookupsleutel gegenereerd om de voorwaarden te verbinden met de prijsvermenigvuldiger.

Dit voorkomt dat dezelfde prijsaannames onafhankelijk van elkaar in meerdere werkbladen worden ingebed.

#### 2. Menu-stamgegevens — `01_Menu_Categories`, `02_Menu_Items`, `03_Item_Sizes`

**Categoriehiërarchie**

`01_Menu_Categories!H3` genereert categoriepaden:

```excel
=LET(
    cat_name, FILTER(B3:B100, E3:E100<>""),
    sub_name, FILTER(F3:F100, E3:E100<>""),
    cat_name & " > " & sub_name
)
````

De resulterende paden bieden leesbare categorierelaties, zoals:

```
Pizza > Create Your Own
Pizza > Gourmet Pizza
Subs > Hot Subs
```

Aantal actieve producten per subcategorie

`01_Menu_Categories!I3`:

```
=LET(
    sub_ids, FILTER(E3:E100, E3:E100<>""),
    COUNTIFS(
        '02_Menu_Items'!D3:D500, sub_ids,
        '02_Menu_Items'!G3:G500, '00_Param_Config'!$B$12
    )
)
```

Categorieconfiguratiestatus

`01_Menu_Categories!J3`:

```
=LET(
    cat_active, FILTER(D3:D100, E3:E100<>""),
    item_cnt, I3#,
    IF(cat_active <> '00_Param_Config'!$B$12, "⏸️ 已停用",
        IF(item_cnt = 0, "⚠️ 空分类预警", "✅ 正常")
    )
)
```

Samengestelde product-maat-sleutel

`03_Item_Sizes` genereert een samengestelde sleutel met behulp van:

```
=Item_ID & "|" & Size_ID
```

Bijvoorbeeld:

```
ITM001|SZ_SM
ITM001|SZ_LG
```

Met deze sleutel kan de werkmap onderscheid maken tussen meerdere maten die bij hetzelfde product horen.

Validatie van product-maat-prijzen

De prijsvalidatielaag controleert of een geconfigureerde maatsprijs ontbreekt, ongeldig is of de gecentraliseerde waarschuwingsdrempel overschrijdt.

De resulterende status wordt gebruikt door zowel de productstamgegevens als het POS-exportproces.

#### 3. Modifierstamgegevens & prijzen — `04_Modifier_Groups`, `05_Modifiers`, `06_Modifier_Pricing`

Aantal actieve modifiers

`04_Modifier_Groups!F3`:

```
=LET(
    grp_ids, FILTER(A3:A100, A3:A100<>""),
    COUNTIFS(
        '05_Modifiers'!B3:B500, grp_ids,
        '05_Modifiers'!E3:E500, '00_Param_Config'!$B$12
    )
)
```

Selectieregel-lookup voor modifiergroepen

`04_Modifier_Groups!G3`:

```
=LET(
    grp_ids, FILTER(A3:A100, A3:A100<>""),
    XLOOKUP(
        grp_ids,
        '09_Selection_Rules'!A3:A100,
        '09_Selection_Rules'!B3:B100,
        -1
    )
)
```

De bijbehorende lookup voor het maximale aantal selecties gebruikt kolom C in `09_Selection_Rules`.

Een retourwaarde van `-1` identificeert een ontbrekende regel.

Dimensies van modifierprijzen

`06_Modifier_Pricing` beheert de relatie tussen:

```
Modifier ID
    +
Size ID
    +
Selection Type
```

De samengestelde prijssleutel ondersteunt het ophalen van de juiste modifierprijs voor een geselecteerde productmaat en hoeveelheidsniveau.

Prijzen voor de hele, linker en rechter portie worden afgeleid van de gecentraliseerde prijsfactoren.

De prijsarchitectuur scheidt de referentieprijs van de plaatsingsberekening, zodat restaurantspecifiek prijsbeleid kan worden aangepast zonder de ordersimulator te herschrijven.

#### 4. Standaardingrediënten & relatiemapping — `07_Default_Modifiers`, `08_Group_Assignment`, `09_Selection_Rules`

Lookupsleutel voor standaardmodifiers

`07_Default_Modifiers!E3`:

```
=LET(
    itm_ids, FILTER(A3:A1000, A3:A1000<>""),
    mod_ids, FILTER(B3:B1000, A3:A1000<>""),
    itm_ids & "|" & mod_ids
)
```

Referentie productnaam

`07_Default_Modifiers!F3`:

```
=LET(
    itm_ids, FILTER(A3:A1000, A3:A1000<>""),
    XLOOKUP(
        itm_ids,
        '02_Menu_Items'!A3:A500,
        '02_Menu_Items'!B3:B500,
        "❌ 商品ID无效"
    )
)
```

Referentie modifiernaam

`07_Default_Modifiers!G3`:

```
=LET(
    mod_ids, FILTER(B3:B1000, A3:A1000<>""),
    XLOOKUP(
        mod_ids,
        '05_Modifiers'!A3:A500,
        '05_Modifiers'!C3:C500,
        "❌ 配料ID无效"
    )
)
```

Validatie van de standaardconfiguratie

`07_Default_Modifiers!H3` controleert:

- Of de product-id bestaat.
- Of de modifier-id bestaat.
- Of dezelfde product-modifier-combinatie wordt gedupliceerd.
- Of de inbegrepen status correct is geconfigureerd.

Toewijzingen op categorie- en itemniveau

`08_Group_Assignment` gebruikt het veld `Level` om de toewijzingsreikwijdte te onderscheiden:

```
Level = Category
    Category ID = valid category
    Item ID = -

Level = Item
    Item ID = valid product
    Category ID = -
```

De toewijzingsvalidatielaag controleert de effectieve reikwijdte en identificeert ongeldige of buitensporige toewijzingen.

Selectieregels

`09_Selection_Rules` gebruikt regels op groepsniveau om te bepalen:

- Minimale selecties.
- Maximale selecties.
- Verplichte status.
- Of meervoudige selectie is toegestaan.

Deze waarden worden gerefereerd door de ordersimulator en de validatieformules voor modifiergroepen.

#### 5. Klantprijssimulatie — `10_Order_Test`

Dit werkblad combineert de basisprijs van het product met de werkelijke modifiertoeslagen die door de geselecteerde orderconfiguratie worden gegenereerd.

Naam van het geselecteerde product

`E3`:

```
=XLOOKUP(
    B3,
    '02_Menu_Items'!A3:A500,
    '02_Menu_Items'!B3:B500 & " (" &
    '02_Menu_Items'!F3:F500 & ")",
    "⚠️ 请选择有效商品ID"
)
```

Basisprijs per maat

`E4`:

```
=XLOOKUP(
    B3 & "|" & B4,
    '03_Item_Sizes'!F3#,
    '03_Item_Sizes'!D3:D1000,
    0
)
```

Totale modifiertoeslag

`E5`:

```
=SUM(J10#)
```

Klantprijs vóór belasting

`E6`:

```
=E4+E5
```

Eindprijs voor de klant inclusief belasting

`E7`:

```
=ROUND(
    E6 * (1 + '00_Param_Config'!$B$4),
    2
)
```

De detailberekeningen voor modifiers volgen de volgende volgorde:

```
Selected Modifier ID
        |
        v
Modifier Master Lookup
        |
        v
Default Ingredient Recognition
        |
        v
Conditional Pricing Rule Lookup
        |
        v
Size & Portion Price Retrieval
        |
        v
Placement Adjustment
        |
        v
Calculated Line Price
        |
        v
Group Selection Validation
```

De eindprijs wordt berekend op basis van de daadwerkelijk gesimuleerde selectie en niet op basis van een handmatig samengestelde offerte.

#### 6. POS-export — `11_POS_Export`

Het POS-exportwerkblad filtert producten dynamisch op basis van hun geconfigureerde beschikbaarheid.

Filteren van beschikbare producten

`A3`:

```
=LET(
    sz_itm_ids,
    FILTER(
        '03_Item_Sizes'!A3:A1000,
        '03_Item_Sizes'!A3:A1000<>""
    ),
    itm_avail,
    XLOOKUP(
        sz_itm_ids,
        '02_Menu_Items'!A3:A500,
        '02_Menu_Items'!G3:G500,
        "No"
    ),
    FILTER(
        sz_itm_ids,
        itm_avail = '00_Param_Config'!$B$11
    )
)
```

Naam van het exportproduct

`B3`:

```
=XLOOKUP(
    A3#,
    '02_Menu_Items'!A3:A500,
    '02_Menu_Items'!B3:B500,
    ""
)
```

Naam van de exportmaat

`D3`:

```
=XLOOKUP(
    A3# & "|" & C3#,
    '03_Item_Sizes'!F3#,
    FILTER(
        '03_Item_Sizes'!C3:C1000,
        '03_Item_Sizes'!A3:A1000<>""
    ),
    ""
)
```

Basisprijs van de export

`E3`:

```
=XLOOKUP(
    A3# & "|" & C3#,
    '03_Item_Sizes'!F3#,
    FILTER(
        '03_Item_Sizes'!D3:D1000,
        '03_Item_Sizes'!A3:A1000<>""
    ),
    0
)
```

POS-synchronisatiestatus

De exportvalidatie combineert productintegriteit en validatie van maatsprijzen.

De beoogde uitvoer is:

```
READY_FOR_POS
HOLD_CHECK_ERROR
```

De export is ontworpen om wijzigingen in de productconfiguratie gemakkelijker te beoordelen en over te zetten. Het is geen directe POS-API-integratie of automatische synchronisatiedienst.

#### 7. Dashboard & configuratiemonitoring — `12_Dashboard`

Het dashboard combineert zes operationele indicatoren:

| KPI                       | Berekeningsdoel                           |
| ------------------------- | --------------------------------------------- |
| Totaal actieve categorieën   | Telt ingeschakelde menucategorieën.               |
| Totaal beschikbare items     | Telt beschikbare producten.                    |
| Totaal herbruikbare modifiers  | Telt ingeschakelde modifiers.                     |
| Totaal aantal prijsregels | Telt geconfigureerde prijsregels.              |
| Totaal configuratieanomalieën    | Aggregeert gedetecteerde configuratie-uitzonderingen. |
| Aantal POS-gereede items      | Telt records die als exportgereed zijn gemarkeerd.       |

Totaal beschikbare items

`C4`:

```
=COUNTIF(
    '02_Menu_Items'!G3:G500,
    '00_Param_Config'!$B$12
)
```

Totaal configuratieanomalieën

`F4`:

```
=COUNTIF('02_Menu_Items'!M3#, "❌*")
+ COUNTIF('03_Item_Sizes'!I3#, "❌*")
+ COUNTIF('04_Modifier_Groups'!I3#, "❌*")
+ COUNTIF('08_Group_Assignment'!I3#, "⚠️*")
+ COUNTIF('09_Selection_Rules'!H3#, "❌*")
```

Aantal POS-gereede records

`G4`:

```
=COUNTIF(
    '11_POS_Export'!I3#,
    "READY_FOR_POS"
)
```

Het dashboard biedt een geconsolideerd overzicht van de gezondheid van de menuconfiguratie, terwijl de onderliggende werkbladen de gedetailleerde records behouden die nodig zijn om afzonderlijke uitzonderingen te onderzoeken en te corrigeren.

### Validatieregels

Het validatiekader is ontworpen om fouten te identificeren voordat deze klantgerichte prijzen of de POS-configuratie beïnvloeden.

| Veld / gebied                | Validatieregel                                                              | Foutgedrag                                                      |
| --------------------------- | ---------------------------------------------------------------------------- | ------------------------------------------------------------------- |
| Categorie-id                 | Moet uniek zijn en verwijzen naar een geldige categorie.                               | Ongeldige categorierelaties worden gemarkeerd.                         |
| Subcategorie-id              | Moet overeenkomen met een bestaande subcategorie.                                  | Ongeldige productclassificatie wordt geïdentificeerd.                       |
| Categorie actief             | Moet de geconfigureerde Ja/Nee-statuswaarden gebruiken.                                | Ongeldige statuswaarden beïnvloeden de categorievalidatie.                   |
| Item-id                     | Moet een unieke productidentificator zijn.                                         | Ongeldige of dubbele productidentiteiten vereisen correctie.         |
| Itemcategorie               | Moet verwijzen naar een geldige categorie en subcategorie.                             | De productintegriteitsvalidatie identificeert inconsistente relaties. |
| Itembeschikbaarheid           | Moet de geconfigureerde activestatus-identificator gebruiken.                            | Ongeldige beschikbaarheid kan correcte filtering verhinderen.                 |
| Itemmaat                   | Moet verwijzen naar een geldige product-maat-combinatie.                         | Ontbrekende of ongeldige product-maat-records worden gemarkeerd.                |
| Basisprijs                  | Moet aanwezig zijn en voldoen aan de geconfigureerde prijscontroles.                   | Ontbrekende of afwijkende prijzen verhinderen POS-gereedheid.                  |
| Modifiergroep-id           | Moet een unieke groepsidentificator zijn.                                           | Ongeldige relaties worden gemarkeerd.                                  |
| Modifier-id                 | Moet uniek zijn en verwijzen naar een geldige modifiergroep.                         | Modifier-lookupfouten worden geïdentificeerd.                              |
| Modifier actief             | Moet de geconfigureerde activestatus-identificator gebruiken.                            | Inactieve modifiers worden uitgesloten van tellingen van actieve modifiers.        |
| Modifierprijs            | Moet de vereiste prijsrecords voor maat en selectietype bevatten.           | Onvolledige prijsmatrices vereisen correctie.                     |
| Modifierniveau               | Moet overeenkomen met de geconfigureerde identificatoren voor modifierniveaus.                 | Ongeldige niveauclassificatie beïnvloedt de prijsvalidatie.             |
| Standaardmodifiertoewijzing    | Product- en modifier-id's moeten bestaan.                                         | Ongeldige id's worden gemarkeerd.                                            |
| Uniciteit van standaardmodifiers | Dezelfde product-modifier-combinatie mag niet worden gedupliceerd.                | Dubbele toewijzingen worden gemarkeerd.                                     |
| Inbegrepen status             | Moet overeenkomen met de geconfigureerde identificator van de inbegrepen status waar van toepassing.       | Onjuiste configuratie van standaardinclusie wordt gemarkeerd.               |
| Toewijzingsniveau            | Moet `Category` of `Item` zijn.                                         | Ongeldige toewijzingsreikwijdte wordt gemarkeerd.                                |
| Toewijzing op categorieniveau   | Vereist een geldige categorie-id en de voorgeschreven item-id-placeholder.         | Ongeldige reikwijdteconfiguratie wordt gemarkeerd.                             |
| Toewijzing op itemniveau       | Vereist een geldige item-id en de voorgeschreven categorie-id-placeholder.         | Ongeldige reikwijdteconfiguratie wordt gemarkeerd.                             |
| Toewijzing van modifiergroep   | Moet verwijzen naar een bestaande modifiergroep.                                   | Ongeldige groepsrelaties worden gemarkeerd.                            |
| Capaciteit van modifiergroep     | Het maximale aantal selecties mag het beschikbare aantal actieve modifiers niet overschrijden. | De selectieconfiguratie vereist correctie.                        |
| Verplichte modifiergroep     | Verplichte groepen moeten een minimale selectie groter dan nul hebben.             | Logisch conflict wordt gemarkeerd.                                        |
| Optionele modifiergroep     | Optionele groepen moeten een minimale selectie van nul hebben.                     | Inconsistente selectieregels worden gemarkeerd.                           |
| Meervoudige selectie          | Moet de geconfigureerde machtiging voor meervoudige selectie volgen.                    | Ongeldige selectieconfiguraties vereisen beoordeling.                    |
| Klanttestproduct       | Moet verwijzen naar een geldig product.                                              | Ongeldige productselectie toont een waarschuwing.                       |
| Klanttestmaat          | Moet verwijzen naar een geldige maat voor het geselecteerde product.                        | Een ontbrekende maatsprijs-lookup geeft een ongeldig of nulprijsresultaat.  |
| Klantmodifierselectie | Moet verwijzen naar een beschikbare modifier.                                        | Ongeldige modifierverwijzingen vereisen correctie.                     |
| Klantportietype       | Moet `Light`, `Normal`, `Extra` of `Xxtra` zijn.                             | Ongeldige selectiewaarden zijn geen geldige prijsinvoer.              |
| Klantplaatsing          | Moet `Whole`, `Left` of `Right` zijn.                                        | Ongeldige plaatsingswaarden zijn geen geldige prijsinvoer.              |
| Voorwaardelijke prijs         | Moet uitkomen op de geconfigureerde inclusie- en hoeveelheidsregel.                  | Onjuiste of ontbrekende prijsconfiguratie vereist onderzoek.  |
| Klantprijs              | Moet de basisprijs, modifiertoeslagen en het geconfigureerde belastingtarief weerspiegelen.      | De ordersimulator legt de resulterende berekening ter beoordeling bloot.   |
| POS-export                  | De validatie van product en maatsprijs moet slagen.                                 | `HOLD_CHECK_ERROR` vervangt `READY_FOR_POS`.                        |

Belangrijk operationeel onderscheid: de validatielaag controleert de geconfigureerde relaties en zakelijke regels die in de werkmap zijn vastgelegd. Zij verifieert niet onafhankelijk of een POS-leverancier elk geëxporteerd veld accepteert of elk prijsmechanisme ondersteunt.

</details>


## Andere tools in deze serie

Meer lichtgewicht zakelijke beslissingsondersteunende tools en operationele werkmappen:

\- [Theseus Workshop](https://theseusworkshop.com) — Een verzameling praktische tools voor terugkerende operationele en financiële problemen van kleine bedrijven.

## Licentie

Dit project is gelicentieerd onder de Apache License 2.0.

Zie het LICENSE-bestand voor de volledige licentievoorwaarden.
