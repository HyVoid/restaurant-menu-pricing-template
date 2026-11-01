[ 🌐 عربي ](README.ar.md) | [ 🇩🇪 Deutsch ](README.de.md) | [ 🇫🇷 Français ](README.fr.md) | [ 🇳🇱 Nederlands ](README.nl.md) | [ 🇪🇸 Español ](README.sp.md) | [ 🇬🇧 English ](README.md)

# Restaurant-Menü-Preisrechner, Modifier-Konfiguration & POS-Vorbereitungs-Toolkit (Excel & Web-Vorlage)
<img width="1920" height="1032" alt="image" src="https://github.com/user-attachments/assets/45a7c5cb-ec7e-4744-a38b-e705561fe0d6" />


[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-Web%20App%20%2B%20Excel%20Template-2251FF)
![Tool Type](https://img.shields.io/badge/Tool-Restaurant%20Menu%20Management%20%26%20POS%20Pricing-051C2C)
![Architecture](https://img.shields.io/badge/Architecture-Formula--Driven%20No--VBA-success)

**Ein kostenloser Online-Restaurant-Menü-Preisrechner und eine Offline-Excel-POS-Konfigurationsvorlage für unabhängige Restaurants, Pizzerien und QSR-Betreiber. Zentralisieren Sie Menü-Stammdaten, berechnen Sie komplexe Modifier-Aufpreise, konfigurieren Sie Belag-Regeln für halbe und ganze Pizzen, verhindern Sie Doppelabbuchungen bei bereits in der Rezeptur enthaltenen Zutaten und exportieren Sie validierte, POS-fertige Menü-Tabellenblätter – ganz ohne VBA-Makros.**

> **Testen Sie Ihre Menü-Preisregeln sofort in Ihrem Browser** — Keine Konto-Registrierung, keine Software-Installation und keine komplexen Tabellenformeln erforderlich.
>
> - 🌐 **[Starten Sie die kostenlose interaktive Restaurant-Menü-Preisrechner-Web-App](https://hyvoid.github.io/restaurant-menu-pricing-template/)** *(Testen Sie Einzel-Sitzungs-Menükonfigurationen, Modifier-Logik und Kundenbestellpreise online)*
> - 📥 **[Laden Sie die Offline-Restaurant-POS-Menü-Konfigurations-Excel-Vorlage (.xlsx) herunter](https://theseusworkshop.com/l/dznbyd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=restaurant-menu-pricing-template)** *(Gebaut für dauerhafte Menü-Stammdatensätze, mehrsaisonale Preisaktualisierungen und betriebliche POS-Prüfprotokolle — mit einer 30-tägigen bedingungslosen Geld-zurück-Garantie)*

---

## Was diese Restaurant-Menü-Verwaltungsvorlage Ihnen hilft nachzuverfolgen (Schmerzpunkt-zu-Lösung-Zuordnung)

| Kern-Menüdaten & Preiskennzahl | Häufiger betrieblicher Schmerzpunkt in Restaurants | Wie dieses Menü-Konfigurations-Toolkit das löst |
|---|---|---|
| **Menü-Stammhierarchie & Größenverfügbarkeit** | Menükategorien, Artikelgrößen (Klein/Mittel/Groß) und saisonale Verfügbarkeit sind verstreut über veraltete PDFs und POS-Exporte. | Erfasst aktive Kategorien, Unterkategorien, Artikel-SKUs und größen-spezifische Verfügbarkeit in einer relationalen Menü-Stammdatenbank. |
| **Basis-Menüpreise & gestaffelte Modifier-Aufpreise** | Mitarbeitende schätzen Zusatzbelag-Gebühren oder berechnen Pauschalgebühren, die Größenunterschiede ignorieren, und verschleiern so die Food-Gewinnmargen. | Berechnet exakte Basispreise plus Standard- und Premium-Modifier-Zuschläge, skaliert nach Produktgröße. |
| **Standard-Rezepturzutaten (keine Doppelabbuchung)** | Kunden, die Spezialartikel bestellen (z. B. Hawaiian- oder Supreme-Pizza), werden doppelt belastet, wenn sie die standardmäßig enthaltenen Beläge behalten. | Ordnet jedem Artikel die standardmäßig enthaltenen Zutaten zu, sodass Standardmengen $0 extra kosten, während „Extra“-Portionen automatisch einen Aufpreis auslösen. |
| **Teilbetrag- & halbe-Pizza-Belagspreise** | Manuelle Berechnung von Belägen für linke Hälfte, rechte Hälfte und ganze Pizza führt zu Verzögerungen an der Kasse und inkonsistenten Kundenrechnungen. | Wendet globale fraktionale Preismultiplikatoren an, um halbe und ganze Belagsportionen automatisch korrekt zu berechnen. |
| **Min/Max-Auswahllimits von Modifier-Gruppen** | Kunden wählen ungültige Kombioptionen oder überspringen Pflichtauswahlen (z. B. Bodenart oder Gargrad), weil die POS-Regeln fehlerhaft sind. | Erzwingt strenge Mindest-, Höchst- und Mengenlimits pro Modifier auf der Ebene der Modifier-Gruppe vor dem Live-Einsatz. |
| **POS-Vorbereitung & QA-Status vor der Menü-Synchronisierung** | Das Hochladen unvollständiger Artikeldatensätze oder ungeprüfter Preise direkt in ein Live-POS zerstört Küchenbons und Online-Bestellungen. | Markiert fehlende Preise und defekte Modifier-Verknüpfungen auf einem Validierungs-Dashboard und isoliert saubere, POS-fertige Exportzeilen. |

---

## Schritt-für-Schritt-Anleitung: So konfigurieren Sie Restaurant-Menüpreise & bereiten POS-Importe vor

Folgen Sie diesem 4-stufigen Menü-Engineering-Workflow, um Ihre Artikeldatenbank zu standardisieren, komplexe Bestellkombinationen zu testen und ein sauberes POS-Import-Tabellenblatt zu erzeugen:

1. **Schritt 1: Globale Restaurant-Steuer-, Währungs- & Fraktionspreisparameter konfigurieren**
   **Aktion:** Öffnen Sie das `Parameters`-Tab und geben Sie Ihren lokalen Umsatzsteuersatz, Ihr Währungssymbol, den Preis-Multiplikator für halbe Beläge und den Premium-Modifier-Aufpreisfaktor ein.
   *Warum das wichtig ist:* Das einmalige Festlegen globaler Preisvariablen schafft eine einheitliche Berechnungsbasis über alle Menükategorien und Artikelgrößen hinweg.

2. **Schritt 2: Menü-Stammliste, Größen & Modifier-Gruppen importieren oder eingeben**
   **Aktion:** Befüllen Sie die strukturierten Eingabetabellen mit Ihren Menükategorien, Gerichtnamen, Größenvarianten, Modifier-Gruppen und Roh-Zutatenpreisen – oder fügen Sie bestehende Zeilen aus Ihrem aktuellen POS-CSV-Export ein.
   *Warum das wichtig ist:* Die Verknüpfung von Artikeln und Modifiern über saubere relationale IDs eliminiert doppelte Dateneingabe, wenn sich Zutatenpreise ändern.

3. **Schritt 3: Standard-Rezepturbestandteile zuordnen & echte Kundenbestellungen simulieren**
   **Aktion:** Weisen Sie Modifier-Gruppen bestimmten Menüartikeln zu, markieren Sie eingebaute Spezialzutaten und führen Sie Testbestellungen im interaktiven **Bestellpreis-Simulator** durch (testen Sie Größenwechsel, linke/rechte halbe Beläge und Zusatzportionen).
   *Warum das wichtig ist:* Die Echtzeit-Simulation von Bestellungen vor und nach Steuern deckt Extras mit zu niedrigen Preisen und Doppelabbuchungsfehler auf, bevor auch nur ein Kunde eine Bestellung aufgibt.

4. **Schritt 4: Konfigurationsqualität prüfen & langfristige Menüpflege mit Excel verankern**
   **Aktion:** Prüfen Sie das automatisierte QA-Dashboard, um unvollständige Artikel-Markierungen zu bereinigen, exportieren Sie Ihre validierte POS-Vorbereitungstabelle und **wechseln Sie von der Einzel-Sitzungs-Browser-Demo zur wiederverwendbaren Excel-Arbeitsmappe**.
   *Warum das wichtig ist:* Während der [kostenlose Browser-Menürechner](https://hyvoid.github.io/restaurant-menu-pricing-template/) ideal für schnelle Einmaltests ist, speichern Browser-Sitzungen historische Änderungen nicht dauerhaft. **[Laden Sie die vollständige Offline-Excel-Menü-Stammvorlage (.xlsx) herunter](https://theseusworkshop.com/l/dznbyd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=restaurant-menu-pricing-template)**, um die dauerhafte Menüdatenbank Ihres Restaurants zu speichern, wiederkehrende vierteljährliche Preisaktualisierungen durchzuführen, historische Prüfprotokolle zu pflegen und Ihre Konfiguration über mehrere Standorte hinweg wiederzuverwenden, ohne bei null anzufangen.

**Konfigurieren Sie globale Preisregeln einmal. Pflegen Sie einen sauberen Menüstamm. Simulieren Sie echte Kunden-Kassenpreise. Exportieren Sie erst nach automatisierter Validierung in Ihr POS.**

---

## Warum Standard-Restaurant-Menü-Tabellen bei Modifier- & Belagspreisen scheitern

Die Verwaltung einer Restaurantspeisekarte wird tückisch komplex, sobald ein Gericht nicht mehr durch einen einzelnen Pauschalpreis definiert ist.

Denken Sie an eine Spezialpizza oder ein anpassbares Kombi-Menü mit größenabhängiger Preisgestaltung, standardmäßig enthaltenen Belägen, optionalen Premium-Extras und kundenwählbaren Mengen. Eine herkömmliche Excel- oder Google-Sheets-Menüliste kann Basisartikelpreise erfassen, aber sie trennt nicht **Zutaten, die bereits in der Basisrezeptur enthalten sind**, von **zusätzlichen Modifiern, die einen Aufpreis erfordern**.

Dieser strukturelle blinde Fleck erzeugt zwei kostspielige geschäftliche Risiken für Gastronomiebetriebe:

1. **Die Doppelabbuchungs-Falle für Kunden:** Kunden werden versehentlich zusätzlich für Standardzutaten belastet, die bereits im Basispreis eines Spezialgerichts enthalten sind – das führt zu Kassenstreitigkeiten und abgebrochenen Online-Bestellungen.
2. **Stille Erosion der Gewinnmargen:** Kassierer und POS-Administratoren wenden inkonsistente Preise auf Extra-Beläge, Größenupgrades oder Halb-Halb-Pizza-Auswahlen an, weil die Preisregeln von Hand berechnet werden.

Zum Beispiel enthält eine **Hawaiian-Pizza** normale Portionen Schinken und Ananas in ihrem Basispreis. Die Beibehaltung der Standardmengen von Schinken und Ananas darf **$0.00** zur Rechnung hinzufügen. Bestellt der Kunde jedoch *extra* Ananas, muss ein größenangepasster Zuschlag ausgelöst werden – und ein Premium-Belag nur auf der **linken Hälfte** sollte einer definierten fraktionalen Preisregel folgen und nicht dem Rätselraten.

Das sind keine einfachen Tippfehler. Sie passieren, weil flache Tabellen die kommerziellen Regeln hinter einer modernen Restaurant-Speisekarte nicht abbilden können.

Dieses Toolkit verwandelt diese Menübeziehungen in ein strukturiertes, formelgetriebenes Konfigurationssystem. Statt doppelte Belagslisten für jeden Artikel zu pflegen und Preise manuell zu prüfen, definieren Sie Zutaten, Modifier-Gruppen und Preisregeln einmal – und wenden sie konsistent auf Ihre gesamte Speisekarte an.

---

## Manuelle Menü-Tabellen vs. strukturiertes POS-Menü-Konfigurations-Toolkit

| Restaurant-Menü-Konfigurations-Herausforderung | Manuelle Tabellen & direkte POS-Bearbeitungen (ohne dieses Tool) | Strukturiertes Menü-Stamm- & Preis-Toolkit (mit diesem Tool) |
|---|---|---|
| **Fragmentierung der Menü-Stammdaten** | Kategorien, Artikelgrößen und Modifier-Preise liegen in unverbundenen Tabs oder unsortierten CSV-Exporten und verursachen Versionschaos bei Menü-Updates. | Relationale Stammtabellen trennen Kategorien, Basisartikel und größen-spezifische Preismatrizen und bewahren dabei automatische ID-Lookups. |
| **Doppelte Modifier- & Belagspflege** | Derselbe Käse, dieselbe Sauce oder derselbe Belag wird über Dutzende Menüartikel hinweg wiederholt eingegeben, was zu uneinheitlichen Namen und veralteten Aufpreispreisen führt. | Jeder Zutaten-Modifier wird einmal in einer zentralen Bibliothek angelegt und über wiederverwendbare Modifier-Gruppen mit mehreren Gerichten verknüpft. |
| **Preisfehler bei Spezialartikeln & Standardrezepturen** | Standardmäßig enthaltene Beläge auf Spezialpizzen oder Burgern werden fälschlich als bezahlte Extras abgerechnet, wenn Kunden ihre Bestellung anpassen. | Das Standardrezeptur-Mapping unterscheidet automatisch eingebaute Standardmengen ($0-Berechnung) von bezahlten „Extra“-Upgrades. |
| **Inkonsistente halbe-Pizza- & Fraktionspreise** | Mitarbeitende berechnen Belagsgebühren für linke Hälfte, rechte Hälfte und ganze Pizza manuell, was zu Rechnungsfehlern und Rätselraten an der Kasse führt. | Globale fraktionale Preisfaktoren berechnen halbe Belagsportionen automatisch aus dem Referenzsatz für ganze Pizzen. |
| **Defekte POS-Modifier-Auswahlregeln** | Mindestauswahlen und maximale Belagsgrenzen werden beim Setup vergessen, sodass unmögliche Bestellungen das Küchendisplay erreichen. | Constraint-Regeln auf Modifier-Gruppen-Ebene erzwingen explizite Min/Max-Auswahllimits und Mengengrenzen pro Artikel während der POS-Vorbereitung. |
| **Live-POS-Bereitstellungs- & Synchronisierungsfehler** | Ungeprüfte Preise und unvollständige Artikelzeilen werden direkt in Live-POS-Terminals eingegeben und verursachen Kassenausfälle während der Stoßzeiten. | Eingebaute QA-Validierung markiert fehlende Daten, simuliert Kundenbons und erzeugt eine saubere, vorab geprüfte POS-Exporttabelle. |

---

## Für wen diese Restaurant-Menü- & Preisvorlage gedacht ist (Zielrollen & Anwendungsfälle)

Gebaut für unabhängige Restaurants, Pizzerien, Schnellrestaurants (QSR), Fast-Casual-Konzepte und Gastronomiebetriebe mit mehreren Standorten, die Menüs in mehreren Größen, gemeinsame Zutaten-Modifier, bedingte Aufpreise und strenge Kundenbestellregeln verwalten.

### Rollenbasierte Suchszenarien & betriebliche Anwendungsfälle

- **Pizzeria-Betreiber & unabhängige Restaurantbetreiber (`Pizza Menu Pricing Calculator & Topping Spreadsheet`)**
  Verlieren Sie kein Geld mehr an nicht abgerechnete Extra-Beläge oder berechnen Sie Kunden auf Spezialpizzen nicht zu viel. Modellieren Sie Multiplikatoren für ganze vs. halbe Pizzen, größenabhängige Zuschläge und enthaltene Rezepturzutaten, bevor Sie Ihre physischen Tafeln oder Ihr Online-Bestellmenü aktualisieren.
- **QSR- & Fast-Casual-Betriebsmanager (`Multi-Size Restaurant Menu Management Template`)**
  Standardisieren Sie Menühierarchien, Kombi-Modifier-Gruppen und saisonale Preisanpassungen über mehrere Produktkategorien hinweg in einer zentralen Offline-Stammdatei – ohne auf fehleranfällige VBA-Skripte angewiesen zu sein.
- **POS-Onboarding-Spezialisten & Menü-Administratoren (`Restaurant POS Menu Staging & Import Checklist`)**
  Bereinigen, strukturieren und validieren Sie rohe Menüdaten, bevor Sie Datenbanken in Toast, Square, Clover, Lightspeed oder eigenen POS-Plattformen aufbauen oder aktualisieren. Finden Sie fehlende Basispreise und defekte Modifier-Min/Max-Regeln, bevor Sie live gehen.
- **Restaurantberater & Menü-Engineers (`Menu Engineering & Modifier Architecture Toolkit`)**
  Liefern Sie Hospitality-Kunden eine wiederholbare, prüfungsfähige Menükonfigurations-Arbeitsmappe. Strukturieren Sie chaotische Kundenmenüs in logische Modifier-Gruppen um und demonstrieren Sie bei Beratungseinsätzen die exakten Kunden-Kassenpreise.

**Für wen dieses Toolkit NICHT gedacht ist:**
Gastronomiebetriebe, die ein Live-Point-of-Sale-Terminal, einen Kreditkarten-Zahlungsabwickler, ein Küchendisplay-System (KDS), einen Rohbestands-Verbrauchstracker oder einen automatisierten Rezeptur-Nährwertrechner suchen. Dieses Toolkit ist ein **Menü-Stammsystem vor dem POS, ein Preissimulator und ein Konfigurations-QA-System** – es bereitet Ihre kommerzielle Menülogik vor und validiert sie, damit Ihre POS-Plattform ohne Preisfehler läuft.

---

## Die Geschäftslogik & Methodik: Versteckte blinde Flecken bei Restaurantpreisen beseitigen

*(Geschrieben für Restaurantbesitzer und Betreiber – keine technischen oder fortgeschrittenen Tabellenkalkulationskenntnisse erforderlich)*

Wenn Kunden Essen bestellen, wirkt Menüpreisgestaltung auf den ersten Blick einfach: Artikel auswählen, Belag hinzufügen, Gesamtsumme zahlen. Hinter der Theke verlieren die meisten unabhängigen Restaurants jedoch Geld oder frustrieren Kunden wegen **drei versteckter kommerzieller blinder Flecken**, die Standard-Tabellen nicht erkennen.

So schützt die Geschäftslogik in diesem Toolkit Ihre Gewinnmargen und das Kundenerlebnis – einfach erklärt:

### 1. Das Problem „Eine Aktualisierung, zehn Fehler“ beseitigen (Single Source of Truth)
* **Der blinde Fleck:** Angenommen, der Großhandelspreis für Mozzarella oder Bacon springt um 20 %. In einer typischen Restaurant-Tabelle wird „Extra Bacon“ separat unter Burgern, Sandwiches, Salaten und Pizzen aufgeführt. Aktualisieren Sie den Preis an drei Stellen und vergessen die vierte, verlieren Sie entweder Gewinnmarge oder verwirren Kunden mit inkonsistenten Preisen.
* **So löst es das Toolkit:** Sie geben jede Zutat und jeden Modifier **einmal** in einer zentralen Liste ein. Ob diese Zutat bei 2 Artikeln oder 50 Artikeln verwendet wird – die Aktualisierung ihres Preises in einer einzigen Zelle aktualisiert automatisch jeden verbundenen Menüartikel und jede Bestellberechnung in der gesamten Arbeitsmappe.

### 2. Die Falle der „Spezialartikel-Doppelabbuchung“ verhindern (bedingte Rezepturpreisgestaltung)
* **Der blinde Fleck:** Wenn ein Kunde eine „Meat Lover's Pizza“ oder einen „Deluxe-Burger“ bestellt, zahlt er bereits einen höheren Basispreis, weil Salami, Wurst oder Käse in der Rezeptur enthalten sind. Behandelt Ihr Menü-Setup jeden Belag gleich, berechnet Ihr System dem Kunden möglicherweise *erneut* etwas, nur weil er die Standardzutaten bestätigt – oder wenn Sie alle Beläge kostenlos machen, um Beschwerden zu vermeiden, verlieren Sie Geld, sobald jemand *doppeltes* Fleisch verlangt.
* **So löst es das Toolkit:** Die Vorlage trennt **„Standardmäßig enthalten“** von **„Bezahlt Extra“**. Sie weiß, dass 1x Salami auf einer Salami-Pizza $0.00 extra kostet, während die Erhöhung der Menge auf 2x (Extra Salami) automatisch den exakten Upgrade-Preis berechnet. Sie schützen das Vertrauen der Kunden, ohne Essen zu verschenken.

### 3. Kassenrätselraten bei halben Belägen & Größen stoppen (proportionale Preislogik)
* **Der blinde Fleck:** Kunden fragen häufig nach Belägen nur auf der linken oder rechten Hälfte einer Pizza oder wechseln zwischen Klein-, Mittel- und Großgrößen. Ohne eine klare Formel berechnen Mitarbeitende entweder den vollen Ganz-Pizza-Belagspreis (verärgert die Kunden) oder halbieren Preise manuell ohne Berücksichtigung des Zubereitungsaufwands (schadet Ihren Margen).
* **So löst es das Toolkit:** Sie legen eine einzige Regel fest – z. B. „Ein Belag auf der halben Pizza kostet 60 % eines Belags auf der ganzen Pizza.“ Das Toolkit berechnet automatisch den exakten, centgenauen Betrag für Platzierungen auf linker Hälfte, rechter Hälfte und ganze Pizza in jeder Pizzagröße.

### 4. Unmögliche Küchenbons blockieren (Min/Max-Auswahl-Schutzleisten)
* **Der blinde Fleck:** Bestellt ein Kunde online ein Kombi-Menü und vergisst, eine Pflichtbeilage zu wählen – oder wählt 12 kostenlose Saucen, obwohl nur 2 erlaubt sind – muss Ihr Küchenpersonal das Kochen unterbrechen, den Kunden anrufen oder die Kosten schlucken.
* **So löst es das Toolkit:** Sie definieren einfache Grenzregeln für jede Optionsgruppe (z. B. *„Mindestens 1 Boden wählen; maximal 1“* oder *„Bis zu 4 Standardbeläge wählen“*). Der eingebaute Bestellsimulator prüft Ihr Menü sofort gegen diese Regeln und zeigt Ihnen, ob ein Kunde versehentlich eine ungültige Bestellung aufgeben könnte.

### 5. Testen, bevor Sie live gehen (POS-vorgelagerte Qualitätssicherung)
* **Der blinde Fleck:** Die meisten Restaurants entdecken Menüpreisfehler *erst nachdem* sie ihr POS-System aktualisiert haben – meist wenn ein Stoßzeiten-Kunde an der Kasse reklamiert oder eine Online-Liefer-App den falschen Gesamtbetrag berechnet.
* **So löst es das Toolkit:** Das Toolkit fungiert als sicherer „Crashtest-Dummy“ für Ihr Menü. Sie können echte Kundenbestellungen simulieren, den exakten Bon vor und nach Steuern sehen und ein automatisches Status-Dashboard prüfen, das fehlende Preise rot markiert. Sie exportieren Ihr Menü erst in Ihr POS, wenn jeder Artikel die Prüfung bestanden hat.

---

## Über den Ersteller

Ich baue praktische, formelgetriebene Business-Tools, die chaotische Betriebsdaten in strukturierte, wiederholbare Entscheidungsunterstützungs-Workflows verwandeln.

Die meisten unabhängigen Restaurants und Kleinunternehmen haben ihre Preisdaten bereits, aber sie sind gefangen in unverbundenen Tabellen, unordentlichen POS-Exporten und undokumentierten Gewohnheiten der Mitarbeitenden. Die eigentliche betriebliche Herausforderung besteht darin, diese Aufzeichnungen zu verbinden, konsistente Preisregeln durchzusetzen und Konfigurationsfehler abzufangen, bevor sie zahlende Kunden erreichen.

Das **Restaurant Menu Pricing Calculator & POS Configuration Toolkit** gibt Gastronomiebetreibern Enterprise-Grade-Menüvalidierung in einem transparenten, schlanken Browser- und Excel-Workflow.

- 🌐 **[Testen Sie jetzt den kostenlosen Browser-Menüpreisrechner](https://hyvoid.github.io/restaurant-menu-pricing-template/)**
- 📥 **[Holen Sie sich die vollständige wiederverwendbare Excel-Menüstamm-Arbeitsmappe (.xlsx)](https://theseusworkshop.com/l/dznbyd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=restaurant-menu-pricing-template)**


## Technische Details

<details>
<summary>Für technische Prüfer, Excel-Praktiker und Entwickler — aufklappen, um Arbeitsmappen-Architektur, Preislogik, Formeln und Validierungsregeln einzusehen.</summary>

### Arbeitsmappen-Architektur

Das Toolkit ist als relationale Menükonfigurations-Datenbank und Preisvalidierungs-Engine konzipiert und nicht als Sammlung unabhängiger Menü-Tabellenblätter.

Seine Architektur folgt drei Prinzipien:

1. **Single Source of Truth:** Jede Kategorie, jedes Produkt, jede Größe, jeder Modifier und jede Preisregel hat einen festen Stammdaten-Speicherort.
2. **Wiederverwendbare Beziehungen:** Menüartikel erben Modifier-Gruppen über Zuweisungen auf Kategorie- oder Artikelebene, statt Konfigurationsdatensätze zu duplizieren.
3. **Automatisierte Validierung und Ausgabe:** Dynamische Formeln verbinden die Stammdaten, berechnen Kundenpreise, erkennen Konfigurationsfehler und erzeugen standardisierte POS-Exportdatensätze.

Die Arbeitsmappe enthält fünf Funktionsmodule und 13 Arbeitsblätter.

#### 1. Globale Parameter & Preisregeln

| Arbeitsblatt | Zuständigkeit |
|---|---|
| `00_Param_Config` | Zentralisiert Währung, Steuersatz, Faktoren für ganze/halbe Portionen, Premium-Modifier-Zuschläge, betriebliche Schwellenwerte, Statuskennzeichnungen und bedingte Preisregeln. |

Dieses Arbeitsblatt ist die globale Steuerungsebene. Preisannahmen und Geschäftsstatus-Kennungen werden zentral gepflegt, statt als feste Werte in einzelnen Berechnungsformeln eingebettet zu sein.

#### 2. Menü-Produktstammdaten

| Arbeitsblatt | Zuständigkeit |
|---|---|
| `01_Menu_Categories` | Pflegt Menükategorien, Unterkategorien, Anzeigereihenfolge und Kategorieverfügbarkeit. |
| `02_Menu_Items` | Speichert einzelne Produktidentitäten, Beschreibungen, Kategoriebeziehungen, Produkttypen und Verfügbarkeit. |
| `03_Item_Sizes` | Speichert größen-spezifische Produktpreise und verknüpft jede Größe mit ihrem übergeordneten Menüartikel. |

Die Trennung zwischen Produktidentität und größen-spezifischer Preisgestaltung verhindert, dass sich der Produktstamm bei jeder neuen Größe oder Preioption horizontal ausweitet.

#### 3. Wiederverwendbare Modifier & mehrdimensionale Preisgestaltung

| Arbeitsblatt | Zuständigkeit |
|---|---|
| `04_Modifier_Groups` | Definiert kundensichtbare Modifier-Gruppen und ihren betrieblichen Geltungsbereich. |
| `05_Modifiers` | Pflegt global wiederverwendbare Zutaten- und Modifier-Identitäten. |
| `06_Modifier_Pricing` | Speichert Modifier-Preise nach Produktgröße und Portionsebene, mit Berechnungen für die Platzierung auf ganzer Pizza, linker Hälfte oder rechter Hälfte. |

Dieselbe Zutat kann über mehrere Produkte hinweg wiederverwendet werden, ohne separate Kopien ihres Stammdatensatzes anzulegen.

#### 4. Beziehungsmapping & Auswahl-Constraints

| Arbeitsblatt | Zuständigkeit |
|---|---|
| `07_Default_Modifiers` | Ordnet Spezialprodukten ihre enthaltenen Zutaten und Standardmengen zu. |
| `08_Group_Assignment` | Weist Modifier-Gruppen entweder auf Kategorieebene oder auf Artikelebene zu. |
| `09_Selection_Rules` | Steuert Mindestauswahlen, Höchstauswahlen, Pflichtgruppen und Mehrfachauswahl-Berechtigungen. |

Dieses Modul bestimmt, welche Modifier-Gruppen für jedes Produkt gelten und ob die resultierende Kundenauswahlkonfiguration logisch gültig ist.

#### 5. Bestellssimulation, POS-Export & Überwachung

| Arbeitsblatt | Zuständigkeit |
|---|---|
| `10_Order_Test` | Simuliert Kundenauswahlen und berechnet Endpreise unter Prüfung der Einhaltung der Auswahlregeln. |
| `11_POS_Export` | Erzeugt eine standardisierte Liste verfügbarer Produkte, Größen, Preise, Modifier-Zusammenfassungen und Synchronisierungsstatus. |
| `12_Dashboard` | Fasst Menügesundheit, Konfigurationsanomalien, Produktverfügbarkeit und POS-Bereitschaft zusammen. |

#### Datenfluss

Der vollständige Datenfluss folgt dieser Sequenz:

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

### Entscheidungsrahmen

Das Toolkit ist um vier betriebliche Entscheidungen herum aufgebaut.

| Entscheidung | Erforderliche Informationen | Entscheidungsergebnis |
| ------------------------------------------ | ---------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------- |
| Was kann verkauft werden? | Kategorie-, Produkt-, Größen- und Verfügbarkeitsdatensätze. | Identifiziert die gültige Menüstruktur und verfügbare Produktkonfigurationen. |
| Was sollen Kunden zahlen? | Basispreis, Modifier-Preisgestaltung, enthaltene Zutaten, Portionsmengen, Platzierung und Steuer. | Bestimmt den berechneten Kundenpreis für eine bestimmte Bestellkonfiguration. |
| Was können Kunden auswählen? | Modifier-Zuweisungen, Gruppenzugehörigkeit, Pflichtstatus und Mindest-/Höchstauswahllimits. | Identifiziert gültige und ungültige Bestellkonfigurationen. |
| Ist das Menü bereit für die POS-Synchronisierung? | Stammdaten-Integrität, Preisvollständigkeit, Modifier-Beziehungen und Validierungsstatus. | Trennt exportbereite Datensätze von korrekturbedürftigen Datensätzen. |

Der Rahmen unterscheidet drei verschiedene Arten von Informationen:

- Produktfakten: Was das Restaurant verkauft und welche Größen verfügbar sind.
- Kommerzielle Regeln: Wie Zutaten, Mengen und Platzierung den Verkaufspreis beeinflussen.
- Betriebliche Constraints: Welche Auswahlen erlaubt sind und ob die Konfiguration bereit für den Einsatz ist.

Diese Trennung verhindert, dass eine Produktbeschreibung, eine Preisannahme und eine Bestelleinschränkung als austauschbare Daten behandelt werden.

### Drei häufige Preis- & Konfigurationsfallen

#### Falle 1 — Doppeltberechnung für bereits enthaltene Zutaten

Die Entscheidung: Den Kundenpreis einer Spezialpizza auf Basis ihrer aufgelisteten Zutaten festlegen.

Der unbemerkte Fehler: Eine herkömmliche Modifier-Berechnung addiert den Standardbelagspreis, unabhängig davon, ob dieser Belag bereits im Spezialprodukt enthalten ist.

Betrachten Sie eine Hawaiian-Pizza mit Schinken und Ananas im Basispreis.

| Kundenauswahl | Falsche Behandlung | Korrekte Behandlung |
| ------------------ | ----------------------------------------------------------- | -------------------------------------------------- |
| Normaler Schinken | Addiert eine Belagsgebühr | Keine zusätzliche Berechnung |
| Light-Ananas | Addiert eine Belagsgebühr | Keine zusätzliche Berechnung |
| Extra-Schinken | Addiert eine Belagsgebühr ohne Berücksichtigung der Enthaltensein | Wendet die konfigurierte Preisregel für Zusatzmengen an |
| Extra-Ananas | Verwendet den Standard-Modifier-Preis ohne Prüfung der Enthaltensein | Wendet die konfigurierte Preisregel für Zusatzmengen an |

Der falsche Ansatz behandelt jede ausgewählte Zutat als neuen Kauf. Der korrigierte Ansatz erkennt, dass eine Zutat bereits Teil der kommerziellen Definition des Produkts sein kann.

Das `07_Default_Modifiers`-Mapping identifiziert enthaltene Zutaten. Die bedingten Preisregeln bestimmen dann den anwendbaren Multiplikator basierend auf Enthalten-Status und gewähltem Portionstyp.

#### Falle 2 — Halb-Pizza-Beläge als unabhängige Festpreise behandeln

Die Entscheidung: Diesen Modifier-Preis berechnen, unabhängig davon, ob ein Belag die gesamte Pizza oder nur eine Hälfte bedeckt.

Der unbemerkte Fehler: Die Produktmenge, die den Belag erhält, ändert sich, aber die Preisberechnung spiegelt die gewählte Platzierung nicht wider.

Nehmen wir zum Beispiel an, ein Belag auf der ganzen Pizza kostet $2.00.

| Platzierung | Falscher Preis | Korrekter Preis |
| --------- | --------------- | ------------- |
| Ganze | $2.00 | $2.00 |
| Linke | $2.00 | $1.00 |
| Rechte | $2.00 | $1.00 |

Der falsche Ansatz berechnet zu viel für Halb-Pizza-Auswahlen, wenn die Preisrichtlinie des Restaurants proportional ist.

Der korrigierte Ansatz leitet Preise für die linke und rechte Hälfte aus der Ganzpreis-Referenz unter Verwendung zentralisierter Preisfaktoren ab.

#### Falle 3 — POS-Menü-Bereitschaft als Dateneingabeaufgabe behandeln

Die Entscheidung: Die neueste Menü-Tabelle nach der Aktualisierung der Produktpreise an den POS-Anbieter senden.

Der unbemerkte Fehler: Eine abgeschlossene Preiseingabe beweist nicht, dass die gesamte Produktkonfiguration gültig ist.

Ein Produkt kann einen Basispreis haben und trotzdem enthalten:

- Einen fehlenden Größendatensatz.
- Eine ungültige Kategoriebeziehung.
- Eine nicht verfügbare Modifier-Gruppe.
- Fehlende Standard-Zutatenzuordnungen.
- Inkonsistente Mindest- und Höchstauswahlregeln.

| Prüfansatz | Risiko | Korrigierter Ansatz |
| ----------------------------------------- | ------------------------------------------------ | ---------------------------------------------------------- |
| Prüfen, ob der Produktname existiert. | Unvollständige Produkte erscheinen bereit. | Produktidentität und Kategoriebeziehungen validieren. |
| Prüfen, ob ein Basispreis existiert. | Einige verfügbare Größen bleiben ohne Preis. | Jede konfigurierte Produkt-Größen-Kombination validieren. |
| Prüfen, ob Modifier-Gruppen aufgeführt sind. | Ungültige oder leere Gruppen können das POS erreichen. | Gruppenzugehörigkeit und Auswahl-Constraints validieren. |
| Das gesamte Menü exportieren. | Fehler werden in die POS-Konfiguration übertragen. | Berechtigte Datensätze exportieren und Synchronisierungsstatus prüfen. |

Der korrigierte Ansatz unterscheidet zwischen abgeschlossener Dateneingabe und betrieblicher Bereitschaft.

### Beispielszenario

Szenario: Testen einer großen Hawaiian-Pizza mit Extra-Ananas auf der linken Hälfte

Ein Restaurant möchte bestätigen, dass seine Spezialpizza-Preisgestaltung korrekt bleibt, wenn Kunden Menge und Platzierung enthaltener Zutaten ändern.

Angenommen, folgende exemplarische Konfiguration:

| Eingabe | Wert |
| -------------------- | -------------- |
| Produkt | Hawaiian-Pizza |
| Größe | Groß |
| Basispreis | $18.00 |
| Enthaltene Zutaten | Schinken, Ananas |
| Gewählter Modifier | Ananas |
| Gewählte Menge | Extra |
| Platzierung | Linke |
| Ganz-Modifier-Preis | $2.00 |
| Faktor linke Hälfte | 0.50 |
| Umsatzsteuer | 8.00% |

Schritt 1 — Die enthaltene Zutat identifizieren.

Ananas ist als enthaltene Zutat für die Hawaiian-Pizza gemappt. Dem Kunden wird daher die standardmäßige, enthaltene Menge nicht erneut berechnet.

Schritt 2 — Den anwendbaren Mengenmultiplikator bestimmen.

Der Kunde wählt Extra statt Normal. Gemäß der exemplarischen Regelmatrix erhält eine enthaltene Zutat, die als Extra gewählt wird, einen Multiplikator von 1.00.

Schritt 3 — Den platzierungsangepassten Modifier-Preis berechnen.

```
Whole Modifier Price = $2.00

Left Half Price = $2.00 × 0.50

Raw Portion Price = $1.00

Applicable Multiplier = 1.00

Additional Charge = $1.00 × 1.00

Additional Charge = $1.00
```

Schritt 4 — Den Kundenpreis berechnen.

| Berechnung | Betrag |
| --------------------------- | ------ |
| Große Hawaiian-Pizza | $18.00 |
| Extra-Ananas — Linke Hälfte | $1.00 |
| Zwischensumme vor Steuern | $19.00 |
| Umsatzsteuer (8%) | $1.52 |
| Endgültiger Kundenpreis | $20.52 |

Analytische Interpretation

Ein herkömmlicher Belagsrechner würde möglicherweise den Standard-Ananaspreis erneut berechnen, selbst wenn die Zutat bereits enthalten ist. Alternativ würde ein Rechner, der die Platzierung ignoriert, die vollen $2.00 für eine Halb-Pizza-Auswahl berechnen.

Das korrekte Ergebnis spiegelt beide kommerziellen Bedingungen wider: den Status der enthaltenen Zutat und die gewählte Platzierung.

Entscheidungsimplikation: Das Restaurant kann den kundensichtbaren Preis validieren, bevor es sein Live-Menü ändert, was das Risiko inkonsistenter Berechnungen und vermeidbarer POS-Korrekturen reduziert.

Dieses Szenario veranschaulicht die konfigurierte Preislogik. Tatsächliche Ergebnisse hängen von der eigenen Preismatrix, den Modifier-Preisen und den Steuereinstellungen des Restaurants ab.


### Formelreferenz

Die Arbeitsmappe verwendet native Excel-Dynamic-Array-Funktionen, um Beziehungen zwischen Menüdatensätzen zu pflegen, Preise zu berechnen und Validierungsausgaben zu erzeugen.

Kernfunktionen umfassen:

| Funktion | Zweck |
|---|---|
| `LET` | Benennt Zwischenberechnungen und hält komplexe Formeln lesbar. |
| `FILTER` | Extrahiert gültige Datensätze und erzeugt dynamisch Ausgabelisten. |
| `XLOOKUP` | Ruft zugehörige Produkt-, Kategorie-, Modifier- und Preisinformationen ab. |
| `COUNTIF` / `COUNTIFS` | Zählt Datensätze und prüft Auswahl- oder Konfigurationsbedingungen. |
| `UNIQUE` | Erzeugt eindeutige Kategorie- und Produktkennungen. |
| `SUM` / `SUMIFS` | Aggregiert Modifier-Gebühren und Geschäftswerte. |
| `IF` / `AND` / `OR` | Wendet Preisbedingungen und Validierungsregeln an. |
| `ROUND` | Steuert die Genauigkeit monetärer Berechnungen. |
| `AVERAGEIFS` | Berechnet durchschnittliche Startpreise auf Kategorieebene. |

#### 1. Globale Parameter — `00_Param_Config`

Das Parameterblatt pflegt zwei Arten von Steuerungen:

**Statische Konfigurationsparameter**

| Parameter | Referenz | Zweck |
|---|---|---|
| Währungssymbol | `B3` | Steuert die monetäre Darstellung. |
| Umsatzsteuersatz | `B4` | Berechnet die endgültigen Kundenpreise. |
| Faktor linke Hälfte | `B5` | Berechnet Modifier-Preise für die linke Hälfte. |
| Faktor rechte Hälfte | `B6` | Berechnet Modifier-Preise für die rechte Hälfte. |
| Faktor ganze Portion | `B7` | Steuert die Preisgestaltung für ganze Portionen. |
| Premium-Modifier-Zuschlag | `B8` | Liefert die konfigurierte Premium-Preisanpassung. |
| Warnschwelle Modifier-Gruppen | `B9` | Markiert übermäßige Modifier-Gruppen-Zuweisungen. |
| Warnschwelle hohe Preise | `B10` | Identifiziert ungewöhnlich hohe Produktpreise. |
| POS-Export-Verfügbarkeitsbaseline | `B11` | Definiert den Status, der für das Filtern beim POS-Export verwendet wird. |
| Standard-Aktivstatus | `B12` | Zentralisiert die Aktivstatus-Kennung. |
| Enthalten-Status | `B13` | Identifiziert standardmäßig enthaltene Zutaten. |
| Standard-Modifier-Stufe | `B14` | Identifiziert die Standard-Modifier-Stufe. |

**Bedingte Preismatrix**

Der Bereich `E3:J10` pflegt die bedingten Preisregeln.

Seine Logik kombiniert:

- Enthalten-Status-Bedingung
- Auswahltyp-Bedingung
- Preismultiplikator-Ergebnis

Ein zusammengesetzter Lookup-Schlüssel wird erzeugt, um die Bedingungen mit dem Preismultiplikator zu verbinden.

Dies verhindert, dass dieselben Preisannahmen unabhängig voneinander in mehrere Arbeitsblätter eingebettet werden.

#### 2. Menü-Stammdaten — `01_Menu_Categories`, `02_Menu_Items`, `03_Item_Sizes`

**Kategoriehierarchie**

`01_Menu_Categories!H3` erzeugt Kategoriepfade:

```excel
=LET(
    cat_name, FILTER(B3:B100, E3:E100<>""),
    sub_name, FILTER(F3:F100, E3:E100<>""),
    cat_name & " > " & sub_name
)
````

Die erzeugten Pfade liefern lesbare Kategoriebeziehungen, z. B.:

```
Pizza > Create Your Own
Pizza > Gourmet Pizza
Subs > Hot Subs
```

Anzahl aktiver Produkte pro Unterkategorie

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

Kategorie-Konfigurationsstatus

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

Produkt-Größen-Kombinationsschlüssel

`03_Item_Sizes` erzeugt einen Kombinationsschlüssel mit:

```
=Item_ID & "|" & Size_ID
```

Zum Beispiel:

```
ITM001|SZ_SM
ITM001|SZ_LG
```

Dieser Schlüssel ermöglicht der Arbeitsmappe, zwischen mehreren Größen desselben Produkts zu unterscheiden.

Produkt-Größen-Preisvalidierung

Die Preisvalidierungsebene prüft, ob ein konfigurierter Größenpreis fehlt, ungültig ist oder die zentrale Warnschwelle überschreitet.

Der resultierende Status wird sowohl vom Produktstamm als auch vom POS-Exportprozess verwendet.

#### 3. Modifier-Stammdaten & Preisgestaltung — `04_Modifier_Groups`, `05_Modifiers`, `06_Modifier_Pricing`

Anzahl aktiver Modifier

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

Auswahlregel-Lookup für Modifier-Gruppen

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

Der entsprechende Höchstauswahl-Lookup verwendet Spalte C in `09_Selection_Rules`.

Ein Rückgabewert von `-1` kennzeichnet eine fehlende Regel.

Modifier-Preisdimensionen

`06_Modifier_Pricing` pflegt die Beziehung zwischen:

```
Modifier ID
    +
Size ID
    +
Selection Type
```

Der zusammengesetzte Preisschlüssel unterstützt das Abrufen des korrekten Modifier-Preises für eine gewählte Produktgröße und Mengenstufe.

Ganz-, Links- und Rechtspreise werden aus den zentralisierten Preisfaktoren abgeleitet.

Die Preisarchitektur trennt den Referenzpreis von der Platzierungsberechnung, sodass restaurant-spezifische Preisrichtlinien angepasst werden können, ohne den Bestellsimulator umzuschreiben.

#### 4. Standardzutaten & Beziehungsmapping — `07_Default_Modifiers`, `08_Group_Assignment`, `09_Selection_Rules`

Standard-Modifier-Lookup-Schlüssel

`07_Default_Modifiers!E3`:

```
=LET(
    itm_ids, FILTER(A3:A1000, A3:A1000<>""),
    mod_ids, FILTER(B3:B1000, A3:A1000<>""),
    itm_ids & "|" & mod_ids
)
```

Produktname-Referenz

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

Modifier-Name-Referenz

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

Validierung der Standardkonfiguration

`07_Default_Modifiers!H3` prüft:

- Ob die Produkt-ID existiert.
- Ob die Modifier-ID existiert.
- Ob dieselbe Produkt-Modifier-Kombination dupliziert ist.
- Ob der Enthalten-Status korrekt konfiguriert ist.

Zuweisungen auf Kategorie- und Artikelebene

`08_Group_Assignment` verwendet das Feld `Level`, um den Zuweisungsumfang zu unterscheiden:

```
Level = Category
    Category ID = valid category
    Item ID = -

Level = Item
    Item ID = valid product
    Category ID = -
```

Die Zuweisungsvalidierungsebene prüft den effektiven Geltungsbereich und identifiziert ungültige oder übermäßige Zuweisungen.

Auswahlregeln

`09_Selection_Rules` verwendet Regeln auf Gruppenebene, um Folgendes zu bestimmen:

- Mindestauswahlen.
- Höchstauswahlen.
- Pflichtstatus.
- Ob Mehrfachauswahlen erlaubt sind.

Diese Werte werden vom Bestellsimulator und den Validierungsformeln für Modifier-Gruppen referenziert.

#### 5. Kundenpreis-Simulation — `10_Order_Test`

Dieses Arbeitsblatt kombiniert den Basispreis des Produkts mit den tatsächlichen Modifier-Gebühren, die aus der gewählten Bestellkonfiguration erzeugt werden.

Name des gewählten Produkts

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

Basisgrößenpreis

`E4`:

```
=XLOOKUP(
    B3 & "|" & B4,
    '03_Item_Sizes'!F3#,
    '03_Item_Sizes'!D3:D1000,
    0
)
```

Gesamter Modifier-Aufpreis

`E5`:

```
=SUM(J10#)
```

Kundenpreis vor Steuern

`E6`:

```
=E4+E5
```

Endgültiger Kundenpreis inklusive Steuern

`E7`:

```
=ROUND(
    E6 * (1 + '00_Param_Config'!$B$4),
    2
)
```

Die Modifier-Detailberechnungen folgen dieser Sequenz:

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

Der Endpreis wird aus der tatsächlichen simulierten Auswahl berechnet und nicht aus einem manuell zusammengestellten Angebot.

#### 6. POS-Export — `11_POS_Export`

Das POS-Export-Arbeitsblatt filtert Produkte dynamisch nach ihrer konfigurierten Verfügbarkeit.

Filterung verfügbarer Produkte

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

Export-Produktname

`B3`:

```
=XLOOKUP(
    A3#,
    '02_Menu_Items'!A3:A500,
    '02_Menu_Items'!B3:B500,
    ""
)
```

Export-Größenname

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

Export-Basispreis

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

POS-Synchronisierungsstatus

Die Exportvalidierung kombiniert Produktintegrität und Größenpreis-Validierung.

Die beabsichtigte Ausgabe ist:

```
READY_FOR_POS
HOLD_CHECK_ERROR
```

Der Export ist so konzipiert, dass Änderungen an der Produktkonfiguration leichter überprüft und übertragen werden können. Es handelt sich nicht um eine direkte POS-API-Integration oder einen automatischen Synchronisierungsdienst.

#### 7. Dashboard & Konfigurationsüberwachung — `12_Dashboard`

Das Dashboard kombiniert sechs betriebliche Kennzahlen:

| KPI | Berechnungszweck |
| ------------------------- | --------------------------------------------- |
| Aktive Kategorien insgesamt | Zählt aktivierte Menükategorien. |
| Verfügbare Artikel insgesamt | Zählt verfügbare Produkte. |
| Wiederverwendbare Modifier insgesamt | Zählt aktivierte Modifier. |
| Preisregeln insgesamt | Zählt konfigurierte Preisregeln. |
| Konfigurationsanomalien insgesamt | Aggregiert erkannte Konfigurationsausnahmen. |
| POS-fertige Artikel | Zählt als exportbereit markierte Datensätze. |

Verfügbare Artikel insgesamt

`C4`:

```
=COUNTIF(
    '02_Menu_Items'!G3:G500,
    '00_Param_Config'!$B$12
)
```

Konfigurationsanomalien insgesamt

`F4`:

```
=COUNTIF('02_Menu_Items'!M3#, "❌*")
+ COUNTIF('03_Item_Sizes'!I3#, "❌*")
+ COUNTIF('04_Modifier_Groups'!I3#, "❌*")
+ COUNTIF('08_Group_Assignment'!I3#, "⚠️*")
+ COUNTIF('09_Selection_Rules'!H3#, "❌*")
```

Anzahl POS-fertiger Datensätze

`G4`:

```
=COUNTIF(
    '11_POS_Export'!I3#,
    "READY_FOR_POS"
)
```

Das Dashboard bietet eine konsolidierte Ansicht des Zustands der Menükonfiguration, während die zugrunde liegenden Arbeitsblätter die detaillierten Aufzeichnungen behalten, die zur Untersuchung und Korrektur einzelner Ausnahmen benötigt werden.

### Validierungsregeln

Das Validierungsframework ist darauf ausgelegt, Fehler zu identifizieren, bevor sie kundensichtbare Preise oder die POS-Konfiguration beeinflussen.

| Feld / Bereich | Validierungsregel | Fehlerverhalten |
| --------------------------- | ---------------------------------------------------------------------------- | ------------------------------------------------------------------- |
| Kategorie-ID | Muss eindeutig sein und auf eine gültige Kategorie verweisen. | Ungültige Kategoriebeziehungen werden markiert. |
| Unterkategorie-ID | Muss einer vorhandenen Unterkategorie entsprechen. | Ungültige Produktklassifizierung wird identifiziert. |
| Kategorie aktiv | Muss die konfigurierten Ja/Nein-Statuswerte verwenden. | Ungültige Statuswerte beeinflussen die Kategorievalidierung. |
| Artikel-ID | Muss eine eindeutige Produktkennung sein. | Ungültige oder doppelte Produktidentitäten erfordern Korrektur. |
| Artikelkategorie | Muss auf eine gültige Kategorie und Unterkategorie verweisen. | Die Produktintegritätsvalidierung identifiziert inkonsistente Beziehungen. |
| Artikelverfügbarkeit | Muss die konfigurierte Aktivstatus-Kennung verwenden. | Ungültige Verfügbarkeit kann korrektes Filtern verhindern. |
| Artikelgröße | Muss auf eine gültige Produkt-Größen-Kombination verweisen. | Fehlende oder ungültige Produkt-Größen-Datensätze werden markiert. |
| Basispreis | Muss vorhanden sein und die konfigurierten Preisprüfungen erfüllen. | Fehlende oder anomale Preise verhindern die POS-Bereitschaft. |
| Modifier-Gruppen-ID | Muss eine eindeutige Gruppenkennung sein. | Ungültige Beziehungen werden markiert. |
| Modifier-ID | Muss eindeutig sein und auf eine gültige Modifier-Gruppe verweisen. | Modifier-Lookup-Fehler werden identifiziert. |
| Modifier aktiv | Muss die konfigurierte Aktivstatus-Kennung verwenden. | Inaktive Modifier werden aus den Zählungen aktiver Modifier ausgeschlossen. |
| Modifier-Preisgestaltung | Muss die erforderlichen Größen- und Auswahltyp-Preisdatensätze enthalten. | Unvollständige Preismatrizen erfordern Korrektur. |
| Modifier-Stufe | Muss den konfigurierten Modifier-Stufen-Kennungen entsprechen. | Ungültige Stufenklassifizierung beeinflusst die Preisvalidierung. |
| Standard-Modifier-Mapping | Produkt- und Modifier-IDs müssen existieren. | Ungültige IDs werden markiert. |
| Standard-Modifier-Eindeutigkeit | Dieselbe Produkt-Modifier-Kombination darf nicht dupliziert werden. | Doppelte Mappings werden markiert. |
| Enthalten-Status | Muss, wo anwendbar, der konfigurierten Enthalten-Status-Kennung entsprechen. | Falsche Konfiguration der standardmäßigen Einbeziehung wird markiert. |
| Zuweisungsebene | Muss entweder `Category` oder `Item` sein. | Ungültiger Zuweisungsumfang wird markiert. |
| Zuweisung auf Kategorieebene | Erfordert eine gültige Kategorie-ID und den vorgeschriebenen Artikel-ID-Platzhalter. | Ungültige Bereichskonfiguration wird markiert. |
| Zuweisung auf Artikelebene | Erfordert eine gültige Artikel-ID und den vorgeschriebenen Kategorie-ID-Platzhalter. | Ungültige Bereichskonfiguration wird markiert. |
| Modifier-Gruppen-Zuweisung | Muss auf eine vorhandene Modifier-Gruppe verweisen. | Ungültige Gruppenbeziehungen werden markiert. |
| Modifier-Gruppen-Kapazität | Die Höchstauswahlzahl darf die Anzahl verfügbarer aktiver Modifier nicht überschreiten. | Die Auswahlkonfiguration erfordert Korrektur. |
| Pflicht-Modifier-Gruppe | Pflichtgruppen müssen eine Mindestauswahl größer null haben. | Logischer Konflikt wird markiert. |
| Optionale Modifier-Gruppe | Optionale Gruppen sollten eine Mindestauswahl von null haben. | Inkonsistente Auswahlregeln werden markiert. |
| Mehrfachauswahl | Muss der konfigurierten Mehrfachauswahl-Berechtigung folgen. | Ungültige Auswahlkonfigurationen erfordern Überprüfung. |
| Kunden-Testprodukt | Muss auf ein gültiges Produkt verweisen. | Ungültige Produktauswahl zeigt eine Warnung an. |
| Kunden-Testgröße | Muss auf eine gültige Größe für das gewählte Produkt verweisen. | Fehlender Größenpreis-Lookup liefert ein ungültiges oder Nullpreis-Ergebnis. |
| Kunden-Modifier-Auswahl | Muss auf einen verfügbaren Modifier verweisen. | Ungültige Modifier-Referenzen erfordern Korrektur. |
| Kunden-Portionstyp | Muss `Light`, `Normal`, `Extra` oder `Xxtra` verwenden. | Ungültige Auswahlwerte sind keine gültigen Preiseingaben. |
| Kunden-Platzierung | Muss `Whole`, `Left` oder `Right` verwenden. | Ungültige Platzierungswerte sind keine gültigen Preiseingaben. |
| Bedingte Preisgestaltung | Muss zur konfigurierten Einbeziehungs- und Mengenregel aufgelöst werden. | Falsche oder fehlende Preiskonfiguration erfordert Untersuchung. |
| Kundenpreis | Muss den Basispreis, Modifier-Gebühren und den konfigurierten Steuersatz widerspiegeln. | Der Bestellsimulator legt die resultierende Berechnung zur Überprüfung offen. |
| POS-Export | Produkt- und Größenpreis-Validierung müssen bestanden werden. | `HOLD_CHECK_ERROR` ersetzt `READY_FOR_POS`. |

Wichtige betriebliche Unterscheidung: Die Validierungsebene prüft die konfigurierten Beziehungen und Geschäftsregeln, die in der Arbeitsmappe dargestellt sind. Sie überprüft nicht unabhängig, ob ein POS-Anbieter jedes exportierte Feld akzeptiert oder jedes Preisverhalten unterstützt.

</details>


## Weitere Tools in dieser Reihe

Weitere schlanke Business-Entscheidungsunterstützungs-Tools und operative Arbeitsmappen:

\- [Theseus Workshop](https://theseusworkshop.com) — Eine Sammlung praktischer Tools für wiederkehrende kleine betriebliche und finanzielle Probleme.

## Lizenz

Dieses Projekt ist unter der Apache License 2.0 lizenziert.

Den vollständigen Lizenztext finden Sie in der LICENSE-Datei.
