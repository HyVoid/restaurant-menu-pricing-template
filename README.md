# Restaurant Menu Pricing Calculator, Modifier Configuration & POS Staging Toolkit (Excel & Web Template)
<img width="1920" height="1032" alt="image" src="https://github.com/user-attachments/assets/45a7c5cb-ec7e-4744-a38b-e705561fe0d6" />


[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-Web%20App%20%2B%20Excel%20Template-2251FF)
![Tool Type](https://img.shields.io/badge/Tool-Restaurant%20Menu%20Management%20%26%20POS%20Pricing-051C2C)
![Architecture](https://img.shields.io/badge/Architecture-Formula--Driven%20No--VBA-success)

**A free online restaurant menu pricing calculator and offline Excel POS configuration template designed for independent restaurants, pizzerias, and QSR operators. Centralize menu master data, calculate complex modifier upcharges, configure half-and-whole pizza topping rules, prevent double-charging on included recipe ingredients, and export validated POS-ready menu sheets without VBA macros.**

> **Test Your Menu Pricing Rules Immediately in Your Browser** — No account registration, software installation, or advanced spreadsheet formulas required.
>
> - 🌐 **[Launch the Free Interactive Restaurant Menu Pricing Web App](YOUR_HTML_DEMO_URL)** *(Test single-session menu configurations, modifier logic, and customer order prices online)*
> - 📥 **[Download the Offline Restaurant POS Menu Configuration Excel Template (.xlsx)](https://theseusworkshop.com/l/dznbyd)** *(Built for permanent menu master records, multi-season price updates, and operational POS audit trails — backed by a 30-day no-questions-asked refund guarantee)*

---

## What This Restaurant Menu Management Template Helps You Track (Pain-Point to Solution Mapping)

| Core Menu Data & Pricing Metric | Common Restaurant Operational Pain Point | How This Menu Configuration Toolkit Solves It |
|---|---|---|
| **Menu Master Hierarchy & Size Availability** | Menu categories, item sizes (Small/Medium/Large), and seasonal availability are scattered across outdated PDFs and POS exports. | Tracks active categories, subcategories, item SKUs, and size-specific availability in one relational menu master database. |
| **Base Menu Prices & Tiered Modifier Upcharges** | Staff guess extra topping charges or charge flat fees that ignore size differences, eroding food profit margins. | Calculates exact base prices plus standard and premium modifier surcharges scaled by product size. |
| **Default Recipe Ingredients (No Double-Charging)** | Customers ordering specialty items (e.g., Hawaiian or Supreme Pizza) get double-charged when keeping standard included toppings. | Maps default included ingredients per item so standard quantities cost $0 extra while "Extra" portions trigger an automatic upcharge. |
| **Fractional & Half-Pizza Topping Pricing** | Manual calculation of left-half, right-half, and whole-pizza toppings leads to checkout delays and inconsistent customer bills. | Applies global fractional pricing multipliers to automatically compute accurate half-portion and whole-portion topping charges. |
| **Modifier Group Min/Max Selection Limits** | Customers select invalid combo options or skip required choices (like crust type or cooking temperature) due to broken POS rules. | Enforces strict minimum, maximum, and per-modifier quantity limits at the modifier-group level before live deployment. |
| **Pre-POS Menu Sync Readiness & QA Status** | Uploading incomplete item records or unverified prices directly into a live POS breaks kitchen tickets and online ordering. | Flags missing prices and broken modifier links on a validation dashboard, isolating clean, POS-ready export rows. |

---

## Step-by-Step Tutorial: How to Configure Restaurant Menu Pricing & Prepare POS Imports

Follow this 4-step menu engineering workflow to standardize your item database, test complex order combinations, and generate a clean POS import sheet:

1. **Step 1: Configure Global Restaurant Tax, Currency & Fractional Pricing Parameters**
   **Action:** Open the `Parameters` tab and input your local sales tax rate, currency symbol, half-topping price multiplier, and premium modifier upcharge factor.
   *Why it matters:* Setting global pricing variables once establishes a unified calculation baseline across every menu category and item size.

2. **Step 2: Import or Enter Your Menu Master List, Sizes & Modifier Groups**
   **Action:** Populate the structured input tables with your menu categories, dish names, size variations, modifier groups, and raw ingredient prices—or paste existing rows from your current POS CSV export.
   *Why it matters:* Linking items and modifiers through clean relational IDs eliminates duplicate data entry when ingredient prices change.

3. **Step 3: Map Default Recipe Inclusions & Simulate Real Customer Orders**
   **Action:** Assign modifier groups to specific menu items, mark built-in specialty ingredients, and run test orders in the interactive **Order Price Simulator** (testing size changes, left/right half toppings, and extra portions).
   *Why it matters:* Real-time pre-tax and post-tax order simulation exposes underpriced add-ons and double-charging bugs before a single customer places an order.

4. **Step 4: Audit Configuration Health & Lock In Long-Term Menu Maintenance with Excel**
   **Action:** Check the automated QA dashboard to clear incomplete item flags, export your validated POS staging table, and **transition from the single-session browser demo to the reusable Excel workbook**.
   *Why it matters:* While the [Free Browser Menu Calculator](YOUR_HTML_DEMO_URL) is ideal for quick one-time testing, browser sessions do not permanently store historical changes. **[Download the Full Offline Excel Menu Master Template (.xlsx)](https://theseusworkshop.com/l/dznbyd)** to save your restaurant's permanent menu database, run recurring quarterly price updates, maintain historical audit logs, and reuse your configuration across multiple store locations without starting from scratch.

**Configure global pricing rules once. Maintain one clean menu master. Simulate real customer checkout prices. Export to your POS only after automated validation.**

---

## Why Standard Restaurant Menu Spreadsheets Fail at Modifier & Topping Pricing

Restaurant menu management becomes deceptively complex the moment a dish is no longer defined by a single flat price.

Consider a specialty pizza or customizable combo meal with size-based pricing, included default toppings, optional premium add-ons, and customer-selectable quantities. A conventional Excel or Google Sheets menu list can record base item prices, but it fails to separate **ingredients already included in the base recipe** from **extra modifiers that require an upcharge**.

This structural blind spot creates two costly commercial risks for food businesses:

1. **The Double-Charging Customer Trap:** Customers are accidentally charged extra for standard ingredients already built into a specialty dish's base price, leading to checkout disputes and abandoned online orders.
2. **Silent Profit Margin Erosion:** Cashiers and POS administrators apply inconsistent prices to extra toppings, size upgrades, or half-and-half pizza selections because pricing rules are calculated by hand.

For example, a **Hawaiian Pizza** includes normal portions of ham and pineapple in its base price. Keeping standard quantities of ham and pineapple must add **$0.00** to the bill. Ordering *extra* pineapple, however, must trigger a size-adjusted surcharge—and placing a premium topping on the **Left Half** only should follow a defined fractional pricing rule rather than guesswork.

These are not simple typing mistakes. They happen because flat spreadsheets cannot model the commercial rules behind a modern restaurant menu.

This toolkit turns those menu relationships into a structured, formula-driven configuration system. Instead of maintaining duplicate topping lists for every item and manually checking prices, you define ingredients, modifier groups, and pricing rules once—and apply them consistently across your entire menu.

---

## Manual Menu Spreadsheets vs. Structured POS Menu Configuration Toolkit

| Restaurant Menu Configuration Challenge | Manual Spreadsheets & Direct POS Edits (Without This Tool) | Structured Menu Master & Pricing Toolkit (With This Tool) |
|---|---|---|
| **Menu Master Data Fragmentation** | Categories, item sizes, and modifier prices live in disconnected tabs or unorganized CSV exports, causing version control chaos during menu updates. | Relational master tables separate categories, base items, and size-specific price matrices while preserving automatic ID lookups. |
| **Duplicate Modifier & Topping Maintenance** | The same cheese, sauce, or topping is typed repeatedly across dozens of menu items, creating mismatched names and outdated upcharge prices. | Each ingredient modifier is created once in a central library and linked to multiple dishes via reusable modifier groups. |
| **Specialty Item & Default Recipe Pricing Bugs** | Standard included toppings on specialty pizzas or burgers are mistakenly billed as paid add-ons when customers customize their order. | Default recipe mapping distinguishes built-in standard quantities ($0 charge) from paid "Extra" upgrades automatically. |
| **Inconsistent Half-Pizza & Fractional Pricing** | Staff manually calculate left-half, right-half, and whole-pizza topping charges, resulting in billing errors and cashier guesswork. | Global fractional pricing factors automatically calculate half-portion topping prices from the whole-pizza reference rate. |
| **Broken POS Modifier Selection Rules** | Minimum required choices and maximum topping caps are forgotten during setup, allowing impossible orders to reach the kitchen display. | Modifier-group constraint rules enforce explicit Min/Max selection limits and per-item quantity caps during staging. |
| **Live POS Deployment & Sync Errors** | Unverified prices and incomplete item rows are entered directly into live POS terminals, causing checkout crashes during peak service hours. | Built-in QA validation flags missing data, simulates customer receipts, and generates a clean, pre-verified POS export table. |

---

## Who This Restaurant Menu & Pricing Template Is For (Target Roles & Use Cases)

Built for independent restaurants, pizzerias, quick-service restaurants (QSR), fast-casual concepts, and multi-unit food operations that manage multi-size menus, shared ingredient modifiers, conditional upcharges, and strict customer ordering rules.

### Role-Based Search Scenarios & Operational Use Cases

- **Pizzeria Owners & Independent Restaurant Operators (`Pizza Menu Pricing Calculator & Topping Spreadsheet`)**
  Stop losing money on unbilled extra toppings or overcharging customers on specialty pizzas. Model whole-pizza vs. half-pizza topping multipliers, size-based surcharges, and included recipe items before updating your physical boards or online ordering menu.
- **QSR & Fast-Casual Operations Managers (`Multi-Size Restaurant Menu Management Template`)**
  Standardize menu hierarchies, combo modifier groups, and seasonal price adjustments across multiple product categories in one centralized offline master file without relying on fragile VBA scripts.
- **POS Onboarding Specialists & Menu Administrators (`Restaurant POS Menu Staging & Import Checklist`)**
  Clean, organize, and validate raw menu data before building or updating databases in Toast, Square, Clover, Lightspeed, or custom POS platforms. Catch missing base prices and broken modifier min/max rules before going live.
- **Restaurant Consultants & Menu Engineers (`Menu Engineering & Modifier Architecture Toolkit`)**
  Deliver a repeatable, audit-ready menu configuration workbook to hospitality clients. Restructure messy client menus into logical modifier groups and demonstrate exact customer checkout pricing during consulting engagements.

**Who this toolkit is NOT designed for:**
Food businesses looking for a live point-of-sale terminal, credit card payment processor, kitchen display system (KDS), raw inventory depletion tracker, or automated recipe nutrition calculator. This toolkit is a **pre-POS menu master, pricing simulator, and configuration QA system**—it prepares and validates your commercial menu logic so your POS platform runs without pricing errors.

---

## The Business Logic & Methodology: Solving Hidden Restaurant Pricing Blind Spots

*(Written for restaurant owners and operators—no technical or advanced spreadsheet background required)*

When customers order food, menu pricing looks simple on the surface: pick an item, add a topping, pay the total. Behind the counter, however, most independent restaurants lose money or frustrate customers because of **three hidden commercial blind spots** that standard spreadsheets fail to catch.

Here is how the business logic inside this toolkit protects your profit margins and customer experience in plain English:

### 1. Eliminating the "One-Update, Ten-Mistakes" Problem (Single Source of Truth)
* **The Blind Spot:** Suppose the wholesale cost of mozzarella cheese or bacon jumps by 20%. In a typical restaurant spreadsheet, "Extra Bacon" is listed separately under Burgers, Sandwiches, Salads, and Pizzas. If you update the price in three places but forget the fourth, you either lose profit margin or confuse customers with inconsistent pricing.
* **How the Toolkit Fixes It:** You enter every ingredient and modifier **once** in a central list. Whether that ingredient is used on 2 items or 50 items, updating its price in one cell automatically updates every connected menu item and order calculation across the entire workbook.

### 2. Preventing the "Specialty Item Double-Charge" Trap (Conditional Recipe Pricing)
* **The Blind Spot:** When a customer orders a "Meat Lover's Pizza" or a "Deluxe Burger," they are already paying a higher base price because pepperoni, sausage, or cheese is included in the recipe. If your menu setup treats every topping the same, your system might accidentally charge the customer *again* just for confirming the standard ingredients—or, if you make all toppings free to avoid complaints, you lose money when someone asks for *double* meat.
* **How the Toolkit Fixes It:** The template separates **"Included by Default"** from **"Paid Extra."** It knows that 1x Pepperoni on a Pepperoni Pizza costs $0.00 extra, while increasing the quantity to 2x (Extra Pepperoni) automatically charges the exact upgrade price. You protect customer trust without giving away free food.

### 3. Stopping Cashier Guesswork on Half-Toppings & Sizes (Proportional Pricing Logic)
* **The Blind Spot:** Customers frequently ask for toppings on only the left half or right half of a pizza, or switch between Small, Medium, and Large sizes. Without a clear formula, staff either charge the full whole-pizza topping price (upsetting the customer) or manually cut prices in half without accounting for preparation labor (hurting your margins).
* **How the Toolkit Fixes It:** You set a single rule—for example, "A half-pizza topping costs 60% of a whole-pizza topping." The toolkit automatically calculates the exact penny-accurate charge for Left-Half, Right-Half, and Whole placements across every pizza size.

### 4. Blocking Impossible Kitchen Tickets (Min/Max Selection Guardrails)
* **The Blind Spot:** If a customer orders a combo meal online and forgets to choose a required side dish—or selects 12 free sauces when only 2 are allowed—your kitchen staff has to stop cooking, call the customer, or absorb the cost.
* **How the Toolkit Fixes It:** You define simple boundary rules for every option group (e.g., *"Must pick at least 1 crust; maximum 1"* or *"Pick up to 4 standard toppings"*). The built-in Order Simulator tests your menu against these rules immediately, showing you if a customer could accidentally place an invalid order.

### 5. Testing Before You Go Live (Pre-POS Quality Assurance)
* **The Blind Spot:** Most restaurants discover menu pricing mistakes *after* they update their POS system—usually when a rush-hour customer complains at the register or an online delivery app charges the wrong total.
* **How the Toolkit Fixes It:** The toolkit acts as a safe "crash test dummy" for your menu. You can simulate real customer orders, see the exact pre-tax and post-tax receipt, and check an automatic health dashboard that highlights missing prices in red. You only export your menu to your POS once every item passes inspection.

---

## About the Creator

I build practical, formula-driven business tools that turn messy operational data into structured, repeatable decision-support workflows.

Most independent restaurants and small businesses already have their pricing data, but it is trapped inside disconnected spreadsheets, messy POS exports, and undocumented staff habits. The real operational challenge is connecting those records, enforcing consistent pricing rules, and catching configuration mistakes before they reach paying customers.

The **Restaurant Menu Pricing Calculator & POS Configuration Toolkit** gives hospitality operators enterprise-grade menu validation inside a transparent, zero-bloat browser and Excel workflow.

- 🌐 **[Try the Free Browser Menu Pricing Calculator Now](YOUR_HTML_DEMO_URL)**
- 📥 **[Get the Full Reusable Excel Menu Master Workbook (.xlsx)](https://theseusworkshop.com/l/dznbyd)**


## Technical Details

<details>
<summary>For technical reviewers, Excel practitioners, and developers — expand to inspect workbook architecture, pricing logic, formulas, and validation rules.</summary>

### Workbook Architecture

The toolkit is designed as a relational menu configuration database and pricing-validation engine rather than a collection of independent menu spreadsheets.

Its architecture follows three principles:

1. **Single Source of Truth:** Each category, product, size, modifier, and pricing rule has a designated master-data location.
2. **Reusable Relationships:** Menu items inherit modifier groups through category-level or item-level assignments instead of duplicating configuration records.
3. **Automated Validation and Output:** Dynamic formulas connect the master data, calculate customer prices, detect configuration errors, and generate standardized POS export records.

The workbook contains five functional modules and 13 worksheets.

#### 1. Global Parameters & Pricing Rules

| Worksheet | Responsibility |
|---|---|
| `00_Param_Config` | Centralizes currency, tax rate, whole/half pricing factors, premium modifier surcharges, operational thresholds, status indicators, and conditional pricing rules. |

This worksheet is the global control layer. Pricing assumptions and business status identifiers are maintained centrally instead of being embedded as fixed values in individual calculation formulas.

#### 2. Menu Product Master Data

| Worksheet | Responsibility |
|---|---|
| `01_Menu_Categories` | Maintains menu categories, subcategories, display order, and category availability. |
| `02_Menu_Items` | Stores individual product identities, descriptions, category relationships, product types, and availability. |
| `03_Item_Sizes` | Stores size-specific product prices and connects each size to its parent menu item. |

The separation between product identity and size-specific pricing prevents the product master from expanding horizontally every time a new size or pricing option is introduced.

#### 3. Reusable Modifiers & Multi-Dimensional Pricing

| Worksheet | Responsibility |
|---|---|
| `04_Modifier_Groups` | Defines customer-facing modifier groups and their operational scope. |
| `05_Modifiers` | Maintains globally reusable ingredients and modifier identities. |
| `06_Modifier_Pricing` | Stores modifier prices by product size and portion level, with whole/left/right placement calculations. |

The same ingredient can be reused across multiple products without creating separate copies of its master record.

#### 4. Relationship Mapping & Selection Constraints

| Worksheet | Responsibility |
|---|---|
| `07_Default_Modifiers` | Maps specialty products to their included ingredients and default quantities. |
| `08_Group_Assignment` | Assigns modifier groups at either category level or individual product level. |
| `09_Selection_Rules` | Controls minimum selections, maximum selections, required groups, and multiple-selection permissions. |

This module determines which modifier groups apply to each product and whether the resulting customer selection configuration is logically valid.

#### 5. Order Simulation, POS Export & Monitoring

| Worksheet | Responsibility |
|---|---|
| `10_Order_Test` | Simulates customer selections and calculates final prices while checking selection-rule compliance. |
| `11_POS_Export` | Generates a standardized list of available products, sizes, prices, modifier summaries, and synchronization status. |
| `12_Dashboard` | Summarizes menu health, configuration anomalies, product availability, and POS readiness. |

#### Data Flow

The complete data flow follows this sequence:

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

### Decision Framework

The toolkit is built around four operational decisions.

| Decision                                   | Information Required                                                                           | Decision Outcome                                                             |
| ------------------------------------------ | ---------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------- |
| What can be sold?                          | Category, product, size, and availability records.                                             | Identifies the valid menu structure and available product configurations.    |
| What should customers pay?                 | Base price, modifier pricing, included ingredients, portion quantities, placement, and tax.    | Determines the calculated customer price for a specific order configuration. |
| What can customers select?                 | Modifier assignments, group membership, required status, and minimum/maximum selection limits. | Identifies valid and invalid ordering configurations.                        |
| Is the menu ready for POS synchronization? | Master-data integrity, pricing completeness, modifier relationships, and validation status.    | Separates records ready for export from records requiring correction.        |

The framework distinguishes three different types of information:

- Product facts: What the restaurant sells and which sizes are available.
- Commercial rules: How ingredients, quantities, and placement affect the selling price.
- Operational constraints: Which selections are permitted and whether the configuration is ready for deployment.

This separation prevents a product description, a pricing assumption, and an ordering restriction from being treated as interchangeable data.

### Three Common Pricing & Configuration Traps

#### Trap 1 — Charging Twice for Ingredients Already Included

The decision: Set a specialty pizza's customer price based on its listed ingredients.

The unnoticed flaw: A conventional modifier calculation adds the standard topping price regardless of whether that topping is already included in the specialty product.

Consider a Hawaiian Pizza with ham and pineapple included in its base price.

| Customer Selection | Incorrect Treatment                                         | Correct Treatment                                  |
| ------------------ | ----------------------------------------------------------- | -------------------------------------------------- |
| Normal ham         | Adds a topping charge                                       | No additional charge                               |
| Light pineapple    | Adds a topping charge                                       | No additional charge                               |
| Extra ham          | Adds a topping charge without considering inclusion         | Applies the configured extra-quantity pricing rule |
| Extra pineapple    | Uses the standard modifier price without checking inclusion | Applies the configured extra-quantity pricing rule |

The incorrect approach treats every selected ingredient as a new purchase. The corrected approach recognizes that an ingredient may already be part of the product's commercial definition.

The `07_Default_Modifiers` mapping identifies included ingredients. The conditional pricing rules then determine the applicable multiplier based on inclusion status and selected portion type.

#### Trap 2 — Treating Half-Pizza Toppings as Independent Fixed Prices

The decision: Charge the same modifier price regardless of whether a topping covers the entire pizza or only one half.

The unnoticed flaw: The amount of product receiving the topping changes, but the pricing calculation does not reflect the selected placement.

For example, assume a whole-pizza topping costs $2.00.

| Placement | Incorrect Price | Correct Price |
| --------- | --------------- | ------------- |
| Whole     | $2.00           | $2.00         |
| Left      | $2.00           | $1.00         |
| Right     | $2.00           | $1.00         |

The incorrect approach overcharges half-pizza selections when the restaurant's policy is proportional pricing.

The corrected approach derives left- and right-half prices from the whole-price reference using centralized pricing factors.

#### Trap 3 — Treating POS Menu Readiness as a Data Entry Task

The decision: Send the latest menu spreadsheet to the POS provider after updating product prices.

The unnoticed flaw: A completed price entry does not establish that the entire product configuration is valid.

A product may have a base price but still contain:

- A missing size record.
- An invalid category relationship.
- An unavailable modifier group.
- Missing default ingredient mappings.
- Inconsistent minimum and maximum selection rules.

| Review Approach                           | Risk                                             | Corrected Approach                                         |
| ----------------------------------------- | ------------------------------------------------ | ---------------------------------------------------------- |
| Check whether the product name exists.    | Incomplete products appear ready.                | Validate product identity and category relationships.      |
| Check whether a base price exists.        | Some available sizes remain unpriced.            | Validate each configured product-size combination.         |
| Check whether modifier groups are listed. | Invalid or empty groups may reach the POS.       | Validate group membership and selection constraints.       |
| Export the entire menu.                   | Errors are transferred to the POS configuration. | Export eligible records and review synchronization status. |

The corrected approach distinguishes data entry completion from operational readiness.

### Example Scenario

Scenario: Testing a Large Hawaiian Pizza with Extra Pineapple on the Left Half

A restaurant wants to confirm that its specialty pizza pricing remains correct when customers change the quantity and placement of included ingredients.

Assume the following illustrative configuration:

| Input                | Value          |
| -------------------- | -------------- |
| Product              | Hawaiian Pizza |
| Size                 | Large          |
| Base Price           | $18.00         |
| Included Ingredients | Ham, Pineapple |
| Selected Modifier    | Pineapple      |
| Selected Quantity    | Extra          |
| Placement            | Left           |
| Whole Modifier Price | $2.00          |
| Left Half Factor     | 0.50           |
| Sales Tax            | 8.00%          |

Step 1 — Identify the included ingredient.

Pineapple is mapped as an included ingredient for Hawaiian Pizza. The customer is therefore not charged for the standard included quantity again.

Step 2 — Determine the applicable quantity multiplier.

The customer selects Extra rather than Normal. Under the illustrative rule matrix, an included ingredient selected as Extra receives a multiplier of 1.00.

Step 3 — Calculate the placement-adjusted modifier price.

```
Whole Modifier Price = $2.00

Left Half Price = $2.00 × 0.50

Raw Portion Price = $1.00

Applicable Multiplier = 1.00

Additional Charge = $1.00 × 1.00

Additional Charge = $1.00
```

Step 4 — Calculate the customer price.

| Calculation                 | Amount |
| --------------------------- | ------ |
| Large Hawaiian Pizza        | $18.00 |
| Extra Pineapple — Left Half | $1.00  |
| Pre-Tax Total               | $19.00 |
| Sales Tax (8%)              | $1.52  |
| Final Customer Price        | $20.52 |

Analytical interpretation

A conventional topping calculator might charge the standard pineapple price again, even when the ingredient is already included. Alternatively, a calculator that ignores placement might charge the full $2.00 for a half-pizza selection.

The correct result reflects both commercial conditions: the included ingredient status and the selected placement.

Decision implication: The restaurant can validate the customer-facing price before changing its live menu, reducing the risk of inconsistent charges and avoidable POS corrections.

This scenario illustrates the configured pricing logic. Actual results depend on the restaurant's own pricing matrix, modifier prices, and tax settings.

````

```markdown
### Formula Reference

The workbook uses native Excel dynamic-array functions to maintain relationships between menu records, calculate prices, and generate validation outputs.

Core functions include:

| Function | Purpose |
|---|---|
| `LET` | Names intermediate calculations and keeps complex formulas readable. |
| `FILTER` | Extracts valid records and dynamically generates output lists. |
| `XLOOKUP` | Retrieves related product, category, modifier, and pricing information. |
| `COUNTIF` / `COUNTIFS` | Counts records and checks selection or configuration conditions. |
| `UNIQUE` | Generates distinct category and product identifiers. |
| `SUM` / `SUMIFS` | Aggregates modifier charges and business values. |
| `IF` / `AND` / `OR` | Applies pricing conditions and validation rules. |
| `ROUND` | Controls monetary calculation precision. |
| `AVERAGEIFS` | Calculates category-level average starting prices. |

#### 1. Global Parameters — `00_Param_Config`

The parameter sheet maintains two types of controls:

**Static configuration parameters**

| Parameter | Reference | Purpose |
|---|---|---|
| Currency Symbol | `B3` | Controls monetary presentation. |
| Sales Tax Rate | `B4` | Calculates final customer prices. |
| Left Half Factor | `B5` | Calculates left-half modifier prices. |
| Right Half Factor | `B6` | Calculates right-half modifier prices. |
| Whole Factor | `B7` | Controls whole-portion pricing. |
| Premium Modifier Surcharge | `B8` | Provides the configured premium pricing adjustment. |
| Modifier Group Warning Threshold | `B9` | Flags excessive modifier-group assignments. |
| High Price Warning Threshold | `B10` | Identifies unusually high product prices. |
| POS Export Availability Baseline | `B11` | Defines the status used for POS export filtering. |
| Default Active Status | `B12` | Centralizes the active-status identifier. |
| Included Status | `B13` | Identifies default included ingredients. |
| Standard Modifier Tier | `B14` | Identifies the standard modifier tier. |

**Conditional pricing matrix**

The `E3:J10` area maintains the conditional pricing rules.

Its logic combines:

- Included Status Condition
- Selection Type Condition
- Pricing Multiplier Result

A composite lookup key is generated to connect the conditions with the pricing multiplier.

This prevents the same pricing assumptions from being independently embedded in multiple worksheets.

#### 2. Menu Master Data — `01_Menu_Categories`, `02_Menu_Items`, `03_Item_Sizes`

**Category hierarchy**

`01_Menu_Categories!H3` generates category paths:

```excel
=LET(
    cat_name, FILTER(B3:B100, E3:E100<>""),
    sub_name, FILTER(F3:F100, E3:E100<>""),
    cat_name & " > " & sub_name
)
````

The resulting paths provide readable category relationships, such as:

```
Pizza > Create Your Own
Pizza > Gourmet Pizza
Subs > Hot Subs
```

Active product count by subcategory

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

Category configuration status

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

Product-size composite key

`03_Item_Sizes` generates a composite key using:

```
=Item_ID & "|" & Size_ID
```

For example:

```
ITM001|SZ_SM
ITM001|SZ_LG
```

This key allows the workbook to distinguish between multiple sizes belonging to the same product.

Product-size price validation

The price validation layer checks whether a configured size price is missing, invalid, or exceeds the centralized warning threshold.

The resulting status is used by both the product master and POS export process.

#### 3. Modifier Master Data & Pricing — `04_Modifier_Groups`, `05_Modifiers`, `06_Modifier_Pricing`

Active modifier count

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

Modifier group selection-rule lookup

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

The corresponding maximum-selection lookup uses column C in `09_Selection_Rules`.

A return value of `-1` identifies a missing rule.

Modifier pricing dimensions

`06_Modifier_Pricing` maintains the relationship between:

```
Modifier ID
    +
Size ID
    +
Selection Type
```

The composite pricing key supports the retrieval of the correct modifier price for a selected product size and quantity level.

Whole, left, and right prices are derived from the centralized pricing factors.

The pricing architecture separates the reference price from the placement calculation, allowing restaurant-specific pricing policies to be adjusted without rewriting the order simulator.

#### 4. Default Ingredients & Relationship Mapping — `07_Default_Modifiers`, `08_Group_Assignment`, `09_Selection_Rules`

Default modifier lookup key

`07_Default_Modifiers!E3`:

```
=LET(
    itm_ids, FILTER(A3:A1000, A3:A1000<>""),
    mod_ids, FILTER(B3:B1000, A3:A1000<>""),
    itm_ids & "|" & mod_ids
)
```

Product name reference

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

Modifier name reference

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

Default configuration validation

`07_Default_Modifiers!H3` checks:

- Whether the product ID exists.
- Whether the modifier ID exists.
- Whether the same product-modifier combination is duplicated.
- Whether the included status is correctly configured.

Category-level and item-level assignments

`08_Group_Assignment` uses the `Level` field to distinguish assignment scope:

```
Level = Category
    Category ID = valid category
    Item ID = -

Level = Item
    Item ID = valid product
    Category ID = -
```

The assignment validation layer checks the effective scope and identifies invalid or excessive assignments.

Selection rules

`09_Selection_Rules` uses group-level rules to determine:

- Minimum selections.
- Maximum selections.
- Required status.
- Whether multiple selections are permitted.

These values are referenced by the order simulator and modifier group validation formulas.

#### 5. Customer Price Simulation — `10_Order_Test`

This worksheet combines the product's base price with the actual modifier charges generated by the selected order configuration.

Selected product name

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

Base size price

`E4`:

```
=XLOOKUP(
    B3 & "|" & B4,
    '03_Item_Sizes'!F3#,
    '03_Item_Sizes'!D3:D1000,
    0
)
```

Total modifier upcharge

`E5`:

```
=SUM(J10#)
```

Pre-tax customer price

`E6`:

```
=E4+E5
```

Final customer price including tax

`E7`:

```
=ROUND(
    E6 * (1 + '00_Param_Config'!$B$4),
    2
)
```

The modifier detail calculations use the following sequence:

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

The final price is calculated from the actual simulated selection rather than from a manually assembled quotation.

#### 6. POS Export — `11_POS_Export`

The POS export worksheet dynamically filters products based on their configured availability.

Available product filtering

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

Export product name

`B3`:

```
=XLOOKUP(
    A3#,
    '02_Menu_Items'!A3:A500,
    '02_Menu_Items'!B3:B500,
    ""
)
```

Export size name

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

Export base price

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

POS synchronization status

The export validation combines product integrity and size-price validation.

The intended output is:

```
READY_FOR_POS
HOLD_CHECK_ERROR
```

The export is designed to make product configuration changes easier to review and transfer. It is not a direct POS API integration or automatic synchronization service.

#### 7. Dashboard & Configuration Monitoring — `12_Dashboard`

The dashboard combines six operational indicators:

| KPI                       | Calculation Purpose                           |
| ------------------------- | --------------------------------------------- |
| Total Active Categories   | Counts enabled menu categories.               |
| Total Available Items     | Counts available products.                    |
| Total Reusable Modifiers  | Counts enabled modifiers.                     |
| Total Pricing Rules Count | Counts configured pricing rules.              |
| Total Config Anomalies    | Aggregates detected configuration exceptions. |
| POS Ready Item Count      | Counts records marked ready for export.       |

Total available items

`C4`:

```
=COUNTIF(
    '02_Menu_Items'!G3:G500,
    '00_Param_Config'!$B$12
)
```

Total configuration anomalies

`F4`:

```
=COUNTIF('02_Menu_Items'!M3#, "❌*")
+ COUNTIF('03_Item_Sizes'!I3#, "❌*")
+ COUNTIF('04_Modifier_Groups'!I3#, "❌*")
+ COUNTIF('08_Group_Assignment'!I3#, "⚠️*")
+ COUNTIF('09_Selection_Rules'!H3#, "❌*")
```

POS-ready record count

`G4`:

```
=COUNTIF(
    '11_POS_Export'!I3#,
    "READY_FOR_POS"
)
```

The dashboard provides a consolidated view of menu configuration health, while the underlying worksheets retain the detailed records needed to investigate and correct individual exceptions.

### Validation Rules

The validation framework is designed to identify errors before they affect customer-facing prices or POS configuration.

| Field / Area                | Validation Rule                                                              | Error Behavior                                                      |
| --------------------------- | ---------------------------------------------------------------------------- | ------------------------------------------------------------------- |
| Category ID                 | Must be unique and reference a valid category.                               | Invalid category relationships are flagged.                         |
| Subcategory ID              | Must correspond to an existing subcategory.                                  | Invalid product classification is identified.                       |
| Category Active             | Must use the configured Yes/No status values.                                | Invalid status values affect category validation.                   |
| Item ID                     | Must be a unique product identifier.                                         | Invalid or duplicate product identities require correction.         |
| Item Category               | Must reference a valid category and subcategory.                             | Product integrity validation identifies inconsistent relationships. |
| Item Availability           | Must use the configured active-status identifier.                            | Invalid availability can prevent correct filtering.                 |
| Item Size                   | Must reference a valid product and size combination.                         | Missing or invalid product-size records are flagged.                |
| Base Price                  | Must be present and satisfy the configured pricing checks.                   | Missing or anomalous prices prevent POS readiness.                  |
| Modifier Group ID           | Must be a unique group identifier.                                           | Invalid relationships are flagged.                                  |
| Modifier ID                 | Must be unique and reference a valid modifier group.                         | Modifier lookup errors are identified.                              |
| Modifier Active             | Must use the configured active-status identifier.                            | Inactive modifiers are excluded from active modifier counts.        |
| Modifier Pricing            | Must contain the required size and selection-type pricing records.           | Incomplete pricing matrices require correction.                     |
| Modifier Tier               | Must correspond to the configured modifier-tier identifiers.                 | Invalid tier classification affects pricing validation.             |
| Default Modifier Mapping    | Product and modifier IDs must exist.                                         | Invalid IDs are flagged.                                            |
| Default Modifier Uniqueness | The same product-modifier combination must not be duplicated.                | Duplicate mappings are flagged.                                     |
| Included Status             | Must match the configured included-status identifier where applicable.       | Incorrect default inclusion configuration is flagged.               |
| Assignment Level            | Must be either `Category` or `Item`.                                         | Invalid assignment scope is flagged.                                |
| Category-Level Assignment   | Requires a valid Category ID and the prescribed Item ID placeholder.         | Invalid scope configuration is flagged.                             |
| Item-Level Assignment       | Requires a valid Item ID and the prescribed Category ID placeholder.         | Invalid scope configuration is flagged.                             |
| Modifier Group Assignment   | Must reference an existing modifier group.                                   | Invalid group relationships are flagged.                            |
| Modifier Group Capacity     | Maximum selection count must not exceed the available active modifier count. | Selection configuration requires correction.                        |
| Required Modifier Group     | Required groups must have a minimum selection greater than zero.             | Logical conflict is flagged.                                        |
| Optional Modifier Group     | Optional groups should have a minimum selection of zero.                     | Inconsistent selection rules are flagged.                           |
| Multiple Selection          | Must follow the configured multiple-selection permission.                    | Invalid selection configurations require review.                    |
| Customer Test Product       | Must reference a valid product.                                              | Invalid product selection displays a warning.                       |
| Customer Test Size          | Must reference a valid size for the selected product.                        | Missing size-price lookup returns an invalid or zero-price result.  |
| Customer Modifier Selection | Must reference an available modifier.                                        | Invalid modifier references require correction.                     |
| Customer Portion Type       | Must use `Light`, `Normal`, `Extra`, or `Xxtra`.                             | Invalid selection values are not valid pricing inputs.              |
| Customer Placement          | Must use `Whole`, `Left`, or `Right`.                                        | Invalid placement values are not valid pricing inputs.              |
| Conditional Pricing         | Must resolve to the configured inclusion and quantity rule.                  | Incorrect or missing pricing configuration requires investigation.  |
| Customer Price              | Must reflect the base price, modifier charges, and configured tax rate.      | The order simulator exposes the resulting calculation for review.   |
| POS Export                  | Product and size-price validation must pass.                                 | `HOLD_CHECK_ERROR` replaces `READY_FOR_POS`.                        |

Important operational distinction: The validation layer checks the configured relationships and business rules represented in the workbook. It does not independently verify whether a POS provider accepts every exported field or supports every pricing behavior.

</details>


## Other Tools in This Series

More lightweight business decision-support tools and operational workbooks:

\- [Theseus Workshop](https://theseusworkshop.com) — A collection of practical tools for recurring small-business operational and financial problems.

## License

This project is licensed under the Apache License 2.0.

See the LICENSE file for the complete license terms.
