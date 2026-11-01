
# Restaurant Menu Master & Pricing Configuration Toolkit

[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-Browser%20%2B%20Excel-2251FF)
![Tool Type](https://img.shields.io/badge/Tool-Menu%20Management%20%26%20Pricing-051C2C)
![Architecture](https://img.shields.io/badge/Architecture-No--VBA-success)

**Manage restaurant menu data, reusable modifiers, complex pricing rules, and POS-ready exports through one structured configuration and price-validation toolkit — free in your browser, with an offline Excel version available.**

> **Start with the free browser version.** No signup, installation, or spreadsheet expertise required.
>
> - 🌐 [Open in Browser](YOUR_HTML_DEMO_URL)
> - 📥 [Get the Excel Version](YOUR_GUMROAD_OR_RELEASE_URL)
>
> The Excel version is designed for permanent records, repeat menu maintenance, and operational audit trails. Purchases are covered by a 30-day, no-questions-asked money-back guarantee.

## What It Helps You Track

- **Menu structure and product availability** — which categories, subcategories, menu items, and product sizes are currently configured and available for sale.
- **Base prices and modifier charges** — how much each product and size costs, including additional charges for standard and premium toppings.
- **Default included ingredients** — which ingredients are already included in specialty products, preventing customers from being charged twice for standard portions.
- **Half-and-whole topping prices** — the correct charge for whole pizzas, left-half toppings, and right-half toppings.
- **Ordering restrictions** — whether customers are selecting the required number of modifiers without exceeding group-level minimum and maximum limits.
- **POS configuration readiness** — which menu items have complete pricing and configuration data, and which still need attention before synchronization.

## Quick Start Workflow

The toolkit follows a simple maintenance cycle: configure the menu once, update individual records when changes occur, validate the results, and export the approved configuration.

1. **Set your restaurant's pricing parameters.** Define the currency symbol, applicable sales tax rate, whole-pizza and half-pizza pricing factors, premium modifier surcharge, and other operational thresholds in the dedicated Parameters sheet. These settings provide a consistent pricing baseline across the workbook.

2. **Enter your menu and modifier data.** Add categories, menu items, available sizes, modifier groups, ingredients, and their prices to the designated input tables. Existing menu information can be transferred from POS exports or other spreadsheets. Use the supplied IDs to establish relationships between records.

3. **Configure business rules and test customer prices.** Define which modifiers belong to each product, which ingredients are included by default, and how many selections customers can make. Use the order simulator to test different product sizes, topping quantities, and placement combinations. Review the calculated pre-tax and final customer prices.

4. **Review configuration health and prepare POS changes.** Check the dashboard for incomplete records and pricing exceptions. Once the relevant records pass validation, use the automatically generated POS export table to prepare a standardized configuration update for the POS provider.

**Set the pricing rules once. Maintain the menu data. Test the customer price. Export only after validation.**

## Why I Built This

Restaurant menu configuration becomes surprisingly difficult when a product is no longer defined by a single name and price.

Consider a specialty pizza with a large base price, several included toppings, optional premium ingredients, and customer-selectable quantities. A conventional menu spreadsheet may record the product price and list its ingredients, but it does not necessarily distinguish between an ingredient already included in the selling price and an additional ingredient that should generate a surcharge.

That creates two different pricing risks.

First, customers can be charged twice for ingredients already included in a specialty product. Second, staff may apply inconsistent prices to extra toppings, different pizza sizes, or half-pizza selections.

For example, a Hawaiian Pizza may already include ham and pineapple. Selecting normal quantities of those ingredients should not generate another charge. Selecting extra pineapple, however, may require an additional charge. A left-half topping should also be priced differently from a whole-pizza topping.

These are not isolated spreadsheet errors. They are failures to represent the actual commercial relationships behind a menu.

I built this toolkit to turn those relationships into a reusable, productized configuration and pricing framework. Instead of maintaining separate ingredient lists and manually checking every product, the same ingredient definitions, pricing rules, and selection constraints can be reused across the menu.

The result is a more controlled workflow: maintain one source of truth, test the resulting customer price, identify configuration errors, and prepare a consistent POS change list.

## Common Restaurant Menu Management Problems This Solves

| Problem | Without This Tool | With This Tool |
|---|---|---|
| **Menu data fragmentation** | Categories, products, sizes, and prices are maintained in disconnected tables, making updates difficult to coordinate. | A structured menu master separates categories, products, and size-specific prices while maintaining their relationships. |
| **Duplicate modifier maintenance** | The same ingredient is recreated for multiple products, resulting in inconsistent names, prices, and availability. | Each modifier is defined once and reused through modifier groups and assignment relationships. |
| **Incorrect specialty product pricing** | Included ingredients may be charged again when customers retain their standard quantities. | Default ingredient mappings distinguish included quantities from additional paid selections. |
| **Inconsistent half-pizza pricing** | Staff manually calculate left-half and right-half topping charges, increasing the likelihood of inconsistent prices. | Centralized pricing factors generate half-portion prices from the whole-portion reference price. |
| **Invalid customer selections** | Required ingredients, minimum selections, and maximum selections are not consistently enforced during menu configuration. | Modifier-group selection rules provide explicit minimum, maximum, and multiple-selection constraints. |
| **POS configuration errors** | Incomplete product records and incorrect prices may be sent to the POS provider without a systematic review. | Automated validation, an order-price simulator, and a POS-ready export help identify problems before synchronization. |

## Who This Is For

Designed for independent restaurants, pizza shops, quick-service restaurants, and multi-product food businesses that need to maintain structured menus with multiple sizes, reusable ingredients, conditional modifier pricing, and customer selection restrictions.

Particularly useful for:

- Restaurant owners maintaining their own menu and pricing configurations.
- Operations managers responsible for menu updates and product consistency.
- Menu administrators preparing product and modifier changes for POS providers.
- Restaurant consultants organizing complex menu structures for clients.

**Who this is not designed for:** Restaurants looking for a complete POS system, payment processor, kitchen display system, inventory management platform, or automated ordering infrastructure. This toolkit focuses on menu master data, pricing configuration, rule validation, and POS preparation. It does not replace the operational POS platform.

No spreadsheet expertise needed to get started. Open the free browser version and begin organizing menu information immediately.

## About

I build lightweight business tools that turn recurring operational problems into structured, reusable decision-support workflows.

Many small businesses already have the necessary information, but it is scattered across spreadsheets, system exports, and manually maintained records. The practical challenge is deciding what information needs to be connected, what rules must be consistent, and what should be checked before an operational decision is made.

The Restaurant Menu Master & Pricing Configuration Toolkit is one example of this approach: a focused tool for maintaining menu relationships and testing pricing outcomes without introducing the complexity of a full enterprise system.


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

## The Business Logic & Methodology

Restaurant menu pricing is not simply a matter of adding ingredient prices to product prices. It requires understanding which ingredients are already included, which selections create additional charges, how portion sizes affect those charges, and which combinations customers are permitted to order.

This toolkit applies several practical business methods to make those commercial relationships explicit.

- Relational master-data management: Separates products, sizes, ingredients, and modifier groups into reusable records. This prevents duplicate maintenance and makes menu-wide changes more consistent.
- Conditional pricing: Distinguishes included ingredients from additional purchases and applies different charges according to the selected quantity. This prevents double charging while preserving revenue from genuine upgrades.
- Proportional portion pricing: Uses whole-product prices as the reference for half-product selections. This creates consistent pricing for left-half, right-half, and whole-pizza toppings without maintaining unrelated prices for each placement.
- Constraint-based configuration: Defines the minimum and maximum number of permitted selections for each modifier group. This helps prevent impossible ordering configurations and reduces manual correction by restaurant staff.
- Pre-deployment quality assurance: Tests representative customer orders, checks configuration exceptions, and separates POS-ready records from incomplete ones. This reduces the chance of transferring incorrect prices or incomplete product relationships into the live ordering environment.

The central commercial objective is straightforward: make the menu's actual pricing and ordering logic consistent before it reaches the customer or the POS system.

## Other Tools in This Series

More lightweight business decision-support tools and operational workbooks:

\- [Theseus Workshop](https://theseusworkshop.com) — A collection of practical tools for recurring small-business operational and financial problems.

## License

This project is licensed under the Apache License 2.0.

See the LICENSE file for the complete license terms.
