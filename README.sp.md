[ 🌐 عربي ](README.ar.md) | [ 🇩🇪 Deutsch ](README.de.md) | [ 🇫🇷 Français ](README.fr.md) | [ 🇳🇱 Nederlands ](README.nl.md) | [ 🇪🇸 Español ](README.sp.md) | [ 🇬🇧 English ](README.md)

# Calculadora de Precios de Menús de Restaurante, Configuración de Modificadores y Kit de Herramientas de Preparación POS (Plantilla Excel y Web)
<img width="1920" height="1032" alt="image" src="https://github.com/user-attachments/assets/45a7c5cb-ec7e-4744-a38b-e705561fe0d6" />


[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-Web%20App%20%2B%20Excel%20Template-2251FF)
![Tool Type](https://img.shields.io/badge/Tool-Restaurant%20Menu%20Management%20%26%20POS%20Pricing-051C2C)
![Architecture](https://img.shields.io/badge/Architecture-Formula--Driven%20No--VBA-success)

**Una calculadora de precios de menús de restaurantes en línea y gratuita, junto con una plantilla Excel offline de configuración POS, diseñada para restaurantes independientes, pizzerías y operadores QSR. Centraliza los datos maestros del menú, calcula cargos adicionales complejos de modificadores, configura reglas de toppings de media pizza y pizza entera, evita el cobro doble de ingredientes ya incluidos en la receta y exporta hojas de menú validadas y listas para POS sin macros VBA.**

> **Pruebe de inmediato sus reglas de precios de menú en su navegador** — No se requiere registro de cuenta, instalación de software ni fórmulas avanzadas de hojas de cálculo.
>
> - 🌐 **[Lance la aplicación web gratuita e interactiva de precios de menús de restaurantes](https://hyvoid.github.io/restaurant-menu-pricing-template/)** *(Pruebe en línea configuraciones de menú de sesión única, lógica de modificadores y precios de pedidos de clientes)*
> - 📥 **[Descargue la plantilla Excel offline de configuración de menús POS para restaurantes (.xlsx)](https://theseusworkshop.com/l/dznbyd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=restaurant-menu-pricing-template)** *(Diseñada para registros maestros de menú permanentes, actualizaciones de precios multi-temporada y trazas de auditoría operativa de POS — respaldada por una garantía de reembolso de 30 días sin preguntas)*

---

## Qué le ayuda a controlar esta plantilla de gestión de menús de restaurante (mapeo de puntos de dolor a soluciones)

| Métrica central de datos y precios del menú | Punto de dolor operativo común en restaurantes | Cómo lo resuelve este kit de herramientas de configuración de menús |
|---|---|---|
| **Jerarquía maestra del menú y disponibilidad de tamaños** | Las categorías del menú, los tamaños de los artículos (Pequeño/Medio/Grande) y la disponibilidad de temporada están dispersos en PDF obsoletos y exportaciones de POS. | Realiza seguimiento de categorías activas, subcategorías, SKU de artículos y disponibilidad por tamaño en una base de datos relacional única de datos maestros del menú. |
| **Precios base del menú y cargos adicionales escalonados de modificadores** | El personal adivina los cargos por toppings extra o aplica tarifas fijas que ignoran las diferencias de tamaño, erosionando los márgenes de beneficio alimentario. | Calcula precios base exactos más recargos estándar y premium de modificadores escalados según el tamaño del producto. |
| **Ingredientes predeterminados de la receta (sin cobro doble)** | Los clientes que piden artículos especiales (p. ej., pizza Hawaiana o Supreme) sufren cobro doble cuando conservan los toppings estándar ya incluidos. | Mapea los ingredientes incluidos predeterminados por artículo, de modo que las cantidades estándar cuesten $0 adicionales, mientras que las porciones "Extra" activan un cargo adicional automático. |
| **Precios fraccionales y de media pizza para toppings** | El cálculo manual de toppings de mitad izquierda, mitad derecha y pizza entera provoca demoras en el pago y facturas inconsistentes para los clientes. | Aplica multiplicadores globales de precios fraccionales para calcular automáticamente cargos precisos de toppings por media porción y porción completa. |
| **Límites mín/máx de selección de grupos de modificadores** | Los clientes seleccionan opciones de combo no válidas u omiten elecciones obligatorias (como el tipo de masa o el punto de cocción) debido a reglas de POS defectuosas. | Aplica límites estrictos de mínimo, máximo y cantidad por modificador a nivel de grupo de modificadores antes del despliegue en vivo. |
| **Preparación y estado de QA para la sincronización previa al POS** | Cargar registros de artículos incompletos o precios no verificados directamente en un POS en vivo rompe los tickets de cocina y los pedidos en línea. | Señala precios faltantes y enlaces de modificadores rotos en un panel de validación, aislando filas limpias y listas para exportar al POS. |

---

## Tutorial paso a paso: cómo configurar los precios del menú de un restaurante y preparar las importaciones al POS

Siga este flujo de trabajo de ingeniería de menús de 4 pasos para estandarizar su base de datos de artículos, probar combinaciones complejas de pedidos y generar una hoja de importación al POS limpia:

1. **Paso 1: Configurar los parámetros globales de impuestos, moneda y precios fraccionales del restaurante**
   **Acción:** Abra la pestaña `Parameters` e introduzca su tasa de impuestos locales, el símbolo de moneda, el multiplicador de precio de topping de media pizza y el factor de cargo adicional de modificadores premium.
   *Por qué importa:* Establecer las variables globales de precios una sola vez crea una base de cálculo unificada para todas las categorías del menú y tamaños de artículos.

2. **Paso 2: Importar o introducir su lista maestra de menú, tamaños y grupos de modificadores**
   **Acción:** Rellene las tablas de entrada estructuradas con sus categorías de menú, nombres de platos, variantes de tamaño, grupos de modificadores y precios base de ingredientes, o pegue filas existentes de su exportación CSV actual del POS.
   *Por qué importa:* Vincular artículos y modificadores mediante IDs relacionales limpios elimina la entrada duplicada de datos cuando cambian los precios de los ingredientes.

3. **Paso 3: Mapear los ingredientes incluidos predeterminados de la receta y simular pedidos reales de clientes**
   **Acción:** Asigne grupos de modificadores a artículos específicos del menú, marque los ingredientes especiales integrados y ejecute pedidos de prueba en el **Simulador de Precio de Pedidos** interactivo (probando cambios de tamaño, toppings de mitad izquierda/derecha y porciones extra).
   *Por qué importa:* La simulación de pedidos en tiempo real, antes y después de impuestos, expone complementos con precios bajos y errores de cobro doble antes de que un solo cliente realice un pedido.

4. **Paso 4: Auditar el estado de la configuración y consolidar el mantenimiento del menú a largo plazo con Excel**
   **Acción:** Revise el panel de QA automatizado para eliminar las marcas de artículos incompletos, exporte su tabla de preparación POS validada y **pase de la demostración de navegador de sesión única al libro de Excel reutilizable**.
   *Por qué importa:* Si bien la [calculadora de menús gratuita en el navegador](https://hyvoid.github.io/restaurant-menu-pricing-template/) es ideal para pruebas rápidas puntuales, las sesiones del navegador no almacenan de forma permanente los cambios históricos. **[Descargue la plantilla Excel completa y offline de datos maestros del menú (.xlsx)](https://theseusworkshop.com/l/dznbyd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=restaurant-menu-pricing-template)** para guardar la base de datos de menú permanente de su restaurante, ejecutar actualizaciones trimestrales de precios recurrentes, mantener registros históricos de auditoría y reutilizar su configuración en varias sucursales sin empezar desde cero.

**Configure las reglas globales de precios una sola vez. Mantenga unos datos maestros del menú limpios y únicos. Simule precios reales de pago de clientes. Exporte a su POS solo después de la validación automatizada.**

---

## Por qué fallan las hojas de cálculo de menús estándar de restaurantes en los precios de modificadores y toppings

La gestión del menú de un restaurante se vuelve engañosamente compleja en el momento en que un plato deja de definirse por un único precio fijo.

Considere una pizza especial o un combo personalizable con precios según el tamaño, toppings predeterminados incluidos, complementos premium opcionales y cantidades seleccionables por el cliente. Una lista de menú convencional en Excel o Google Sheets puede registrar los precios base de los artículos, pero no logra separar **los ingredientes ya incluidos en la receta base** de **los modificadores extra que requieren un cargo adicional**.

Este punto ciego estructural crea dos riesgos comerciales costosos para los negocios de comida:

1. **La trampa del cobro doble al cliente:** Los clientes reciben cargos accidentales por ingredientes estándar ya incluidos en el precio base de un plato especial, lo que genera disputas en el pago y pedidos en línea abandonados.
2. **Erosión silenciosa del margen de beneficio:** Los cajeros y administradores de POS aplican precios inconsistentes a los toppings extra, las mejoras de tamaño o las selecciones de media y mitad de pizza porque las reglas de precios se calculan a mano.

Por ejemplo, una **pizza Hawaiana** incluye porciones normales de jamón y piña en su precio base. Mantener las cantidades estándar de jamón y piña debe añadir **$0.00** a la cuenta. Sin embargo, pedir piña *extra* debe activar un recargo ajustado por tamaño, y colocar un topping premium solo en la **Mitad Izquierda** debe seguir una regla de precios fraccional definida en lugar de conjeturas.

No se trata de simples errores de tecleo. Ocurren porque las hojas de cálculo planas no pueden modelar las reglas comerciales detrás de un menú de restaurante moderno.

Este kit de herramientas convierte esas relaciones del menú en un sistema de configuración estructurado y basado en fórmulas. En lugar de mantener listas de toppings duplicadas para cada artículo y verificar precios manualmente, usted define ingredientes, grupos de modificadores y reglas de precios una sola vez, y los aplica de manera consistente en todo su menú.

---

## Hojas de cálculo de menús manuales frente a kit de herramientas estructurado de configuración de menús POS

| Desafío de configuración del menú del restaurante | Hojas de cálculo manuales y ediciones directas en POS (sin esta herramienta) | Datos maestros estructurados del menú y kit de herramientas de precios (con esta herramienta) |
|---|---|---|
| **Fragmentación de los datos maestros del menú** | Las categorías, los tamaños de artículos y los precios de modificadores viven en pestañas desconectadas o exportaciones CSV desordenadas, causando caos de control de versiones durante las actualizaciones del menú. | Las tablas maestras relacionales separan categorías, artículos base y matrices de precios por tamaño, preservando las búsquedas automáticas de IDs. |
| **Mantenimiento duplicado de modificadores y toppings** | El mismo queso, salsa o topping se teclea repetidamente en docenas de artículos del menú, creando nombres inconsistentes y precios de cargo adicional desactualizados. | Cada modificador de ingrediente se crea una vez en una biblioteca central y se vincula a varios platos mediante grupos de modificadores reutilizables. |
| **Errores de precios en artículos especiales y recetas predeterminadas** | Los toppings estándar incluidos en pizzas especiales o hamburguesas se facturan por error como complementos de pago cuando los clientes personalizan su pedido. | El mapeo de recetas predeterminadas distingue automáticamente las cantidades estándar integradas (cargo de $0) de las mejoras de pago "Extra". |
| **Precios inconsistentes de media pizza y fraccionales** | El personal calcula manualmente los cargos de toppings de mitad izquierda, mitad derecha y pizza entera, lo que genera errores de facturación y conjeturas del cajero. | Los factores globales de precios fraccionales calculan automáticamente los precios de toppings de media porción a partir de la tarifa de referencia de pizza entera. |
| **Reglas de selección de modificadores de POS rotas** | Las elecciones mínimas obligatorias y los límites máximos de toppings se olvidan durante la configuración, permitiendo que pedidos imposibles lleguen a la pantalla de cocina. | Las reglas de restricción a nivel de grupo de modificadores aplican límites explícitos de selección mín/máx y topes de cantidad por artículo durante la preparación. |
| **Errores de despliegue y sincronización en POS en vivo** | Precios no verificados y filas de artículos incompletas se introducen directamente en terminales POS en vivo, provocando fallos de pago en las horas pico. | La validación de QA integrada señala los datos faltantes, simula recibos de clientes y genera una tabla de exportación al POS limpia y preverificada. |

---

## Para quién es esta plantilla de menú y precios de restaurante (roles objetivo y casos de uso)

Diseñada para restaurantes independientes, pizzerías, restaurantes de comida rápida (QSR), conceptos fast-casual y operaciones alimentarias multi-local que gestionan menús de varios tamaños, modificadores de ingredientes compartidos, cargos adicionales condicionales y reglas estrictas de pedidos de clientes.

### Escenarios de búsqueda por rol y casos de uso operativos

- **Propietarios de pizzerías y operadores de restaurantes independientes (`Pizza Menu Pricing Calculator & Topping Spreadsheet`)**
  Deje de perder dinero por toppings extra no facturados o de cobrar de más a los clientes en pizzas especiales. Modele los multiplicadores de toppings de pizza entera frente a media pizza, los recargos por tamaño y los artículos incluidos en la receta antes de actualizar sus pizarras físicas o su menú de pedidos en línea.
- **Gerentes de operaciones QSR y fast-casual (`Multi-Size Restaurant Menu Management Template`)**
  Estandarice jerarquías de menú, grupos de modificadores de combos y ajustes de precios de temporada en varias categorías de productos en un único archivo maestro offline centralizado, sin depender de frágiles scripts VBA.
- **Especialistas de incorporación de POS y administradores de menús (`Restaurant POS Menu Staging & Import Checklist`)**
  Limpie, organice y valide datos de menú sin procesar antes de crear o actualizar bases de datos en Toast, Square, Clover, Lightspeed o plataformas POS personalizadas. Detecte precios base faltantes y reglas mín/máx de modificadores rotas antes del lanzamiento.
- **Consultores de restaurantes e ingenieros de menús (`Menu Engineering & Modifier Architecture Toolkit`)**
  Entregue a sus clientes de hostelería un libro de configuración de menús repetible y listo para auditorías. Reestructure menús desordenados de clientes en grupos de modificadores lógicos y demuestre los precios exactos de pago de los clientes durante las consultorías.

**Para quién NO está diseñado este kit de herramientas:**
Negocios de comida que buscan un terminal de punto de venta en vivo, un procesador de pagos con tarjeta de crédito, un sistema de pantalla de cocina (KDS), un rastreador de consumo de inventario en bruto o una calculadora automática de nutrición de recetas. Este kit de herramientas es un **sistema de datos maestros de menú previo al POS, simulador de precios y QA de configuración**: prepara y valida la lógica de su menú comercial para que su plataforma POS funcione sin errores de precios.

---

## La lógica de negocio y metodología: resolver los puntos ciegos ocultos de los precios en restaurantes

*(Escrito para propietarios y operadores de restaurantes; no se requieren conocimientos técnicos ni avanzados de hojas de cálculo)*

Cuando los clientes piden comida, los precios del menú parecen simples a primera vista: elija un artículo, añada un topping, pague el total. Sin embargo, detrás del mostrador, la mayoría de los restaurantes independientes pierden dinero o frustran a los clientes por **tres puntos ciegos comerciales ocultos** que las hojas de cálculo estándar no detectan.

Aquí se explica, en lenguaje sencillo, cómo la lógica de negocio de este kit de herramientas protege sus márgenes de beneficio y la experiencia del cliente:

### 1. Eliminar el problema de "una actualización, diez errores" (única fuente de verdad)
* **El punto ciego:** Suponga que el costo mayorista del queso mozzarella o del tocino sube un 20%. En una hoja de cálculo típica de restaurante, "Extra Bacon" aparece por separado en Hamburguesas, Sándwiches, Ensaladas y Pizzas. Si actualiza el precio en tres lugares pero olvida el cuarto, pierde margen de beneficio o confunde a los clientes con precios inconsistentes.
* **Cómo lo corrige el kit de herramientas:** Introduce cada ingrediente y modificador **una sola vez** en una lista central. Ya sea que ese ingrediente se use en 2 artículos o en 50, actualizar su precio en una sola celda actualiza automáticamente cada artículo del menú conectado y cada cálculo de pedido en todo el libro.

### 2. Prevenir la trampa del "cobro doble en artículos especiales" (precios condicionales de receta)
* **El punto ciego:** Cuando un cliente pide una "Pizza Meat Lover" o una "Hamburguesa Deluxe", ya está pagando un precio base más alto porque el pepperoni, la salchicha o el queso están incluidos en la receta. Si la configuración de su menú trata todos los toppings por igual, su sistema podría cobrar al cliente *otra vez* solo por confirmar los ingredientes estándar o, si hace todos los toppings gratis para evitar quejas, pierde dinero cuando alguien pide carne *doble*.
* **Cómo lo corrige el kit de herramientas:** La plantilla separa **"Incluido por defecto"** de **"Extra de pago"**. Sabe que 1x Pepperoni en una Pizza de Pepperoni cuesta $0.00 adicional, mientras que aumentar la cantidad a 2x (Pepperoni Extra) cobra automáticamente el precio exacto de la mejora. Protege la confianza del cliente sin regalar comida.

### 3. Acabar con las conjeturas del cajero en toppings de media pizza y tamaños (lógica de precios proporcional)
* **El punto ciego:** Los clientes frecuentemente piden toppings solo en la mitad izquierda o derecha de una pizza, o cambian entre tamaños Pequeño, Medio y Grande. Sin una fórmula clara, el personal cobra el precio completo del topping de pizza entera (molestando al cliente) o reduce los precios a la mitad manualmente sin contabilizar la mano de obra de preparación (perjudicando sus márgenes).
* **Cómo lo corrige el kit de herramientas:** Usted define una única regla —por ejemplo, "un topping de media pizza cuesta el 60% de un topping de pizza entera". El kit de herramientas calcula automáticamente el cargo exacto, preciso al centavo, para las posiciones Mitad Izquierda, Mitad Derecha y Entera en todos los tamaños de pizza.

### 4. Bloquear tickets de cocina imposibles (barandillas de selección mín/máx)
* **El punto ciego:** Si un cliente pide un combo en línea y olvida elegir un acompañamiento obligatorio, o selecciona 12 salsas gratis cuando solo se permiten 2, su personal de cocina debe dejar de cocinar, llamar al cliente o asumir el costo.
* **Cómo lo corrige el kit de herramientas:** Usted define reglas de límites simples para cada grupo de opciones (p. ej., *"Debe elegir al menos 1 masa; máximo 1"* o *"Elija hasta 4 toppings estándar"*). El simulador de pedidos integrado prueba su menú contra estas reglas de inmediato y le muestra si un cliente podría realizar accidentalmente un pedido no válido.

### 5. Probar antes del lanzamiento (QA previo al POS)
* **El punto ciego:** La mayoría de los restaurantes descubren los errores de precios del menú *después* de actualizar su sistema POS, generalmente cuando un cliente en hora pico se queja en la caja o una aplicación de entrega en línea cobra un total incorrecto.
* **Cómo lo corrige el kit de herramientas:** El kit actúa como un "maniquí de prueba de choque" seguro para su menú. Puede simular pedidos reales de clientes, ver el recibo exacto antes y después de impuestos, y consultar un panel de salud automático que resalta en rojo los precios faltantes. Solo exporta su menú al POS cuando cada artículo pasa la inspección.

---

## Sobre el creador

Creo herramientas de negocio prácticas y basadas en fórmulas que convierten datos operativos desordenados en flujos de trabajo estructurados y repetibles de apoyo a la toma de decisiones.

La mayoría de los restaurantes independientes y pequeños negocios ya tienen sus datos de precios, pero están atrapados en hojas de cálculo desconectadas, exportaciones de POS desordenadas y hábitos del personal no documentados. El verdadero desafío operativo es conectar esos registros, aplicar reglas de precios consistentes y detectar errores de configuración antes de que lleguen a los clientes que pagan.

La **Calculadora de Precios de Menús de Restaurante y Kit de Configuración POS** brinda a los operadores de hostelería una validación de menús de nivel empresarial dentro de un flujo de trabajo transparente en navegador y Excel, sin sobrecarga innecesaria.

- 🌐 **[Pruebe ahora la calculadora gratuita de precios de menús en el navegador](https://hyvoid.github.io/restaurant-menu-pricing-template/)**
- 📥 **[Obtenga el libro Excel completo y reutilizable de datos maestros del menú (.xlsx)](https://theseusworkshop.com/l/dznbyd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=restaurant-menu-pricing-template)**


## Detalles técnicos

<details>
<summary>Para revisores técnicos, profesionales de Excel y desarrolladores: despliegue para inspeccionar la arquitectura del libro, la lógica de precios, las fórmulas y las reglas de validación.</summary>

### Arquitectura del libro

El kit de herramientas está diseñado como una base de datos relacional de configuración de menús y un motor de validación de precios, y no como una colección de hojas de cálculo de menús independientes.

Su arquitectura sigue tres principios:

1. **Única fuente de verdad:** Cada categoría, producto, tamaño, modificador y regla de precios tiene una ubicación designada en los datos maestros.
2. **Relaciones reutilizables:** Los artículos del menú heredan grupos de modificadores mediante asignaciones a nivel de categoría o de artículo, en lugar de duplicar registros de configuración.
3. **Validación y salida automatizadas:** Fórmulas dinámicas conectan los datos maestros, calculan los precios de los clientes, detectan errores de configuración y generan registros estandarizados de exportación al POS.

El libro contiene cinco módulos funcionales y 13 hojas de cálculo.

#### 1. Parámetros globales y reglas de precios

| Hoja de cálculo | Responsabilidad |
|---|---|
| `00_Param_Config` | Centraliza la moneda, la tasa de impuestos, los factores de precio entero/media, los recargos de modificadores premium, los umbrales operativos, los indicadores de estado y las reglas de precios condicionales. |

Esta hoja de cálculo es la capa de control global. Los supuestos de precios y los identificadores de estado del negocio se mantienen centralmente en lugar de incrustarse como valores fijos en fórmulas de cálculo individuales.

#### 2. Datos maestros de productos del menú

| Hoja de cálculo | Responsabilidad |
|---|---|
| `01_Menu_Categories` | Mantiene las categorías del menú, las subcategorías, el orden de visualización y la disponibilidad de las categorías. |
| `02_Menu_Items` | Almacena las identidades individuales de los productos, sus descripciones, las relaciones de categoría, los tipos de producto y la disponibilidad. |
| `03_Item_Sizes` | Almacena los precios de productos por tamaño y conecta cada tamaño con su artículo de menú principal. |

La separación entre la identidad del producto y los precios por tamaño evita que los datos maestros de productos se expandan horizontalmente cada vez que se introduce un nuevo tamaño u opción de precios.

#### 3. Modificadores reutilizables y precios multidimensionales

| Hoja de cálculo | Responsabilidad |
|---|---|
| `04_Modifier_Groups` | Define los grupos de modificadores visibles para el cliente y su alcance operativo. |
| `05_Modifiers` | Mantiene ingredientes y identidades de modificadores globalmente reutilizables. |
| `06_Modifier_Pricing` | Almacena los precios de los modificadores por tamaño de producto y nivel de porción, con cálculos de posición entera/izquierda/derecha. |

El mismo ingrediente puede reutilizarse en varios productos sin crear copias separadas de su registro maestro.

#### 4. Mapeo de relaciones y restricciones de selección

| Hoja de cálculo | Responsabilidad |
|---|---|
| `07_Default_Modifiers` | Mapea los productos especiales con sus ingredientes incluidos y cantidades predeterminadas. |
| `08_Group_Assignment` | Asigna grupos de modificadores a nivel de categoría o de producto individual. |
| `09_Selection_Rules` | Controla las selecciones mínimas, las selecciones máximas, los grupos obligatorios y los permisos de selección múltiple. |

Este módulo determina qué grupos de modificadores se aplican a cada producto y si la configuración de selección resultante del cliente es lógicamente válida.

#### 5. Simulación de pedidos, exportación al POS y monitoreo

| Hoja de cálculo | Responsabilidad |
|---|---|
| `10_Order_Test` | Simula las selecciones del cliente y calcula los precios finales mientras verifica el cumplimiento de las reglas de selección. |
| `11_POS_Export` | Genera una lista estandarizada de productos disponibles, tamaños, precios, resúmenes de modificadores y estado de sincronización. |
| `12_Dashboard` | Resume la salud del menú, las anomalías de configuración, la disponibilidad de productos y la preparación del POS. |

#### Flujo de datos

El flujo de datos completo sigue esta secuencia:

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

### Marco de decisiones

El kit de herramientas se construye en torno a cuatro decisiones operativas.

| Decisión                                   | Información requerida                                                                           | Resultado de la decisión                                                             |
| ------------------------------------------ | ---------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------- |
| ¿Qué se puede vender?                      | Registros de categoría, producto, tamaño y disponibilidad.                                             | Identifica la estructura válida del menú y las configuraciones de producto disponibles.    |
| ¿Qué deben pagar los clientes?                 | Precio base, precios de modificadores, ingredientes incluidos, cantidades de porción, posición e impuestos.    | Determina el precio calculado al cliente para una configuración de pedido específica. |
| ¿Qué pueden seleccionar los clientes?                 | Asignaciones de modificadores, pertenencia a grupos, estado obligatorio y límites de selección mínimos/máximos. | Identifica configuraciones de pedido válidas y no válidas.                        |
| ¿Está el menú listo para la sincronización con el POS? | Integridad de los datos maestros, integridad de precios, relaciones de modificadores y estado de validación.    | Separa los registros listos para exportar de los que requieren corrección.        |

El marco distingue tres tipos diferentes de información:

- Hechos del producto: qué vende el restaurante y qué tamaños están disponibles.
- Reglas comerciales: cómo los ingredientes, las cantidades y la posición afectan el precio de venta.
- Restricciones operativas: qué selecciones están permitidas y si la configuración está lista para el despliegue.

Esta separación evita que una descripción de producto, un supuesto de precios y una restricción de pedidos se traten como datos intercambiables.

### Tres trampas comunes de precios y configuración

#### Trampa 1 — Cobrar dos veces por ingredientes ya incluidos

La decisión: Fijar el precio al cliente de una pizza especial en función de sus ingredientes indicados.

El defecto no detectado: Un cálculo convencional de modificadores añade el precio estándar del topping sin importar si ese topping ya está incluido en el producto especial.

Considere una pizza Hawaiana con jamón y piña incluidos en su precio base.

| Selección del cliente | Tratamiento incorrecto                                         | Tratamiento correcto                                  |
| ------------------ | ----------------------------------------------------------- | -------------------------------------------------- |
| Jamón normal         | Añade un cargo por topping                                       | Sin cargo adicional                               |
| Piña light    | Añade un cargo por topping                                       | Sin cargo adicional                               |
| Jamón extra          | Añade un cargo por topping sin considerar la inclusión         | Aplica la regla configurada de precios por cantidad extra |
| Piña extra    | Usa el precio estándar del modificador sin verificar la inclusión | Aplica la regla configurada de precios por cantidad extra |

El enfoque incorrecto trata cada ingrediente seleccionado como una nueva compra. El enfoque corregido reconoce que un ingrediente puede formar parte de la definición comercial del producto.

El mapeo de `07_Default_Modifiers` identifica los ingredientes incluidos. Las reglas de precios condicionales determinan entonces el multiplicador aplicable según el estado de inclusión y el tipo de porción seleccionado.

#### Trampa 2 — Tratar los toppings de media pizza como precios fijos independientes

La decisión: Cobrar el mismo precio de modificador sin importar si un topping cubre toda la pizza o solo una mitad.

El defecto no detectado: La cantidad de producto que recibe el topping cambia, pero el cálculo de precios no refleja la posición seleccionada.

Por ejemplo, suponga que un topping de pizza entera cuesta $2.00.

| Posición | Precio incorrecto | Precio correcto |
| --------- | --------------- | ------------- |
| Entera     | $2.00           | $2.00         |
| Izquierda      | $2.00           | $1.00         |
| Derecha     | $2.00           | $1.00         |

El enfoque incorrecto cobra de más las selecciones de media pizza cuando la política del restaurante es el precio proporcional.

El enfoque corregido deriva los precios de mitad izquierda y mitad derecha a partir del precio de referencia de la pizza entera, usando factores de precios centralizados.

#### Trampa 3 — Tratar la preparación del menú para el POS como una tarea de entrada de datos

La decisión: Enviar la última hoja de cálculo del menú al proveedor de POS después de actualizar los precios de los productos.

El defecto no detectado: Una entrada de precios completada no establece que toda la configuración del producto sea válida.

Un producto puede tener un precio base y aun así contener:

- Un registro de tamaño faltante.
- Una relación de categoría no válida.
- Un grupo de modificadores no disponible.
- Mapeos de ingredientes predeterminados faltantes.
- Reglas de selección mínima y máxima inconsistentes.

| Enfoque de revisión                           | Riesgo                                             | Enfoque corregido                                         |
| ----------------------------------------- | ------------------------------------------------ | ---------------------------------------------------------- |
| Verificar si el nombre del producto existe.    | Los productos incompletos parecen estar listos.                | Validar la identidad del producto y las relaciones de categoría.      |
| Verificar si existe un precio base.        | Algunos tamaños disponibles permanecen sin precios.            | Validar cada combinación producto-tamaño configurada.         |
| Verificar si los grupos de modificadores están listados. | Grupos no válidos o vacíos pueden llegar al POS.       | Validar la pertenencia a grupos y las restricciones de selección.       |
| Exportar todo el menú.                   | Los errores se transfieren a la configuración del POS. | Exportar los registros elegibles y revisar el estado de sincronización. |

El enfoque corregido distingue la finalización de la entrada de datos de la preparación operativa.

### Escenario de ejemplo

Escenario: Probar una pizza Hawaiana Grande con piña extra en la mitad izquierda

Un restaurante quiere confirmar que sus precios de pizza especial siguen siendo correctos cuando los clientes cambian la cantidad y la posición de los ingredientes incluidos.

Suponga la siguiente configuración ilustrativa:

| Entrada                | Valor          |
| -------------------- | -------------- |
| Producto              | Pizza Hawaiana |
| Tamaño                 | Grande          |
| Precio Base           | $18.00         |
| Ingredientes Incluidos | Jamón, Piña |
| Modificador Seleccionado    | Piña      |
| Cantidad Seleccionada    | Extra          |
| Posición            | Izquierda           |
| Precio del Modificador Entero | $2.00          |
| Factor de Mitad Izquierda     | 0.50           |
| Impuesto sobre Ventas            | 8.00%          |

Paso 1 — Identificar el ingrediente incluido.

La piña está mapeada como ingrediente incluido de la pizza Hawaiana. Por lo tanto, no se vuelve a cobrar al cliente la cantidad estándar incluida.

Paso 2 — Determinar el multiplicador de cantidad aplicable.

El cliente selecciona Extra en lugar de Normal. Según la matriz ilustrativa de reglas, un ingrediente incluido seleccionado como Extra recibe un multiplicador de 1.00.

Paso 3 — Calcular el precio del modificador ajustado por posición.

```
Whole Modifier Price = $2.00

Left Half Price = $2.00 × 0.50

Raw Portion Price = $1.00

Applicable Multiplier = 1.00

Additional Charge = $1.00 × 1.00

Additional Charge = $1.00
```

Paso 4 — Calcular el precio al cliente.

| Cálculo                 | Monto |
| --------------------------- | ------ |
| Pizza Hawaiana Grande        | $18.00 |
| Piña Extra — Mitad Izquierda | $1.00  |
| Total antes de impuestos               | $19.00 |
| Impuesto sobre Ventas (8%)              | $1.52  |
| Precio Final al Cliente        | $20.52 |

Interpretación analítica

Una calculadora de toppings convencional podría cobrar de nuevo el precio estándar de la piña, incluso cuando el ingrediente ya está incluido. Alternativamente, una calculadora que ignora la posición podría cobrar los $2.00 completos por una selección de media pizza.

El resultado correcto refleja ambas condiciones comerciales: el estado de ingrediente incluido y la posición seleccionada.

Implicación de decisión: El restaurante puede validar el precio visible al cliente antes de modificar su menú en vivo, reduciendo el riesgo de cargos inconsistentes y correcciones de POS evitables.

Este escenario ilustra la lógica de precios configurada. Los resultados reales dependen de la matriz de precios, los precios de modificadores y la configuración de impuestos del propio restaurante.


### Referencia de fórmulas

El libro utiliza funciones nativas de matriz dinámica de Excel para mantener las relaciones entre los registros del menú, calcular precios y generar salidas de validación.

Las funciones principales incluyen:

| Función | Propósito |
|---|---|
| `LET` | Nombra los cálculos intermedios y mantiene legibles las fórmulas complejas. |
| `FILTER` | Extrae registros válidos y genera dinámicamente listas de salida. |
| `XLOOKUP` | Recupera información relacionada de productos, categorías, modificadores y precios. |
| `COUNTIF` / `COUNTIFS` | Cuenta registros y verifica condiciones de selección o configuración. |
| `UNIQUE` | Genera identificadores distintos de categorías y productos. |
| `SUM` / `SUMIFS` | Agrega cargos de modificadores y valores del negocio. |
| `IF` / `AND` / `OR` | Aplica condiciones de precios y reglas de validación. |
| `ROUND` | Controla la precisión de los cálculos monetarios. |
| `AVERAGEIFS` | Calcula los precios iniciales promedio a nivel de categoría. |

#### 1. Parámetros globales — `00_Param_Config`

La hoja de parámetros mantiene dos tipos de controles:

**Parámetros de configuración estáticos**

| Parámetro | Referencia | Propósito |
|---|---|---|
| Símbolo de Moneda | `B3` | Controla la presentación monetaria. |
| Tasa de Impuesto sobre Ventas | `B4` | Calcula los precios finales al cliente. |
| Factor de Mitad Izquierda | `B5` | Calcula los precios de modificadores de mitad izquierda. |
| Factor de Mitad Derecha | `B6` | Calcula los precios de modificadores de mitad derecha. |
| Factor Entero | `B7` | Controla los precios de porción completa. |
| Recargo de Modificador Premium | `B8` | Proporciona el ajuste de precios premium configurado. |
| Umbral de Advertencia de Grupos de Modificadores | `B9` | Señala asignaciones excesivas de grupos de modificadores. |
| Umbral de Advertencia de Precio Alto | `B10` | Identifica precios de productos inusualmente altos. |
| Línea Base de Disponibilidad para la Exportación al POS | `B11` | Define el estado utilizado para el filtrado de exportación al POS. |
| Estado Activo Predeterminado | `B12` | Centraliza el identificador de estado activo. |
| Estado Incluido | `B13` | Identifica los ingredientes incluidos predeterminados. |
| Nivel de Modificador Estándar | `B14` | Identifica el nivel de modificador estándar. |

**Matriz de precios condicionales**

El área `E3:J10` mantiene las reglas de precios condicionales.

Su lógica combina:

- Condición de Estado Incluido
- Condición de Tipo de Selección
- Resultado del Multiplicador de Precios

Se genera una clave de búsqueda compuesta para conectar las condiciones con el multiplicador de precios.

Esto evita que los mismos supuestos de precios queden incrustados de forma independiente en varias hojas de cálculo.

#### 2. Datos maestros del menú — `01_Menu_Categories`, `02_Menu_Items`, `03_Item_Sizes`

**Jerarquía de categorías**

`01_Menu_Categories!H3` genera las rutas de categorías:

```excel
=LET(
    cat_name, FILTER(B3:B100, E3:E100<>""),
    sub_name, FILTER(F3:F100, E3:E100<>""),
    cat_name & " > " & sub_name
)
````

Las rutas resultantes proporcionan relaciones de categorías legibles, tales como:

```
Pizza > Create Your Own
Pizza > Gourmet Pizza
Subs > Hot Subs
```

Recuento de productos activos por subcategoría

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

Estado de configuración de la categoría

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

Clave compuesta producto-tamaño

`03_Item_Sizes` genera una clave compuesta usando:

```
=Item_ID & "|" & Size_ID
```

Por ejemplo:

```
ITM001|SZ_SM
ITM001|SZ_LG
```

Esta clave permite al libro distinguir entre varios tamaños que pertenecen al mismo producto.

Validación de precios producto-tamaño

La capa de validación de precios verifica si un precio de tamaño configurado falta, no es válido o excede el umbral de advertencia centralizado.

El estado resultante lo utilizan tanto los datos maestros de productos como el proceso de exportación al POS.

#### 3. Datos maestros y precios de modificadores — `04_Modifier_Groups`, `05_Modifiers`, `06_Modifier_Pricing`

Recuento de modificadores activos

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

Búsqueda de reglas de selección del grupo de modificadores

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

La búsqueda correspondiente de selección máxima usa la columna C en `09_Selection_Rules`.

Un valor de retorno de `-1` identifica una regla faltante.

Dimensiones de precios de modificadores

`06_Modifier_Pricing` mantiene la relación entre:

```
Modifier ID
    +
Size ID
    +
Selection Type
```

La clave de precios compuesta permite recuperar el precio correcto del modificador para un tamaño de producto y nivel de cantidad seleccionados.

Los precios entero, izquierda y derecha se derivan de los factores de precios centralizados.

La arquitectura de precios separa el precio de referencia del cálculo de posición, permitiendo ajustar las políticas de precios específicas del restaurante sin reescribir el simulador de pedidos.

#### 4. Ingredientes predeterminados y mapeo de relaciones — `07_Default_Modifiers`, `08_Group_Assignment`, `09_Selection_Rules`

Clave de búsqueda de modificadores predeterminados

`07_Default_Modifiers!E3`:

```
=LET(
    itm_ids, FILTER(A3:A1000, A3:A1000<>""),
    mod_ids, FILTER(B3:B1000, A3:A1000<>""),
    itm_ids & "|" & mod_ids
)
```

Referencia de nombre de producto

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

Referencia de nombre de modificador

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

Validación de la configuración predeterminada

`07_Default_Modifiers!H3` verifica:

- Si el ID del producto existe.
- Si el ID del modificador existe.
- Si la misma combinación producto-modificador está duplicada.
- Si el estado incluido está correctamente configurado.

Asignaciones a nivel de categoría y de artículo

`08_Group_Assignment` usa el campo `Level` para distinguir el alcance de la asignación:

```
Level = Category
    Category ID = valid category
    Item ID = -

Level = Item
    Item ID = valid product
    Category ID = -
```

La capa de validación de asignaciones verifica el alcance efectivo e identifica asignaciones no válidas o excesivas.

Reglas de selección

`09_Selection_Rules` usa reglas a nivel de grupo para determinar:

- Selecciones mínimas.
- Selecciones máximas.
- Estado obligatorio.
- Si se permiten selecciones múltiples.

El simulador de pedidos y las fórmulas de validación de grupos de modificadores hacen referencia a estos valores.

#### 5. Simulación del precio al cliente — `10_Order_Test`

Esta hoja de cálculo combina el precio base del producto con los cargos reales de modificadores generados por la configuración de pedido seleccionada.

Nombre del producto seleccionado

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

Precio base del tamaño

`E4`:

```
=XLOOKUP(
    B3 & "|" & B4,
    '03_Item_Sizes'!F3#,
    '03_Item_Sizes'!D3:D1000,
    0
)
```

Cargo adicional total de modificadores

`E5`:

```
=SUM(J10#)
```

Precio al cliente antes de impuestos

`E6`:

```
=E4+E5
```

Precio final al cliente incluyendo impuestos

`E7`:

```
=ROUND(
    E6 * (1 + '00_Param_Config'!$B$4),
    2
)
```

Los cálculos detallados de modificadores usan la siguiente secuencia:

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

El precio final se calcula a partir de la selección realmente simulada y no de una cotización ensamblada manualmente.

#### 6. Exportación al POS — `11_POS_Export`

La hoja de exportación al POS filtra dinámicamente los productos según su disponibilidad configurada.

Filtrado de productos disponibles

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

Nombre del producto exportado

`B3`:

```
=XLOOKUP(
    A3#,
    '02_Menu_Items'!A3:A500,
    '02_Menu_Items'!B3:B500,
    ""
)
```

Nombre del tamaño exportado

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

Precio base exportado

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

Estado de sincronización con el POS

La validación de exportación combina la integridad del producto y la validación de precios por tamaño.

La salida prevista es:

```
READY_FOR_POS
HOLD_CHECK_ERROR
```

La exportación está diseñada para facilitar la revisión y transferencia de los cambios de configuración de productos. No es una integración directa con la API del POS ni un servicio de sincronización automática.

#### 7. Panel de control y monitoreo de la configuración — `12_Dashboard`

El panel de control combina seis indicadores operativos:

| KPI                       | Propósito del cálculo                           |
| ------------------------- | --------------------------------------------- |
| Total de Categorías Activas   | Cuenta las categorías de menú habilitadas.               |
| Total de Artículos Disponibles     | Cuenta los productos disponibles.                    |
| Total de Modificadores Reutilizables  | Cuenta los modificadores habilitados.                     |
| Total de Reglas de Precios Configuradas | Cuenta las reglas de precios configuradas.              |
| Total de Anomalías de Configuración    | Agrega las excepciones de configuración detectadas. |
| Recuento de Artículos Listos para POS      | Cuenta los registros marcados como listos para exportar.       |

Total de artículos disponibles

`C4`:

```
=COUNTIF(
    '02_Menu_Items'!G3:G500,
    '00_Param_Config'!$B$12
)
```

Total de anomalías de configuración

`F4`:

```
=COUNTIF('02_Menu_Items'!M3#, "❌*")
+ COUNTIF('03_Item_Sizes'!I3#, "❌*")
+ COUNTIF('04_Modifier_Groups'!I3#, "❌*")
+ COUNTIF('08_Group_Assignment'!I3#, "⚠️*")
+ COUNTIF('09_Selection_Rules'!H3#, "❌*")
```

Recuento de registros listos para POS

`G4`:

```
=COUNTIF(
    '11_POS_Export'!I3#,
    "READY_FOR_POS"
)
```

El panel de control proporciona una vista consolidada de la salud de la configuración del menú, mientras que las hojas subyacentes conservan los registros detallados necesarios para investigar y corregir las excepciones individuales.

### Reglas de validación

El marco de validación está diseñado para identificar errores antes de que afecten los precios visibles al cliente o la configuración del POS.

| Campo / Área                | Regla de validación                                                              | Comportamiento ante errores                                                      |
| --------------------------- | ---------------------------------------------------------------------------- | ------------------------------------------------------------------- |
| ID de Categoría                 | Debe ser único y hacer referencia a una categoría válida.                               | Las relaciones de categoría no válidas se marcan.                         |
| ID de Subcategoría              | Debe corresponder a una subcategoría existente.                                  | Se identifica la clasificación de producto no válida.                       |
| Categoría Activa             | Debe usar los valores de estado Sí/No configurados.                                | Los valores de estado no válidos afectan la validación de categorías.                   |
| ID de Artículo                     | Debe ser un identificador de producto único.                                         | Las identidades de producto no válidas o duplicadas requieren corrección.         |
| Categoría del Artículo               | Debe hacer referencia a una categoría y subcategoría válidas.                             | La validación de integridad del producto identifica relaciones inconsistentes. |
| Disponibilidad del Artículo           | Debe usar el identificador de estado activo configurado.                            | Una disponibilidad no válida puede impedir el filtrado correcto.                 |
| Tamaño del Artículo                   | Debe hacer referencia a una combinación válida de producto y tamaño.                         | Los registros producto-tamaño faltantes o no válidos se marcan.                |
| Precio Base                  | Debe estar presente y cumplir las comprobaciones de precios configuradas.                   | Los precios faltantes o anómalos impiden la preparación para el POS.                  |
| ID de Grupo de Modificadores           | Debe ser un identificador de grupo único.                                           | Las relaciones no válidas se marcan.                                  |
| ID de Modificador                 | Debe ser único y hacer referencia a un grupo de modificadores válido.                         | Se identifican los errores de búsqueda de modificadores.                              |
| Modificador Activo             | Debe usar el identificador de estado activo configurado.                            | Los modificadores inactivos se excluyen de los recuentos de modificadores activos.        |
| Precios de Modificadores            | Debe contener los registros de precios por tamaño y tipo de selección requeridos.           | Las matrices de precios incompletas requieren corrección.                     |
| Nivel de Modificador               | Debe corresponder a los identificadores de nivel de modificador configurados.                 | Una clasificación de nivel no válida afecta la validación de precios.             |
| Mapeo de Modificadores Predeterminados    | Los ID de producto y de modificador deben existir.                                         | Los ID no válidos se marcan.                                            |
| Unicidad de Modificadores Predeterminados | La misma combinación producto-modificador no debe duplicarse.                | Los mapeos duplicados se marcan.                                     |
| Estado Incluido             | Debe coincidir con el identificador de estado incluido configurado cuando corresponda.       | Una configuración incorrecta de inclusión predeterminada se marca.               |
| Nivel de Asignación            | Debe ser `Category` o `Item`.                                         | Se marca el alcance de asignación no válido.                             |
| Asignación a Nivel de Categoría   | Requiere un ID de Categoría válido y el marcador de posición de ID de Artículo prescrito.         | Se marca la configuración de alcance no válida.                             |
| Asignación a Nivel de Artículo       | Requiere un ID de Artículo válido y el marcador de posición de ID de Categoría prescrito.         | Se marca la configuración de alcance no válida.                             |
| Asignación de Grupo de Modificadores   | Debe hacer referencia a un grupo de modificadores existente.                                   | Las relaciones de grupo no válidas se marcan.                            |
| Capacidad del Grupo de Modificadores     | El recuento máximo de selecciones no debe exceder el recuento disponible de modificadores activos. | La configuración de selección requiere corrección.                        |
| Grupo de Modificadores Obligatorio     | Los grupos obligatorios deben tener una selección mínima mayor que cero.             | Se marca el conflicto lógico.                                        |
| Grupo de Modificadores Opcional     | Los grupos opcionales deben tener una selección mínima de cero.                     | Las reglas de selección inconsistentes se marcan.                           |
| Selección Múltiple          | Debe seguir el permiso de selección múltiple configurado.                    | Las configuraciones de selección no válidas requieren revisión.                    |
| Producto de Prueba del Cliente       | Debe hacer referencia a un producto válido.                                              | Una selección de producto no válida muestra una advertencia.                       |
| Tamaño de Prueba del Cliente          | Debe hacer referencia a un tamaño válido para el producto seleccionado.                        | Una búsqueda de precio por tamaño faltante devuelve un resultado no válido o de precio cero.  |
| Selección de Modificador del Cliente | Debe hacer referencia a un modificador disponible.                                        | Las referencias de modificadores no válidas requieren corrección.                     |
| Tipo de Porción del Cliente       | Debe usar `Light`, `Normal`, `Extra` o `Xxtra`.                             | Los valores de selección no válidos no son entradas de precios válidas.              |
| Posición del Cliente          | Debe usar `Whole`, `Left` o `Right`.                                        | Los valores de posición no válidos no son entradas de precios válidas.              |
| Precios Condicionales         | Debe resolverse a la regla de inclusión y cantidad configurada.                  | Una configuración de precios incorrecta o faltante requiere investigación.  |
| Precio al Cliente              | Debe reflejar el precio base, los cargos de modificadores y la tasa de impuestos configurada.      | El simulador de pedidos expone el cálculo resultante para su revisión.   |
| Exportación al POS                  | La validación del producto y del precio por tamaño debe aprobarse.                                 | `HOLD_CHECK_ERROR` reemplaza a `READY_FOR_POS`.                        |

Distinción operativa importante: La capa de validación verifica las relaciones configuradas y las reglas de negocio representadas en el libro. No verifica de forma independiente si un proveedor de POS acepta cada campo exportado ni si admite cada comportamiento de precios.

</details>


## Otras herramientas de esta serie

Más herramientas ligeras de apoyo a la toma de decisiones comerciales y libros operativos:

\- [Theseus Workshop](https://theseusworkshop.com) — Una colección de herramientas prácticas para problemas operativos y financieros recurrentes de pequeños negocios.

## Licencia

Este proyecto está licenciado bajo la Apache License 2.0.

Consulte el archivo LICENSE para ver los términos completos de la licencia.
