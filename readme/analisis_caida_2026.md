# Análisis Estratégico — Caída de Ventas 2026
## Rumbo Verde | Enero–Mayo 2026 vs 2025

**Autor:** [Tu Nombre] | Jefe de Tienda & Analista BI  
**Fecha:** Junio 2026  
**Herramientas:** Python · Pandas · Matplotlib · Datos reales Bsale

---

## 1. Contexto

Rumbo Verde es una tienda chilena de alimentos y productos naturales que opera con canales mixtos: tienda física/web propia y Uber Eats. Entre enero y mayo de 2026, las ventas acumularon **$158.9M CLP**, representando una caída de **-4.9% respecto al mismo período de 2025** ($167.2M).

Este documento descompone la caída para identificar sus causas reales, cuantificar el impacto de cada factor y proponer acciones concretas.

---

## 2. Magnitud del problema

| Indicador | Ene–May 2025 | Ene–May 2026 | Variación |
|---|---|---|---|
| Venta Bruta | $167.2M | $158.9M | **-4.9%** |
| Boletas | 4.893 | 4.926 | +0.7% |
| Unidades vendidas | 17.168 | 14.878 | **-13.3%** |
| Boleta promedio | $34.400 | $32.275 | -6.2% |
| Unidades / ticket | 3.50 | 3.04 | **-13.4%** |
| % Margen bruto | 38.6% | 35.5% | **-3.1pp** |
| Descuentos entregados | $8.15M | $5.86M | -28.1% |
| % Uber Eats | 10.2% | 21.5% | **+11.3pp** |

---

## 3. Diagnóstico — ¿Dónde está la caída?

### 3.1 El tráfico NO es el problema

Las boletas crecieron +0.7% — hay **más clientes** llegando a la tienda en 2026 que en 2025. El problema no es adquisición de clientes.

### 3.2 El problema es lo que compra cada cliente

Las unidades por ticket cayeron de **3.50 a 3.04 (-13.4%)** y la boleta promedio bajó de $34.400 a $32.275 (-6.2%). Los clientes están comprando **menos artículos por visita**.

### 3.3 Cuantificación de la caída por componente

Caída total: **-$8.27M CLP**

| Componente | Impacto | % de la caída |
|---|---|---|
| Caída en ticket promedio | -$10.40M | 126% |
| Recuperación por más boletas | +$1.07M | -13% (compensa parcialmente) |
| **Total neto** | **-$8.27M** | **100%** |

**Conclusión: la caída es 100% un problema de ticket, no de tráfico.**

### 3.4 ¿Por qué cayó el ticket si el precio unitario subió?

Aquí está la paradoja del período:

- El precio promedio por artículo **subió +8.5%** ($9.826 → $10.664)
- Pero las unidades por ticket **cayeron -13.4%**

Los clientes están comprando artículos más caros pero **en menor cantidad**. Esto sugiere que el alza de precios está generando un efecto sustitución: el cliente elige 2-3 productos en vez de 4-5.

---

## 4. El factor Uber Eats

Uber Eats pasó de representar el 10.2% al 21.5% de las ventas en el período — un crecimiento de +11.3pp en solo un año.

**Implicancia en el margen:**  
El margen cayó -3.1pp (de 38.6% a 35.5%). Parte de esta compresión se explica porque Uber Eats tiene comisiones de plataforma que reducen el margen neto por venta, y su participación se duplicó en el período.

**Implicancia en el ticket:**  
El ticket promedio en Uber Eats tiende a ser menor que en tienda física — el cliente de delivery compra 1-2 productos específicos, no hace una compra completa de despensa. Al crecer Uber como canal, arrastra el ticket promedio global hacia abajo.

---

## 5. Comportamiento mes a mes

| Mes | Var. Ventas | Var. Boletas | Var. Ticket | Var. Margen |
|---|---|---|---|---|
| Enero | -8.3% | -1.3% | -7.1% | +1.5pp |
| Febrero | -7.3% | +12.4% | -17.6% | -6.0pp |
| Marzo | **+13.5%** | +13.9% | -0.4% | -4.3pp |
| Abril | -13.7% | -5.1% | -9.0% | -4.8pp |
| Mayo | -7.9% | -12.5% | +5.2% | -1.8pp |

**Marzo es la excepción positiva** — el único mes con crecimiento en ventas (+13.5%), impulsado por alto volumen de boletas (+13.9%). Investigar qué ocurrió en marzo 2026 (campaña, evento, activación) para replicarlo.

**Febrero y abril son los meses más críticos** — combinan caída en boletas y caída en ticket simultáneamente.

---

## 6. Hipótesis explicativas

A partir del análisis de datos, propongo tres hipótesis no excluyentes:

**H1 — Efecto precio:** El alza de precios promedio (+8.5%) está reduciendo la cantidad de ítems por compra. Los clientes ajustan su canasta manteniendo el número de visitas pero comprando menos unidades.

**H2 — Mix de canal:** El crecimiento acelerado de Uber Eats (canal de ticket bajo) está comprimiendo tanto el ticket promedio global como el margen. El negocio crece en transacciones pero no en valor por transacción.

**H3 — Ausencia de producto ancla:** En 2025, los meses fuertes correlacionaban con alta presencia de Wellplus y productos de alto ticket. Si el mix de productos cambió (menos suplementos premium, más snacks o bebidas de menor valor), el ticket caería naturalmente.

---

## 7. Recomendaciones

### Inmediatas (0–30 días)

**Estrategia de bundle:** Crear combos de 2-3 productos complementarios con precio especial. El objetivo es subir las unidades/ticket de 3.04 a 3.5+. Ejemplo: "Kit Bienestar Semanal" con 3R Desconecta + Omega 3 + Té adaptógeno.

**Optimizar catálogo Uber Eats:** Priorizar en la plataforma los productos de mayor ticket y margen (suplementos, aceites esenciales DoTerra, línea 3R) en vez de snacks de bajo valor. Uber ya es el 21.5% del negocio — vale la pena gestionarlo estratégicamente.

### Corto plazo (30–90 días)

**Investigar marzo 2026:** Es el único mes con crecimiento. Identificar la causa (¿campaña? ¿nuevo producto? ¿activación?) y diseñar un plan para replicarlo en julio–agosto, que históricamente son los meses más fuertes del año.

**Revisión de política de precios:** El alza de +8.5% en precio unitario promedio puede estar en el límite de lo que el cliente absorbe. Evaluar si hay categorías donde el precio superó el umbral de decisión.

### Mediano plazo (90–180 días)

**KPI de seguimiento mensual:** Incorporar "unidades/ticket" como KPI de alerta temprana en los reportes. Una caída sostenida bajo 3.0 debe activar una revisión de estrategia.

**Análisis de cohorte de clientes:** Cruzar los datos de clientes registrados con el período para evaluar si la caída proviene de clientes que compraban más y ahora compran menos, o de nuevos clientes con ticket estructuralmente menor.

---

## 8. Conclusión

La caída de -4.9% en ventas 2026 **no es una crisis de tráfico** — los clientes siguen llegando. Es una señal de que el negocio está creciendo en volumen de transacciones pero perdiendo valor por transacción, comprimido simultáneamente por el alza de precios y el crecimiento del canal Uber Eats.

La buena noticia: el margen, aunque comprimido, sigue siendo alto (35.5%) y el negocio mantiene solidez operacional. El período julio–agosto 2026 será determinante — históricamente son los meses más fuertes del año y representan la oportunidad real de cerrar 2026 en positivo respecto a 2025.

---

*Análisis construido sobre matriz de 53 meses (enero 2022 – mayo 2026) con datos reales exportados desde Bsale. Procesado con Python/Pandas.*
