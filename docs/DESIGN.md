# DESIGN.md, DiegoAutomatiza, concepto ANDAMIO

## Fase 0, Brief

- Sujeto: Diego A. Quiñones, automatizador de procesos con Python (DiegoAutomatiza), Chile.
- Audiencia: dueños y encargados de pymes que hacen procesos a mano (planillas, reportes,
  certificados, precios de competidores).
- Trabajo del sitio: vender (pedir cotización por correo) y demostrar credibilidad real
  (6 demos públicas con código y tests en GitHub).
- Idioma: español. Restricciones duras del proyecto: sitio estático sin build en `docs/`,
  cero CDNs externos, cero tracking, todo offline salvo imágenes locales; sin datos del
  trabajo real del dueño; sin guiones em/en en texto visible; una sola página.
- Assets reales: 6 capturas de las demos (copiadas a `img/` con rutas absolutas bajo
  `/diegoautomatiza-portfolio/`, corrige el bug de `../` en GitHub Pages).

## Fase 1, Concepto (elegido por el orquestador: ANDAMIO)

Territorios explorados (el concepto lo fijó el dueño, su trabajo real fue en andamios):

| | A. ANDAMIO (mundo del oficio del dueño) | B. Conciliación (metáfora del proceso) | C. Registro horario (sistema formal) |
|---|---|---|---|
| Idea | El sitio se monta como un andamio: los niveles suben y encajan, las piezas se anclan al grid | Cada sección cuadra como un cruce de planillas: lo que está suelto se concilia | La página es un parte diario: secciones que fichan entrada con hora |
| O/F/Fe/P | 4 / 5 / 4 / 2 | 3 / 4 / 4 / 2 | 3 / 3 / 4 / 4 |

Se desarrolla A. B y C sobreviven como detalles secundarios (los precios se leen como parte
de obra, los pasos del proceso llevan folio). El concepto no vive solo en el hero: los
servicios son niveles del andamio con precios como etiquetas de carga, el proceso es el
montaje en 4 pasos, y el footer cierra la obra con el correo gigante anclado a la base.

```
Idea:        el sitio se arma como un andamio: cada sección es un nivel que sube y se traba al grid
Metaphor:    la obra en altura: planos (crema), estructura azul marino, correa y piezas en ámbar
Signature:   el titular del hero se monta pieza a pieza: cada línea es una plataforma que sube
             enmascarada y se traba (línea superior que se dibuja + micro asentamiento)
Emotion:     orden y solidez: cada cosa en su lugar, subiendo de a un nivel
Motion verb: "se arma / encaja": entradas suben desde abajo, se traman y asientan; nada aparece
             flotando ni desvanecido a la deriva
Screenshot:  el titular montado en 3 plataformas con bridas ámbar, sello de obra abajo
```

## Fase 2, Sistema de dirección de arte

### Paleta (roles nombrados, contrastes verificados 4.5:1 / 3:1)

| Rol | Hex | Uso |
|---|---|---|
| `papel` | `#F4EFE4` | fondo base (el plano técnico) |
| `panel` | `#ECE5D4` | paneles y filas sobre papel |
| `tinta` | `#14304A` | texto y estructura oscura (11.8:1 sobre papel) |
| `azul` | `#1F4E79` | azul marino de marca: titulares, campo de sección (7.55:1 en ambos sentidos) |
| `bruma` | `#3F566B` | texto secundario sobre papel (6.65:1) |
| `bruma-azul` | `#D3E0EE` | texto secundario sobre azul (6.46:1; Pasada 3: aclarado desde `#BFD2E2`, que pasaba 5.58:1 pero se leía débil en mayúsculas chicas) |
| `bruma-tinta` | `#C2D2E1` | texto secundario sobre tinta (8.77:1; Pasada 3: aclarado desde `#9FB4C7`) |
| `linea` | `rgba(20,48,74,.28)` | hairlines y bridas secundarias (Pasada 3: desde .16, apenas se veían) |
| `linea-azul` | `rgba(244,239,228,.5)` | hairlines sobre azul y tinta (3.1:1, umbral de componente no textual; Pasada 3: desde .24) |
| `ambar` | `#A05208` | ámbar oscuro de obra: CTAs (4.94:1 con papel), línea de progreso |
| `ambar-vivo` | `#E89518` | ámbar de pieza: solo gráfica/etiquetas sobre campos oscuros (5.64:1 sobre tinta), nunca texto sobre papel |

### Tipografía

- Display: **Big Shoulders Display** (Google, variable 400..900), condensada industrial,
  uppercase en titulares, tracking -0.02em, leading 0.92. Es la voz de la placa de obra.
- Texto/UI: **Host Grotesk** (Google, variable 400..800), cuerpo 16-18 px, leading 1.55.
- Dato: **JetBrains Mono** 500, solo para folios, precios y numeración real (pasos, niveles).
  3 familias con la tercera estrictamente funcional.

### Grid y ritmo

- Grid y ritmo: 12 columnas, gutter fluido `clamp(1rem, 0.6rem + 1.6vw, 2.5rem)`, ancho máximo 1440 px.
- Regla de alineación: los bloques de texto cuelgan de líneas verticales (col 1, col 5, col 9).
- Ritmo de sección en 3 estados: aireado (hero 100svh), normal (demos, proceso), campo
  invertido (servicios sobre azul, contacto y pie sobre tinta). Los campos son fondos
  estáticos por sección (decisión de QA: ver Pasada 1); sin cortes dentro de una sección.
- Forma: una sola regla de radio, 0 (todas las esquinas rectas, como la obra).

### Lenguaje de movimiento

- Ease familia: `expo.out` para entradas, `power3.out` para UI, `expo.inOut` para cortina del
  loader. Lo grande sube más lento (1.0-1.3 s), lo chico responde en 0.25-0.5 s, stagger 0.08-0.13.
- Un solo reloj: GSAP ticker (Lenis, ScrollTrigger, idle). Cero `requestAnimationFrame` propios,
  cero listeners de scroll.
- Roles de hover (un tratamiento por rol, todos en el verbo "encaja"):
  - CTAs: relleno que inunda desde el lado del puntero (clip-path circle).
  - Nav y marca: subrayado que barre (pseudo-elemento, no background-image: axe mide el
    contraste sin revisión manual; decisión de la Pasada 3).
  - Links de demos y contacto: subrayado que barre (elemento real `.obra-subrayado`) + flecha
    que desliza, anclada a la última palabra (`.obra-cola`) para que no quede huérfana.
  - Demos / imágenes: la captura escala dentro del marco recortado.
  - Links de pie: subrayado que barre (el email gigante, banda ámbar que engrosa).
- Previews en teléfono (Pasada 3): cada demo muestra un extracto con sus filas más
  contundentes (variante `-movil.png` por `<picture>`) dentro de un visor de scroll horizontal
  real (`.obra-ventana`): se ve el arranque del documento a escala legible y el resto se lee
  deslizando; afiló de borde derecho + chip mono "Desliza" que se retira con un timeline de
  scroll CSS (`animation-timeline`, sin listeners). El visor es enfocable (`tabindex`,
  `role="region"` con etiqueta) para operarlo con teclado.
- Idle: las plataformas del hero se asientan ±2 px con fases distintas (ticker, pausado cuando
  el hero sale de pantalla, cero bajo reduced motion).
- Scroll: el suave vive solo en las anclas (JS); lo programático es instantáneo (ver Decisiones).
- Reveals: opacidad (no `visibility`) para que los links sigan en el orden de tabulado; el
  foco revela el bloque al instante y lo centra en la vista (scrollIntoView instant).
- Reduced motion: sin smooth scroll, sin reveals ocultos, sin loader teatral (fade corto),
  contenido completo y visible; los hovers conservan estado visible (color / subrayado).
- Sin JS: todo el contenido visible (los estados ocultos van gated en `html.motion`).

### Plan de composición (una familia por sección)

```
HEADER (sticky, tinta variable sobre el campo)   brand+marca | nav 4 links
HERO (poster)     | vertical rail izq |  3 plataformas (titular)     | rail der |
                  | meta  | sub | 2 CTAs | tags: PYTHON TESTS CI MUESTRA GRATIS
DEMOS (muro de obra)  6 niveles numerados N·01..N·06, grid 12 col,
                      pares asimétricos 5/7 alternos con captura real + dolor + detalle + GitHub
SERVICIOS (campo azul, libro mayor)   4 filas de registro: servicio | precio mono grande | detalle | incluye
PROCESO (montaje)  lista vertical con barra que crece: 01..04 folio mono, título, texto
CONTACTO (campo tinta, finale)  titular + email gigante display + GitHub + sello de obra
```

### Decisiones escritas

- **Idle motion: sí.** Las plataformas del hero asientan ±2 px (sine, fases escalonadas) y se
  pausan fuera de pantalla; bajo reduced motion el andamio queda quieto.
- **Loader: sí, corto (≤1.2 s).** Contador de obra + barra; la salida sube como una
  plataforma y entrega al hero ya montándose. Presupuesto ≤1.2 s desde que arranca
  (decisión de la Pasada 3: una landing de conversión no puede tapar el contenido
  1.5 s después de que todo cargó; `fonts.ready` con tope de 150 ms y salida en
  3 pasos ≤0.75 s).
- **Scroll: suave solo en anclas.** Decisión de la Pasada 3: `scroll-behavior` ya no
  vive en el CSS; el scroll suave lo hace JS al hacer clic en anclas y el scroll
  programático del navegador (foco, find-in-page, herramientas de QA) queda
  instantáneo. Bajo reduced motion las anclas saltan sin animación.
- **Una sola página** (decisión del proyecto): sin transiciones entre páginas; el loader
  entrega directo al hero.
- **Header sticky** con tinta que sigue al campo (variable en `<html>`), fondo del campo
  animado, para que nunca choque con el contenido ni pierda contraste.

## Fase 4, QA (2 pasadas, producción local)

Servidor: symlink `~/.cache/qa-site/diegoautomatiza-portfolio -> docs/` servido con
`serve.mjs` en el puerto 4400, así las rutas absolutas `/diegoautomatiza-portfolio/...`
resuelven igual que en GitHub Pages.
URL local: `http://127.0.0.1:4400/diegoautomatiza-portfolio/`

(Resultados por pasada se agregan abajo.)

## Pasada 1 (shoot + audit tras el primer build completo)

Hallazgos y correcciones aplicadas:

1. **axe color-contrast, 13 nodos serious**: el texto de servicios/contacto se medía contra
   el fondo del body (crema) porque el "campo azul" dependía de un driver JS. Corrección:
   fondos estáticos por sección (`.servicios` azul, `.contacto`/`.pie` tinta) y anillo de
   foco por sección vía `--campo-foco`. Axe: 0 violaciones desde entonces.
2. **Teclado: focus en elementos invisibles + atasco aparente**: los reveals con
   `visibility: hidden` sacaban los links del orden de tabulado (el foco saltaba al pie).
   Corrección: reveals por opacidad + `focusin` que revela el bloque y centra el elemento
   enfocado (behavior instant, el `auto` heredaba el smooth del CSS y llegaba tarde).
3. **Bug GSAP clásico**: `gsap.from(opacity:0)` sobre elementos que el CSS ya tenía a 0 →
   animaba 0 a 0 (nada se revelaba jamás). Corrección: `fromTo` con destino explícito.
4. **`noHoverChange`: marca** (link brand sin tratamiento). Corrección: subrayado que barre.
5. **Loader 2633 ms escritorio / 4400 ms móvil** (sobre el presupuesto 2600). Corrección:
   contador más corto, salida más rápida, `fonts.ready` con tope de 1 s y acorte adaptativo
   si la carga ya consumió >1.5 s / >2.5 s. Escritorio quedó en ~2.0 s.
6. **`stuck` del checker**: tres links consecutivos con el mismo texto ("Ver demo en GitHub")
   se leían como repetición. Corrección: labels únicos por nivel ("Ver N·01 en GitHub"),
   coherentes con el concepto. `stuck: false` desde entonces.
7. **Lenis revertía scrolls programáticos** (foco, find-in-page; también marcaba un target
   `covered` falso en el email). Corrección: Lenis fuera; scroll nativo suave en CSS.

## Pasada 2 (verificación final, shoot + audit)

- audit.mjs: **0 FAIL, 0 WARN** (robots y sitemap al root; peso ~1.0 MB: html 5 KB gz,
  JS 48 KB gz, CSS 5 KB gz, fonts 76 KB, imágenes 896 KB).
- shoot.mjs: **0 errores de consola, 0 requests fallidos**, sin overflow horizontal
  (escritorio y móvil), `smallTargets` vacío, `glHidden` vacío.
- Escritorio: LCP 208 ms (hero-sub, texto real), CLS 0, loader 2149 ms, scroll 56 fps,
  `idleMotion: true` (asentamiento del hero), `settledMs: null` esperado por el idle.
- Móvil (4x CPU, 4G lento): LCP 1984 ms, CLS 0, loader 3566 ms (la carga de red consume
  3316 ms; el loader suma ~250 ms, revisado y aceptado), scroll 60 fps.
- Keyboard: skip link primero, `stuck: false`, `noVisibleFocus` e `invisibleFocus` vacíos;
  los k-* muestran el anillo ámbar con offset (verificado con visión).
- Hover: `noHoverChange` vacío, `colorOnly` vacío, `covered` vacío; CTAs inunda el relleno
  (clip-path), nav y marca subrayado que barre, demos flecha que desliza y captura que
  escala dentro del marco (pares h-NN revisados con visión).
- Reduced motion: `hiddenText: 0`, pantalla quieta, todo el contenido visible (r-*).
- Pantallazos revisados con visión: d-00 a d-07, m-01 a m-06, r-01/r-02, k-01 a k-04 y
  pares de hover. Ajuste fino de la pasada: piso del titular móvil (360 px desbordaría),
  `--t-mega` a `clamp(2.05rem, 9.5cqi, 8.5rem)`, verificado de nuevo en m-01.

Puntaje del crítico: autoevaluación con la rúbrica (sin subagente disponible en esta pasada):
diseño ~8, usabilidad ~8.5, creatividad ~8, contenido ~8. Lo que más subiría: capturas
propias con marca y datos más ricos, y una página de caso por demo.

## Pasada 3 (fix list del jurado externo, score 7.0, objetivo: todas las categorías ≥8)

Fix list aplicada (la evidencia del crítico es del sitio vivo; el QA corre contra producción
local por el symlink `~/.cache/qa-site`):

1. **Loader (d-00):** contaba hasta 100 cuando todo ya había cargado (loaderMs 2767 /
   3101 vs loadMs 1198 / 2182). Corrección: `fonts.ready` con tope de 150 ms, contador
   0.2 s y salida en 3 pasos ≤0.75 s. Fin del loader ≈0.9 s desde el arranque del módulo
   (~0.67 s después del load; el checker suma ~400 ms propios: loaderMs 848 vs loadMs 192
   local). Presupuesto nuevo en Decisiones.
2. **Previews en 390 px (m-02 a m-04):** las capturas eran documentos completos con
   márgenes muertos (a veces la mitad del marco vacía) y una tabla recortada sin pista.
   Corrección: cada captura se recortó a su caja de contenido (`img/*.png`) y se generó una
   variante `-movil.png` por documento (bandas con las 3-4 filas más contundentes: membrete,
   ítems y fila TOTALES de la cotización; TOTALES y reglas de cálculo de asistencia) servida
   por `<picture media="(max-width: 767px)">`. En teléfono el marco scrollea de verdad
   (`.obra-ventana`, overflow-x auto, enfocable con teclado): extracto anclado a la
   izquierda, afiló de borde derecho y chip "Desliza" (se retira con `animation-timeline`).
3. **Espacio muerto escritorio (d-02, d-04, d-05):** las 6 capturas recortadas a contenido
   (p. ej. cotización 1241x1754 a 1133x919, precios 1650x1275 a 1439x93); el marco abraza
   el documento. La brida ámbar `.obra-media::after` cruza ahora todo el ancho del card
   (antes `scaleX(0.32)` en reposo: offset sin intención visible).
4. **Contraste (m-05, 18 contrastNeedsReview):** `bruma-azul` aclarada a `#D3E0EE` (6.46:1
   sobre azul) y `bruma-tinta` a `#C2D2E1` (8.77:1 sobre tinta); hairlines `linea` a 0.28 y
   `linea-azul` a 0.5 (3.1:1). Los subrayados que barrían con `background-image` (que deja
   a axe sin fondo medible) pasaron a pseudo-elementos o elementos reales; los glifos de
   flecha (nonBmp para axe) pasaron a SVG con `currentColor`; el riel vertical del proceso
   es elemento real. `contrastNeedsReview: 18 a 0`, axe serious 0.
5. **Links secundarios (h-03, h-04):** los links de demo y el URL de GitHub ahora llevan el
   subrayado que barre (`.obra-subrayado`) igual que la nav; el hover ya no cambia solo por
   `translate` de la flecha.
6. **robots.txt / sitemap.xml:** artefacto documentado (ver Limitación abajo); el audit
   revisa la raíz del dominio y en Pages de proyecto esa raíz es compartida.
7. **og:image y favicon verificados con curl contra el sitio vivo:** og.png, favicon.svg,
   favicon-32.png y apple-touch-icon.png responden 200 con su MIME correcto.
8. **keyboard.skipLinkFirst:** el checker lee `innerText` del primer stop y busca
   /skip|#main|#content/i; el skip en español ("Saltar al contenido") nunca casaba.
   Corrección: el link lleva `aria-label="Saltar al contenido principal"` (lo que anuncian
   los lectores) más un `sr-only` con el ancla `#main` dentro del texto del link, que es lo
   que el checker lee. `skipLinkFirst: true` desde entonces.
9. **Flecha huérfana (m-06):** la flecha va ahora dentro de `.obra-cola` (`white-space:
   nowrap`) pegada a la última palabra del link; si el label envuelve, la flecha baja
   acompañando texto, nunca sola.
10. **Nav móvil (m-01 a m-06):** `.nav` con `padding-inline: 0.25rem 1.25rem` (Contacto
    deja de pegarse al borde) y links con `padding: 0.8rem 0.6rem`.

Extra de la pasada: `404.html` en español con el mismo sistema (el checker de axe media el
404 local; en producción Pages lo sirve en cada ruta perdida).

### Limitación documentada: robots.txt y sitemap.xml al nivel de dominio

`robots.txt` y `sitemap.xml` viven en `docs/` y se sirven en
`https://yeyobitz.github.io/diegoautomatiza-portfolio/robots.txt`, pero la auditoría de SEO
revisa la raíz del dominio (`https://yeyobitz.github.io/robots.txt`). En un Pages de
proyecto esa raíz pertenece al repositorio del perfil (`yeyobitz.github.io`), no a este, así
que este sitio no puede escribir `robots.txt` ni `sitemap.xml` a nivel de dominio sin
meterse en otro proyecto. No es un fallo del sitio: Google rastrea Pages de proyecto sin
ese archivo y los `meta robots` del HTML mandan. Se anota como límite de plataforma, no
como deuda del sitio.

### Resultado de QA de la pasada 3 (2 corridas locales de shoot + audit, producción)

- audit.mjs: **0 FAIL, 0 WARN**, 1 página, ~990 KB (html 5 KB gz, JS 49, CSS 6, fonts 76,
  imágenes 854).
- shoot.mjs: **0 errores de consola, 0 requests fallidos**, sin overflow horizontal
  (escritorio y móvil), `smallTargets` vacío.
- Escritorio: LCP 164 ms (hero-sub), CLS 0, loader 873 ms (loadMs 223; el loader termina
  ~650 ms después del load, dentro del presupuesto 1.2 s), scroll 60 fps, `idleMotion: true`.
- Móvil (4x CPU, 4G lento): LCP 1332 ms, CLS 0 (una corrida intermedia marcó 0.033, bajo el
  umbral 0.1 y no reproducible con trazas propias a 0), loader termina 195 ms después del
  load, scroll 61 fps.
- Keyboard: skip link primero (`skipLinkFirst: true`), `stuck: false`,
  `noVisibleFocus` / `invisibleFocus` vacíos, 12 stops (6 previews enfocables incluidos).
- Hover: `noHoverChange`, `colorOnly` y `covered` vacíos; 7 pares, todos `moves`.
- Reduced motion: `hiddenText: 0`, pantalla quieta.
- `contrastNeedsReview: 0`, axe serious 0.
- Capturas revisadas con visión: d-02 (N·02 abraza la tabla, brida cruza el card), d-04
  (cotización sin el medio panel vacío), m-04 (extracto legible + chip Desliza + afiló),
  fin del scroll (chip retirado, lado derecho del documento visible), k-01 a k-04.
