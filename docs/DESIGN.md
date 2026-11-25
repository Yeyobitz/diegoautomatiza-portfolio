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
| `bruma-azul` | `#BFD2E2` | texto secundario sobre azul (5.58:1) |
| `linea` | `rgba(20,48,74,.16)` | hairlines y bridas secundarias |
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
  - Nav: subrayado que barre.
  - Demos / imágenes: la captura escala dentro del marco recortado; flecha que desliza en links.
  - Links de pie: subrayado que barre (el email gigante).
- Idle: las plataformas del hero se asientan ±2 px con fases distintas (ticker, pausado cuando
  el hero sale de pantalla, cero bajo reduced motion).
- Scroll: nativo (sin librería de smooth scroll). Decisión de QA: Lenis revertía los scrolls
  programáticos del navegador (foco de teclado, find-in-page), dejando el foco fuera de la
  vista; se usa `scroll-behavior: smooth` en CSS y el verbo de movimiento queda intacto.
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
- **Loader: sí, corto (≈1.6 s).** Contador de obra + barra; la salida sube como una plataforma
  y entrega al hero ya montándose. Presupuesto ≤2.6 s.
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
