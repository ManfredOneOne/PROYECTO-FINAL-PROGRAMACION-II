<p align="center"><img src="logo_umg.png" alt="Logo de la Universidad Mariano Gálvez de Guatemala" width="140"></p>

# CODEBRAN · Proyecto Final de Programación II

**Bodega y cuantificación de obra**

**Universidad Mariano Gálvez de Guatemala**
Facultad de Ingeniería · Ingeniería Civil · Sección «B» · Antigua Guatemala

**Integrantes**

- Obed Barrientos
- Mayelin Gisel García
- Dervin Velásquez
- Manfred Urízar Ruano

---

## 1. Descripción

**CODEBRAN** es una herramienta para llevar el control de la bodega de una obra y calcular los materiales que pide cada elemento constructivo. El inventario y los cálculos son **independientes**: calcular no lee ni descuenta nada de la bodega. Cada cálculo se puede descargar como **memoria de cálculo en Excel** con todas sus fórmulas y con dibujos (sección típica e isométrico). Se entrega en dos versiones con la misma lógica de cálculo:

| Archivo | Qué es | Cómo se usa |
|---|---|---|
| `index.html` | Entrada principal: abre el programa web (sirve para publicarlo en un servidor o abrirlo con doble clic) | Abrir con doble clic en el navegador |
| `proyecto_final_v3.html` | Página web con cuentas de usuario, inventario, calculadoras (incluidos los muros de contención), dibujos, memoria en Excel, tablas de normativa y gestión de proyectos públicos y privados | Abrir con doble clic en el navegador |
| `proyecto_final.py` | Programa de consola en Python con cuentas de usuario, el inventario, las mismas calculadoras y la misma memoria en Excel con fórmulas y dibujos (sin gestión de proyectos) | `python proyecto_final.py` |
| `firestore.rules` | Reglas de la base de datos en línea (Firebase) para usar cuentas en cualquier dispositivo | Se pegan en la consola de Firebase |
| `manifest.webmanifest`, `sw.js`, `iconos/` | Permiten instalar la página en el teléfono y abrirla sin internet | Automático |
| `logo_umg.png` | Logo de la universidad | — |
| `respaldo_antes_de_cambios/` | Versiones anteriores del programa | Solo de consulta |

El programa en Python solo usa la librería estándar (no hay que instalar nada) y funciona con Python 3.8 o más reciente. Crea junto a él los archivos `cuentas_programa.json` (cuentas), `datos_proyecto.json` (bodega y cálculos) y la carpeta `memorias/` (Excel descargados).

Las cantidades son valores de referencia para un estimado. El diseño final debe revisarlo un ingeniero civil colegiado.

### Uso desde el teléfono

La página se adapta a cualquier pantalla y no necesita instalarse: basta con abrirla en el navegador del teléfono, la tableta o la computadora.

- **Menú inferior fijo:** Inicio, Inventario, Calcular, Muros y **Más**, siempre a la mano del pulgar y sin desplazar el menú hacia los lados.
- **Botón «Más»:** abre una hoja con todos los módulos (también Carreteras, Normativa y Obras), cada uno con una descripción corta. Se cierra con «Cerrar», tocando fuera de la hoja o con la tecla Esc.
- **Tarjetas de módulos en el Inicio:** acceso directo a cada módulo con un toque.
- **Engranaje de ajustes ⚙ (arriba a la derecha):** al tocarlo muestra el **modo de pantalla Claro / Oscuro / Auto** y el **idioma ES / EN**. «Auto» sigue la configuración del dispositivo, y la elección queda guardada en ese navegador. Se cierra al tocar fuera o con Esc.
- **Usuario debajo del engranaje:** círculo con la foto de perfil (o las iniciales). Al tocarlo muestra el nombre y el correo, «Cambiar foto», «Quitar foto», «Mi cuenta y contraseña» y «Cerrar sesión».
- **Vista de bienvenida:** antes del Inicio aparece la portada con el logo y los integrantes, el saludo «Bienvenido» y una escena animada de **Tikal** (Templo I del Gran Jaguar, Templo II, Acrópolis Norte, selva y guacamayas de día; estrellas, luna y luciérnagas de noche). En el cielo se lee **GUATEMALA** con un símbolo patrio dentro de cada letra. Ahí mismo se **inicia sesión o se crea la cuenta**. Vuelve a aparecer al cerrar sesión.
- **Menú lateral (computadora):** se despliega al pasar el cursor; el logo de la universidad conserva su tamaño con el menú cerrado y junto a él se lee **CODEBRAN**.
- **Tarjetas de módulos de vidrio líquido:** cada módulo es una tarjeta redondeada con una **escena animada** que corresponde a su título (bodega con montacargas, plano con calculadora, muro de block, carretera con tránsito, libro de normas, obra con grúa, cronograma y el ingeniero de «Mi cuenta»), de día en modo claro y de noche en modo oscuro. El nombre y la descripción van sobre un panel de vidrio esmerilado.
- **Gestión de proyectos:** módulo del menú (en el teléfono, dentro de «Más») para proyectos públicos y privados: renglones de trabajo, registro de avance físico por el supervisor, estimaciones y pagos, control por estado y alertas. Se describe en la sección 2.
- **Escena animada de Antigua a color:** volcanes, La Merced, el Arco de Santa Catalina, la Catedral, casas de colores, edificios, una obra con grúa y la calle con tránsito. En modo claro se ve de día (sol, cielo azul y aves) y en modo oscuro de noche (luna, estrellas, ventanas y faroles encendidos, luces de los vehículos y lava en el Volcán de Fuego). Ya no tiene botones encima.
- **Idioma ES / EN:** dentro del engranaje (también en la vista de bienvenida). La elección queda guardada en el navegador. Se traduce la interfaz; los datos que escribe el usuario (nombres, renglones, materiales) y los archivos descargados quedan como están.
- **Botones y campos de 44 px o más** en pantallas táctiles, y texto de 16 px en los campos para que el teléfono no haga zoom al escribir.

### Cuentas de usuario

No hay administrador: **cada persona crea su propia cuenta** con su nombre, su **correo electrónico** y una contraseña (mínimo 6 caracteres), y después entra con su correo y su contraseña.

- **Sin conexión a Firebase** (como viene el archivo): las cuentas se guardan en ese navegador y la contraseña se guarda cifrada (SHA-256 con sal), nunca tal cual.
- **Con Firebase** (pegando la configuración al inicio de `proyecto_final_v3.html`, activando «Correo electrónico/contraseña» en Firebase Authentication y publicando `firestore.rules`): la cuenta sirve en cualquier dispositivo, Firebase envía un **correo para verificar** la cuenta antes de entrar y la opción **«¿Olvidaste tu contraseña?»** envía un enlace al correo para crear una nueva.
- **Foto de perfil:** se sube desde la computadora o desde la galería (o la cámara) del teléfono; se recorta en cuadrado y se reduce a 320 × 320 px.
- **Programa en Python:** al iniciar pide «Iniciar sesión» o «Crear cuenta nueva» con correo y contraseña; la contraseña se guarda como huella PBKDF2-SHA256 con sal en `cuentas_programa.json`. El menú 8, «Mi cuenta y contraseña», permite cambiarla.

### Memoria de cálculo en Excel

En **Calcular**, **Muros** y **Carreteras**, el botón **«Descargar Excel»** (en Python, la pregunta «¿Descargar la memoria de cálculo en Excel?») crea un archivo `.xlsx` con:

- **Dibujos al inicio:** la **sección típica** con sus cotas y el acero, y una **vista isométrica** que muestra de dónde sale el volumen (por ejemplo, «Largo × Ancho × Espesor = 5 × 4 × 0.12 = 2.400 m³»). En muros y carreteras el isométrico se arma con la misma sección, repetida a lo largo del tramo.
- **Datos de entrada** en celdas amarillas: si se cambian, Excel recalcula todo.
- **Constantes:** desperdicio, factor de volumen seco, volumen y peso del saco, largo de la varilla.
- **Cálculo paso a paso con fórmulas de Excel** (`ROUNDUP`, `SUM`, `SUMIF`…) y una columna **«Fórmula en Excel»** que muestra la fórmula de cada resultado.
- En muros y carreteras: revisión normativa, estabilidad, capacidad hidráulica, mezclas por parte, acero por diámetro, resumen de materiales y otros materiales.

En los cálculos de **Calcular** todo es fórmula, desde las medidas hasta las varillas. En **Muros y Carreteras** el volumen de cada parte va como dato (lo calcula la geometría del programa) y de ahí en adelante todo es fórmula. El muro de gaviones no lleva concreto ni acero, por eso su Excel trae dibujos y cantidades, sin fórmulas de mezcla. Los dibujos se ven en Excel 2016 o más reciente (en el Excel de Python van en SVG con una copia PNG sin textos para programas más antiguos).

## 2. Módulos

1. **Inventario de bodega:** entradas, salidas, correcciones de conteo y kardex por material, con precio y mínimo de existencia. Es **independiente de los cálculos**: ningún cálculo lo lee ni descuenta material de la bodega. El material se elige de una lista desplegable agrupada por tipo (primero los que ya están en bodega; «Otro material» permite escribir uno nuevo), y la unidad de medida se elige de una lista que cambia según el tipo de material (ver sección 3).
2. **Elementos de concreto reforzado:** losa, viga, columna, zapata, muro, cimiento corrido, solera y mocheta. Calcula cemento, arena, piedrín, agua y hierro, con el grado de hierro según la medida de la varilla.
3. **Muros:** cinco calculadoras en la misma vista (en Python, menú 3):
   - **Muro de block:** block, mortero de pegado, cal, concreto fluido en celdas, soleras y acero (AGIES NSE 7.4, NTG 41054).
   - **Muro de contención de concreto reforzado (voladizo):** pantalla y zapata. Calcula el concreto y el acero de la pantalla, la punta y el talón por resistencia (ACI 318), y revisa el cortante en la base de la pantalla.
   - **Muro de contención de concreto ciclópeo:** muro de gravedad con 30 a 50 % de piedra bola. Calcula el concreto y la piedra bola.
   - **Muro de gaviones:** cajas de 2 × 1 m escalonadas al frente. Calcula cajas por hilada, malla, piedra de relleno, alambre galvanizado y geotextil.
   - **Muro de mampostería de piedra:** muro de gravedad con 65 a 75 % de piedra y mortero 1 : 3 o 1 : 4 de cemento y arena.

   En los cuatro muros de contención se piden los datos del relleno y del suelo (peso unitario, ángulo de fricción, sobrecarga, desplante, fricción base-suelo, capacidad soporte y coeficiente sísmico kh) y se revisa la estabilidad con y sin sismo: volteo, deslizamiento, excentricidad y presión en el suelo (sección 4). Cada uno tiene dibujo de la sección, memoria de cálculo descargable (HTML y Excel con fórmulas y dibujos), lloraderos, filtro de grava y excavación.
4. **Carreteras y drenaje:** pavimento rígido, bordillo fundido, cuneta tipo L, cuneta tipo V y canal, con la capacidad hidráulica de cunetas y canales.
5. **Normativa:** tablas de acero, cemento, concreto, block, mortero, tipos de muro y criterios de pavimento rígido.
6. **Obras y cálculos guardados:** suma de materiales por obra y exportación a CSV. Para quitar un cálculo, el botón «Quitar» pregunta «¿Quitar este cálculo?» con los botones **Sí** y **No**; solo «Sí» lo borra, y «No» o la tecla Esc cancelan.
7. **Gestión de proyectos** (solo en la página web; se guarda en el navegador, sin servidor):
   - **Tipo de proyecto Público o Privado.** Datos comunes: nombre del proyecto, constructora, representante legal de la empresa, residente responsable de la empresa, supervisor de obras, ubicación, monto, plazo y descripción.
     - **Público:** entidad contratante, número de contrato, número de operación Guatecompras (NOG, con enlace a la consulta), No. de acta de inicio, número de convenio, código SNIP, fuente de financiamiento, modalidad de compra y anticipo.
     - **Privado:** propietario y NIT, número de contrato o cotización aprobada, licencia de construcción municipal, modalidad de contratación (suma alzada, precios unitarios, administración o llave en mano, cada una con su recomendación de control), anticipo y retención de garantía.
   - **Estados:** Planificado, En gestión, En ejecución, Suspendido y Finalizado. El botón «Control del estado» de cada proyecto muestra:
     - una lista de verificación por estado, distinta para proyectos públicos y privados (tabla abajo);
     - datos de control del estado actual: avance físico, monto pagado o cobrado, número y fecha de la última estimación, fecha y motivo de la suspensión, fecha de recepción, monto liquidado y garantía;
     - barras de avance físico, avance financiero y tiempo transcurrido (los días suspendidos se descuentan del plazo);
     - fianzas de la Ley de Contrataciones del Estado en quetzales (sostenimiento de oferta 1 a 5 %, cumplimiento 10 %, anticipo 100 % del anticipo, conservación de obra 15 % por 18 meses, saldos deudores 5 %) o, en privados, anticipo, retención y garantía de cumplimiento;
     - historial de cambios de estado con fecha y nota.
   - **Aviso al cambiar de estado:** lista lo que falta (puntos sin marcar en etapas anteriores, contrato, NOG, acta de inicio, supervisor, licencia) y pide confirmar.
   - **Alertas de control:** plazo vencido o por vencer, atraso del avance físico contra el tiempo, más pagado que ejecutado, estimaciones atrasadas, suspensión de más de 30 días o reanudación vencida, fianza de conservación por vencer, liquidación pendiente después de 90 días y garantía de obra vigente.
   - **Indicadores y gráficos:** proyectos por estado, avance físico contra lo programado por proyecto y «Avance físico al día» (cada proyecto en ejecución con su último registro y un botón para registrar avance). La lista se filtra por tipo, estado y búsqueda (nombre, NOG, contrato) y se descarga en CSV.
   - **Botones en cada proyecto:** «Registrar renglones de trabajo» y «Registrar avance físico», además de «Abrir expediente», «Editar» y «Eliminar».
   - **Registrar avance físico** (reemplaza el antiguo formulario de tareas): se elige el proyecto y el **supervisor de obras** (el municipal o de la entidad en proyectos públicos; el del propietario en privados) anota, renglón por renglón, solo la **cantidad ejecutada** en la unidad del renglón, con fecha y observaciones. Mientras escribe, la página calcula sola el ejecutado acumulado, el % del renglón y el avance físico del proyecto. Debajo está la lista de registros de avance físico de todos los proyectos. El avance físico del proyecto se pondera con el costo de cada renglón.
   - **Expediente del proyecto** (botón «Abrir expediente» en cada tarjeta), con ocho pestañas:
     - **Resumen:** avance físico, avance financiero, programado a hoy, días por vencer, monto vigente, pagado, saldo y anticipo por amortizar; **curva S** (programado según el cronograma por renglón contra físico y financiero acumulados, con tabla de datos); alertas y todos los datos del proyecto.
     - **Renglones de trabajo:** tabla como la del presupuesto oficial: No., descripción, cantidad, unidad, precio unitario, costo total y **% del contrato** (costo del renglón ÷ valor del contrato × 100, con dos decimales), agrupada por módulos con subtotales, total y **monto en letras**. Incluye la plantilla «Escuela: aulas, dirección, bodega, cocina, mobiliario y baños» con 78 renglones (Q 1,700,000.00), y al pegar desde Excel reconoce las filas de módulo. Se cargan los renglones típicos de una tipología (edificación, pavimentación, carretera, agua potable, drenaje sanitario, muro de contención o puente), se agregan uno por uno o se pegan desde Excel (con inicio y fin programados opcionales). Meta física del proyecto (por ejemplo, m² de pavimento).
       - **Cronograma por renglón:** cada renglón lleva su inicio y fin programados; el botón «Proponer fechas del cronograma» los reparte en orden y con traslape dentro del plazo para ajustarlos después. Un diagrama de barras (Gantt) muestra el plazo de cada renglón, su avance ejecutado y lo programado a hoy, y marca los atrasados.
     - **Avance físico:** el mismo registro por renglón y el historial del proyecto, con el avance acumulado después de cada registro.
     - **Estimaciones y pagos:** anticipo (fecha de pago, amortizado y por amortizar) y estimaciones con la cantidad medida de cada renglón; calcula bruto, amortización del anticipo, retención (privados) y líquido. Cada estimación pasa por Presentada, Aprobada, Pagada o Rechazada, con fecha de pago y número de CUR (públicos) o factura FEL (privados). Avisa si una cantidad pasa de lo contratado.
     - **Modificaciones:** órdenes de cambio, de trabajo suplementario, acuerdos de trabajo extra, ampliaciones de plazo y reducciones; ajustan el monto vigente y la fecha de fin vigente.
     - **Control del estado:** la lista de verificación por estado, los datos de control, las fianzas o garantías y el historial.
     - **Cotejo oficial:** en públicos se anotan los valores del SNIP (SEGEPLAN), SICOIN/SIGES (MINFIN), Guatecompras y el Consejo de Desarrollo (CODEDE/COMUDE), con enlaces a cada sistema, y se comparan con este control (tolerancia 3 puntos en % y 0.5 % en montos). En privados se coteja con lo facturado en FEL, los depósitos del banco y el estado de cuenta del propietario, y se vigila el vencimiento de la licencia. La página no se conecta en automático a esos sistemas: no tienen un servicio público de consulta y piden usuario institucional.
     - **Bitácora:** anotaciones con fecha, folio y autor.
   - Las tareas de la versión anterior se reemplazaron por el registro de avance físico. Los proyectos guardados con la versión anterior se convierten solos: «En curso» pasa a «En ejecución», «En pausa» a «Suspendido» y «Completado» a «Finalizado»; quedan como privados y el responsable pasa a ser el residente.

### Control por estado

| Estado | Proyecto público | Proyecto privado |
|---|---|---|
| Planificado | Preinversión, registro SNIP, planos y especificaciones, presupuesto y cronograma, estudio de suelos y topografía, evaluación ambiental (MARN), disponibilidad presupuestaria, convenio de financiamiento | Topografía y suelos, diseño aprobado por el propietario, presupuesto y cronograma, uso del suelo o alineación municipal, solicitud de licencia, evaluación ambiental si aplica |
| En gestión | Bases en Guatecompras (NOG), junta de cotización o licitación, ofertas con fianza de sostenimiento, adjudicación, contrato aprobado, fianzas de cumplimiento y anticipo, registro en la Contraloría | Cotización aceptada, contrato privado con firmas legalizadas, licencia otorgada, anticipo, seguro de obra, IGSS, servicios provisionales |
| En ejecución | Acta de inicio, bitácora autorizada por la Contraloría, anticipo, supervisor nombrado, estimaciones al día, ensayos de laboratorio, informes de supervisión, órdenes de cambio aprobadas, fianzas vigentes | Bitácora, avances cobrados, órdenes de cambio firmadas, ensayos de calidad, planilla e IGSS, informe semanal con fotos, bodega y kardex, inspecciones municipales |
| Suspendido | Acta de suspensión con motivo, aviso a la afianzadora, inventario de obra y materiales, solicitud de prórroga, protección de la obra, acta de reanudación | Aviso por escrito al propietario, constancia de obra ejecutada, cobro de lo ejecutado, protección de la obra, acuerdo de reanudación |
| Finalizado | Aviso de finalización, comisión receptora, acta de recepción, fianzas de conservación y saldos deudores, comisión liquidadora y acta de liquidación, planos finales, finiquito y cierre en Guatecompras | Pruebas de instalaciones, pendientes resueltos, acta de entrega firmada, liquidación y devolución de la retención, carta de garantía, planos y manuales, aviso de finalización a la municipalidad |

Los porcentajes de fianzas y plazos son de referencia (Decreto 57-92 y sus reformas); confírmalos con las bases del evento y el contrato.

## 3. Datos de cálculo

| Constante | Valor |
|---|---|
| Desperdicio | 5 % |
| Factor de volumen seco (concreto) | 1.54 |
| Factor de volumen seco (mortero) | 1.25 |
| Saco de cemento | 42.5 kg = 0.0283 m³ |
| Largo comercial de varilla | 6 m |
| Block por m² (sisa de 1 cm) | 12.5 |
| Peso unitario: concreto reforzado / ciclópeo / mampostería de piedra | 2.4 / 2.3 / 2.2 t/m³ |
| Gavión (piedra con 30 % de porosidad) | 0.70 × peso de la piedra |
| Vacíos de la piedra suelta | 35 % (piedra suelta = volumen sólido / 0.65) |
| Malla de gavión / alambre de amarre | 1.25 kg/m² / 8 % del peso de la malla |

### Unidades de medida por tipo de material

Al elegir el material, la página (y el programa de consola) ofrece las unidades que se usan para ese tipo de material. Siempre queda la opción «Otra unidad» para escribir una distinta.

| Tipo de material | Se reconoce por | Unidades disponibles |
|---|---|---|
| Cemento | cemento | sacos de 42.5 kg, kg, lb, quintales, toneladas |
| Cal y aglomerantes | cal, yeso, pegablock | kg, lb, quintales, sacos o bolsas, toneladas |
| Acero de refuerzo | varilla, hierro, acero, pasajunta | varillas de 6 m, quintales, kg, lb, toneladas, metros lineales |
| Alambre y clavos | alambre, clavo, tornillo | lb, kg, quintales, rollos, cajas |
| Agregados | arena, piedrín, grava, selecto, base granular | m³, yd³, pie³ |
| Block, ladrillo y adoquín | block, ladrillo, adoquín | unidades, cientos, millares, m² |
| Láminas, mallas y paneles | lámina, tablayeso, malla, geotextil | hojas, m², rollos, piezas, metros |
| Madera y formaleta | madera, tabla, regla, parral, plywood | piezas, pies tablares, varas, metros, hojas |
| Tubería | tubo, tubería, PVC | tubos, metros, pies, piezas |
| Líquidos | pintura, sellador, aditivo, membrana | galones, litros, cubetas de 5 gal, toneles de 55 gal |
| Otros | cualquier otro nombre | piezas, unidades, kg, lb, metros, m², m³, litros, galones, rollos, cajas, juegos, global |

**Conversiones.** Si un material ya existe en bodega y la entrada se registra en otra unidad, la cantidad se convierte a la unidad de bodega antes de sumarla (por ejemplo, 2 quintales de cemento = 2.13 sacos). Si no hay forma de convertir, el programa avisa y no registra la entrada.

| Equivalencia | Valor |
|---|---|
| 1 quintal | 100 lb = 45.36 kg |
| 1 libra | 0.4536 kg |
| 1 tonelada | 1,000 kg |
| 1 yd³ | 0.7646 m³ |
| 1 pie³ | 0.0283 m³ |
| 1 galón | 3.785 litros |
| 1 pie | 0.3048 m |
| 1 ciento / 1 millar | 100 / 1,000 unidades |
| 1 saco de cemento | 42.5 kg |
| 1 varilla | 6 m; su peso depende del número (tabla de acero) |

### Concreto según dosificación

| f'c | Cemento : arena : piedrín | Agua/cemento | Uso recomendado |
|---|---|---|---|
| 2,000 psi (140 kg/cm²) | 1 : 3 : 5 | 0.70 | Cimiento ciclópeo, banquetas, fundiciones sin carga |
| 2,500 psi (175 kg/cm²) | 1 : 2 : 4 | 0.62 | Bordillos, cunetas, soleras y mochetas de vivienda |
| 3,000 psi (210 kg/cm²) | 1 : 2 : 3 | 0.55 | Losas, vigas, columnas, zapatas, canales |
| 4,000 psi (280 kg/cm²) | 1 : 1.5 : 2.5 | 0.46 | Pavimento rígido (MR ≈ 4.5 MPa), elementos muy cargados |

## 4. Fórmulas principales

**Concreto**

- Volumen con desperdicio: `Vt = V × 1.05`
- Sacos de cemento: `Sacos = 1.54 × c / (c + a + p) × Vt / 0.0283`
- Arena: `1.54 × a / (c + a + p) × Vt` · Piedrín: `1.54 × p / (c + a + p) × Vt`
- Agua: `Sacos × 42.5 × (a/c)`
- Hierro de elementos: `Vt × cuantía (kg/m³)`; varillas = `kg / (peso por metro × 6)`

**Muro de block**

- Área de block: `A = L × H − vanos − soleras × L × 0.20`
- Block: `N = A × 12.5 × 1.05`
- Mortero: `V = A × 7.375 m de junta/m² × 0.01 m × espesor × 1.10`
- Celdas con refuerzo: `n = L / s + 1`; concreto fluido = `n × altura × área de celda`

**Avance físico y financiero de proyectos**

- Monto de cada renglón: `Mr = cantidad × precio unitario`; peso: `Pr = Mr / Σ Mr`; % del contrato: `Mr / valor del contrato × 100` (dos decimales)
- El avance físico sale de los registros de avance físico; si no hay, de las estimaciones
- Avance físico ponderado: `AF = Σ min(1, ejecutado / contratado) × Pr × 100`. Sin renglones: `ejecutado de la meta / meta × 100`; sin meta: valor manual
- Estimación: `bruto = Σ cantidad medida × P.U.`; `amortización = bruto × % anticipo`; `retención = bruto × % retención` (privados); `líquido = bruto − amortización − retención`
- Monto vigente: `monto original + Σ modificaciones`; fin vigente: `fin original + Σ días de ampliación`
- Avance financiero: `(anticipo pagado + Σ líquido de estimaciones pagadas) / monto vigente × 100`
- Tiempo transcurrido: `(días desde el inicio − días suspendidos) / días del plazo vigente × 100`
- Programado de un renglón a una fecha: `f = (fecha − inicio) / (fin − inicio)`, entre 0 y 1
- Avance programado del proyecto: `AP = Σ f × Pr × 100` (curva S programada). Sin cronograma: línea recta entre el inicio y el fin vigente; los renglones sin fechas usan el plazo del proyecto
- Alerta de atraso si lo programado supera al avance físico en más de 15 puntos (también por renglón); alerta de sobrepago si lo pagado (sin el anticipo por amortizar) supera al avance físico en más de 10 puntos

**Muros de contención (por metro de muro)**

- Empuje activo de Rankine: `Ka = tan²(45° − φ/2)`; `Pa = ½ Ka γ H² + Ka q H`
- Empuje pasivo frente a la base: `Pp = ½ Kp γ Df²`, con `Kp = 1 / Ka`
- Sismo (Mononobe-Okabe): `θ = arctan kh`; `KAE = cos²(φ − θ) / (cos θ [1 + √(sen φ sen(φ − θ) / cos θ)]²)`; `ΔPAE = ½ γ H² (KAE − Ka)` a 0.6 H, más la inercia del muro `kh × ΣW`
- Volteo: `FSv = ΣMr / Ma ≥ 2.0` (≥ 1.5 con sismo)
- Deslizamiento: `FSd = (μ ΣW + Pp) / Pa ≥ 1.5` (≥ 1.1 con sismo)
- Excentricidad: `x = (ΣMr − Ma) / ΣW`, `e = B/2 − x ≤ B/6` (≤ B/4 con sismo)
- Presión en el suelo: `qmáx = ΣW / B × (1 + 6e / B) ≤ qa` (≤ 1.33 qa con sismo)
- Pantalla en voladizo: `Mu = 1.6 (Ka γ hp³ / 6 + Ka q hp² / 2) + 1.0 ME`; acero por resistencia con `As mín = 0.0015 b t` (vertical), `0.0020 b t` (horizontal) y `0.0018 b h` (zapata); cortante `φVc = 0.75 × 0.53 √f'c b d`
- Gravedad (ciclópeo o mampostería): `V = (Bc hc + (B + c)/2 × h) × L`; piedra suelta = `V × % piedra / 0.65 × 1.05`
- Gaviones: cajas por hilada = `⌈L / 2⌉ × ancho`; malla por caja = `2 (2·1 + 2h + 1h) + h`; piedra = `0.70 V / 0.65 × 1.05`
- Lloraderos PVC Ø 3 pulg: una fila a cada 1.5 m de alto y uno a cada 1.5 m de largo

**Pavimento rígido**

- Volumen: `V = L × A × e`
- Juntas transversales: `⌈L / sj⌉ − 1`; separación máxima `≤ 24 × e` y `≤ 4.5 m`
- Pasajuntas lisas: diámetro ≈ `e / 8`, 0.45 m a cada 0.30 m
- Barras de amarre: #4 de 0.75 m a cada 0.75 m en la junta longitudinal

**Bordillo y cunetas**

- Bordillo trapezoidal: `A = (b1 + b2) / 2 × h`
- Cuneta tipo L: `A = a × (t + a·Sx/2) + b × h`; capacidad por Izzard (HEC-22): `Q = 0.376 / n × Sx^(5/3) × S^(1/2) × T^(8/3)`
- Cuneta tipo V: `ℓ = √((B/2)² + h²)`; concreto = `2 × ℓ × t × L`

**Canal y Manning**

- `Q = (1/n) × A × R^(2/3) × S^(1/2)`, con `R = A / P`
- Froude: `Fr = V / √(g × A / T)`; subcrítico si `Fr < 1`
- Velocidad aceptable en concreto: entre 0.6 m/s (no sedimenta) y 4.5 m/s (no erosiona)

## 5. Normativa de referencia

### Acero de refuerzo (COGUANOR NTG 36011 / ASTM A615)

| Varilla | Pulg | Ø mm | Área cm² | kg/m | Grado usual | Uso típico |
|---|---|---|---|---|---|---|
| #2 | 1/4 | 6.35 | 0.32 | 0.249 | 40 | Estribos, eslabones, amarres |
| #3 | 3/8 | 9.53 | 0.71 | 0.560 | 40 o 60 | Losas, soleras, mochetas, muros de block |
| #4 | 1/2 | 12.70 | 1.27 | 0.994 | 40 o 60 | Vigas, columnas, zapatas, barras de amarre |
| #5 | 5/8 | 15.88 | 1.98 | 1.552 | 60 | Vigas y columnas de mayor carga |
| #6 | 3/4 | 19.05 | 2.85 | 2.235 | 60 | Edificios; pasajuntas de losas de 15 cm |
| #7 | 7/8 | 22.23 | 3.88 | 3.042 | 60 | Edificios y puentes; pasajuntas de 18 cm |
| #8 | 1 | 25.40 | 5.07 | 3.973 | 60 | Edificios y puentes; pasajuntas de 20 a 25 cm |

- **Grado 40:** fy = 2,810 kg/cm² (40,000 psi). Vivienda de 1 y 2 niveles (NSE 4).
- **Grado 60:** fy = 4,200 kg/cm² (60,000 psi). Marcos sismorresistentes (NSE 7.1), edificios, puentes y obras viales.
- **Hierro comercial:** no normado; no usar en elementos estructurales.

### Cemento (NTG 41095 / ASTM C1157)

| Cemento | Tipo | Resistencia a 28 días | Uso |
|---|---|---|---|
| Uso general (UGC) | GU | 4,060 psi (28 MPa) | Concreto de losas, vigas, columnas, pavimentos |
| Estructural / alta resistencia | GU o HE | 5,000 psi (35 MPa) | Alta resistencia, prefabricados |
| Resistente a sulfatos | MS / HS | según fabricante | Cunetas, canales, alcantarillas, suelos agresivos |
| Cemento de mampostería | ASTM C91 | menor | Solo morteros y repellos |

Los 4,060 psi son la resistencia del cemento en cubos de mortero de laboratorio, no la del concreto.

### Block de concreto (COGUANOR NTG 41054)

| Clase | f'm promedio (área neta) | Mínimo individual | Absorción | Marca | Uso |
|---|---|---|---|---|---|
| A | 133 kg/cm² (13.0 MPa) | 113 kg/cm² | ≤ 10 % | Azul | Muros de carga, mampostería reforzada |
| B | 100 kg/cm² (9.8 MPa) | 85 kg/cm² | ≤ 15 % | Rojo | Vivienda de un nivel con carga moderada |
| C | 66 kg/cm² (6.5 MPa) | 56 kg/cm² | ≤ 20 % | Verde | Muros divisorios y cercos sin carga |

Medidas reales: 9×19×39, 14×19×39 y 19×19×39 cm.

### Mortero de pegado (AGIES NSE 7.4 / ASTM C270)

| Tipo | Cemento : cal : arena | Resistencia mínima | Uso |
|---|---|---|---|
| M | 1 : ¼ : 3 | 175 kg/cm² | Muros de carga, cimientos, bajo el nivel del suelo |
| S | 1 : ½ : 4½ | 125 kg/cm² | Mampostería reforzada en general |
| N | 1 : 1 : 6 | 50 kg/cm² | Muros interiores sin carga y tabiques |

### Tipos de muros que se construyen en Guatemala

| Tipo | Norma | Descripción |
|---|---|---|
| Mampostería reforzada de block | AGIES NSE 7.4 | Block clase A, refuerzo vertical en celdas rellenas (máx. a cada 1.20 m) y soleras de humedad, intermedia y corona |
| Mampostería confinada | AGIES NSE 4 y NSE 7.4 | Block o ladrillo confinado por columnas (mochetas) y soleras de concreto |
| Ladrillo de barro cocido | NSE 7.4 | Ladrillo tayuyo o perforado, confinado o reforzado |
| Muro de concreto reforzado | AGIES NSE 7.1 (ACI 318) | Muros de corte, sótanos, tanques y cisternas |
| Muros de contención | AGIES NSE 2.1 | En voladizo de concreto reforzado (B ≈ 0.5 a 0.7 H), de gravedad de concreto ciclópeo o de mampostería de piedra (base 0.5 a 0.7 H, corona ≥ 0.30 a 0.40 m), o de gaviones (base 0.5 a 0.8 H); requieren estudio de suelos, revisión sísmica y drenaje (lloraderos y filtro) |
| Muros no estructurales | NTG 41055 | Tabiques de block de 9 cm, tablayeso o paneles |
| Adobe y bajareque | No recomendado en obra nueva | Sistemas de tierra muy vulnerables en sismos |

### Pavimento rígido y drenaje (DGC 2001, AASHTO 93)

| Concepto | Valor de referencia |
|---|---|
| Concreto para losa | MR ≥ 4.5 MPa (≈ f'c 4,000 psi) |
| Espesor de losa | 15 cm mínimo en calles; 18 a 25 cm en carreteras |
| Juntas transversales | ≤ 24 × espesor y ≤ 4.5 m; losas con largo/ancho ≤ 1.25 |
| Pasajuntas | Varilla lisa Ø ≈ e/8, 45 cm, a cada 30 cm |
| Barras de amarre | #4 corrugada de 75 cm a cada 75 cm |
| Base granular | 15 cm compactada al 95 % Proctor modificado |
| Bordillo fundido | 15 cm arriba, 20 cm abajo, 45 cm de alto; juntas a cada 3 m |
| Cunetas y canales | 2,500 a 3,000 psi, 10 a 15 cm de espesor, juntas a cada 3 m |
| n de Manning | Concreto 0.013 a 0.017; mampostería de piedra 0.025; tierra 0.025 a 0.030 |

## 6. Referencias

- AGIES, Normas de Seguridad Estructural NSE 2018: NSE 2.1, NSE 4, NSE 7.1 y NSE 7.4.
- ACI 318, requisitos de reglamento para concreto estructural (flexión, cortante y acero mínimo de muros y zapatas).
- Mononobe-Okabe, empuje sísmico de tierras sobre muros de contención.
- Congreso de la República de Guatemala, Ley de Contrataciones del Estado (Decreto 57-92) y sus reformas; Guatecompras.
- SEGEPLAN, Sistema Nacional de Inversión Pública (SNIP); Ministerio de Finanzas Públicas, SICOIN y SIGES.
- COGUANOR NTG 36011 (ASTM A615), barras de acero corrugado para refuerzo de concreto.
- COGUANOR NTG 41095 (ASTM C1157), cementos hidráulicos, especificaciones por desempeño.
- COGUANOR NTG 41054 y NTG 41055, bloques huecos de concreto para muros.
- Dirección General de Caminos, *Especificaciones generales para construcción de carreteras y puentes* (2001).
- AASHTO, *Guía para el diseño de estructuras de pavimento* (1993); FHWA HEC-22, drenaje urbano.
