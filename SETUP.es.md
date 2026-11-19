# Guía de instalación

Esta guía te lleva desde cero hasta tener la hoja de cálculo y el bot de Telegram funcionando. Está pensada para personas que nunca han usado Apps Script. *In English? Read [SETUP.md](SETUP.md).*

**¿Solo quieres la hoja de cálculo, sin bot?** Haz el Paso 1 y el Paso 7 y listo. El bot es opcional.

**¿Quieres probarlo antes?** Sube uno de los archivos de `demo/` (datos inventados) igual que en el Paso 1 y échale un vistazo.

## Qué necesitas

Unos 30 minutos, una sola vez, en un ordenador. Después todo se hace desde Telegram en el móvil. Sin programar, sin servidor y sin coste.

- **Una cuenta de Google** (una cuenta normal de Gmail sirve; algunas cuentas de trabajo o de centros educativos bloquean los scripts que hacen falta).
- **Telegram** en el móvil.
- **Un ordenador con navegador** para la instalación (Chrome funciona mejor). La app de Hojas de cálculo del móvil no permite hacer estos pasos.
- **Los dos archivos** de este repositorio: la plantilla (en español `templates/Presupuesto_Familiar.xlsx` o en inglés `templates/Family_Budget.xlsx`) y el archivo del bot `bot/FamilyBudgetBot.gs`. Descárgalos desde la última [Release](../../releases) o con **Code ▸ Download ZIP**.

Los pasos usan los nombres de la plantilla en español. En la plantilla inglesa los mismos menús salen en inglés (por ejemplo *Family Budget* en lugar de *Presupuesto familiar*).

## Paso 1: Sube la hoja de cálculo a Google Drive

La hoja tiene que convertirse en una hoja de Google de verdad, no en un archivo de Excel, o los siguientes pasos no aparecerán.

1. Entra en [drive.google.com](https://drive.google.com), haz clic en **+ Nuevo ▸ Subir archivo** y elige la plantilla.
2. Haz doble clic en el archivo subido para abrirlo.
3. Haz clic en **Archivo ▸ Guardar como Hojas de cálculo de Google**. Se abre una copia nueva en otra pestaña.
4. Mira arriba en la copia nueva: junto al nombre del archivo **no debe aparecer la etiqueta verde "XLSX"**. A partir de ahora usa solo esta copia; puedes borrar la de Excel.
5. Haz clic en **Archivo ▸ Configuración** y elige:
   - **Configuración regional**: tu país (por ejemplo España o Argentina).
   - **Zona horaria**: la tuya.
6. Abre la pestaña **Movimientos** y borra la fila 2 (la fila gris de ejemplo).

### Atajo: enlace "Hacer una copia"

Si la página del proyecto ofrece un enlace **Hacer una copia**, haz clic con tu cuenta de Google abierta. Obtienes tu propia hoja de Google con el código del bot ya dentro, así que puedes saltarte los puntos 1–4 del Paso 1 y los puntos 1–4 del Paso 3. Sigue con Archivo ▸ Configuración (Paso 1, punto 5) y luego el Paso 2.

## Paso 2: Crea tu bot de Telegram

El ayudante oficial de Telegram, BotFather, crea el bot por ti y te da su contraseña (el "token").

1. En Telegram, busca **@BotFather** (tiene una marca azul de verificado) y pulsa **Iniciar**.
2. Envía `/newbot`.
3. Escribe un **nombre** que verá la gente, por ejemplo *Presupuesto familiar*.
4. Escribe un **nombre de usuario** que termine en `bot`, por ejemplo *familia_garcia_bot*. Si está ocupado, prueba otro.
5. BotFather responde con un **token** largo, del tipo `123456789:AAH...xyz`. Cópialo y guárdalo en un sitio seguro. **No lo compartas nunca**: quien lo tenga controla el bot.
6. Opcional pero recomendado: envía `/setjoingroups`, elige tu bot y luego **Disable**, para que nadie pueda añadirlo a un grupo.

## Paso 3: Pega el código del bot en la hoja

El bot vive dentro de tu hoja, así que abre el editor de código **desde la propia hoja**, no desde Google Drive.

1. En tu hoja de Google, haz clic en **Extensiones ▸ Apps Script**. Se abre una pestaña nueva con un poco de código.
2. En tu ordenador, abre *FamilyBudgetBot.gs* con un editor de texto (en Mac: clic derecho ▸ Abrir con ▸ TextEdit; en Windows: Bloc de notas). Selecciónalo todo y cópialo.
3. Vuelve a Apps Script, haz clic dentro del código, selecciónalo todo y bórralo. Pega el código del bot.
4. Haz clic en el icono de **guardar** (o pulsa Ctrl+S / ⌘S).
5. Haz clic en el **engranaje** (Configuración del proyecto) a la izquierda y pon la misma **zona horaria** que en el Paso 1.
6. Opcional: haz clic en "Proyecto sin título" arriba y ponle un nombre, por ejemplo *Bot del presupuesto*.

¿No ves el menú Extensiones? El archivo sigue siendo la versión de Excel: vuelve al Paso 1, punto 3.

## Paso 4: Publícalo

Al publicarlo, el bot recibe una dirección web privada a la que Telegram envía lo que pulsas.

1. En Apps Script, haz clic en el botón azul **Implementar ▸ Nueva implementación**.
2. Haz clic en el engranaje junto a *Seleccionar tipo* y elige **Aplicación web**.
3. Pon **Ejecutar como: Yo** y **Quién tiene acceso: Cualquier usuario**.
4. Haz clic en **Implementar**, luego en **Autorizar acceso** y elige tu cuenta de Google.
5. Si Google dice *"Google no ha verificado esta aplicación"*: haz clic en **Configuración avanzada ▸ Ir a … (no seguro)**. Es normal: es tu propio script, no una app de una tienda.
6. Marca **Seleccionar todo** y haz clic en **Continuar**. El primer permiso debe hablar solo de las hojas de cálculo en las que está instalado este script, no de todas.
7. Copia la **URL de la aplicación web**. Termina en `/exec`.

**¿"Cualquier usuario" abre mi hoja a cualquiera?** No. Solo deja que Telegram llegue al bot sin iniciar sesión en Google. Tu hoja sigue siendo privada, la dirección solo funciona con una clave secreta que el bot añade solo, y el bot ignora a cualquiera que no esté en tu lista.

## Paso 5: Conecta el bot y activa la automatización

1. Vuelve a la pestaña de la hoja y **recarga la página**.
2. Espera unos segundos: aparece un menú nuevo **Presupuesto familiar** al final de la barra de menús, después de *Ayuda*.
3. Haz clic en **Presupuesto familiar ▸ 1. Conectar bot de Telegram…**. Pega el **token** del Paso 2 y después la **URL** del Paso 4.
4. Deberías ver *"Conectado a @tu_bot"*.
5. Haz clic en **Presupuesto familiar ▸ 2. Activar automatización diaria**. Desde ese momento, cada mañana hacia las 7:00 el bot añade tus pagos recurrentes y envía los avisos de facturas.
6. Opcional, si pagas en más de una moneda: **Presupuesto familiar ▸ 3. Preparar monedas, subcategorías y calendarios**.

## Paso 6: Entra tú e invita a tu familia

El bot solo responde a las personas de tu lista; los demás no reciben nada.

**Tú (la primera vez)**

1. En Telegram, abre tu bot y envía `/start`. Te responde con tu **ID de Telegram**, un número del tipo *123456789*.
2. En la hoja, abre la pestaña **Configuración**. Busca **IDs de Telegram con acceso total** y pega tu número en la celda amarilla de su derecha.
3. Envía `/start` otra vez. Aparecen los botones del menú.

**Todos los demás**

1. En la hoja: **Presupuesto familiar ▸ 4. Palabra secreta para nuevos usuarios…** y escribe una palabra, por ejemplo *mango42*.
2. Dile esa palabra en privado a la persona que quieres añadir. Ella abre el bot y la envía.
3. Recibes un mensaje de Telegram con su nombre y dos botones: **Acceso total** (puede añadir movimientos) o **Solo lectura** (solo puede consultar). Un toque y queda añadida.
4. Envía `/start` y ya está dentro.

Cada persona elige su idioma con `/english` o `/espanol`. Para quitar a alguien, borra su número en Configuración.

### La lista de acceso: quién puede usar el bot

El bot solo obedece a los números de dos celdas de la pestaña **Configuración**. Quien no esté no recibe ninguna respuesta, ni siquiera un error.

| Celda de Configuración | Quién va ahí | Qué puede hacer |
| --- | --- | --- |
| **IDs de Telegram con acceso total** | Tú y quien vaya a añadir movimientos | Todo: añadir, deshacer, marcar facturas pagadas, añadir personas nuevas, exportar (si está activado) |
| **IDs de Telegram solo lectura** | Abuelos, una gestoría… | Buscar, Últimos, Este mes, Próximos y avisos |

- Escribe solo los números, separados por comas: `123456789, 987654321`. El número va en la celda amarilla a la derecha de la etiqueta.
- Los cambios funcionan al momento; no hay que volver a publicar nada.
- Mientras la primera lista esté vacía, el bot enseña su ID a quien le escriba, para que puedas configurarlo. Rellénala enseguida.

**Bloquea las dos celdas** (si otras personas pueden editar la hoja)

1. En **Configuración**, selecciona las dos celdas amarillas junto a las etiquetas de IDs.
2. Haz clic en **Datos ▸ Proteger hojas e intervalos** y luego en **Establecer permisos**.
3. Elige **Solo tú** y haz clic en **Hecho**.

Ahora solo tú puedes cambiar quién tiene acceso. Los botones de "añadir" de Telegram siguen funcionando, porque el bot actúa en tu nombre.

## Paso 7: Rellena lo básico

Las celdas amarillas son las que rellenas tú; todo lo demás se calcula solo. Calcula unos 20 minutos.

| Pestaña | Qué hacer |
| --- | --- |
| **Configuración** | Cambia *Hijo/a 1*, *Adulto 1*… por nombres reales (columna *Para (persona)*). Ajusta las categorías: primero ingresos, debajo gastos; escribe las nuevas en filas amarillas vacías. |
| **Recurrentes** | Una fila por cada cosa que se repite: sueldos, alquiler, colegio, suscripciones, seguros. Pon el importe, cada cuánto y la fecha de un pago real. |
| **Facturas e impuestos** | Próximos vencimientos cuyo importe cambia (impuestos, excursiones). El bot te avisa 3 días antes. |

Deja los presupuestos mensuales vacíos los 2–3 primeros meses. Después la pestaña **Año** muestra lo que de verdad gastas de media, y pones presupuestos un poco por debajo.

No cambies el nombre de las pestañas ni de los títulos de columna, y no insertes ni borres filas en Configuración: vacía el nombre en su lugar.

## El día a día

Apunta un pago en cuanto lo haces: son unos 10 segundos. Ejemplo: pulsa **➖ Gasto ▸ Supermercado ▸ Hoy**, escribe `45,20`, pulsa **Saltar**, elige para quién es y quién pagó, y luego **✅ Guardar**.

| Botón | Qué hace |
| --- | --- |
| ➖ Gasto / ➕ Ingreso / 🧾 Impuesto | Añade un movimiento. Al elegir la categoría, el bot muestra los últimos movimientos de esa categoría, para que veas si ya está apuntado. |
| 🔎 Buscar | Encuentra movimientos pasados: escribe una palabra como *dentista*, si quieres con un mes o un año (*dentista 2026*). Los resultados muestran un total por persona. |
| 🕘 Últimos | Los 10 últimos movimientos, con para quién es cada uno y quién lo apuntó. Filtra por categoría o por persona. |
| 📊 Este mes | Ingresos, gastos y presupuesto usado en lo que va de mes. |
| 📅 Próximos | Facturas y pagos recurrentes de los próximos 30 días. |
| ↩️ Deshacer | Borra el movimiento que acabas de añadir. |
| 📤 Exportar | Opcional, desactivado por defecto. Envía un mes o un año como archivo CSV al chat. Mira abajo. |

- **Fechas:** pulsa Hoy, Ayer o el calendario. Si la escribes, pon primero el día: `25/10`.
- **Otras monedas:** escribe la moneda después del importe, por ejemplo `42000 ars`. Pon el tipo de cambio de cada mes en la pestaña *Tipos de cambio*.
- **Gráficos y análisis:** abre la hoja. *Panel* muestra el año, *Análisis* ordena tus categorías para cualquier mes.

### Exportar desde el móvil (opcional)

Desactivado por defecto. Para activarlo, abre **Configuración** y pon **Permitir exportar desde Telegram** en **Sí**. Entonces solo las personas de **IDs de Telegram con acceso total** ven un botón **📤 Exportar**; las de solo lectura, nunca. Eligen un mes o un año y el bot envía un archivo CSV que se abre en Excel o en Hojas de cálculo de Google. Vuelve a ponerlo en **No** y el botón desaparece en el siguiente /start.

Un archivo CSV se puede reenviar a cualquiera, así que actívalo solo si todas las personas con acceso total pueden sacar los datos de la hoja.

## Actualizar a una versión nueva

Cuando se publique una versión nueva de `bot/FamilyBudgetBot.gs` en este repositorio, tus datos, ajustes y lista de personas se quedan como están. Solo cambia el código.

1. Hoja ▸ **Extensiones ▸ Apps Script**.
2. Selecciona todo el código, bórralo, pega el archivo nuevo y guarda.
3. **Implementar ▸ Gestionar implementaciones**, haz clic en el **lápiz**, en *Versión* elige **Nueva versión** y luego **Implementar**. No uses *Nueva implementación*: cambiaría la dirección web.
4. Cada persona envía `/start` una vez en Telegram para ver los botones nuevos.

## Si algo falla

| Lo que ves | Qué hacer |
| --- | --- |
| No hay menú **Extensiones** | El archivo sigue siendo Excel (etiqueta verde "XLSX"). **Archivo ▸ Guardar como Hojas de cálculo de Google** y usa la copia nueva. En el móvil no existe el menú Extensiones: usa un ordenador. |
| Apps Script se abrió como un "Proyecto sin título" aparte | Se abrió desde Drive. Ciérralo y ábrelo otra vez desde la hoja: **Extensiones ▸ Apps Script**. |
| No sale el menú **Presupuesto familiar** al recargar | En Apps Script, elige **onOpen** en la lista de funciones junto a *Ejecutar*, haz clic en **Ejecutar**, permítelo y vuelve a la hoja. |
| El bot no contesta nada | **Presupuesto familiar ▸ Comprobar estado del bot**. Si ves "302" o mensajes pendientes: publica una versión nueva (*Actualizar*, paso 3) y vuelve a usar **1. Conectar bot de Telegram** con el mismo token y la misma URL. |
| El bot contesta con ⚠️ y un mensaje | Algo de la hoja no cuadra (a menudo una pestaña o columna renombrada). El mensaje dice qué. |
| *"Los datos que has introducido infringen las reglas de validación"* | **Presupuesto familiar ▸ Reparar los desplegables de Movimientos**. |
| Meses o números en otro idioma | **Archivo ▸ Configuración ▸ Configuración regional**: elige tu país. |
| Una persona no consigue entrar | Comprueba que su número está en **Configuración**, o vuelve a enviarle la palabra secreta. |

Cualquier cosa rota en la hoja se puede recuperar con **Archivo ▸ Historial de versiones**.

## Privacidad y seguridad, en palabras sencillas

Tus datos se quedan en tu propio Google Drive; no se envía nada a nadie, salvo tus mensajes con Telegram.

- **Solo entra tu lista.** El bot solo responde a los IDs de Telegram de Configuración, y solo en chats privados. Para todos los demás se queda callado.
- **El script solo puede abrir esta hoja**, no el resto de tu Drive.
- **Guarda dos cosas en secreto:** el token de BotFather y la palabra secreta. Si alguna se filtra, pide un token nuevo a BotFather (`/revoke`) y vuelve a usar **1. Conectar bot de Telegram**, o cambia la palabra secreta.
- **Las personas de solo lectura** pueden mirar y buscar, pero no pueden añadir, cambiar ni borrar nada.
- **Ninguna IA** lee tus datos: el bot son solo botones y reglas.
- Para compartir la hoja con alguien que solo deba mirar: **Compartir ▸ su correo ▸ Lector**.
