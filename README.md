# Personal · Pedidos y licencias

App web en Google Apps Script para que el personal haga sus pedidos desde el celular y Recursos Humanos los responda en un solo lugar.

## Qué hace

**Para cada empleado** (entra con usuario y contraseña):
- Ve su saldo de vacaciones de la temporada (1/10 al 30/4), calculado por antigüedad según la LCT: 14, 21, 28 o 35 días corridos.
- Ve sus vacaciones pendientes de temporadas anteriores, con su fecha de vencimiento.
- Ve sus días a favor en el banco de días.
- Pide vacaciones, días, francos compensatorios, adelantos de sueldo y licencias especiales, y avisa ausencias.
- Sube certificados (foto o PDF), que se guardan solos en `Legajos/Apellido Nombre` en Drive.
- Avisa a su encargado por WhatsApp con el mensaje ya escrito.

**Para Recursos Humanos:**
- Bandeja de pedidos por responder, con aviso si otra persona falta esos días y control del tope de adelanto.
- Calendario de quién falta.
- Nómina con ingreso, sueldo, encargado, usuario y contraseña temporal de cada uno.
- Vacaciones pendientes de temporadas anteriores, con vencimiento. Al pedir vacaciones se usan primero.
- Banco de días: carga de feriados o francos trabajados (por persona o por sector entero) y descuento automático al aprobar un franco.
- Historial con filtros.
- Tope de adelantos como % del sueldo neto, por período de liquidación del 26 al 25.
- Encargados y usuarios de RRHH.

Todos los datos quedan en una planilla de Google con estas hojas: Nomina, Pedidos, Banco, VacPendientes, Encargados, Admins y Config.

## Archivos

| Archivo | Qué es |
|---|---|
| `src/Codigo.gs` | Servidor: reglas, usuarios, planilla y Drive |
| `src/Index.html` | Pantalla de la app (empleados y RRHH) |
| `src/Datos.ejemplo.gs` | Formato de los datos iniciales. Copialo como `Datos.gs` con tus datos reales |
| `src/appsscript.json` | Configuración del proyecto: zona horaria de Argentina y app web |

> **Privacidad:** `Datos.gs` lleva DNI, CUIL y sueldos. Está excluido en `.gitignore`: no lo subas al repositorio, y mantené el repositorio **privado**.

## Instalación desde cero

1. En [script.google.com](https://script.google.com) creá un proyecto nuevo.
2. Creá tres archivos y pegá el contenido de cada uno:
   - **Codigo** (secuencia de comandos) con `Codigo.gs`.
   - **Datos** (secuencia de comandos) con tu `Datos.gs`.
   - **Index** (HTML) con `Index.html`.
3. Ejecutá `setup()` una vez. En el registro de ejecución aparece la contraseña temporal del usuario `rrhh`.
4. Publicá la app: **Implementar → Nueva implementación → Aplicación web**, con *Ejecutar como: Yo* y *Quién tiene acceso: Cualquier persona*.
5. Entrá a la URL con `rrhh` y elegí tu contraseña.

Si ya tenés una instalación anterior, ejecutá `migrar()` en lugar de `setup()`. Agrega las hojas y columnas nuevas sin borrar datos.

## Actualizar el código

Después de cambiar el código, usá **Implementar → Gestionar implementaciones → editar → Nueva versión**. Así la URL no cambia. Si en cambio creás una implementación nueva, la URL cambia y los empleados pierden el acceso.

### Con clasp (opcional)

[clasp](https://github.com/google/clasp) sincroniza este repositorio con el proyecto de Apps Script:

```bash
npm install -g @google/clasp
clasp login
cp .clasp.json.ejemplo .clasp.json   # poné el ID del proyecto: Configuración del proyecto → ID de la secuencia de comandos
clasp push                            # sube src/ al proyecto
```

En el editor, el archivo tiene que llamarse **Codigo**, sin tilde. Si se llama "Código", renombralo antes del primer `clasp push`; si no, clasp crea un segundo archivo y aparece el error *Identifier has already been declared*.

## Integraciones (opcionales)

Se configuran en **Configuración del proyecto → Propiedades del script**. Las claves van ahí y no en el código, así no terminan en el repositorio.

| Propiedad | Para qué |
|---|---|
| `SLACK_WEBHOOK` | Webhook entrante del canal privado `#rrhh`. Avisa cada movimiento: pedidos nuevos (con saldo, tope y quién más falta), cancelaciones, respuestas, certificados, banco de días, vacaciones pendientes, cambios de datos y de usuarios, y bloqueos por intentos fallidos. Cada aviso de pedido trae un link que abre ese pedido en la app. |
| `TRELLO_KEY`, `TRELLO_TOKEN` | Credenciales de la API de Trello. |
| `TRELLO_LIST_ID` | Lista donde se crea una tarjeta "Revisar pedido de … de …" por cada pedido nuevo. Al responder, la tarjeta se cierra; la de un adelanto aprobado queda como recordatorio para descontarlo, con vencimiento el 25. Ejecutá `configurarListaTrello()` para cargarla sola. |
| `TRELLO_LIST_HECHO` | Opcional. Lista a la que se mueven los pedidos respondidos; sin ella, se archivan. `configurarListaTrello()` la detecta si se llama Hecho, Listo o Terminado. |

El calendario **Vacaciones** de Google Calendar lo crea `migrar()` (o `setup()`). Las vacaciones, los pedidos de día y los francos aprobados se cargan ahí como eventos de día completo, y se quitan si la aprobación se deshace.

Para probar las tres conexiones, ejecutá `probarIntegraciones()`. Para recibir todas las mañanas un resumen de lo que está sin responder, ejecutá una vez `activarResumenDiario()`.
