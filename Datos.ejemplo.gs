/**
 * EJEMPLO de datos iniciales. Solo se usan una vez, cuando ejecutás setup().
 * Copiá este archivo como Datos.gs y completalo con tu nómina real.
 * Datos.gs está en .gitignore: NO lo subas al repositorio (tiene DNI, CUIL y sueldos).
 */
const NOMINA_INICIAL = [
  // Clave, Apellido, Nombre, DNI, CUIL, Ingreso (AAAA-MM-DD), Tarea, Sector, Encargado, Salario neto
  ["30111222", "Pérez", "Juan Carlos", "30111222", "20-30111222-3", "2015-03-01", "Operario", "Elaboracion", "Cristian", 2500000],
  ["gomez-ana", "Gómez", "Ana", "", "", "2024-06-10", "Administrativa", "Administrativa", "Maximo", 2000000]
];

const ENCARGADOS_INICIALES = [
  // Nombre, Teléfono (código de área sin 0 y sin 15)
  ["Cristian", "1100000000"],
  ["Maximo", "1100000001"]
];

// ID de la carpeta "Legajos" en Drive. Si no existe o no hay acceso, setup() crea una nueva.
const LEGAJOS_ID_INICIAL = "";
const TOPE_ADELANTO_PCT_INICIAL = 30;
