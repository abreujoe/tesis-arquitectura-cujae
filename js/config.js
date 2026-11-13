/* =====================================================================
   CONFIGURACIÓN DEL SITIO
   Este es el único archivo que hace falta tocar para cambiar textos
   generales. Los datos de las tesis NO están aquí: salen del Excel
   que está en la carpeta "datos" (ver LEEME.md).
   ===================================================================== */
window.CONFIG = {
  // Título que se ve arriba y en la pestaña del navegador
  titulo: "Tesis de Arquitectura CUJAE",
  subtitulo: "Registro de trabajos de diploma de la Facultad de Arquitectura, Universidad Tecnológica de La Habana José Antonio Echeverría",

  // Archivo de datos. Para actualizar el sitio basta con reemplazar este archivo.
  excel: "datos/registro.xlsx",
  hoja: "Registro",            // nombre de la hoja con las tesis

  // Fecha que se muestra como "datos actualizados el…". Si se deja vacío,
  // el sitio usa la fecha en que se subió el Excel (cuando el servidor la informa).
  fechaDatos: "",

  // Cómo avisar de un error. Deja vacío lo que no quieras mostrar.
  contactoCorreo: "",          // p. ej. "nombre@correo.com"
  repositorio: "",             // p. ej. "https://github.com/usuario/tesis-arquitectura-cujae"

  // Filas que no se muestran ni se cuentan: las que en la columna
  // "Tipo de aviso" tengan alguno de estos textos.
  excluirAvisos: ["No es tesis"],

  // Mostrar en la ficha de cada tesis las notas de la columna "Observaciones"
  mostrarObservaciones: true,

  // Cuántas tesis se muestran de entrada y con cada "Mostrar más"
  porPagina: 30
};
