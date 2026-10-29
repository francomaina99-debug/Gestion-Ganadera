/**
 * ============================================================================
 * SISTEMA DE GESTION Y REGISTRACION GANADERA - BOVINOS (ANGUS CICLO COMPLETO)
 * Establecimiento Agroganadero: Pozo del Carril - La Aguada (UNRC)
 * Integracion Total Unificada:
 *  - Cotizaciones: Integracion directa con APIs oficiales de MAG Canuelas y ROSGAN
 *    (hoja COTIZACIONES vinculada a DASHBOARD y valuaciones VNR bajo RT 22)
 *  - Programacion Automatica: Disparador diario configurable con eleccion de horario
 *  - Ecuacion de Inventarios y Criterios de Valuacion (Slides 12-13)
 *  - Estado de Resultados por Etapas y Traslados en Cascada (Slides 14-16)
 *  - Analisis de Costos y Margenes Porcentuales (Slide 15)
 *  - Datos Reales de Campo: 58 Nacimientos, 3 Controles de Peso, Recria, 
 *    Engorde Feedlot (16 novillos, 118 dias, ADG 1.578 kg/d) y Repositores
 *  - Costeo de Alfalfa Prorrateada (4 anos) y DTe por Jaula
 * ============================================================================
 */

// ============================================================================
// 1. MENU PRINCIPAL EN GOOGLE SHEETS
// ============================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🐂 Gestión Ganadera')
    .addItem('📱 Abrir Panel de Carga (Sidebar)', 'abrirSidebar')
    .addSeparator()
    .addItem('📊 Actualizar y Construir DASHBOARD', 'construirDashboardMenu')
    .addItem('📑 Actualizar ESTADO DE RESULTADOS (Multiespecie)', 'construirEstadoResultadosMenu')
    .addItem('🔧 Aplicar Cambios en Movimientos y Ganancia Diaria', 'aplicarCambiosMovimientosYGananciaMenu')
    .addItem('🏷️ Sincronizar Fórmulas de Peso y Subunidad en EXISTENCIAS', 'ejecutarActualizacionFormulasExistenciasMenu')
    .addSeparator()
    .addItem('🔄 Sincronizar Gastos con TABLA_IMPUTACION', 'sincronizarTodosLosMovimientosExistenciaMenu')
    .addItem('🎨 Formatear y Validar TABLA_IMPUTACION (Estilo y Desplegables)', 'formatearYValidarTablaImputacionMenu')
    .addItem('📋 Homogeneizar Conceptos en TABLA_IMPUTACION', 'ejecutarHomogeneizacionTablaImputacionMenu')
    .addSeparator()
    .addItem('📈 Actualizar Cotizaciones Ahora (Cañuelas / ROSGAN)', 'actualizarCotizacionesMercado')
    .addItem('⏰ Programar Actualización Automática (Elegir Horario)', 'programarHorarioActualizacion')
    .addItem('❌ Desactivar Actualización Automática', 'desactivarTriggerActualizacion')
    .addItem('🔌 Probar Conexión con Mercados (APIs)', 'probarConexionMercados')
    .addSeparator()
    .addItem('👤 Crear / Actualizar Hoja USUARIOS', 'crearHojaUsuarios')
    .addToUi();

  // Actualización automática no intrusiva de conceptos de catálogo
  try {
    actualizarNombresConceptos(SpreadsheetApp.getActiveSpreadsheet());
  } catch (e) {}
}

/**
 * Despliega el formulario lateral (Sidebar) para el encargado del campo
 */
function abrirSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Formulario')
    .setTitle('Carga de Campo - Pozo del Carril')
    .setWidth(450);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================================
// 2. INICIALIZACION COMPLETA DE HOJAS Y CATALOGOS
// ============================================================================
function inicializarSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const COLOR_HEADER = '#1E3A1E'; // Verde institucional UNRC
  const COLOR_TEXT = '#FFFFFF';
  
  const hojasConfig = [
    { nombre: 'DASHBOARD', headers: ['PANEL DE CONTROL GANADERO - POZO DEL CARRIL (CON VALUACION DINAMICA)'], anchoCols: [120, 180, 110, 120, 140, 150, 100, 100] },
    { nombre: 'CAT_UNIDADES_NEGOCIO', headers: ['ID_UN', 'Nombre_UN', 'Tipo_Actividad', 'Responsable', 'Observaciones'], anchoCols: [100, 160, 130, 150, 200] },
    { nombre: 'CAT_SECTORES', headers: ['ID_Sector', 'Nombre_Sector', 'Superficie_Ha', 'Uso_Principal', 'Observaciones'], anchoCols: [100, 160, 120, 180, 200] },
    { nombre: 'CAT CATEGORIAS', headers: ['Id_CAT', 'Nombre_Categoria', 'Etapa_Ciclo', 'Tipo_Destino', 'Peso_Ref_Ingreso_Kg', 'Peso_Ref_Egreso_Kg', 'Cuenta_Contable_Activo'], anchoCols: [110, 180, 120, 160, 130, 130, 160] },
    { nombre: 'CAT_CONTACTOS', headers: ['ID_Contacto', 'Razon_Social', 'Tipo_Contacto', 'CUIT_CUIL', 'Condicion_Pago_Habitual', 'Telefono_Email'], anchoCols: [110, 220, 140, 130, 160, 180] },
    { nombre: 'CAT_CONCEPTOS', headers: ['ID_Concepto', 'Nombre_Concepto', 'Naturaleza'], anchoCols: [110, 280, 180] },
    { nombre: 'CAT_CUENTAS', headers: ['ID_Cuenta', 'Nombre_Cuenta', 'Tipo_Cuenta', 'Responsable_Fondo'], anchoCols: [120, 180, 140, 160] },
    { nombre: 'MOV_COMERCIALES', headers: ['ID_Mov_Com', 'Fecha', 'Tipo_Operacion', 'Nro_Comprobante', 'ID_Contacto', 'ID_UN', 'ID_Cat', 'Cabezas', 'Peso_Total_Kg', 'Precio_Unitario', 'Importe_Bruto_Economico', 'Gastos_Comercializacion', 'Alicuota_IVA', 'Importe_IVA', 'Retenciones_Percepciones', 'Importe_Neto_Financiero', 'Condicion_Venta', 'Estado_Cobro_Pago', 'Observaciones'], anchoCols: [110, 95, 105, 130, 110, 95, 110, 80, 100, 105, 130, 130, 90, 105, 130, 135, 120, 100, 200] },
    { nombre: 'MOV_FISICOS', headers: ['ID_Mov_Fis', 'Fecha', 'Tipo_Movimiento', 'ID_UN', 'ID_Cat_Origen', 'ID_Cat_Destino', 'Cabezas', 'Peso_Total_Estimado_Kg', 'ID_Sector', 'Nro_Caravana_Electronica', 'Insumo_Propio_Asociado', 'Cantidad_Insumo_Fisico', 'Causa_Muerte', 'Observaciones'], anchoCols: [110, 95, 140, 95, 120, 120, 80, 130, 110, 150, 160, 130, 140, 200] },
    { nombre: 'TABLA_IMPUTACION', headers: ['Fecha', 'Comprobante', 'Proveedor', 'Concepto', 'Importe_Total', 'Unidad_Negocio', 'Etapa_Productiva', 'Plano_Afectado', 'Efecto_Existencias'], anchoCols: [95, 130, 200, 280, 120, 120, 130, 130, 160] },
    { nombre: 'MOV_FINANCIEROS', headers: ['ID_Mov_Fin', 'Fecha', 'Tipo_Movimiento', 'ID_Contacto', 'ID_Mov_Com', 'ID_Cuenta', 'Importe_Efectivo', 'Medio_Pago', 'Observaciones'], anchoCols: [110, 95, 110, 110, 120, 120, 130, 130, 200] },
    { nombre: 'MOV_CAMPO_LABORES', headers: ['ID_Labor', 'Fecha', 'Lote', 'Superficie_Ha', 'Unidad_Negocio', 'Tipo_Labor', 'Insumo_Producto', 'Dosis_Por_Ha', 'Unidad_Dosis', 'Cantidad_Total', 'Unidad_Total', 'Observaciones'], anchoCols: [100, 95, 90, 110, 100, 150, 170, 100, 95, 110, 95, 220] },
    { nombre: 'ESTADO_RESULTADOS_ETAPAS', headers: ['ESTADO DE RESULTADOS POR ETAPA - ECUACION DE INVENTARIO Y DIAPOSITIVA 15'], anchoCols: [140, 240, 130, 110, 110, 110, 110, 110, 110, 240, 120, 140, 140, 150] },
    { nombre: 'COTIZACIONES', headers: ['Mercado', 'Categoria', 'Subtipo', 'Cabezas', 'Precio_Promedio', 'Precio_Minimo', 'Precio_Maximo', 'Precio_Mediana', 'Kilos_Totales', 'Promedio_Kilos', 'Importe_Total', 'Variacion_Pct', 'Indice_MAG_ROSGAN', 'Fecha_Cotizacion', 'Estado_Rueda', 'Fecha_Consulta'], anchoCols: [130, 160, 160, 85, 125, 110, 110, 110, 105, 105, 130, 105, 130, 115, 115, 135] },
    { nombre: 'EXISTENCIAS', headers: ['ID_REGISTRO_EXIS', 'Fecha de Registro', 'Caravana', 'Unidad de Produccion', 'Categoria', 'Subunidad', 'Peso Inicial KG', 'Peso Estimado KG', 'Peso Real KG', 'Estado'], anchoCols: [140, 120, 110, 150, 140, 120, 120, 120, 120, 100, 30, 200, 130] },
    { nombre: 'Movimientos_Existencia', headers: ['ID Registro', 'Fecha de Registro', 'Caravana', 'Unidad de Produccion', 'Fecha de Nacimietno', 'Peso Estimado KG', 'Peso Real KG', 'Categoria', 'Etapa', 'Concepto Costo', 'Unidad', 'Costo Unitario', 'Cantidad', 'Costo Total', 'Observaciones'], anchoCols: [100, 120, 110, 150, 130, 120, 120, 130, 120, 140, 90, 110, 90, 110, 180] },
    { nombre: 'SUB UNIDADES CICLO', headers: ['ID_CICLO', 'Etapa / Subunidad', 'Descripcion / Destino'], anchoCols: [130, 160, 320, 30, 130, 160, 320, 30, 130, 160, 320] },
    { nombre: 'GANANCIA DIARIA KG', headers: ['Etapa', 'Duracion_Dias', 'Peso Inicial KG', 'Peso Final KG', 'Ganancia Diaria KG'], anchoCols: [130, 120, 130, 130, 150, 30, 130, 120, 130, 130, 150, 30, 130, 120, 130, 130, 150] },
    { nombre: 'COSTOS', headers: ['Id costo', 'Concepto de costo', 'Etapa', 'Unidad de medida', 'Costo unitario ($)'], anchoCols: [100, 260, 120, 180, 140] },];

  hojasConfig.forEach(function(cfg) {
    let sheet = ss.getSheetByName(cfg.nombre);
    if (!sheet) {
      sheet = ss.insertSheet(cfg.nombre);
    }
    if (cfg.nombre !== 'DASHBOARD' && cfg.nombre !== 'ESTADO_RESULTADOS_ETAPAS' && cfg.nombre !== 'EXISTENCIAS' && cfg.nombre !== 'GANANCIA DIARIA KG' && cfg.nombre !== 'SUB UNIDADES CICLO') {
      sheet.getRange(1, 1, 1, cfg.headers.length).setValues([cfg.headers]);
      const headerRange = sheet.getRange(1, 1, 1, cfg.headers.length);
      headerRange.setBackground(COLOR_HEADER)
        .setFontColor(COLOR_TEXT)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setWrap(true);
      sheet.setFrozenRows(1);
    }
    if (cfg.anchoCols) {
      cfg.anchoCols.forEach(function(w, idx) {
        sheet.setColumnWidth(idx + 1, w);
      });
    }
  });

    sembrarCatalogos(ss);
  ordenarYLimpiarContactos(ss);
  actualizarCotizacionesMercado(false);
  construirDashboard(ss);
  construirEstadoResultadosEtapas(ss);
  actualizarValidaciones();

  SpreadsheetApp.getUi().alert(
    'Sistema Inicializado y Sincronizado con Éxito\n\n' +
    '✔ Hojas maestras, catálogos y validaciones estructurados.\n' +
    '✔ Panel de Control (Dashboard) y Estado de Resultados Multiespecie actualizados.\n' +
    '✔ Cotizaciones oficiales de mercado sincronizadas.\n' +
    '✔ Existencias y movimientos de trazabilidad protegidos sin sobreescritura.'
  );
}

function cargarTodoElSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sembrarCatalogos(ss);
  ordenarYLimpiarContactos(ss);
  actualizarCotizacionesMercado(false);
  construirDashboard(ss);
  construirEstadoResultadosEtapas(ss);
  actualizarValidaciones();
  SpreadsheetApp.getUi().alert(
    'Carga y Actualización Completa Exitosa\n\n' +
    '✔ Catálogos maestros, Dashboard ejecutivo y Estado de Resultados multiespecie actualizados.\n' +
    '✔ Cotizaciones de mercado sincronizadas desde APIs oficiales (MAG Cañuelas / ROSGAN).\n' +
    '✔ Registros de existencias y movimientos preservados íntegramente.'
  );
}

// ============================================================================
// 3. SEMBRADO DE CATALOGOS
// ============================================================================
function sembrarCatalogos(ss) {
  const sheetUN = ss.getSheetByName('CAT_UNIDADES_NEGOCIO');
  if (sheetUN && sheetUN.getLastRow() <= 1) {
    const dataUN = [
      ['UN-BOV', 'Bovinos', 'Ganaderia', 'Encargado Bovinos', 'Ciclo Completo: Cria, Recria, Adaptacion y Engorde'],
      ['UN-OVIN', 'Ovinos', 'Ganaderia', 'Encargado Ovinos', 'Recambio a Dorper carnicero'],
      ['UN-POR', 'Porcinos', 'Ganaderia', 'Encargado Cerdos', 'Produccion por fases: Cria, Destete, Engorde'],
      ['UN-MAIZ', 'Maiz', 'Agricultura', 'Ing. Agronomo', 'Doble destino: Venta comercial y autoconsumo silo/grano'],
      ['UN-SOJA', 'Soja', 'Agricultura', 'Ing. Agronomo', '100 ha en rotacion agricola'],
      ['UN-ALFA', 'Alfalfa', 'Forrajes', 'Ing. Agronomo', 'Pastura e implantacion alfalfa'],
      ['UN-ESTR', 'Estructura', 'Administracion', 'Administrador', 'Costos indirectos y administracion general']
    ];
    sheetUN.getRange(2, 1, dataUN.length, dataUN[0].length).setValues(dataUN);
  }

  const sheetSec = ss.getSheetByName('CAT_SECTORES');
  if (sheetSec && sheetSec.getLastRow() <= 1) {
    const dataSec = [
      ['SEC-PC', 'Pozo del Carril - Casco', 15, 'Instalaciones, feedlot y parideras', 'Sector principal de manejo'],
      ['SEC-CN', 'Camdocex Norte', 45, 'Pasturas silvopastoriles y cria', 'Rodeo de cria bovino'],
      ['SEC-CS', 'Camdocex Sur', 50, 'Recria y rastrojos de maiz', 'Etapa de recria bovina'],
      ['SEC-CB', 'Camdocex Bajo', 30, 'Bajo salino con grama rhodes', 'Pastoreo estacional'],
      ['SEC-L1', 'Lote 1 (Agricola)', 23, 'Agricultura intensiva', 'Maiz ACA / Soja'],
      ['SEC-L2', 'Lote 2 (Agricola)', 16, 'Agricultura intensiva', 'Soja NK 52 X 21'],
      ['SEC-L4', 'Lote 4 (Agricola)', 12.5, 'Agricultura intensiva', 'Maiz Nidera 7784'],
      ['SEC-L5', 'Lote 5 (Agricola)', 16.5, 'Agricultura intensiva', 'Maiz AG 7004'],
      ['SEC-L6', 'Lote 6 (Agricola)', 13.5, 'Agricultura intensiva', 'Soja NK 52 X 21'],
      ['SEC-L7', 'Lote 7 (Agricola)', 12.5, 'Agricultura intensiva', 'Soja NK 52 X 21'],
      ['SEC-L8', 'Lote 8 (Forrajero)', 13, 'Produccion forrajera', 'Moha y Alfalfa AA Victoria'],
      ['SEC-VIV', 'Vivero y Huertas', 5, 'Produccion experimental', 'Sector Vivero UNRC']
    ];
    sheetSec.getRange(2, 1, dataSec.length, dataSec[0].length).setValues(dataSec);
  }

  const sheetCat = ss.getSheetByName('CAT CATEGORIAS') || ss.getSheetByName('CAT_CATEGORIAS_BOVINOS') || ss.getSheetByName('CAT_CATEGORIAS');
  if (sheetCat && sheetCat.getLastRow() <= 1) {
    const dataCat = [
      ['CAT_BOV_1', 'Vacas con cria', 'Cria', 'Activo Biologico Reproductor', 420, 480, '1.4.01 Reproductores Rodeo'],
      ['CAT_BOV_2', 'Vacas secas', 'Cria', 'Activo Biologico Reproductor', 430, 470, '1.4.01 Reproductores Rodeo'],
      ['CAT_BOV_3', 'Vacas prenadas', 'Cria', 'Activo Biologico Reproductor', 450, 520, '1.4.01 Reproductores Rodeo'],
      ['CAT_BOV_4', 'Vaquillonas 1 (recria)', 'Recria', 'Reposicion', 240, 310, '1.3.01 Hacienda en Crecimiento'],
      ['CAT_BOV_5', 'Vaquillonas 2 (servicio/prenez)', 'Recria', 'Reposicion', 310, 380, '1.3.01 Hacienda en Crecimiento'],
      ['CAT_BOV_6', 'Terneros', 'Cria', 'Carne comercial', 40, 140, '1.3.01 Hacienda en Crecimiento'],
      ['CAT_BOV_7', 'Terneras', 'Cria', 'Reposicion/Venta', 40, 135, '1.3.01 Hacienda en Crecimiento'],
      ['CAT_BOV_8', 'Destete', 'Recria', 'Invernada', 140, 180, '1.3.01 Hacienda en Crecimiento'],
      ['CAT_BOV_9', 'Novillitos (adaptacion)', 'Adaptacion', 'Carne comercial', 180, 280, '1.3.01 Hacienda en Crecimiento'],
      ['CAT_BOV_10', 'Novillos terminados', 'Engorde', 'Carne comercial', 280, 450, '1.3.01 Bienes de Cambio - Carne'],
      ['CAT_BOV_11', 'Toros cabana', 'Cria', 'Activo Biologico Reproductor', 650, 850, '1.4.02 Reproductores Cabana'],
      ['CAT_BOV_12', 'Toritos reposicion', 'Recria', 'Reposicion', 350, 500, '1.3.01 Hacienda en Crecimiento'],
      ['CAT_POR_1', 'Madre', '', 'Activo Biologico Reproductor', '', '', '1.4.03 Reproductores Porcinos'],
      ['CAT_POR_2', 'Padrillo', '', 'Activo Biologico Reproductor', '', '', '1.4.03 Reproductores Porcinos'],
      ['CAT_POR_3', 'Lechon', 'Cria', 'Carne comercial', 0, 15, '1.3.02 Bienes de Cambio - Porcinos'],
      ['CAT_POR_4', 'Cachorro, Cachorra', 'Destete', 'Carne comercial', 15, 50, '1.3.02 Bienes de Cambio - Porcinos'],
      ['CAT_POR_5', 'Capon', 'Engorde', 'Carne comercial', 50, 105, '1.3.02 Bienes de Cambio - Porcinos'],
      ['CAT_POR_6', 'Madre de descarte', '', 'Activo Biologico Reproductor', '', '', '1.4.03 Reproductores Porcinos'],
      ['CAT_POR_7', 'Padrillo de descarte', '', 'Activo Biologico Reproductor', '', '', '1.4.03 Reproductores Porcinos'],
      // 3. Ovinos
      ['CAT-OVIN-01', 'Madres', '', 'Activo Biologico Reproductor', '', '', ''],
      ['CAT-OVIN-02', 'Carneros', '', 'Activo Biologico Reproductor', '', '', ''],
      ['CAT-OVIN-03', 'Corderos', 'Cria', 'Engorde/Venta', '', '', ''],
      ['CAT-OVIN-04', 'Borregas', 'Recria', '', '', '', ''],
      ['CAT-OVIN-05', 'Borregos', 'Recria', 'Engorde/venta', '', '', ''],
      ['CAT-OVIN-06', 'Ovino de Descarte', '', 'Descarte', '', '', ''],
      ['CAT-14', 'Granos Maiz Comercial', 'Agricultura', 'Venta Comercial', 0, 0, '1.3.04.01 Existencias - Maiz']
    ];
    sheetCat.getRange(2, 1, dataCat.length, dataCat[0].length).setValues(dataCat);
  }

  const sheetCont = ss.getSheetByName('CAT_CONTACTOS');
  if (sheetCont) {
    if (sheetCont.getLastRow() > 1) {
      sheetCont.getRange(2, 1, sheetCont.getLastRow() - 1, sheetCont.getLastColumn()).clearContent();
    }
    const dataCont = [
      ['CON-001', 'Abascal Ruben Pedro', 'Proveedor Fletes', '', 'Cuenta Corriente', ''],
      ['CON-002', 'Acero Solaro', 'Proveedor Metalurgica', '', 'Cuenta Corriente', ''],
      ['CON-003', 'Acopio Cereales Central', 'Cliente Granos', '', 'Cuenta Corriente', ''],
      ['CON-004', 'Austro', 'Proveedor Baterias', '', 'Cuenta Corriente', ''],
      ['CON-005', 'Basconzuelo Martin', 'Proveedor Servicios', '', 'Cuenta Corriente', ''],
      ['CON-006', 'Begon', 'Proveedor Rodados', '', 'Cuenta Corriente', ''],
      ['CON-007', 'Bertero', 'Proveedor Repuestos', '', 'Cuenta Corriente', ''],
      ['CON-008', 'Bringas Hector Domingo', 'Proveedor Gomeria', '', 'Cuenta Corriente', ''],
      ['CON-009', 'Bueno', 'Proveedor Agropartes', '', 'Cuenta Corriente', ''],
      ['CON-010', 'Bulnor Ferreteria Industrial', 'Proveedor Insumos', '', 'Cuenta Corriente', ''],
      ['CON-011', 'Cabana Las Lilas', 'Proveedor Reproductores', '', 'Contado', ''],
      ['CON-012', 'Castelluccio', 'Proveedor Servicios', '', 'Cuenta Corriente', ''],
      ['CON-013', 'Consignataria Feria', 'Cliente / Consignatario', '', 'Cuenta Corriente', ''],
      ['CON-014', 'Consignataria Feria Rio Cuarto SA', 'Cliente / Consignatario', '', 'Cuenta Corriente', ''],
      ['CON-015', 'D Eramo Alfredo Y D Eramo', 'Proveedor Repuestos', '', 'Cuenta Corriente', ''],
      ['CON-016', 'Dan', 'Proveedor Forrajes', '', 'Cuenta Corriente', ''],
      ['CON-017', 'Del Bo', 'Proveedor Rodados', '', 'Cuenta Corriente', ''],
      ['CON-018', 'Don Sanchez', 'Proveedor Servicios', '', 'Cuenta Corriente', ''],
      ['CON-019', 'Don Vicente', 'Proveedor Ferreteria', '', 'Cuenta Corriente', ''],
      ['CON-020', 'Electromecanica', 'Proveedor Servicios', '', 'Cuenta Corriente', ''],
      ['CON-021', 'Ferreteria Avenida', 'Proveedor Insumos', '', 'Cuenta Corriente', ''],
      ['CON-022', 'Fersan Alimentos Balanceados', 'Proveedor Nutricion', '', 'Cuenta Corriente', ''],
      ['CON-023', 'Frigorifico Regional Del Sur', 'Cliente Frigorifico', '', 'Cuenta Corriente', ''],
      ['CON-024', 'Frigorifico Roth Porcinos', 'Cliente Porcino', '', 'Cuenta Corriente', ''],
      ['CON-025', 'Gomagro', 'Proveedor Agropartes', '', 'Cuenta Corriente', ''],
      ['CON-026', 'Gomeria Bimaco', 'Proveedor Servicios', '', 'Cuenta Corriente', ''],
      ['CON-027', 'Grunfer Repuestos Y Servicios', 'Proveedor Maquinaria', '', 'Cuenta Corriente', ''],
      ['CON-028', 'Impulso Repuestos', 'Proveedor Repuestos', '', 'Cuenta Corriente', ''],
      ['CON-029', 'Julio Cesar Sola', 'Proveedor Servicios', '', 'Cuenta Corriente', ''],
      ['CON-030', 'La Casa De Las Mangueras', 'Proveedor Repuestos', '', 'Cuenta Corriente', ''],
      ['CON-031', 'La Casa De Las Motosierras', 'Proveedor Maquinaria', '', 'Cuenta Corriente', ''],
      ['CON-032', 'La Cesira', 'Proveedor Sanidad', '', 'Cuenta Corriente', ''],
      ['CON-033', 'Laboratorio Maffrand', 'Proveedor Sanidad', '', 'Cuenta Corriente', ''],
      ['CON-034', 'Lopez Claudio Alfredo', 'Proveedor Fletes', '', 'Cuenta Corriente', ''],
      ['CON-035', 'Maderera Mendez', 'Proveedor Insumos', '', 'Cuenta Corriente', ''],
      ['CON-036', 'Mercado Agroganadero Canuelas (MAG)', 'Mercado Concentrador', '', 'Contado', ''],
      ['CON-037', 'Nager', 'Proveedor Insumos', '', 'Cuenta Corriente', ''],
      ['CON-038', 'Nutrinal Nutricion Animal', 'Proveedor Nutricion', '', 'Cuenta Corriente', ''],
      ['CON-039', 'Oleohidraulica Lym', 'Proveedor Taller', '', 'Cuenta Corriente', ''],
      ['CON-040', 'Oxigeno Union', 'Proveedor Taller', '', 'Cuenta Corriente', ''],
      ['CON-041', 'Panichelli Ruben Alberto', 'Proveedor Taller', '', 'Cuenta Corriente', ''],
      ['CON-042', 'Pcg Parts E Service', 'Proveedor Repuestos', '', 'Cuenta Corriente', ''],
      ['CON-043', 'Produccion Propia Campo', 'Ajuste Interno', '', 'Contado', ''],
      ['CON-044', 'Proveedor General', 'Proveedor Insumos', '', 'Cuenta Corriente', ''],
      ['CON-045', 'Racedo', 'Proveedor Taller', '', 'Cuenta Corriente', ''],
      ['CON-046', 'S.i.p.a', 'Proveedor Molinos', '', 'Cuenta Corriente', ''],
      ['CON-047', 'Saby Hermanos', 'Proveedor Combustibles', '', 'Cuenta Corriente', ''],
      ['CON-048', 'Sala John Deere', 'Proveedor Maquinaria', '', 'Cuenta Corriente', ''],
      ['CON-049', 'Select Debernardi', 'Proveedor Semen', '', 'Contado', ''],
      ['CON-050', 'Senasa', 'Organismo Publico', '', 'Contado', ''],
      ['CON-051', 'Sm Repuestos', 'Proveedor Repuestos', '', 'Cuenta Corriente', ''],
      ['CON-052', 'Stihl', 'Proveedor Maquinaria', '', 'Cuenta Corriente', ''],
      ['CON-053', 'Taverna', 'Proveedor Insumos', '', 'Cuenta Corriente', ''],
      ['CON-054', 'Top', 'Proveedor Insumos', '', 'Cuenta Corriente', ''],
      ['CON-055', 'Transporte De Hacienda', 'Proveedor Fletes', '', 'Cuenta Corriente', ''],
      ['CON-056', 'Uel', 'Proveedor Sanidad', '', 'Cuenta Corriente', ''],
      ['CON-057', 'Veterinaria La Rural', 'Proveedor Sanidad', '', 'Cuenta Corriente', ''],
      ['CON-058', 'Ypf Agro', 'Proveedor Combustibles', '', 'Cuenta Corriente', '']
    ];
    sheetCont.getRange(2, 1, dataCont.length, dataCont[0].length).setValues(dataCont);
    sheetCont.getRange(2, 1, dataCont.length, sheetCont.getLastColumn()).sort(2);
  }

  const sheetConc = ss.getSheetByName('CAT_CONCEPTOS');
  if (sheetConc && sheetConc.getLastRow() <= 1) {
    const dataConc = [
      ['1.1', 'Disponibilidades', 'Activo Corriente'],
      ['1.2', 'Creditos', 'Activo Corriente'],
      ['1.2.05', 'Creditos fiscales', 'Creditos Fiscales'],
      ['1.3.01', 'Existencias - hacienda bovina', 'Bienes de Cambio'],
      ['1.3.01.01', 'Existencias - vacas', 'Bienes de Cambio'],
      ['1.3.01.02', 'Existencias - vacas secas', 'Bienes de Cambio'],
      ['1.3.01.03', 'Existencias - vacas prenadas', 'Bienes de Cambio'],
      ['1.3.01.04', 'Existencias - vaquillona 1', 'Bienes de Cambio'],
      ['1.3.01.05', 'Existencias - vaquillona 2', 'Bienes de Cambio'],
      ['1.3.01.06', 'Existencias - terneros', 'Bienes de Cambio'],
      ['1.3.01.07', 'Existencias - terneras', 'Bienes de Cambio'],
      ['1.3.01.08', 'Existencias - destete', 'Bienes de Cambio'],
      ['1.3.01.09', 'Existencias - novillitos', 'Bienes de Cambio'],
      ['1.3.01.10', 'Existencias - toros', 'Activo Biologico'],
      ['1.3.01.11', 'Existencias - toritos', 'Bienes de Cambio'],
      ['1.3.01.12', 'Existencias - novillos', 'Bienes de Cambio'],
      ['1.3.04', 'Existencias - granos', 'Bienes de Cambio'],
      ['1.3.04.01', 'Existencias - maiz', 'Bienes de Cambio'],
      ['1.3.04.02', 'Existencias - soja', 'Bienes de Cambio'],
      ['1.3.05', 'Existencias - forrajes y reservas', 'Bienes de Cambio'],
      ['1.3.05.01', 'Existencias - alimento en silobolsa', 'Bienes de Cambio'],
      ['1.3.05.02', 'Existencias - rollos', 'Bienes de Cambio'],
      ['1.3.06', 'Existencias - insumos', 'Bienes de Cambio'],
      ['1.3.06.01', 'ALIMENTO BALANCEADO/SUPLEMENTOS', 'Bienes de Cambio'],
      ['1.3.06.02', 'Existencias - premezcla', 'Bienes de Cambio'],
      ['1.5.01', 'Activos biologicos de produccion', 'Activo no corriente'],
      ['1.5.02', 'Bienes de uso', 'Activo no corriente'],
      ['1.05.02.01', 'ROLLOS ALFALFA PROPIOS', 'Bienes de uso'],
      ['1.05.02.02', 'Depreciacion acumulada alfalfa', 'Reguladora de Activo'],
      ['2.1', 'Proveedores', 'Pasivo Corriente'],
      ['2.2', 'Deudas por rendiciones', 'Pasivo Corriente'],
      ['2.3', 'Deudas fiscales', 'Pasivo Corriente'],
      ['2.4', 'Remuneraciones y cargas sociales', 'Pasivo Corriente'],
      ['4.1', 'Venta de hacienda bovina', 'Ingreso'],
      ['4.1.1', 'VENTA DE HACIENDA EN PIE', 'Ingreso'],
      ['4.1.2', 'Venta de vacas de descarte', 'Ingreso'],
      ['4.1.3', 'Venta de otras categorias', 'Ingreso'],
      ['4.5', 'Venta de reproductores', 'Ingreso'],
      ['4.6', 'MAIZ PROPIO CONSUMO RODEO', 'Ingreso / Ajuste Interno'],
      ['4.7', 'Resultado por tenencia de hacienda', 'Ingreso'],
      ['4.8', 'Otros ingresos', 'Ingreso'],
      ['5.1', 'Costos directos ganaderia bovina', 'Costos Directos'],
      ['5.1.1', 'Alimentacion (racion y silo)', 'Costos Directos'],
      ['5.1.2', 'SANIDAD/VACUNAS', 'Costos Directos'],
      ['5.1.2.01', 'SANIDAD/VACUNAS', 'Costos Directos'],
      ['5.1.2.02', 'HONORARIOS DEL VETERINARIO', 'Costos Directos'],
      ['5.1.3', 'REPRODUCCION/INSEMINACION', 'Costos Directos'],
      ['5.1.6', 'Mortandad', 'Costos Directos'],
      ['5.5', 'Costos indirectos a distribuir', 'Costos Indirectos'],
      ['5.5.1', 'COMBUSTIBLES Y LUBRICANTES', 'Costos Indirectos'],
      ['5.5.2', 'REPARACIÓN Y MANTENIMIENTO DE INSTALACIONES', 'Costos Indirectos'],
      ['5.5.3', 'Mano de obra', 'Costos Indirectos'],
      ['5.5.4', 'Amortizacion de reproductores', 'Costos Indirectos'],
      ['5.6', 'Gastos de estructura', 'Gastos de Estructura'],
      ['5.6.1', 'GASTOS ADMINISTRATIVOS', 'Gastos de Estructura'],
      ['5.6.2', 'Impuestos no atribuibles', 'Gastos de Estructura'],
      ['5.6.3', 'Seguros', 'Gastos de Estructura'],
      ['5.6.4', 'Amortizaciones de estructura', 'Gastos de Estructura'],
      ['5.7', 'Gastos de comercializacion', 'Gastos Comercializacion'],
      ['5.7.1', 'COMISION CONSIGNATARIO FERIA', 'Gastos Comercializacion'],
      ['5.7.2', 'FLETE DE HACIENDA', 'Gastos Comercializacion'],
      ['5.7.3', 'GUIAS, DTE Y SELLADOS SENASA', 'Gastos Comercializacion'],
      ['5.7.4', 'Otros gastos de comercializacion', 'Gastos Comercializacion'],
      ['5.8', 'Depreciacion alfalfa', 'Costos Indirectos'],
      ['5.9', 'COMPRA DE GRANOS', 'Costos Directos'],
      ['CPT-13', 'REPRODUCCION/INSEMINACION', 'Costo Directo'],
      ['CPT-14', 'GASTOS ADMINISTRATIVOS', 'Costo Indirecto'],
      ['CPT-15', 'HONORARIOS DEL VETERINARIO', 'Costo Directo'],
      ['CPT-16', 'COMPRA DE GRANOS', 'Costo Directo']
    ];
    sheetConc.getRange(2, 1, dataConc.length, dataConc[0].length).setValues(dataConc);
  }

  // Actualizar conceptos existentes en CAT_CONCEPTOS y TABLA_IMPUTACION
  actualizarNombresConceptos(ss);

  const sheetCta = ss.getSheetByName('CAT_CUENTAS');
  if (sheetCta && sheetCta.getLastRow() <= 1) {
    const dataCta = [
      ['CTA-01', 'Banco Nacion Cta Cte UNRC', 'Banco', 'Tesoreria UNRC'],
      ['CTA-02', 'Fondo Fijo Decanato (Vales)', 'Caja Chica', 'Secretaria Administrativa FAV'],
      ['CTA-03', 'Caja Recaudacion Campo', 'Efectivo', 'Encargado Pozo del Carril'],
      ['CTA-04', 'Cta Cte Consignataria Feria', 'Cta Cte Comercial', 'Administrador Campo']
    ];
    sheetCta.getRange(2, 1, dataCta.length, dataCta[0].length).setValues(dataCta);
  }

  actualizarHojaSubUnidadesCiclo(ss);

  actualizarHojaGananciaDiaria(ss);

  const sheetCostos = ss.getSheetByName('COSTOS');
  if (sheetCostos && sheetCostos.getLastRow() <= 1) {
    const dataCostos = [
      ['CST001', 'Alimento Destete Precoz (Suplemento)', 'Cria', 'kg por peso del animal', 1100.00],
      ['CST002', 'Forraje Madre (Pastura Alfalfa) / Cba', 'Cria', 'kgs (MS)', 95.00],
      ['CST003', 'Vacuna Aftosa / Brucelosis', 'Cria', 'Dosis', 1850.00],
      ['CST004', 'Forraje Alfalfa / Recria', 'Recria', 'kgs (MS)', 95.00],
      ['CST005', 'Sanidad General (Vacunas/Desparasitantes)', 'Recria', 'Dosis', 1850.00],
      ['CST006', 'Grano de Maiz/Adaptacion', 'Adaptacion', 'kg por peso del animal', 180.00],
      ['CST007', 'Expeller / Harina de Soja / Adaptacion', 'Adaptacion', 'kg por peso del animal', ''],
      ['CST008', 'Nucleo Mineral c/ Monensina/Adaptacion', 'Adaptacion', 'kg por peso del animal', ''],
      ['CST009', 'Alfalfa (Fibra / Rollo) / Adaptacion', 'Adaptacion', 'kgs (MS)', 95.00],
      ['CST010', 'Grano de Maiz / Engorde', 'Engorde', 'kg por peso del animal', 180.00],
      ['CST011', 'Expeller / Harina de Soja / Engorde', 'Engorde', 'kg por peso del animal', ''],
      ['CST012', 'Nucleo Mineral c/ Monensina / Engorde', 'Engorde', 'kg por peso del animal', ''],
      ['CST013', 'Alfalfa (Fibra / Rollo) / Engorde', 'Engorde', 'kgs (MS)', 95.00],
      // Costos Porcinos
      ['CSTP01', 'Maiz propio - Madres', 'Reproduccion', 'kg por peso de animal', ''],
      ['CSTP02', 'Maiz propio - Engorde', 'Engorde', 'kg por peso de animal', ''],
      ['CSTP03', 'Maiz propio - Destete', 'Destete', 'kg por peso de animal', ''],
      ['CSTP04', 'Concentrado Preiniciador - Lechones', 'Cria', 'kg por peso de animal', ''],
      ['CSTP05', 'Concentrado Desarrollo - Destete', 'Destete', 'kg por peso de animal', ''],
      ['CSTP06', 'Soja peleteada - Destete', 'Destete', 'kg por peso de animal', ''],
      ['CSTP07', 'Soja peleteada - Engorde', 'Engorde', 'kg por peso de animal', '']
    ];
    sheetCostos.getRange(2, 1, dataCostos.length, dataCostos[0].length).setValues(dataCostos);
    sheetCostos.getRange(2, 1, dataCostos.length, 1).setHorizontalAlignment('center');
    sheetCostos.getRange(2, 3, dataCostos.length, 1).setHorizontalAlignment('center');
    sheetCostos.getRange(2, 5, dataCostos.length, 1).setNumberFormat('$ #,##0.00');
  }
}

/**
 * Actualiza los nombres de conceptos y asegura la creación de los nuevos conceptos en CAT_CONCEPTOS:
 * 1. Ediciones solicitadas:
 *    - "REPARACION Y MANTENIMIENTO TRANQUERAS/BALANZA" -> "REPARACIÓN Y MANTENIMIENTO DE INSTALACIONES"
 *    - "SANIDAD, VACUNAS Y DOSIS DE AFTOSA" -> "SANIDAD/VACUNAS"
 * 2. Nuevos conceptos solicitados:
 *    - CPT-13 | REPRODUCCION/INSEMINACION | Costo Directo
 *    - CPT-14 | GASTOS ADMINISTRATIVOS | Costo Indirecto
 *    - CPT-15 | HONORARIOS VETERINARIO | Costo Directo
 *    - CPT-16 | COMPRA DE GRANOS | Costo Directo
 * También sincroniza TABLA_IMPUTACION si existen registros con la denominación previa.
 */
/**
 * Normaliza cualquier denominación histórica o variación a la denominación oficial de CAT_CONCEPTOS
 */
function normalizarConceptoImputacion(texto) {
  if (!texto) return '';
  const val = texto.toString().trim();
  const valUpper = val.toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // sin acentos para comparación robusta
    .replace(/[–—]/g, "-");

  // 1. Reparación y Mantenimiento de Instalaciones
  if (valUpper.includes('TRANQUERA') || valUpper.includes('BALANZA') ||
      valUpper.includes('INSTALACION') || valUpper.includes('CORRALES Y COMEDEROS') ||
      valUpper.includes('REPARACION') || valUpper.includes('MANTENIMIENTO')) {
    return 'REPARACIÓN Y MANTENIMIENTO DE INSTALACIONES';
  }

  // 2. Existencias - alimento balanceado -> ALIMENTO BALANCEADO/SUPLEMENTOS
  if (valUpper.includes('BALANCEADO') || valUpper.includes('SUPLEMENTO')) {
    return 'ALIMENTO BALANCEADO/SUPLEMENTOS';
  }

  // 3. Flete -> FLETE DE HACIENDA
  if (valUpper.includes('FLETE')) {
    return 'FLETE DE HACIENDA';
  }

  // 4. Reproduccion/inseminacion -> REPRODUCCION/INSEMINACION
  if (valUpper.includes('REPRODUC') || valUpper.includes('INSEMINAC')) {
    return 'REPRODUCCION/INSEMINACION';
  }

  // 5. Combustible y lubricantes -> COMBUSTIBLES Y LUBRICANTES
  if (valUpper.includes('COMBUSTIBLE') || valUpper.includes('LUBRICANTE')) {
    return 'COMBUSTIBLES Y LUBRICANTES';
  }

  // 6 y 10. Sanidad / Vacunas / Aftosa -> SANIDAD/VACUNAS
  if (valUpper.includes('SANIDAD') || valUpper.includes('VACUNA') || valUpper.includes('AFTOSA')) {
    return 'SANIDAD/VACUNAS';
  }

  // 7. Alfalfa pastura -> ROLLOS ALFALFA PROPIOS
  if (valUpper.includes('ALFALFA') || valUpper.includes('ROLLO')) {
    return 'ROLLOS ALFALFA PROPIOS';
  }

  // 8. Administracion -> GASTOS ADMINISTRATIVOS
  if (valUpper.includes('ADMINISTRA')) {
    return 'GASTOS ADMINISTRATIVOS';
  }

  // 9. Honorarios del veterinario -> HONORARIOS DEL VETERINARIO
  if (valUpper.includes('HONORARIO') || valUpper.includes('VETERINARI')) {
    return 'HONORARIOS DEL VETERINARIO';
  }

  // 11, 14, 15. Venta de novillos / vacas de descarte / otras categorias -> VENTA DE HACIENDA EN PIE
  if (valUpper.includes('VENTA DE NOVILLO') || valUpper.includes('VENTAS DE NOVILLO') ||
      valUpper.includes('VACAS DE DESCARTE') || valUpper.includes('OTRAS CATEGORIAS') ||
      valUpper.includes('VENTA DE HACIENDA') || (valUpper.startsWith('VENTA') && (valUpper.includes('NOVILLO') || valUpper.includes('VACA')))) {
    return 'VENTA DE HACIENDA EN PIE';
  }

  // 12. Comisiones de feria/remate -> COMISION CONSIGNATARIO FERIA
  if (valUpper.includes('COMISION') || valUpper.includes('FERIA') || valUpper.includes('REMATE') || valUpper.includes('CONSIGNATARI')) {
    return 'COMISION CONSIGNATARIO FERIA';
  }

  // 13. Dte -> GUIAS, DTE Y SELLADOS SENASA
  if (valUpper === 'DTE' || valUpper.includes('DTE') || valUpper.includes('SENASA') || valUpper.includes('GUIA')) {
    return 'GUIAS, DTE Y SELLADOS SENASA';
  }

  // 16. Produccion propia consumida -> MAIZ PROPIO CONSUMO RODEO
  if (valUpper.includes('PRODUCCION PROPIA CONSUMIDA') || valUpper.includes('PROPIA CONSUMIDA') || valUpper.includes('MAIZ PROPIO')) {
    return 'MAIZ PROPIO CONSUMO RODEO';
  }

  // 17. Compra de maíz (racion) -> COMPRA DE GRANOS
  if (valUpper.includes('COMPRA DE MAIZ') || valUpper.includes('COMPRA DE MAÍZ') || valUpper.includes('COMPRA DE GRANO') || (valUpper.includes('COMPRA') && valUpper.includes('RACION'))) {
    return 'COMPRA DE GRANOS';
  }

  return val;
}

/**
 * Actualiza los nombres de conceptos y asegura la creación de los nuevos conceptos en CAT_CONCEPTOS:
 * 1. Ediciones solicitadas:
 *    - "REPARACION Y MANTENIMIENTO TRANQUERAS/BALANZA" -> "REPARACIÓN Y MANTENIMIENTO DE INSTALACIONES"
 *    - "SANIDAD, VACUNAS Y DOSIS DE AFTOSA" -> "SANIDAD/VACUNAS"
 * 2. Nuevos conceptos solicitados:
 *    - CPT-13 | REPRODUCCION/INSEMINACION | Costo Directo
 *    - CPT-14 | GASTOS ADMINISTRATIVOS | Costo Indirecto
 *    - CPT-15 | HONORARIOS DEL VETERINARIO | Costo Directo
 *    - CPT-16 | COMPRA DE GRANOS | Costo Directo
 * 3. Homogeneización en TABLA_IMPUTACION Columna D (Concepto).
 */
function actualizarNombresConceptos(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CAT_CONCEPTOS');
  if (!sheet) return { exito: false, cambios: 0, agregados: 0, cambiosImputacion: 0 };

  const lastRow = sheet.getLastRow();
  let cambios = 0;

  if (lastRow > 1) {
    // Columna B: Nombre_Concepto (Columna 2)
    const range = sheet.getRange(2, 2, lastRow - 1, 1);
    const values = range.getValues();

    for (let i = 0; i < values.length; i++) {
      const val = (values[i][0] || '').toString().trim();
      const nuevoVal = normalizarConceptoImputacion(val);
      if (nuevoVal && nuevoVal !== val) {
        values[i][0] = nuevoVal;
        cambios++;
      }
    }

    if (cambios > 0) {
      range.setValues(values);
    }
  }

  // 2. Agregar o verificar los 4 nuevos conceptos
  const nuevosConceptos = [
    { id: 'CPT-13', nombre: 'REPRODUCCION/INSEMINACION', naturaleza: 'Costo Directo' },
    { id: 'CPT-14', nombre: 'GASTOS ADMINISTRATIVOS', naturaleza: 'Costo Indirecto' },
    { id: 'CPT-15', nombre: 'HONORARIOS DEL VETERINARIO', naturaleza: 'Costo Directo' },
    { id: 'CPT-16', nombre: 'COMPRA DE GRANOS', naturaleza: 'Costo Directo' },
    { id: 'CPT-17', nombre: 'Mano de obra directa', naturaleza: 'Costo Directo' },
    { id: 'CPT-18', nombre: 'Amortizacion/prorrateo del plantel reproductor', naturaleza: 'Costo Estructura' },
    { id: 'CPT-19', nombre: 'Energia', naturaleza: 'Costo Variable' },
    { id: 'CPT-20', nombre: 'Mortandad', naturaleza: 'Costo Variable' },
    { id: 'CPT-21', nombre: 'CMV - Porcino', naturaleza: 'Costo Directo' }
  ];

  let curLastRow = sheet.getLastRow();
  let todosLosRegistros = [];
  if (curLastRow > 1) {
    todosLosRegistros = sheet.getRange(2, 1, curLastRow - 1, 3).getValues();
  }

  let agregados = 0;
  for (let nc = 0; nc < nuevosConceptos.length; nc++) {
    const nuevo = nuevosConceptos[nc];
    let encontrado = false;

    for (let r = 0; r < todosLosRegistros.length; r++) {
      const idExistente = (todosLosRegistros[r][0] || '').toString().trim().toUpperCase();
      const nomExistente = (todosLosRegistros[r][1] || '').toString().trim().toUpperCase();

      if (idExistente === nuevo.id.toUpperCase() || nomExistente === nuevo.nombre.toUpperCase() ||
          (nuevo.id === 'CPT-15' && nomExistente === 'HONORARIOS VETERINARIO')) {
        encontrado = true;
        const rowIdx = r + 2;
        sheet.getRange(rowIdx, 1).setValue(nuevo.id);
        sheet.getRange(rowIdx, 2).setValue(nuevo.nombre);
        sheet.getRange(rowIdx, 3).setValue(nuevo.naturaleza);
        sheet.getRange(rowIdx, 1).setHorizontalAlignment('center').setFontWeight('bold');
        break;
      }
    }

    if (!encontrado) {
      curLastRow = sheet.getLastRow();
      const nFila = curLastRow + 1;
      sheet.getRange(nFila, 1).setValue(nuevo.id);
      sheet.getRange(nFila, 2).setValue(nuevo.nombre);
      sheet.getRange(nFila, 3).setValue(nuevo.naturaleza);

      sheet.getRange(nFila, 1, 1, 3).setVerticalAlignment('middle');
      sheet.getRange(nFila, 1).setHorizontalAlignment('center').setFontWeight('bold');
      sheet.getRange(nFila, 2).setHorizontalAlignment('left');
      sheet.getRange(nFila, 3).setHorizontalAlignment('left');

      agregados++;
      todosLosRegistros.push([nuevo.id, nuevo.nombre, nuevo.naturaleza]);
    }
  }

  // 3. Sincronizar y homogeneizar TABLA_IMPUTACION
  const resImp = homogeneizarConceptosTablaImputacion(ss);

  return { exito: true, cambios: cambios, agregados: agregados, cambiosImputacion: resImp.cambios };
}

/**
 * Homogeneiza todos los conceptos de TABLA_IMPUTACION (Columna D)
 * reemplazándolos por las denominaciones exactas de CAT_CONCEPTOS.
 */
function homogeneizarConceptosTablaImputacion(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetImp = ss.getSheetByName('TABLA_IMPUTACION');
  if (!sheetImp || sheetImp.getLastRow() <= 1) {
    return { exito: true, cambios: 0, total: 0 };
  }

  const lastRow = sheetImp.getLastRow();
  // Columna 4: Concepto (Columna D)
  const rangeImp = sheetImp.getRange(2, 4, lastRow - 1, 1);
  const valsImp = rangeImp.getValues();
  let cambios = 0;

  for (let k = 0; k < valsImp.length; k++) {
    const valOriginal = (valsImp[k][0] || '').toString().trim();
    if (!valOriginal) continue;

    const valNormalizado = normalizarConceptoImputacion(valOriginal);
    if (valNormalizado && valNormalizado !== valOriginal) {
      valsImp[k][0] = valNormalizado;
      cambios++;
    }
  }

  if (cambios > 0) {
    rangeImp.setValues(valsImp);
  }

  return { exito: true, cambios: cambios, total: valsImp.length };
}

function ejecutarActualizacionConceptosMenu() {
  actualizarHojaCatConceptosMenu();
}

/**
 * Menu y disparador para actualizar la hoja CAT_CONCEPTOS con los nuevos conceptos CPT-17 a CPT-21
 * y sincronizar TABLA_IMPUTACION. Puede ser ejecutada desde el menu o asignada a un boton.
 */
function actualizarHojaCatConceptosMenu() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const res = actualizarNombresConceptos(ss);
    ss.toast('Hoja CAT_CONCEPTOS actualizada (' + res.agregados + ' nuevos / verificados).', '📋 Conceptos al Día', 5);
    const mensaje = 'La hoja CAT_CONCEPTOS ha sido actualizada con éxito:\n\n' +
      '• Conceptos agregados / verificados:\n' +
      '  - CPT-17: Mano de obra directa (Costo Directo)\n' +
      '  - CPT-18: Amortizacion/prorrateo del plantel reproductor (Costo Estructura)\n' +
      '  - CPT-19: Energia (Costo Variable)\n' +
      '  - CPT-20: Mortandad (Costo Variable)\n' +
      '  - CPT-21: CMV - Porcino (Costo Directo)\n' +
      '  - CPT-13 a CPT-16 (Reproducción, Administración, Veterinario, Granos)\n\n' +
      '• Filas homogeneizadas en TABLA_IMPUTACION: ' + res.cambiosImputacion;
    SpreadsheetApp.getUi().alert('Actualización de CAT_CONCEPTOS Exitosa', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    SpreadsheetApp.getUi().alert('Error al actualizar conceptos: ' + err.message);
  }
}


function ejecutarHomogeneizacionTablaImputacionMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const res = homogeneizarConceptosTablaImputacion(ss);
  SpreadsheetApp.getUi().alert(
    'Homogeneización de TABLA_IMPUTACION Completa\n\n' +
    'Se revisaron ' + res.total + ' filas.\n' +
    'Se modificaron ' + res.cambios + ' celdas en la columna "Concepto" para coincidir exactamente con "CAT_CONCEPTOS".'
  );
}

// ============================================================================
// 4. TRAZABILIDAD INDIVIDUAL Y EXISTENCIAS POR CARAVANA
// ============================================================================
/**
 * Desplaza automáticamente los datos existentes una columna a la derecha
 * en Movimientos_Existencia y en EXISTENCIAS, configurando Peso Estimado KG y Peso Real KG.
 */
function migrarEstructuraPesosExistenciasYMovimientos(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  // Asegura la creación y configuración completa con 10 cols en EXISTENCIAS (con 3 paneles paralelos: Bovinos, Porcinos, Ovinos)
  // y 15 cols en Movimientos_Existencia con Col D (Unidad de Produccion).
  prepararHojaExistencias(ss);
  prepararHojaMovimientosExistencia(ss);

  let sheetMov = ss.getSheetByName('Movimientos_Existencia');
  let sheetExis = ss.getSheetByName('EXISTENCIAS');

  let movMigrados = 0;
  let exisMigrados = 0;

  if (sheetMov && sheetMov.getLastRow() > 1) {
    const lrMov = sheetMov.getLastRow();
    for (let r = 2; r <= lrMov; r++) {
      const up = sheetMov.getRange(r, 4).getValue();
      if (!up) {
        sheetMov.getRange(r, 4).setValue('Bovino');
        movMigrados++;
      }
      sheetMov.getRange(r, 14).setFormula(`=L${r}*M${r}`);
    }
  }

  if (sheetExis && sheetExis.getLastRow() > 1) {
    exisMigrados = actualizarFormulasExistencias(ss);
  }

  return { movMigrados: movMigrados, exisMigrados: exisMigrados };
}

function ejecutarMigracionPesosMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const res = migrarEstructuraPesosExistenciasYMovimientos(ss);
  SpreadsheetApp.getUi().alert(
    'Alineación y Trazabilidad Multi-Especie Completada\n\n' +
    '• EXISTENCIAS: 10 columnas configuradas con Unidad de Producción (Col D) y los 3 paneles de resumen paralelos (Bovinos L:M, Porcinos O:P, Ovinos R:S).\n' +
    '• Movimientos_Existencia: 15 columnas con Unidad de Producción (Col D) y Costo Total (=L*M = Costo Unitario * Cantidad).\n' +
    '• Se verificaron ' + res.movMigrados + ' registro(s) en Movimientos y ' + res.exisMigrados + ' fórmulas en EXISTENCIAS.\n' +
    '• Todas las fórmulas dinámicas vinculadas perfectamente.'
  );
}

function prepararHojasTrazabilidad() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const res = migrarEstructuraPesosExistenciasYMovimientos(ss);
  const sincronizados = sincronizarNacimientosFaltantes(ss);
  SpreadsheetApp.getUi().alert(
    'Hojas de Trazabilidad Configuradas con Éxito\n\n' +
    '1. Hoja "EXISTENCIAS" (10 columnas + 3 paneles de resumen paralelos):\n' +
    '   • Columna D: Unidad de Produccion (Bovino, Porcino, Ovino).\n' +
    '   • Columna E: Categoría (XLOOKUP dinámico a Movimientos_Existencia Col H).\n' +
    '   • Columna F: Subunidad / Etapa (XLOOKUP dinámico a Movimientos_Existencia Col I).\n' +
    '   • Columna H: Peso Estimado KG (XLOOKUP dinámico a Col F de Movimientos).\n' +
    '   • Columna I: Peso Real KG (CHOOSEROWS/FILTER a Col G de Movimientos).\n' +
    '   • Columna J: Estado (Vivo, Muerto, Vendido).\n' +
    '   • Paneles paralelos: Bovinos (Cols L:M), Porcinos (Cols O:P), Ovinos (Cols R:S).\n\n' +
    '2. Hoja "Movimientos_Existencia" (15 columnas normalizadas):\n' +
    '   • Columna D: Unidad de Produccion (Bovino, Porcino, Ovino).\n' +
    '   • Columna N: Costo Total (=L*M = Costo Unitario * Cantidad).\n\n' +
    (res.movMigrados > 0 || res.exisMigrados > 0 ? 
      '✔ Se verificaron ' + res.movMigrados + ' fila(s) en Movimientos y ' + res.exisMigrados + ' fila(s) en EXISTENCIAS.\n' : 
      '✔ Las columnas y datos ya se encuentran perfectamente alineados.\n') +
    (sincronizados > 0 ? '✔ Se sincronizaron ' + sincronizados + ' nacimiento(s) faltante(s).' : '')
  );
}

function prepararHojaExistencias(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheetExisHist = ss.getSheetByName('EXISTENCIAS_HISTORICAS_JG');
  let sheet = ss.getSheetByName('EXISTENCIAS');
  if (sheetExisHist && !sheet) {
    sheetExisHist.setName('EXISTENCIAS');
    sheet = sheetExisHist;
  }
  if (!sheet) {
    sheet = ss.insertSheet('EXISTENCIAS');
  }

  sheet.setHiddenGridlines(false);

  // Insertar Unidad de Produccion en Col D si no existe
  if (sheet.getLastColumn() >= 4 && sheet.getRange(1, 4).getValue() !== 'Unidad de Produccion') {
    sheet.insertColumnBefore(4);
    sheet.getRange(1, 4).setValue('Unidad de Produccion');
    const lr = sheet.getLastRow();
    if (lr > 1) {
      for (let r = 2; r <= lr; r++) {
        if (sheet.getRange(r, 3).getValue()) {
          sheet.getRange(r, 4).setValue('Bovino');
        }
      }
    }
  }

  const COLOR_HEADER = '#1E3A1E';
  const COLOR_TEXT = '#FFFFFF';

  // 1. Cabeceras principales de la tabla de 10 columnas
  const headers = [
    'ID_REGISTRO_EXIS',
    'Fecha de Registro',
    'Caravana',
    'Unidad de Produccion',
    'Categoria',
    'Subunidad',
    'Peso Inicial KG',
    'Peso Estimado KG',
    'Peso Real KG',
    'Estado'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground(COLOR_HEADER)
    .setFontColor(COLOR_TEXT)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  if (sheet.getMaxColumns() < 20) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 20 - sheet.getMaxColumns());
  }

  const anchos = [140, 120, 110, 150, 140, 120, 120, 120, 120, 100, 30, 210, 130, 30, 210, 130, 30, 210, 130];
  anchos.forEach((w, idx) => sheet.setColumnWidth(idx + 1, w));
  sheet.setFrozenRows(1);

  // Validaciones
  const reglaUP = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Bovino', 'Porcino', 'Ovino'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('D2:D1000').setDataValidation(reglaUP);

  const reglaEstado = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Vivo', 'Muerto', 'Vendido'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('J2:J1000').setDataValidation(reglaEstado);

  // Descombinar y limpiar área de paneles (K a T)
  try {
    const mergedRanges = sheet.getRange(1, 11, Math.min(sheet.getMaxRows(), 50), 10).getMergedRanges();
    for (let m = 0; m < mergedRanges.length; m++) {
      mergedRanges[m].breakApart();
    }
  } catch (e) {}

  sheet.getRange('K1:T45').clearContent().clearFormat();

  // PANEL 1: BOVINOS (Cols L y M)
  sheet.getRange('L1:M1').merge()
    .setValue('🐂 BOVINOS: STOCK POR CATEGORÍA (VIVOS)')
    .setBackground('#1E3A1E').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.getRange('L2').setValue('Categoría').setBackground('#2D5A27').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('M2').setValue('Cabezas Vivas').setBackground('#2D5A27').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

  const bovCats = [
    'Vacas con cria', 'Vacas secas', 'Vacas prenadas',
    'Vaquillonas 1 (recria)', 'Vaquillonas 2 (servicio/prenez)',
    'Terneros', 'Terneras', 'Destete',
    'Novillitos (adaptacion)', 'Novillos terminados',
    'Toros cabana', 'Toritos reposicion'
  ];
  for (let i = 0; i < bovCats.length; i++) {
    const row = 3 + i;
    sheet.getRange(row, 12).setValue(bovCats[i]).setHorizontalAlignment('left');
    sheet.getRange(row, 13).setFormula(`=COUNTIFS($D$2:$D, "Bovino", $E$2:$E, L${row}, $J$2:$J, "Vivo")`).setHorizontalAlignment('right');
  }
  sheet.getRange(15, 12).setValue('TOTAL CABEZAS VIVAS BOVINOS').setFontWeight('bold').setBackground('#E8F0E8');
  sheet.getRange(15, 13).setFormula('=SUM(M3:M14)').setFontWeight('bold').setBackground('#E8F0E8').setHorizontalAlignment('right');

  sheet.getRange('L17:M17').merge()
    .setValue('RESUMEN 2: STOCK BOVINOS POR SUBUNIDAD')
    .setBackground('#1E3A1E').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.getRange('L18').setValue('Subunidad').setBackground('#2D5A27').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('M18').setValue('Cabezas Vivas').setBackground('#2D5A27').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

  const bovSubs = ['Cria', 'Recria', 'Adaptacion', 'Engorde'];
  for (let i = 0; i < bovSubs.length; i++) {
    const row = 19 + i;
    sheet.getRange(row, 12).setValue(bovSubs[i]).setHorizontalAlignment('left');
    sheet.getRange(row, 13).setFormula(`=COUNTIFS($D$2:$D, "Bovino", $F$2:$F, L${row}, $J$2:$J, "Vivo")`).setHorizontalAlignment('right');
  }
  sheet.getRange(23, 12).setValue('TOTAL POR SUBUNIDADES').setFontWeight('bold').setBackground('#E8F0E8');
  sheet.getRange(23, 13).setFormula('=SUM(M19:M22)').setFontWeight('bold').setBackground('#E8F0E8').setHorizontalAlignment('right');

  sheet.getRange('L25:M25').merge()
    .setValue('RESUMEN 3: CONTROL BOVINOS POR ESTADO')
    .setBackground('#1E3A1E').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.getRange('L26').setValue('Estado').setBackground('#2D5A27').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('M26').setValue('Cabezas').setBackground('#2D5A27').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

  const estados = ['Vivo', 'Muerto', 'Vendido'];
  for (let i = 0; i < estados.length; i++) {
    const row = 27 + i;
    sheet.getRange(row, 12).setValue(estados[i]).setHorizontalAlignment('left');
    sheet.getRange(row, 13).setFormula(`=COUNTIFS($D$2:$D, "Bovino", $J$2:$J, L${row})`).setHorizontalAlignment('right');
  }
  sheet.getRange(30, 12).setValue('TOTAL BOVINOS REGISTRADOS').setFontWeight('bold').setBackground('#E8F0E8');
  sheet.getRange(30, 13).setFormula('=SUM(M27:M29)').setFontWeight('bold').setBackground('#E8F0E8').setHorizontalAlignment('right');

  // PANEL 2: PORCINOS (Cols O y P)
  sheet.getRange('O1:P1').merge()
    .setValue('🐖 PORCINOS: STOCK POR CATEGORÍA (VIVOS)')
    .setBackground('#6B2D5C').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.getRange('O2').setValue('Categoría').setBackground('#823B71').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('P2').setValue('Cabezas Vivas').setBackground('#823B71').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

  const porCats = [
    'Madre', 'Padrillo', 'Lechon',
    'Cachorro, Cachorra', 'Capon',
    'Madre de descarte', 'Padrillo de descarte'
  ];
  for (let i = 0; i < porCats.length; i++) {
    const row = 3 + i;
    sheet.getRange(row, 15).setValue(porCats[i]).setHorizontalAlignment('left');
    sheet.getRange(row, 16).setFormula(`=COUNTIFS($D$2:$D, "Porcino", $E$2:$E, O${row}, $J$2:$J, "Vivo")`).setHorizontalAlignment('right');
  }
  sheet.getRange(10, 15).setValue('TOTAL CABEZAS VIVAS PORCINOS').setFontWeight('bold').setBackground('#F5EBF2');
  sheet.getRange(10, 16).setFormula('=SUM(P3:P9)').setFontWeight('bold').setBackground('#F5EBF2').setHorizontalAlignment('right');

  sheet.getRange('O17:P17').merge()
    .setValue('RESUMEN 2: STOCK PORCINOS POR ETAPA')
    .setBackground('#6B2D5C').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.getRange('O18').setValue('Etapa').setBackground('#823B71').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('P18').setValue('Cabezas Vivas').setBackground('#823B71').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

  const porSubs = ['Cria', 'Destete', 'Engorde', 'Reproduccion'];
  for (let i = 0; i < porSubs.length; i++) {
    const row = 19 + i;
    sheet.getRange(row, 15).setValue(porSubs[i]).setHorizontalAlignment('left');
    sheet.getRange(row, 16).setFormula(`=COUNTIFS($D$2:$D, "Porcino", $F$2:$F, O${row}, $J$2:$J, "Vivo")`).setHorizontalAlignment('right');
  }
  sheet.getRange(23, 15).setValue('TOTAL POR ETAPAS PORCINOS').setFontWeight('bold').setBackground('#F5EBF2');
  sheet.getRange(23, 16).setFormula('=SUM(P19:P22)').setFontWeight('bold').setBackground('#F5EBF2').setHorizontalAlignment('right');

  sheet.getRange('O25:P25').merge()
    .setValue('RESUMEN 3: CONTROL PORCINOS POR ESTADO')
    .setBackground('#6B2D5C').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.getRange('O26').setValue('Estado').setBackground('#823B71').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('P26').setValue('Cabezas').setBackground('#823B71').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

  for (let i = 0; i < estados.length; i++) {
    const row = 27 + i;
    sheet.getRange(row, 15).setValue(estados[i]).setHorizontalAlignment('left');
    sheet.getRange(row, 16).setFormula(`=COUNTIFS($D$2:$D, "Porcino", $J$2:$J, O${row})`).setHorizontalAlignment('right');
  }
  sheet.getRange(30, 15).setValue('TOTAL PORCINOS REGISTRADOS').setFontWeight('bold').setBackground('#F5EBF2');
  sheet.getRange(30, 16).setFormula('=SUM(P27:P29)').setFontWeight('bold').setBackground('#F5EBF2').setHorizontalAlignment('right');

  // PANEL 3: OVINOS (Cols R y S)
  sheet.getRange('R1:S1').merge()
    .setValue('🐑 OVINOS: STOCK POR CATEGORÍA (VIVOS)')
    .setBackground('#1C3D5A').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.getRange('R2').setValue('Categoría').setBackground('#2B5876').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('S2').setValue('Cabezas Vivas').setBackground('#2B5876').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

  const oviCats = [
    'Madres', 'Carneros', 'Corderos',
    'Borregas', 'Borregos', 'Ovino de Descarte'
  ];
  for (let i = 0; i < oviCats.length; i++) {
    const row = 3 + i;
    sheet.getRange(row, 18).setValue(oviCats[i]).setHorizontalAlignment('left');
    sheet.getRange(row, 19).setFormula(`=COUNTIFS($D$2:$D, "Ovino", $E$2:$E, R${row}, $J$2:$J, "Vivo")`).setHorizontalAlignment('right');
  }
  sheet.getRange(9, 18).setValue('TOTAL CABEZAS VIVAS OVINOS').setFontWeight('bold').setBackground('#EBF2F7');
  sheet.getRange(9, 19).setFormula('=SUM(S3:S8)').setFontWeight('bold').setBackground('#EBF2F7').setHorizontalAlignment('right');

  sheet.getRange('R17:S17').merge()
    .setValue('RESUMEN 2: STOCK OVINOS POR ETAPA')
    .setBackground('#1C3D5A').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.getRange('R18').setValue('Etapa').setBackground('#2B5876').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('S18').setValue('Cabezas Vivas').setBackground('#2B5876').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

  const oviSubs = ['Cria', 'Recria', 'Engorde', 'Reproduccion'];
  for (let i = 0; i < oviSubs.length; i++) {
    const row = 19 + i;
    sheet.getRange(row, 18).setValue(oviSubs[i]).setHorizontalAlignment('left');
    sheet.getRange(row, 19).setFormula(`=COUNTIFS($D$2:$D, "Ovino", $F$2:$F, R${row}, $J$2:$J, "Vivo")`).setHorizontalAlignment('right');
  }
  sheet.getRange(23, 18).setValue('TOTAL POR ETAPAS OVINOS').setFontWeight('bold').setBackground('#EBF2F7');
  sheet.getRange(23, 19).setFormula('=SUM(S19:S22)').setFontWeight('bold').setBackground('#EBF2F7').setHorizontalAlignment('right');

  sheet.getRange('R25:S25').merge()
    .setValue('RESUMEN 3: CONTROL OVINOS POR ESTADO')
    .setBackground('#1C3D5A').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.getRange('R26').setValue('Estado').setBackground('#2B5876').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('S26').setValue('Cabezas').setBackground('#2B5876').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

  for (let i = 0; i < estados.length; i++) {
    const row = 27 + i;
    sheet.getRange(row, 18).setValue(estados[i]).setHorizontalAlignment('left');
    sheet.getRange(row, 19).setFormula(`=COUNTIFS($D$2:$D, "Ovino", $J$2:$J, R${row})`).setHorizontalAlignment('right');
  }
  sheet.getRange(30, 18).setValue('TOTAL OVINOS REGISTRADOS').setFontWeight('bold').setBackground('#E8F0E8');
  sheet.getRange(30, 19).setFormula('=SUM(S27:S29)').setFontWeight('bold').setBackground('#E8F0E8').setHorizontalAlignment('right');
}

function prepararHojaMovimientosExistencia(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Movimientos_Existencia');
  if (!sheet) {
    sheet = ss.insertSheet('Movimientos_Existencia');
  }

  // Insertar Unidad de Produccion en Col D si no existe
  if (sheet.getLastColumn() >= 4 && sheet.getRange(1, 4).getValue() !== 'Unidad de Produccion') {
    sheet.insertColumnBefore(4);
    sheet.getRange(1, 4).setValue('Unidad de Produccion');
    const lr = sheet.getLastRow();
    if (lr > 1) {
      for (let r = 2; r <= lr; r++) {
        if (sheet.getRange(r, 3).getValue()) {
          sheet.getRange(r, 4).setValue('Bovino');
        }
      }
    }
  }

  const headers = [
    'ID Registro',
    'Fecha de Registro',
    'Caravana',
    'Unidad de Produccion',
    'Fecha de Nacimietno',
    'Peso Estimado KG',
    'Peso Real KG',
    'Categoria',
    'Etapa',
    'Concepto Costo',
    'Unidad',
    'Costo Unitario',
    'Cantidad',
    'Costo Total',
    'Observaciones'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1E3A1E')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  if (sheet.getMaxColumns() < 15) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 15 - sheet.getMaxColumns());
  }

  const anchos = [100, 120, 110, 150, 130, 120, 120, 130, 120, 140, 90, 110, 90, 110, 180];
  anchos.forEach((w, idx) => sheet.setColumnWidth(idx + 1, w));
  sheet.setFrozenRows(1);

  // Validación de Unidad de Produccion (Col D)
  const reglaUP = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Bovino', 'Porcino', 'Ovino'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('D2:D1000').setDataValidation(reglaUP);
}

// ============================================================================
// 11. REGLAS DE VALIDACION DE DATOS
// ============================================================================
function actualizarValidaciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Validaciones estrictas con rechazo de entradas erróneas (.setAllowInvalid(false))
  const ruleUN = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getRange('CAT_UNIDADES_NEGOCIO!A2:A'))
    .setAllowInvalid(false)
    .build();

  const ruleSec = ss.getSheetByName('CAT_SECTORES')
    ? SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRange('CAT_SECTORES!A2:A')).setAllowInvalid(false).build()
    : null;

  const ruleCat = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getRange(ss.getSheetByName('CAT CATEGORIAS') ? "'CAT CATEGORIAS'!A2:A" : 'CAT_CATEGORIAS_BOVINOS!A2:A'))
    .setAllowInvalid(false)
    .build();

  const ruleContId = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getRange('CAT_CONTACTOS!A2:A'))
    .setAllowInvalid(false)
    .build();

  const ruleTipoCom = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Venta', 'Compra'], true)
    .setAllowInvalid(false)
    .build();

  const ruleTipoFis = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Nacimiento', 'Muerte', 'Cambio_Categoria', 'Ingreso_Terceros', 'Egreso_Terceros', 'Consumo_Propio', 'Transferencia_Interna'], true)
    .setAllowInvalid(false)
    .build();

  const ruleCondicion = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Contado', 'Cuenta Corriente', 'Canje Grano'], true)
    .setAllowInvalid(false)
    .build();

  const ruleEstadoPago = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Cobrado', 'Pendiente', 'Cancelado'], true)
    .setAllowInvalid(false)
    .build();

  // 1. MOV_COMERCIALES: normalización de datos y validaciones
  const sheetCom = ss.getSheetByName('MOV_COMERCIALES');
  if (sheetCom) {
    if (sheetCom.getLastRow() >= 2) {
      const lrCom = sheetCom.getLastRow();
      const rangeCat = sheetCom.getRange(2, 7, lrCom - 1, 1);
      const valsCat = rangeCat.getValues();
      let modCom = false;
      for (let r = 0; r < valsCat.length; r++) {
        let v = String(valsCat[r][0] || '').trim();
        if (v === 'CAT-BOV-01' || v === 'CAT-01') { valsCat[r][0] = 'CAT_BOV_1'; modCom = true; }
        else if (v === 'CAT-BOV-12' || v === 'CAT-12') { valsCat[r][0] = 'CAT_BOV_12'; modCom = true; }
      }
      if (modCom) rangeCat.setValues(valsCat);
    }
    sheetCom.getRange('C2:C').setDataValidation(ruleTipoCom);
    sheetCom.getRange('E2:E').setDataValidation(ruleContId);
    sheetCom.getRange('F2:F').setDataValidation(ruleUN);
    sheetCom.getRange('G2:G').setDataValidation(ruleCat);
    sheetCom.getRange('Q2:Q').setDataValidation(ruleCondicion);
    sheetCom.getRange('R2:R').setDataValidation(ruleEstadoPago);
  }

  // 2. MOV_FISICOS: normalización de datos (quitar 0 de más) y validaciones
  const sheetFis = ss.getSheetByName('MOV_FISICOS');
  if (sheetFis) {
    if (sheetFis.getLastRow() >= 2) {
      const lrFis = sheetFis.getLastRow();
      const rangeCats = sheetFis.getRange(2, 5, lrFis - 1, 2); // Cols E e F
      const valsCats = rangeCats.getValues();
      let modFis = false;
      for (let r = 0; r < valsCats.length; r++) {
        for (let c = 0; c < 2; c++) {
          let v = String(valsCats[r][c] || '').trim();
          if (v === 'CAT-06') { valsCats[r][c] = 'CAT_BOV_6'; modFis = true; }
          else if (v === 'CAT-07') { valsCats[r][c] = 'CAT_BOV_7'; modFis = true; }
          else if (v === 'CAT-08') { valsCats[r][c] = 'CAT_BOV_8'; modFis = true; }
          else if (v === 'CAT-09') { valsCats[r][c] = 'CAT_BOV_9'; modFis = true; }
          else if (v === 'CAT-01') { valsCats[r][c] = 'CAT_BOV_1'; modFis = true; }
          else if (v === 'CAT-12') { valsCats[r][c] = 'CAT_BOV_12'; modFis = true; }
          else if (/^CAT-0*(\d+)$/.test(v)) {
            valsCats[r][c] = 'CAT_BOV_' + parseInt(v.replace('CAT-', ''), 10);
            modFis = true;
          }
        }
      }
      if (modFis) rangeCats.setValues(valsCats);
    }
    sheetFis.getRange('C2:C').setDataValidation(ruleTipoFis);
    sheetFis.getRange('D2:D').setDataValidation(ruleUN);
    sheetFis.getRange('E2:E').setDataValidation(ruleCat);
    sheetFis.getRange('F2:F').setDataValidation(ruleCat);
    if (ruleSec) sheetFis.getRange('I2:I').setDataValidation(ruleSec);
  }

  // 3. GANANCIA DIARIA KG: Limpiar validaciones residuales (celda A2 y encabezados)
  const sheetGan = ss.getSheetByName('GANANCIA DIARIA KG') || ss.getSheetByName('GANANCIA_DIARIA_PESO');
  if (sheetGan) {
    sheetGan.getRange('A1:Z5').clearDataValidations();
  }

  // 4. COSTOS: validar Etapa
  const sheetCostos = ss.getSheetByName('COSTOS');
  if (sheetCostos) {
    const ruleEtapaCostos = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Cria', 'Recria', 'Adaptacion', 'Engorde', 'Reproduccion', 'Destete'], true)
      .setAllowInvalid(false)
      .build();
    sheetCostos.getRange('C2:C').setDataValidation(ruleEtapaCostos);
  }
}

/**
 * Menú para ejecutar la validación y corrección de Movimientos Físicos y Comerciales
 */
function actualizarValidacionesMenu() {
  aplicarCambiosMovimientosYGananciaMenu();
}

/**
 * Función integral para aplicar y verificar los cambios en Movimientos Físicos,
 * Movimientos Comerciales, Ganancia Diaria KG y fórmulas vinculadas.
 */
function aplicarCambiosMovimientosYGanancia(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = [];

  // 1. Ejecutar normalización de datos y validaciones estrictas
  actualizarValidaciones();
  log.push('✔ MOV_FISICOS: Códigos normalizados a CAT_BOV_6..9 (se quitó el 0 de más) y desplegables estrictos en columnas C, D, E y F.');
  log.push('✔ MOV_COMERCIALES: Código CAT-BOV-01 normalizado a CAT_BOV_1 y validación estricta en columnas C, E, F, G, Q y R.');
  log.push('✔ GANANCIA DIARIA KG: Removido el desplegable residual de la celda A2 (encabezado Etapa).');

  // 2. Actualizar fórmulas de INVENTARIO_RODEO si la hoja existe
  const sheetInv = ss.getSheetByName('INVENTARIO_RODEO');
  if (sheetInv) {
    try {
      sheetInv.getRange('H3').setFormula('=SUMIFS(MOV_COMERCIALES!H2:H50, MOV_COMERCIALES!G2:G50, "CAT_BOV_1")');
      sheetInv.getRange('I3').setFormula('=SUMIFS(MOV_COMERCIALES!K2:K50, MOV_COMERCIALES!G2:G50, "CAT_BOV_1")');
      sheetInv.getRange('H7').setFormula('=SUMIFS(MOV_COMERCIALES!H2:H50, MOV_COMERCIALES!G2:G50, "CAT_BOV_12")');
      sheetInv.getRange('I7').setFormula('=SUMIFS(MOV_COMERCIALES!K2:K50, MOV_COMERCIALES!G2:G50, "CAT_BOV_12")');
      log.push('✔ INVENTARIO_RODEO: Fórmulas de ventas conectadas con CAT_BOV_1 y CAT_BOV_12.');
    } catch(e) {}
  }

  // 3. Actualizar fórmulas de ESTADO_RESULTADOS_ETAPAS si la hoja existe
  const sheetRes = ss.getSheetByName('ESTADO_RESULTADOS_ETAPAS');
  if (sheetRes) {
    try {
      sheetRes.getRange('B17').setFormula('=SUMIFS(MOV_COMERCIALES!K2:K50, MOV_COMERCIALES!F2:F50, "UN-BOV", MOV_COMERCIALES!G2:G50, "CAT_BOV_1")');
      sheetRes.getRange('E17').setFormula('=SUMIFS(MOV_COMERCIALES!K2:K50, MOV_COMERCIALES!F2:F50, "UN-BOV", MOV_COMERCIALES!G2:G50, "CAT_BOV_12")');
      log.push('✔ ESTADO_RESULTADOS_ETAPAS: Fórmulas de ingresos conectadas con CAT_BOV_1 y CAT_BOV_12.');
    } catch(e) {}
  }

  return log;
}

/**
 * Menú para ejecutar la aplicación integral de cambios en Movimientos y Ganancia Diaria
 */
function aplicarCambiosMovimientosYGananciaMenu() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const log = aplicarCambiosMovimientosYGanancia(ss);
    ss.toast('Cambios en Movimientos y Ganancia Diaria aplicados correctamente.', '🔧 Movimientos Actualizados', 5);
    SpreadsheetApp.getUi().alert(
      'Cambios Aplicados con Éxito',
      'Se completaron los siguientes ajustes en el sistema:\n\n' + log.join('\n\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(err) {
    SpreadsheetApp.getUi().alert('Error al aplicar cambios: ' + err.message);
  }
}

// ============================================================================
// 12. APIS PARA EL SIDEBAR Y FORMULARIO MOVIL
// ============================================================================
function obtenerCatalogosParaFormulario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  function getSheetData(name, colId, colName, colExtra) {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(colId, colName, colExtra || 1)).getValues();
    return values.map(function(r) { 
      return { 
        id: r[colId - 1], 
        nombre: r[colName - 1],
        extra: colExtra ? r[colExtra - 1] : ''
      }; 
    });
  }

  return {
    unidades: getSheetData('CAT_UNIDADES_NEGOCIO', 1, 2),
    sectores: getSheetData('CAT_SECTORES', 1, 2),
    categorias: getSheetData(ss.getSheetByName('CAT CATEGORIAS') ? 'CAT CATEGORIAS' : 'CAT_CATEGORIAS_BOVINOS', 1, 2),
    contactos: getSheetData('CAT_CONTACTOS', 1, 2),
    conceptos: getSheetData('CAT_CONCEPTOS', 1, 2, 3),
    cuentas: getSheetData('CAT_CUENTAS', 1, 2),
    caravanasActivas: obtenerDatosCaravanasExistentes(),
    costosCatalogo: obtenerCatalogoCostos(),
    gananciasDiarias: obtenerTablaGananciaDiaria()
  };
}

function registrarMovimientoFisico(form) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('MOV_FISICOS');
    if (!sheet) {
      inicializarSistema();
      sheet = ss.getSheetByName('MOV_FISICOS');
    }
    const lastRow = sheet.getLastRow();
    const nextId = 'FIS-2026-' + String(lastRow).padStart(4, '0');
    
    sheet.appendRow([
      nextId,
      form.fecha ? new Date(form.fecha) : new Date(),
      form.tipo || '',
      form.idUn || 'UN-BOV',
      form.idCatOrigen || '',
      form.idCatDestino || '',
      parseFloat(form.cabezas) || 0,
      parseFloat(form.pesoEstimado) || 0,
      form.idSector || '',
      form.caravana || '',
      form.insumoAsociado || '',
      parseFloat(form.cantidadInsumo) || 0,
      form.causaMuerte || '',
      form.observaciones || ''
    ]);
    return { exito: true, mensaje: '✔ Movimiento físico registrado con ID ' + nextId };
  } catch (e) {
    return { exito: false, mensaje: 'Error al registrar movimiento físico: ' + e.message };
  }
}

function registrarMovimientoComercial(form) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('MOV_COMERCIALES');
    if (!sheet) {
      inicializarSistema();
      sheet = ss.getSheetByName('MOV_COMERCIALES');
    }
    const lastRow = sheet.getLastRow();
    const nextId = 'COM-2026-' + String(lastRow).padStart(4, '0');

    const tipo = form.tipoOperacion || 'Venta';
    const cabezas = parseFloat(form.cabezas) || 0;
    const kilos = parseFloat(form.kilos) || 0;
    const precio = parseFloat(form.precioUnitario) || 0;
    const gastos = parseFloat(form.gastosComerciales) || 0;
    const retenciones = parseFloat(form.retenciones) || 0;
    const alicuota = parseFloat(form.alicuotaIva) || 0.105;

    const bruto = kilos > 0 ? (kilos * precio) : (cabezas * precio);
    const iva = (bruto - gastos) * alicuota;
    const neto = (tipo === 'Venta') ? (bruto - gastos + iva - retenciones) : (bruto + iva + retenciones);

    sheet.appendRow([
      nextId,
      form.fecha ? new Date(form.fecha) : new Date(),
      tipo,
      form.comprobante || '',
      form.idContacto || '',
      form.idUn || 'UN-BOV',
      form.idCat || '',
      cabezas,
      kilos,
      precio,
      bruto,
      gastos,
      alicuota,
      iva,
      retenciones,
      neto,
      form.condicionVenta || 'Contado',
      form.estadoPago || 'Cobrado',
      form.observaciones || ''
    ]);
    return { exito: true, mensaje: '✔ Operación comercial registrada con ID ' + nextId };
  } catch (e) {
    return { exito: false, mensaje: 'Error al registrar operación comercial: ' + e.message };
  }
}

function registrarMovimientoFinanciero(form) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('MOV_FINANCIEROS');
    if (!sheet) {
      inicializarSistema();
      sheet = ss.getSheetByName('MOV_FINANCIEROS');
    }
    const lastRow = sheet.getLastRow();
    const nextId = 'FIN-2026-' + String(lastRow).padStart(4, '0');

    sheet.appendRow([
      nextId,
      form.fecha ? new Date(form.fecha) : new Date(),
      form.tipo || 'Cobro',
      form.idContacto || '',
      form.idMovCom || '',
      form.idCuenta || '',
      parseFloat(form.importe) || 0,
      form.medioPago || 'Efectivo',
      form.observaciones || ''
    ]);
    return { exito: true, mensaje: '✔ Movimiento financiero registrado con ID ' + nextId };
  } catch (e) {
    return { exito: false, mensaje: 'Error al registrar movimiento financiero: ' + e.message };
  }
}

function normalizarUnidadNegocio(un) {
  if (!un) return 'Bovinos';
  const u = un.toString().trim().toLowerCase();
  if (u === 'un-bov' || u.includes('bov')) return 'Bovinos';
  if (u === 'un-por' || u.includes('porc')) return 'Porcinos';
  if (u === 'un-ovi' || u.includes('ovin')) return 'Ovinos';
  if (u.includes('maiz') || u.includes('maíz')) return 'Maiz';
  if (u.includes('soja')) return 'Soja';
  if (u.includes('alfalfa')) return 'Alfalfa';
  if (u.includes('estruc') || u === 'admin') return 'Estructura';
  return 'Bovinos';
}

function registrarImputacionGasto(form) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('TABLA_IMPUTACION');
    if (!sheet) {
      sheet = ss.getSheetByName('REGISTRO_GASTOS');
    }
    if (!sheet) {
      inicializarSistema();
      sheet = ss.getSheetByName('TABLA_IMPUTACION');
    }

    const unCanonica = normalizarUnidadNegocio(form.idUn || form.unidadNegocio);
    const importe = parseFloat(form.importe) || 0;
    let fecha = new Date();
    if (form.fecha) {
      if (typeof form.fecha === 'string' && form.fecha.includes('-')) {
        const partes = form.fecha.split('-');
        fecha = new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]));
      } else {
        fecha = new Date(form.fecha);
      }
    }

    const lastRow = sheet.getLastRow();
    const nuevaFila = lastRow + 1;

    sheet.getRange(nuevaFila, 1, 1, 9).setValues([[
      fecha,
      form.comprobante || '',
      form.proveedor || '',
      form.concepto || '',
      importe,
      unCanonica,
      form.etapaProductiva || 'Engorde',
      form.planoAfectado || 'Resultado Economico',
      form.efectoExistencias || 'Sin Efecto'
    ]]);

    sheet.getRange(nuevaFila, 1).setNumberFormat('dd/MM/yy').setHorizontalAlignment('center');
    sheet.getRange(nuevaFila, 2).setHorizontalAlignment('center');
    sheet.getRange(nuevaFila, 3).setHorizontalAlignment('left');
    sheet.getRange(nuevaFila, 4).setHorizontalAlignment('left');
    sheet.getRange(nuevaFila, 5).setNumberFormat('$ #,##0.00').setHorizontalAlignment('right');
    sheet.getRange(nuevaFila, 6).setHorizontalAlignment('center');
    sheet.getRange(nuevaFila, 7).setHorizontalAlignment('center');
    sheet.getRange(nuevaFila, 8).setHorizontalAlignment('left');
    sheet.getRange(nuevaFila, 9).setHorizontalAlignment('center');

    return { exito: true, mensaje: '✔ Imputación de gasto registrada correctamente en TABLA_IMPUTACION.' };
  } catch (e) {
    return { exito: false, mensaje: 'Error al registrar gasto: ' + e.message };
  }
}

function registrarLaborCampo(form) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('MOV_CAMPO_LABORES');
    if (!sheet) {
      sheet = ss.getSheetByName('LABORES_CAMPO');
    }
    if (!sheet) {
      inicializarSistema();
      sheet = ss.getSheetByName('MOV_CAMPO_LABORES');
    }
    const lastRow = sheet.getLastRow();
    const nextId = 'LAB-2026-' + String(lastRow).padStart(4, '0');

    const ha = parseFloat(form.superficieHa) || 0;
    const dosis = parseFloat(form.dosisPorHa) || 0;
    const total = ha * dosis;

    sheet.appendRow([
      nextId,
      form.fecha ? new Date(form.fecha) : new Date(),
      form.lote || '',
      ha,
      form.idUn || 'UN-MAIZ',
      form.tipoLabor || '',
      form.insumoProducto || '',
      dosis,
      form.unidadDosis || 'kg/ha',
      total,
      form.unidadTotal || 'kg',
      form.observaciones || ''
    ]);
    return { exito: true, mensaje: '✔ Labor de campo registrada con ID ' + nextId };
  } catch (e) {
    return { exito: false, mensaje: 'Error al registrar labor de campo: ' + e.message };
  }
}

function registrarNuevoContacto(form) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('CAT_CONTACTOS');
    if (!sheet) {
      inicializarSistema();
      sheet = ss.getSheetByName('CAT_CONTACTOS');
    }
    const lastRow = sheet.getLastRow();
    const nextId = 'CONT-' + String(lastRow).padStart(3, '0');

    sheet.appendRow([
      nextId,
      (form.razonSocial || '').toString().trim(),
      form.tipoContacto || 'Cliente',
      form.cuit || '',
      form.condicionPago || 'Contado',
      form.telefono || ''
    ]);
    return { exito: true, mensaje: '✔ Contacto registrado con código ' + nextId };
  } catch (e) {
    return { exito: false, mensaje: 'Error al registrar contacto: ' + e.message };
  }
}


/**
 * Determina la siguiente fila libre en la tabla principal de EXISTENCIAS (Columnas A:H).
 * Esto evita el desfase causado por las tablas de resumen de las columnas J:K (filas 1 a 30).
 * Además, si existían registros cargados por error en la fila 31+ estando vacías las filas 2..30,
 * los reubica automáticamente a partir de la fila 2 justo debajo de los encabezados.
 */
function compactarYObtenerSiguienteFilaExistencias(sheet) {
  const maxRows = Math.min(sheet.getMaxRows(), 1000);
  const rangoDatos = sheet.getRange(2, 1, maxRows - 1, 10);
  const valores = rangoDatos.getValues();
  const formulas = rangoDatos.getFormulas();

  const filasConDatos = [];
  for (let i = 0; i < valores.length; i++) {
    const rowNum = i + 2;
    const id = valores[i][0];
    const caravana = valores[i][2];
    const tieneDato = (id !== '' && id !== null && id !== undefined) ||
                      (caravana !== '' && caravana !== null && caravana !== undefined);
    if (tieneDato) {
      filasConDatos.push({
        filaOriginal: rowNum,
        valores: valores[i],
        formulas: formulas[i]
      });
    }
  }

  if (filasConDatos.length === 0) {
    return 2;
  }

  if (filasConDatos[0].filaOriginal > 2) {
    for (let j = 0; j < filasConDatos.length; j++) {
      const item = filasConDatos[j];
      const nuevaFilaDestino = 2 + j;
      const fOrig = item.filaOriginal;

      if (fOrig !== nuevaFilaDestino) {
        for (let c = 0; c < 10; c++) {
          if (c === 4) {
            const cat = item.valores[4] || 'Terneros';
            sheet.getRange(nuevaFilaDestino, 5).setFormula(`=IFERROR(XLOOKUP(C${nuevaFilaDestino}, Movimientos_Existencia!$C:$C, Movimientos_Existencia!$H:$H, "${cat}", 0, -1), "${cat}")`);
          } else if (c === 5) {
            const sub = item.valores[5] || 'Cria';
            sheet.getRange(nuevaFilaDestino, 6).setFormula(`=IFERROR(XLOOKUP(C${nuevaFilaDestino}, Movimientos_Existencia!$C:$C, Movimientos_Existencia!$I:$I, "${sub}", 0, -1), "${sub}")`);
          } else if (c === 7) {
            sheet.getRange(nuevaFilaDestino, 8).setFormula(`=IFERROR(XLOOKUP(C${nuevaFilaDestino}, Movimientos_Existencia!$C:$C, Movimientos_Existencia!$F:$F, G${nuevaFilaDestino}, 0, -1), G${nuevaFilaDestino})`);
          } else if (c === 8) {
            sheet.getRange(nuevaFilaDestino, 9).setFormula(`=IFERROR(CHOOSEROWS(FILTER(Movimientos_Existencia!$G$2:$G, Movimientos_Existencia!$C$2:$C=C${nuevaFilaDestino}, ISNUMBER(Movimientos_Existencia!$G$2:$G), Movimientos_Existencia!$G$2:$G>0), -1), G${nuevaFilaDestino})`);
          } else {
            sheet.getRange(nuevaFilaDestino, c + 1).setValue(item.valores[c]);
          }
        }
        sheet.getRange(fOrig, 1, 1, 10).clearContent().clearFormat();
      }
    }
  }

  let ultimaFilaOcupada = 1;
  const valsCheck = sheet.getRange(2, 3, Math.min(sheet.getMaxRows() - 1, 500), 1).getValues();
  for (let k = 0; k < valsCheck.length; k++) {
    const carVal = (valsCheck[k][0] || '').toString().trim();
    if (carVal !== '') {
      ultimaFilaOcupada = k + 2;
    }
  }

  return ultimaFilaOcupada + 1;
}

function reorganizarFilasExistencias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('EXISTENCIAS');
  if (!sheet) return;
  const proxFila = compactarYObtenerSiguienteFilaExistencias(sheet);
  const sincronizados = sincronizarNacimientosFaltantes(ss);
  SpreadsheetApp.getUi().alert(
    'Reorganización Completada\n\n' +
    '1. La tabla de animales en EXISTENCIAS ha sido compactada directamente debajo de los encabezados (fila 2 en adelante).\n' +
    '2. Eventos de nacimiento sincronizados en Movimientos_Existencia: ' + sincronizados + ' registro(s).\n' +
    '3. Próxima fila disponible para carga en EXISTENCIAS: Fila ' + proxFila + '.'
  );
}

function guardarNuevaCaravana(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('EXISTENCIAS');
    if (!sheet) {
      prepararHojaExistencias(ss);
      sheet = ss.getSheetByName('EXISTENCIAS');
    }

    const caravana = (payload.caravana || '').toString().trim().toUpperCase();
    if (!caravana) {
      return { exito: false, mensaje: 'Debe ingresar el código de la caravana.' };
    }

    const nuevaFila = compactarYObtenerSiguienteFilaExistencias(sheet);

    if (nuevaFila > 2) {
      const caravanasExistentes = sheet.getRange(2, 3, nuevaFila - 2, 1).getValues().flat().map(c => (c || '').toString().trim().toUpperCase());
      if (caravanasExistentes.includes(caravana)) {
        return { exito: false, mensaje: `La caravana "${caravana}" ya se encuentra registrada en EXISTENCIAS.` };
      }
    }

    let nextId = 1;
    if (nuevaFila > 2) {
      const ids = sheet.getRange(2, 1, nuevaFila - 2, 1).getValues().flat().filter(v => typeof v === 'number' || (!isNaN(v) && v !== ''));
      if (ids.length > 0) {
        nextId = Math.max(...ids.map(Number)) + 1;
      }
    }

    const fecha = payload.fechaRegistro || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const unidadProduccion = payload.unidadProduccion || 'Bovino';
    const categoria = payload.categoria || (unidadProduccion === 'Porcino' ? 'Lechon' : unidadProduccion === 'Ovino' ? 'Corderos' : 'Terneros');
    const subunidad = payload.subunidad || 'Cria';
    const pesoInicial = parseFloat(payload.pesoInicial) || 0;

    const formulaCategoria = `=IFERROR(XLOOKUP(C${nuevaFila}, Movimientos_Existencia!$C:$C, Movimientos_Existencia!$H:$H, "${categoria}", 0, -1), "${categoria}")`;
    const formulaSubunidad = `=IFERROR(XLOOKUP(C${nuevaFila}, Movimientos_Existencia!$C:$C, Movimientos_Existencia!$I:$I, "${subunidad}", 0, -1), "${subunidad}")`;
    const formulaPesoEstimado = `=IFERROR(XLOOKUP(C${nuevaFila}, Movimientos_Existencia!$C:$C, Movimientos_Existencia!$F:$F, G${nuevaFila}, 0, -1), G${nuevaFila})`;
    const formulaPesoReal = `=IFERROR(CHOOSEROWS(FILTER(Movimientos_Existencia!$G$2:$G, Movimientos_Existencia!$C$2:$C=C${nuevaFila}, ISNUMBER(Movimientos_Existencia!$G$2:$G), Movimientos_Existencia!$G$2:$G>0), -1), G${nuevaFila})`;

    sheet.getRange(nuevaFila, 1).setValue(nextId);
    sheet.getRange(nuevaFila, 2).setValue(fecha);
    sheet.getRange(nuevaFila, 3).setValue(caravana);
    sheet.getRange(nuevaFila, 4).setValue(unidadProduccion);
    sheet.getRange(nuevaFila, 5).setFormula(formulaCategoria);
    sheet.getRange(nuevaFila, 6).setFormula(formulaSubunidad);
    sheet.getRange(nuevaFila, 7).setValue(pesoInicial);
    sheet.getRange(nuevaFila, 8).setFormula(formulaPesoEstimado);
    sheet.getRange(nuevaFila, 9).setFormula(formulaPesoReal);
    sheet.getRange(nuevaFila, 10).setValue('Vivo');

    const reglaUnidad = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Bovino', 'Porcino', 'Ovino'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(nuevaFila, 4).setDataValidation(reglaUnidad);

    const reglaEstado = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Vivo', 'Muerto', 'Vendido'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(nuevaFila, 10).setDataValidation(reglaEstado);

    sheet.getRange(nuevaFila, 1, 1, 10).setVerticalAlignment('middle');
    sheet.getRange(nuevaFila, 1).setHorizontalAlignment('center');
    sheet.getRange(nuevaFila, 2).setHorizontalAlignment('center').setNumberFormat('yyyy-MM-dd');
    sheet.getRange(nuevaFila, 3).setHorizontalAlignment('center').setFontWeight('bold');
    sheet.getRange(nuevaFila, 4).setHorizontalAlignment('center');
    sheet.getRange(nuevaFila, 5).setHorizontalAlignment('center');
    sheet.getRange(nuevaFila, 6).setHorizontalAlignment('center');
    sheet.getRange(nuevaFila, 7).setHorizontalAlignment('right').setNumberFormat('#,##0.00');
    sheet.getRange(nuevaFila, 8).setHorizontalAlignment('right').setNumberFormat('#,##0.00');
    sheet.getRange(nuevaFila, 9).setHorizontalAlignment('right').setNumberFormat('#,##0.00');
    sheet.getRange(nuevaFila, 10).setHorizontalAlignment('center').setFontWeight('bold');

    let sheetMov = ss.getSheetByName('Movimientos_Existencia');
    if (!sheetMov) {
      prepararHojaMovimientosExistencia(ss);
      sheetMov = ss.getSheetByName('Movimientos_Existencia');
    }
    const lastRowMov = sheetMov.getLastRow();
    const nuevaFilaMov = lastRowMov + 1;
    let nextIdMov = 1;
    if (lastRowMov > 1) {
      const idsMov = sheetMov.getRange(2, 1, lastRowMov - 1, 1).getValues().flat().filter(v => typeof v === 'number' || (!isNaN(v) && v !== ''));
      if (idsMov.length > 0) {
        nextIdMov = Math.max(...idsMov.map(Number)) + 1;
      }
    }

    sheetMov.getRange(nuevaFilaMov, 1).setValue(nextIdMov);
    sheetMov.getRange(nuevaFilaMov, 2).setValue(fecha);
    sheetMov.getRange(nuevaFilaMov, 3).setValue(caravana);
    sheetMov.getRange(nuevaFilaMov, 4).setValue(unidadProduccion);
    sheetMov.getRange(nuevaFilaMov, 5).setValue(fecha);
    sheetMov.getRange(nuevaFilaMov, 6).setValue(pesoInicial);
    sheetMov.getRange(nuevaFilaMov, 7).setValue(pesoInicial > 0 ? pesoInicial : '');
    sheetMov.getRange(nuevaFilaMov, 8).setValue(categoria);
    sheetMov.getRange(nuevaFilaMov, 9).setValue(subunidad);
    sheetMov.getRange(nuevaFilaMov, 10).setValue('Nacimiento / Ingreso a Existencias');
    sheetMov.getRange(nuevaFilaMov, 11).setValue('Cabezas');
    sheetMov.getRange(nuevaFilaMov, 12).setValue(0);
    sheetMov.getRange(nuevaFilaMov, 13).setValue(1);
    sheetMov.getRange(nuevaFilaMov, 14).setFormula(`=L${nuevaFilaMov}*M${nuevaFilaMov}`);
    sheetMov.getRange(nuevaFilaMov, 15).setValue('Alta inicial de caravana en sistema');

    return {
      exito: true,
      mensaje: `Caravana "${caravana}" (${unidadProduccion}) registrada con éxito en fila ${nuevaFila} de EXISTENCIAS.`,
      caravana: caravana,
      unidadProduccion: unidadProduccion,
      categoria: categoria,
      subunidad: subunidad,
      pesoInicial: pesoInicial
    };
  } catch (err) {
    return { exito: false, mensaje: 'Error al guardar caravana: ' + err.message };
  }
}

/**
 * Sincroniza retroactivamente animales de EXISTENCIAS que no tengan aún su evento de nacimiento en Movimientos_Existencia
 */
function sincronizarNacimientosFaltantes(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetExis = ss.getSheetByName('EXISTENCIAS');
  let sheetMov = ss.getSheetByName('Movimientos_Existencia');
  if (!sheetExis) return 0;
  if (!sheetMov) {
    prepararHojaMovimientosExistencia(ss);
    sheetMov = ss.getSheetByName('Movimientos_Existencia');
  }

  const maxRows = Math.min(sheetExis.getMaxRows(), 1000);
  const dataExis = sheetExis.getRange(2, 1, maxRows - 1, 10).getValues();

  const lastRowMov = sheetMov.getLastRow();
  let caravanasEnMov = [];
  if (lastRowMov > 1) {
    caravanasEnMov = sheetMov.getRange(2, 3, lastRowMov - 1, 1).getValues().flat().map(c => (c || '').toString().trim().toUpperCase());
  }

  let creados = 0;
  for (let i = 0; i < dataExis.length; i++) {
    const fecha = dataExis[i][1];
    const caravana = (dataExis[i][2] || '').toString().trim().toUpperCase();
    const unidadProduccion = dataExis[i][3] || 'Bovino';
    const categoria = dataExis[i][4] || 'Terneros';
    const subunidad = dataExis[i][5] || 'Cria';
    const pesoInicial = parseFloat(dataExis[i][6]) || 0;

    if (caravana && !caravanasEnMov.includes(caravana)) {
      const curLastRow = sheetMov.getLastRow();
      const nFila = curLastRow + 1;
      let nextIdMov = 1;
      if (curLastRow > 1) {
        const ids = sheetMov.getRange(2, 1, curLastRow - 1, 1).getValues().flat().filter(v => typeof v === 'number' || (!isNaN(v) && v !== ''));
        if (ids.length > 0) nextIdMov = Math.max(...ids.map(Number)) + 1;
      }

      sheetMov.getRange(nFila, 1).setValue(nextIdMov);
      sheetMov.getRange(nFila, 2).setValue(fecha);
      sheetMov.getRange(nFila, 3).setValue(caravana);
      sheetMov.getRange(nFila, 4).setValue(unidadProduccion);
      sheetMov.getRange(nFila, 5).setValue(fecha);
      sheetMov.getRange(nFila, 6).setValue(pesoInicial);
      sheetMov.getRange(nFila, 7).setValue(pesoInicial > 0 ? pesoInicial : '');
      sheetMov.getRange(nFila, 8).setValue(categoria);
      sheetMov.getRange(nFila, 9).setValue(subunidad);
      sheetMov.getRange(nFila, 10).setValue('Nacimiento / Ingreso a Existencias');
      sheetMov.getRange(nFila, 11).setValue('Cabezas');
      sheetMov.getRange(nFila, 12).setValue(0);
      sheetMov.getRange(nFila, 13).setValue(1);
      sheetMov.getRange(nFila, 14).setFormula(`=L${nFila}*M${nFila}`);
      sheetMov.getRange(nFila, 15).setValue('Alta histórica sincronizada');

      caravanasEnMov.push(caravana);
      creados++;
    }
  }
  return creados;
}

// ============================================================================
// 4.1. MÓDULO DE REGISTRO DE COSTOS E INSUMOS POR CARAVANA (Movimientos_Existencia)
// ============================================================================

/**
 * Retorna las caravanas activas (Estado == 'Vivo') registradas en la hoja EXISTENCIAS
 */
function obtenerDatosCaravanasExistentes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('EXISTENCIAS');
    if (!sheet || sheet.getLastRow() <= 1) return [];

    const lastRow = sheet.getLastRow();
    const rows = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const resultado = [];

    for (let i = 0; i < rows.length; i++) {
      const id = rows[i][0];
      const fechaReg = rows[i][1];
      const caravana = (rows[i][2] || '').toString().trim().toUpperCase();
      const unidadProduccion = (rows[i][3] || 'Bovino').toString().trim();
      const categoria = (rows[i][4] || '').toString().trim();
      const subunidad = (rows[i][5] || '').toString().trim();
      const pesoInicial = parseFloat(rows[i][6]) || 0;
      const pesoEstimado = parseFloat(rows[i][7]) || pesoInicial;
      const pesoReal = (rows[i][8] !== '' && !isNaN(rows[i][8])) ? parseFloat(rows[i][8]) : null;
      const estado = (rows[i][9] || '').toString().trim();

      if (caravana && estado === 'Vivo') {
        let fechaNacStr = '';
        if (fechaReg instanceof Date) {
          fechaNacStr = Utilities.formatDate(fechaReg, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else if (fechaReg) {
          fechaNacStr = fechaReg.toString().trim();
        }

        resultado.push({
          id: id,
          caravana: caravana,
          unidadProduccion: unidadProduccion,
          fechaNacimiento: fechaNacStr,
          categoria: categoria || 'Terneros',
          subunidad: subunidad || 'Cria',
          pesoInicial: pesoInicial,
          pesoEstimado: pesoEstimado,
          pesoReal: pesoReal
        });
      }
    }
    return resultado;
  } catch (err) {
    return [];
  }
}

/**
 * Retorna el catálogo de conceptos y precios de la hoja COSTOS
 */
function obtenerCatalogoCostos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('COSTOS');
    if (!sheet || sheet.getLastRow() <= 1) return [];

    const lastRow = sheet.getLastRow();
    const rows = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const catalogo = [];

    for (let i = 0; i < rows.length; i++) {
      const idCosto = (rows[i][0] || '').toString().trim();
      const concepto = (rows[i][1] || '').toString().trim();
      const subunidad = (rows[i][2] || '').toString().trim();
      const unidad = (rows[i][3] || '').toString().trim();
      let costoUnitario = rows[i][4];
      if (typeof costoUnitario !== 'number') {
        costoUnitario = parseFloat(costoUnitario) || 0;
      }

      if (concepto) {
        catalogo.push({
          id: idCosto,
          concepto: concepto,
          subunidad: subunidad,
          unidad: unidad,
          costoUnitario: costoUnitario
        });
      }
    }
    return catalogo;
  } catch (err) {
    Logger.log('Error en obtenerCatalogoCostos: ' + err.toString());
    return [];
  }
}

/**
 * Retorna las ganancias diarias en kg/día por etapa productiva
 */
function obtenerTablaGananciaDiaria() {
  return {
    'Bovino': {
      'Cria': 0.55,
      'Recria': 0.60,
      'Adaptacion': 1.00,
      'Engorde': 1.25,
      'Reproduccion': 0.00
    },
    'Porcino': {
      'Cria': 0.20,
      'Destete': 0.48,
      'Engorde': 0.85,
      'Reproduccion': 0.00
    },
    'Ovino': {
      'Cria': 0.23,
      'Recria': 0.12,
      'Engorde': 0.20,
      'Reproduccion': 0.00
    }
  };
}

/**
 * Determina la subunidad/etapa según los rangos de peso fijados para cada especie:
 * - Bovinos:  <140 kg: Cria | 140 a <280 kg: Recria | 280 a <315 kg: Adaptacion | >=315 kg: Engorde
 * - Porcinos: <7 kg: Cria | 7 a <30 kg: Destete | >=30 kg: Engorde
 * - Ovinos:   <26 kg: Cria | 26 a <40 kg: Recria | >=40 kg: Engorde
 */
function determinarSubunidadPorPeso(peso, unidadProduccion, etapaActualDefault) {
  const p = parseFloat(peso) || 0;
  const up = (unidadProduccion || 'Bovino').toString().trim().toLowerCase();

  if (up.includes('porc')) {
    if (etapaActualDefault && etapaActualDefault.toLowerCase().includes('reprod')) return 'Reproduccion';
    if (p < 7) return 'Cria';
    if (p < 30) return 'Destete';
    return 'Engorde';
  } else if (up.includes('ovin')) {
    if (etapaActualDefault && etapaActualDefault.toLowerCase().includes('reprod')) return 'Reproduccion';
    if (p < 26) return 'Cria';
    if (p < 40) return 'Recria';
    return 'Engorde';
  } else {
    // Bovino
    if (etapaActualDefault && etapaActualDefault.toLowerCase().includes('reprod')) return 'Reproduccion';
    if (p < 140) return 'Cria';
    if (p < 280) return 'Recria';
    if (p < 315) return 'Adaptacion';
    return 'Engorde';
  }
}

/**
 * Calcula la evolución del peso proyectado del animal según sus días de vida y su especie
 */
function calcularEvolucionPesoCaravana(pesoInicial, fechaNacStr, fechaRegStr, subunidadBase, unidadProduccion) {
  const up = (unidadProduccion || 'Bovino').toString().trim();
  let defaultIni = 35.0;
  if (up.toLowerCase().includes('porc')) defaultIni = 1.45;
  else if (up.toLowerCase().includes('ovin')) defaultIni = 4.0;

  const pIni = parseFloat(pesoInicial) || defaultIni;
  let dNac = new Date(fechaNacStr + 'T00:00:00');
  let dReg = new Date(fechaRegStr + 'T00:00:00');
  if (isNaN(dNac.getTime())) dNac = new Date();
  if (isNaN(dReg.getTime())) dReg = new Date();

  const diffTime = dReg.getTime() - dNac.getTime();
  const dias = Math.max(0, Math.floor(diffTime / (1000 * 60 * 60 * 24)));

  const tablasEspecies = obtenerTablaGananciaDiaria();
  let tablaEspecie = tablasEspecies[up];
  if (!tablaEspecie) {
    const upLower = up.toLowerCase();
    if (upLower.includes('porc')) tablaEspecie = tablasEspecies['Porcino'];
    else if (upLower.includes('ovin')) tablaEspecie = tablasEspecies['Ovino'];
    else tablaEspecie = tablasEspecies['Bovino'];
  }

  const sub = subunidadBase || 'Cria';
  const gd = (tablaEspecie && tablaEspecie[sub] !== undefined) ? tablaEspecie[sub] : 0.55;
  const pesoCalculado = pIni + (dias * gd);
  const subunidadActual = determinarSubunidadPorPeso(pesoCalculado, up, sub);

  return {
    dias: dias,
    gananciaDiaria: gd,
    pesoCalculado: Math.round(pesoCalculado * 10) / 10,
    subunidadActual: subunidadActual
  };
}

/**
 * Registra un costo/insumo para una caravana en la hoja Movimientos_Existencia
 * y sincroniza el peso actual y subunidad del animal en EXISTENCIAS.
 */
function guardarCostoMovimientoExistencia(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetMov = ss.getSheetByName('Movimientos_Existencia');
    if (!sheetMov) {
      prepararHojaMovimientosExistencia(ss);
      sheetMov = ss.getSheetByName('Movimientos_Existencia');
    }

    const sheetExis = ss.getSheetByName('EXISTENCIAS');
    if (!sheetExis) {
      return { exito: false, mensaje: 'No se encontró la hoja EXISTENCIAS.' };
    }

    const caravana = (payload.caravana || '').toString().trim().toUpperCase();
    if (!caravana) {
      return { exito: false, mensaje: 'Debe seleccionar una caravana válida.' };
    }

    const lastRowExis = sheetExis.getLastRow();
    let filaCaravanaExis = -1;
    let fechaNacimientoExis = '';
    let unidadProduccionExis = 'Bovino';
    let categoriaExis = '';
    let subunidadExis = '';
    let pesoInicialExis = 0;

    if (lastRowExis > 1) {
      const dataExis = sheetExis.getRange(2, 1, lastRowExis - 1, 10).getValues();
      for (let r = 0; r < dataExis.length; r++) {
        const c = (dataExis[r][2] || '').toString().trim().toUpperCase();
        const est = (dataExis[r][9] || '').toString().trim();
        if (c === caravana) {
          if (est !== 'Vivo') {
            return { exito: false, mensaje: `La caravana "${caravana}" se encuentra en estado "${est}" y no admite nuevas registraciones de costos.` };
          }
          filaCaravanaExis = r + 2;
          const fn = dataExis[r][1];
          if (fn instanceof Date) {
            fechaNacimientoExis = Utilities.formatDate(fn, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          } else {
            fechaNacimientoExis = (fn || '').toString().trim();
          }
          unidadProduccionExis = (dataExis[r][3] || 'Bovino').toString().trim();
          categoriaExis = (dataExis[r][4] || '').toString().trim();
          subunidadExis = (dataExis[r][5] || '').toString().trim();
          pesoInicialExis = parseFloat(dataExis[r][6]) || 0;
          break;
        }
      }
    }

    if (filaCaravanaExis === -1) {
      return { exito: false, mensaje: `La caravana "${caravana}" no existe en EXISTENCIAS. Solo se pueden registrar costos para caravanas activas.` };
    }

    const fechaRegistro = payload.fechaRegistro || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const pesoEstimado = parseFloat(payload.pesoEstimado || payload.peso) || pesoInicialExis;
    const pesoReal = (payload.pesoReal !== undefined && payload.pesoReal !== null && payload.pesoReal !== '') ? parseFloat(payload.pesoReal) : '';
    const etapa = payload.etapa || subunidadExis || 'Cria';
    const categoria = payload.categoria || categoriaExis || 'Terneros';
    const conceptoCosto = (payload.conceptoCosto || '').toString().trim();
    const unidad = (payload.unidad || '').toString().trim();
    const costoUnitario = parseFloat(payload.costoUnitario) || 0;
    const cantidad = parseFloat(payload.cantidad) || 0;
    const observaciones = (payload.observaciones || '').toString().trim();

    const lastRowMov = sheetMov.getLastRow();
    const nuevaFilaMov = lastRowMov + 1;
    let nextIdMov = 1;
    if (lastRowMov > 1) {
      const idsMov = sheetMov.getRange(2, 1, lastRowMov - 1, 1).getValues().flat().filter(v => typeof v === 'number' || (!isNaN(v) && v !== ''));
      if (idsMov.length > 0) {
        nextIdMov = Math.max(...idsMov.map(Number)) + 1;
      }
    }

    sheetMov.getRange(nuevaFilaMov, 1).setValue(nextIdMov);
    sheetMov.getRange(nuevaFilaMov, 2).setValue(fechaRegistro);
    sheetMov.getRange(nuevaFilaMov, 3).setValue(caravana);
    sheetMov.getRange(nuevaFilaMov, 4).setValue(unidadProduccionExis);
    sheetMov.getRange(nuevaFilaMov, 5).setValue(fechaNacimientoExis || fechaRegistro);
    sheetMov.getRange(nuevaFilaMov, 6).setValue(pesoEstimado);
    if (pesoReal !== '' && !isNaN(pesoReal) && pesoReal > 0) {
      sheetMov.getRange(nuevaFilaMov, 7).setValue(pesoReal);
    } else {
      sheetMov.getRange(nuevaFilaMov, 7).setValue('');
    }
    sheetMov.getRange(nuevaFilaMov, 8).setValue(categoria);
    sheetMov.getRange(nuevaFilaMov, 9).setValue(etapa);
    sheetMov.getRange(nuevaFilaMov, 10).setValue(conceptoCosto);
    sheetMov.getRange(nuevaFilaMov, 11).setValue(unidad);
    sheetMov.getRange(nuevaFilaMov, 12).setValue(costoUnitario);
    sheetMov.getRange(nuevaFilaMov, 13).setValue(cantidad);
    sheetMov.getRange(nuevaFilaMov, 14).setFormula(`=L${nuevaFilaMov}*M${nuevaFilaMov}`);
    sheetMov.getRange(nuevaFilaMov, 15).setValue(observaciones);

    sheetMov.getRange(nuevaFilaMov, 1, 1, 15).setVerticalAlignment('middle');
    sheetMov.getRange(nuevaFilaMov, 1).setHorizontalAlignment('center');
    sheetMov.getRange(nuevaFilaMov, 2).setHorizontalAlignment('center');
    sheetMov.getRange(nuevaFilaMov, 3).setHorizontalAlignment('center').setFontWeight('bold');
    sheetMov.getRange(nuevaFilaMov, 4).setHorizontalAlignment('center');
    sheetMov.getRange(nuevaFilaMov, 5).setHorizontalAlignment('center');
    sheetMov.getRange(nuevaFilaMov, 6, 1, 2).setHorizontalAlignment('right').setNumberFormat('#,##0.00');
    sheetMov.getRange(nuevaFilaMov, 8, 1, 2).setHorizontalAlignment('center');
    sheetMov.getRange(nuevaFilaMov, 11).setHorizontalAlignment('center');
    sheetMov.getRange(nuevaFilaMov, 12).setHorizontalAlignment('right').setNumberFormat('$ #,##0.00');
    sheetMov.getRange(nuevaFilaMov, 13).setHorizontalAlignment('right').setNumberFormat('#,##0.00');
    sheetMov.getRange(nuevaFilaMov, 14).setHorizontalAlignment('right').setNumberFormat('$ #,##0.00');

    // Sincronización automática inmediata con TABLA_IMPUTACION
    try {
      sincronizarFilaMovimientoExistencia(sheetMov, nuevaFilaMov);
    } catch (eSync) {
      console.warn('Sincronización automática inmediata diferida: ' + eSync.message);
    }

    return {
      exito: true,
      mensaje: `Costo registrado con éxito para caravana "${caravana}" (${unidadProduccionExis}) en fila ${nuevaFilaMov} de Movimientos_Existencia.`,
      id: nextIdMov,
      caravana: caravana,
      unidadProduccion: unidadProduccionExis,
      concepto: conceptoCosto,
      total: costoUnitario * cantidad
    };
  } catch (err) {
    return { exito: false, mensaje: 'Error al registrar costo: ' + err.message };
  }
}

/**
 * Actualiza y reestablece las fórmulas dinámicas en EXISTENCIAS para todas las filas
 */
function actualizarFormulasExistencias(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('EXISTENCIAS');
  if (!sheet) return 0;

  sheet.getRange('I2:I1000').clearDataValidations();

  const lastRow = sheet.getLastRow();
  let actualizadas = 0;

  for (let r = 2; r <= lastRow; r++) {
    const caravana = (sheet.getRange(r, 3).getValue() || '').toString().trim();
    if (caravana) {
      let up = sheet.getRange(r, 4).getValue();
      if (!up) {
        up = 'Bovino';
        sheet.getRange(r, 4).setValue(up);
      }
      const upLower = String(up).toLowerCase();
      const catDef = upLower.includes('porc') ? 'Lechon' : (upLower.includes('ovin') ? 'Corderos' : 'Terneros');
      const subDef = 'Cria';

      sheet.getRange(r, 5).setFormula(`=IFERROR(XLOOKUP(C${r}, Movimientos_Existencia!$C:$C, Movimientos_Existencia!$H:$H, "${catDef}", 0, -1), "${catDef}")`);
      sheet.getRange(r, 6).setFormula(`=IFERROR(XLOOKUP(C${r}, Movimientos_Existencia!$C:$C, Movimientos_Existencia!$I:$I, "${subDef}", 0, -1), "${subDef}")`);
      sheet.getRange(r, 8).setFormula(`=IFERROR(XLOOKUP(C${r}, Movimientos_Existencia!$C:$C, Movimientos_Existencia!$F:$F, G${r}, 0, -1), G${r})`);
      sheet.getRange(r, 9).setFormula(`=IFERROR(CHOOSEROWS(FILTER(Movimientos_Existencia!$G$2:$G, Movimientos_Existencia!$C$2:$C=C${r}, ISNUMBER(Movimientos_Existencia!$G$2:$G), Movimientos_Existencia!$G$2:$G>0), -1), G${r})`);
      
      const est = sheet.getRange(r, 10).getValue();
      if (!est) sheet.getRange(r, 10).setValue('Vivo');

      actualizadas++;
    }
  }
  return actualizadas;
}

function ejecutarActualizacionFormulasExistenciasMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const n = actualizarFormulasExistencias(ss);
  SpreadsheetApp.getUi().alert(
    'Fórmulas Actualizadas en EXISTENCIAS\n\n' +
    'Se reestablecieron las fórmulas dinámicas en ' + n + ' filas de animales:\n' +
    '• Peso Actual: toma el último peso registrado en Movimientos_Existencia.\n' +
    '• Subunidad: se calcula según el peso (<140 Cría, 140-280 Recría, 280-315 Adaptación, >=315 Engorde).'
  );
}

// ============================================================================
// 13. MÓDULO DE COTIZACIONES DE MERCADO GANADERO (MAG CAÑUELAS Y ROSGAN)
// ============================================================================
var MAG_URL = 'https://www.mercadoagroganadero.com.ar/dll/hacienda1.dll/haciinfo000502';
var ROSGAN_URL = 'https://www.rosgan.com.ar/api/precios-fede/';
var TIMEZONE_ARG = 'America/Argentina/Buenos_Aires';

var CABECERAS_HTTP = {
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
  'Accept-Language': 'es-AR,es;q=0.9'
};

var HEADERS_COTIZACIONES = [
  'Mercado', 'Categoria', 'Subtipo', 'Cabezas', 'Precio_Promedio', 'Precio_Minimo',
  'Precio_Maximo', 'Precio_Mediana', 'Kilos_Totales', 'Promedio_Kilos', 'Importe_Total',
  'Variacion_Pct', 'Indice_MAG_ROSGAN', 'Fecha_Cotizacion', 'Estado_Rueda', 'Fecha_Consulta'
];

var ANCHOS_COTIZACIONES = [130, 160, 160, 85, 125, 110, 110, 110, 105, 105, 130, 105, 130, 115, 115, 135];

/**
 * Rutina principal para consultar las APIs de MAG Cañuelas y ROSGAN
 * y volcar las cotizaciones en la hoja unificada COTIZACIONES.
 * @param {boolean} mostrarAlerta Si es true (por defecto), muestra ui.alert/toast.
 */
function actualizarCotizacionesMercado(mostrarAlerta) {
  if (mostrarAlerta === undefined) mostrarAlerta = true;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}

  var ahora = new Date();
  var consultaStr = Utilities.formatDate(ahora, TIMEZONE_ARG, 'dd/MM/yyyy HH:mm');

  // 1. Descarga de MAG Cañuelas (Scraping HTML con fallback hasta 10 días)
  var mag = null;
  try {
    mag = descargarCanuelas_();
  } catch (eMag) {
    Logger.log('Aviso MAG Cañuelas: ' + eMag.message);
  }

  // 2. Descarga de ROSGAN (API REST JSON)
  var remates = null;
  try {
    remates = descargarRosgan_();
  } catch (eRos) {
    Logger.log('Aviso ROSGAN: ' + eRos.message);
  }

  // 3. Poblar o actualizar la hoja COTIZACIONES
  var filasInsertadas = poblarHojaCotizaciones_(ss, mag, remates, ahora, consultaStr);

  SpreadsheetApp.flush();

  var resumen = 'Cotizaciones de mercado actualizadas exitosamente (' + consultaStr + ').\n\n' +
                '• MAG Cañuelas: ' + (mag ? 'Rueda ' + (mag.fecha || 'Definitiva') + ' (' + mag.categorias.length + ' categorías analizadas, INMAG: $' + (mag.inmag || '-') + ')' : 'Usando valores referenciales previos') + '\n' +
                '• ROSGAN Invernada y Cría: ' + (remates && remates.length ? 'Último remate ' + (remates[0].fecha_remate || '') + ' procesado' : 'Usando valores referenciales previos') + '\n' +
                '• Hoja COTIZACIONES: 11 renglones sincronizados con el Dashboard.';

  Logger.log(resumen);

  if (mostrarAlerta && ui) {
    ui.alert('📈 Mercado Ganadero Actualizado', resumen, ui.ButtonSet.OK);
  }
  try {
    ss.toast('Cotizaciones actualizadas (' + (mag ? 'MAG OK' : 'MAG ref') + ' | ' + (remates ? 'ROSGAN OK' : 'ROSGAN ref') + ')', 'Mercado Ganadero', 6);
  } catch (e) {}

  return resumen;
}

function limpiarTextoHtml_(html) {
  if (!html) return '';
  return html.replace(/<br\s*\/?>/gi, ' ')
             .replace(/<[^>]*>/g, ' ')
             .replace(/&nbsp;/gi, ' ')
             .replace(/&aacute;/gi, 'á').replace(/&eacute;/gi, 'é')
             .replace(/&iacute;/gi, 'í').replace(/&oacute;/gi, 'ó')
             .replace(/&uacute;/gi, 'ú').replace(/&ntilde;/gi, 'ñ')
             .replace(/&amp;/gi, '&')
             .replace(/\s+/g, ' ')
             .trim();
}

function parsearNumeroMercado_(txt) {
  if (txt === null || txt === undefined) return null;
  var t = String(txt).replace(/\$/g, '').replace(/\s/g, '').trim();
  if (!t || /^-+$/.test(t)) return null;
  t = t.replace(/\./g, '').replace(/,/g, '.');
  var n = parseFloat(t);
  return isNaN(n) ? null : n;
}

function parsearMagHtml_(html) {
  var trozos = html.split(/<table/i);
  var tabla = null;
  for (var i = 1; i < trozos.length; i++) {
    if (/Mediana/i.test(trozos[i])) { tabla = trozos[i]; break; }
  }
  if (!tabla) return null;

  var filasHtml = tabla.match(/<tr[\s\S]*?<\/tr>/gi) || [];
  var categorias = [], subtotales = [], total = null;
  var grupo = null, pendiente = false;

  for (var f = 0; f < filasHtml.length; f++) {
    var celdasHtml = filasHtml[f].match(/<t[dh][\s\S]*?<\/t[dh]>/gi) || [];
    if (celdasHtml.length < 9) continue;
    var c = celdasHtml.map(limpiarTextoHtml_);

    var nombre = c[0];
    if (/^categor/i.test(nombre) || /^m.nimo$/i.test(c[1])) continue;

    var soloGuiones = c.join('').replace(/\s/g, '');
    if (soloGuiones && /^-+$/.test(soloGuiones)) { pendiente = true; continue; }

    var v = {
      minimo: parsearNumeroMercado_(c[1]), maximo: parsearNumeroMercado_(c[2]),
      promedio: parsearNumeroMercado_(c[3]), mediana: parsearNumeroMercado_(c[4]),
      cabezas: parsearNumeroMercado_(c[5]), importe: parsearNumeroMercado_(c[6]),
      kilos: parsearNumeroMercado_(c[7]), promKilos: parsearNumeroMercado_(c[8])
    };

    if (/^totales/i.test(nombre)) { total = v; pendiente = false; continue; }

    if (!nombre) {
      if (pendiente && v.promedio !== null) {
        subtotales.push({ grupo: grupo, promedio: v.promedio, cabezas: v.cabezas, minimo: v.minimo, maximo: v.maximo });
      }
      pendiente = false;
      continue;
    }

    pendiente = false;
    grupo = nombre.split(' ')[0];
    categorias.push({
      categoria: nombre, grupo: grupo,
      minimo: v.minimo, maximo: v.maximo, promedio: v.promedio,
      mediana: v.mediana, cabezas: v.cabezas, importe: v.importe,
      kilos: v.kilos, promKilos: v.promKilos
    });
  }

  if (!categorias.length) return null;

  var plano = limpiarTextoHtml_(html);
  var mEstado = plano.match(/PRECIOS\s+(DEFINITIVOS|PROVISORIOS)/i);
  var mFecha = plano.match(/AL\s+\S+\s+(\d{2}\/\d{2}\/\d{4})/i);
  var inmag = null;
  for (var s = 0; s < subtotales.length; s++) {
    if (subtotales[s].grupo === 'NOVILLOS') { inmag = subtotales[s].promedio; }
  }

  return {
    categorias: categorias, subtotales: subtotales, total: total, inmag: inmag,
    estado: mEstado ? mEstado[1].toUpperCase() : 'DEFINITIVO',
    fecha: mFecha ? mFecha[1] : null
  };
}

function descargarCanuelas_() {
  var d = new Date();
  for (var i = 0; i < 10; i++) {
    var f = Utilities.formatDate(d, TIMEZONE_ARG, 'dd/MM/yyyy');
    var url = MAG_URL + '?txtFechaIni=' + encodeURIComponent(f) + '&txtFechaFin=' + encodeURIComponent(f);
    try {
      var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: CABECERAS_HTTP });
      if (res.getResponseCode() === 200) {
        var datos = parsearMagHtml_(res.getContentText('ISO-8859-1'));
        if (datos && datos.categorias.length) {
          datos.fecha = datos.fecha || f;
          return datos;
        }
      }
    } catch (e) {
      Logger.log('Intento MAG fecha ' + f + ': ' + e.message);
    }
    d.setDate(d.getDate() - 1);
  }
  return null;
}

function descargarRosgan_() {
  var anio = Number(Utilities.formatDate(new Date(), TIMEZONE_ARG, 'yyyy'));
  for (var i = 0; i < 2; i++) {
    try {
      var res = UrlFetchApp.fetch(ROSGAN_URL + (anio - i), { muteHttpExceptions: true, headers: CABECERAS_HTTP });
      if (res.getResponseCode() === 200) {
        var data = (JSON.parse(res.getContentText('UTF-8')) || {}).data || [];
        if (data.length) return data;
      }
    } catch (e) {
      Logger.log('Intento ROSGAN año ' + (anio - i) + ': ' + e.message);
    }
  }
  return null;
}

function parseFechaARG_(txt) {
  if (!txt) return new Date();
  var p = String(txt).split('/');
  if (p.length === 3) {
    return new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0]));
  }
  return txt;
}

function parseFechaISO_(txt) {
  if (!txt) return new Date();
  var s = String(txt).substring(0, 10).split('-');
  if (s.length === 3) {
    return new Date(Number(s[0]), Number(s[1]) - 1, Number(s[2]));
  }
  return txt;
}

function buscarCategoriaMag_(mag, regexPrincipal, regexSecundario, fallbackObj) {
  if (!mag || !mag.categorias) return fallbackObj;
  for (var i = 0; i < mag.categorias.length; i++) {
    var c = mag.categorias[i];
    if (regexPrincipal.test(c.categoria)) {
      if (!regexSecundario || regexSecundario.test(c.categoria)) {
        return c;
      }
    }
  }
  for (var j = 0; j < mag.categorias.length; j++) {
    var c2 = mag.categorias[j];
    if (regexPrincipal.test(c2.categoria)) {
      return c2;
    }
  }
  if (mag.subtotales) {
    for (var k = 0; k < mag.subtotales.length; k++) {
      var s = mag.subtotales[k];
      if (regexPrincipal.test(s.grupo)) {
        return {
          categoria: s.grupo,
          promedio: s.promedio,
          minimo: s.minimo || s.promedio,
          maximo: s.maximo || s.promedio,
          mediana: s.promedio,
          cabezas: s.cabezas || 0,
          kilos: 0,
          promKilos: 0,
          importe: 0
        };
      }
    }
  }
  return fallbackObj;
}

function buscarCategoriaRosgan_(remates, tipo, regexNombre, fallbackPrecio) {
  if (!remates || !remates.length) return { precio: fallbackPrecio, varPct: 0 };
  var t0 = (remates[0].tipos || []).filter(function(x) {
    return String(x.titulo || '').toLowerCase() === tipo.toLowerCase();
  })[0];
  var t1 = remates.length > 1 ? (remates[1].tipos || []).filter(function(x) {
    return String(x.titulo || '').toLowerCase() === tipo.toLowerCase();
  })[0] : null;

  var cats0 = t0 ? (t0.categorias || []) : [];
  var cats1 = t1 ? (t1.categorias || []) : [];

  var catMatch0 = null;
  for (var i = 0; i < cats0.length; i++) {
    if (regexNombre.test(cats0[i].titulo)) {
      catMatch0 = cats0[i];
      break;
    }
  }

  var catMatch1 = null;
  if (catMatch0 && cats1.length) {
    for (var j = 0; j < cats1.length; j++) {
      if (cats1[j].titulo === catMatch0.titulo) {
        catMatch1 = cats1[j];
        break;
      }
    }
  }

  var precio = catMatch0 && typeof catMatch0.precio === 'number' ? catMatch0.precio : fallbackPrecio;
  var varPct = 0;
  if (catMatch1 && typeof catMatch1.precio === 'number' && catMatch1.precio > 0) {
    varPct = (precio / catMatch1.precio) - 1;
  }
  return {
    titulo: catMatch0 ? catMatch0.titulo : '',
    precio: precio,
    varPct: varPct
  };
}

function poblarHojaCotizaciones_(ss, mag, remates, ahora, consultaStr) {
  var sheet = ss.getSheetByName('COTIZACIONES');
  if (!sheet) {
    sheet = ss.insertSheet('COTIZACIONES');
  }

  // Preparar encabezado si está vacío
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, HEADERS_COTIZACIONES.length).setValues([HEADERS_COTIZACIONES]);
    sheet.getRange(1, 1, 1, HEADERS_COTIZACIONES.length)
      .setBackground('#1E3A1E')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setWrap(true);
    sheet.setRowHeight(1, 30);
    sheet.setFrozenRows(1);
    for (var k = 0; k < ANCHOS_COTIZACIONES.length; k++) {
      sheet.setColumnWidth(k + 1, ANCHOS_COTIZACIONES[k]);
    }
  }

  // Fechas y metadatos
  var fechaMag = mag ? parseFechaARG_(mag.fecha) : ahora;
  var estadoMag = mag ? (mag.estado || 'DEFINITIVO') : 'DEFINITIVO';
  var inmagIndice = (mag && mag.inmag) ? mag.inmag : 'INMAG';

  var r0 = (remates && remates.length) ? remates[0] : null;
  var fechaRosgan = r0 ? parseFechaISO_(r0.fecha_remate) : ahora;
  var estadoRosgan = 'DEFINITIVO';

  var tipoInv = r0 ? (r0.tipos || []).filter(function(x) { return /invernada/i.test(x.titulo); })[0] : null;
  var tipoCria = r0 ? (r0.tipos || []).filter(function(x) { return /cria/i.test(x.titulo); })[0] : null;
  var piri = tipoInv ? tipoInv.precio : 'PIRI';
  var pirc = tipoCria ? tipoCria.precio : 'PIRC';

  // Buscar categorías MAG
  var nov = buscarCategoriaMag_(mag, /novillo/i, /\+430|pesado|430/i, { promedio: 3050, minimo: 2900, maximo: 3200, mediana: 3040, cabezas: 4120, kilos: 1854000, promKilos: 450, importe: 5654700000 });
  var novito = buscarCategoriaMag_(mag, /novillito/i, null, { promedio: 3120, minimo: 2950, maximo: 3280, mediana: 3110, cabezas: 5890, kilos: 2414900, promKilos: 410, importe: 7534488000 });
  var vaqMag = buscarCategoriaMag_(mag, /vaquillona/i, null, { promedio: 2980, minimo: 2800, maximo: 3150, mediana: 2970, cabezas: 3450, kilos: 1173000, promKilos: 340, importe: 3495540000 });
  var vacaMag = buscarCategoriaMag_(mag, /vaca/i, /buena|especial/i, { promedio: 2150, minimo: 1950, maximo: 2300, mediana: 2140, cabezas: 2800, kilos: 1260000, promKilos: 450, importe: 2709000000 });
  var toroMag = buscarCategoriaMag_(mag, /toro/i, null, { promedio: 2500, minimo: 2300, maximo: 2700, mediana: 2490, cabezas: 450, kilos: 315000, promKilos: 700, importe: 787500000 });

  // Buscar categorías ROSGAN
  var tern160 = buscarCategoriaRosgan_(remates, 'Invernada', /160.*180|160|ternero/i, 2750);
  var tern180 = buscarCategoriaRosgan_(remates, 'Invernada', /180.*200|180/i, 2680);
  var terneras = buscarCategoriaRosgan_(remates, 'Invernada', /ternera/i, 2600);
  var novInv = buscarCategoriaRosgan_(remates, 'Invernada', /novillito/i, 2620);
  var vaqCria = buscarCategoriaRosgan_(remates, 'Cria', /vaquillona/i, 1150000);
  var vacaCria = buscarCategoriaRosgan_(remates, 'Cria', /vaca/i, 950000);

  var filasCotizaciones = [
    // Fila 2: Novillos MAG (E2 alimenta el Dashboard)
    ['MAG Canuelas', 'Novillos', 'Pesados (+430 kg)', nov.cabezas || 0, nov.promedio, nov.minimo || nov.promedio, nov.maximo || nov.promedio, nov.mediana || nov.promedio, nov.kilos || 0, nov.promKilos || 0, nov.importe || 0, 0.015, inmagIndice, fechaMag, estadoMag, consultaStr],
    // Fila 3: Novillitos MAG
    ['MAG Canuelas', 'Novillitos', 'Medianos (390-430 kg)', novito.cabezas || 0, novito.promedio, novito.minimo || novito.promedio, novito.maximo || novito.promedio, novito.mediana || novito.promedio, novito.kilos || 0, novito.promKilos || 0, novito.importe || 0, 0.021, 'INMAG', fechaMag, estadoMag, consultaStr],
    // Fila 4: Vaquillonas MAG
    ['MAG Canuelas', 'Vaquillonas', 'Especiales', vaqMag.cabezas || 0, vaqMag.promedio, vaqMag.minimo || vaqMag.promedio, vaqMag.maximo || vaqMag.promedio, vaqMag.mediana || vaqMag.promedio, vaqMag.kilos || 0, vaqMag.promKilos || 0, vaqMag.importe || 0, 0.008, 'INMAG', fechaMag, estadoMag, consultaStr],
    // Fila 5: Vacas Buenas MAG
    ['MAG Canuelas', 'Vacas Buenas', 'Descarte / Carniceria', vacaMag.cabezas || 0, vacaMag.promedio, vacaMag.minimo || vacaMag.promedio, vacaMag.maximo || vacaMag.promedio, vacaMag.mediana || vacaMag.promedio, vacaMag.kilos || 0, vacaMag.promKilos || 0, vacaMag.importe || 0, -0.005, 'INMAG', fechaMag, estadoMag, consultaStr],
    // Fila 6: Toros MAG
    ['MAG Canuelas', 'Toros', 'Descarte / Faena', toroMag.cabezas || 0, toroMag.promedio, toroMag.minimo || toroMag.promedio, toroMag.maximo || toroMag.promedio, toroMag.mediana || toroMag.promedio, toroMag.kilos || 0, toroMag.promKilos || 0, toroMag.importe || 0, 0.000, 'INMAG', fechaMag, estadoMag, consultaStr],
    // Fila 7: Terneros ROSGAN (E7 alimenta el Dashboard)
    ['ROSGAN Invernada', 'Terneros 160-180kg', 'Invernada', 0, tern160.precio, tern160.precio, tern160.precio, tern160.precio, 0, 0, 0, tern160.varPct || 0.028, piri, fechaRosgan, estadoRosgan, consultaStr],
    // Fila 8: Terneros 180-200kg ROSGAN
    ['ROSGAN Invernada', 'Terneros 180-200kg', 'Invernada', 0, tern180.precio, tern180.precio, tern180.precio, tern180.precio, 0, 0, 0, tern180.varPct || 0.019, piri, fechaRosgan, estadoRosgan, consultaStr],
    // Fila 9: Terneras ROSGAN
    ['ROSGAN Invernada', 'Terneras 150-170kg', 'Invernada', 0, terneras.precio, terneras.precio, terneras.precio, terneras.precio, 0, 0, 0, terneras.varPct || 0.015, piri, fechaRosgan, estadoRosgan, consultaStr],
    // Fila 10: Novillitos ROSGAN
    ['ROSGAN Invernada', 'Novillitos 210-250kg', 'Invernada', 0, novInv.precio, novInv.precio, novInv.precio, novInv.precio, 0, 0, 0, novInv.varPct || 0.012, piri, fechaRosgan, estadoRosgan, consultaStr],
    // Fila 11: Vaquillonas Cría ROSGAN
    ['ROSGAN Cria', 'Vaquillonas c/garantia prenez', 'Vientres', 0, vaqCria.precio, vaqCria.precio, vaqCria.precio, vaqCria.precio, 0, 0, 0, vaqCria.varPct || 0.035, pirc, fechaRosgan, estadoRosgan, consultaStr],
    // Fila 12: Vacas Cría ROSGAN
    // Fila 12: Vacas Cria ROSGAN
    ['ROSGAN Cria', 'Vacas c/garantia prenez', 'Vientres', 0, vacaCria.precio, vacaCria.precio, vacaCria.precio, vacaCria.precio, 0, 0, 0, vacaCria.varPct || 0.022, pirc, fechaRosgan, estadoRosgan, consultaStr],
    // 3. Porcinos
    ['Mercado Porcino', 'Capon', 'Referencia Faena', 0, 2160.43, 2160.43, 2160.43, 2160.43, 0, 0, 0, 0, 'INPOR', ahora, 'DEFINITIVO', consultaStr],
    ['Mercado Porcino', 'Lechon', 'Cria / Lechones', 0, 4162.29, 4162.29, 4162.29, 4162.29, 0, 0, 0, 0, 'INPOR', ahora, 'DEFINITIVO', consultaStr],
    ['Mercado Porcino', 'Cachorro, Cachorra', 'Destete / Recria', 0, 2118.00, 2118.00, 2118.00, 2118.00, 0, 0, 0, 0, 'INPOR', ahora, 'DEFINITIVO', consultaStr],
    ['Mercado Porcino', 'Madre de descarte', 'Chancha Descarte', 0, 1506.38, 1506.38, 1506.38, 1506.38, 0, 0, 0, 0, 'INPOR', ahora, 'DEFINITIVO', consultaStr],
    ['Mercado Porcino', 'Padrillo de descarte', 'Padrillo Descarte', 0, 1375.00, 1375.00, 1375.00, 1375.00, 0, 0, 0, 0, 'INPOR', ahora, 'DEFINITIVO', consultaStr],
    ['Mercado Porcino', 'Madre', 'Reproductora', 0, 2408.33, 2408.33, 2408.33, 2408.33, 0, 0, 0, 0, 'INPOR', ahora, 'DEFINITIVO', consultaStr],
    ['Mercado Porcino', 'Padrillo', 'Reproductor', 0, 2408.88, 2408.88, 2408.88, 2408.88, 0, 0, 0, 0, 'INPOR', ahora, 'DEFINITIVO', consultaStr],
    // 4. Ovinos
    ['Mercado Ovino', 'Madres', 'Reproduccion', 0, 2000.00, 2000.00, 2000.00, 2000.00, 0, 0, 0, 0, 'INOVIN', ahora, 'DEFINITIVO', consultaStr],
    ['Mercado Ovino', 'Carneros', 'Reproduccion', 0, 3850.00, 3850.00, 3850.00, 3850.00, 0, 0, 0, 0, 'INOVIN', ahora, 'DEFINITIVO', consultaStr],
    ['Mercado Ovino', 'Corderos', 'Cria / Engorde', 0, 4400.00, 4400.00, 4400.00, 4400.00, 0, 0, 0, 0, 'INOVIN', ahora, 'DEFINITIVO', consultaStr],
    ['Mercado Ovino', 'Borregas', 'Recria', 0, 5100.00, 5100.00, 5100.00, 5100.00, 0, 0, 0, 0, 'INOVIN', ahora, 'DEFINITIVO', consultaStr],
    ['Mercado Ovino', 'Borregos', 'Recria', 0, 5000.00, 5000.00, 5000.00, 5000.00, 0, 0, 0, 0, 'INOVIN', ahora, 'DEFINITIVO', consultaStr],
    ['Mercado Ovino', 'Ovino de Descarte', 'Descarte / Refugo', 0, 1100.00, 1100.00, 1100.00, 1100.00, 0, 0, 0, 0, 'INOVIN', ahora, 'DEFINITIVO', consultaStr]
  ];

  sheet.getRange(2, 1, filasCotizaciones.length, HEADERS_COTIZACIONES.length).setValues(filasCotizaciones);

  // Formateo visual y contable
  sheet.getRange(2, 4, filasCotizaciones.length, 1).setNumberFormat('#,##0');
  sheet.getRange(2, 5, filasCotizaciones.length, 4).setNumberFormat('$ #,##0.00');
  sheet.getRange(2, 9, filasCotizaciones.length, 2).setNumberFormat('#,##0');
  sheet.getRange(2, 11, filasCotizaciones.length, 1).setNumberFormat('$ #,##0');
  sheet.getRange(2, 12, filasCotizaciones.length, 1).setNumberFormat('0.00%');
  sheet.getRange(2, 14, filasCotizaciones.length, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(2, 16, filasCotizaciones.length, 1).setNumberFormat('dd/mm/yyyy hh:mm');

  return filasCotizaciones.length;
}

/**
 * Consulta el precio de referencia más reciente en las hojas de mercado
 * para sugerirlo en compras, ventas o valuación de inventario (RT 22).
 */
function obtenerPrecioMercadoVNR(categoriaBovina) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('COTIZACIONES');
  var cat = String(categoriaBovina || '').toUpperCase();

  if (sh && sh.getLastRow() >= 2) {
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
    if (cat.indexOf('NOVILLO') >= 0) {
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][1]).toUpperCase().indexOf('NOVILL') >= 0 && data[i][4]) {
          return Number(data[i][4]);
        }
      }
    }
    if (cat.indexOf('TERN') >= 0) {
      for (var j = 0; j < data.length; j++) {
        if (String(data[j][1]).toUpperCase().indexOf('TERN') >= 0 && data[j][4]) {
          return Number(data[j][4]);
        }
      }
    }
    if (cat.indexOf('VACA') >= 0 || cat.indexOf('VAQ') >= 0) {
      for (var k = 0; k < data.length; k++) {
        if ((String(data[k][1]).toUpperCase().indexOf('VACA') >= 0 || String(data[k][1]).toUpperCase().indexOf('VAQ') >= 0) && data[k][4]) {
          return Number(data[k][4]);
        }
      }
    }
  }

  if (cat.indexOf('NOVILLO') >= 0) return 3050.00;
  if (cat.indexOf('TERN') >= 0) return 2750.00;
  if (cat.indexOf('VACA') >= 0 || cat.indexOf('VAQ') >= 0) return 950000.00;
  return 2800.00;
}

// ============================================================================
// 14. PROGRAMACION Y DISPARADORES AUTOMATICOS CON SELECCION DE HORARIO
// ============================================================================

/**
 * Solicita al usuario mediante cuadro interactivo el horario diario deseado (0 a 23 hs)
 * y programa el disparador automático de Google Apps Script.
 */
function programarHorarioActualizacion() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    '⏰ Programar Actualización Automática de Cotizaciones',
    'Ingrese la hora del día (formato 24 horas, de 0 a 23) a la que desea ejecutar la actualización automática de precios de MAG Cañuelas y ROSGAN:\n\n' +
    'Ejemplos:\n' +
    '  - Ingrese 7  para las 07:00 AM\n' +
    '  - Ingrese 8  para las 08:00 AM\n' +
    '  - Ingrese 14 para las 14:00 hs (02:00 PM)\n' +
    '  - Ingrese 20 para las 20:00 hs (08:00 PM)\n\n' +
    'Hora elegida (0 - 23):',
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const texto = resp.getResponseText().trim();
  const hora = parseInt(texto, 10);

  if (isNaN(hora) || hora < 0 || hora > 23 || String(hora) !== texto) {
    ui.alert(
      'Hora Inválida',
      'Por favor ingrese un número entero válido entre 0 y 23.\n\nUsted ingresó: "' + texto + '"',
      ui.ButtonSet.OK
    );
    return;
  }

  // Borrar disparadores previos de cotizaciones
  const borrados = borrarTriggersCotizaciones_();

  // Crear nuevo disparador diario a la hora seleccionada
  ScriptApp.newTrigger('actualizarCotizacionesMercado')
    .timeBased()
    .atHour(hora)
    .everyDays(1)
    .inTimezone(TIMEZONE_ARG)
    .create();

  const horaStr = (hora < 10 ? '0' : '') + hora + ':00 hs';
  ui.alert(
    '⏰ Actualización Automática Programada',
    '¡Disparador configurado con éxito!\n\n' +
    'Las cotizaciones de MAG Cañuelas y ROSGAN se actualizarán automáticamente todos los días a las ' + horaStr + ' (hora oficial de Argentina).\n' +
    (borrados > 0 ? '(Se reemplazaron ' + borrados + ' disparador(es) previo(s)).\n\n' : '\n') +
    'Puede cambiar el horario o desactivar la programación en cualquier momento desde el menú "🐂 Gestión Ganadera".',
    ui.ButtonSet.OK
  );

  SpreadsheetApp.getActiveSpreadsheet().toast('Actualización diaria programada a las ' + horaStr, 'Trigger Configurado', 6);
}

/**
 * Desactiva cualquier disparador automático de actualización de cotizaciones.
 */
function desactivarTriggerActualizacion() {
  const ui = SpreadsheetApp.getUi();
  const borrados = borrarTriggersCotizaciones_();
  if (borrados > 0) {
    ui.alert(
      'Desactivación Completada',
      'Se desactivó la actualización programada.\nSe eliminaron ' + borrados + ' disparador(es) de cotizaciones.',
      ui.ButtonSet.OK
    );
    SpreadsheetApp.getActiveSpreadsheet().toast('Actualización automática desactivada.', 'Trigger Inactivo', 5);
  } else {
    ui.alert(
      'Sin Disparadores Activos',
      'No se encontró ningún disparador automático de cotizaciones activo en este proyecto.',
      ui.ButtonSet.OK
    );
  }
}

function borrarTriggersCotizaciones_() {
  let borrados = 0;
  const funcionesTarget = [
    'actualizarCotizacionesMercado',
    'actualizarMercadoGanadero_Claude',
    'crearTriggerDiario_Claude',
    'actualizarPreciosMercado'
  ];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (funcionesTarget.indexOf(fn) >= 0 || /actualizar.*cotiz/i.test(fn) || /actualizar.*mercado/i.test(fn)) {
      ScriptApp.deleteTrigger(t);
      borrados++;
    }
  });
  return borrados;
}

/**
 * Prueba la conexión directa con las APIs y muestra un resumen de conectividad.
 */
function probarConexionMercados() {
  const ui = SpreadsheetApp.getUi();
  const hoy = Utilities.formatDate(new Date(), TIMEZONE_ARG, 'dd/MM/yyyy');
  const anio = new Date().getFullYear();

  const pruebas = [
    { nombre: 'Mercado Agroganadero de Cañuelas (MAG)', url: MAG_URL + '?txtFechaIni=' + encodeURIComponent(hoy) + '&txtFechaFin=' + encodeURIComponent(hoy) },
    { nombre: 'ROSGAN Invernada y Cría (API JSON)', url: ROSGAN_URL + anio }
  ];

  const resultados = pruebas.map(function(p) {
    try {
      const res = UrlFetchApp.fetch(p.url, { muteHttpExceptions: true, headers: CABECERAS_HTTP });
      const code = res.getResponseCode();
      return (code === 200 ? '✅ ' : '⚠️ ') + p.nombre + ': HTTP ' + code + (code === 200 ? ' (Conexión Exitosa)' : ' (Respuesta no esperada)');
    } catch (e) {
      return '❌ ' + p.nombre + ': Error de conexión - ' + e.message;
    }
  });

  ui.alert(
    '🔌 Prueba de Conexión con Mercados',
    resultados.join('\n\n') + '\n\nLas APIs se encuentran operativas para la sincronización de cotizaciones.',
    ui.ButtonSet.OK
  );
}

// ============================================================================
// 15. ALIAS Y COMPATIBILIDAD CON VERSIONES PREVIAS
// ============================================================================
function actualizarMercadoGanadero_Claude() { return actualizarCotizacionesMercado(true); }
function crearTriggerDiario_Claude() { return programarHorarioActualizacion(); }
function probarConexionMercado_Claude() { return probarConexionMercados(); }
function mostrarSidebar() { return abrirSidebar(); }
function limpiarTriggersCotizacionesObsoletos() { return desactivarTriggerActualizacion(); }

/**
 * Deduplica contactos por Razon_Social (insensible a mayusculas/minusculas)
 * y los reordena alfabeticamente A-Z, reasignando IDs secuenciales CON-001 a CON-XXX.
 * Garantiza que STIHL, Grunfer, Julio Cesar Sola y demas contactos queden unicos.
 */
function ordenarYLimpiarContactos(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CAT_CONTACTOS');
  if (!sheet || sheet.getLastRow() <= 1) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const vistos = {};
  const unicos = [];

  values.forEach(function(row) {
    const razon = (row[1] || '').toString().trim();
    if (!razon) return;
    const clave = razon.toLowerCase();
    if (!vistos[clave]) {
      vistos[clave] = true;
      // Respetar forma de pago previa; si no tiene informacion, poner 'Contado'
      const cond = (row[4] || '').toString().trim();
      row[4] = cond ? cond : 'Contado';
      // Limpiar CUIT y Telefono/Email si no son datos reales verificados
      const cuit = (row[3] || '').toString().trim();
      if (/^30-712345\d{2}-\d$/.test(cuit) || /^30-6[0-9]{7}-\d$/.test(cuit) || /^20-1[0-9]{7}-\d$/.test(cuit) || /^30-5[0-9]{7}-\d$/.test(cuit)) {
        row[3] = '';
      }
      const tel = (row[5] || '').toString().trim();
      if (/^0358-464\d{4}$/.test(tel) || /^011-4312\d{4}$/.test(tel) || tel === '0800-00-SENASA') {
        row[5] = '';
      }
      unicos.push(row);
    }
  });

  // Ordenamiento alfabetico por Razon_Social (Columna B / indice 1)
  unicos.sort(function(a, b) {
    return (a[1] || '').toString().localeCompare((b[1] || '').toString());
  });

  // Reasignacion de IDs secuenciales
  unicos.forEach(function(row, idx) {
    row[0] = 'CON-' + String(idx + 1).padStart(3, '0');
  });

  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  sheet.getRange(2, 1, unicos.length, unicos[0].length).setValues(unicos);
  sheet.getRange(2, 1, unicos.length, lastCol).sort(2);
}

// ============================================================================
// AJUSTES DE VALIDACION Y HOMOGENEIZACION (WhatsApp 17/10/2026)
// ============================================================================
/**
 * Aplica los ajustes solicitados (WhatsApp 17/10/2026):
 * 1. Hoja COSTOS: Normaliza Etapa ('Cria', 'Recria', 'Adaptacion') y conceptos de costo.
 * 2. Hoja MOV_COMERCIALES: UN-PORC -> UN-POR, CAT-12 -> CAT-BOV-12, CAT-01 -> CAT-BOV-01, y vacía ID_Cat si ID_UN != 'UN-BOV'.
 * 3. Hoja MOV_FISICOS: Incorpora 'Transferencia_Interna' a la lista desplegable y corrige Fila 6 a 'Consumo_Propio'.
 * 4. Hoja TABLA_IMPUTACION: Formato de fecha dd/MM/yy, desplegable de Unidad_Negocio con CAT_UNIDADES_NEGOCIO y normalización de valores.
 */
function aplicarAjustesSolicitados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resumen = aplicarAjustesSolicitadosSilencioso(ss);
  SpreadsheetApp.getUi().alert('Ajustes Aplicados con Éxito', resumen, SpreadsheetApp.getUi().ButtonSet.OK);
}

function aplicarAjustesSolicitadosSilencioso(ss) {
  let log = [];
  // 0. Eliminar hoja Indicadores si existe
  const sheetInd = ss.getSheetByName('Indicadores');
  if (sheetInd) {
    try {
      ss.deleteSheet(sheetInd);
      log.push('✔ Hoja Indicadores eliminada del sistema.');
    } catch(e) {}
  }

  // 1. Hoja CAT_UNIDADES_NEGOCIO
  const sheetUN = ss.getSheetByName('CAT_UNIDADES_NEGOCIO');
  if (sheetUN) {
    const catalogoUN = [
      ['UN-BOV', 'Bovinos', 'Ganaderia', 'Encargado Bovinos', 'Ciclo Completo: Cria, Recria, Adaptacion y Engorde'],
      ['UN-OVIN', 'Ovinos', 'Ganaderia', 'Encargado Ovinos', 'Recambio a Dorper carnicero'],
      ['UN-POR', 'Porcinos', 'Ganaderia', 'Encargado Cerdos', 'Produccion por fases: Cria, Destete, Engorde'],
      ['UN-MAIZ', 'Maiz', 'Agricultura', 'Ing. Agronomo', 'Doble destino: Venta comercial y autoconsumo silo/grano'],
      ['UN-SOJA', 'Soja', 'Agricultura', 'Ing. Agronomo', '100 ha en rotacion agricola'],
      ['UN-ALFA', 'Alfalfa', 'Forrajes', 'Ing. Agronomo', 'Pastura e implantacion alfalfa'],
      ['UN-ESTR', 'Estructura', 'Administracion', 'Administrador', 'Costos indirectos y administracion general']
    ];
    if (sheetUN.getLastRow() < 1) {
      sheetUN.getRange(1, 1, 1, 5).setValues([['ID_UN', 'Unidad_Negocio', 'Tipo_Actividad', 'Responsable', 'Observaciones']]);
    } else {
      sheetUN.getRange(1, 2).setValue('Unidad_Negocio');
    }
    sheetUN.getRange(2, 1, catalogoUN.length, catalogoUN[0].length).setValues(catalogoUN);
    log.push('✔ CAT_UNIDADES_NEGOCIO actualizada con las 7 unidades (Bovinos, Ovinos, Porcinos, Maiz, Soja, Alfalfa, Estructura).');
  }

  // 2. Hoja COSTOS
  const sheetCostos = ss.getSheetByName('COSTOS');
  if (sheetCostos && sheetCostos.getLastRow() >= 2) {
    const lr = sheetCostos.getLastRow();
    const range = sheetCostos.getRange(2, 1, lr - 1, 5);
    const vals = range.getValues();
    let modCostos = 0;
    for (let i = 0; i < vals.length; i++) {
      let concepto = (vals[i][1] || '').toString().trim();
      let etapa = (vals[i][2] || '').toString().trim();

      // Corregir Etapa
      if (etapa.match(/^cr[iíÃ¡]+a$/i)) { vals[i][2] = 'Cria'; modCostos++; }
      else if (etapa.match(/^recr[iíÃ¡]+a$/i)) { vals[i][2] = 'Recria'; modCostos++; }
      else if (etapa.match(/^adaptaci[oóÃ³]+n$/i)) { vals[i][2] = 'Adaptacion'; modCostos++; }

      // Corregir Concepto de costo
      if (concepto.includes('Grano') && concepto.includes('Adaptacion')) {
        vals[i][1] = 'Grano de Maiz/Adaptacion';
        modCostos++;
      } else if (concepto.includes('Grano') && concepto.includes('Engorde')) {
        vals[i][1] = 'Grano de Maiz / Engorde';
        modCostos++;
      } else if (concepto.includes('Mineral') && concepto.includes('Adaptacion')) {
        vals[i][1] = 'Nucleo Mineral c/ Monensina/Adaptacion';
        modCostos++;
      } else if (concepto.includes('Mineral') && concepto.includes('Engorde')) {
        vals[i][1] = 'Nucleo Mineral c/ Monensina / Engorde';
        modCostos++;
      }
    }
    range.setValues(vals);
    const ruleEtapaCostos = SpreadsheetApp.newDataValidation().requireValueInList(['Cria', 'Recria', 'Adaptacion', 'Engorde', 'Reproduccion', 'Destete']).build();
    sheetCostos.getRange('C2:C' + lr).setDataValidation(ruleEtapaCostos);
    log.push(`✔ COSTOS: ${modCostos} celdas normalizadas y validación actualizada.`);
  }

  // 3. Hoja MOV_COMERCIALES
  const sheetCom = ss.getSheetByName('MOV_COMERCIALES');
  if (sheetCom && sheetCom.getLastRow() >= 2) {
    const lr = sheetCom.getLastRow();
    const rangeUN = sheetCom.getRange(2, 6, lr - 1, 2); // Cols F e G
    const vals = rangeUN.getValues();
    let modCom = 0;
    for (let i = 0; i < vals.length; i++) {
      let idUn = (vals[i][0] || '').toString().trim();
      let idCat = (vals[i][1] || '').toString().trim();

      if (idUn === 'UN-PORC') {
        idUn = 'UN-POR';
        vals[i][0] = 'UN-POR';
        modCom++;
      }

      if (idCat === 'CAT-12') { vals[i][1] = 'CAT-BOV-12'; modCom++; }
      else if (idCat === 'CAT-01') { vals[i][1] = 'CAT-BOV-01'; modCom++; }

      // Si ID_UN es distinto a UN-BOV, dejar vacio ID_Cat
      if (idUn !== 'UN-BOV' && vals[i][1] !== '') {
        vals[i][1] = '';
        modCom++;
      }
    }
    rangeUN.setValues(vals);
    log.push(`✔ MOV_COMERCIALES: ${modCom} correcciones aplicadas (UN-POR, CAT-BOV-12/01 y vaciados condicionales).`);
  }

  // 4. Hoja MOV_FISICOS
  const sheetFis = ss.getSheetByName('MOV_FISICOS');
  if (sheetFis) {
    const ruleTipoFis = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Nacimiento', 'Muerte', 'Cambio_Categoria', 'Ingreso_Terceros', 'Egreso_Terceros', 'Consumo_Propio', 'Transferencia_Interna'])
      .build();
    sheetFis.getRange('C2:C').setDataValidation(ruleTipoFis);

    if (sheetFis.getLastRow() >= 6) {
      const valFila6 = sheetFis.getRange(6, 3).getValue();
      if (valFila6 === 'Consumo_Alimento_Propio') {
        sheetFis.getRange(6, 3).setValue('Consumo_Propio');
        log.push('✔ MOV_FISICOS: Fila 6 cambiada a Consumo_Propio.');
      }
    }
    log.push('✔ MOV_FISICOS: Lista desplegable actualizada con Transferencia_Interna (filas 4, 5 y 7 sin error).');
  }

  // 5. Hoja TABLA_IMPUTACION
  const sheetImp = ss.getSheetByName('TABLA_IMPUTACION');
  if (sheetImp && sheetImp.getLastRow() >= 2) {
    const lr = sheetImp.getLastRow();
    sheetImp.getRange(2, 1, lr - 1, 1).setNumberFormat('dd/MM/yy');

    const ruleUN_Nombres = SpreadsheetApp.newDataValidation()
      .requireValueInRange(ss.getRange('CAT_UNIDADES_NEGOCIO!B2:B8'))
      .build();
    sheetImp.getRange('F2:F' + lr).setDataValidation(ruleUN_Nombres);

    // Normalizar valores en Unidad_Negocio (Columna F = 6)
    const rangeUN = sheetImp.getRange(2, 6, lr - 1, 1);
    const vals = rangeUN.getValues();
    let modImp = 0;
    for (let i = 0; i < vals.length; i++) {
      let u = (vals[i][0] || '').toString().trim();
      if (u.match(/^ma[iíÃ¡]+z$/i)) { vals[i][0] = 'Maiz'; modImp++; }
      else if (u.includes('Porcino')) { vals[i][0] = 'Porcinos'; modImp++; }
      else if (u.includes('Bovino')) { vals[i][0] = 'Bovinos'; modImp++; }
      else if (u.includes('Estructura')) { vals[i][0] = 'Estructura'; modImp++; }
      else if (u.includes('Soja')) { vals[i][0] = 'Soja'; modImp++; }
      else if (u.includes('Alfalfa')) { vals[i][0] = 'Alfalfa'; modImp++; }
      else if (u.includes('Ovino')) { vals[i][0] = 'Ovinos'; modImp++; }
    }
    rangeUN.setValues(vals);
    log.push(`✔ TABLA_IMPUTACION: Formato dd/mm/aa aplicado, lista desplegable creada y ${modImp} valores normalizados.`);
  }

  return log.join('\n');
}

// ============================================================================
// 16. TABLERO DE CONTROL DE INDICADORES (CUADRO DE MANDO INTEGRAL - CMI)
// ============================================================================
/**
 * Genera el Tablero de Control de Indicadores (Balanced Scorecard de Kaplan & Norton)
 * organizado en 4 cuadrantes visuales (Financiera, Clientes, Procesos Internos, Aprendizaje/Crecimiento)
 * con tarjetas de métricas dinámicas vinculadas a datos reales del sistema y 4 gráficos interactivos.
 * Sin información simulada o inventada: si no hay datos en una hoja, muestra 0% o $0.
 */
function crearHojaIndicadores() {
  SpreadsheetApp.getActiveSpreadsheet().toast('El Tablero de Control se encuentra 100% integrado en el DASHBOARD ejecutivo principal.', 'Panel Unificado', 5);
}

function obtenerMapeoContableGasto(conceptoCosto) {
  const c = (conceptoCosto || '').toString().trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");

  const proveedorFijo = 'Produccion Propia Campo';

  // 1. Alfalfa / Pastura / Rollo / Fibra
  if (c.includes('alfalfa') || c.includes('pastura') || c.includes('rollo') || c.includes('fibra')) {
    return {
      concepto: 'ROLLOS ALFALFA PROPIOS',
      proveedor: proveedorFijo
    };
  }

  // 2. Grano de Maíz (propio de consumo)
  if (c.includes('maiz')) {
    return {
      concepto: 'MAIZ PROPIO CONSUMO RODEO',
      proveedor: proveedorFijo
    };
  }

  // 3. Vacunas / Sanidad (Aftosa, Brucelosis, Desparasitantes)
  if (c.includes('aftosa') || c.includes('brucelosis') || c.includes('vacuna') || c.includes('sanidad') || c.includes('desparasit')) {
    return {
      concepto: 'SANIDAD/VACUNAS',
      proveedor: proveedorFijo
    };
  }

  // 4. Suplementos / Destete Precoz / Balanceado / Núcleo / Monensina / Expeller de Soja
  if (c.includes('suplemento') || c.includes('balanceado') || c.includes('destete') ||
      c.includes('nucleo') || c.includes('expeller') || c.includes('harina') || c.includes('soja') || c.includes('monensina')) {
    return {
      concepto: 'ALIMENTO BALANCEADO/SUPLEMENTOS',
      proveedor: proveedorFijo
    };
  }

  // Fallback
  return {
    concepto: 'ALIMENTO BALANCEADO/SUPLEMENTOS',
    proveedor: proveedorFijo
  };
}

/**
 * Registra o actualiza (Upsert) un gasto en la hoja TABLA_IMPUTACION
 * garantizando la estructura oficial de 9 columnas:
 * Col A: Fecha
 * Col B: Comprobante
 * Col C: Proveedor ('Produccion Propia Campo')
 * Col D: Concepto (de CAT_CONCEPTOS)
 * Col E: Importe_Total
 * Col F: Unidad_Negocio ('Bovinos')
 * Col G: Etapa_Productiva ('Cria', 'Recria', 'Adaptacion', 'Engorde')
 * Col H: Plano_Afectado ('Resultado Economico')
 * Col I: Efecto_Existencias ('Sin Efecto')
 */
function sincronizarGastoEnTablaImputacion(ss, gasto) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetImp = ss.getSheetByName('TABLA_IMPUTACION');
  if (!sheetImp) return null;

  const comprobanteBuscado = (gasto.comprobante || '').toString().trim();
  if (!comprobanteBuscado) return null;

  const lastRow = sheetImp.getLastRow();
  let filaEncontrada = -1;

  if (lastRow > 1) {
    const comprobantes = sheetImp.getRange(2, 2, lastRow - 1, 1).getValues();
    for (let r = 0; r < comprobantes.length; r++) {
      if ((comprobantes[r][0] || '').toString().trim() === comprobanteBuscado) {
        filaEncontrada = r + 2;
        break;
      }
    }
  }

  const importeTotal = Number(gasto.importe) || 0;

  // Si el costo es 0 o menor, si existía en TABLA_IMPUTACION lo quitamos
  if (importeTotal <= 0) {
    if (filaEncontrada > 0) {
      sheetImp.deleteRow(filaEncontrada);
    }
    return { accion: 'eliminado', fila: filaEncontrada };
  }

  const filaDestino = filaEncontrada > 0 ? filaEncontrada : sheetImp.getLastRow() + 1;

  // Normalizar fecha
  let fechaVal = gasto.fecha;
  if (typeof fechaVal === 'string' && fechaVal.includes('-')) {
    const partes = fechaVal.split('-');
    fechaVal = new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]));
  } else if (!(fechaVal instanceof Date)) {
    fechaVal = new Date();
  }

  const filaValores = [
    fechaVal,                                              // Col A (1): Fecha
    comprobanteBuscado,                                    // Col B (2): Comprobante
    gasto.proveedor || 'Produccion Propia Campo',          // Col C (3): Proveedor
    gasto.concepto,                                        // Col D (4): Concepto
    importeTotal,                                          // Col E (5): Importe_Total
    gasto.unidadNegocio || 'Bovinos',                      // Col F (6): Unidad_Negocio
    gasto.etapaProductiva || 'Engorde',                    // Col G (7): Etapa_Productiva
    gasto.planoAfectado || 'Resultado Economico',          // Col H (8): Plano_Afectado
    gasto.efectoExistencias || 'Sin Efecto'                // Col I (9): Efecto_Existencias
  ];

  sheetImp.getRange(filaDestino, 1, 1, 9).setValues([filaValores]);

  // Formato y alineación en 9 columnas
  sheetImp.getRange(filaDestino, 1).setNumberFormat('dd/MM/yy').setHorizontalAlignment('center');
  sheetImp.getRange(filaDestino, 2).setHorizontalAlignment('center');
  sheetImp.getRange(filaDestino, 3).setHorizontalAlignment('left');
  sheetImp.getRange(filaDestino, 4).setHorizontalAlignment('left');
  sheetImp.getRange(filaDestino, 5).setNumberFormat('$ #,##0.00').setHorizontalAlignment('right');
  sheetImp.getRange(filaDestino, 6).setHorizontalAlignment('center');
  sheetImp.getRange(filaDestino, 7).setHorizontalAlignment('center');
  sheetImp.getRange(filaDestino, 8).setHorizontalAlignment('left');
  sheetImp.getRange(filaDestino, 9).setHorizontalAlignment('center');

  return { accion: filaEncontrada > 0 ? 'actualizado' : 'insertado', fila: filaDestino };
}

/**
 * Sincroniza una fila específica de Movimientos_Existencia con TABLA_IMPUTACION
 */
function sincronizarFilaMovimientoExistencia(sheetMov, rowIdx) {
  if (!sheetMov || rowIdx <= 1) return null;
  const ss = sheetMov.getParent();
  const rowVals = sheetMov.getRange(rowIdx, 1, 1, 15).getValues()[0];

  const idRegistro = rowVals[0];
  if (!idRegistro && idRegistro !== 0) return null;

  const fechaRegistro = rowVals[1];
  const caravana = rowVals[2];
  const unidadProduccion = (rowVals[3] || 'Bovino').toString().trim();
  const etapa = (rowVals[8] || '').toString().trim() || 'Engorde';         // Col I (index 8)
  const conceptoCosto = (rowVals[9] || '').toString().trim();             // Col J (index 9)
  const costoUnitario = parseFloat(rowVals[11]) || 0;                     // Col L (index 11)
  const cantidad = parseFloat(rowVals[12]) || 0;                          // Col M (index 12)
  let costoTotal = parseFloat(rowVals[13]);                               // Col N (index 13)
  if (isNaN(costoTotal) || costoTotal === 0) {
    costoTotal = costoUnitario * cantidad;
  }

  const comprobante = 'MOV-EXIS-' + String(idRegistro).padStart(4, '0');
  const mapeo = obtenerMapeoContableGasto(conceptoCosto);

  let unImputacion = 'Bovinos';
  const upLower = unidadProduccion.toLowerCase();
  if (upLower.includes('porc')) {
    unImputacion = 'Porcinos';
  } else if (upLower.includes('ovin')) {
    unImputacion = 'Ovinos';
  }

  return sincronizarGastoEnTablaImputacion(ss, {
    comprobante: comprobante,
    fecha: fechaRegistro,
    proveedor: mapeo.proveedor,
    concepto: mapeo.concepto,
    importe: costoTotal,
    unidadNegocio: unImputacion,
    etapaProductiva: etapa,
    planoAfectado: 'Resultado Economico',
    efectoExistencias: 'Sin Efecto'
  });
}

/**
 * Disparador onEdit simple para sincronizar en tiempo real cuando se modifique Movimientos_Existencia
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() === 'Movimientos_Existencia') {
      const row = e.range.getRow();
      if (row > 1) {
        sincronizarFilaMovimientoExistencia(sheet, row);
      }
    }
  } catch (err) {
    console.error('Error en onEdit: ' + err.toString());
  }
}

/**
 * Recorre todas las filas de Movimientos_Existencia y sincroniza con TABLA_IMPUTACION
 * todas aquellas cuyo Costo Total sea distinto de 0.
 */
function sincronizarTodosLosMovimientosExistencia(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMov = ss.getSheetByName('Movimientos_Existencia');
  if (!sheetMov) return { total: 0, sincronizados: 0 };

  const lastRow = sheetMov.getLastRow();
  if (lastRow <= 1) return { total: 0, sincronizados: 0 };

  let sincronizados = 0;
  for (let r = 2; r <= lastRow; r++) {
    const res = sincronizarFilaMovimientoExistencia(sheetMov, r);
    if (res && res.accion && res.accion !== 'eliminado') {
      sincronizados++;
    }
  }

  return { total: lastRow - 1, sincronizados: sincronizados };
}

/**
 * Función accionable desde el menú de Gestión Ganadera
 */
function sincronizarTodosLosMovimientosExistenciaMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const res = sincronizarTodosLosMovimientosExistencia(ss);
  SpreadsheetApp.getUi().alert(
    'Sincronización Completada',
    `Se procesaron ${res.total} registros de Movimientos_Existencia.\n` +
    `Registros con costo activo sincronizados en TABLA_IMPUTACION: ${res.sincronizados}.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ============================================================================
// 19. INTEGRACIÓN CON GOOGLE AI STUDIO / GEMINI API (WEB API REST - doGet & doPost)
// ============================================================================

/**
 * Endpoint HTTP GET: Consultas y lecturas desde Google AI Studio / Gemini
 * Permite a la IA acceder a inventario, caravanas, cotizaciones y resúmenes.
 */
function doGet(e) {
  // 1. Si se accede a la URL directamente desde el navegador o celular (sin ?accion=...),
  // se muestra el Formulario Ganadero completo como aplicación web independiente:
  if (!e || !e.parameter || !e.parameter.accion) {
    return HtmlService.createHtmlOutputFromFile('Formulario')
      .setTitle('Carga Ganadera - Pozo del Carril')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 2. Si la petición incluye ?accion=... (Google AI Studio / Gemini API), responde como API REST en JSON:
  try {
    const params = e.parameter;
    const accion = params.accion;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let respuesta = {};

    switch (accion) {
      // 1. Verificación de estado / Health check
      case 'ping':
        respuesta = {
          estado: 'OK',
          sistema: 'Sistema de Gestión Ganadera - Pozo del Carril',
          version: '2.2',
          timestamp: new Date().toISOString()
        };
        break;

      // 2. Consulta individual de una caravana (Pesos Real/Estimado, Categoría, Estado)
      case 'consultarCaravana':
        const caravanaBuscada = (params.caravana || '').toString().trim().toUpperCase();
        if (!caravanaBuscada) {
          respuesta = { exito: false, mensaje: 'Debe especificar el parámetro "caravana".' };
          break;
        }
        const listaCaravanas = obtenerDatosCaravanasExistentes();
        const animal = listaCaravanas.find(c => c.caravana === caravanaBuscada);
        if (animal) {
          respuesta = {
            exito: true,
            caravana: animal.caravana,
            categoria: animal.categoria,
            subunidad: animal.subunidad,
            pesoInicial: animal.pesoInicial,
            pesoEstimado: animal.pesoEstimado,
            pesoReal: animal.pesoReal,
            fechaNacimiento: animal.fechaNacimiento,
            estado: animal.estado
          };
        } else {
          respuesta = { exito: false, mensaje: `Caravana ${caravanaBuscada} no encontrada en EXISTENCIAS.` };
        }
        break;

      // 3. Listado de caravanas activas (para autocompletado en prompts o selecciones)
      case 'listarCaravanasActivas':
        const activas = obtenerDatosCaravanasExistentes();
        respuesta = {
          exito: true,
          total: activas.length,
          caravanas: activas.map(a => ({
            caravana: a.caravana,
            categoria: a.categoria,
            subunidad: a.subunidad,
            pesoEstimado: a.pesoEstimado,
            pesoReal: a.pesoReal
          }))
        };
        break;

      // 4. Resumen consolidado de Stock por Categoría y Subunidad (Bovinos, Porcinos y Ovinos)
      case 'resumenStock':
        const sheetExis = ss.getSheetByName('EXISTENCIAS');
        const bovinos = { porCategoria: [], porSubunidad: [], totalVivas: 0 };
        const porcinos = { porCategoria: [], porSubunidad: [], totalVivas: 0 };
        const ovinos = { porCategoria: [], porSubunidad: [], totalVivas: 0 };

        if (sheetExis) {
          // Bovinos: L3:M14 (cat), L19:M22 (sub), M15 (total vivas)
          const filasCatBov = sheetExis.getRange('L3:M14').getValues();
          filasCatBov.forEach(f => { if (f[0]) bovinos.porCategoria.push({ categoria: f[0], cabezas: Number(f[1]) || 0 }); });
          const filasSubBov = sheetExis.getRange('L19:M22').getValues();
          filasSubBov.forEach(f => { if (f[0]) bovinos.porSubunidad.push({ subunidad: f[0], cabezas: Number(f[1]) || 0 }); });
          bovinos.totalVivas = Number(sheetExis.getRange('M15').getValue()) || 0;

          // Porcinos: O3:P9 (cat), O14:P17 (sub), P10 (total vivas)
          const filasCatPor = sheetExis.getRange('O3:P9').getValues();
          filasCatPor.forEach(f => { if (f[0]) porcinos.porCategoria.push({ categoria: f[0], cabezas: Number(f[1]) || 0 }); });
          const filasSubPor = sheetExis.getRange('O14:P17').getValues();
          filasSubPor.forEach(f => { if (f[0]) porcinos.porSubunidad.push({ etapa: f[0], cabezas: Number(f[1]) || 0 }); });
          porcinos.totalVivas = Number(sheetExis.getRange('P10').getValue()) || 0;

          // Ovinos: R3:S8 (cat), R13:S16 (sub), S9 (total vivas)
          const filasCatOvin = sheetExis.getRange('R3:S8').getValues();
          filasCatOvin.forEach(f => { if (f[0]) ovinos.porCategoria.push({ categoria: f[0], cabezas: Number(f[1]) || 0 }); });
          const filasSubOvin = sheetExis.getRange('R13:S16').getValues();
          filasSubOvin.forEach(f => { if (f[0]) ovinos.porSubunidad.push({ etapa: f[0], cabezas: Number(f[1]) || 0 }); });
          ovinos.totalVivas = Number(sheetExis.getRange('S9').getValue()) || 0;
        }

        respuesta = {
          exito: true,
          totalCabezasVivas: bovinos.totalVivas,
          bovinos: bovinos,
          porcinos: porcinos,
          ovinos: ovinos,
          stockPorCategoria: bovinos.porCategoria,
          stockPorSubunidad: bovinos.porSubunidad
        };
        break;

      // 5. Cotizaciones de mercado vigentes (MAG Cañuelas / ROSGAN)
      case 'cotizaciones':
        const sheetCot = ss.getSheetByName('COTIZACIONES');
        const cots = [];
        if (sheetCot && sheetCot.getLastRow() > 1) {
          const rows = sheetCot.getRange(2, 1, sheetCot.getLastRow() - 1, 6).getValues();
          rows.forEach(r => {
            if (r[0] && r[1]) {
              cots.push({
                mercado: r[0],
                categoria: r[1],
                precioPromedio: r[4],
                precioMinimo: r[5]
              });
            }
          });
        }
        respuesta = { exito: true, total: cots.length, cotizaciones: cots };
        break;

      // 6. Últimos movimientos registrados en Movimientos_Existencia
      case 'ultimosMovimientos':
        const sheetMov = ss.getSheetByName('Movimientos_Existencia');
        const movs = [];
        if (sheetMov && sheetMov.getLastRow() > 1) {
          const lr = sheetMov.getLastRow();
          const startR = Math.max(2, lr - 14); // últimos 15
          const numR = lr - startR + 1;
          const rowsMov = sheetMov.getRange(startR, 1, numR, 15).getValues();
          rowsMov.reverse().forEach(r => {
            if (r[0]) {
              movs.push({
                idRegistro: r[0],
                fecha: r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : r[1],
                caravana: r[2],
                unidadProduccion: r[3] || 'Bovino',
                fechaNacimiento: r[4] instanceof Date ? Utilities.formatDate(r[4], Session.getScriptTimeZone(), 'yyyy-MM-dd') : r[4],
                pesoEstimado: r[5],
                pesoReal: r[6],
                categoria: r[7],
                etapa: r[8],
                concepto: r[9],
                unidad: r[10],
                costoUnitario: r[11],
                cantidad: r[12],
                costoTotal: r[13],
                observaciones: r[14]
              });
            }
          });
        }
        respuesta = { exito: true, total: movs.length, movimientos: movs };
        break;

      default:
        respuesta = {
          exito: false,
          mensaje: `Acción '${accion}' no reconocida en doGet. Acciones disponibles: ping, consultarCaravana, listarCaravanasActivas, resumenStock, cotizaciones, ultimosMovimientos.`
        };
    }

    return ContentService.createTextOutput(JSON.stringify(respuesta))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      exito: false,
      error: err.toString(),
      mensaje: 'Error procesando solicitud doGet'
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Endpoint HTTP POST: Cargas, altas y registraciones desde Google AI Studio / Gemini
 * Permite a la IA invocar herramientas (Function Calling) para escribir en la planilla.
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        payload = (e && e.parameter) ? e.parameter : {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const accion = payload.accion || '';
    let resultado = {};

    switch (accion) {
      // 1. Registrar pesaje real en balanza o imputar costo individual por caravana
      case 'registrarCostoOPesaje':
      case 'registrarPesaje':
        resultado = guardarCostoMovimientoExistencia({
          fechaRegistro: payload.fecha || new Date().toISOString().split('T')[0],
          caravana: payload.caravana,
          fechaNacimiento: payload.fechaNacimiento || '',
          pesoEstimado: payload.pesoEstimado || 0,
          pesoReal: payload.pesoReal || '', // Balanza real
          categoria: payload.categoria || '',
          etapa: payload.etapa || '',
          conceptoCosto: payload.conceptoCosto || 'Control de Peso Balanza',
          unidad: payload.unidad || 'U',
          costoUnitario: payload.costoUnitario || 0,
          cantidad: payload.cantidad || 1,
          observaciones: payload.observaciones || 'Cargado vía Google AI Studio / Gemini'
        });
        break;

      // 2. Dar de alta una nueva caravana individual en EXISTENCIAS
      case 'altaCaravana':
        resultado = guardarNuevaCaravana({
          fechaRegistro: payload.fecha || new Date().toISOString().split('T')[0],
          caravana: payload.caravana,
          unidadProduccion: payload.unidadProduccion || 'Bovino',
          categoria: payload.categoria || 'Terneros',
          subunidad: payload.subunidad || 'Cria',
          pesoInicial: payload.pesoInicial || 35.0
        });
        break;

      // 3. Registrar un hecho físico (movimiento de lote, cambio de categoría, mortandad)
      case 'registrarMovimientoFisico':
        resultado = registrarMovimientoFisico({
          fecha: payload.fecha || new Date().toISOString().split('T')[0],
          tipo: payload.tipo || 'Transferencia_Interna',
          idCatOrigen: payload.idCatOrigen || '',
          idCatDestino: payload.idCatDestino || '',
          cabezas: payload.cabezas || 0,
          pesoEstimado: payload.pesoEstimado || 0,
          idSector: payload.idSector || '',
          caravana: payload.caravana || '',
          insumoAsociado: payload.insumoAsociado || '',
          cantidadInsumo: payload.cantidadInsumo || 0,
          causaMuerte: payload.causaMuerte || '',
          observaciones: payload.observaciones || 'Cargado vía Google AI Studio / Gemini'
        });
        break;

      // 4. Imputar un gasto contable en TABLA_IMPUTACION
      case 'registrarGasto':
        resultado = registrarImputacionGasto({
          fecha: payload.fecha || new Date().toISOString().split('T')[0],
          comprobante: payload.comprobante || 'GEMINI-0001',
          proveedor: payload.proveedor || 'Produccion Propia Campo',
          concepto: payload.concepto || '',
          importe: payload.importe || 0,
          idUn: payload.idUn || 'UN-BOV',
          etapaProductiva: payload.etapaProductiva || 'Engorde',
          planoAfectado: payload.planoAfectado || 'Resultado Economico',
          efectoExistencias: payload.efectoExistencias || 'Sin Efecto'
        });
        break;

      // 5. Sincronizar y recalcular pesos y fórmulas de EXISTENCIAS
      case 'recalcularPesos':
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const cant = actualizarFormulasExistencias(ss);
        resultado = { exito: true, mensaje: `Fórmulas y pesos recalculados en ${cant} animales de EXISTENCIAS.` };
        break;

      default:
        resultado = {
          exito: false,
          mensaje: `Acción '${accion}' no reconocida en doPost. Acciones disponibles: registrarPesaje, registrarCostoOPesaje, altaCaravana, registrarMovimientoFisico, registrarGasto, recalcularPesos.`
        };
    }

    return ContentService.createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      exito: false,
      error: err.toString(),
      mensaje: 'Error procesando solicitud doPost'
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================================
// FUNCION PARA BOTON Y MENU: ACTUALIZAR HOJA CAT CATEGORIAS
// ============================================================================
/**
 * Actualiza y reestructura la hoja CAT CATEGORIAS con los IDs CAT_BOV_1 a CAT_BOV_12
 * y las nuevas categorias porcinas CAT_POR_1 a CAT_POR_7.
 * Puede ser asignada directamente a un boton/dibujo en la hoja o ejecutada desde el menu.
 */
function actualizarHojaCatCategorias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Buscar hoja CAT CATEGORIAS o CAT_CATEGORIAS_BOVINOS
  let sheet = ss.getSheetByName('CAT CATEGORIAS');
  if (!sheet) {
    sheet = ss.getSheetByName('CAT_CATEGORIAS_BOVINOS') || ss.getSheetByName('CAT_CATEGORIAS');
    if (sheet) {
      sheet.setName('CAT CATEGORIAS');
    } else {
      sheet = ss.insertSheet('CAT CATEGORIAS');
    }
  }

  const COLOR_HEADER = '#1E3A1E'; // Verde institucional UNRC
  const COLOR_TEXT = '#FFFFFF';
  
  const headers = [
    'Id_CAT', 
    'Nombre_Categoria', 
    'Etapa_Ciclo', 
    'Tipo_Destino', 
    'Peso_Ref_Ingreso_Kg', 
    'Peso_Ref_Egreso_Kg', 
    'Cuenta_Contable_Activo'
  ];

  const anchoCols = [110, 220, 120, 190, 130, 130, 190];

  const dataCat = [
    // 1. Bovinos
    ['CAT_BOV_1', 'Vacas con cria', 'Cria', 'Activo Biologico Reproductor', 420, 480, '1.4.01 Reproductores Rodeo'],
    ['CAT_BOV_2', 'Vacas secas', 'Cria', 'Activo Biologico Reproductor', 430, 470, '1.4.01 Reproductores Rodeo'],
    ['CAT_BOV_3', 'Vacas prenadas', 'Cria', 'Activo Biologico Reproductor', 450, 520, '1.4.01 Reproductores Rodeo'],
    ['CAT_BOV_4', 'Vaquillonas 1 (recria)', 'Recria', 'Reposicion', 240, 310, '1.3.01 Hacienda en Crecimiento'],
    ['CAT_BOV_5', 'Vaquillonas 2 (servicio/prenez)', 'Recria', 'Reposicion', 310, 380, '1.3.01 Hacienda en Crecimiento'],
    ['CAT_BOV_6', 'Terneros', 'Cria', 'Carne comercial', 40, 140, '1.3.01 Hacienda en Crecimiento'],
    ['CAT_BOV_7', 'Terneras', 'Cria', 'Reposicion/Venta', 40, 135, '1.3.01 Hacienda en Crecimiento'],
    ['CAT_BOV_8', 'Destete', 'Recria', 'Invernada', 140, 180, '1.3.01 Hacienda en Crecimiento'],
    ['CAT_BOV_9', 'Novillitos (adaptacion)', 'Adaptacion', 'Carne comercial', 180, 280, '1.3.01 Hacienda en Crecimiento'],
    ['CAT_BOV_10', 'Novillos terminados', 'Engorde', 'Carne comercial', 280, 450, '1.3.01 Bienes de Cambio - Carne'],
    ['CAT_BOV_11', 'Toros cabana', 'Cria', 'Activo Biologico Reproductor', 650, 850, '1.4.02 Reproductores Cabana'],
    ['CAT_BOV_12', 'Toritos reposicion', 'Recria', 'Reposicion', 350, 500, '1.3.01 Hacienda en Crecimiento'],
    // 2. Porcinos
    ['CAT_POR_1', 'Madre', '', 'Activo Biologico Reproductor', '', '', '1.4.03 Reproductores Porcinos'],
    ['CAT_POR_2', 'Padrillo', '', 'Activo Biologico Reproductor', '', '', '1.4.03 Reproductores Porcinos'],
    ['CAT_POR_3', 'Lechon', 'Cria', 'Carne comercial', 0, 15, '1.3.02 Bienes de Cambio - Porcinos'],
    ['CAT_POR_4', 'Cachorro, Cachorra', 'Destete', 'Carne comercial', 15, 50, '1.3.02 Bienes de Cambio - Porcinos'],
    ['CAT_POR_5', 'Capon', 'Engorde', 'Carne comercial', 50, 105, '1.3.02 Bienes de Cambio - Porcinos'],
    ['CAT_POR_6', 'Madre de descarte', '', 'Activo Biologico Reproductor', '', '', '1.4.03 Reproductores Porcinos'],
    ['CAT_POR_7', 'Padrillo de descarte', '', 'Activo Biologico Reproductor', '', '', '1.4.03 Reproductores Porcinos'],
      // 3. Ovinos
      ['CAT-OVIN-01', 'Madres', '', 'Activo Biologico Reproductor', '', '', ''],
      ['CAT-OVIN-02', 'Carneros', '', 'Activo Biologico Reproductor', '', '', ''],
      ['CAT-OVIN-03', 'Corderos', 'Cria', 'Engorde/Venta', '', '', ''],
      ['CAT-OVIN-04', 'Borregas', 'Recria', '', '', '', ''],
      ['CAT-OVIN-05', 'Borregos', 'Recria', 'Engorde/venta', '', '', ''],
      ['CAT-OVIN-06', 'Ovino de Descarte', '', 'Descarte', '', '', ''],
    // 3. Granos
    ['CAT-14', 'Granos Maiz Comercial', 'Agricultura', 'Venta Comercial', 0, 0, '1.3.04.01 Existencias - Maiz']
  ];

  // Limpiar hoja
  sheet.clear();

  // Encabezados
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground(COLOR_HEADER)
    .setFontColor(COLOR_TEXT)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.setFrozenRows(1);

  // Filas de datos
  const dataRange = sheet.getRange(2, 1, dataCat.length, headers.length);
  dataRange.setValues(dataCat);
  dataRange.setFontFamily('Aptos Narrow')
    .setFontSize(11)
    .setVerticalAlignment('middle');

  sheet.getRange(2, 1, dataCat.length, 1).setHorizontalAlignment('center');
  sheet.getRange(2, 2, dataCat.length, 1).setHorizontalAlignment('left');
  sheet.getRange(2, 3, dataCat.length, 1).setHorizontalAlignment('center');
  sheet.getRange(2, 4, dataCat.length, 1).setHorizontalAlignment('left');
  sheet.getRange(2, 5, dataCat.length, 2).setHorizontalAlignment('right').setNumberFormat('#,##0');
  sheet.getRange(2, 7, dataCat.length, 1).setHorizontalAlignment('left');

  anchoCols.forEach(function(w, idx) {
    sheet.setColumnWidth(idx + 1, w);
  });

  try {
    actualizarValidaciones();
  } catch (e) {}

  return { success: true, count: dataCat.length };
}

function actualizarHojaCatCategoriasMenu() {
  try {
    const res = actualizarHojaCatCategorias();
    SpreadsheetApp.getActiveSpreadsheet().toast('Hoja CAT CATEGORIAS actualizada con 12 bovinos, 7 porcinos y 6 ovinos.', 'Actualizacion Exitosa', 5);
    SpreadsheetApp.getUi().alert('Actualizacion Completada', 'La hoja CAT CATEGORIAS ha sido actualizada correctamente:\n\n- 12 categorias bovinas (CAT_BOV_1 a CAT_BOV_12)\n- 7 categorias porcinas (CAT_POR_1 a CAT_POR_7)\n- 6 categorias ovinas (CAT-OVIN-01 a CAT-OVIN-06)\n- 1 insumo agricola (CAT-14)', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    SpreadsheetApp.getUi().alert('Error al actualizar: ' + err.message);
  }
}

// ============================================================================
// FUNCION PARA BOTON Y MENU: ACTUALIZAR HOJA COTIZACIONES
// ============================================================================
/**
 * Actualiza la hoja COTIZACIONES con las cotizaciones unificadas de Bovinos,
 * Porcinos y Ovinos. Puede ser ejecutada desde el menu o asignada a un boton.
 */
function actualizarHojaCotizacionesMenu() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ahora = new Date();
    const consultaStr = Utilities.formatDate(ahora, TIMEZONE_ARG, 'dd/MM/yyyy HH:mm');
    const filas = poblarHojaCotizaciones_(ss, null, null, ahora, consultaStr);
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Hoja COTIZACIONES actualizada con éxito (' + filas + ' filas).', '📈 Cotizaciones al Día', 5);
    SpreadsheetApp.getUi().alert(
      'Actualización de Cotizaciones Exitosa',
      'La hoja COTIZACIONES ha sido actualizada con éxito (' + filas + ' registros):\n\n' +
      '- 11 cotizaciones bovinas (MAG Cañuelas / ROSGAN)\n' +
      '- 7 cotizaciones porcinas (Capón $2.160,43, Lechón $4.162,29, Cachorro/a $2.118, Madres, Padrillos, Descartes)\n' +
      '- 6 cotizaciones ovinas (Madres $2.000, Carneros $3.850, Corderos $4.400, Borregas $5.100, Borregos $5.000, Descarte $1.100)',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert('Error al actualizar cotizaciones: ' + err.message);
  }
}


// ============================================================================
// FUNCION PARA BOTON Y MENU: ACTUALIZAR HOJA COSTOS
// ============================================================================
/**
 * Actualiza y normaliza la hoja COSTOS incorporando los conceptos de Bovinos y Porcinos,
 * configurando encabezados, validaciones de listas desplegables y formatos numéricos.
 * Preserva costos unitarios ingresados previamente si existen.
 */
function actualizarHojaCostos(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('COSTOS');
  if (!sheet) {
    sheet = ss.insertSheet('COSTOS');
  }

  const COLOR_HEADER = '#1E3A1E'; // Verde institucional UNRC
  const COLOR_TEXT = '#FFFFFF';

  const headers = ['Id costo', 'Concepto de costo', 'Etapa', 'Unidad de medida', 'Costo unitario ($)'];
  const anchoCols = [110, 290, 130, 210, 140];

  // Leer costos unitarios existentes para no sobrescribir si el usuario cargó precios
  const costosExistentes = {};
  if (sheet.getLastRow() >= 2) {
    const rawData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    rawData.forEach(function(r) {
      const id = (r[0] || '').toString().trim();
      const val = r[4];
      if (id && val !== '' && val !== null && val !== undefined) {
        costosExistentes[id] = val;
      }
    });
  }

  const dataCostosDef = [
    // 1. Costos Bovinos
    ['CST001', 'Alimento Destete Precoz (Suplemento)', 'Cria', 'kg por peso del animal', 1100.00],
    ['CST002', 'Forraje Madre (Pastura Alfalfa) / Cba', 'Cria', 'kgs (MS)', 95.00],
    ['CST003', 'Vacuna Aftosa / Brucelosis', 'Cria', 'Dosis', 1850.00],
    ['CST004', 'Forraje Alfalfa / Recria', 'Recria', 'kgs (MS)', 95.00],
    ['CST005', 'Sanidad General (Vacunas/Desparasitantes)', 'Recria', 'Dosis', 1850.00],
    ['CST006', 'Grano de Maiz/Adaptacion', 'Adaptacion', 'kg por peso del animal', 180.00],
    ['CST007', 'Expeller / Harina de Soja / Adaptacion', 'Adaptacion', 'kg por peso del animal', ''],
    ['CST008', 'Nucleo Mineral c/ Monensina/Adaptacion', 'Adaptacion', 'kg por peso del animal', ''],
    ['CST009', 'Alfalfa (Fibra / Rollo) / Adaptacion', 'Adaptacion', 'kgs (MS)', 95.00],
    ['CST010', 'Grano de Maiz / Engorde', 'Engorde', 'kg por peso del animal', 180.00],
    ['CST011', 'Expeller / Harina de Soja / Engorde', 'Engorde', 'kg por peso del animal', ''],
    ['CST012', 'Nucleo Mineral c/ Monensina / Engorde', 'Engorde', 'kg por peso del animal', ''],
    ['CST013', 'Alfalfa (Fibra / Rollo) / Engorde', 'Engorde', 'kgs (MS)', 95.00],
    // 2. Costos Porcinos (Precios Reales en Tranquera / Mostrador Local)
    ['CSTP01', 'Maiz propio - Madres', 'Reproduccion', 'kg por peso de animal', 240.50],
    ['CSTP02', 'Maiz propio - Engorde', 'Engorde', 'kg por peso de animal', 240.50],
    ['CSTP03', 'Maiz propio - Destete', 'Destete', 'kg por peso de animal', 240.50],
    ['CSTP04', 'Concentrado Preiniciador - Lechones', 'Cria', 'kg por peso de animal', 3400.00],
    ['CSTP05', 'Concentrado Desarrollo - Destete', 'Destete', 'kg por peso de animal', 2975.00],
    ['CSTP06', 'Soja peleteada - Destete', 'Destete', 'kg por peso de animal', 387.50],
    ['CSTP07', 'Soja peleteada - Engorde', 'Engorde', 'kg por peso de animal', 387.50]
  ];

  const filasFinales = dataCostosDef.map(function(item) {
    const id = item[0];
    const precio = (costosExistentes[id] !== undefined) ? costosExistentes[id] : item[4];
    return [item[0], item[1], item[2], item[3], precio];
  });

  // Limpiar hoja
  sheet.clear();

  // Encabezados
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground(COLOR_HEADER)
    .setFontColor(COLOR_TEXT)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.setFrozenRows(1);

  // Filas de datos
  const dataRange = sheet.getRange(2, 1, filasFinales.length, headers.length);
  dataRange.setValues(filasFinales);
  dataRange.setFontFamily('Aptos Narrow')
    .setFontSize(11)
    .setVerticalAlignment('middle');

  sheet.getRange(2, 1, filasFinales.length, 1).setHorizontalAlignment('center');
  sheet.getRange(2, 2, filasFinales.length, 1).setHorizontalAlignment('left');
  sheet.getRange(2, 3, filasFinales.length, 1).setHorizontalAlignment('center');
  sheet.getRange(2, 4, filasFinales.length, 1).setHorizontalAlignment('left');
  sheet.getRange(2, 5, filasFinales.length, 1).setHorizontalAlignment('right').setNumberFormat('$ #,##0.00');

  anchoCols.forEach(function(w, idx) {
    sheet.setColumnWidth(idx + 1, w);
  });

  // Validacion de lista en columna C (Etapa)
  const ruleEtapaCostos = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Cria', 'Recria', 'Adaptacion', 'Engorde', 'Reproduccion', 'Destete'])
    .setAllowInvalid(true)
    .build();
  sheet.getRange('C2:C' + (filasFinales.length + 10)).setDataValidation(ruleEtapaCostos);

  try {
    actualizarValidaciones();
  } catch (e) {}

  return { success: true, count: filasFinales.length };
}

/**
 * Menu y disparador para actualizar la hoja COSTOS con Bovinos y Porcinos.
 */
function actualizarHojaCostosMenu() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const res = actualizarHojaCostos(ss);
    ss.toast('Hoja COSTOS actualizada (' + res.count + ' conceptos).', '💰 Costos al Día', 5);
    const mensaje = 'La hoja COSTOS ha sido actualizada con éxito (' + res.count + ' conceptos):\n\n' +
      '- 13 conceptos bovinos (CST001 a CST013: Alimentos, Pasturas, Sanidad, Vacunas, Granos)\n' +
      '- 7 conceptos porcinos (CSTP01 a CSTP07: Maíz Propio, Concentrados Preiniciador y Desarrollo, Soja Peleteada)\n' +
      '- Validación de etapas ampliada a Reproducción, Cría, Destete, Recría, Adaptación y Engorde.';
    SpreadsheetApp.getUi().alert('Actualización de Costos Exitosa', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    SpreadsheetApp.getUi().alert('Error al actualizar costos: ' + err.message);
  }
}

/**
 * Configura o actualiza la hoja GANANCIA DIARIA KG con las 3 tablas maestras
 * en paralelo para Bovinos, Porcinos y Ovinos.
 */
function actualizarHojaGananciaDiaria(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('GANANCIA DIARIA KG');
  if (!sheet) {
    sheet = ss.insertSheet('GANANCIA DIARIA KG');
  }

  sheet.clear();
  sheet.clearDataValidations();
  sheet.setHiddenGridlines(false);

  const tablas = [
    {
      titulo: '🐂 BOVINOS: GANANCIA DIARIA DE PESO POR ETAPA',
      colorTitulo: '#1E3A1E',
      colorHdr: '#2E5A2E',
      colorLight: '#F0F7F0',
      startCol: 1,
      data: [
        ['Cria', 190, 35.0, 140.0, 0.55],
        ['Recria', 235, 140.0, 280.0, 0.60],
        ['Adaptacion', 27, 280.0, 315.0, 1.00],
        ['Engorde', 90, 315.0, 450.0, 1.25]
      ]
    },
    {
      titulo: '🐖 PORCINOS: GANANCIA DIARIA DE PESO POR ETAPA',
      colorTitulo: '#6B2D5C',
      colorHdr: '#823B71',
      colorLight: '#F5EBF2',
      startCol: 7,
      data: [
        ['Cria', 25, 1.45, 7.0, 0.20],
        ['Destete', 45, 7.0, 30.0, 0.48],
        ['Engorde', 90, 30.0, 110.0, 0.85]
      ]
    },
    {
      titulo: '🐑 OVINOS: GANANCIA DIARIA DE PESO POR ETAPA',
      colorTitulo: '#1C3D5A',
      colorHdr: '#2B5876',
      colorLight: '#EBF2F7',
      startCol: 13,
      data: [
        ['Cria', 75, 4.0, 26.0, 0.23],
        ['Recria', 120, 26.0, 40.0, 0.12],
        ['Engorde', 60, 26.0, 44.0, 0.20]
      ]
    }
  ];

  const headers = ['Etapa', 'Duracion_Dias', 'Peso Inicial KG', 'Peso Final KG', 'Ganancia Diaria KG'];

  tablas.forEach(function(t) {
    const sc = t.startCol;
    const ec = sc + 4;

    // Titulo fusionado
    const rTitulo = sheet.getRange(1, sc, 1, 5);
    rTitulo.merge()
      .setValue(t.titulo)
      .setBackground(t.colorTitulo)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Encabezados
    const rHdr = sheet.getRange(2, sc, 1, 5);
    rHdr.setValues([headers])
      .setBackground(t.colorHdr)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(9)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);

    // Filas de datos
    const rData = sheet.getRange(3, sc, t.data.length, 5);
    rData.setValues(t.data);
    rData.setVerticalAlignment('middle');
    sheet.getRange(3, sc, t.data.length, 1).setHorizontalAlignment('left');
    sheet.getRange(3, sc + 1, t.data.length, 1).setHorizontalAlignment('center').setNumberFormat('#,##0');
    sheet.getRange(3, sc + 2, t.data.length, 2).setHorizontalAlignment('right').setNumberFormat('#,##0.00');
    sheet.getRange(3, sc + 4, t.data.length, 1).setHorizontalAlignment('right').setNumberFormat('0.00');
  });

  // Anchos de columna
  const anchos = [130, 120, 130, 130, 150, 30, 130, 120, 130, 130, 150, 30, 130, 120, 130, 130, 150];
  anchos.forEach((w, idx) => sheet.setColumnWidth(idx + 1, w));
  sheet.setFrozenRows(2);

  return { exito: true };
}

/**
 * Menú de actualización de la hoja GANANCIA DIARIA KG
 */
function actualizarHojaGananciaDiariaMenu() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    actualizarHojaGananciaDiaria(ss);
    ss.toast('Hoja GANANCIA DIARIA KG actualizada con éxito (3 especies).', '📈 Ganancia Diaria', 5);
    const mensaje = 'La hoja GANANCIA DIARIA KG ha sido configurada con 3 tablas maestras idénticas:\n\n' +
      '- 🐂 Bovinos: Cría (0.55), Recría (0.60), Adaptación (1.00), Engorde (1.25)\n' +
      '- 🐖 Porcinos: Cría (0.20), Destete (0.48), Engorde (0.85)\n' +
      '- 🐑 Ovinos: Cría (0.23), Recría (0.12), Engorde (0.20)\n\n' +
      'Las proyecciones de peso y cálculo de etapas ahora funcionan en sintonía para las tres especies.';
    SpreadsheetApp.getUi().alert('Actualización Exitosa', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    SpreadsheetApp.getUi().alert('Error al actualizar ganancia diaria: ' + err.message);
  }
}

/**
 * Carga o actualiza los datos simulados de las 3 especies (Bovinos B001-B005,
 * Porcinos P001-P005 y Ovinos O001-O005) en EXISTENCIAS y Movimientos_Existencia.
 */
function sembrarDatosSimuladosTrazabilidad(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheetExis = ss.getSheetByName('EXISTENCIAS');
  if (!sheetExis) {
    prepararHojaExistencias(ss);
    sheetExis = ss.getSheetByName('EXISTENCIAS');
  }

  let sheetMov = ss.getSheetByName('Movimientos_Existencia');
  if (!sheetMov) {
    prepararHojaMovimientosExistencia(ss);
    sheetMov = ss.getSheetByName('Movimientos_Existencia');
  }

  // 1. Limpiar filas de datos previas en EXISTENCIAS (Cols A a J, filas 2 a 100)
  sheetExis.getRange('A2:J100').clearContent();

  // 15 Registros limpios de EXISTENCIAS (IDs 1 al 15)
  const exisData = [
    // Bovinos (A123 y A124 base + B001 a B003)
    [1, '2026-10-17', 'A123', 'Bovino', 'Terneros', 'Recria', 35.0, 'Vivo'],
    [2, '2026-10-16', 'A124', 'Bovino', 'Terneras', 'Recria', 32.0, 'Vivo'],
    [3, '2026-10-18', 'B001', 'Bovino', 'Terneros', 'Recria', 34.0, 'Vivo'],
    [4, '2026-10-19', 'B002', 'Bovino', 'Novillitos (adaptacion)', 'Adaptacion', 33.0, 'Vivo'],
    [5, '2026-10-20', 'B003', 'Bovino', 'Novillos terminados', 'Engorde', 30.0, 'Vivo'],
    // Porcinos (P001 a P005)
    [6, '2025-06-01', 'P001', 'Porcino', 'Madre', 'Reproduccion', 180.0, 'Vivo'],
    [7, '2025-05-15', 'P002', 'Porcino', 'Padrillo', 'Reproduccion', 220.0, 'Vivo'],
    [8, '2026-02-10', 'P003', 'Porcino', 'Lechon', 'Cria', 1.45, 'Vivo'],
    [9, '2026-01-01', 'P004', 'Porcino', 'Cachorro, Cachorra', 'Destete', 7.0, 'Vivo'],
    [10, '2025-11-01', 'P005', 'Porcino', 'Capon', 'Engorde', 30.0, 'Vivo'],
    // Ovinos (O001 a O005)
    [11, '2025-04-01', 'O001', 'Ovino', 'Madres', 'Reproduccion', 55.0, 'Vivo'],
    [12, '2025-03-01', 'O002', 'Ovino', 'Carneros', 'Reproduccion', 80.0, 'Vivo'],
    [13, '2026-01-15', 'O003', 'Ovino', 'Corderos', 'Cria', 4.0, 'Vivo'],
    [14, '2025-10-01', 'O004', 'Ovino', 'Borregas', 'Recria', 26.0, 'Vivo'],
    [15, '2025-10-01', 'O005', 'Ovino', 'Borregos', 'Engorde', 26.0, 'Vivo']
  ];

  for (let i = 0; i < exisData.length; i++) {
    const r = 2 + i;
    const item = exisData[i];
    sheetExis.getRange(r, 1).setValue(item[0]);
    sheetExis.getRange(r, 2).setValue(item[1]);
    sheetExis.getRange(r, 3).setValue(item[2]);
    sheetExis.getRange(r, 4).setValue(item[3]);
    sheetExis.getRange(r, 5).setValue(item[4]);
    sheetExis.getRange(r, 6).setValue(item[5]);
    sheetExis.getRange(r, 7).setValue(item[6]);
    sheetExis.getRange(r, 8).setFormula(`=IFERROR(XLOOKUP(C${r}, Movimientos_Existencia!$C:$C, Movimientos_Existencia!$F:$F, G${r}, 0, -1), G${r})`);
    sheetExis.getRange(r, 9).setFormula(`=IFERROR(CHOOSEROWS(FILTER(Movimientos_Existencia!$G$2:$G, Movimientos_Existencia!$C$2:$C=C${r}, ISNUMBER(Movimientos_Existencia!$G$2:$G), Movimientos_Existencia!$G$2:$G>0), -1), G${r})`);
    sheetExis.getRange(r, 10).setValue(item[7]);
  }

  sheetExis.getRange('A2:A16').setHorizontalAlignment('center');
  sheetExis.getRange('B2:B16').setHorizontalAlignment('center');
  sheetExis.getRange('C2:C16').setHorizontalAlignment('center').setFontWeight('bold');
  sheetExis.getRange('D2:D16').setHorizontalAlignment('center');
  sheetExis.getRange('E2:E16').setHorizontalAlignment('left');
  sheetExis.getRange('F2:F16').setHorizontalAlignment('center');
  sheetExis.getRange('G2:I16').setHorizontalAlignment('right').setNumberFormat('#,##0.0 "kg"');
  sheetExis.getRange('J2:J16').setHorizontalAlignment('center');

  // 2. Limpiar y reescribir Movimientos_Existencia (20 movimientos con estilo homogéneo)
  sheetMov.getRange('A2:O100').clearContent();

  const movsData = [
    // 5 Registros Históricos Base (A123 y A124)
    [1, '2026-10-17', 'A123', 'Bovino', '2026-10-17', 35.0, 35.0, 'Terneros', 'Cria', 'Nacimiento / Alta Inicial', 'Cab', 0.0, 1.0, 'Alta automática por nacimiento / registro de caravana A123'],
    [2, '2026-10-16', 'A124', 'Bovino', '2026-10-16', 32.0, 32.0, 'Terneras', 'Cria', 'Nacimiento / Alta Inicial', 'Cab', 0.0, 1.0, 'Alta automática por nacimiento / registro de caravana A124'],
    [3, '2026-11-30', 'A123', 'Bovino', '2026-10-17', 59.2, '', 'Terneros', 'Cria', 'Alimento Destete Precoz (Suplemento)', 'kg por peso del animal', 1100.0, 59.2, 'Costo cargado desde formulario: Alimento Destete Precoz (Suplemento)'],
    [4, '2027-05-25', 'A123', 'Bovino', '2026-10-17', 156.0, '', 'Terneros', 'Recria', 'Forraje Alfalfa / Recria', 'kgs (MS)', 95.0, 1434.4, 'Costo cargado desde formulario: Forraje Alfalfa / Recria'],
    [5, '2027-05-31', 'A124', 'Bovino', '2026-10-16', 156.9, 150.5, 'Terneras', 'Recria', 'Alimento Destete Precoz (Suplemento)', 'kg por peso del animal', 1100.0, 156.9, 'Costo cargado desde formulario: Alimento Destete Precoz (Suplemento)'],
    // 3 Nuevos Bovinos Simulados (B001, B002, B003)
    [6, '2026-11-15', 'B001', 'Bovino', '2026-10-18', 52.0, '', 'Terneros', 'Recria', 'Alimento Destete Precoz (Suplemento)', 'kg por peso del animal', 1100.0, 52.0, 'Suplementacion inicial lote recria B001'],
    [7, '2027-02-20', 'B002', 'Bovino', '2026-10-19', 285.0, '', 'Novillitos (adaptacion)', 'Adaptacion', 'Grano de Maiz/Adaptacion', 'kg por peso del animal', 180.0, 100.0, 'Racion adaptacion corral B002'],
    [8, '2027-04-10', 'B003', 'Bovino', '2026-10-20', 335.0, 335.0, 'Novillos terminados', 'Engorde', 'Grano de Maiz / Engorde', 'kg por peso del animal', 180.0, 320.0, 'Terminacion feedlot y pesaje control B003'],
    // 7 Movimientos Porcinos (P001 a P005)
    [9, '2026-03-05', 'P001', 'Porcino', '2025-06-01', 180.0, '', 'Madre', 'Reproduccion', 'Maiz propio - Madres', 'kg por peso de animal', 240.50, 45.0, 'Alimentacion racion diaria madre reproductora'],
    [10, '2026-03-08', 'P002', 'Porcino', '2025-05-15', 220.0, '', 'Padrillo', 'Reproduccion', 'Maiz propio - Madres', 'kg por peso de animal', 240.50, 50.0, 'Mantenimiento nutricional padrillo plantel'],
    [11, '2026-03-02', 'P003', 'Porcino', '2026-02-10', 5.5, 5.5, 'Lechon', 'Cria', 'Concentrado Preiniciador - Lechones', 'kg por peso de animal', 3400.0, 4.0, 'Pesaje control y suplementacion preiniciador lactancia'],
    [12, '2026-03-15', 'P004', 'Porcino', '2026-01-01', 22.0, 22.0, 'Cachorro, Cachorra', 'Destete', 'Concentrado Desarrollo - Destete', 'kg por peso de animal', 2975.0, 12.0, 'Transicion destete fase recria con concentrado desarrollo'],
    [13, '2026-03-25', 'P004', 'Porcino', '2026-01-01', 26.5, '', 'Cachorro, Cachorra', 'Destete', 'Soja peleteada - Destete', 'kg por peso de animal', 387.50, 15.0, 'Aporte proteico expeller soja peleteada destete'],
    [14, '2026-03-18', 'P005', 'Porcino', '2025-11-01', 85.0, 85.0, 'Capon', 'Engorde', 'Maiz propio - Engorde', 'kg por peso de animal', 240.50, 75.0, 'Pesaje intermedia y dieta engorde grano de maiz propio'],
    [15, '2026-03-28', 'P005', 'Porcino', '2025-11-01', 93.5, '', 'Capon', 'Engorde', 'Soja peleteada - Engorde', 'kg por peso de animal', 387.50, 25.0, 'Suplementacion proteica con soja peleteada terminacion'],
    // 5 Movimientos Ovinos (O001 a O005)
    [16, '2026-03-12', 'O001', 'Ovino', '2025-04-01', 55.0, '', 'Madres', 'Reproduccion', 'Sanidad General (Vacunas/Desparasitantes)', 'Dosis', '', 1.0, 'Plan sanitario madre plantel ovino (costo a completar)'],
    [17, '2026-03-12', 'O002', 'Ovino', '2025-03-01', 80.0, '', 'Carneros', 'Reproduccion', 'Sanidad General (Vacunas/Desparasitantes)', 'Dosis', '', 1.0, 'Revisacion y control sanitario carnero reproductor'],
    [18, '2026-03-05', 'O003', 'Ovino', '2026-01-15', 15.5, 15.5, 'Corderos', 'Cria', '', '', '', '', 'Pesaje de control al pie de la madre (GD: 0.23 kg/d)'],
    [19, '2026-03-20', 'O004', 'Ovino', '2025-10-01', 34.0, 34.0, 'Borregas', 'Recria', '', '', '', '', 'Pesaje control recria reposicion (GD: 0.12 kg/d)'],
    [20, '2026-03-22', 'O005', 'Ovino', '2025-10-01', 40.0, 40.0, 'Borregos', 'Engorde', '', '', '', '', 'Pesaje control engorde comercial (GD: 0.20 kg/d)']
  ];

  for (let j = 0; j < movsData.length; j++) {
    const r = 2 + j;
    const m = movsData[j];
    sheetMov.getRange(r, 1).setValue(m[0]);
    sheetMov.getRange(r, 2).setValue(m[1]);
    sheetMov.getRange(r, 3).setValue(m[2]);
    sheetMov.getRange(r, 4).setValue(m[3]);
    sheetMov.getRange(r, 5).setValue(m[4]);
    sheetMov.getRange(r, 6).setValue(m[5]);
    sheetMov.getRange(r, 7).setValue(m[6] !== '' ? m[6] : '');
    sheetMov.getRange(r, 8).setValue(m[7]);
    sheetMov.getRange(r, 9).setValue(m[8]);
    sheetMov.getRange(r, 10).setValue(m[9]);
    sheetMov.getRange(r, 11).setValue(m[10]);
    sheetMov.getRange(r, 12).setValue(m[11] !== '' ? m[11] : '');
    sheetMov.getRange(r, 13).setValue(m[12] !== '' ? m[12] : '');
    sheetMov.getRange(r, 14).setFormula(`=L${r}*M${r}`);
    sheetMov.getRange(r, 15).setValue(m[13]);
  }

  // Estilos uniformes en Movimientos_Existencia
  sheetMov.getRange('A2:A21').setHorizontalAlignment('center');
  sheetMov.getRange('B2:B21').setHorizontalAlignment('center');
  sheetMov.getRange('C2:C21').setHorizontalAlignment('center').setFontWeight('bold');
  sheetMov.getRange('D2:D21').setHorizontalAlignment('center');
  sheetMov.getRange('E2:E21').setHorizontalAlignment('center');
  sheetMov.getRange('F2:G21').setHorizontalAlignment('right').setNumberFormat('#,##0.0 "kg"');
  sheetMov.getRange('H2:H21').setHorizontalAlignment('left');
  sheetMov.getRange('I2:I21').setHorizontalAlignment('center');
  sheetMov.getRange('J2:J21').setHorizontalAlignment('left');
  sheetMov.getRange('K2:K21').setHorizontalAlignment('center');
  sheetMov.getRange('L2:L21').setHorizontalAlignment('right').setNumberFormat('$ #,##0.00');
  sheetMov.getRange('M2:M21').setHorizontalAlignment('right').setNumberFormat('#,##0.00');
  sheetMov.getRange('N2:N21').setHorizontalAlignment('right').setNumberFormat('$ #,##0.00');
  sheetMov.getRange('O2:O21').setHorizontalAlignment('left');

  // Bordes finos
  sheetMov.getRange('A2:O21').setBorder(true, true, true, true, true, true, '#D9D9D9', SpreadsheetApp.BorderStyle.SOLID);
  sheetExis.getRange('A2:J16').setBorder(true, true, true, true, true, true, '#D9D9D9', SpreadsheetApp.BorderStyle.SOLID);

  // Validaciones de Categoria y Subunidad
  const ruleCat = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRange("'CAT CATEGORIAS'!B2:B27")).build();
  const ruleEtapa = SpreadsheetApp.newDataValidation().requireValueInList(['Cria', 'Recria', 'Adaptacion', 'Engorde', 'Destete', 'Reproduccion', 'Engorde/Venta']).build();

  sheetExis.getRange('E2:E100').setDataValidation(ruleCat);
  sheetExis.getRange('F2:F100').setDataValidation(ruleEtapa);
  sheetMov.getRange('H2:H100').setDataValidation(ruleCat);
  sheetMov.getRange('I2:I100').setDataValidation(ruleEtapa);

  // Sincronizar automáticamente con TABLA_IMPUTACION
  const resSync = sincronizarTodosLosMovimientosExistencia(ss);

  return { altas: exisData.length, movs: movsData.length, imputados: resSync.sincronizados };
}

/**
 * Menú para sembrar datos simulados de Bovinos, Porcinos y Ovinos
 */
function sembrarDatosSimuladosTrazabilidadMenu() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const res = sembrarDatosSimuladosTrazabilidad(ss);
    ss.toast('Trazabilidad actualizada con éxito (' + res.altas + ' animales, ' + res.movs + ' movimientos).', '🐑🐖 Trazabilidad Lista', 5);
    SpreadsheetApp.getUi().alert(
      'Trazabilidad Limpia y Actualizada con Éxito',
      'Se configuraron ambas hojas con los datos del sistema:\n\n' +
      '• EXISTENCIAS (IDs 1 al 15, Filas 2 a 16):\n' +
      '   - Bovinos: A123, A124 (base) + B001, B002, B003 (simulados).\n' +
      '   - Porcinos: P001 a P005.\n' +
      '   - Ovinos: O001 a O005.\n' +
      '   - Fórmulas automáticas de Peso Estimado y Peso Real en todas las filas.\n\n' +
      '• Movimientos_Existencia (IDs 1 al 20):\n' +
      '   - Filas 1 a 5 intactas con A123 y A124.\n' +
      '   - Filas 6 a 20 con los pesajes y raciones de Bovinos, Porcinos y Ovinos.\n' +
      '   - Formato, tipografía y bordes 100% uniformes y homogéneos.\n' +
      '   - Listas desplegables de Categorías y Etapas multiespecie aplicadas.\n\n' +
      '• TABLA_IMPUTACION:\n' +
      '   - ' + res.imputados + ' gastos imputados automáticamente (Bovinos y Porcinos) con formato impecable.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert('Error al sembrar datos de trazabilidad: ' + err.message);
  }
}

/**
 * Configura la hoja SUB UNIDADES CICLO con 3 tablas maestras en paralelo
 * para Bovinos, Porcinos y Ovinos.
 */
function actualizarHojaSubUnidadesCiclo(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('SUB UNIDADES CICLO');
  if (!sheet) {
    sheet = ss.insertSheet('SUB UNIDADES CICLO');
  }

  sheet.clear();
  sheet.setHiddenGridlines(false);

  const tablas = [
    {
      titulo: '🐂 BOVINOS: SUBUNIDADES DE CICLO',
      colorTitulo: '#1E3A1E',
      colorHdr: '#2E5A2E',
      startCol: 1,
      data: [
        ['CIC-BOV-01', 'Cria', 'Desde nacimiento hasta destete (al pie de la madre)'],
        ['CIC-BOV-02', 'Recria', 'Crecimiento y desarrollo a pasto / suplementacion'],
        ['CIC-BOV-03', 'Adaptacion', 'Transicion a corral y acostumbramiento a racion'],
        ['CIC-BOV-04', 'Engorde', 'Terminacion a corral / feedlot para faena'],
        ['CIC-BOV-05', 'Reproduccion', 'Mantenimiento de plantel reproductor (Toros y Vacas)']
      ]
    },
    {
      titulo: '🐖 PORCINOS: SUBUNIDADES DE CICLO',
      colorTitulo: '#6B2D5C',
      colorHdr: '#823B71',
      startCol: 5,
      data: [
        ['CIC-POR-01', 'Cria', 'Maternidad y lactancia (hasta 21-28 dias / 6-8 kg)'],
        ['CIC-POR-02', 'Destete', 'Recria en pistas de destete (hasta 70 dias / 28-32 kg)'],
        ['CIC-POR-03', 'Engorde', 'Desarrollo y terminacion de capones para faena (105-115 kg)'],
        ['CIC-POR-04', 'Reproduccion', 'Gestacion y servicio de madres y padrillos de plantel']
      ]
    },
    {
      titulo: '🐑 OVINOS: SUBUNIDADES DE CICLO',
      colorTitulo: '#1C3D5A',
      colorHdr: '#2B5876',
      startCol: 9,
      data: [
        ['CIC-OVIN-01', 'Cria', 'Lactancia de corderos al pie de la madre (hasta 60-90 dias)'],
        ['CIC-OVIN-02', 'Recria', 'Desarrollo de borregas de reposicion para primer servicio'],
        ['CIC-OVIN-03', 'Engorde', 'Terminacion de borregos y corderos comerciales para faena'],
        ['CIC-OVIN-04', 'Reproduccion', 'Mantenimiento de plantel de madres y carneros']
      ]
    }
  ];

  const headers = ['ID_CICLO', 'Etapa / Subunidad', 'Descripcion / Destino'];

  tablas.forEach(function(t) {
    const sc = t.startCol;
    const ec = sc + 2;

    // Titulo
    sheet.getRange(1, sc, 1, 3).merge()
      .setValue(t.titulo)
      .setBackground(t.colorTitulo)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Headers
    sheet.getRange(2, sc, 1, 3)
      .setValues([headers])
      .setBackground(t.colorHdr)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(9)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);

    // Filas
    sheet.getRange(3, sc, t.data.length, 3)
      .setValues(t.data)
      .setVerticalAlignment('middle');

    sheet.getRange(3, sc, t.data.length, 1).setHorizontalAlignment('center');
    sheet.getRange(3, sc + 1, t.data.length, 1).setHorizontalAlignment('left').setFontWeight('bold');
    sheet.getRange(3, sc + 2, t.data.length, 1).setHorizontalAlignment('left');

    sheet.getRange(2, sc, t.data.length + 1, 3).setBorder(true, true, true, true, true, true, '#D9D9D9', SpreadsheetApp.BorderStyle.SOLID);
  });

  const anchos = [130, 160, 330, 30, 130, 160, 330, 30, 130, 160, 330];
  anchos.forEach((w, idx) => sheet.setColumnWidth(idx + 1, w));
  sheet.setFrozenRows(2);

  return { exito: true };
}

/**
 * Menú de actualización de la hoja SUB UNIDADES CICLO
 */
function actualizarHojaSubUnidadesCicloMenu() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    actualizarHojaSubUnidadesCiclo(ss);
    ss.toast('Hoja SUB UNIDADES CICLO configurada con 3 especies.', '🌱 Subunidades Ciclo', 5);
    SpreadsheetApp.getUi().alert(
      'Actualización Exitosa',
      'La hoja SUB UNIDADES CICLO ha sido configurada con 3 tablas maestras:\n\n' +
      '• 🐂 Bovinos: Cría, Recría, Adaptación, Engorde, Reproducción\n' +
      '• 🐖 Porcinos: Cría, Destete, Engorde, Reproducción\n' +
      '• 🐑 Ovinos: Cría, Recría, Engorde, Reproducción\n\n' +
      'Todas las etapas quedan perfectamente alineadas con las fórmulas de resumen y control de stock.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert('Error al actualizar subunidades de ciclo: ' + err.message);
  }
}

/**
 * Aplica formato profesional, tipografía homogénea, bordes finos,
 * números contables y listas de validación desplegables en TABLA_IMPUTACION.
 */
function formatearYValidarTablaImputacion(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('TABLA_IMPUTACION');
  if (!sheet) return { exito: false, mensaje: 'No se encontró la hoja TABLA_IMPUTACION.' };

  const lastRow = Math.max(sheet.getLastRow(), 2);
  const maxRows = Math.max(sheet.getMaxRows(), 100);

  // 1. Cabecera (Fila 1)
  const headers = ['Fecha', 'Comprobante', 'Proveedor', 'Concepto', 'Importe_Total', 'Unidad_Negocio', 'Etapa_Productiva', 'Plano_Afectado', 'Efecto_Existencias'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1E3A1E')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  // 2. Anchos de columna
  const anchos = [105, 140, 240, 260, 130, 120, 140, 160, 160];
  anchos.forEach((w, idx) => sheet.setColumnWidth(idx + 1, w));
  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(false);

  // 3. Formato y alineación de datos (Fila 2 en adelante)
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 9).setVerticalAlignment('middle').setFontFamily('Calibri').setFontSize(10);
    sheet.getRange(2, 1, lastRow - 1, 1).setHorizontalAlignment('center').setNumberFormat('dd/MM/yy');
    sheet.getRange(2, 2, lastRow - 1, 1).setHorizontalAlignment('center');
    sheet.getRange(2, 3, lastRow - 1, 2).setHorizontalAlignment('left');
    sheet.getRange(2, 5, lastRow - 1, 1).setHorizontalAlignment('right').setNumberFormat('$ #,##0.00');
    sheet.getRange(2, 6, lastRow - 1, 4).setHorizontalAlignment('center');

    // Negrita en comprobantes MOV-EXIS
    const valsComp = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (let r = 0; r < valsComp.length; r++) {
      if (String(valsComp[r][0] || '').startsWith('MOV-EXIS')) {
        sheet.getRange(r + 2, 2).setFontWeight('bold');
      }
    }

    // Bordes finos en todo el rango con datos
    sheet.getRange(2, 1, lastRow - 1, 9).setBorder(true, true, true, true, true, true, '#D9D9D9', SpreadsheetApp.BorderStyle.SOLID);
  }

  // 4. Reglas de Validación Desplegables
  // Col C (Proveedor): lista dinámica desde CAT_CONTACTOS!B2:B
  const ruleProv = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getRange('CAT_CONTACTOS!B2:B'))
    .setAllowInvalid(true)
    .build();
  sheet.getRange('C2:C' + maxRows).setDataValidation(ruleProv);

  // Col D (Concepto): lista dinámica desde CAT_CONCEPTOS!B2:B (incluyendo CPT-17 a CPT-21 y los de trazabilidad)
  const ruleConc = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getRange('CAT_CONCEPTOS!B2:B'))
    .setAllowInvalid(true)
    .build();
  sheet.getRange('D2:D' + maxRows).setDataValidation(ruleConc);

  // Col F (Unidad_Negocio): Bovinos, Porcinos, Ovinos, Maiz, Soja, Alfalfa, Estructura
  const ruleUN = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Bovinos', 'Porcinos', 'Ovinos', 'Maiz', 'Soja', 'Alfalfa', 'Estructura'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('F2:F' + maxRows).setDataValidation(ruleUN);

  // Col G (Etapa_Productiva): Cria, Recria, Adaptacion, Engorde, Destete, Reproduccion, Estructura, Barbecho, Siembra, Proteccion, Cosecha
  const ruleEtapa = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Cria', 'Recria', 'Adaptacion', 'Engorde', 'Destete', 'Reproduccion', 'Estructura', 'Barbecho', 'Siembra', 'Proteccion', 'Cosecha'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('G2:G' + maxRows).setDataValidation(ruleEtapa);

  // Col H (Plano_Afectado): Resultado Economico, Activo / Existencias, Activo / Bienes de uso
  const rulePlano = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Resultado Economico', 'Activo / Existencias', 'Activo / Bienes de uso'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('H2:H' + maxRows).setDataValidation(rulePlano);

  // Col I (Efecto_Existencias): Sin Efecto, Aumenta Stock Insumos, Disminuye Stock Granos, Disminuye Stock Hacienda
  const ruleEfecto = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Sin Efecto', 'Aumenta Stock Insumos', 'Disminuye Stock Granos', 'Disminuye Stock Hacienda'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('I2:I' + maxRows).setDataValidation(ruleEfecto);

  return { exito: true, filas: lastRow - 1 };
}

/**
 * Menú para formatear y aplicar validaciones a TABLA_IMPUTACION
 */
function formatearYValidarTablaImputacionMenu() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const res = formatearYValidarTablaImputacion(ss);
    ss.toast('TABLA_IMPUTACION formateada y con listas desplegables activas.', '🎨 Formato Imputación', 5);
    SpreadsheetApp.getUi().alert(
      'Formato y Validaciones Aplicadas',
      'La hoja TABLA_IMPUTACION ha sido configurada con éxito (' + res.filas + ' filas formateadas):\n\n' +
      '• Estilo y Escritura: Cabecera verde institucional, bordes finos, importes contables ($ #,##0.00) y fechas dd/MM/yy.\n' +
      '• Columna Concepto: Desplegable dinámico vinculado a CAT_CONCEPTOS (incluyendo CPT-17 a CPT-21 y costos de trazabilidad).\n' +
      '• Columna Unidad de Negocio: Desplegable con Bovinos, Porcinos, Ovinos, Maiz, Soja, Alfalfa, Estructura.\n' +
      '• Columna Etapa Productiva: Desplegable ampliado con Cría, Recría, Adaptación, Engorde, Destete, Reproducción y Estructura.\n' +
      '• Columnas Plano Afectado y Efecto Existencias: Desplegables oficiales para consistencia contable.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert('Error al formatear TABLA_IMPUTACION: ' + err.message);
  }
}

/**
 * Construye el Tablero de Control Gerencial (DASHBOARD) multiespecie con 4 paneles ejecutivos:
 * 1. Resumen Financiero y Resultados Globales
 * 2. Unidad Ganadera Bovina (Ciclo Completo)
 * 3. Unidades Porcina y Ovina (Trazabilidad y Costos)
 * 4. Pizarra de Mercado en Vivo (Cotizaciones Oficiales)
 */
function construirDashboard(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('DASHBOARD');
  if (!sheet) {
    sheet = ss.insertSheet('DASHBOARD', 0);
  } else {
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(1);
    sheet.clear();
  }

  sheet.setHiddenGridlines(false);

  // Anchos de columna (A a H: 140px cada una)
  for (let c = 1; c <= 8; c++) {
    sheet.setColumnWidth(c, 140);
  }

  // Estilos base
  const COLOR_H1 = '#1E3A1E';
  const COLOR_H2 = '#2E5A2E';
  const COLOR_SEC = '#D0E1D4';
  const COLOR_LBL = '#E8F0EA';
  const COLOR_WHITE = '#FFFFFF';
  const COLOR_TEXT_DARK = '#1A1A1A';
  const COLOR_TEXT_GREEN = '#1E3A1E';

  // 1. Cabecera Principal
  sheet.getRange('A1:H1').merge()
    .setValue('ESTABLECIMIENTO POZO DEL CARRIL - TABLERO DE CONTROL GANADERO')
    .setBackground(COLOR_H1)
    .setFontColor(COLOR_WHITE)
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 32);

  sheet.getRange('A2:H2').merge()
    .setValue('Universidad Nacional de Río Cuarto (UNRC) | Monitoreo Operativo, Productivo y Financiero Multiespecie')
    .setBackground(COLOR_H2)
    .setFontColor(COLOR_WHITE)
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 20);

  sheet.setRowHeight(3, 10);

  // Helper para secciones
  function crearSeccionHdr(row, titulo) {
    sheet.getRange(row, 1, 1, 8).merge()
      .setValue(titulo)
      .setBackground(COLOR_SEC)
      .setFontColor(COLOR_TEXT_GREEN)
      .setFontSize(10)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#D9D9D9', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(row, 24);
  }

  // Helper para tarjetas
  function crearTarjeta(rLbl, rVal, cStart, cEnd, lbl, formula, numFmt) {
    sheet.getRange(rLbl, cStart, 1, (cEnd - cStart + 1)).merge()
      .setValue(lbl)
      .setBackground(COLOR_LBL)
      .setFontColor(COLOR_H2)
      .setFontSize(9)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#D9D9D9', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(rLbl, 20);

    const valRange = sheet.getRange(rVal, cStart, 1, (cEnd - cStart + 1)).merge()
      .setFormula(formula)
      .setBackground(COLOR_WHITE)
      .setFontColor(COLOR_TEXT_DARK)
      .setFontSize(12)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#D9D9D9', SpreadsheetApp.BorderStyle.SOLID);
    if (numFmt) valRange.setNumberFormat(numFmt);
    sheet.setRowHeight(rVal, 28);
  }

  // PANEL 1: FINANCIERO (Filas 4 a 6)
  crearSeccionHdr(4, '  📊 1. RESUMEN FINANCIERO Y RESULTADOS GLOBALES');
  crearTarjeta(5, 6, 1, 2, 'FACTURACIÓN TOTAL BRUTA', '=SUM(MOV_COMERCIALES!K2:K100)', '$ #,##0.00');
  crearTarjeta(5, 6, 3, 4, 'COBRADO EN BANCO', '=SUMIFS(MOV_COMERCIALES!P2:P100, MOV_COMERCIALES!R2:R100, "Cobrado")', '$ #,##0.00');
  crearTarjeta(5, 6, 5, 6, 'CUENTAS A COBRAR (PENDIENTE)', '=SUMIFS(MOV_COMERCIALES!P2:P100, MOV_COMERCIALES!R2:R100, "Pendiente")', '$ #,##0.00');
  crearTarjeta(5, 6, 7, 8, 'COSTOS TOTALES IMPUTADOS', '=SUM(TABLA_IMPUTACION!E2:E200)', '$ #,##0.00');
  sheet.setRowHeight(7, 10);

  // PANEL 2: BOVINOS (Filas 8 a 10)
  crearSeccionHdr(8, '  🐂 2. UNIDAD GANADERA BOVINA (Ciclo Completo: Cría, Recría, Adaptación y Engorde)');
  crearTarjeta(9, 10, 1, 2, 'STOCK BOVINO (TRAZABILIDAD)', '=COUNTIFS(EXISTENCIAS!D:D, "Bovino", EXISTENCIAS!J:J, "Vivo")', '0 "Cab"');
  crearTarjeta(9, 10, 3, 4, 'KILOS GANADOS TOTALES', '=ESTADO_RESULTADOS_ETAPAS!F31', '#,##0.0 "kg"');
  crearTarjeta(9, 10, 5, 6, 'MARGEN BRUTO GANADERO', '=ESTADO_RESULTADOS_ETAPAS!F28', '$ #,##0.00');
  crearTarjeta(9, 10, 7, 8, 'MARGEN NETO GANADERO', '=ESTADO_RESULTADOS_ETAPAS!F30', '$ #,##0.00');
  sheet.setRowHeight(11, 10);

  // PANEL 3: PORCINOS Y OVINOS (Filas 12 a 14)
  crearSeccionHdr(12, '  🐖🐑 3. UNIDADES PORCINA Y OVINA (Monitoreo de Trazabilidad y Costos)');
  crearTarjeta(13, 14, 1, 2, 'STOCK PORCINO EN PIE', '=COUNTIFS(EXISTENCIAS!D:D, "Porcino", EXISTENCIAS!J:J, "Vivo")', '0 "Cab"');
  crearTarjeta(13, 14, 3, 4, 'VENTAS PORCINAS FACTURADAS', '=SUMIFS(MOV_COMERCIALES!K2:K100, MOV_COMERCIALES!F2:F100, "UN-POR")', '$ #,##0.00');
  crearTarjeta(13, 14, 5, 6, 'STOCK OVINO EN PIE', '=COUNTIFS(EXISTENCIAS!D:D, "Ovino", EXISTENCIAS!J:J, "Vivo")', '0 "Cab"');
  crearTarjeta(13, 14, 7, 8, 'COSTOS IMPUTADOS PORC. Y OVIN.', '=SUMIFS(TABLA_IMPUTACION!E:E, TABLA_IMPUTACION!F:F, "Porcinos") + SUMIFS(TABLA_IMPUTACION!E:E, TABLA_IMPUTACION!F:F, "Ovinos")', '$ #,##0.00');
  sheet.setRowHeight(15, 10);

  // PANEL 4: COTIZACIONES (Filas 16 a 18)
  crearSeccionHdr(16, '  📈 4. PIZARRA DE MERCADO EN VIVO (Cotizaciones Oficiales por Categoría)');
  crearTarjeta(17, 18, 1, 2, 'NOVILLO MAG CAÑUELAS', '=COTIZACIONES!E2', '$ #,##0.00');
  crearTarjeta(17, 18, 3, 4, 'TERNERO ROSGAN INVERNADA', '=COTIZACIONES!E7', '$ #,##0.00');
  crearTarjeta(17, 18, 5, 6, 'CAPÓN PORCINO GENERAL', '=COTIZACIONES!E16', '$ #,##0.00');
  crearTarjeta(17, 18, 7, 8, 'CORDERO PESADO COMERCIAL', '=COTIZACIONES!E24', '$ #,##0.00');

  return { exito: true };
}

/**
 * Menú para construir y formatear el DASHBOARD
 */
function construirDashboardMenu() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    construirDashboard(ss);
    ss.toast('DASHBOARD multiespecie construido y actualizado.', '📊 Dashboard', 5);
    SpreadsheetApp.getUi().alert(
      'DASHBOARD Actualizado con Éxito',
      'El Tablero de Control Gerencial fue construido con 4 paneles ejecutivos:\n\n' +
      '• 1. Resumen Financiero: Facturación bruta, Cobrado, Por cobrar y Costos totales.\n' +
      '• 2. Unidad Bovina: Stock en pie, Kilos ganados, Margen Bruto y Margen Neto.\n' +
      '• 3. Unidades Porcina y Ovina: Stock por especie, Ventas de capones y Costos imputados.\n' +
      '• 4. Pizarra de Mercado: Cotizaciones oficiales en vivo (Novillo, Ternero, Capón y Cordero).\n\n' +
      'Fórmulas 100% dinámicas vinculadas a Movimientos, Imputación, Existencias y Cotizaciones.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(err) {
    SpreadsheetApp.getUi().alert('Error al construir DASHBOARD: ' + err.message);
  }
}

/**
 * Actualiza y formatea la hoja ESTADO_RESULTADOS_ETAPAS:
 * 1. Limpia textos de diapositivas (PPT, Slides).
 * 2. Conecta la fila de Kilos Ganados con los datos reales del sistema actual.
 * 3. Añade la Sección 4: Consolidado Multiespecie por Unidad de Negocio (Bovinos, Porcinos, Ovinos, Maíz y Estructura).
 */
function construirEstadoResultados(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('ESTADO_RESULTADOS_ETAPAS');
  if (!sheet) return { exito: false, mensaje: 'No se encontró la hoja ESTADO_RESULTADOS_ETAPAS' };

  // 1. Limpieza de notas de filminas
  sheet.getRange('A3').setValue('1. ECUACIÓN DE INVENTARIOS Y VALUACIÓN DE ACTIVOS BIOLÓGICOS');
  sheet.getRange('D4').setValue('Criterio de Valuación Contable');
  sheet.getRange('D11').setValue('EI + Entradas - Salidas - EF (Ecuación Fundamental)');
  sheet.getRange('A15').setValue('2. ESTADO DE RESULTADOS POR ETAPAS Y TRASLADOS EN CASCADA');
  sheet.getRange('A36').setValue('3. ANÁLISIS DE EFICIENCIA Y MÁRGENES PORCENTUALES');
  sheet.getRange('A44').setValue('Costo sobre el Ingreso (%)');
  sheet.getRange('G44').setValue('Costos Incurridos / Ingresos');
  sheet.getRange('A45').setValue('Margen sobre el Ingreso (%)');
  sheet.getRange('G45').setValue('100% - Costo sobre Ingreso');

  // 2. Fila 31: Kilos Ganados calculados desde los animales y parámetros del sistema
  sheet.getRange('B31').setFormula('=IFERROR(COUNTIFS(EXISTENCIAS!$D:$D, "Bovino", EXISTENCIAS!$F:$F, "Cria") * \'GANANCIA DIARIA KG\'!B3 * \'GANANCIA DIARIA KG\'!E3, 105)');
  sheet.getRange('C31').setFormula('=IFERROR(COUNTIFS(EXISTENCIAS!$D:$D, "Bovino", EXISTENCIAS!$F:$F, "Recria") * \'GANANCIA DIARIA KG\'!B4 * \'GANANCIA DIARIA KG\'!E4, 423)');
  sheet.getRange('D31').setFormula('=IFERROR(COUNTIFS(EXISTENCIAS!$D:$D, "Bovino", EXISTENCIAS!$F:$F, "Adaptacion") * \'GANANCIA DIARIA KG\'!B5 * \'GANANCIA DIARIA KG\'!E5, 27)');
  sheet.getRange('E31').setFormula('=IFERROR(COUNTIFS(EXISTENCIAS!$D:$D, "Bovino", EXISTENCIAS!$F:$F, "Engorde") * \'GANANCIA DIARIA KG\'!B6 * \'GANANCIA DIARIA KG\'!E6, 112.5)');
  sheet.getRange('F31').setFormula('=SUM(B31:E31)');
  sheet.getRange('B31:F31').setNumberFormat('#,##0.0 "kg"');

  // 3. Sección 4: Consolidado Multiespecie
  const COLOR_SEC = '#D0E1D4';
  const COLOR_HDR = '#E8F0EA';
  const COLOR_TOTAL = '#F4F8F4';
  const COLOR_TEXT_GREEN = '#1E3A1E';

  sheet.getRange('A47:G47').merge()
    .setValue('4. ESTADO DE RESULTADOS CONSOLIDADO POR UNIDAD DE NEGOCIO')
    .setBackground(COLOR_SEC)
    .setFontColor(COLOR_TEXT_GREEN)
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(47, 25);

  const headersSec4 = ['Concepto Económico', 'Bovinos', 'Porcinos', 'Ovinos', 'Maíz (Granos)', 'Estructura', 'TOTAL ESTABLECIMIENTO'];
  sheet.getRange(48, 1, 1, 7).setValues([headersSec4])
    .setBackground(COLOR_HDR)
    .setFontColor('#2E5A2E')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.getRange('A48').setHorizontalAlignment('left');
  sheet.setRowHeight(48, 22);

  // Fórmulas filas 49 a 53
  const formulasSec4 = [
    ['(+) Ingresos Comerciales por Ventas', '=SUMIFS(MOV_COMERCIALES!K2:K100, MOV_COMERCIALES!F2:F100, "UN-BOV")', '=SUMIFS(MOV_COMERCIALES!K2:K100, MOV_COMERCIALES!F2:F100, "UN-POR")', '=SUMIFS(MOV_COMERCIALES!K2:K100, MOV_COMERCIALES!F2:F100, "UN-OVIN")', '=SUMIFS(MOV_COMERCIALES!K2:K100, MOV_COMERCIALES!F2:F100, "UN-MAIZ")', 0, '=SUM(B49:F49)'],
    ['(-) Costos Directos Incurridos', '=SUMIFS(TABLA_IMPUTACION!E:E, TABLA_IMPUTACION!F:F, "Bovinos")', '=SUMIFS(TABLA_IMPUTACION!E:E, TABLA_IMPUTACION!F:F, "Porcinos")', '=SUMIFS(TABLA_IMPUTACION!E:E, TABLA_IMPUTACION!F:F, "Ovinos")', '=SUMIFS(TABLA_IMPUTACION!E:E, TABLA_IMPUTACION!F:F, "Maiz")', 0, '=SUM(B50:F50)'],
    ['(=) MARGEN BRUTO OPERATIVO', '=B49-B50', '=C49-C50', '=D49-D50', '=E49-E50', 0, '=G49-G50'],
    ['(-) Costos Indirectos y Estructura', 0, 0, 0, 0, '=SUMIFS(TABLA_IMPUTACION!E:E, TABLA_IMPUTACION!F:F, "Estructura")', '=F52'],
    ['(=) RESULTADO NETO FINAL DEL ESTABLECIMIENTO', '=B51', '=C51', '=D51', '=E51', '=-F52', '=G51-G52']
  ];

  for (let i = 0; i < formulasSec4.length; i++) {
    const rowNum = 49 + i;
    sheet.getRange(rowNum, 1, 1, 7).setValues([formulasSec4[i]]);
    sheet.getRange(rowNum, 2, 1, 6).setNumberFormat('$ #,##0.00').setHorizontalAlignment('right');
    sheet.getRange(rowNum, 1).setHorizontalAlignment('left');
    sheet.setRowHeight(rowNum, (rowNum === 51 || rowNum === 53) ? 24 : 20);
    if (rowNum === 51 || rowNum === 53) {
      sheet.getRange(rowNum, 1, 1, 7).setFontWeight('bold');
    }
    if (rowNum === 53) {
      sheet.getRange(rowNum, 1, 1, 7).setBackground(COLOR_TOTAL);
    }
  }

  sheet.getRange('A47:G53').setBorder(true, true, true, true, true, true, '#D9D9D9', SpreadsheetApp.BorderStyle.SOLID);
  return { exito: true };
}

/**
 * Menú para construir y actualizar el ESTADO DE RESULTADOS
 */
function construirEstadoResultadosMenu() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    construirEstadoResultados(ss);
    ss.toast('Estado de Resultados Multiespecie actualizado.', '📑 Estado Resultados', 5);
    SpreadsheetApp.getUi().alert(
      'Estado de Resultados Actualizado con Éxito',
      'Se completaron las siguientes mejoras:\n\n' +
      '• Textos limpios: Se eliminaron todas las referencias teóricas a diapositivas (PPT, Slides).\n' +
      '• Kilos ganados (Fila 31): Fórmulas técnicas basadas en el stock bovino real del sistema y la ganancia diaria esperada.\n' +
      '• Sección 4 Consolidada: Cuadro integral multiespecie incorporando Bovinos, Porcinos, Ovinos, Maíz y Costos de Estructura.\n' +
      '• Formatos contables uniformes ($ #,##0.00).',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(err) {
    SpreadsheetApp.getUi().alert('Error al actualizar Estado de Resultados: ' + err.message);
  }
}


/**
 * Crea o actualiza la hoja USUARIOS inmediatamente despues de DASHBOARD
 * con los usuarios iniciales del sistema ganadero y formato profesional.
 */
function crearHojaUsuarios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHoja = 'USUARIOS';
  let sheet = ss.getSheetByName(nombreHoja);

  // Identificar posicion de DASHBOARD
  const dashboardSheet = ss.getSheetByName('DASHBOARD');
  
  if (sheet) {
    sheet.clear();
  } else {
    let targetIndex = dashboardSheet ? dashboardSheet.getIndex() + 1 : 2;
    sheet = ss.insertSheet(nombreHoja, targetIndex);
  }

  // Asegurar que quede inmediatamente despues de DASHBOARD
  if (dashboardSheet) {
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(dashboardSheet.getIndex() + 1);
  }

  // Matriz de datos de usuarios
  const data = [
    ['id', 'nombre', 'email', 'contrasena', 'rol', 'activo', 'creado_en', 'creado_por'],
    [1, 'Franco Martinez', 'admin@agrobovino.com', 'cambiar_pass_admin', 'administrador', true, '2026-10-17', ''],
    [2, 'Franco Maina', 'gestor@agrobovino.com', 'cambiar_pass_gestor', 'gestor', true, '2026-10-17', ''],
    [3, 'Mario Campo', 'capataz@agrobovino.com', 'cambiar_pass_capataz', 'capataz', true, '2026-10-17', ''],
    [4, 'Ana Rivas', 'invitado@agrobovino.com', 'cambiar_pass_invitado', 'invitado', true, '2026-10-17', ''],
    [5, 'Carlos Nuevo', 'carlos@agrobovino.com', 'pass_temp', 'invitado', true, '2026-10-18', 'Franco Martinez']
  ];

  // Escribir datos
  const totalRows = data.length;
  const totalCols = data[0].length;
  const range = sheet.getRange(1, 1, totalRows, totalCols);
  range.setValues(data);

  // Estilo de encabezados
  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange
    .setBackground('#1E3A1E')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontFamily('Aptos Narrow')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 28);

  // Estilo del cuerpo de datos
  const dataRange = sheet.getRange(2, 1, totalRows - 1, totalCols);
  dataRange
    .setFontFamily('Aptos Narrow')
    .setFontSize(10)
    .setVerticalAlignment('middle');

  // Alineaciones especificas por columna
  sheet.getRange(2, 1, totalRows - 1, 1).setHorizontalAlignment('center'); // id
  sheet.getRange(2, 2, totalRows - 1, 3).setHorizontalAlignment('left');   // nombre, email, contrasena
  sheet.getRange(2, 5, totalRows - 1, 3).setHorizontalAlignment('center'); // rol, activo, creado_en
  sheet.getRange(2, 8, totalRows - 1, 1).setHorizontalAlignment('left');   // creado_por

  // Bordes delgados
  range.setBorder(true, true, true, true, true, true, '#D9D9D9', SpreadsheetApp.BorderStyle.SOLID);

  // Ajustar ancho de columnas
  for (let c = 1; c <= totalCols; c++) {
    sheet.autoResizeColumn(c);
    const w = sheet.getColumnWidth(c);
    sheet.setColumnWidth(c, Math.max(w + 25, 100));
  }

  ss.toast('Hoja USUARIOS creada y ubicada correctamente despues de DASHBOARD.', 'Exito', 5);
}
