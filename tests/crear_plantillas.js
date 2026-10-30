/**
 * Crea:
 *  - plantilla-ejemplo.xlsx  (DATOS FICTICIOS — seguro para el repo público)
 *  - /workspace/recibos_web_test/setiembre-real.xlsx  (datos reales, NO se sube al repo)
 */
const fs = require('fs');
const XLSX = require('../vendor/xlsx-0.20.3.full.min.js');
const E = require('../js/excel.js');
const C = require('../js/calculo.js');
E.usar(XLSX);

// Logo ficticio: cuadrado naranja con "EJ"
function logoFicticio() {
  // PNG 64x64 naranja mínimo vía canvas no disponible en node puro → usamos un PNG 1x1 y lo documentamos,
  // o generamos con un PNG embebido simple.
  // PNG 64×64 naranja sólido (comprimido a mano es largo). Usamos un JPEG data URL corto pintado con pure JS:
  // Mejor: leer el logo real solo para la prueba local; en el ejemplo público no hay logo (o uno genérico).
  return '';
}

const ficticio = {
  config: {
    EDIFICIO_NOMBRE: 'EDIFICIO EJEMPLO',
    DIRECCION_LINEA_1: 'Dirección: Av. Los Olivos N°123 /',
    DIRECCION_LINEA_2: 'Urb. Las Flores. Lima',
    PRESIDENTE_NOMBRE: 'Ana Pérez',
    PRESIDENTE_CELULAR: '999000111',
    TESORERO_NOMBRE: 'Luis Gómez',
    TESORERO_CELULAR: '999000222',
    BANCO: 'BCP',
    CUENTA_BANCARIA: '19100000000000',
    CCI: '00219100000000000000',
    NOTA_PAGO: 'Por favor, envíe su comprobante de pago al tesorero vía WhatsApp indicando su número de departamento.',
    EMPRESA_AGUA: 'Sedapal', EMPRESA_LUZ: 'Luz del Sur',
    NUM_DPTOS: 10, DIA_VENCIMIENTO: 5, SALDO_INICIAL_FONDO: 4097.67
  },
  deptos: [
    { unidad: 101, propietario: 'Juan Demo', correo: 'dpto101@ejemplo.com' },
    { unidad: 102, propietario: 'María Ejemplo', correo: 'dpto102@ejemplo.com, inquilino102@ejemplo.com' },
    { unidad: 201, propietario: 'Pedro Prueba', correo: 'dpto201@ejemplo.com' },
    { unidad: 202, propietario: 'Lucía Sample', correo: 'dpto202@ejemplo.com' },
    { unidad: 301, propietario: 'Carlos Ficticio', correo: 'dpto301@ejemplo.com' },
    { unidad: 302, propietario: 'Sofía Modelo', correo: 'dpto302@ejemplo.com' },
    { unidad: 401, propietario: 'Diego Ensayo', correo: 'dpto401@ejemplo.com' },
    { unidad: 402, propietario: 'Elena Dummy', correo: 'dpto402@ejemplo.com' },
    { unidad: 501, propietario: 'Miguel Test', correo: 'dpto501@ejemplo.com' },
    { unidad: 502, propietario: 'Laura Muestra', correo: 'dpto502@ejemplo.com' }
  ],
  // Mismos números que setiembre real (para que el cálculo de ejemplo coincida), nombres inventados
  lecturas: [
    { mes: '2026-09', unidad: 101, anterior: 1281.328, actual: 1290.107 },
    { mes: '2026-09', unidad: 102, anterior: 1690.315, actual: 1727.609 },
    { mes: '2026-09', unidad: 201, anterior: 39.858, actual: 53.561 },
    { mes: '2026-09', unidad: 202, anterior: 674.427, actual: 688.229 },
    { mes: '2026-09', unidad: 301, anterior: 459.125, actual: 465.722 },
    { mes: '2026-09', unidad: 302, anterior: 974.338, actual: 993.461 },
    { mes: '2026-09', unidad: 401, anterior: 1135.912, actual: 1151.465 },
    { mes: '2026-09', unidad: 402, anterior: 291.616, actual: 295.599 },
    { mes: '2026-09', unidad: 501, anterior: 1802.587, actual: 1827.461 },
    { mes: '2026-09', unidad: 502, anterior: 771.98, actual: 776.69 }
  ],
  servicios: [{ mes: '2026-09', emision: '2026-09-25', vencimiento: '2026-10-05', aguaImporte: 679.4, aguaM3: 154, luzImporte: 510.4 }],
  gastos: [
    { mes: '2026-09', num: 1, concepto: 'Salario del personal', monto: 1573.34, tipo: 'FIJO', copiar: 'SÍ' },
    { mes: '2026-09', num: 2, concepto: 'Consumo de luz', monto: '', tipo: 'LUZ', copiar: 'SÍ' },
    { mes: '2026-09', num: 3, concepto: 'Mantenimiento de ascensor', monto: 349.34, tipo: 'FIJO', copiar: 'SÍ' },
    { mes: '2026-09', num: 4, concepto: 'Jardinero / Mantenimiento de áreas verdes', monto: 90, tipo: 'FIJO', copiar: 'SÍ' },
    { mes: '2026-09', num: 5, concepto: 'Agua común', monto: '', tipo: 'AGUA_COMUN', copiar: 'SÍ' },
    { mes: '2026-09', num: 6, concepto: 'Fondo de contingencia', monto: 800, tipo: 'FONDO', copiar: 'SÍ' },
    { mes: '2026-09', num: 7, concepto: 'Recarga de celular conserje', monto: 5, tipo: 'FIJO', copiar: 'SÍ' },
    { mes: '2026-09', num: 8, concepto: 'Domingos - días (6,13,20,27)', monto: 200, tipo: 'FIJO', copiar: 'SÍ' },
    { mes: '2026-09', num: 9, concepto: 'Mantenimiento de elevador discapacitados', monto: 200, tipo: 'EXTRA', copiar: 'NO' },
    { mes: '2026-09', num: 10, concepto: 'Compra de temporizador con base para tablero de mando (elevador discapacitados)', monto: 400, tipo: 'EXTRA', copiar: 'NO' },
    { mes: '2026-09', num: 11, concepto: 'Gasfitero (desmontaje e instalación de inodoro baño semisotano conserje) por atoro, se cambio pernos de anclaje, anillo de cera con silicona', monto: 180, tipo: 'EXTRA', copiar: 'NO' },
    { mes: '2026-09', num: 12, concepto: 'Materiales de limpieza', monto: 0, tipo: 'FIJO', copiar: 'SÍ' }
  ],
  ajustes: [{ mes: '2026-09', unidad: 501, concepto: 'Crédito a favor', monto: -5 }],
  pagos: [],
  fondo: [{ mes: '2026-09', saldoAnterior: 4097.67, aporte: 800, gastosFondo: 0, saldo: 4897.67 }],
  logo: ''
};

// Verifica cálculo
const r = C.calcularMes(Object.assign({ mes: '2026-09' }, ficticio));
if (r.unidades.find(u => u.unidad == 101).total !== 471.99) throw new Error('101 != 471.99: ' + r.unidades.find(u => u.unidad == 101).total);
if (r.totalFacturado !== 4987.48) throw new Error('total');
console.log('cálculo OK', r.unidades.find(u => u.unidad == 101).total);

// Plantilla pública: sin historial; Setiembre (ficticio) escrito en la pestaña "Mes a facturar"
const bEj = E.aplicarBorrador ? E.prepararBorrador(ficticio, '2026-09') : null;
bEj.servicios.emision = ''; bEj.existe = false;
const vacioHist = Object.assign({}, ficticio, { lecturas: [], servicios: [], gastos: [], ajustes: [], fondo: [] });
const wbEj = E.libroNuevo(vacioHist, { borrador: bEj });
fs.writeFileSync('/workspace/recibos_web/plantilla-ejemplo.xlsx', XLSX.write(wbEj,{bookType:'xlsx',type:'buffer'}));

// Real: a partir del verificacion_google + nombres reales
const realJson = JSON.parse(fs.readFileSync('/workspace/sistema_recibos/ejemplo_setiembre.json', 'utf8'));
const L = E.leerLibro(XLSX.read(fs.readFileSync('/workspace/sistema_recibos/verificacion_google.xlsx')));
const real = L.datos;
// Asegura lecturas 501/502 completas desde json si faltan
realJson.lecturas.forEach(function (l) {
  if (!real.lecturas.some(x => x.mes === l.mes && String(x.unidad) === String(l.unidad))) real.lecturas.push(l);
});
fs.mkdirSync('/workspace/recibos_web_test', { recursive: true });
fs.writeFileSync('/workspace/recibos_web_test/setiembre-real.xlsx', XLSX.write(E.libroNuevo(real),{bookType:'xlsx',type:'buffer'}));
// Real "como lo usará ella": historial hasta Agosto + Setiembre escrito en la pestaña "Mes a facturar"
const bReal = E.prepararBorrador(real, '2026-09');
const histReal = Object.assign({}, real, {
  lecturas: real.lecturas.filter(x => x.mes < '2026-09'), servicios: real.servicios.filter(x => x.mes < '2026-09'),
  gastos: real.gastos.filter(x => x.mes < '2026-09'), ajustes: real.ajustes.filter(x => x.mes < '2026-09'), fondo: real.fondo.filter(x => x.mes < '2026-09') });
fs.writeFileSync('/workspace/recibos_web_test/setiembre-real-mes.xlsx', XLSX.write(E.libroNuevo(histReal, { borrador: bReal }),{bookType:'xlsx',type:'buffer'}));
console.log('plantillas OK');
