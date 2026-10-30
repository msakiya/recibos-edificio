const assert = require('assert');
const fs = require('fs');
const XLSX = require('../vendor/xlsx-0.20.3.full.min.js');
const C = require('../js/calculo.js');
const E = require('../js/excel.js');
const M = require('../js/media.js');
E.usar(XLSX);

const datos = JSON.parse(fs.readFileSync('/workspace/sistema_recibos/ejemplo_setiembre.json', 'utf8'));
const res = C.calcularMes(datos);
const u = n => res.unidades.find(x => x.unidad === n);
assert.strictEqual(u(101).total, 471.99);
assert.strictEqual(u(102).total, 597.80);
assert.strictEqual(u(501).total, 538.00);
assert.strictEqual(res.totalFacturado, 4987.48);
assert.strictEqual(res.fondo.saldo, 4897.67);
assert.strictEqual(res.agua.comun, 24.63);
assert.strictEqual(res.comunPorUnidad, 433.26);
assert.strictEqual(C.nombreEsDeUnidad('101.jpg', 101), true);
assert.strictEqual(C.nombreEsDeUnidad('Dpto 101.jpeg', 101), true);
assert.strictEqual(C.nombreEsDeUnidad('1010.jpg', 101), false);
assert.strictEqual(C.nombreEsDeUnidad('201.jpg', 101), false);
assert.strictEqual(C.parseNumero('1290,107'), 1290.107);
assert.strictEqual(C.parseNumero('1,290.107'), 1290.107);

// Excel ficticio
const wb = XLSX.read(fs.readFileSync('/workspace/recibos_web/plantilla-ejemplo.xlsx'));
const L = E.leerLibro(wb);
assert.strictEqual(L.errores.length, 0);
assert.strictEqual(L.datos.config.EDIFICIO_NOMBRE, 'EDIFICIO EJEMPLO');
assert.strictEqual(L.datos.config.CUENTA_BANCARIA, '19100000000000'); // ficticia
assert.ok(!/De la Gala|Aracelly|19117030817045/.test(JSON.stringify(L.datos)), 'no datos reales en plantilla pública');
const ME = E.cargarMes(wb, L.datos);
assert.deepStrictEqual(ME.errores, []);
assert.strictEqual(ME.borrador.mes, '2026-09');
const { res: r2 } = E.aplicarBorrador(L.datos, ME.borrador);
assert.strictEqual(r2.unidades.find(x => x.unidad == 101).total, 471.99);
assert.strictEqual(r2.totalFacturado, 4987.48);
assert.strictEqual(r2.fondo.saldo, 4897.67);

// Real
const wbR = XLSX.read(fs.readFileSync('/workspace/recibos_web_test/setiembre-real.xlsx'));
const LR = E.leerLibro(wbR);
const { res: r3 } = E.aplicarBorrador(LR.datos, E.prepararBorrador(LR.datos, '2026-09'));
assert.strictEqual(r3.unidades.find(x => x.unidad == 101).total, 471.99);

// Real con pestaña "Mes a facturar" (como lo llenará ella) + ida y vuelta al mes siguiente
{
  const wbM = XLSX.read(fs.readFileSync('/workspace/recibos_web_test/setiembre-real-mes.xlsx'), { cellDates: true });
  const LM = E.leerLibro(wbM);
  const MM = E.cargarMes(wbM, LM.datos);
  assert.deepStrictEqual(MM.errores, []);
  const ap = E.aplicarBorrador(LM.datos, MM.borrador);
  const u = n => ap.res.unidades.find(x => x.unidad == n).total;
  assert.strictEqual(u(101), 471.99); assert.strictEqual(u(102), 597.80); assert.strictEqual(u(501), 538.00);
  assert.strictEqual(ap.res.totalFacturado, 4987.48); assert.strictEqual(ap.res.fondo.saldo, 4897.67);
  // Genera: historial + pestaña para Octubre
  E.escribirMes(wbM, ap.datos, '2026-09');
  E.escribirHojaMes(wbM, ap.datos, '2026-10');
  const wb2 = XLSX.read(XLSX.write(wbM, { bookType: 'xlsx', type: 'buffer' }), { cellDates: true });
  assert.strictEqual(wb2.SheetNames[0], 'Mes a facturar');
  const L2 = E.leerLibro(wb2);
  const M2 = E.cargarMes(wb2, L2.datos);
  const txt = M2.errores.join(' | ');
  assert.ok(/importe del recibo de agua .*celda B5/.test(txt), txt);
  assert.ok(/m³ facturados .*celda B6/.test(txt), txt);
  assert.ok(/lectura actual .*101 \(D16\)/.test(txt), txt);
  const H = E.leerHojaMes(wb2, L2.datos);
  assert.strictEqual(E.mesDeTexto(H.campos.mes), '2026-10');
  assert.strictEqual(H.lecturas[0].anterior, 1290.107);            // anterior = actual de setiembre
  assert.strictEqual(H.campos.saldoAnterior, 4897.67);             // fondo arrastrado
  assert.ok(!H.gastos.some(g => g.tipo === 'EXTRA' && g.concepto));  // extras no se copian
  console.log('ida y vuelta Setiembre → Octubre OK');
}

// Media clasificación
const clas = M.clasificar([
  { name: '101.jpg', type: 'image/jpeg' },
  { name: 'Dpto 102.jpeg', type: 'image/jpeg' },
  { name: 'recibo-de-luz-2026.09.19.jpg', type: 'image/jpeg' },
  { name: 'recibo-de-agua-setiembre.pdf', type: 'application/pdf' },
  { name: 'comprobante 301.jpg', type: 'image/jpeg' },
  { name: 'foto_random.png', type: 'image/png' },
  { name: '1010.jpg', type: 'image/jpeg' }
], [101, 102, 201, 301, 1010]);
assert.strictEqual(clas.asignados.filter(a => a.tipo === 'medidor' && a.unidad === '101').length, 1);
assert.strictEqual(clas.asignados.filter(a => a.tipo === 'medidor' && a.unidad === '102').length, 1);
assert.strictEqual(clas.asignados.filter(a => a.tipo === 'luz').length, 1);
assert.strictEqual(clas.asignados.filter(a => a.tipo === 'agua').length, 1);
assert.strictEqual(clas.asignados.filter(a => a.tipo === 'pago' && a.unidad === '301').length, 1);
assert.ok(clas.sinReconocer.some(s => s.file.name === 'foto_random.png'));
assert.ok(clas.asignados.some(a => a.unidad === '1010') || clas.sinReconocer.some(s => /1010/.test(s.file.name)));

console.log('TODAS LAS PRUEBAS NODE OK');
