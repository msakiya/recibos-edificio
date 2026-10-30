// Pruebas rápidas de correo.js (node): teléfono, WhatsApp móvil/escritorio, mailto, HTML.
const assert = require('assert');
const C = require('../js/correo.js');
assert.strictEqual(C.telefonoWhatsapp('987654321'), '51987654321');
assert.strictEqual(C.telefonoWhatsapp(987654321), '51987654321');
assert.strictEqual(C.telefonoWhatsapp('987 654 321'), '51987654321');
assert.strictEqual(C.telefonoWhatsapp('+51 987-654-321'), '51987654321');
assert.strictEqual(C.telefonoWhatsapp('51987654321'), '51987654321');
assert.strictEqual(C.telefonoWhatsapp(''), '');
assert.strictEqual(C.telefonoWhatsapp('01 4445555'), '');   // fijo: no es WhatsApp
assert.strictEqual(C.telefonoBonito('987654321'), '987 654 321');
const cfg = { EDIFICIO_NOMBRE: 'EDIFICIO EJEMPLO', CORREO_NOMBRE_EDIFICIO: 'Edificio Ejemplo', BANCO: 'BCP', CUENTA_BANCARIA: '191', CCI: '002',
  TITULAR_CUENTA: 'Luis Gómez', TESORERO_CELULAR: '999000222' };
const link = 'https://drive.google.com/file/d/X/view';
let r = C.armar(cfg, '101', 'Setiembre 2026', 'S/ 471.99', 'a@b.pe', { link, telefono: '999001101', movil: true });
assert.ok(r.whatsapp.startsWith('https://wa.me/51999001101?text='));
assert.ok(decodeURIComponent(r.whatsapp).includes('Puedes ver tu recibo aquí 👉 ' + link));
assert.ok(r.mailto.startsWith('mailto:a@b.pe?subject=Recibo%20de%20Gastos%20Comunes%20-%20Setiembre%202026%20-%20Dpto%20101&body='));
assert.ok(decodeURIComponent(r.mailto).includes('VER RECIBO: ' + link));
assert.ok(r.texto.includes('Importante: Enviar constancia de pago al WhatsApp: 999 000 222'));   // por defecto: celular del tesorero
r = C.armar(Object.assign({}, cfg, { WHATSAPP_CONSTANCIA: '987654321' }), '101', 'Setiembre 2026', 'S/ 471.99', '', { link, telefono: '', movil: false });
assert.strictEqual(r.whatsapp, ''); assert.strictEqual(r.telefono, '');
assert.ok(r.html.includes('>VER RECIBO</a>') && r.html.includes('987 654 321') && r.html.includes('Titular de la Cuenta:'));
r = C.armar(cfg, '101', 'Setiembre 2026', 'S/ 471.99', '', { link: '', telefono: '999001101' });
assert.ok(r.whatsapp.startsWith('https://web.whatsapp.com/send?phone=51999001101&text='));
assert.ok(!r.html.includes('>VER RECIBO</a>'));
console.log('CORREO/WHATSAPP (unit) OK');
