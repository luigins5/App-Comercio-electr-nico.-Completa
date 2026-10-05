/**
 * Traductor seguro para la Base de Datos de la Tienda (Google Sheets)
 * Diseñado específicamente para la arquitectura de este proyecto (App Comercio Completo):
 * - Respeta la arquitectura JDB (JSON en columna Data)
 * - Mantiene intactos IDs, códigos, precios, imágenes, fechas y claves del sistema
 * - Traduce únicamente textos visibles (nombres, descripciones, notas, etiquetas)
 * - Conserva nombres de hojas y cabeceras exactas para no romper el backend
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌐 Tienda')
    .addItem('Traducir datos al Español (Seguro)', 'traducirBaseDeDatosTienda')
    .addSeparator()
    .addItem('Cargar Menú de Ejemplo en Español (seedMenu)', 'ejecutarSeedMenuDesdeMenu')
    .addToUi();
}

/**
 * Permite cargar el menú de bebidas en español predefinido en Code.gs
 */
function ejecutarSeedMenuDesdeMenu() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert(
    'Cargar Menú en Español',
    '¿Deseas poblar las categorías, productos y complementos en español predefinidos para la tienda?',
    ui.ButtonSet.YES_NO
  );
  if (res === ui.Button.YES) {
    withLock_(function() {
      seedMenu_();
    });
    ui.alert('¡Listo!', 'El menú en español ha sido cargado exitosamente en la base de datos.', ui.ButtonSet.OK);
  }
}

/**
 * Traduce de forma inteligente y segura los registros de la base de datos
 */
function traducirBaseDeDatosTienda() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    'Traducción Segura de la Base de Datos',
    'Esta función traducirá al español los textos visibles de:\n' +
    '• Categorías (CATEGORIES)\n' +
    '• Productos (PRODUCTS)\n' +
    '• Complementos (ADDONS)\n' +
    '• Métodos de pago (PAYMENT_METHODS)\n' +
    '• Ajustes visibles (SETTINGS)\n' +
    '• Etiquetas de roles (ROLES)\n\n' +
    'Mantendrá 100% intactas las claves JSON, IDs, fórmulas, nombres de columnas y la conexión con la app web.\n\n' +
    '¿Deseas continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Iniciando traducción segura de la tienda...', 'Traductor Tienda', -1);

  var stats = { categorias: 0, productos: 0, addons: 0, metodos: 0, ajustes: 0, roles: 0 };
  var cache = {};

  function traducirTexto(txt) {
    if (!txt || typeof txt !== 'string') return txt;
    var trimmed = txt.trim();
    if (!trimmed || !/[a-zA-Z]/.test(trimmed)) return txt;
    if (cache[trimmed] !== undefined) return cache[trimmed];
    try {
      var res = LanguageApp.translate(trimmed, '', 'es');
      cache[trimmed] = res;
      return res;
    } catch (e) {
      cache[trimmed] = txt;
      return txt;
    }
  }

  withLock_(function() {
    // 1. CATEGORÍAS (JDB)
    var shCat = sh_(CAT_SHEET);
    if (shCat && shCat.getLastRow() > 1) {
      var rowsCat = shCat.getRange(2, 1, shCat.getLastRow() - 1, 2).getValues();
      var changedCat = false;
      rowsCat.forEach(function(r) {
        var arr = safeParse_(r[1], []);
        if (Array.isArray(arr)) {
          arr.forEach(function(item) {
            if (item.name) {
              var trName = traducirTexto(item.name);
              if (trName !== item.name) { item.name = trName; stats.categorias++; changedCat = true; }
            }
            if (item.tagline) {
              var trTag = traducirTexto(item.tagline);
              if (trTag !== item.tagline) { item.tagline = trTag; changedCat = true; }
            }
          });
          r[1] = JSON.stringify(arr);
        }
      });
      if (changedCat) {
        shCat.getRange(2, 2, rowsCat.length, 1).setValues(rowsCat.map(function(r) { return [r[1]]; }));
        jdbDrop_(CAT_SHEET);
      }
    }

    // 2. PRODUCTOS (JDB)
    var shProd = sh_(PROD_SHEET);
    if (shProd && shProd.getLastRow() > 1) {
      var rowsProd = shProd.getRange(2, 1, shProd.getLastRow() - 1, 2).getValues();
      var changedProd = false;
      rowsProd.forEach(function(r) {
        var arr = safeParse_(r[1], []);
        if (Array.isArray(arr)) {
          arr.forEach(function(item) {
            if (item.name) {
              var trName = traducirTexto(item.name);
              if (trName !== item.name) { item.name = trName; stats.productos++; changedProd = true; }
            }
            if (item.description) {
              var trDesc = traducirTexto(item.description);
              if (trDesc !== item.description) { item.description = trDesc; changedProd = true; }
            }
            if (Array.isArray(item.sizes)) {
              item.sizes.forEach(function(sz) {
                if (sz.name) sz.name = traducirTexto(sz.name);
              });
            }
          });
          r[1] = JSON.stringify(arr);
        }
      });
      if (changedProd) {
        shProd.getRange(2, 2, rowsProd.length, 1).setValues(rowsProd.map(function(r) { return [r[1]]; }));
        jdbDrop_(PROD_SHEET);
      }
    }

    // 3. COMPLEMENTOS (ADDONS - JDB)
    var shAddon = sh_(ADDON_SHEET);
    if (shAddon && shAddon.getLastRow() > 1) {
      var rowsAdd = shAddon.getRange(2, 1, shAddon.getLastRow() - 1, 2).getValues();
      var changedAdd = false;
      rowsAdd.forEach(function(r) {
        var arr = safeParse_(r[1], []);
        if (Array.isArray(arr)) {
          arr.forEach(function(item) {
            if (item.name) {
              var trName = traducirTexto(item.name);
              if (trName !== item.name) { item.name = trName; stats.addons++; changedAdd = true; }
            }
          });
          r[1] = JSON.stringify(arr);
        }
      });
      if (changedAdd) {
        shAddon.getRange(2, 2, rowsAdd.length, 1).setValues(rowsAdd.map(function(r) { return [r[1]]; }));
        jdbDrop_(ADDON_SHEET);
      }
    }

    // 4. MÉTODOS DE PAGO (PAYMENT_METHODS - JDB)
    var shPm = sh_(PM_SHEET);
    if (shPm && shPm.getLastRow() > 1) {
      var rowsPm = shPm.getRange(2, 1, shPm.getLastRow() - 1, 2).getValues();
      var changedPm = false;
      rowsPm.forEach(function(r) {
        var arr = safeParse_(r[1], []);
        if (Array.isArray(arr)) {
          arr.forEach(function(item) {
            if (item.name) {
              var trName = traducirTexto(item.name);
              if (trName !== item.name) { item.name = trName; stats.metodos++; changedPm = true; }
            }
            if (item.instructions) {
              var trInst = traducirTexto(item.instructions);
              if (trInst !== item.instructions) { item.instructions = trInst; changedPm = true; }
            }
          });
          r[1] = JSON.stringify(arr);
        }
      });
      if (changedPm) {
        shPm.getRange(2, 2, rowsPm.length, 1).setValues(rowsPm.map(function(r) { return [r[1]]; }));
        jdbDrop_(PM_SHEET);
      }
    }

    // 5. AJUSTES (SETTINGS)
    var shSet = sh_(SETTINGS_SHEET);
    if (shSet && shSet.getLastRow() > 1) {
      var rowsSet = shSet.getRange(2, 1, shSet.getLastRow() - 1, 2).getValues();
      var translatableKeys = [
        'shop_name', 'about_text', 'hero_title', 'hero_subtitle',
        'pause_message', 'delivery_note', 'footer_note'
      ];
      var changedSet = false;
      rowsSet.forEach(function(r) {
        var key = String(r[0] || '').trim();
        if (translatableKeys.indexOf(key) !== -1 && r[1]) {
          var trVal = traducirTexto(String(r[1]));
          if (trVal !== String(r[1])) {
            r[1] = trVal;
            stats.ajustes++;
            changedSet = true;
          }
        } else if ((key === 'sugar_levels' || key === 'ice_levels') && r[1]) {
          var lvls = safeParse_(r[1], null);
          if (Array.isArray(lvls)) {
            var trLvls = lvls.map(function(l) { return traducirTexto(String(l)); });
            r[1] = JSON.stringify(trLvls);
            stats.ajustes++;
            changedSet = true;
          }
        }
      });
      if (changedSet) {
        shSet.getRange(2, 1, rowsSet.length, 2).setValues(rowsSet);
        CACHE_.remove(SET_KEY);
        _m.set = null;
      }
    }

    // 6. ROLES (ROLES)
    var shRoles = sh_(ROLES_SHEET);
    if (shRoles && shRoles.getLastRow() > 1) {
      var rowsRoles = shRoles.getRange(2, 1, shRoles.getLastRow() - 1, ROLE_HEAD.length).getValues();
      var changedRoles = false;
      rowsRoles.forEach(function(r) {
        var label = r[ROLE_C.LABEL];
        if (label) {
          var trLabel = traducirTexto(String(label));
          if (trLabel !== String(label)) {
            r[ROLE_C.LABEL] = trLabel;
            stats.roles++;
            changedRoles = true;
          }
        }
      });
      if (changedRoles) {
        shRoles.getRange(2, 1, rowsRoles.length, ROLE_HEAD.length).setValues(rowsRoles);
        bustRoles_();
      }
    }

    // Purgar caché del frontend para que tome efecto inmediato
    bustFront_();
  });

  ss.toast('¡Traducción completada con éxito!', 'Traductor Tienda', 5);
  ui.alert(
    '¡Traducción de Base de Datos Completada!',
    'Se tradujeron de forma segura:\n' +
    '• Categorías: ' + stats.categorias + '\n' +
    '• Productos: ' + stats.productos + '\n' +
    '• Complementos: ' + stats.addons + '\n' +
    '• Métodos de pago: ' + stats.metodos + '\n' +
    '• Ajustes de tienda: ' + stats.ajustes + '\n' +
    '• Roles: ' + stats.roles + '\n\n' +
    'La base de datos sigue 100% compatible y operativa con la tienda web.',
    ui.ButtonSet.OK
  );
}
