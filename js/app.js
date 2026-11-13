/* =====================================================================
   Tesis de Arquitectura CUJAE — aplicación
   Lee el Excel de datos/registro.xlsx en el navegador y construye el
   buscador, los filtros, las estadísticas y la página "Sobre el registro".
   No hay que tocar este archivo para actualizar los datos.
   ===================================================================== */
(function () {
  "use strict";
  var C = window.CONFIG || {};
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };

  // ---------- utilidades de texto ----------
  function nfc(s) { return (s == null ? "" : String(s)).normalize("NFC"); }
  // minúsculas y sin tildes, conservando la longitud (para poder resaltar)
  function norm(s) {
    var out = "", i, ch, d;
    s = nfc(s);
    for (i = 0; i < s.length; i++) {
      ch = s[i];
      if (s.charCodeAt(i) < 128) { out += ch.toLowerCase(); continue; }
      d = ch.normalize("NFD")[0].toLowerCase();
      out += d.length === 1 ? d : ch.toLowerCase().charAt(0);
    }
    return out;
  }
  function esc(s) { return String(s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
  function plural(n, uno, varios) { return n.toLocaleString("es") + " " + (n === 1 ? uno : varios); }
  function num(n) { return n.toLocaleString("es"); }
  function resaltar(texto, terminos) {
    if (!terminos || !terminos.length) return esc(texto);
    var n = norm(texto), marcas = [], i, p, t;
    for (i = 0; i < terminos.length; i++) {
      t = terminos[i]; p = n.indexOf(t);
      while (p !== -1) { marcas.push([p, p + t.length]); p = n.indexOf(t, p + t.length); }
    }
    if (!marcas.length) return esc(texto);
    marcas.sort(function (a, b) { return a[0] - b[0]; });
    var out = "", pos = 0;
    marcas.forEach(function (m) {
      if (m[1] <= pos) return;
      var a = Math.max(m[0], pos);
      out += esc(texto.slice(pos, a)) + "<mark>" + esc(texto.slice(a, m[1])) + "</mark>"; pos = m[1];
    });
    return out + esc(texto.slice(pos));
  }

  // ---------- lectura del Excel ----------
  var TESIS = [], PERSONAS = [], ANIOS = [], META = { total: 0, excluidas: 0, fecha: "", verif: null };

  function columnas(cab) {
    var h = cab.map(function (c) { return norm(c).replace(/\s+/g, " ").trim(); });
    function buscar(f) { for (var i = 0; i < h.length; i++) if (h[i] && f(h[i])) return i; return -1; }
    var noOriginal = function (x) { return x.indexOf("tal como") === -1; };
    return {
      id: buscar(function (x) { return /^n[.º°o]*$/.test(x) || x === "numero" || x === "#"; }),
      anio: buscar(function (x) { return x === "ano" || x.indexOf("ano ") === 0; }),
      autores: buscar(function (x) { return x.indexOf("autor") === 0 || x.indexOf("diplomante") === 0; }),
      titulo: buscar(function (x) { return x.indexOf("titulo") === 0; }),
      tutores: buscar(function (x) { return x.indexOf("tutor") === 0 && noOriginal(x); }),
      asesores: buscar(function (x) { return x.indexOf("asesor") === 0 && noOriginal(x); }),
      resumen: buscar(function (x) { return x === "resumen"; }),
      archivo: buscar(function (x) { return x.indexOf("archivo") === 0; }),
      pag: buscar(function (x) { return x.indexOf("pag") === 0; }),
      obs: buscar(function (x) { return x.indexOf("observ") === 0; }),
      aviso: buscar(function (x) { return x.indexOf("tipo de aviso") === 0; }),
      verif: buscar(function (x) { return x.indexOf("verific") === 0; }),
      enlace: buscar(function (x) { return x.indexOf("enlace") !== -1 || x.indexOf("link") !== -1 || x === "url"; })
    };
  }
  function nombres(celda) {
    return nfc(celda).split(/\n|;/).map(function (x) { return x.trim(); }).filter(Boolean).map(function (x) {
      var m = x.match(/\s*\(([^)]*)\)\s*$/);
      return { n: m ? x.slice(0, m.index).trim() : x, rol: m ? m[1] : "" };
    });
  }
  function procesar(libro) {
    var hoja = libro.Sheets[C.hoja] || libro.Sheets[libro.SheetNames[0]];
    if (!hoja) throw new Error("El Excel no tiene ninguna hoja.");
    var filas = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: "", raw: true });
    if (filas.length < 2) throw new Error("La hoja «" + (C.hoja || libro.SheetNames[0]) + "» está vacía.");
    var col = columnas(filas[0]), faltan = [];
    if (col.titulo < 0) faltan.push("Título"); if (col.autores < 0) faltan.push("Autor(es)"); if (col.anio < 0) faltan.push("Año");
    if (faltan.length) throw new Error("En la primera fila de la hoja no encuentro estas columnas: " + faltan.join(", ") + ". Los encabezados deben conservar esos nombres.");
    var excl = (C.excluirAvisos || []).map(norm), v = { auto: 0, imagen: 0, corregido: 0, cotejar: 0, total: 0 };
    var cel = function (f, i) { return i < 0 ? "" : nfc(f[i]).trim(); };
    TESIS = []; META.total = 0; META.excluidas = 0;
    filas.slice(1).forEach(function (f, k) {
      var titulo = cel(f, col.titulo), autores = nombres(cel(f, col.autores));
      if (!titulo && !autores.length && !cel(f, col.archivo)) return;
      META.total++;
      var aviso = cel(f, col.aviso), an = norm(aviso);
      if (excl.some(function (e) { return e && an.indexOf(e) !== -1; })) { META.excluidas++; return; }
      var anio = parseInt(cel(f, col.anio), 10);
      var t = {
        id: cel(f, col.id) || String(k + 1), anio: isNaN(anio) ? null : anio, titulo: titulo || "(Sin título)",
        autores: autores.map(function (a) { return a.n; }), tutores: nombres(cel(f, col.tutores)), asesores: nombres(cel(f, col.asesores)),
        resumen: cel(f, col.resumen), archivo: cel(f, col.archivo), pag: cel(f, col.pag), obs: cel(f, col.obs), aviso: aviso,
        verif: cel(f, col.verif), enlace: cel(f, col.enlace)
      };
      t._t = norm(t.titulo); t._a = norm(t.autores.join(" | "));
      t._tu = norm(t.tutores.concat(t.asesores).map(function (x) { return x.n; }).join(" | "));
      t._r = norm(t.resumen); t._todo = t._t + " \n " + t._a + " \n " + t._tu + " \n " + t._r;
      if (t.resumen && t.verif) {
        var vn = norm(t.verif); v.total++;
        if (vn.indexOf("cotejar") !== -1) v.cotejar++;
        else if (vn.indexOf("automaticamente") !== -1) v.auto++;
        else if (vn.indexOf("corrig") !== -1) v.corregido++;
        else v.imagen++;
      }
      TESIS.push(t);
    });
    META.verif = col.verif >= 0 && v.total ? v : null;
    var ys = TESIS.map(function (t) { return t.anio; }).filter(Boolean);
    ANIOS = []; if (ys.length) for (var y = Math.min.apply(null, ys); y <= Math.max.apply(null, ys); y++) ANIOS.push(y);
    PERSONAS = personas(TESIS);
  }
  function personas(lista) {
    var m = {};
    lista.forEach(function (t) {
      function suma(arr, campo) {
        var vistos = {};
        arr.forEach(function (p) {
          if (vistos[p.n]) return; vistos[p.n] = 1;
          var o = m[p.n] || (m[p.n] = { n: p.n, tutor: 0, asesor: 0, min: null, max: null, _n: norm(p.n) });
          o[campo]++;
          if (t.anio) { o.min = o.min === null ? t.anio : Math.min(o.min, t.anio); o.max = o.max === null ? t.anio : Math.max(o.max, t.anio); }
        });
      }
      suma(t.tutores, "tutor"); suma(t.asesores, "asesor");
    });
    return Object.keys(m).map(function (k) { m[k].total = m[k].tutor + m[k].asesor; return m[k]; })
      .sort(function (a, b) { return b.tutor - a.tutor || b.total - a.total || a.n.localeCompare(b.n, "es"); });
  }

  // ---------- estado y dirección (hash) ----------
  var E = { vista: "registro", q: "", ambito: "todo", desde: "", hasta: "", tutor: "", solo: false, autoria: "", resumen: "", orden: "reciente", visibles: 0, base: "todas" };
  var ordenManual = false;
  function leerHash() {
    var h = location.hash.replace(/^#\/?/, ""), p = h.split("?"), q = new URLSearchParams(p[1] || "");
    E.vista = ["registro", "estadisticas", "acerca"].indexOf(p[0]) !== -1 ? p[0] : "registro";
    E.q = q.get("q") || ""; E.ambito = q.get("en") || "todo"; E.desde = q.get("desde") || ""; E.hasta = q.get("hasta") || "";
    E.tutor = q.get("tutor") || ""; E.solo = q.get("solo") === "1"; E.autoria = q.get("autoria") || ""; E.resumen = q.get("resumen") || "";
    E.orden = q.get("orden") || (E.q ? "relevancia" : "reciente"); ordenManual = !!q.get("orden"); E.base = q.get("base") || "todas";
  }
  function hash(vista) {
    var q = new URLSearchParams();
    if (E.q) q.set("q", E.q); if (E.ambito !== "todo") q.set("en", E.ambito);
    if (E.desde) q.set("desde", E.desde); if (E.hasta) q.set("hasta", E.hasta);
    if (E.tutor) q.set("tutor", E.tutor); if (E.solo) q.set("solo", "1");
    if (E.autoria) q.set("autoria", E.autoria); if (E.resumen) q.set("resumen", E.resumen);
    if (ordenManual) q.set("orden", E.orden); if (E.base !== "todas" && (vista || E.vista) === "estadisticas") q.set("base", E.base);
    var s = q.toString();
    return "#/" + (vista || E.vista) + (s ? "?" + s : "");
  }
  function ir(vista) { var h = hash(vista); if (location.hash === h) pintar(); else location.hash = h; }

  // ---------- búsqueda y filtros ----------
  function terminos(q) {
    var out = [], re = /"([^"]+)"|(\S+)/g, m;
    while ((m = re.exec(norm(q)))) { var t = (m[1] || m[2]).replace(/^[.,;:]+|[.,;:]+$/g, ""); if (t) out.push(t); }
    return out;
  }
  function filtrar(sinAnios) {
    var ts = terminos(E.q), campo = { todo: "_todo", titulo: "_t", autor: "_a", tutor: "_tu", resumen: "_r" }[E.ambito] || "_todo";
    var d = parseInt(E.desde, 10), h = parseInt(E.hasta, 10);
    return TESIS.filter(function (t) {
      var i;
      for (i = 0; i < ts.length; i++) if (t[campo].indexOf(ts[i]) === -1) return false;
      if (!sinAnios) { if (!isNaN(d) && (!t.anio || t.anio < d)) return false; if (!isNaN(h) && (!t.anio || t.anio > h)) return false; }
      if (E.tutor) {
        var es = function (p) { return p.n === E.tutor; };
        if (!(t.tutores.some(es) || (!E.solo && t.asesores.some(es)))) return false;
      }
      if (E.autoria === "1" && t.autores.length !== 1) return false;
      if (E.autoria === "2" && t.autores.length < 2) return false;
      if (E.resumen === "si" && !t.resumen) return false;
      if (E.resumen === "no" && t.resumen) return false;
      return true;
    });
  }
  function puntos(t, ts) {
    var p = 0;
    ts.forEach(function (x) {
      if (t._t.indexOf(x) !== -1) p += 6; if (t._a.indexOf(x) !== -1) p += 5; if (t._tu.indexOf(x) !== -1) p += 3;
      var c = t._r.split(x).length - 1; p += Math.min(c, 4);
    });
    return p;
  }
  function ordenar(lista) {
    var ts = terminos(E.q), cmp = function (a, b) { return a.localeCompare(b, "es", { sensitivity: "base" }); };
    var porAnio = function (a, b) { return (b.anio || 0) - (a.anio || 0) || cmp(a.autores[0] || "", b.autores[0] || ""); };
    var f = {
      reciente: porAnio,
      antiguo: function (a, b) { return (a.anio || 9999) - (b.anio || 9999) || cmp(a.autores[0] || "", b.autores[0] || ""); },
      titulo: function (a, b) { return cmp(a.titulo.replace(/^[^A-Za-zÀ-ÿ0-9]+/, ""), b.titulo.replace(/^[^A-Za-zÀ-ÿ0-9]+/, "")); },
      autor: function (a, b) { return cmp(a.autores[0] || "~", b.autores[0] || "~") || porAnio(a, b); },
      relevancia: function (a, b) { return puntos(b, ts) - puntos(a, ts) || porAnio(a, b); }
    }[E.orden] || porAnio;
    return lista.slice().sort(f);
  }

  // ---------- perfil de años (cabecera) ----------
  var arrastre = null;
  function pintarPerfil() {
    var cont = $("#perfil-barras"), eje = $("#perfil-eje"), base = filtrar(true), cuenta = {};
    base.forEach(function (t) { if (t.anio) cuenta[t.anio] = (cuenta[t.anio] || 0) + 1; });
    var max = Math.max.apply(null, ANIOS.map(function (y) { return cuenta[y] || 0; }).concat([1]));
    var d = parseInt(E.desde, 10), h = parseInt(E.hasta, 10), hay = !isNaN(d) || !isNaN(h);
    if (isNaN(d)) d = ANIOS[0]; if (isNaN(h)) h = ANIOS[ANIOS.length - 1];
    cont.innerHTML = ANIOS.map(function (y, i) {
      var n = cuenta[y] || 0, cls = "b" + (n ? "" : " vacia") + (hay ? (y >= d && y <= h ? " sel" : "") : " dentro");
      return '<div class="' + cls + '" data-i="' + i + '" data-y="' + y + '" data-n="' + n + '" style="height:' + (n ? Math.max(4, Math.round(n / max * 100)) : 0) + '%"></div>';
    }).join("");
    eje.innerHTML = ANIOS.map(function (y, i) { return "<span>" + (i === 0 || i === ANIOS.length - 1 || (y % 5 === 0 && i > 1 && i < ANIOS.length - 2) ? y : "") + "</span>"; }).join("");
    $("#perfil-ayuda").innerHTML = hay
      ? "Período elegido: " + (d === h ? d : d + "–" + h) + '. <button type="button" id="perfil-todo">Ver todos los años</button>'
      : "Tesis por año" + (E.q || E.tutor || E.autoria || E.resumen ? " para esta búsqueda" : "") + ". Pulsa un año o arrastra sobre las barras para elegir un período.";
    var b = $("#perfil-todo"); if (b) b.onclick = function () { E.desde = ""; E.hasta = ""; ir(); };
  }
  function iniciarPerfil() {
    var cont = $("#perfil-barras");
    function idx(ev) {
      var r = cont.getBoundingClientRect(), i = Math.floor((ev.clientX - r.left) / r.width * ANIOS.length);
      return Math.max(0, Math.min(ANIOS.length - 1, i));
    }
    function marcar(a, b) {
      var lo = Math.min(a, b), hi = Math.max(a, b);
      $$(".b", cont).forEach(function (el, i) { el.classList.toggle("sel", i >= lo && i <= hi); el.classList.remove("dentro"); });
    }
    cont.addEventListener("pointerdown", function (ev) { arrastre = { a: idx(ev), b: idx(ev) }; cont.setPointerCapture(ev.pointerId); marcar(arrastre.a, arrastre.b); ev.preventDefault(); });
    cont.addEventListener("pointermove", function (ev) {
      var i = idx(ev);
      if (arrastre) { arrastre.b = i; marcar(arrastre.a, arrastre.b); globo(ANIOS[Math.min(arrastre.a, i)] + "–" + ANIOS[Math.max(arrastre.a, i)], ev); }
      else { var el = cont.children[i]; if (el) globo("<b>" + el.dataset.y + "</b>: " + plural(+el.dataset.n, "tesis", "tesis"), ev); }
    });
    cont.addEventListener("pointerup", function () {
      if (!arrastre) return;
      var lo = Math.min(arrastre.a, arrastre.b), hi = Math.max(arrastre.a, arrastre.b); arrastre = null; globo();
      E.desde = String(ANIOS[lo]); E.hasta = String(ANIOS[hi]); ir("registro");
    });
    cont.addEventListener("pointercancel", function () { arrastre = null; pintarPerfil(); });
    cont.addEventListener("pointerleave", function () { if (!arrastre) globo(); });
  }
  function globo(html, ev) {
    var g = $("#globo");
    if (!html) { g.hidden = true; return; }
    g.innerHTML = html; g.hidden = false;
    g.style.left = Math.max(70, Math.min(window.innerWidth - 70, ev.clientX)) + "px"; g.style.top = (ev.clientY - 14) + "px";
  }

  // ---------- vista: registro ----------
  var actual = [];
  function pintarRegistro() {
    actual = ordenar(filtrar(false));
    var ts = terminos(E.q), n = actual.length;
    $("#conteo").innerHTML = n === TESIS.length
      ? plural(n, "tesis", "tesis") + " <small>en el registro</small>"
      : plural(n, "tesis", "tesis") + " <small>de " + num(TESIS.length) + "</small>";
    pintarActivos();
    if (!E.visibles) E.visibles = C.porPagina || 30;
    var lista = $("#lista");
    if (!n) {
      lista.innerHTML = '<li class="vacio" style="display:block"><h3>Ninguna tesis coincide con esta búsqueda</h3><p>Prueba con menos palabras, revisa la ortografía o quita algún filtro. La búsqueda no distingue mayúsculas ni tildes.</p></li>';
    } else {
      lista.innerHTML = actual.slice(0, E.visibles).map(function (t) { return item(t, ts); }).join("");
    }
    var mas = $("#mas"); mas.hidden = n <= E.visibles;
    if (!mas.hidden) mas.textContent = "Mostrar " + Math.min(C.porPagina || 30, n - E.visibles) + " más (quedan " + num(n - E.visibles) + ")";
  }
  function fragmento(t, ts) {
    if (!ts.length || !t.resumen || (E.ambito !== "todo" && E.ambito !== "resumen")) return "";
    var p = -1, i;
    for (i = 0; i < ts.length && p === -1; i++) p = t._r.indexOf(ts[i]);
    if (p === -1) return "";
    var a = Math.max(0, p - 110), b = Math.min(t.resumen.length, p + 190);
    if (a > 0) { var s = t.resumen.indexOf(" ", a); if (s !== -1 && s < p) a = s + 1; }
    if (b < t.resumen.length) { var e = t.resumen.lastIndexOf(" ", b); if (e > p) b = e; }
    return '<p class="extracto">' + (a > 0 ? "… " : "") + resaltar(t.resumen.slice(a, b).replace(/\n/g, " "), ts) + (b < t.resumen.length ? " …" : "") + "</p>";
  }
  function persona(p, ts) {
    return '<button type="button" data-tutor="' + esc(p.n) + '" title="Ver las tesis de ' + esc(p.n) + '">' + resaltar(p.n, E.ambito === "todo" || E.ambito === "tutor" ? ts : []) + "</button>" + (p.rol ? " (" + esc(p.rol) + ")" : "");
  }
  function item(t, ts) {
    var enT = E.ambito === "todo" || E.ambito === "titulo" ? ts : [], enA = E.ambito === "todo" || E.ambito === "autor" ? ts : [];
    var h = '<li data-id="' + esc(t.id) + '"><div class="anio' + (t.anio ? "" : " sin") + '">' + (t.anio || "s. f.") + '</div><div class="tesis">';
    h += '<h3><button type="button" class="titulo" aria-expanded="false">' + resaltar(t.titulo, enT) + "</button></h3>";
    h += '<p class="autores">' + (t.autores.length ? t.autores.map(function (a) { return resaltar(a, enA); }).join("; ") : "Autor no identificado") + "</p>";
    if (t.tutores.length) h += '<p class="tutores">' + (t.tutores.length > 1 ? "Tutores: " : "Tutor: ") + t.tutores.map(function (p) { return persona(p, ts); }).join("; ") + "</p>";
    if (t.asesores.length) h += '<p class="tutores">' + (t.asesores.length > 1 ? "Asesores: " : "Asesor: ") + t.asesores.map(function (p) { return persona(p, ts); }).join("; ") + "</p>";
    h += fragmento(t, ts);
    h += '<p class="abrir"><button type="button" class="boton-plano ver">' + (t.resumen ? "Leer el resumen" : "Ver la ficha") + "</button></p>";
    return h + "</div></li>";
  }
  function detalle(t) {
    var ts = E.ambito === "todo" || E.ambito === "resumen" ? terminos(E.q) : [];
    var h = '<div class="detalle">';
    h += t.resumen
      ? '<h4>Resumen</h4><div class="resumen">' + resaltar(t.resumen, ts) + "</div>"
      : '<p class="sin-resumen">Esta tesis no tiene un apartado de resumen en su documento, así que el registro lo deja en blanco.</p>';
    h += '<dl class="datos-ficha">';
    h += "<dt>N.º en el registro</dt><dd>" + esc(t.id) + "</dd>";
    if (t.archivo) h += "<dt>Archivo</dt><dd>" + esc(t.archivo) + (t.pag ? " (resumen en la pág. " + esc(t.pag) + ")" : "") + "</dd>";
    if (t.enlace && /^https?:\/\//i.test(t.enlace)) h += '<dt>Documento</dt><dd><a href="' + esc(t.enlace) + '" rel="noopener" target="_blank">Abrir el PDF</a></dd>';
    if (t.resumen && t.verif) h += "<dt>Comprobación del resumen</dt><dd>" + esc(t.verif) + "</dd>";
    if (C.mostrarObservaciones && t.obs) h += "<dt>Notas del registro</dt><dd>" + esc(t.obs) + "</dd>";
    return h + "</dl></div>";
  }
  function pintarActivos() {
    var f = [];
    if (E.q) f.push(["q", "Texto", "«" + E.q + "»" + (E.ambito !== "todo" ? " en " + { titulo: "título", autor: "autor", tutor: "tutor", resumen: "resumen" }[E.ambito] : "")]);
    if (E.desde || E.hasta) f.push(["anios", "Años", (E.desde || ANIOS[0]) + "–" + (E.hasta || ANIOS[ANIOS.length - 1])]);
    if (E.tutor) f.push(["tutor", E.solo ? "Tutor" : "Tutor o asesor", E.tutor]);
    if (E.autoria) f.push(["autoria", "Autoría", E.autoria === "1" ? "un solo autor" : "en equipo"]);
    if (E.resumen) f.push(["resumen", "Resumen", E.resumen === "si" ? "con resumen" : "sin resumen"]);
    $("#activos").innerHTML = f.map(function (x) {
      return '<button type="button" class="ficha-filtro" data-quitar="' + x[0] + '" title="Quitar este filtro">' + esc(x[1]) + ": <b>" + esc(x[2]) + '</b><i aria-hidden="true">×</i><span class="oculto"> (quitar)</span></button>';
    }).join("");
  }
  function quitar(k) {
    if (k === "q") { E.q = ""; E.ambito = "todo"; if (!ordenManual) E.orden = "reciente"; }
    if (k === "anios") { E.desde = ""; E.hasta = ""; }
    if (k === "tutor") { E.tutor = ""; E.solo = false; }
    if (k === "autoria") E.autoria = ""; if (k === "resumen") E.resumen = "";
    E.visibles = 0; ir("registro");
  }
  function sincronizarControles() {
    $("#q").value = E.q;
    $$('input[name="ambito"]').forEach(function (r) { r.checked = r.value === E.ambito; });
    $("#f-desde").value = E.desde; $("#f-hasta").value = E.hasta;
    $("#f-tutor").value = E.tutor; $("#f-solo-tutor").checked = E.solo;
    $$('input[name="f-autoria"]').forEach(function (r) { r.checked = r.value === E.autoria; });
    $$('input[name="f-resumen"]').forEach(function (r) { r.checked = r.value === E.resumen; });
    $("#orden").value = E.orden; $("#est-base").value = E.base;
  }
  function descargarCSV() {
    var q = function (s) { return '"' + String(s == null ? "" : s).replace(/"/g, '""') + '"'; };
    var filas = [["N.º", "Año", "Autor(es)", "Título", "Tutor(es)", "Asesores", "Resumen"].map(q).join(";")];
    var nom = function (a) { return a.map(function (p) { return p.n + (p.rol ? " (" + p.rol + ")" : ""); }).join(" / "); };
    actual.forEach(function (t) { filas.push([t.id, t.anio || "", t.autores.join(" / "), t.titulo, nom(t.tutores), nom(t.asesores), t.resumen].map(q).join(";")); });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob(["﻿" + filas.join("\r\n")], { type: "text/csv;charset=utf-8" }));
    a.download = "tesis-arquitectura-cujae.csv"; document.body.appendChild(a); a.click(); a.remove();
  }

  // ---------- combo de tutores ----------
  function iniciarCombo() {
    var inp = $("#f-tutor"), ul = $("#f-tutor-lista"), activo = -1, ops = [];
    function abrir() {
      var t = norm(inp.value.trim());
      ops = PERSONAS.filter(function (p) { return !t || p._n.indexOf(t) !== -1; }).slice(0, 60);
      ul.innerHTML = ops.length ? ops.map(function (p, i) { return '<li role="option" id="op-' + i + '" data-i="' + i + '">' + esc(p.n) + "<small>" + p.total + "</small></li>"; }).join("") : '<li aria-disabled="true">Ningún nombre coincide</li>';
      ul.hidden = false; inp.setAttribute("aria-expanded", "true"); activo = -1;
    }
    function cerrar() { ul.hidden = true; inp.setAttribute("aria-expanded", "false"); }
    function elegir(i) { if (!ops[i]) return; E.tutor = ops[i].n; E.visibles = 0; cerrar(); ir("registro"); }
    function mover(d) {
      if (ul.hidden) abrir();
      activo = Math.max(0, Math.min(ops.length - 1, activo + d));
      $$("li", ul).forEach(function (li, i) { li.setAttribute("aria-selected", i === activo ? "true" : "false"); if (i === activo) li.scrollIntoView({ block: "nearest" }); });
      inp.setAttribute("aria-activedescendant", "op-" + activo);
    }
    inp.addEventListener("focus", abrir); inp.addEventListener("input", abrir);
    inp.addEventListener("keydown", function (ev) {
      if (ev.key === "ArrowDown") { mover(1); ev.preventDefault(); }
      else if (ev.key === "ArrowUp") { mover(-1); ev.preventDefault(); }
      else if (ev.key === "Enter") { elegir(activo < 0 ? 0 : activo); ev.preventDefault(); }
      else if (ev.key === "Escape") { cerrar(); inp.value = E.tutor; }
    });
    inp.addEventListener("blur", function () { setTimeout(function () { cerrar(); if (!inp.value.trim() && E.tutor) quitar("tutor"); else inp.value = E.tutor; }, 160); });
    ul.addEventListener("mousedown", function (ev) { var li = ev.target.closest("li[data-i]"); if (li) { ev.preventDefault(); elegir(+li.dataset.i); } });
  }

  // ---------- vista: estadísticas ----------
  var ordenTabla = { k: "tutor", asc: false }, tablaCompleta = false;
  var VACIAS = "de la del las los el en y a para por con un una su sus al se que como entre desde sobre e o u the of and in no es lo sin ante bajo hacia hasta segun caso estudio".split(" ");
  function barras(cont, datos, opc) {
    opc = opc || {};
    var max = Math.max.apply(null, datos.map(function (d) { return d.v; }).concat([1]));
    cont.innerHTML = datos.map(function (d) {
      var nombre = opc.accion ? '<button type="button" class="nombre" data-' + opc.accion + '="' + esc(d.k) + '" title="' + esc(d.n) + '">' + esc(d.n) + "</button>" : '<span class="nombre" title="' + esc(d.n) + '">' + esc(d.n) + "</span>";
      return '<div class="fila-barra">' + nombre + '<span class="pista"><i style="width:' + (d.v / max * 100).toFixed(1) + '%"></i></span><span class="valor">' + num(d.v) + "</span></div>";
    }).join("") || '<p class="nota">Sin datos.</p>';
  }
  function pintarEstadisticas() {
    var base = E.base === "busqueda" ? filtrar(false) : TESIS, n = base.length;
    var hayFiltro = E.q || E.desde || E.hasta || E.tutor || E.autoria || E.resumen;
    $("#est-nota").textContent = E.base === "busqueda"
      ? (hayFiltro ? "Calculado sobre las " + num(n) + " tesis que cumplen la búsqueda y los filtros actuales del registro." : "No hay ninguna búsqueda ni filtro activo: se muestran todas las tesis.")
      : "";
    var per = personas(base), conTutor = per.filter(function (p) { return p.tutor > 0; });
    var ys = base.map(function (t) { return t.anio; }).filter(Boolean), aut = {};
    base.forEach(function (t) { t.autores.forEach(function (a) { aut[norm(a)] = 1; }); });
    var conRes = base.filter(function (t) { return t.resumen; }).length;
    var cifras = [
      [num(n), "tesis"],
      [ys.length ? Math.min.apply(null, ys) + "–" + Math.max.apply(null, ys) : "—", "años que abarca"],
      [num(Object.keys(aut).length), "autores"],
      [num(conTutor.length), "tutores distintos"],
      [n ? Math.round(conRes / n * 100) + " %" : "—", "tienen resumen (" + num(conRes) + ")"]
    ];
    $("#est-cifras").innerHTML = cifras.map(function (c) { return '<div class="cifra"><b>' + c[0] + "</b><span>" + c[1] + "</span></div>"; }).join("");

    // por año
    var cuenta = {}; base.forEach(function (t) { if (t.anio) cuenta[t.anio] = (cuenta[t.anio] || 0) + 1; });
    var max = Math.max.apply(null, ANIOS.map(function (y) { return cuenta[y] || 0; }).concat([1]));
    var paso = max <= 10 ? 2 : max <= 30 ? 10 : max <= 60 ? 20 : 25, tope = Math.ceil(max / paso) * paso, guias = "", ejeY = "";
    for (var g = 0; g <= tope; g += paso) { guias += '<div class="ga-guia" style="bottom:' + (g / tope * 100) + '%"></div>'; ejeY += '<span style="bottom:' + (g / tope * 100) + '%">' + g + "</span>"; }
    $("#est-anios").innerHTML = '<div class="ga-lienzo"><div class="ga-eje-y">' + ejeY + '</div><div class="ga-plot">' + guias + '<div class="ga-barras">' +
      ANIOS.map(function (y) {
        var v = cuenta[y] || 0;
        return '<button type="button" data-anio="' + y + '" data-n="' + v + '" style="height:' + (v / tope * 100) + '%" aria-label="' + y + ": " + v + ' tesis">' + (v === max ? "<em>" + v + "</em>" : "") + "</button>";
      }).join("") + '</div></div><div class="ga-eje-x">' +
      ANIOS.map(function (y, i) { return "<span>" + (i === 0 || i === ANIOS.length - 1 || (y % 5 === 0 && i > 1 && i < ANIOS.length - 2) ? y : "") + "</span>"; }).join("") + "</div></div>";
    $("#est-anios-tabla").innerHTML = '<div class="mini-tabla">' + ANIOS.map(function (y) { return "<div><span>" + y + "</span><b>" + (cuenta[y] || 0) + "</b></div>"; }).join("") + "</div>";

    // tutores
    barras($("#est-tutores-top"), conTutor.slice(0, 20).map(function (p) { return { k: p.n, n: p.n, v: p.tutor }; }), { accion: "tutor" });
    pintarTablaTutores(per);

    // autoría y nº de tutores
    var da = { 1: 0, 2: 0, 3: 0, 4: 0 }, dt = { 0: 0, 1: 0, 2: 0, 3: 0, 4: 0 };
    base.forEach(function (t) { if (t.autores.length) da[Math.min(4, t.autores.length)]++; dt[Math.min(4, t.tutores.length)]++; });
    barras($("#est-autoria"), [{ n: "Un autor", v: da[1] }, { n: "Dos autores", v: da[2] }, { n: "Tres autores", v: da[3] }, { n: "Cuatro o más", v: da[4] }]);
    barras($("#est-ntutores"), [{ n: "Un tutor", v: dt[1] }, { n: "Dos tutores", v: dt[2] }, { n: "Tres tutores", v: dt[3] }, { n: "Cuatro o más", v: dt[4] }, { n: "Sin tutor identificado", v: dt[0] }]);

    // palabras de los títulos
    var pal = {}, forma = {};
    base.forEach(function (t) {
      var vistos = {};
      (t.titulo.match(/[A-Za-zÀ-ÿ]{4,}/g) || []).forEach(function (w) {
        var k = norm(w); if (VACIAS.indexOf(k) !== -1 || vistos[k]) return; vistos[k] = 1;
        pal[k] = (pal[k] || 0) + 1; var f = w.toLowerCase(); (forma[k] = forma[k] || {})[f] = (forma[k][f] || 0) + 1;
      });
    });
    var top = Object.keys(pal).sort(function (a, b) { return pal[b] - pal[a] || a.localeCompare(b); }).slice(0, 30).map(function (k) {
      var f = Object.keys(forma[k]).sort(function (a, b) { return forma[k][b] - forma[k][a]; })[0];
      return { k: f, n: f, v: pal[k] };
    });
    barras($("#est-palabras"), top, { accion: "palabra" });

    // períodos
    var cub = {};
    base.forEach(function (t) {
      if (!t.anio) return;
      var ini = Math.floor((t.anio - 1) / 5) * 5 + 1, c = cub[ini] || (cub[ini] = { n: 0, p: {} }); c.n++;
      t.tutores.forEach(function (p) { c.p[p.n] = (c.p[p.n] || 0) + 1; });
    });
    var ultimo = ANIOS[ANIOS.length - 1];
    $("#est-periodos").innerHTML = Object.keys(cub).sort().map(function (ini) {
      var c = cub[ini], fin = Math.min(+ini + 4, ultimo), desde = Math.max(+ini, ANIOS[0]);
      var tops = Object.keys(c.p).sort(function (a, b) { return c.p[b] - c.p[a] || a.localeCompare(b, "es"); }).slice(0, 5);
      return "<div><h3>" + (desde === fin ? fin : desde + "–" + fin) + " <small>(" + plural(c.n, "tesis", "tesis") + ")</small></h3><ol>" +
        tops.map(function (k) { return "<li>" + esc(k) + "<span>" + c.p[k] + "</span></li>"; }).join("") + "</ol></div>";
    }).join("");
  }
  function pintarTablaTutores(per) {
    per = per || personas(E.base === "busqueda" ? filtrar(false) : TESIS);
    var t = norm($("#est-tutor-buscar").value.trim());
    var filas = per.filter(function (p) { return !t || p._n.indexOf(t) !== -1; });
    var k = ordenTabla.k, s = ordenTabla.asc ? 1 : -1;
    filas.sort(function (a, b) {
      if (k === "n") return s * a.n.localeCompare(b.n, "es");
      return s * ((a[k] || 0) - (b[k] || 0)) || b.total - a.total || a.n.localeCompare(b.n, "es");
    });
    var total = filas.length, lim = tablaCompleta || t ? total : 25;
    var th = function (key, txt, cls) { return '<th scope="col"' + (cls ? ' class="' + cls + '"' : "") + (ordenTabla.k === key ? ' aria-sort="' + (ordenTabla.asc ? "ascending" : "descending") + '"' : "") + '><button type="button" data-orden="' + key + '">' + txt + "</button></th>"; };
    $("#est-tutores-tabla").innerHTML = "<thead><tr>" + th("n", "Nombre") + th("tutor", "Como tutor", "n") + th("asesor", "Como asesor", "n") + th("total", "Total", "n") + th("min", "Primera tesis", "n") + th("max", "Última tesis", "n") + "</tr></thead><tbody>" +
      filas.slice(0, lim).map(function (p) {
        return '<tr><td><button type="button" data-tutor="' + esc(p.n) + '">' + esc(p.n) + '</button></td><td class="n">' + p.tutor + '</td><td class="n">' + p.asesor + '</td><td class="n">' + p.total + '</td><td class="n">' + (p.min || "—") + '</td><td class="n">' + (p.max || "—") + "</td></tr>";
      }).join("") + "</tbody>";
    var b = $("#est-tutores-mas"); b.hidden = lim >= total; b.textContent = "Ver la tabla completa (" + num(total) + " personas)";
  }

  // ---------- vista: sobre el registro ----------
  function pintarAcerca() {
    var n = TESIS.length, conRes = TESIS.filter(function (t) { return t.resumen; }).length, v = META.verif;
    var ys = ANIOS.length ? ANIOS[0] + " y " + ANIOS[ANIOS.length - 1] : "";
    var avisos = function (txt) { txt = norm(txt); return TESIS.filter(function (t) { return norm(t.aviso).indexOf(txt) !== -1; }).length; };
    var h = "";
    h += '<p class="entrada">Este sitio reúne en un solo lugar los datos básicos de ' + num(n) + " tesis de la Facultad de Arquitectura de la CUJAE defendidas entre " + ys + ": año, autores, título, tutores y, cuando el documento lo trae, su resumen. Sirve para saber qué se ha investigado y proyectado en la facultad, quién lo hizo y con quién.</p>";
    h += "<p>El registro se construyó a partir de los documentos en PDF de las tesis. No es una lista oficial de la facultad: contiene las tesis cuyo PDF estaba disponible, así que puede faltar alguna de esos años. Los PDF no se publican aquí; cada ficha indica el nombre del archivo de origen.</p>";

    h += "<h2>Cómo usarlo</h2><ul>";
    h += "<li><strong>Buscar.</strong> Escribe una o varias palabras; aparecen las tesis que las contienen todas. La búsqueda no distingue mayúsculas ni tildes. Para buscar una frase exacta, ponla entre comillas: «\"centro histórico\"».</li>";
    h += "<li><strong>Acotar.</strong> Elige si buscas en todo, solo en títulos, autores, tutores o resúmenes. A la izquierda puedes filtrar por años, por tutor, por tesis individuales o en equipo, y por si tienen resumen.</li>";
    h += "<li><strong>Años.</strong> Las barras de la cabecera muestran cuántas tesis hay cada año; pulsa una o arrastra sobre varias para quedarte con un período.</li>";
    h += "<li><strong>Compartir.</strong> La dirección de la página guarda la búsqueda y los filtros: cópiala para enviar a alguien exactamente la misma lista. También puedes descargarla en un archivo CSV que abre Excel.</li></ul>";

    h += "<h2>Cómo se hizo</h2>";
    h += "<p>Cada tesis se leyó de su PDF. Los datos de la portada (título, autores, tutores, año) se tomaron de la portada o, si faltaban, de la declaración de autoría. El resumen se copió tal cual aparece en el apartado «Resumen» del documento.</p>";
    h += "<p>La lectura de los PDF y la transcripción se hicieron con programas que extraen el texto de los documentos y con ayuda de un asistente de inteligencia artificial. Por eso se añadieron las comprobaciones que se describen más abajo: ningún resumen se dio por bueno sin compararlo con su origen.</p>";
    h += "<h3>Los resúmenes son literales</h3><p>No se redactó, acortó ni reescribió ningún resumen. Los únicos cambios son mecánicos: unir los renglones, unir las palabras cortadas con guion a fin de renglón y reponer letras que algunos PDF guardan como un solo signo (fi, fl, ti). Si el original tiene una errata, la errata se conserva.</p>";
    h += "<p>Cuando la tesis no tiene un apartado llamado «Resumen», el campo queda vacío. No se tomaron como resumen las introducciones.</p>";
    h += "<h3>Tesis escaneadas</h3><p>Algunos PDF son imágenes escaneadas o guardan el texto de forma ilegible. En esos casos el resumen se obtuvo por reconocimiento óptico de caracteres (OCR) y después se comparó con la imagen de la página.</p>";
    h += "<h3>Títulos, nombres y años</h3><ul>";
    h += "<li><strong>Títulos.</strong> Se transcribieron de la portada corrigiendo tildes y ortografía; los nombres propios se dejaron como están.</li>";
    h += "<li><strong>Autores.</strong> Figuran como en la portada, sin grados académicos.</li>";
    h += "<li><strong>Tutores y asesores.</strong> Una misma persona aparece escrita de muchas maneras en distintas tesis (con o sin segundo apellido, con iniciales, con erratas). Para poder contar, cada persona se unificó bajo un solo nombre. Los cotutores se cuentan como tutores; los asesores y consultantes van aparte.</li>";
    h += "<li><strong>Año.</strong> Es el de la portada. Un curso «2005-2006» se anota como 2006.</li></ul>";

    h += "<h2>Qué se comprobó</h2>";
    if (v) {
      h += "<p>Cada resumen pasó por una de dos comprobaciones. La ficha de cada tesis dice cuál le tocó.</p>";
      h += '<table class="cuadro"><tbody>';
      h += "<tr><td>Resúmenes en el registro</td><td>" + num(conRes) + "</td></tr>";
      h += "<tr><td>Comparados automáticamente, carácter por carácter, con el texto interno del PDF: idénticos</td><td>" + num(v.auto) + "</td></tr>";
      h += "<tr><td>Comparados palabra por palabra con la imagen de la página en una lectura independiente: idénticos</td><td>" + num(v.imagen) + "</td></tr>";
      h += "<tr><td>Comparados con la imagen y corregidos (erratas de lectura, palabras pegadas, signos)</td><td>" + num(v.corregido) + "</td></tr>";
      if (v.cotejar) h += "<tr><td>No se pudieron comprobar enteros, porque el PDF no muestra algunas palabras en la página</td><td>" + num(v.cotejar) + "</td></tr>";
      h += "</tbody></table>";
    } else {
      h += "<p>Los resúmenes se compararon con el texto del PDF o con la imagen de la página de la que proceden.</p>";
    }
    h += "<p>Los títulos, autores y tutores se contrastaron además con un registro anterior y con las listas oficiales de defensas de los años en que existen.</p>";

    h += "<h2>Errores que puede tener</h2><p>El registro se hizo con cuidado, pero conviene leerlo sabiendo dónde puede fallar.</p><ul>";
    h += "<li><strong>Erratas en resúmenes leídos por OCR.</strong> Aunque se revisaron contra la imagen, puede quedar alguna letra o signo mal leído en las tesis escaneadas.</li>";
    h += "<li><strong>Tesis sin resumen.</strong> " + num(n - conRes) + " tesis no tienen resumen aquí porque su documento no trae ese apartado" + (avisos("incompleto") ? ", o porque el PDF está incompleto" : "") + ".</li>";
    h += "<li><strong>Nombres unificados.</strong> Al reunir las variantes de un nombre se pudo juntar por error a dos personas distintas, o dejar separada a una misma persona. Los casos dudosos son pocos, pero existen.</li>";
    h += "<li><strong>Años.</strong> En algunas tesis el año de la portada no coincide con el del curso en que se archivó (pasa sobre todo con defensas de diciembre y enero). Se anotó el de la portada.</li>";
    h += "<li><strong>Datos que el PDF no trae.</strong> Unas pocas tesis no indican tutor, o su portada es una imagen difícil de leer; esos campos pueden estar vacíos o incompletos.</li>";
    h += "<li><strong>Tesis que faltan o se repiten.</strong> Solo están las tesis con PDF disponible. Algún documento puede aparecer dos veces si se guardó en dos archivos.</li>";
    h += "<li><strong>Estadísticas.</strong> Todos los recuentos salen de estos datos, de modo que heredan sus límites: los años con pocos PDF no significan necesariamente pocas tesis defendidas.</li></ul>";

    h += "<h2>Avisar de un error</h2>";
    var vias = [];
    if (C.contactoCorreo) vias.push('escribe a <a href="mailto:' + esc(C.contactoCorreo) + '">' + esc(C.contactoCorreo) + "</a>");
    if (C.repositorio) vias.push('abre una incidencia en <a href="' + esc(C.repositorio) + '/issues" rel="noopener">el repositorio del proyecto</a>');
    h += "<p>Si encuentras un dato equivocado" + (vias.length ? ", " + vias.join(" o ") + "." : ", comunícalo a quien mantiene este registro.") + " Indica el número de la tesis en el registro (aparece en su ficha) y qué habría que corregir.</p>";

    h += "<h2>Datos</h2><p>Todo el sitio se genera a partir de una hoja de Excel, que puedes <a href=\"" + esc(C.excel) + '" download>descargar completa</a>.' + (META.fecha ? " Datos actualizados el " + esc(META.fecha) + "." : "") + "</p>";
    $("#acerca-cuerpo").innerHTML = h;
  }

  // ---------- pintado general ----------
  function pintar() {
    leerHash();
    document.body.dataset.vista = E.vista;
    $$(".secciones a").forEach(function (a) { if (a.dataset.vista === E.vista) a.setAttribute("aria-current", "page"); else a.removeAttribute("aria-current"); a.href = hash(a.dataset.vista); });
    $$(".vista").forEach(function (v) { v.hidden = v.id !== "vista-" + E.vista; });
    sincronizarControles();
    if (E.vista === "registro") { pintarPerfil(); pintarRegistro(); }
    if (E.vista === "estadisticas") pintarEstadisticas();
    if (E.vista === "acerca") pintarAcerca();
    document.title = ({ registro: "", estadisticas: "Estadísticas · ", acerca: "Sobre el registro · " }[E.vista]) + (C.titulo || "Tesis");
  }

  function eventos() {
    window.addEventListener("hashchange", function () { E.visibles = 0; pintar(); });
    $("#form-buscar").addEventListener("submit", function (ev) { ev.preventDefault(); E.q = $("#q").value.trim(); if (!ordenManual) E.orden = E.q ? "relevancia" : "reciente"; E.visibles = 0; ir("registro"); });
    var espera; $("#q").addEventListener("input", function () { clearTimeout(espera); espera = setTimeout(function () { E.q = $("#q").value.trim(); if (!ordenManual) E.orden = E.q ? "relevancia" : "reciente"; E.visibles = 0; history.replaceState(null, "", hash("registro")); pintarPerfil(); pintarRegistro(); $("#orden").value = E.orden; }, 220); });
    $$('input[name="ambito"]').forEach(function (r) { r.addEventListener("change", function () { E.ambito = r.value; E.visibles = 0; ir("registro"); }); });
    $("#f-desde").addEventListener("change", function (e) { E.desde = e.target.value; if (E.hasta && E.desde && +E.hasta < +E.desde) E.hasta = E.desde; E.visibles = 0; ir(); });
    $("#f-hasta").addEventListener("change", function (e) { E.hasta = e.target.value; if (E.hasta && E.desde && +E.hasta < +E.desde) E.desde = E.hasta; E.visibles = 0; ir(); });
    $("#f-solo-tutor").addEventListener("change", function (e) { E.solo = e.target.checked; E.visibles = 0; ir(); });
    $$('input[name="f-autoria"]').forEach(function (r) { r.addEventListener("change", function () { E.autoria = r.value; E.visibles = 0; ir(); }); });
    $$('input[name="f-resumen"]').forEach(function (r) { r.addEventListener("change", function () { E.resumen = r.value; E.visibles = 0; ir(); }); });
    $("#f-limpiar").addEventListener("click", function () { E.q = ""; E.ambito = "todo"; E.desde = E.hasta = E.tutor = E.autoria = E.resumen = ""; E.solo = false; ordenManual = false; E.orden = "reciente"; E.visibles = 0; ir("registro"); });
    $("#orden").addEventListener("change", function (e) { E.orden = e.target.value; ordenManual = true; E.visibles = 0; ir(); });
    $("#mas").addEventListener("click", function () { var antes = E.visibles; E.visibles += C.porPagina || 30; pintarRegistro(); var li = $("#lista").children[antes]; if (li) { var b = $("button.titulo", li); if (b) b.focus({ preventScroll: true }); } });
    $("#descargar-csv").addEventListener("click", descargarCSV);
    $("#filtros-abrir").addEventListener("click", function (e) { var c = $("#filtros-cuerpo"), a = c.classList.toggle("abierto"); e.currentTarget.setAttribute("aria-expanded", a ? "true" : "false"); });
    $("#activos").addEventListener("click", function (ev) { var b = ev.target.closest("[data-quitar]"); if (b) quitar(b.dataset.quitar); });
    $("#lista").addEventListener("click", function (ev) {
      var tu = ev.target.closest("[data-tutor]");
      if (tu) { E.tutor = tu.dataset.tutor; E.solo = false; E.visibles = 0; ir("registro"); window.scrollTo({ top: $("#contenido").offsetTop - 10 }); return; }
      var b = ev.target.closest("button.titulo, button.ver"); if (!b) return;
      var li = b.closest("li"), t = TESIS.filter(function (x) { return x.id === li.dataset.id; })[0], d = $(".detalle", li), tit = $("button.titulo", li), ver = $("button.ver", li);
      if (d) { d.remove(); tit.setAttribute("aria-expanded", "false"); ver.textContent = t.resumen ? "Leer el resumen" : "Ver la ficha"; }
      else { $(".tesis", li).insertAdjacentHTML("beforeend", detalle(t)); tit.setAttribute("aria-expanded", "true"); ver.textContent = "Cerrar"; var ex = $(".extracto", li); if (ex) ex.remove(); }
    });
    // estadísticas
    $("#est-base").addEventListener("change", function (e) { E.base = e.target.value; ir("estadisticas"); });
    $("#est-tutor-buscar").addEventListener("input", function () { pintarTablaTutores(); });
    $("#est-tutores-mas").addEventListener("click", function () { tablaCompleta = true; pintarTablaTutores(); });
    $("#vista-estadisticas").addEventListener("click", function (ev) {
      var o = ev.target.closest("[data-orden]"); if (o) { var k = o.dataset.orden; ordenTabla = { k: k, asc: ordenTabla.k === k ? !ordenTabla.asc : k === "n" || k === "min" }; pintarTablaTutores(); return; }
      var conservar = E.base === "busqueda";
      function limpio() { if (!conservar) { E.q = ""; E.ambito = "todo"; E.desde = E.hasta = E.tutor = E.autoria = E.resumen = ""; E.solo = false; } E.visibles = 0; }
      var tu = ev.target.closest("[data-tutor]"); if (tu) { limpio(); E.tutor = tu.dataset.tutor; E.solo = !!tu.closest("#est-tutores-top"); ir("registro"); window.scrollTo(0, 0); return; }
      var an = ev.target.closest("[data-anio]"); if (an) { limpio(); E.desde = E.hasta = an.dataset.anio; globo(); ir("registro"); window.scrollTo(0, 0); return; }
      var pa = ev.target.closest("[data-palabra]"); if (pa) { limpio(); E.q = pa.dataset.palabra; E.ambito = "titulo"; ir("registro"); window.scrollTo(0, 0); }
    });
    $("#est-anios").addEventListener("pointermove", function (ev) { var b = ev.target.closest("[data-anio]"); if (b) globo("<b>" + b.dataset.anio + "</b>: " + plural(+b.dataset.n, "tesis", "tesis"), ev); else globo(); });
    $("#est-anios").addEventListener("pointerleave", function () { globo(); });
    iniciarPerfil(); iniciarCombo();
  }

  // ---------- arranque ----------
  function listo() {
    var op = '<option value="">—</option>' + ANIOS.map(function (y) { return "<option>" + y + "</option>"; }).join("");
    $("#f-desde").innerHTML = op; $("#f-hasta").innerHTML = op;
    $("#estado").hidden = true;
    $("#pie-texto").innerHTML = esc(C.titulo || "") + ". " + plural(TESIS.length, "tesis", "tesis") + " en el registro" + (META.fecha ? "; datos actualizados el " + esc(META.fecha) : "") + '. <a href="#/acerca">Cómo se hizo y qué errores puede tener</a>.';
    if (!eventos.hecho) { eventos(); eventos.hecho = true; }
    pintar();
  }
  function fallo(titulo, detalle, permitirArchivo) {
    var e = $("#estado"); e.hidden = false; e.className = "estado error";
    e.innerHTML = "<h2>" + esc(titulo) + "</h2><p>" + esc(detalle) + "</p>" +
      (permitirArchivo ? '<p class="soltar"><label>Para ver el sitio con un Excel de tu computadora, elígelo aquí: <input type="file" id="archivo-local" accept=".xlsx,.xlsm,.xls"></label></p>' : "");
    var inp = $("#archivo-local");
    if (inp) inp.addEventListener("change", function () {
      var f = inp.files[0]; if (!f) return;
      var r = new FileReader();
      r.onload = function () { try { procesar(XLSX.read(new Uint8Array(r.result), { type: "array" })); META.fecha = ""; e.className = "estado"; listo(); } catch (err) { fallo("No pude leer ese Excel", err.message, true); } };
      r.readAsArrayBuffer(f);
    });
  }
  function arrancar() {
    if (C.titulo) { $("#sitio-titulo").textContent = C.titulo; document.title = C.titulo; }
    $("#sitio-subtitulo").textContent = C.subtitulo || "";
    if (typeof XLSX === "undefined") return fallo("Falta un archivo del sitio", "No se cargó js/vendor/xlsx.mini.min.js. Comprueba que la carpeta «js» se subió completa.", false);
    fetch(C.excel, { cache: "no-cache" }).then(function (r) {
      if (!r.ok) throw new Error("http " + r.status);
      var lm = r.headers.get("last-modified");
      META.fecha = C.fechaDatos || (lm ? new Date(lm).toLocaleDateString("es", { day: "numeric", month: "long", year: "numeric" }) : "");
      return r.arrayBuffer();
    }).then(function (buf) {
      try { procesar(XLSX.read(new Uint8Array(buf), { type: "array" })); listo(); }
      catch (err) { fallo("El Excel no tiene el formato esperado", err.message, true); }
    }).catch(function () {
      fallo("No se pudo cargar el registro", location.protocol === "file:"
        ? "Estás abriendo la página directamente desde una carpeta, y así el navegador no deja leer el archivo «" + C.excel + "». En GitHub Pages funciona sin hacer nada."
        : "No encuentro el archivo «" + C.excel + "». Comprueba que está subido en esa carpeta y con ese nombre.", true);
    });
  }
  arrancar();
})();
