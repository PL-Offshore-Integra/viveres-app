import { useState, useEffect, useCallback, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./lib/supabase";

const USUARIO = "Comprador";
const PORTAL_URL = "https://integra.ploffshore.com";
const BASES = ["Golondrina de Mar", "Atlantic Dama", "Parana Ports"];
const UNIDADES_PEDIDO = ["Kg", "Litros", "Unidad", "Caja", "Bolsa", "Atado", "Cajón", "Ristra", "Lata", "Pote", "Docena", "Bandeja"];
const UNIDADES_ANALISIS = ["Kg", "Litros"];
const PLAZO_PAGO_OPTIONS = ["Contado", "15 días", "30 días", "45 días", "60 días", "90 días"];

const TEMP_COLOR = {
  "Seco":        { bg: "#FEF9C3", color: "#92400E", border: "#FDE68A", dot: "#EAB308" },
  "Refrigerado": { bg: "#DBEAFE", color: "#1E40AF", border: "#BFDBFE", dot: "#3B82F6" },
  "Congelado":   { bg: "#EDE9FE", color: "#4C1D95", border: "#DDD6FE", dot: "#8B5CF6" },
};

const STATUS_PEDIDO = {
  borrador:  { label: "Borrador",             color: "b-gray" },
  enviado:   { label: "Enviado al comprador", color: "b-blue" },
  aprobado:  { label: "Aprobado",             color: "b-green" },
  rechazado: { label: "Rechazado",            color: "b-red" },
};

const TRACKER_STATUS = {
  pendiente:  { label: "Pendiente",  color: "b-amber" },
  en_camino:  { label: "En camino",  color: "b-blue" },
  entregado:  { label: "Entregado",  color: "b-green" },
};

const fmt = (n) => n != null ? new Intl.NumberFormat("es-AR", { maximumFractionDigits: 3 }).format(n) : "—";
const fmtDate = d => d ? new Date(d).toLocaleDateString("es-AR") : "—";
// Cantidad pedida = lo que cargó el requisitor, siempre fija.
// Cantidad autorizada = lo que definió el comprador al aprobar (puede ser
// distinta). Para cualquier cálculo o vista "de lo que realmente se
// compra/entrega" hay que usar esta cantidad efectiva.
const cantEfectiva = (it) => (it?.cantidad_autorizada != null ? it.cantidad_autorizada : (it?.cantidad_pedida || 0));

//  API 
const api = {
  async getCatalogo() {
    const { data, error } = await supabase.from("viveres_catalogo").select("*").eq("activo", true).order("categoria").order("descripcion");
    if (error) throw error;
    return data || [];
  },
  async getParametros() {
    const { data, error } = await supabase.from("viveres_parametros_dieta").select("*");
    if (error) throw error;
    return data || [];
  },
  async getPedidos(filtros = {}) {
    let q = supabase.from("viveres_pedidos").select("*, viveres_pedido_items(*)").order("created_at", { ascending: false });
    if (filtros.status) q = q.eq("status", filtros.status);
    if (filtros.statuses) q = q.in("status", filtros.statuses);
    const { data, error } = await q;
    if (error) throw error;
    return data || [];
  },
  async crearPedido(pedido, items) {
    const { proyecto, ...resto } = pedido;
    const { data: nuevo, error } = await supabase
      .from("viveres_pedidos")
      .insert([{ ...resto, fecha_pedido: pedido.fecha_pedido || null, fecha_necesaria: pedido.fecha_necesaria || null }])
      .select()
      .single();
    if (error) throw error;
    if (items?.length) {
      // cantidad_pedida es la que carga el requisitor: queda fija para
      // siempre. Lo que el comprador autoriza se guarda aparte, en
      // cantidad_autorizada (ver ModalRevisar), sin tocar este valor.
      const { error: errItems } = await supabase
        .from("viveres_pedido_items")
        .insert(items.map(it => ({ ...it, pedido_id: nuevo.id })));
      if (errItems) throw errItems;
    }
    return nuevo;
  },
  async actualizarPedido(id, cambios) {
    const { proyecto, ...resto } = cambios;
    const { data, error } = await supabase
      .from("viveres_pedidos")
      .update({ ...resto, updated_at: new Date().toISOString() })
      .eq("id", id)
      .select()
      .single();
    if (error) throw error;
    return data;
  },
  async actualizarItems(pedidoId, items) {
    const { error: errDel } = await supabase
      .from("viveres_pedido_items")
      .delete()
      .eq("pedido_id", pedidoId);
    if (errDel) throw errDel;
    if (items?.length) {
      const { error: errIns } = await supabase
        .from("viveres_pedido_items")
        .insert(items.map(it => ({ ...it, pedido_id: pedidoId })));
      if (errIns) throw errIns;
    }
  },
  async eliminarPedido(id) {
    const { error: errItems } = await supabase.from("viveres_pedido_items").delete().eq("pedido_id", id);
    if (errItems) throw errItems;
    const { error } = await supabase.from("viveres_pedidos").delete().eq("id", id);
    if (error) throw error;
  },
  async subirRemito(file, pedidoId) {
    const path = `viveres/remitos/${pedidoId}/${Date.now()}_${file.name}`;
    const { error } = await supabase.storage.from("cotizaciones").upload(path, file, { upsert: true });
    if (error) throw error;
    const { data } = supabase.storage.from("cotizaciones").getPublicUrl(path);
    return data.publicUrl;
  },
  async getSolicitantes() {
    const { data, error } = await supabase
      .from("viveres_solicitantes")
      .select("*")
      .eq("activo", true)
      .order("nombre");
    if (error) throw error;
    return data || [];
  },
  async crearSolicitante(nombre) {
    const { data, error } = await supabase
      .from("viveres_solicitantes")
      .insert([{ nombre: nombre.trim(), activo: true }])
      .select()
      .maybeSingle();
    if (error) throw error;
    return data;
  },
  async eliminarSolicitante(id) {
    const { error } = await supabase
      .from("viveres_solicitantes")
      .delete()
      .eq("id", id);
    if (error) throw error;
  },
  async getStockVuelta() {
    const { data, error } = await supabase
      .from("viveres_stock_vuelta")
      .select("*, viveres_stock_vuelta_items(*)")
      .order("fecha", { ascending: false })
      .order("created_at", { ascending: false });
    if (error) throw error;
    return data || [];
  },
  async crearStockVuelta(cabecera, items) {
    const { data: nuevo, error } = await supabase
      .from("viveres_stock_vuelta")
      .insert([{ ...cabecera }])
      .select()
      .single();
    if (error) throw error;
    if (items?.length) {
      const { error: errItems } = await supabase
        .from("viveres_stock_vuelta_items")
        .insert(items.map(it => ({ ...it, stock_vuelta_id: nuevo.id })));
      if (errItems) throw errItems;
    }
    return nuevo;
  },
  async eliminarStockVuelta(id) {
    const { error: errItems } = await supabase.from("viveres_stock_vuelta_items").delete().eq("stock_vuelta_id", id);
    if (errItems) throw errItems;
    const { error } = await supabase.from("viveres_stock_vuelta").delete().eq("id", id);
    if (error) throw error;
  },
};

//  CSS 
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}

/*  TOKENS · INTEGRA Brand Book v1.0 
   Los nombres de variable son los que ya usaba esta app: cambian los valores,
   no los selectores. Navy = estructura, nunca acción. Un solo color de acción.
    */
:root{
  --navy:#082F4E;--blue:#056D76;--mid:#4A5560;--light:#C9D0D6;
  --bg:#FAFBFC;--surface:#FFFFFF;--surface2:#F4F6F8;--surface3:#E4E8EC;
  --border:#E4E8EC;--border2:#C9D0D6;
  --text:#0F1419;--muted:#4A5560;--muted2:#7A8792;
  --accent:#056D76;--accent2:#0E7A5F;--warn:#8F5A0B;--danger:#B3261E;
  --purple:#4A5560;--teal:#056D76;--orange:#8F5A0B;
  --mono:'IBM Plex Mono',monospace;--sans:'IBM Plex Sans',sans-serif;--r:4px;--r2:4px;
  --nav:#082F4E;--action:#056D76;--action-press:#04565D;
  --tr:color 120ms cubic-bezier(.2,0,.38,.9),background-color 120ms cubic-bezier(.2,0,.38,.9),border-color 120ms cubic-bezier(.2,0,.38,.9);
}
/* Instancia: se activa con <html data-instance="pl-offshore"> en index.html */
[data-instance="pl-offshore"]{--nav:#002247;--action:#002247;--blue:#002247;--accent:#002247}
[data-instance="clean-sea"]{--nav:#1B3765;--action:#006945;--blue:#006945;--accent:#006945}
[data-instance="terramare"]{--nav:#213363;--action:#1F5285;--blue:#1F5285;--accent:#1F5285}

body{background:var(--bg);color:var(--text);font-family:var(--sans);font-size:15px;line-height:1.55;min-height:100vh;overflow-x:hidden}
*:focus-visible{outline:2px solid var(--action);outline-offset:2px}
.app{display:flex;min-height:100vh;overflow-x:hidden}

/*  NAVEGACIÓN LATERAL · 240px, colapsa a iconos en mobile  */
.sidebar{width:240px;min-width:240px;background:var(--nav);display:flex;flex-direction:column}
.sidebar-header{border-bottom:1px solid rgba(255,255,255,.14)}
.sidebar-logo-wrap{padding:14px 16px;display:flex;align-items:center;gap:12px;height:56px}
.sidebar-logo-img{width:28px;height:28px;object-fit:contain;border-radius:var(--r);border:0;background:rgba(255,255,255,.14)}
.sidebar-logo-main{font-size:14px;font-weight:600;color:#fff;letter-spacing:0;text-transform:none}
.sidebar-logo-sub{font-family:var(--mono);font-size:11px;color:rgba(255,255,255,.72);margin-top:2px;letter-spacing:.06em;text-transform:uppercase}
.nav-section{padding:16px 16px 6px;font-family:var(--mono);font-size:11px;letter-spacing:.08em;color:rgba(255,255,255,.72);text-transform:uppercase}
.ni{display:flex;align-items:center;gap:10px;padding:9px 16px;font-size:14px;font-weight:500;cursor:pointer;color:rgba(255,255,255,.72);border-left:3px solid transparent;transition:var(--tr);user-select:none;min-height:36px}
.ni:hover{color:#fff;background:rgba(255,255,255,.08)}
.ni.active{color:#fff;border-left-color:var(--action);background:rgba(255,255,255,.12);font-weight:500}
.ni.sub{padding-left:34px;font-size:13px;font-weight:400}
.ni.sub.active{font-weight:500}
.ni.back{color:rgba(255,255,255,.72);font-size:13px;border-top:1px solid rgba(255,255,255,.14);margin-top:6px}
.ni.back:hover{color:#fff}
.ni-icon{font-size:14px;width:16px;text-align:center;flex-shrink:0}
.ni-badge{margin-left:auto;background:rgba(255,255,255,.14);color:#fff;font-family:var(--mono);font-size:11px;font-weight:500;padding:2px 7px;border-radius:3px;min-width:20px;text-align:center}
.ni-badge.amber{background:rgba(255,255,255,.14)}
.ni-badge.gray{background:rgba(255,255,255,.14);color:rgba(255,255,255,.72)}

/*  BARRA SUPERIOR · 56px  */
.main{flex:1;display:flex;flex-direction:column;overflow:hidden;min-width:0}
.topbar{background:var(--surface);border-bottom:1px solid var(--border);padding:0 24px;height:56px;display:flex;align-items:center;justify-content:space-between}
.topbar-title{font-family:var(--mono);font-size:11px;font-weight:500;letter-spacing:.08em;color:var(--muted);text-transform:uppercase}
.content{flex:1;overflow-y:auto;overflow-x:hidden;padding:24px;background:var(--bg)}

/*  PANELES · blancos, borde 1px, radio 4, sin sombra  */
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:24px;margin-bottom:16px}
.card-title{font-family:var(--mono);font-size:11px;font-weight:500;letter-spacing:.08em;color:var(--muted);text-transform:uppercase;margin-bottom:16px;display:flex;align-items:center;justify-content:space-between;gap:12px}

/*  KPIs  */
.stats{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;margin-bottom:24px}
.stat{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:16px 18px}
.stat-label{font-family:var(--mono);font-size:11px;color:var(--muted);font-weight:500;letter-spacing:.08em;margin-bottom:8px;text-transform:uppercase}
.stat-value{font-family:var(--mono);font-size:30px;font-weight:600;color:var(--navy);font-variant-numeric:tabular-nums}
.va{color:var(--navy)}.vg{color:var(--accent2)}.vr{color:var(--danger)}.vp{color:var(--muted)}.vm{color:var(--warn)}.vgr{color:var(--muted)}

/*  TABLAS · fila 40px, regla marcada de 2px navy, dato en mono  */
.table-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
th{font-family:var(--mono);font-size:11px;font-weight:500;letter-spacing:.08em;color:var(--muted);text-transform:uppercase;padding:10px 12px;text-align:left;border-bottom:2px solid var(--navy);white-space:nowrap;background:var(--surface)}
td{padding:12px;border-bottom:1px solid var(--border);vertical-align:middle}
tr:last-child td{border-bottom:none}
tr.click:hover td{background:var(--surface2);cursor:pointer}
.tracker-table th{font-family:var(--mono);font-size:11px;font-weight:500;letter-spacing:.08em;color:var(--muted);text-transform:uppercase;padding:10px 12px;text-align:left;border-bottom:2px solid var(--navy);white-space:nowrap;background:var(--surface);position:sticky;top:0;z-index:2}
.tracker-table th.sortable{cursor:pointer;user-select:none}
.tracker-table th.sortable:hover{color:var(--navy)}
.tracker-table td{padding:12px;border-bottom:1px solid var(--border);vertical-align:middle}
.tracker-table tr:hover td{background:var(--surface2);cursor:pointer}
.tracker-table tr:last-child td{border-bottom:none}

/*  FILTROS  */
.filter-row{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap;align-items:center}
.filter-input,.filter-select{background:var(--surface);border:1px solid var(--border2);border-radius:var(--r);color:var(--text);font-family:var(--sans);font-size:14px;height:36px;padding:0 10px;outline:none;min-width:150px;transition:var(--tr)}
.filter-select{cursor:pointer}
.filter-input:focus,.filter-select:focus{border-width:2px;border-color:var(--action);padding:0 9px}

/*  BADGES DE ESTADO · fondo tenue, texto de estado, mono caja alta  */
.badge{display:inline-flex;align-items:center;font-family:var(--mono);font-size:11px;font-weight:500;padding:3px 8px;border-radius:3px;white-space:nowrap;letter-spacing:.06em;text-transform:uppercase}
.b-amber{background:#FBF1E3;color:#8F5A0B;border:0}
.b-blue{background:#E6F1F2;color:#056D76;border:0}
.b-teal{background:#E8F3EF;color:#0E7A5F;border:0}
.b-red{background:#FAEAE8;color:#B3261E;border:0}
.b-purple{background:#F4F6F8;color:#4A5560;border:0}
.b-orange{background:#FBF1E3;color:#8F5A0B;border:0}
.b-green{background:#E8F3EF;color:#0E7A5F;border:0}
.b-gray{background:#F4F6F8;color:#4A5560;border:0}
.urgdot{width:6px;height:6px;border-radius:50%;display:inline-block;margin-right:6px;flex-shrink:0}

/*  BOTONES · un solo primario por vista. Nada se mueve al presionar  */
.btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;font-family:var(--sans);font-size:14px;font-weight:500;letter-spacing:0;height:36px;padding:0 16px;border-radius:var(--r);border:1px solid transparent;cursor:pointer;transition:var(--tr);white-space:nowrap;text-transform:none}
.btn-primary{background:var(--action);color:#fff}
.btn-primary:hover{background:var(--navy)}
.btn-primary:active{background:var(--action-press)}
.btn-success{background:var(--accent2);color:#fff}
.btn-success:hover{background:#0B6249}
.btn-danger{background:var(--surface);color:var(--danger);border-color:var(--border2)}
.btn-danger:hover{background:#FAEAE8;border-color:var(--danger)}
.btn-ghost{background:var(--surface);color:var(--muted);border-color:var(--border2)}
.btn-ghost:hover{color:var(--text);background:var(--surface2)}
.btn-warn{background:var(--surface);color:var(--warn);border-color:var(--border2)}
.btn-warn:hover{background:#FBF1E3;border-color:var(--warn)}
.btn-cond{background:var(--surface);color:var(--muted);border-color:var(--border2)}
.btn-cond:hover{background:var(--surface2)}
.btn-confirm{background:var(--surface);color:var(--warn);border-color:var(--border2)}
.btn-confirm:hover{background:#FBF1E3}
.btn-sm{height:28px;padding:0 12px;font-size:13px}
.btn:disabled{background:var(--surface3);color:var(--muted2);border-color:transparent;cursor:not-allowed}

/*  CAPAS FLOTANTES · la única sombra del sistema  */
.overlay{position:fixed;inset:0;background:rgba(15,20,25,.45);display:flex;align-items:flex-start;justify-content:center;z-index:100;padding:24px;overflow-y:auto}
.modal{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);width:100%;max-width:860px;margin:auto;box-shadow:0 8px 24px rgba(15,20,25,.14)}
.modal-lg{max-width:1120px}
.mhdr{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;padding:20px 24px;border-bottom:1px solid var(--border);background:var(--surface);border-radius:var(--r) var(--r) 0 0}
.mtitle{font-size:18px;font-weight:600;letter-spacing:0;color:var(--navy)}
.mbody{padding:24px}
.mftr{padding:16px 24px;border-top:1px solid var(--border);display:flex;justify-content:flex-end;gap:8px;background:var(--surface2);border-radius:0 0 var(--r) var(--r)}
.mclose{background:none;border:none;color:var(--muted);font-size:20px;cursor:pointer;line-height:1;transition:var(--tr)}
.mclose:hover{color:var(--navy)}
@keyframes fadeIn{from{opacity:1}to{opacity:1}}
@keyframes slideUp{from{opacity:1}to{opacity:1}}

/*  FORMULARIOS · campo 36px, foco borde 2px  */
.fg{display:flex;flex-direction:column;gap:6px}
.fg label{font-family:var(--mono);font-size:11px;color:var(--muted);letter-spacing:.08em;text-transform:uppercase;font-weight:500}
.fg input,.fg select,.fg textarea{background:var(--surface);border:1px solid var(--border2);border-radius:var(--r);color:var(--text);font-family:var(--sans);font-size:14px;height:36px;padding:0 12px;outline:none;transition:var(--tr)}
.fg textarea{resize:vertical;min-height:72px;height:auto;padding:10px 12px}
.fg input:focus,.fg select:focus,.fg textarea:focus{border-width:2px;border-color:var(--action);padding:0 11px}
.fg textarea:focus{padding:9px 11px}
.form-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
.form-grid-3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;margin-bottom:16px}
.form-section{font-family:var(--mono);font-size:11px;font-weight:500;letter-spacing:.08em;color:var(--muted);text-transform:uppercase;margin:32px 0 16px;padding-bottom:8px;border-bottom:1px solid var(--border)}
.items-edit th{font-family:var(--mono);font-size:11px;background:var(--surface)}
.items-edit td{padding:6px 8px}
.items-edit input,.items-edit select{background:var(--surface);border:1px solid var(--border2);border-radius:var(--r);color:var(--text);font-family:var(--mono);font-size:13px;height:32px;padding:0 8px;width:100%;outline:none;transition:var(--tr)}
.items-edit input:focus,.items-edit select:focus{border-width:2px;border-color:var(--action);padding:0 7px}

/*  TRAZABILIDAD  */
.tl{list-style:none}
.tl-item{display:flex;gap:12px;padding-bottom:16px;position:relative}
.tl-item:not(:last-child)::before{content:'';position:absolute;left:11px;top:24px;bottom:0;width:1px;background:var(--border)}
.tl-dot{width:24px;height:24px;border-radius:50%;background:var(--surface2);border:1px solid var(--border2);display:flex;align-items:center;justify-content:center;font-size:11px;flex-shrink:0;z-index:1}
.tl-dot.c{border-color:var(--action);color:var(--action);background:#E6F1F2}
.tl-dot.a{border-color:var(--accent2);color:var(--accent2);background:#E8F3EF}
.tl-dot.r{border-color:var(--danger);color:var(--danger);background:#FAEAE8}
.tl-dot.u{border-color:var(--warn);color:var(--warn);background:#FBF1E3}
.tl-ev{font-size:14px;font-weight:500;color:var(--navy)}
.tl-meta{font-family:var(--mono);font-size:11px;color:var(--muted);margin-top:4px}

/*  FILA DE REQUISICIÓN · el estado va en el borde izquierdo de 3px  */
.req-row{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:16px 18px;margin-bottom:12px;cursor:pointer;transition:var(--tr)}
.req-row:hover{border-color:var(--navy)}
.req-row.unread{border-left:3px solid var(--action)}
.req-row.devuelto{border-left:3px solid var(--warn)}
.req-row.pend-confirm{border-left:3px solid var(--warn)}
.req-title{font-weight:600;font-size:15px;margin-bottom:6px;color:var(--navy)}
.req-meta{display:flex;gap:16px;font-size:13px;color:var(--muted);flex-wrap:wrap;align-items:center}

/*  AVISOS  */
.notif{position:fixed;bottom:24px;right:24px;background:var(--surface);border:1px solid var(--border);border-left-width:3px;border-radius:var(--r);padding:14px 16px;font-size:14px;z-index:300;max-width:360px;display:flex;align-items:center;gap:12px;box-shadow:0 8px 24px rgba(15,20,25,.14)}
.n-green{border-left-color:var(--accent2)}.n-red{border-left-color:var(--danger)}.n-amber{border-left-color:var(--warn)}.n-blue{border-left-color:var(--action)}
.info-box{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:14px 16px;font-size:14px}
.info-box.accent{border-left:3px solid var(--action)}
.info-box.warn{border-left:3px solid var(--warn)}
.info-box.danger{border-left:3px solid var(--danger)}
.info-box.orange{border-left:3px solid var(--warn)}

/*  UTILIDADES  */
.flex-gap{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.flex-between{display:flex;justify-content:space-between;align-items:center;gap:12px}
.mt8{margin-top:8px}.mt12{margin-top:12px}.mt16{margin-top:16px}
.mb8{margin-bottom:8px}.mb12{margin-bottom:12px}.mb16{margin-bottom:16px}
.text-mono{font-family:var(--mono);font-variant-numeric:tabular-nums}
.text-muted{color:var(--muted)}
.empty-state{text-align:center;padding:48px 24px;color:var(--muted);font-size:15px}
.loading{display:flex;align-items:center;justify-content:center;padding:48px;color:var(--muted);gap:12px;font-size:15px}
.spin{animation:spin 1s linear infinite}
@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
.kbar{margin-bottom:12px}
.kbar-lbl{display:flex;justify-content:space-between;font-size:13px;margin-bottom:6px}
.kbar-track{height:6px;background:var(--surface3);border-radius:3px;overflow:hidden;border:0}
.kbar-fill{height:100%;border-radius:3px}
.tabs-row{display:flex;gap:0;border-bottom:1px solid var(--border);margin-bottom:24px;overflow-x:auto}
.tab{font-size:14px;font-weight:500;padding:10px 16px;cursor:pointer;color:var(--muted);border-bottom:2px solid transparent;transition:var(--tr);text-transform:none;letter-spacing:0;margin-bottom:-1px;white-space:nowrap}
.tab:hover{color:var(--navy)}
.tab.active{color:var(--action);border-bottom-color:var(--action)}
.grupo-chip{display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;border-radius:3px;font-family:var(--mono);font-size:12px;font-weight:500;background:var(--surface2);color:var(--navy);border:1px solid var(--border);flex-shrink:0}
.tag{display:inline-block;font-family:var(--mono);font-size:11px;padding:3px 7px;background:var(--surface2);border:1px solid var(--border);border-radius:3px;color:var(--muted);letter-spacing:.06em;text-transform:uppercase}
.fecha-chip{display:inline-flex;flex-direction:column;gap:2px;font-family:var(--mono);font-size:11px;color:var(--text);white-space:nowrap;font-variant-numeric:tabular-nums}
.fecha-chip span:first-child{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.08em}
.tracker-simple-row{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:16px}
.tracker-simple-row.en-curso{border-left:3px solid var(--warn)}
.tracker-simple-row.entregado{border-left:3px solid var(--accent2)}
.req-row-actions{display:flex;flex-direction:row;gap:8px;margin-top:12px;padding-top:12px;border-top:1px solid var(--border);justify-content:flex-end}
.cotiz-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:16px;margin-bottom:16px}

/*  MOBILE  */
@media (max-width: 768px) {
  .app { flex-direction: column; }
  .sidebar { display: none; }
  .main { width: 100%; padding-bottom: 72px; }
  .topbar { padding: 0 16px; }
  .content { padding: 16px; }
  .card { padding: 16px; margin-bottom: 12px; }
  .stats { grid-template-columns: 1fr 1fr; gap: 12px; }
  .stat { padding: 14px; }
  .stat-value { font-size: 24px; }
  .form-grid, .form-grid-3 { grid-template-columns: 1fr; gap: 12px; }
  .table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
  table { font-size: 13px; min-width: 540px; }
  th, td { padding: 10px 8px; }
  .tracker-table th, .tracker-table td { padding: 10px 8px; }
  .filter-row { flex-direction: column; align-items: stretch; }
  .filter-input, .filter-select { min-width: unset; width: 100%; }
  .btn { height: 44px; padding: 0 14px; }
  .btn-sm { height: 36px; }
  .mftr { flex-wrap: wrap; gap: 8px; }
  .mftr .btn { flex: 1; justify-content: center; }
  .overlay { padding: 0; align-items: flex-end; }
  .modal { border-radius: var(--r) var(--r) 0 0; max-width: 100%; max-height: 92vh; overflow-y: auto; }
  .modal-lg { max-width: 100%; }
  .req-meta { gap: 10px; }
  .req-title { font-size: 15px; }
  .tabs-row { overflow-x: auto; -webkit-overflow-scrolling: touch; }
  .tab { font-size: 13px; padding: 10px 12px; }
  .notif { bottom: 88px; right: 12px; left: 12px; max-width: unset; }
  .items-edit { font-size: 13px; }
  .items-edit th, .items-edit td { padding: 6px; }
  .items-edit table { min-width: 380px; }
  .req-row-actions{flex-direction:column;gap:8px;width:100%}
  .req-row-actions .btn{width:100%}
  .mftr{flex-direction:column;align-items:stretch;gap:8px}
  .mftr .btn{width:100%;flex:unset}
  .mftr .btn-success{order:-3}.mftr .btn-primary{order:-2}.mftr .btn-danger{order:-1}
  .card-title{flex-direction:column;align-items:flex-start;gap:10px}
  .card-title .btn{width:100%}
  .filter-row .btn{width:100%}
  .form-footer-actions{flex-direction:column !important;align-items:stretch !important}
  .form-footer-actions .btn{width:100%}
  .cotiz-grid{grid-template-columns:1fr !important}
  .req-row .flex-between{flex-direction:column;align-items:flex-start;gap:10px}
  .req-row .flex-between > .flex-gap:last-child{width:100%;flex-direction:column;gap:8px}
  .req-row .flex-between > .flex-gap:last-child .btn{width:100%}
}

/*  NAVEGACIÓN INFERIOR (solo mobile)  */
@media (max-width: 768px) {
  .mobile-nav {
    display: flex !important;
    position: fixed; bottom: 0; left: 0; right: 0;
    background: var(--nav); border-top: 1px solid rgba(255,255,255,.14);
    z-index: 50; height: 64px;
    justify-content: space-around; align-items: center;
    padding: 0 4px; overflow-x: auto;
  }
  .mobile-nav-item {
    display: flex; flex-direction: column; align-items: center; gap: 3px;
    cursor: pointer; padding: 8px; border-radius: var(--r);
    color: rgba(255,255,255,.72); transition: var(--tr); flex: 1;
    position: relative; min-width: 48px; min-height: 48px; justify-content: center;
  }
  .mobile-nav-item.active { color: #fff; background: rgba(255,255,255,.12); }
  .mobile-nav-item:hover { color: #fff; }
  .mobile-nav-icon { font-size: 16px; line-height: 1; }
  .mobile-nav-label { font-family: var(--mono); font-size: 11px; font-weight: 500; letter-spacing: .06em; text-transform: uppercase; text-align: center; }
  .mobile-nav-badge {
    position: absolute; top: 4px; right: 8px;
    background: rgba(255,255,255,.14); color: #fff;
    font-family: var(--mono); font-size: 10px; font-weight: 500;
    padding: 1px 5px; border-radius: 3px; min-width: 16px; text-align: center;
  }
  .mobile-nav-badge.amber { background: rgba(255,255,255,.14); }
  .mobile-nav-badge.gray { background: rgba(255,255,255,.14); }
}
@media (min-width: 769px) {
  .mobile-nav { display: none !important; }
}

/*  ARMAZÓN · shell del prototipo 
   La navegación del módulo es BLANCA con borde derecho; el navy es la barra
   superior. El ítem activo lleva borde izquierdo de 3px en el color de acción.
    */
.shell{display:grid;grid-template-columns:248px minmax(0,1fr);align-items:stretch;min-height:100vh}
.shell.is-collapsed{grid-template-columns:68px minmax(0,1fr)}

.appbar{height:56px;background:var(--nav);display:flex;align-items:center;gap:24px;padding:0 24px;flex:0 0 auto}
.appbar-iso{height:26px;width:auto;object-fit:contain;display:block;flex:0 0 auto}
.appbar-div{width:1px;height:24px;background:rgba(255,255,255,.14);flex:0 0 auto}
.appbar-instance{font:500 14px/1.2 var(--sans);color:#fff;white-space:nowrap;flex:0 0 auto}
.appbar-search{flex:1;max-width:380px;display:flex;align-items:center;gap:10px;height:32px;padding:0 12px;background:rgba(255,255,255,.10);border:0;border-radius:var(--r);font:400 14px/1.2 var(--sans);color:rgba(255,255,255,.72)}
.appbar-search::placeholder{color:rgba(255,255,255,.72)}
.appbar-tools{margin-left:auto;display:flex;align-items:center;gap:16px}
.appbar-avatar{width:28px;height:28px;border-radius:var(--r);background:rgba(255,255,255,.14);color:#fff;font-family:var(--mono);font-size:12px;font-weight:500;line-height:28px;text-align:center;flex:0 0 auto}
.appbar-user{font:500 13px/1.25 var(--sans);color:#fff;white-space:nowrap}
.appbar-link{background:none;border:0;padding:0;cursor:pointer;font:500 13px/1.2 var(--sans);color:rgba(255,255,255,.86);white-space:nowrap}
.appbar-link:hover{color:#fff;text-decoration:underline}

.sidebar{width:auto;min-width:0;background:var(--surface);border-right:1px solid var(--border);display:flex;flex-direction:column}
.sidebar-header{border-bottom:1px solid var(--border);padding:16px;display:flex;align-items:center;gap:12px;min-height:69px}
.sidebar-logo-img{width:32px;height:32px;object-fit:contain;border:0;border-radius:0;background:none;flex:0 0 auto}
.sidebar-logo-main{font:600 15px/1.3 var(--sans);color:var(--navy);letter-spacing:0;text-transform:none}
.sidebar-logo-sub{font-family:var(--mono);font-size:11px;font-weight:500;color:var(--muted);letter-spacing:.06em;text-transform:uppercase;margin-top:2px}
.sidebar-nav{flex:1;padding:12px 0;overflow-y:auto}
.nav-section{padding:14px 16px 8px;font-family:var(--mono);font-size:11px;font-weight:500;letter-spacing:.08em;color:var(--muted);text-transform:uppercase;text-align:left}
.ni{display:flex;align-items:center;gap:12px;width:100%;padding:9px 16px 9px 13px;background:transparent;border:0;border-left:3px solid transparent;cursor:pointer;text-align:left;font:400 14px/1.3 var(--sans);color:var(--muted);transition:var(--tr);min-height:38px}
.ni:hover{background:var(--surface2);color:var(--navy)}
.ni.active{background:var(--surface2);border-left-color:var(--action);color:var(--navy);font-weight:500}
.ni-ico{display:block;flex:0 0 auto;color:var(--muted2)}
.ni.active .ni-ico{color:var(--action)}
.ni-label{flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ni-badge{margin-left:auto;font-family:var(--mono);font-size:11px;font-weight:500;color:var(--muted);background:var(--surface2);padding:3px 6px;border-radius:3px;min-width:22px;text-align:center;border:1px solid var(--border)}
.ni.active .ni-badge{color:var(--action);background:var(--surface);border-color:var(--border2)}
.ni-badge.amber{color:var(--warn)}
.ni-badge.gray{color:var(--muted)}
.sidebar-foot{border-top:1px solid var(--border);padding:12px 8px;display:flex;flex-direction:column;gap:2px}
.sidebar-foot-btn{display:flex;align-items:center;gap:12px;width:100%;padding:9px 10px;background:none;border:0;border-radius:var(--r);cursor:pointer;font:500 13px/1.2 var(--sans);color:var(--muted);transition:var(--tr)}
.sidebar-foot-btn:hover{background:var(--surface2);color:var(--navy)}
.sidebar-foot-meta{padding:8px 10px 0;font-family:var(--mono);font-size:11px;font-weight:500;line-height:1.6;letter-spacing:.06em;color:var(--muted2)}
.shell.is-collapsed .sidebar-header{justify-content:center;padding:16px 8px}
.shell.is-collapsed .ni{justify-content:center;padding:9px 8px 9px 5px}
.shell.is-collapsed .sidebar-foot-btn{justify-content:center}

/*  encabezado de pantalla  */
.pagehead{background:var(--surface);border-bottom:1px solid var(--border);padding:16px 24px;flex:0 0 auto}
.crumb{display:flex;align-items:center;gap:8px;font:400 13px/1.2 var(--sans);color:var(--muted)}
.crumb button{background:none;border:0;padding:0;cursor:pointer;font:400 13px/1.2 var(--sans);color:var(--action)}
.crumb button:hover{text-decoration:underline;color:var(--navy)}
.crumb-current{color:var(--text)}
.pagehead-row{display:flex;align-items:flex-end;justify-content:space-between;gap:24px;margin-top:10px}
.pagehead h1{font:600 24px/1.25 var(--sans);color:var(--navy);margin:0}
.pagehead p{font:400 13px/1.45 var(--sans);color:var(--muted);margin:6px 0 0;max-width:70ch}
.pagehead-actions{display:flex;gap:8px;flex:0 0 auto}

@media (max-width:768px){
  .shell,.shell.is-collapsed{grid-template-columns:1fr}
  .sidebar{display:none}
  .appbar{gap:12px;padding:0 16px}
  .appbar-search,.appbar-instance{display:none}
  .pagehead{padding:14px 16px}
  .pagehead-row{flex-direction:column;align-items:stretch;gap:12px}
  .pagehead-actions .btn{flex:1}
  .main{padding-bottom:72px}
}

`;

//  HELPERS 
function Notif({ msg, onClose }) {
  if (!msg) return null;
  const cls = { success: "n-green", error: "n-red", warn: "n-amber", info: "n-blue" }[msg.type] || "n-blue";
  return <div className={`notif ${cls}`}><span>{msg.text}</span><button onClick={onClose} style={{ marginLeft: "auto", background: "none", border: "none", color: "var(--muted)", cursor: "pointer" }}>✕</button></div>;
}

function FG({ label, hint, children, full }) {
  return <div className="fg" style={full ? { gridColumn: "1/-1" } : {}}>
    {label && <label>{label}</label>}
    {children}
    {hint && <div style={{ fontSize: 10, color: "var(--muted2)", marginTop: 2 }}>{hint}</div>}
  </div>;
}

function TempBadge({ temp }) {
  const tc = TEMP_COLOR[temp] || { bg: "#F3F4F6", color: "#6B7280", border: "#E5E7EB", dot: "#9CA3AF" };
  return <span style={{ display: "inline-flex", alignItems: "center", gap: 4, fontSize: 10, fontWeight: 600, color: tc.color, background: tc.bg, border: `1px solid ${tc.border}`, borderRadius: 4, padding: "2px 6px" }}>
    <span style={{ width: 5, height: 5, borderRadius: "50%", background: tc.dot, display: "inline-block" }} />{temp}
  </span>;
}

function calcDieta(items, paxDias) {
  const grupos = {};
  items.forEach(it => {
    const total = (it.stock_actual || 0) + (it.cantidad_pedida || 0);
    const porPaxDia = paxDias > 0 ? (total * (it.volumen_peso || 1)) / paxDias : 0;
    grupos[it.categoria] = (grupos[it.categoria] || 0) + porPaxDia;
  });
  return grupos;
}

function exportarParaProveedor(pedido, items) {
  // Al proveedor le mandamos lo AUTORIZADO por el comprador (si ya se
  // definió); la cantidad pedida original queda solo como referencia.
  const rows = items.filter(it => cantEfectiva(it) > 0).map(it => ({
    "Categoría": it.categoria || "", "Temperatura": it.temperatura || "",
    "Descripción": it.descripcion || "", "Unidad de pedido": it.unidad || "",
    "Cantidad pedida (original)": it.cantidad_pedida || 0,
    "Cantidad a comprar (autorizada)": cantEfectiva(it), "Observaciones": "",
  }));
  const grupos = {};
  items.filter(it => cantEfectiva(it) > 0).forEach(it => {
    const total = cantEfectiva(it) * (it.volumen_peso || 1);
    if (!grupos[it.categoria]) grupos[it.categoria] = { total: 0, unidad: it.unidad_analisis || "Kg" };
    grupos[it.categoria].total += total;
  });
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Pedido Víveres");
  const resumen = Object.entries(grupos).map(([cat, d]) => ({ "Categoría": cat, [`Total (${d.unidad})`]: Math.round(d.total * 100) / 100 }));
  if (resumen.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(resumen), "Resumen");
  XLSX.writeFile(wb, `viveres_${(pedido.base_buque || "pedido").replace(/ /g, "_")}_${(pedido.fecha_pedido || "").slice(0, 10)}.xlsx`);
}

//  FORM PEDIDO 
function FormPedido({ pedidoInicial, catalogoInicial, parametros, solicitantes = [], stockVuelta = [], onSave, onCancel, notify }) {
  const [step, setStep] = useState(1);
  const [catalogo] = useState(catalogoInicial || []);
  const [saving, setSaving] = useState(false);
  const [cabecera, setCabecera] = useState({
    empresa: "Parana Logistica", base_buque: pedidoInicial?.base_buque || "",
    pax: pedidoInicial?.pax || 12, dias: pedidoInicial?.dias || 15,
    fecha_pedido: pedidoInicial?.fecha_pedido || new Date().toISOString().split("T")[0],
    fecha_necesaria: pedidoInicial?.fecha_necesaria || "",
    solicitado_por: pedidoInicial?.solicitado_por || "",
    observaciones: pedidoInicial?.observaciones || "",
  });
  const [items, setItems] = useState(() => {
    const ex = pedidoInicial?.viveres_pedido_items || [];
    return catalogo.map(c => {
      const found = ex.find(e => e.catalogo_id === c.id);
      return { catalogo_id: c.id, descripcion: c.descripcion, categoria: c.categoria, subcategoria: c.subcategoria || "", temperatura: c.temperatura || "", unidad: c.unidad || "Unidad", unidad_analisis: c.unidad_analisis || "Kg", volumen_peso: c.volumen_peso || 1, stock_actual: found?.stock_actual || 0, cantidad_pedida: found?.cantidad_pedida || 0 };
    });
  });
  const [itemsManuales, setItemsManuales] = useState(() => {
    if (!pedidoInicial?.viveres_pedido_items) return [];
    return pedidoInicial.viveres_pedido_items.filter(it => !it.catalogo_id).map(it => ({ ...it, id: it.id || `m_${Date.now()}_${Math.random()}` }));
  });
  const [filtroCateg, setFiltroCateg] = useState("");
  const [filtroTemp, setFiltroTemp] = useState("");
  const [busqueda, setBusqueda] = useState("");

  const blankManual = () => ({ id: `m_${Date.now()}_${Math.random()}`, catalogo_id: null, descripcion: "", categoria: "Almacén", temperatura: "Seco", unidad: "Unidad", unidad_analisis: "Kg", volumen_peso: 1, stock_actual: 0, cantidad_pedida: 0 });
  const setCab = (k, v) => setCabecera(c => ({ ...c, [k]: v }));
  const stockEditadoManualmente = useRef(false);
  const setItem = (id, k, v) => {
    if (k === "stock_actual") stockEditadoManualmente.current = true;
    setItems(prev => prev.map(it => it.catalogo_id === id ? { ...it, [k]: parseFloat(v) || 0 } : it));
  };
  const setManual = (i, k, v) => { const arr = [...itemsManuales]; arr[i] = { ...arr[i], [k]: v }; setItemsManuales(arr); };
  const setManualNum = (i, k, v) => { const arr = [...itemsManuales]; arr[i] = { ...arr[i], [k]: parseFloat(v) || 0 }; setItemsManuales(arr); };

  // Al elegir el buque, si hay un registro de "stock vuelta a puerto" para ese buque,
  // completamos el stock a bordo con lo cargado en ese registro (el más reciente).
  // No pisamos ediciones manuales que el usuario ya haya hecho en esta carga.
  const registroStockVuelta = cabecera.base_buque
    ? stockVuelta.find(r => r.base_buque === cabecera.base_buque) || null
    : null;
  useEffect(() => {
    if (pedidoInicial) return; // en edición de un pedido existente no tocamos lo ya cargado
    if (!registroStockVuelta) return;
    if (stockEditadoManualmente.current) return;
    const stockPorCatalogo = {};
    (registroStockVuelta.viveres_stock_vuelta_items || []).forEach(it => {
      if (it.catalogo_id) stockPorCatalogo[it.catalogo_id] = it.stock;
    });
    setItems(prev => prev.map(it => stockPorCatalogo[it.catalogo_id] !== undefined ? { ...it, stock_actual: stockPorCatalogo[it.catalogo_id] } : it));
    notify(`Stock a bordo completado con el registro de vuelta a puerto del ${fmtDate(registroStockVuelta.fecha)}`, "info");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [registroStockVuelta]);

  const paxDias = (cabecera.pax || 0) * (cabecera.dias || 0);
  const todosItems = [...items, ...itemsManuales];
  const dietaActual = calcDieta(todosItems, paxDias);
  const itemsConPedido = todosItems.filter(it => it.cantidad_pedida > 0 && (it.descripcion || "").trim());
  const categorias = [...new Set(catalogo.map(c => c.categoria))].sort();
  const temperaturas = [...new Set(catalogo.map(c => c.temperatura).filter(Boolean))];
  const itemsFiltrados = items.filter(it => {
    if (filtroCateg && it.categoria !== filtroCateg) return false;
    if (filtroTemp && it.temperatura !== filtroTemp) return false;
    if (busqueda && !it.descripcion.toLowerCase().includes(busqueda.toLowerCase())) return false;
    return true;
  });

  const handleGuardar = async (status = "borrador") => {
    if (!cabecera.base_buque || !cabecera.solicitado_por) { alert("Completá Base/Buque y Solicitado por"); return; }
    setSaving(true);
    try {
      const itemsAGuardar = [...items.filter(it => it.cantidad_pedida > 0 || it.stock_actual > 0), ...itemsManuales.filter(it => it.descripcion.trim() && (it.cantidad_pedida > 0 || it.stock_actual > 0))].map(({ id: _id, ...rest }) => rest);
      await onSave({ ...cabecera, status }, itemsAGuardar, status);
    } catch (e) { notify("Error: " + e.message, "error"); }
    finally { setSaving(false); }
  };

  if (step === 1) return (
    <div className="card">
      <div className="card-title">Datos del pedido</div>
      <div className="form-grid-3">
        <FG label="Base / Buque *"><select value={cabecera.base_buque} onChange={e => setCab("base_buque", e.target.value)}><option value="">Seleccionar...</option>{BASES.map(b => <option key={b}>{b}</option>)}</select></FG>
        <FG label="Solicitado por *"><select value={cabecera.solicitado_por} onChange={e => setCab("solicitado_por", e.target.value)}><option value="">Seleccionar...</option>{solicitantes.map(s => <option key={s.id} value={s.nombre}>{s.nombre}</option>)}{cabecera.solicitado_por && !solicitantes.some(s => s.nombre === cabecera.solicitado_por) && <option value={cabecera.solicitado_por}>{cabecera.solicitado_por}</option>}</select></FG>
        <FG label="Proyecto"><input value={cabecera.proyecto || ""} onChange={e => setCab("proyecto", e.target.value)} placeholder="Ej: OP-2026-003" /></FG>
      </div>
      <div className="form-grid">
        <FG label="PAX"><input type="number" value={cabecera.pax} onChange={e => setCab("pax", parseInt(e.target.value) || 0)} min={1} /></FG>
        <FG label="Días"><input type="number" value={cabecera.dias} onChange={e => setCab("dias", parseInt(e.target.value) || 0)} min={1} /></FG>
        <FG label="Fecha del pedido"><input type="date" value={cabecera.fecha_pedido} onChange={e => setCab("fecha_pedido", e.target.value)} /></FG>
        <FG label="Fecha necesaria"><input type="date" value={cabecera.fecha_necesaria} onChange={e => setCab("fecha_necesaria", e.target.value)} /></FG>
      </div>
      <FG label="Observaciones"><textarea value={cabecera.observaciones} onChange={e => setCab("observaciones", e.target.value)} placeholder="Notas adicionales..." /></FG>
      {cabecera.pax > 0 && cabecera.dias > 0 && <div className="info-box accent mt12" style={{ fontSize: 12 }}>Total: <strong>{cabecera.pax} PAX × {cabecera.dias} días = {paxDias} raciones</strong></div>}
      {registroStockVuelta && (
        <div className="info-box accent mt12" style={{ fontSize: 12 }}>
          Hay un registro de <strong>stock a la vuelta a puerto</strong> del <strong>{fmtDate(registroStockVuelta.fecha)}</strong> para {cabecera.base_buque}. Se va a usar como stock a bordo inicial en el paso siguiente (podés editarlo ítem por ítem).
        </div>
      )}
      <div className="form-footer-actions mt16">
        <button className="btn btn-ghost" onClick={onCancel}>Cancelar</button>
        <button className="btn btn-primary" onClick={() => { if (!cabecera.base_buque || !cabecera.solicitado_por) { alert("Completá Base/Buque y Solicitado por"); return; } setStep(2); }}>Continuar → Cargar ítems</button>
      </div>
    </div>
  );

  return (
    <div>
      <div className="form-grid" style={{ marginBottom: 16 }}>
        <div className="card" style={{ margin: 0 }}>
          <div className="card-title">Datos del pedido</div>
          <div style={{ fontSize: 13, fontWeight: 600, color: "var(--navy)" }}>{cabecera.base_buque}</div>
          <div style={{ fontSize: 12, color: "var(--muted)", marginTop: 4 }}>PL Offshore · {cabecera.pax} PAX · {cabecera.dias} días · <strong>{paxDias} raciones</strong></div>
          <div style={{ fontSize: 12, color: "var(--muted)", marginTop: 2 }}>Por: {cabecera.solicitado_por}</div>
          <button className="btn btn-ghost btn-sm mt8" onClick={() => setStep(1)}>← Editar datos</button>
        </div>
        <div className="card" style={{ margin: 0 }}>
          <div className="card-title">Control de dieta — análisis / persona / día</div>
          <div className="dieta-grid">
            {parametros.map(p => {
              const val = dietaActual[p.grupo] || 0;
              const status = val === 0 ? "yellow" : val < p.min ? "red" : val > p.max ? "red" : "green";
              const colors = { green: { bg: "#D1FAE5", color: "#065F46" }, red: { bg: "#FEE2E2", color: "#991B1B" }, yellow: { bg: "#FEF9C3", color: "#92400E" } };
              return <div key={p.grupo} className="dieta-chip" style={{ background: colors[status].bg }}><span style={{ fontSize: 10, color: colors[status].color, fontWeight: 600 }}>{p.grupo}</span><span style={{ fontFamily: "var(--mono)", fontSize: 11, color: colors[status].color }}>{val.toFixed(2)} / {p.max} {p.unidad_medida}</span></div>;
            })}
          </div>
        </div>
      </div>

      <div className="tabs-row">
        <div className={`tab ${filtroCateg === "" ? "active" : ""}`} onClick={() => setFiltroCateg("")}>Todos</div>
        {categorias.map(cat => {
          const cnt = items.filter(it => it.categoria === cat && it.cantidad_pedida > 0).length;
          return <div key={cat} className={`tab ${filtroCateg === cat ? "active" : ""}`} onClick={() => setFiltroCateg(cat)}>{cat}{cnt > 0 && <span style={{ marginLeft: 6, background: "var(--accent2)", color: "#fff", fontSize: 9, fontWeight: 700, padding: "1px 5px", borderRadius: 8, fontFamily: "var(--mono)" }}>{cnt}</span>}</div>;
        })}
        <div className={`tab ${filtroCateg === "__manual__" ? "active" : ""}`} onClick={() => setFiltroCateg("__manual__")}>
           Ingreso manual
          {itemsManuales.filter(it => it.cantidad_pedida > 0 && it.descripcion.trim()).length > 0 && <span style={{ marginLeft: 6, background: "var(--blue)", color: "#fff", fontSize: 9, fontWeight: 700, padding: "1px 5px", borderRadius: 8, fontFamily: "var(--mono)" }}>{itemsManuales.filter(it => it.cantidad_pedida > 0 && it.descripcion.trim()).length}</span>}
        </div>
      </div>

      {filtroCateg === "__manual__" ? (
        <div style={{ marginBottom: 90 }}>
          <div className="info-box accent mb12" style={{ fontSize: 11 }}>Agregá productos que no están en el catálogo.</div>
          <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
            <button className="btn btn-primary" onClick={() => setItemsManuales([...itemsManuales, blankManual()])}>+ Agregar ítem manual</button>
          </div>
          {itemsManuales.length === 0 ? (
            <div className="manual-empty">
              <div style={{ fontSize: 36 }}></div>
              <div style={{ fontSize: 14, fontWeight: 600, color: "var(--navy)" }}>Sin ítems manuales</div>
              <div style={{ fontSize: 12, color: "var(--muted)" }}>Hacé click en "+ Agregar ítem manual" para agregar productos que no están en el catálogo</div>
              <button className="btn btn-primary mt8" onClick={() => setItemsManuales([...itemsManuales, blankManual()])}>+ Agregar primer ítem</button>
            </div>
          ) : (
            <div>
              {itemsManuales.map((it, i) => {
                const totalAnalisis = (it.cantidad_pedida || 0) * (it.volumen_peso || 1);
                return (
                  <div key={it.id} className="manual-row">
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                      <div style={{ fontFamily: "var(--mono)", fontSize: 10, color: "var(--muted)", fontWeight: 600 }}>ÍTEM {i + 1}</div>
                      <button onClick={() => setItemsManuales(itemsManuales.filter((_, j) => j !== i))} style={{ background: "none", border: "none", color: "var(--danger)", cursor: "pointer", fontSize: 18, lineHeight: 1 }}>✕</button>
                    </div>
                    <div className="manual-grid-3">
                      <FG label="Temperatura"><select value={it.temperatura} onChange={e => setManual(i, "temperatura", e.target.value)}><option>Seco</option><option>Refrigerado</option><option>Congelado</option></select></FG>
                      <FG label="Categoría"><select value={it.categoria} onChange={e => setManual(i, "categoria", e.target.value)}>{categorias.map(c => <option key={c}>{c}</option>)}<option>Otro</option></select></FG>
                      <FG label="Descripción *"><input value={it.descripcion} onChange={e => setManual(i, "descripcion", e.target.value)} placeholder="Nombre del producto..." /></FG>
                    </div>
                    <div className="manual-grid-5">
                      <FG label="Unidad pedido"><select value={it.unidad} onChange={e => setManual(i, "unidad", e.target.value)}>{UNIDADES_PEDIDO.map(u => <option key={u}>{u}</option>)}</select></FG>
                      <FG label="Unidad análisis"><select value={it.unidad_analisis || "Kg"} onChange={e => setManual(i, "unidad_analisis", e.target.value)}>{UNIDADES_ANALISIS.map(u => <option key={u}>{u}</option>)}</select></FG>
                      <FG label="Vol/Peso x unidad"><input type="number" step="0.001" min="0" value={it.volumen_peso || ""} onChange={e => setManual(i, "volumen_peso", parseFloat(e.target.value) || 1)} placeholder="1" /></FG>
                      <FG label="Stock actual"><input type="number" min={0} value={it.stock_actual || ""} onChange={e => setManualNum(i, "stock_actual", e.target.value)} placeholder="0" /></FG>
                      <FG label="Cantidad pedida"><input type="number" min={0} value={it.cantidad_pedida || ""} onChange={e => setManualNum(i, "cantidad_pedida", e.target.value)} placeholder="0" style={{ background: it.cantidad_pedida > 0 ? "#DCFCE7" : undefined, fontWeight: it.cantidad_pedida > 0 ? 700 : 400, borderColor: it.cantidad_pedida > 0 ? "#86EFAC" : undefined }} /></FG>
                    </div>
                    {totalAnalisis > 0 && <div style={{ marginTop: 8, fontSize: 11, color: "var(--accent)", fontFamily: "var(--mono)" }}>→ Total análisis: {(totalAnalisis).toFixed(3)} {it.unidad_analisis || "Kg"}</div>}
                  </div>
                );
              })}
              <div style={{ display: "flex", justifyContent: "center", marginTop: 12 }}>
                <button onClick={() => setItemsManuales([...itemsManuales, blankManual()])} style={{ background: "transparent", color: "var(--blue)", border: "2px dashed var(--blue)", borderRadius: "var(--r)", padding: "10px 24px", fontFamily: "var(--sans)", fontSize: 12, fontWeight: 600, cursor: "pointer", width: "100%" }}>+ Agregar otro ítem manual</button>
              </div>
            </div>
          )}
        </div>
      ) : (
        <div>
          <div className="filter-row" style={{ marginBottom: 12 }}>
            <input className="filter-input" placeholder=" Buscar ítem..." value={busqueda} onChange={e => setBusqueda(e.target.value)} />
            <select className="filter-select" value={filtroTemp} onChange={e => setFiltroTemp(e.target.value)}>
              <option value="">Todas las temperaturas</option>
              {temperaturas.map(t => <option key={t}>{t}</option>)}
            </select>
            {(filtroTemp || busqueda) && <button className="btn btn-ghost btn-sm" onClick={() => { setFiltroTemp(""); setBusqueda(""); }}>✕</button>}
            <span style={{ marginLeft: "auto", fontFamily: "var(--mono)", fontSize: 11, color: "var(--muted)" }}>{itemsFiltrados.length} visibles</span>
          </div>
          <div className="card" style={{ padding: 0, overflow: "hidden", marginBottom: 90 }}>
            <div className="table-wrap">
              <table className="tracker-table">
                <thead><tr><th>Temp.</th><th>Categoría</th><th>Descripción</th><th>Unidad pedido</th><th>× Kg/L</th><th style={{ width: 80 }}>Stock</th><th style={{ width: 100 }}>Pedido</th><th>Total</th><th>Análisis/PAX/día</th></tr></thead>
                <tbody>
                  {itemsFiltrados.map(it => {
                    const total = (it.stock_actual || 0) + (it.cantidad_pedida || 0);
                    const totalAnalisis = total * (it.volumen_peso || 1);
                    const porPaxDia = paxDias > 0 ? totalAnalisis / paxDias : 0;
                    return (
                      <tr key={it.catalogo_id} style={{ background: it.cantidad_pedida > 0 ? "#F0FDF4" : "inherit" }}>
                        <td><TempBadge temp={it.temperatura} /></td>
                        <td style={{ fontSize: 11, color: "var(--muted)" }}>{it.categoria}</td>
                        <td style={{ fontWeight: it.cantidad_pedida > 0 ? 600 : 400, fontSize: 12 }}>{it.descripcion}</td>
                        <td style={{ fontSize: 11, color: "var(--muted)" }}>{it.unidad}</td>
                        <td style={{ fontSize: 10, color: "var(--muted2)", fontFamily: "var(--mono)" }}>{it.volumen_peso !== 1 ? `×${it.volumen_peso}` : "—"} {it.unidad_analisis || "Kg"}</td>
                        <td><input type="number" min={0} value={it.stock_actual || ""} placeholder="0" onChange={e => setItem(it.catalogo_id, "stock_actual", e.target.value)} style={{ width: 70, background: "var(--surface)", border: "1px solid var(--border)", borderRadius: "var(--r)", fontFamily: "var(--mono)", fontSize: 12, padding: "4px 8px", outline: "none", textAlign: "right" }} /></td>
                        <td><input type="number" min={0} value={it.cantidad_pedida || ""} placeholder="0" onChange={e => setItem(it.catalogo_id, "cantidad_pedida", e.target.value)} style={{ width: 80, background: it.cantidad_pedida > 0 ? "#DCFCE7" : "var(--surface)", border: `1px solid ${it.cantidad_pedida > 0 ? "#86EFAC" : "var(--border)"}`, borderRadius: "var(--r)", fontFamily: "var(--mono)", fontSize: 12, padding: "4px 8px", outline: "none", textAlign: "right", fontWeight: it.cantidad_pedida > 0 ? 700 : 400 }} /></td>
                        <td className="text-mono" style={{ fontSize: 11, color: total > 0 ? "var(--navy)" : "var(--muted2)" }}>{total > 0 ? total : "—"}</td>
                        <td className="text-mono" style={{ fontSize: 11, color: porPaxDia > 0 ? "var(--accent)" : "var(--muted2)" }}>{porPaxDia > 0 ? porPaxDia.toFixed(3) : "—"}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      <div className="fixed-action-bar">
        <div style={{ flex: 1 }}>
          {itemsConPedido.length === 0 ? <span style={{ fontSize: 12, color: "rgba(255,255,255,.5)" }}>Sin ítems seleccionados</span> :
            <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
              {[...new Set(itemsConPedido.map(it => it.categoria))].map(cat => (
                <div key={cat} style={{ display: "flex", alignItems: "center", gap: 4 }}>
                  <span style={{ fontSize: 10, color: "rgba(255,255,255,.5)" }}>{cat}</span>
                  <span style={{ fontSize: 11, fontFamily: "var(--mono)", fontWeight: 700, color: "#fff", background: "rgba(255,255,255,.15)", borderRadius: 4, padding: "1px 6px" }}>{itemsConPedido.filter(it => it.categoria === cat).length}</span>
                </div>
              ))}
            </div>
          }
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: "#fff", fontFamily: "var(--mono)" }}>{itemsConPedido.length} ítem{itemsConPedido.length !== 1 ? "s" : ""}</div>
          <button className="btn btn-ghost" onClick={() => setStep(1)} style={{ color: "rgba(255,255,255,.7)", borderColor: "rgba(255,255,255,.2)" }}>← Volver</button>
          <button className="btn" onClick={() => handleGuardar("borrador")} disabled={saving} style={{ background: "rgba(255,255,255,.15)", color: "#fff", borderColor: "rgba(255,255,255,.2)" }}>Guardar borrador</button>
          <button className="btn btn-success" onClick={() => handleGuardar("enviado")} disabled={saving || itemsConPedido.length === 0}>{saving ? "Enviando..." : "✓ Enviar al comprador"}</button>
        </div>
      </div>
    </div>
  );
}

//  PAGE: NUEVO PEDIDO 
function PageNuevo({ notify, onSaved, onCancel }) {
  const [catalogo, setCatalogo] = useState([]);
  const [parametros, setParametros] = useState([]);
  const [solicitantes, setSolicitantes] = useState([]);
  const [stockVuelta, setStockVuelta] = useState([]);
  const [loading, setLoading] = useState(true);
  useEffect(() => {
    // getStockVuelta() se carga aparte: si esa consulta falla (tablas nuevas de
    // Supabase con algún problema, por ejemplo), no queremos que se caiga el
    // catálogo ni los solicitantes, que son imprescindibles para armar el pedido.
    Promise.all([api.getCatalogo(), api.getParametros(), api.getSolicitantes()])
      .then(([cat, par, sol]) => { setCatalogo(cat); setParametros(par); setSolicitantes(sol); })
      .catch(e => notify("Error al cargar datos: " + e.message, "error"))
      .finally(() => setLoading(false));
    api.getStockVuelta()
      .then(sv => setStockVuelta(sv))
      .catch(e => console.error("No se pudo cargar el historial de stock vuelta a puerto:", e.message));
  }, [notify]);
  if (loading) return <div className="loading"><span className="spin">◌</span> Cargando catálogo...</div>;
  return <FormPedido catalogoInicial={catalogo} parametros={parametros} solicitantes={solicitantes} stockVuelta={stockVuelta} onSave={async (cab, items) => { await api.crearPedido(cab, items); onSaved(); }} onCancel={onCancel} notify={notify} />;
}

//  MODAL: REVISAR PEDIDO 
function ModalRevisar({ pedido, onClose, onActualizado, notify }) {
  const [loading, setLoading] = useState(true);
  const [modo, setModo] = useState("detalle");
  const [motivoRechazo, setMotivoRechazo] = useState("");
  const [saving, setSaving] = useState(false);
  const [itemsEdit, setItemsEdit] = useState([]);
  const [aprobadoPor, setAprobadoPor] = useState("");

  useEffect(() => {
    // cantidad_pedida (lo que cargó el requisitor) nunca se toca acá.
    // cantidad_autorizada es lo que el comprador puede ajustar; arranca
    // igual a lo pedido y desde ahí se edita.
    const raw = (pedido.viveres_pedido_items || [])
      .filter(it => it.cantidad_pedida > 0)
      .map(it => ({ ...it, cantidad_autorizada: it.cantidad_autorizada ?? it.cantidad_pedida, _eliminado: false }));
    setItemsEdit(raw);
    setLoading(false);
  }, [pedido]);

  const itemsVisibles = itemsEdit.filter(it => !it._eliminado);
  const huboCambios = itemsEdit.some(
    it => it._eliminado || it.cantidad_autorizada !== it.cantidad_pedida
  );

  const setCantidad = (id, val) => {
    setItemsEdit(prev =>
      prev.map(it => it.id === id ? { ...it, cantidad_autorizada: parseFloat(val) || 0 } : it)
    );
  };

  const eliminarItem = (id) => {
    setItemsEdit(prev =>
      prev.map(it => it.id === id ? { ...it, _eliminado: true } : it)
    );
  };

  const restaurarItem = (id) => {
    setItemsEdit(prev =>
      prev.map(it => it.id === id ? { ...it, _eliminado: false, cantidad_autorizada: it.cantidad_pedida } : it)
    );
  };

  const handleAprobar = async () => {
    if (itemsVisibles.length === 0) {
      alert("No quedan ítems en el pedido. Rechazalo en cambio.");
      return;
    }
    if (!aprobadoPor.trim()) {
      alert("Ingresá quién aprueba el pedido");
      return;
    }
    setSaving(true);
    try {
      // cantidad_pedida viaja intacta dentro de "rest" — nunca se pisa.
      // Solo se guarda/actualiza cantidad_autorizada.
      const itemsAGuardar = itemsVisibles.map(
        ({ _eliminado, ...rest }) => rest
      );
      await api.actualizarItems(pedido.id, itemsAGuardar);
      await api.actualizarPedido(pedido.id, {
        status: "aprobado",
        fecha_aprobacion: new Date().toISOString(),
        tracker_status: "pendiente",
        aprobado_por: aprobadoPor.trim(),
      });
      notify("Pedido aprobado", "success");
      onActualizado();
    } catch (e) {
      notify("Error: " + e.message, "error");
    } finally {
      setSaving(false);
    }
  };

  const handleRechazar = async () => {
    if (!motivoRechazo.trim()) return alert("Ingresá un motivo");
    setSaving(true);
    try {
      await api.actualizarPedido(pedido.id, {
        status: "rechazado",
        observaciones: motivoRechazo,
      });
      notify("Pedido rechazado", "warn");
      onActualizado();
    } finally {
      setSaving(false);
    }
  };

  if (loading) return (
    <div className="overlay">
      <div className="modal">
        <div className="mbody"><div className="loading"><span className="spin">◌</span></div></div>
      </div>
    </div>
  );

  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal modal-lg">
        {/* HEADER */}
        <div className="mhdr">
          <div>
            <div className="mtitle"> {pedido.base_buque} — Pedido de Víveres</div>
            <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 4 }}>
              PL Offshore · {pedido.pax} PAX · {pedido.dias} días · {pedido.solicitado_por}
              {pedido.fecha_necesaria && (
                <span style={{ color: "var(--warn)", marginLeft: 8 }}>Nec: {fmtDate(pedido.fecha_necesaria)}</span>
              )}
            </div>
          </div>
          <button className="mclose" onClick={onClose}>✕</button>
        </div>

        {/* BODY */}
        <div className="mbody">
          <div className="tabs-row">
            <div className={`tab ${modo === "detalle" ? "active" : ""}`} onClick={() => setModo("detalle")}>Detalle</div>
            <div
              className={`tab ${modo === "rechazar" ? "active" : ""}`}
              onClick={() => setModo("rechazar")}
              style={{ color: modo === "rechazar" ? "var(--danger)" : undefined, borderBottomColor: modo === "rechazar" ? "var(--danger)" : undefined }}
            >
              Rechazar
            </div>
          </div>

          {/* TAB DETALLE */}
          {modo === "detalle" && (
            <div>
              {huboCambios && (
                <div className="info-box warn mb12" style={{ fontSize: 12, display: "flex", alignItems: "center", gap: 8 }}>
                  <span></span>
                  <span>Hay <strong>modificaciones sin aprobar</strong>. Usá "✓ Aprobar" para confirmarlas.</span>
                </div>
              )}

              <div className="table-wrap">
                <table className="items-edit">
                  <thead>
                    <tr>
                      <th>Categoría</th>
                      <th>Temp.</th>
                      <th>Descripción</th>
                      <th>Unidad</th>
                      <th style={{ width: 90, textAlign: "right" }}>Cant. original</th>
                      <th style={{ width: 120, textAlign: "right" }}>Cant. aprobada</th>
                      <th style={{ width: 32 }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {itemsEdit.length === 0 ? (
                      <tr>
                        <td colSpan={7} style={{ textAlign: "center", padding: 24, color: "var(--muted2)" }}>Sin ítems pedidos</td>
                      </tr>
                    ) : (
                      itemsEdit.map(it => {
                        const modificado = !it._eliminado && it.cantidad_autorizada !== it.cantidad_pedida;
                        return (
                          <tr
                            key={it.id}
                            style={{
                              opacity: it._eliminado ? 0.45 : 1,
                              background: it._eliminado ? "#FEF2F2" : modificado ? "#FFFBEB" : "inherit",
                              transition: "all .15s",
                            }}
                          >
                            <td style={{ fontSize: 11, color: "var(--muted)" }}>{it.categoria}</td>
                            <td><TempBadge temp={it.temperatura} /></td>
                            <td style={{
                              fontWeight: 500, fontSize: 12,
                              textDecoration: it._eliminado ? "line-through" : "none",
                              color: it._eliminado ? "var(--muted2)" : "var(--text)",
                            }}>
                              {it.descripcion}
                            </td>
                            <td style={{ fontSize: 11, color: "var(--muted)" }}>{it.unidad}</td>
                            {/* Cantidad original — lo que cargó el requisitor, fijo, no se toca */}
                            <td style={{ fontFamily: "var(--mono)", fontSize: 12, color: "var(--muted)", textAlign: "right" }}>
                              {it.cantidad_pedida}
                            </td>
                            {/* Cantidad autorizada — la define/edita el comprador */}
                            <td>
                              {it._eliminado ? (
                                <button
                                  onClick={() => restaurarItem(it.id)}
                                  style={{ background: "none", border: "1px solid var(--border)", borderRadius: "var(--r)", fontSize: 10, color: "var(--muted)", cursor: "pointer", padding: "3px 8px", fontFamily: "var(--sans)" }}
                                >
                                  ↩ Restaurar
                                </button>
                              ) : (
                                <div style={{ position: "relative" }}>
                                  <input
                                    type="number"
                                    min={0}
                                    value={it.cantidad_autorizada}
                                    onChange={e => setCantidad(it.id, e.target.value)}
                                    style={{
                                      width: "100%",
                                      background: modificado ? "#FEF9C3" : "var(--surface)",
                                      border: `1px solid ${modificado ? "#FDE68A" : "var(--border)"}`,
                                      borderRadius: "var(--r)",
                                      fontFamily: "var(--mono)",
                                      fontSize: 12,
                                      fontWeight: modificado ? 700 : 400,
                                      padding: "5px 8px",
                                      outline: "none",
                                      textAlign: "right",
                                      color: modificado ? "#92400E" : "var(--text)",
                                    }}
                                  />
                                  {modificado && (
                                    <span
                                      title="Cantidad modificada"
                                      style={{ position: "absolute", right: -8, top: -6, width: 8, height: 8, borderRadius: "50%", background: "var(--warn)", display: "block" }}
                                    />
                                  )}
                                </div>
                              )}
                            </td>
                            {/* Botón eliminar */}
                            <td>
                              {!it._eliminado && (
                                <button
                                  onClick={() => eliminarItem(it.id)}
                                  title="Eliminar ítem"
                                  style={{ background: "none", border: "none", color: "var(--muted2)", cursor: "pointer", fontSize: 15, lineHeight: 1, padding: "2px 4px", borderRadius: 4, transition: "color .12s" }}
                                  onMouseEnter={e => e.currentTarget.style.color = "var(--danger)"}
                                  onMouseLeave={e => e.currentTarget.style.color = "var(--muted2)"}
                                >
                                  ✕
                                </button>
                              )}
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>

              {/* Resumen */}
              {itemsEdit.length > 0 && (
                <div style={{ marginTop: 10, fontSize: 11, color: "var(--muted)", fontFamily: "var(--mono)", display: "flex", gap: 16 }}>
                  <span>{itemsVisibles.length} ítem{itemsVisibles.length !== 1 ? "s" : ""} activos</span>
                  {itemsEdit.filter(it => it._eliminado).length > 0 && (
                    <span style={{ color: "var(--danger)" }}>
                      {itemsEdit.filter(it => it._eliminado).length} eliminado{itemsEdit.filter(it => it._eliminado).length !== 1 ? "s" : ""}
                    </span>
                  )}
                  {itemsEdit.filter(it => !it._eliminado && it.cantidad_autorizada !== it.cantidad_pedida).length > 0 && (
                    <span style={{ color: "var(--warn)" }}>
                      {itemsEdit.filter(it => !it._eliminado && it.cantidad_autorizada !== it.cantidad_pedida).length} modificado{itemsEdit.filter(it => !it._eliminado && it.cantidad_autorizada !== it.cantidad_pedida).length !== 1 ? "s" : ""}
                    </span>
                  )}
                </div>
              )}

              <div className="mt12 flex-gap">
                <button className="btn btn-ghost btn-sm" onClick={() => exportarParaProveedor(pedido, itemsVisibles)}>
                  ↓ Exportar para proveedor
                </button>
              </div>

              <div className="form-section">Aprobación</div>
              <div className="form-grid">
                <FG label="Aprobado por *" hint="Queda guardado como respaldo junto con las cantidades originales y autorizadas.">
                  <input
                    value={aprobadoPor}
                    onChange={e => setAprobadoPor(e.target.value)}
                    placeholder="Nombre de quién aprueba..."
                  />
                </FG>
              </div>
            </div>
          )}

          {/* TAB RECHAZAR */}
          {modo === "rechazar" && (
            <div>
              <div className="info-box danger mb12" style={{ fontSize: 12 }}>
                El pedido quedará registrado como rechazado.
              </div>
              <FG label="Motivo *">
                <textarea value={motivoRechazo} onChange={e => setMotivoRechazo(e.target.value)} placeholder="Explicá por qué se rechaza..." style={{ minHeight: 100 }} />
              </FG>
            </div>
          )}
        </div>

        {/* FOOTER */}
        <div className="mftr">
          <button className="btn btn-ghost" onClick={onClose}>Cerrar</button>
          {modo === "rechazar" && (
            <button className="btn btn-danger" onClick={handleRechazar} disabled={saving || !motivoRechazo.trim()}>
              {saving ? "..." : "✕ Confirmar rechazo"}
            </button>
          )}
          {modo === "detalle" && (
            <button
              className="btn btn-success"
              onClick={handleAprobar}
              disabled={saving || itemsVisibles.length === 0 || !aprobadoPor.trim()}
            >
              {saving ? "Aprobando..." : huboCambios ? "✓ Aprobar con cambios" : "✓ Aprobar"}
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

//  MODAL: TRACKER EDITAR 
function ModalTrackerEditar({ pedido, onClose, onSave, notify }) {
  const remitoInputId = `remito-input-${pedido.id}`;
  const [form, setForm] = useState({
    tracker_status: pedido.tracker_status || "pendiente",
    nro_remito: pedido.nro_remito || "",
    fecha_entrega: pedido.fecha_entrega ? pedido.fecha_entrega.slice(0, 10) : "",
    tracker_notas: pedido.tracker_notas || "",
  });
  const [saving, setSaving] = useState(false);
  const [uploading, setUploading] = useState(false);
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));

  const handleUploadRemito = async (file) => {
    if (!file) return;
    setUploading(true);
    try {
      const url = await api.subirRemito(file, pedido.id);
      const updated = await api.actualizarPedido(pedido.id, { remito_url: url, nro_remito: form.nro_remito || file.name });
      notify("Remito adjuntado", "success");
      onSave(updated);
    } catch (e) { notify("Error al subir remito: " + e.message, "error"); }
    finally { setUploading(false); }
  };

  const handleSave = async () => {
    setSaving(true);
    try {
      const cambios = {
        tracker_status: form.tracker_status,
        nro_remito: form.nro_remito || null,
        tracker_notas: form.tracker_notas || null,
        fecha_entrega: form.fecha_entrega ? new Date(form.fecha_entrega).toISOString() : null,
      };
      const updated = await api.actualizarPedido(pedido.id, cambios);
      notify("Tracker actualizado", "success");
      onSave(updated);
    } finally { setSaving(false); }
  };

  // El tracker/entrega refleja lo AUTORIZADO por el comprador (lo que
  // realmente se despacha), no la cantidad pedida original.
  const items = (pedido.viveres_pedido_items || []).filter(it => it.cantidad_pedida > 0);

  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <div className="mhdr">
          <div>
            <div className="mtitle"> Tracker — {pedido.base_buque}</div>
            <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 4 }}>PL Offshore · {pedido.pax} PAX · {pedido.dias} días · {pedido.solicitado_por}</div>
          </div>
          <button className="mclose" onClick={onClose}>✕</button>
        </div>
        <div className="mbody">
          <div className="fecha-chain">
            <div className={`fecha-step ${pedido.created_at ? "done" : ""}`}>
              <div style={{ fontSize: 20 }}></div>
              <div className="fecha-step-label">Solicitud</div>
              <div className="fecha-step-val">{pedido.created_at ? fmtDate(pedido.created_at) : "—"}</div>
            </div>
            <div className="fecha-arrow">→</div>
            <div className={`fecha-step ${pedido.fecha_aprobacion ? "done" : ""}`}>
              <div style={{ fontSize: 20 }}></div>
              <div className="fecha-step-label">Aprobación</div>
              <div className="fecha-step-val">{pedido.fecha_aprobacion ? fmtDate(pedido.fecha_aprobacion) : "—"}</div>
              {pedido.aprobado_por && <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 2 }}>{pedido.aprobado_por}</div>}
            </div>
            <div className="fecha-arrow">→</div>
            <div className={`fecha-step ${pedido.fecha_entrega ? "done" : ""}`}>
              <div style={{ fontSize: 20 }}></div>
              <div className="fecha-step-label">Entrega</div>
              <div className="fecha-step-val">{pedido.fecha_entrega ? fmtDate(pedido.fecha_entrega) : "—"}</div>
            </div>
          </div>

          <div className="form-section">Estado</div>
          <div className="form-grid">
            <FG label="Estado del pedido">
              <select value={form.tracker_status} onChange={e => set("tracker_status", e.target.value)}>
                <option value="pendiente">Pendiente</option>
                <option value="en_camino">En camino</option>
                <option value="entregado">Entregado</option>
              </select>
            </FG>
            <FG label="Fecha de entrega"><input type="date" value={form.fecha_entrega} onChange={e => set("fecha_entrega", e.target.value)} /></FG>
          </div>

          <div className="form-section">Remito</div>
          <div className="form-grid">
            <FG label="N° Remito"><input value={form.nro_remito} onChange={e => set("nro_remito", e.target.value)} placeholder="Ej: 0001-00001234" /></FG>
            <FG label="Remito firmado (PDF / imagen)">
              {pedido.remito_url
                ? <a href={pedido.remito_url} target="_blank" rel="noreferrer" style={{ fontSize: 12, color: "var(--blue)", display: "flex", alignItems: "center", gap: 4, marginTop: 6 }}> Ver remito adjunto</a>
                : <>
                    <input type="file" id={remitoInputId} accept=".pdf,.jpg,.jpeg,.png" style={{ display: "none" }} onChange={e => handleUploadRemito(e.target.files[0])} />
                    <button className="btn btn-ghost btn-sm" style={{ marginTop: 4 }} onClick={() => document.getElementById(remitoInputId).click()} disabled={uploading}>
                      {uploading ? " Subiendo..." : " Adjuntar remito"}
                    </button>
                  </>
              }
            </FG>
          </div>

          <FG label="Notas" full><textarea value={form.tracker_notas} onChange={e => set("tracker_notas", e.target.value)} placeholder="Observaciones sobre la entrega..." style={{ minHeight: 60 }} /></FG>

          {items.length > 0 && <>
            <div className="form-section">Ítems del pedido ({items.length})</div>
            <div className="table-wrap">
              <table>
                <thead><tr><th>Categoría</th><th>Descripción</th><th>Cant. pedida</th><th>Cant. autorizada</th><th>Unidad</th></tr></thead>
                <tbody>
                  {items.map((it, i) => <tr key={i}><td style={{ fontSize: 11, color: "var(--muted)" }}>{it.categoria}</td><td style={{ fontWeight: 500, fontSize: 12 }}>{it.descripcion}</td><td className="text-mono" style={{ fontSize: 12, color: "var(--muted)" }}>{it.cantidad_pedida}</td><td className="text-mono" style={{ fontWeight: 700, color: "var(--accent2)" }}>{cantEfectiva(it)}</td><td style={{ fontSize: 11, color: "var(--muted)" }}>{it.unidad}</td></tr>)}
                </tbody>
              </table>
            </div>
          </>}
        </div>
        <div className="mftr">
          <button className="btn btn-ghost" onClick={onClose}>Cancelar</button>
          <button className="btn btn-primary" onClick={handleSave} disabled={saving}>{saving ? "Guardando..." : "Guardar"}</button>
        </div>
      </div>
    </div>
  );
}

//  PAGE: TRACKER 
function PageTracker({ notify }) {
  const [pedidos, setPedidos] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState(null);
  const [filtroStatus, setFiltroStatus] = useState("");
  const [filtroBase, setFiltroBase] = useState("");
  const [eliminando, setEliminando] = useState(null);

  const load = useCallback(async () => {
    setLoading(true);
    try { setPedidos(await api.getPedidos({ statuses: ["aprobado", "enviado"] })); }
    finally { setLoading(false); }
  }, []);

  useEffect(() => { load(); }, [load]);

  const handleEliminar = async (e, p) => {
    e.stopPropagation();
    if (!window.confirm(`¿Eliminar el pedido de ${p.base_buque} (${p.pax} PAX × ${p.dias} días)? Se borran también sus ítems. Esta acción no se puede deshacer.`)) return;
    setEliminando(p.id);
    const backup = pedidos;
    // optimistic update
    setPedidos(prev => prev.filter(x => x.id !== p.id));
    try {
      await api.eliminarPedido(p.id);
      notify("Pedido eliminado", "warn");
    } catch (err) {
      setPedidos(backup); // revert
      notify("Error al eliminar: " + (err?.message || "desconocido"), "error");
    } finally {
      setEliminando(null);
    }
  };

  const filtrados = pedidos.filter(p => {
    if (filtroStatus && (p.tracker_status || "pendiente") !== filtroStatus) return false;
    if (filtroBase && p.base_buque !== filtroBase) return false;
    return true;
  });

  const bases = [...new Set(pedidos.map(p => p.base_buque).filter(Boolean))].sort();
  const stats = {
    total: pedidos.length,
    pendiente: pedidos.filter(p => !p.tracker_status || p.tracker_status === "pendiente").length,
    en_camino: pedidos.filter(p => p.tracker_status === "en_camino").length,
    entregado: pedidos.filter(p => p.tracker_status === "entregado").length,
  };

  return (
    <div>
      <div className="stats">
        {[
          { label: "Total aprobados", val: stats.total, color: "var(--blue)" },
          { label: "Pendientes", val: stats.pendiente, color: "var(--warn)" },
          { label: "En camino", val: stats.en_camino, color: "var(--blue)" },
          { label: "Entregados", val: stats.entregado, color: "var(--accent2)" },
        ].map(s => (
          <div key={s.label} className="stat">
            <div className="stat-label">{s.label}</div>
            <div className="stat-value" style={{ color: s.color }}>{s.val}</div>
          </div>
        ))}
      </div>

      <div className="filter-row">
        <select className="filter-select" value={filtroStatus} onChange={e => setFiltroStatus(e.target.value)}>
          <option value="">Todos los estados</option>
          <option value="pendiente">Pendiente</option>
          <option value="en_camino">En camino</option>
          <option value="entregado">Entregado</option>
        </select>
        <select className="filter-select" value={filtroBase} onChange={e => setFiltroBase(e.target.value)}>
          <option value="">Todos los barcos</option>
          {bases.map(b => <option key={b}>{b}</option>)}
        </select>
        {(filtroStatus || filtroBase) && <button className="btn btn-ghost btn-sm" onClick={() => { setFiltroStatus(""); setFiltroBase(""); }}>✕ Limpiar</button>}
        <span style={{ marginLeft: "auto", fontFamily: "var(--mono)", fontSize: 11, color: "var(--muted)" }}>{filtrados.length} de {pedidos.length}</span>
      </div>

      {loading ? <div className="loading"><span className="spin">◌</span></div> :
        filtrados.length === 0 ? <div className="empty-state"><div style={{ fontSize: 28, marginBottom: 8 }}></div>Sin pedidos aprobados</div> :
        <div className="card" style={{ padding: 0, overflow: "hidden" }}>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Base/Barco</th>
                  <th>PAX × Días</th>
                  <th>Solicitante</th>
                  <th>Estado</th>
                  <th> Solicitud</th>
                  <th> Aprobación</th>
                  <th>Aprobado por</th>
                  <th> Entrega</th>
                  <th>Remito</th>
                  <th>Notas</th>
                  <th style={{ width: 90, textAlign: "center" }}>Acciones</th>
                </tr>
              </thead>
              <tbody>
                {filtrados.map(p => {
                  const st = p.tracker_status || "pendiente";
                  const stInfo = TRACKER_STATUS[st] || { label: st, color: "b-gray" };
                  return (
                    <tr key={p.id} className="click" onClick={() => setSelected(p)}>
                      <td style={{ fontWeight: 600, fontSize: 12 }}>{p.base_buque}</td>
                      <td className="text-mono" style={{ fontSize: 11, color: "var(--muted)" }}>{p.pax} × {p.dias}</td>
                      <td style={{ fontSize: 12 }}>{p.solicitado_por}</td>
                      <td><span className={`badge ${stInfo.color}`}>{stInfo.label}</span></td>
                      <td className="text-mono" style={{ fontSize: 11, color: "var(--muted)" }}>{p.created_at ? fmtDate(p.created_at) : "—"}</td>
                      <td className="text-mono" style={{ fontSize: 11, color: p.fecha_aprobacion ? "var(--accent2)" : "var(--muted2)" }}>{p.fecha_aprobacion ? fmtDate(p.fecha_aprobacion) : "—"}</td>
                      <td style={{ fontSize: 11, color: "var(--muted)" }}>{p.aprobado_por || "—"}</td>
                      <td className="text-mono" style={{ fontSize: 11, color: p.fecha_entrega ? "var(--accent2)" : "var(--muted2)" }}>{p.fecha_entrega ? fmtDate(p.fecha_entrega) : "—"}</td>
                      <td>{p.remito_url
                        ? <a href={p.remito_url} target="_blank" rel="noreferrer" onClick={e => e.stopPropagation()} style={{ fontSize: 11, color: "var(--blue)" }}> {p.nro_remito || "Ver"}</a>
                        : <span style={{ fontSize: 11, color: "var(--muted2)" }}>{p.nro_remito || "—"}</span>
                      }</td>
                      <td style={{ fontSize: 11, color: "var(--muted)", maxWidth: 150, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{p.tracker_notas || "—"}</td>
                      <td style={{ textAlign: "center" }} onClick={e => e.stopPropagation()}>
                        <button
                          className="btn btn-danger btn-sm"
                          onClick={e => handleEliminar(e, p)}
                          disabled={eliminando === p.id}
                          title="Eliminar pedido"
                        >
                          {eliminando === p.id ? "..." : "✕"}
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      }
      {selected && <ModalTrackerEditar pedido={selected} onClose={() => setSelected(null)} onSave={(updated) => { setSelected(null); setPedidos(prev => prev.map(p => p.id === updated.id ? { ...p, ...updated } : p)); }} notify={notify} />}
    </div>
  );
}

//  FORM: STOCK VUELTA A PUERTO
function FormStockVuelta({ catalogoInicial, solicitantes = [], onSave, onCancel, notify }) {
  const [catalogo] = useState(catalogoInicial || []);
  const [saving, setSaving] = useState(false);
  const [cabecera, setCabecera] = useState({
    base_buque: "",
    solicitado_por: "",
    fecha: new Date().toISOString().split("T")[0],
    observaciones: "",
  });
  const [items, setItems] = useState(() => catalogo.map(c => ({
    catalogo_id: c.id,
    descripcion: c.descripcion,
    categoria: c.categoria,
    temperatura: c.temperatura || "",
    unidad_analisis: c.unidad_analisis || "Kg",
    stock: "",
  })));
  const [filtroCateg, setFiltroCateg] = useState("");
  const [busqueda, setBusqueda] = useState("");

  const setCab = (k, v) => setCabecera(c => ({ ...c, [k]: v }));
  const setStock = (id, v) => setItems(prev => prev.map(it => it.catalogo_id === id ? { ...it, stock: v } : it));

  const categorias = [...new Set(catalogo.map(c => c.categoria))].sort();
  const itemsFiltrados = items.filter(it => {
    if (filtroCateg && it.categoria !== filtroCateg) return false;
    if (busqueda && !it.descripcion.toLowerCase().includes(busqueda.toLowerCase())) return false;
    return true;
  });
  const completados = items.filter(it => it.stock !== "" && it.stock != null).length;

  const handleGuardar = async () => {
    if (!cabecera.base_buque || !cabecera.solicitado_por || !cabecera.fecha) {
      alert("Completá Base/Buque, Solicitante y Fecha");
      return;
    }
    setSaving(true);
    try {
      const itemsAGuardar = items
        .filter(it => it.stock !== "" && it.stock != null)
        .map(it => ({
          catalogo_id: it.catalogo_id,
          descripcion: it.descripcion,
          categoria: it.categoria,
          unidad_analisis: it.unidad_analisis,
          stock: parseFloat(it.stock) || 0,
        }));
      if (itemsAGuardar.length === 0) { alert("Cargá el stock de al menos un ítem"); setSaving(false); return; }
      await onSave(cabecera, itemsAGuardar);
    } catch (e) {
      notify("Error: " + e.message, "error");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div>
      <div className="card">
        <div className="card-title">Datos de la vuelta a puerto</div>
        <div className="form-grid-3">
          <FG label="Base / Buque *">
            <select value={cabecera.base_buque} onChange={e => setCab("base_buque", e.target.value)}>
              <option value="">Seleccionar...</option>
              {BASES.map(b => <option key={b}>{b}</option>)}
            </select>
          </FG>
          <FG label="Solicitante *">
            <select value={cabecera.solicitado_por} onChange={e => setCab("solicitado_por", e.target.value)}>
              <option value="">Seleccionar...</option>
              {solicitantes.map(s => <option key={s.id} value={s.nombre}>{s.nombre}</option>)}
            </select>
          </FG>
          <FG label="Fecha *"><input type="date" value={cabecera.fecha} onChange={e => setCab("fecha", e.target.value)} /></FG>
        </div>
        <FG label="Observaciones"><textarea value={cabecera.observaciones} onChange={e => setCab("observaciones", e.target.value)} placeholder="Notas adicionales..." /></FG>
      </div>

      <div className="filter-row" style={{ marginBottom: 12 }}>
        <input className="filter-input" placeholder=" Buscar ítem..." value={busqueda} onChange={e => setBusqueda(e.target.value)} />
        <select className="filter-select" value={filtroCateg} onChange={e => setFiltroCateg(e.target.value)}>
          <option value="">Todas las categorías</option>
          {categorias.map(c => <option key={c}>{c}</option>)}
        </select>
        {(filtroCateg || busqueda) && <button className="btn btn-ghost btn-sm" onClick={() => { setFiltroCateg(""); setBusqueda(""); }}>✕</button>}
        <span style={{ marginLeft: "auto", fontFamily: "var(--mono)", fontSize: 11, color: "var(--muted)" }}>{completados} de {items.length} cargados</span>
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden", marginBottom: 16 }}>
        <div className="table-wrap">
          <table className="items-edit">
            <thead>
              <tr>
                <th>Temp.</th>
                <th>Categoría</th>
                <th>Descripción</th>
                <th>Unidad análisis</th>
                <th style={{ width: 110 }}>Stock a bordo</th>
              </tr>
            </thead>
            <tbody>
              {itemsFiltrados.map(it => {
                const cargado = it.stock !== "" && it.stock != null;
                return (
                  <tr key={it.catalogo_id} style={{ background: cargado ? "#F0FDF4" : "inherit" }}>
                    <td><TempBadge temp={it.temperatura} /></td>
                    <td style={{ fontSize: 11, color: "var(--muted)" }}>{it.categoria}</td>
                    <td style={{ fontSize: 12 }}>{it.descripcion}</td>
                    <td style={{ fontSize: 11, color: "var(--muted)" }}>{it.unidad_analisis}</td>
                    <td>
                      <input
                        type="number"
                        step="0.001"
                        min="0"
                        value={it.stock}
                        placeholder="0"
                        onChange={e => setStock(it.catalogo_id, e.target.value)}
                        style={{
                          width: 90, textAlign: "right",
                          fontWeight: cargado ? 700 : 400,
                          background: cargado ? "#DCFCE7" : "var(--surface)",
                          border: `1px solid ${cargado ? "#86EFAC" : "var(--border2)"}`,
                          borderRadius: "var(--r)", fontFamily: "var(--mono)", fontSize: 12,
                          padding: "4px 8px", outline: "none",
                        }}
                      />
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      <div className="form-footer-actions" style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <button className="btn btn-ghost" onClick={onCancel}>Cancelar</button>
        <button className="btn btn-primary" onClick={handleGuardar} disabled={saving}>{saving ? "Guardando..." : "Guardar registro"}</button>
      </div>
    </div>
  );
}

//  MODAL: DETALLE STOCK VUELTA A PUERTO
function ModalStockVueltaDetalle({ registro, onClose }) {
  const items = registro.viveres_stock_vuelta_items || [];
  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal modal-lg">
        <div className="mhdr">
          <div>
            <div className="mtitle"> {registro.base_buque} — Stock a la vuelta a puerto</div>
            <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 4 }}>
              {registro.solicitado_por} · {fmtDate(registro.fecha)}
            </div>
          </div>
          <button className="mclose" onClick={onClose}>✕</button>
        </div>
        <div className="mbody">
          {registro.observaciones && <div className="info-box accent mb12" style={{ fontSize: 12 }}>{registro.observaciones}</div>}
          <div className="table-wrap">
            <table className="items-edit">
              <thead>
                <tr><th>Categoría</th><th>Descripción</th><th>Unidad</th><th style={{ width: 90 }}>Stock</th></tr>
              </thead>
              <tbody>
                {items.map(it => (
                  <tr key={it.id}>
                    <td style={{ fontSize: 11, color: "var(--muted)" }}>{it.categoria}</td>
                    <td style={{ fontSize: 12 }}>{it.descripcion}</td>
                    <td style={{ fontSize: 11, color: "var(--muted)" }}>{it.unidad_analisis}</td>
                    <td className="text-mono" style={{ fontSize: 12, fontWeight: 600 }}>{fmt(it.stock)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
        <div className="mftr"><button className="btn btn-ghost" onClick={onClose}>Cerrar</button></div>
      </div>
    </div>
  );
}

//  PAGE: STOCK VUELTA A PUERTO
function PageStockVuelta({ notify }) {
  const [registros, setRegistros] = useState([]);
  const [catalogo, setCatalogo] = useState([]);
  const [solicitantes, setSolicitantes] = useState([]);
  const [loading, setLoading] = useState(true);
  const [mostrarForm, setMostrarForm] = useState(false);
  const [detalle, setDetalle] = useState(null);
  const [eliminando, setEliminando] = useState(null);
  const [filtroBase, setFiltroBase] = useState("");

  const load = useCallback(async () => {
    setLoading(true);
    // Cargamos catálogo/solicitantes por separado del historial: si el historial
    // falla (por ejemplo, porque todavía no se crearon las tablas en Supabase),
    // igual queremos que el catálogo completo esté disponible para el registro nuevo.
    try {
      const [cat, sol] = await Promise.all([api.getCatalogo(), api.getSolicitantes()]);
      setCatalogo(cat); setSolicitantes(sol);
    } catch (e) {
      notify("Error al cargar el catálogo: " + e.message, "error");
    }
    try {
      const regs = await api.getStockVuelta();
      setRegistros(regs);
    } catch (e) {
      notify("Error al cargar el historial de stock vuelta a puerto: " + e.message, "error");
    } finally {
      setLoading(false);
    }
  }, [notify]);

  useEffect(() => { load(); }, [load]);

  const handleGuardar = async (cabecera, items) => {
    await api.crearStockVuelta(cabecera, items);
    notify("Stock a la vuelta a puerto guardado", "success");
    setMostrarForm(false);
    load();
  };

  const handleEliminar = async (r) => {
    if (!window.confirm(`¿Eliminar el registro de ${r.base_buque} del ${fmtDate(r.fecha)}?`)) return;
    setEliminando(r.id);
    try {
      await api.eliminarStockVuelta(r.id);
      notify("Registro eliminado", "warn");
      setRegistros(prev => prev.filter(x => x.id !== r.id));
    } catch (e) {
      notify("Error: " + e.message, "error");
    } finally {
      setEliminando(null);
    }
  };

  const bases = [...new Set(registros.map(r => r.base_buque).filter(Boolean))].sort();
  const filtrados = registros.filter(r => !filtroBase || r.base_buque === filtroBase);

  if (mostrarForm) {
    return (
      <FormStockVuelta
        catalogoInicial={catalogo}
        solicitantes={solicitantes}
        onSave={handleGuardar}
        onCancel={() => setMostrarForm(false)}
        notify={notify}
      />
    );
  }

  return (
    <div>
      <div className="filter-row mb12">
        <select className="filter-select" value={filtroBase} onChange={e => setFiltroBase(e.target.value)}>
          <option value="">Todos los barcos</option>
          {bases.map(b => <option key={b}>{b}</option>)}
        </select>
        {filtroBase && <button className="btn btn-ghost btn-sm" onClick={() => setFiltroBase("")}>✕</button>}
        <span style={{ marginLeft: "auto", fontFamily: "var(--mono)", fontSize: 11, color: "var(--muted)" }}>{filtrados.length} de {registros.length}</span>
        <button
          className="btn btn-primary btn-sm"
          onClick={() => setMostrarForm(true)}
          disabled={loading || catalogo.length === 0}
          title={!loading && catalogo.length === 0 ? "No hay ítems en el catálogo. Cargalos primero en la sección Catálogo." : undefined}
        >
          + Nuevo registro
        </button>
      </div>

      {!loading && catalogo.length === 0 && (
        <div className="info-box warn mb12" style={{ fontSize: 12 }}>
          El catálogo está vacío: no hay ítems para cargar stock. Agregá productos en la sección <strong>Catálogo</strong> primero.
        </div>
      )}

      {loading ? <div className="loading"><span className="spin">◌</span></div> :
        filtrados.length === 0 ? <div className="empty-state"><div style={{ fontSize: 28, marginBottom: 8 }}></div>Sin registros de stock a la vuelta a puerto</div> :
        <div className="card" style={{ padding: 0, overflow: "hidden" }}>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Base/Barco</th>
                  <th>Solicitante</th>
                  <th>Fecha</th>
                  <th>Ítems cargados</th>
                  <th>Notas</th>
                  <th style={{ width: 90, textAlign: "center" }}>Acciones</th>
                </tr>
              </thead>
              <tbody>
                {filtrados.map(r => (
                  <tr key={r.id} className="click" onClick={() => setDetalle(r)}>
                    <td style={{ fontWeight: 600, fontSize: 12 }}>{r.base_buque}</td>
                    <td style={{ fontSize: 12 }}>{r.solicitado_por}</td>
                    <td className="text-mono" style={{ fontSize: 11, color: "var(--muted)" }}>{fmtDate(r.fecha)}</td>
                    <td className="text-mono" style={{ fontSize: 11, color: "var(--muted)" }}>{(r.viveres_stock_vuelta_items || []).length}</td>
                    <td style={{ fontSize: 11, color: "var(--muted)", maxWidth: 150, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.observaciones || "—"}</td>
                    <td style={{ textAlign: "center" }} onClick={e => e.stopPropagation()}>
                      <button className="btn btn-danger btn-sm" onClick={() => handleEliminar(r)} disabled={eliminando === r.id} title="Eliminar registro">
                        {eliminando === r.id ? "..." : "✕"}
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      }

      {detalle && <ModalStockVueltaDetalle registro={detalle} onClose={() => setDetalle(null)} />}
    </div>
  );
}

//  PAGE: INBOX
function PageInbox({ notify, onNeedRefresh }) {
  const [pedidos, setPedidos] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState(null);
  const [eliminando, setEliminando] = useState(null);

  const load = useCallback(async () => {
    setLoading(true);
    try { setPedidos(await api.getPedidos({ status: "enviado" })); }
    finally { setLoading(false); }
  }, []);

  useEffect(() => { load(); }, [load]);

  const handleEliminar = async (e, p) => {
    e.stopPropagation();
    if (!window.confirm(`¿Eliminar el pedido de ${p.base_buque}? Esta acción no se puede deshacer.`)) return;
    setEliminando(p.id);
    try {
      await api.eliminarPedido(p.id);
      notify("Pedido eliminado", "warn");
      load();
      onNeedRefresh();
    } catch (err) {
      notify("Error al eliminar: " + err.message, "error");
    } finally {
      setEliminando(null);
    }
  };

  return (
    <div>
      {loading ? <div className="loading"><span className="spin">◌</span></div> :
        pedidos.length === 0
          ? <div className="empty-state"><div style={{ fontSize: 28, marginBottom: 8 }}></div>Sin pedidos pendientes</div>
          : pedidos.map(p => {
              const cnt = (p.viveres_pedido_items || []).filter(it => it.cantidad_pedida > 0).length;
              return (
                <div key={p.id} className="req-row unread" onClick={() => setSelected(p)}>
                  <div className="flex-gap mb8">
                    <span className="text-mono" style={{ fontSize: 11, color: "var(--accent)" }}>{fmtDate(p.fecha_pedido)}</span>
                    <span className="badge b-blue">Víveres</span>
                    <span style={{ marginLeft: "auto", fontSize: 10, color: "var(--muted)" }}>PL Offshore</span>
                  </div>
                  <div className="req-title"> {p.base_buque} — {p.pax} PAX × {p.dias} días</div>
                  <div className="req-meta">
                    <span>{p.solicitado_por}</span><span>·</span><span>{cnt} ítems</span>
                    {p.fecha_necesaria && <><span>·</span><span style={{ color: "var(--warn)" }}>Necesario: {fmtDate(p.fecha_necesaria)}</span></>}
                  </div>
                  <div className="req-row-actions" onClick={e => e.stopPropagation()}>
                    <button
                      className="btn btn-danger btn-sm"
                      onClick={e => handleEliminar(e, p)}
                      disabled={eliminando === p.id}
                    >
                      {eliminando === p.id ? "..." : "✕ Eliminar"}
                    </button>
                  </div>
                </div>
              );
            })
      }
      {selected && (
        <ModalRevisar
          pedido={selected}
          onClose={() => setSelected(null)}
          onActualizado={() => { setSelected(null); notify("Pedido actualizado", "success"); load(); onNeedRefresh(); }}
          notify={notify}
        />
      )}
    </div>
  );
}

//  PAGE: HISTORIAL 
function PageHistorial({ onNuevo, notify }) {
  const [pedidos, setPedidos] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState(null);
  const [eliminando, setEliminando] = useState(null);

  const load = useCallback(async () => {
    setLoading(true);
    try { setPedidos(await api.getPedidos()); }
    catch (err) { notify("Error al cargar pedidos: " + (err?.message || "desconocido"), "error"); }
    finally { setLoading(false); }
  }, [notify]);

  useEffect(() => { load(); }, [load]);

  const handleEliminar = async (e, p) => {
    e.stopPropagation();
    if (!window.confirm(`¿Eliminar el pedido de ${p.base_buque} del ${fmtDate(p.fecha_pedido)} (${p.pax} PAX × ${p.dias} días)? Se borran también sus ítems. Esta acción no se puede deshacer.`)) return;
    setEliminando(p.id);
    const backup = pedidos;
    // optimistic update
    setPedidos(prev => prev.filter(x => x.id !== p.id));
    try {
      await api.eliminarPedido(p.id);
      notify("Pedido eliminado", "warn");
    } catch (err) {
      setPedidos(backup); // revert
      notify("Error al eliminar: " + (err?.message || "desconocido"), "error");
    } finally {
      setEliminando(null);
    }
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 12 }}>
        <div style={{ fontSize: 13, color: "var(--muted)" }}>{pedidos.length} pedidos registrados</div>
        <button className="btn btn-primary btn-sm" onClick={onNuevo}>+ Nuevo pedido</button>
      </div>
      {loading ? <div className="loading"><span className="spin">◌</span></div> :
        pedidos.length === 0 ? <div className="empty-state"><div style={{ fontSize: 28, marginBottom: 8 }}></div>Sin pedidos</div> :
        pedidos.map(p => {
          const s = STATUS_PEDIDO[p.status] || { label: p.status, color: "b-gray" };
          const cnt = (p.viveres_pedido_items || []).filter(it => it.cantidad_pedida > 0).length;
          return <div key={p.id} className="req-row" onClick={() => setSelected(p)}>
            <div className="flex-between mb8"><div className="flex-gap"><span className="text-mono" style={{ fontSize: 11, color: "var(--accent)" }}>{fmtDate(p.fecha_pedido)}</span><span className={`badge ${s.color}`}>{s.label}</span></div><span style={{ fontSize: 10, color: "var(--muted)" }}>PL Offshore</span></div>
            <div className="req-title">{p.base_buque} — {p.pax} PAX × {p.dias} días</div>
            <div className="req-meta"><span>{p.solicitado_por}</span><span>·</span><span>{cnt} ítems</span>{p.fecha_necesaria && <><span>·</span><span style={{ color: "var(--warn)" }}>Nec: {fmtDate(p.fecha_necesaria)}</span></>}</div>
            <div className="req-row-actions" onClick={e => e.stopPropagation()}>
              <button
                className="btn btn-danger btn-sm"
                onClick={e => handleEliminar(e, p)}
                disabled={eliminando === p.id}
                title="Eliminar pedido"
              >
                {eliminando === p.id ? "..." : "✕ Eliminar"}
              </button>
            </div>
          </div>;
        })
      }
      {selected && <ModalRevisar pedido={selected} onClose={() => setSelected(null)} onActualizado={() => { setSelected(null); load(); }} notify={notify} />}
    </div>
  );
}

//  PAGE: CATÁLOGO 
const CATEGORIAS_CATALOGO = ["Almacén","Bebidas","Carnicería","Electro","Fiambrería","Frutas","Huevos","Lácteos","Limpieza","Pan","Pastas","Pescadería","Quesos","Snack y Postres","Verduras","Otro"];

//  PAGE SOLICITANTES 
function PageSolicitantes({ notify }) {
  const [solicitantes, setSolicitantes] = useState([]);
  const [loading, setLoading] = useState(true);
  const [nuevo, setNuevo] = useState("");
  const [saving, setSaving] = useState(false);
  const [eliminandoId, setEliminandoId] = useState(null);

  useEffect(() => {
    api.getSolicitantes()
      .then(d => { setSolicitantes(d); setLoading(false); })
      .catch(e => { notify("Error: " + e.message, "error"); setLoading(false); });
  }, [notify]);

  const handleAgregar = async () => {
    const nombre = nuevo.trim();
    if (!nombre) return;
    if (solicitantes.some(s => s.nombre.toLowerCase() === nombre.toLowerCase())) {
      notify("Ese nombre ya existe", "warn");
      return;
    }
    setSaving(true);
    try {
      const data = await api.crearSolicitante(nombre);
      if (data) {
        setSolicitantes(prev => [...prev, data].sort((a, b) => a.nombre.localeCompare(b.nombre, "es")));
      }
      setNuevo("");
      notify("Solicitante agregado", "success");
    } catch (e) {
      notify("Error: " + e.message, "error");
    } finally {
      setSaving(false);
    }
  };

  const handleEliminar = async (s) => {
    if (!window.confirm(`¿Quitar a "${s.nombre}" de la lista de solicitantes?`)) return;
    setEliminandoId(s.id);
    const previo = solicitantes;
    setSolicitantes(prev => prev.filter(x => x.id !== s.id)); // optimista
    try {
      await api.eliminarSolicitante(s.id);
      notify("Solicitante eliminado", "warn");
    } catch (e) {
      setSolicitantes(previo); // revert
      notify("Error: " + e.message, "error");
    } finally {
      setEliminandoId(null);
    }
  };

  const handleKey = (e) => { if (e.key === "Enter") handleAgregar(); };

  return (
    <div>
      <div className="card" style={{ maxWidth: 640 }}>
        <div className="card-title">Agregar solicitante</div>
        <div className="info-box accent mb12" style={{ fontSize: 12 }}>
          Los nombres cargados acá aparecen en el selector "Solicitado por" al crear un pedido. Estandarizá los nombres para mantener el historial consistente.
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
          <div className="fg" style={{ flex: 1 }}>
            <label>Nombre y apellido</label>
            <input value={nuevo} onChange={e => setNuevo(e.target.value)} onKeyDown={handleKey} placeholder="Ej: Juan Pérez" />
          </div>
          <button className="btn btn-primary" onClick={handleAgregar} disabled={saving || !nuevo.trim()}>
            {saving ? "Guardando..." : "+ Agregar"}
          </button>
        </div>
      </div>

      {loading ? <div className="loading"><span className="spin">◌</span></div> :
        <div className="card" style={{ padding: 0, overflow: "hidden", maxWidth: 640 }}>
          <div className="card-title" style={{ padding: "16px 24px 0" }}>
            Solicitantes habilitados
            <span style={{ fontFamily: "var(--mono)", fontSize: 11, color: "var(--muted)", textTransform: "none", letterSpacing: 0 }}>{solicitantes.length} activo{solicitantes.length !== 1 ? "s" : ""}</span>
          </div>
          {solicitantes.length === 0 ? (
            <div className="manual-empty" style={{ margin: 24 }}>
              <div style={{ fontSize: 14, fontWeight: 600, color: "var(--navy)" }}>Sin solicitantes cargados</div>
              <div style={{ fontSize: 12, color: "var(--muted)" }}>Agregá el primer nombre usando el formulario de arriba.</div>
            </div>
          ) : (
            <div className="table-wrap">
              <table className="items-edit">
                <thead>
                  <tr>
                    <th>Nombre</th>
                    <th style={{ width: 70 }}></th>
                  </tr>
                </thead>
                <tbody>
                  {solicitantes.map(s => {
                    const eliminando = eliminandoId === s.id;
                    return (
                      <tr key={s.id}>
                        <td style={{ fontSize: 13, fontWeight: 500 }}>{s.nombre}</td>
                        <td>
                          <button
                            onClick={() => handleEliminar(s)}
                            disabled={eliminando}
                            title="Quitar solicitante"
                            style={{ background: "none", border: "none", color: "var(--muted2)", cursor: "pointer", fontSize: 14, padding: "3px 5px", borderRadius: 4, opacity: eliminando ? 0.5 : 1, transition: "color .12s" }}
                            onMouseEnter={e => e.currentTarget.style.color = "var(--danger)"}
                            onMouseLeave={e => e.currentTarget.style.color = "var(--muted2)"}
                          >
                            {eliminando ? "..." : "✕"}
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      }
    </div>
  );
}

function PageCatalogo({ notify }) {
  const [catalogo, setCatalogo] = useState([]);
  const [loading, setLoading] = useState(true);
  const [busqueda, setBusqueda] = useState("");
  const [filtroCateg, setFiltroCateg] = useState("");
  const [modal, setModal] = useState(false);
  const [saving, setSaving] = useState(false);
  const [savingId, setSavingId] = useState(null);
  const [eliminandoId, setEliminandoId] = useState(null);
  const [editados, setEditados] = useState({}); // id -> campos modificados
  const [form, setForm] = useState({ codigo: "", categoria: "Almacén", subcategoria: "", temperatura: "Seco", descripcion: "", unidad: "Unidad", unidad_analisis: "Kg", volumen_peso: "1", stock: "0" });

  useEffect(() => { api.getCatalogo().then(d => { setCatalogo(d); setLoading(false); }); }, []);

  const categorias = [...new Set(catalogo.map(c => c.categoria))].sort();
  const filtrado = catalogo.filter(c => {
    if (filtroCateg && c.categoria !== filtroCateg) return false;
    if (busqueda && !c.descripcion.toLowerCase().includes(busqueda.toLowerCase()) && !c.codigo?.toLowerCase().includes(busqueda.toLowerCase())) return false;
    return true;
  });

  // Edición inline
  const setcampo = (id, campo, valor) => {
    setEditados(prev => ({ ...prev, [id]: { ...(prev[id] || {}), [campo]: valor } }));
    setCatalogo(prev => prev.map(c => c.id === id ? { ...c, [campo]: valor } : c));
  };

  const getVal = (c, campo) => editados[c.id]?.[campo] !== undefined ? editados[c.id][campo] : c[campo];
  const tienecambios = (id) => !!editados[id] && Object.keys(editados[id]).length > 0;

  const handleGuardarFila = async (c) => {
    if (!tieneambios(c.id)) return;
    setSavingId(c.id);
    try {
      const cambios = editados[c.id];
      if (cambios.volumen_peso !== undefined) cambios.volumen_peso = parseFloat(cambios.volumen_peso) || 1;
      if (cambios.stock !== undefined) cambios.stock = parseFloat(cambios.stock) || 0;
      const { error } = await supabase.from("viveres_catalogo").update(cambios).eq("id", c.id);
      if (error) throw error;
      setEditados(prev => { const n = { ...prev }; delete n[c.id]; return n; });
      notify("Guardado", "success");
    } catch (e) {
      notify("Error: " + e.message, "error");
    } finally {
      setSavingId(null);
    }
  };

  const handleEliminarFila = async (c) => {
    if (!window.confirm(`¿Eliminar "${c.descripcion}" del catálogo?`)) return;
    setEliminandoId(c.id);
    try {
      const { error } = await supabase.from("viveres_catalogo").delete().eq("id", c.id);
      if (error) throw error;
      setCatalogo(prev => prev.filter(x => x.id !== c.id));
      notify("Ítem eliminado", "warn");
    } catch (e) {
      notify("Error: " + e.message, "error");
    } finally {
      setEliminandoId(null);
    }
  };

  // Nuevo ítem
  const setF = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const handleAgregar = async () => {
    if (!form.descripcion.trim()) return alert("La descripción es obligatoria");
    setSaving(true);
    try {
      const { data, error } = await supabase.from("viveres_catalogo").insert([{ ...form, volumen_peso: parseFloat(form.volumen_peso) || 1, stock: parseFloat(form.stock) || 0, activo: true }]).select().single();
      if (error) throw error;
      setCatalogo(prev => [...prev, data]);
      setModal(false);
      setForm({ codigo: "", categoria: "Almacén", subcategoria: "", temperatura: "Seco", descripcion: "", unidad: "Unidad", unidad_analisis: "Kg", volumen_peso: "1", stock: "0" });
      notify("Ítem agregado", "success");
    } catch (e) { alert("Error: " + e.message); }
    finally { setSaving(false); }
  };

  const inStyle = {
    background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 4,
    color: "var(--text)", fontFamily: "var(--mono)", fontSize: 11, padding: "3px 6px",
    outline: "none", width: "100%",
  };
  const inStyleMod = { ...inStyle, background: "#FEF9C3", border: "1px solid #FDE68A", fontWeight: 600 };

  // helper para saber si un campo fue modificado
  const mod = (c, campo) => editados[c.id]?.[campo] !== undefined;
  // helper para guardar (typo fix)
  const tieneambios = (id) => !!editados[id] && Object.keys(editados[id]).length > 0;

  return (
    <div>
      <div className="filter-row mb12">
        <input className="filter-input" placeholder=" Buscar..." value={busqueda} onChange={e => setBusqueda(e.target.value)} style={{ minWidth: 250 }} />
        <select className="filter-select" value={filtroCateg} onChange={e => setFiltroCateg(e.target.value)}>
          <option value="">Todas las categorías</option>
          {categorias.map(c => <option key={c}>{c}</option>)}
        </select>
        {(busqueda || filtroCateg) && <button className="btn btn-ghost btn-sm" onClick={() => { setBusqueda(""); setFiltroCateg(""); }}>✕</button>}
        <span style={{ marginLeft: "auto", fontFamily: "var(--mono)", fontSize: 11, color: "var(--muted)" }}>{filtrado.length} de {catalogo.length}</span>
        <button className="btn btn-primary btn-sm" onClick={() => setModal(true)}>+ Agregar ítem</button>
      </div>

      {Object.keys(editados).length > 0 && (
        <div className="info-box warn mb12" style={{ fontSize: 12, display: "flex", alignItems: "center", gap: 8 }}>
          <span></span>
          <span><strong>{Object.keys(editados).length} ítem{Object.keys(editados).length !== 1 ? "s" : ""} con cambios sin guardar.</strong> Usá el botón  de cada fila para guardar.</span>
        </div>
      )}

      {loading ? <div className="loading"><span className="spin">◌</span></div> :
        <div className="card" style={{ padding: 0, overflow: "hidden" }}>
          <div className="table-wrap">
            <table className="items-edit">
              <thead>
                <tr>
                  <th>Código</th>
                  <th>Categoría</th>
                  <th>Temp.</th>
                  <th>Descripción</th>
                  <th>Unidad pedido</th>
                  <th>Unidad análisis</th>
                  <th style={{ width: 70 }}>Stock</th>
                  <th>Vol/Peso</th>
                  <th style={{ width: 70 }}></th>
                </tr>
              </thead>
              <tbody>
                {filtrado.map(c => {
                  const hayC = tieneambios(c.id);
                  const guardando = savingId === c.id;
                  const eliminando = eliminandoId === c.id;
                  return (
                    <tr key={c.id} style={{ background: hayC ? "#FFFBEB" : "inherit" }}>
                      <td>
                        <input
                          value={getVal(c, "codigo") || ""}
                          onChange={e => setcampo(c.id, "codigo", e.target.value)}
                          style={mod(c, "codigo") ? inStyleMod : inStyle}
                          placeholder="—"
                        />
                      </td>
                      <td>
                        <select
                          value={getVal(c, "categoria") || ""}
                          onChange={e => setcampo(c.id, "categoria", e.target.value)}
                          style={mod(c, "categoria") ? { ...inStyleMod, fontFamily: "var(--sans)" } : { ...inStyle, fontFamily: "var(--sans)" }}
                        >
                          {CATEGORIAS_CATALOGO.map(cat => <option key={cat}>{cat}</option>)}
                        </select>
                      </td>
                      <td>
                        <select
                          value={getVal(c, "temperatura") || "Seco"}
                          onChange={e => setcampo(c.id, "temperatura", e.target.value)}
                          style={mod(c, "temperatura") ? { ...inStyleMod, fontFamily: "var(--sans)" } : { ...inStyle, fontFamily: "var(--sans)" }}
                        >
                          <option>Seco</option>
                          <option>Refrigerado</option>
                          <option>Congelado</option>
                        </select>
                      </td>
                      <td style={{ minWidth: 180 }}>
                        <input
                          value={getVal(c, "descripcion") || ""}
                          onChange={e => setcampo(c.id, "descripcion", e.target.value)}
                          style={mod(c, "descripcion") ? { ...inStyleMod, fontWeight: 700 } : { ...inStyle, fontWeight: 500 }}
                        />
                      </td>
                      <td>
                        <select
                          value={getVal(c, "unidad") || "Unidad"}
                          onChange={e => setcampo(c.id, "unidad", e.target.value)}
                          style={mod(c, "unidad") ? { ...inStyleMod, fontFamily: "var(--sans)" } : { ...inStyle, fontFamily: "var(--sans)" }}
                        >
                          {UNIDADES_PEDIDO.map(u => <option key={u}>{u}</option>)}
                        </select>
                      </td>
                      <td>
                        <select
                          value={getVal(c, "unidad_analisis") || "Kg"}
                          onChange={e => setcampo(c.id, "unidad_analisis", e.target.value)}
                          style={mod(c, "unidad_analisis") ? { ...inStyleMod, fontFamily: "var(--sans)" } : { ...inStyle, fontFamily: "var(--sans)" }}
                        >
                          {UNIDADES_ANALISIS.map(u => <option key={u}>{u}</option>)}
                        </select>
                      </td>
                      <td style={{ width: 70 }}>
                        <input
                          type="number"
                          step="1"
                          min="0"
                          value={getVal(c, "stock") ?? 0}
                          onChange={e => setcampo(c.id, "stock", e.target.value)}
                          style={mod(c, "stock") ? inStyleMod : inStyle}
                        />
                      </td>
                      <td style={{ width: 80 }}>
                        <input
                          type="number"
                          step="0.001"
                          min="0"
                          value={getVal(c, "volumen_peso") ?? 1}
                          onChange={e => setcampo(c.id, "volumen_peso", e.target.value)}
                          style={mod(c, "volumen_peso") ? inStyleMod : inStyle}
                        />
                      </td>
                      <td>
                        <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
                          {hayC && (
                            <button
                              onClick={() => handleGuardarFila(c)}
                              disabled={guardando}
                              title="Guardar cambios"
                              style={{ background: "var(--accent2)", border: "none", borderRadius: 4, color: "#fff", cursor: "pointer", fontSize: 13, padding: "3px 7px", opacity: guardando ? 0.5 : 1 }}
                            >
                              {guardando ? "..." : ""}
                            </button>
                          )}
                          <button
                            onClick={() => handleEliminarFila(c)}
                            disabled={eliminando}
                            title="Eliminar ítem"
                            style={{ background: "none", border: "none", color: "var(--muted2)", cursor: "pointer", fontSize: 14, padding: "3px 5px", borderRadius: 4, opacity: eliminando ? 0.5 : 1, transition: "color .12s" }}
                            onMouseEnter={e => e.currentTarget.style.color = "var(--danger)"}
                            onMouseLeave={e => e.currentTarget.style.color = "var(--muted2)"}
                          >
                            {eliminando ? "..." : "✕"}
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      }

      {modal && (
        <div className="overlay" onClick={e => e.target === e.currentTarget && setModal(false)}>
          <div className="modal" style={{ maxWidth: 600 }}>
            <div className="mhdr"><div className="mtitle">Agregar ítem al catálogo</div><button className="mclose" onClick={() => setModal(false)}>✕</button></div>
            <div className="mbody">
              <div className="form-grid">
                <FG label="Código"><input value={form.codigo} onChange={e => setF("codigo", e.target.value)} placeholder="Ej: NAV001" /></FG>
                <FG label="Temperatura *"><select value={form.temperatura} onChange={e => setF("temperatura", e.target.value)}><option>Seco</option><option>Refrigerado</option><option>Congelado</option></select></FG>
                <FG label="Categoría *"><select value={form.categoria} onChange={e => setF("categoria", e.target.value)}>{CATEGORIAS_CATALOGO.map(c => <option key={c}>{c}</option>)}</select></FG>
                <FG label="Subcategoría"><input value={form.subcategoria} onChange={e => setF("subcategoria", e.target.value)} /></FG>
              </div>
              <FG label="Descripción *" full><input value={form.descripcion} onChange={e => setF("descripcion", e.target.value)} placeholder="Nombre completo del producto" /></FG>
              <div className="form-grid-3 mt12">
                <FG label="Unidad de pedido" hint="Cómo se pide al proveedor"><select value={form.unidad} onChange={e => setF("unidad", e.target.value)}>{UNIDADES_PEDIDO.map(u => <option key={u}>{u}</option>)}</select></FG>
                <FG label="Unidad de análisis" hint="Para el cálculo de dieta"><select value={form.unidad_analisis} onChange={e => setF("unidad_analisis", e.target.value)}>{UNIDADES_ANALISIS.map(u => <option key={u}>{u}</option>)}</select></FG>
                <FG label="Vol/Peso por unidad" hint="Ej: 1 lata = 0.170 Kg"><input type="number" step="0.001" min="0" value={form.volumen_peso} onChange={e => setF("volumen_peso", e.target.value)} placeholder="1" /></FG>
                <FG label="Stock" hint="Cantidad disponible actual"><input type="number" step="1" min="0" value={form.stock} onChange={e => setF("stock", e.target.value)} placeholder="0" /></FG>
              </div>
              {form.volumen_peso && parseFloat(form.volumen_peso) !== 1 && (
                <div className="info-box accent mt8" style={{ fontSize: 11 }}>Ejemplo: 3 {form.unidad} → {(3 * parseFloat(form.volumen_peso)).toFixed(3)} {form.unidad_analisis}</div>
              )}
            </div>
            <div className="mftr">
              <button className="btn btn-ghost" onClick={() => setModal(false)}>Cancelar</button>
              <button className="btn btn-primary" onClick={handleAgregar} disabled={saving}>{saving ? "Guardando..." : "Agregar"}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

//  LOGIN PAGE 
function LoginPage() {
  const [email, setEmail]     = useState("");
  const [pass, setPass]       = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError]     = useState("");

  const handleLogin = async () => {
    setLoading(true); setError("");
    try {
      const { error: e } = await supabase.auth.signInWithPassword({ email, password: pass });
      if (e) setError("Credenciales incorrectas. Verificá tu email y contraseña.");
    } catch {
      setError("Error de conexión. Verificá tu red e intentá nuevamente.");
    } finally {
      setLoading(false);
    }
  };

  const handleKey = (e) => { if (e.key === "Enter") handleLogin(); };

  const loginCSS = `
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
    .login-page{min-height:100vh;display:grid;grid-template-columns:minmax(0,1fr) 560px;background:#FFFFFF;font-family:'IBM Plex Sans',sans-serif;color:#0F1419;text-align:left}
    .login-bg-overlay,.login-bg-lines{display:none}
    .login-split{display:contents}
    .login-left{display:flex;flex-direction:column;justify-content:space-between;gap:48px;padding:56px 64px;background:#002247;border:0;text-align:left}
    .login-left-integra-wrap{margin:0}
    .login-left-integra-img{height:52px;width:auto;object-fit:contain;opacity:1;display:block}
    .login-left-divider{width:100%;height:1px;background:rgba(255,255,255,.14);margin:24px 0}
    .login-left-company{display:flex;align-items:center;gap:14px;margin:0}
    .login-left-company-logo{width:40px;height:40px;border-radius:4px;object-fit:contain;border:0;background:rgba(255,255,255,.14);padding:4px}
    .login-left-company-name{font:600 24px/1.25 'IBM Plex Sans',sans-serif;color:#fff;letter-spacing:0}
    .login-left-line{width:56px;height:3px;background:#F8BC05;margin:24px 0}
    .login-left-sub{font:400 15px/1.55 'IBM Plex Sans',sans-serif;color:rgba(255,255,255,.82);max-width:420px;font-style:normal}
    .login-right{width:auto;display:flex;align-items:center;justify-content:center;padding:56px 64px;background:#FFFFFF}
    .login-card{width:100%;max-width:420px;background:transparent;border:0;border-radius:0;padding:0;backdrop-filter:none;text-align:left}
    .login-card-eyebrow{font:500 11px/1.2 'IBM Plex Mono',monospace;letter-spacing:.08em;color:#4A5560;text-transform:uppercase;margin-bottom:12px}
    .login-card-title{font:600 24px/1.25 'IBM Plex Sans',sans-serif;color:#082F4E;margin-bottom:8px}
    .login-card-sub{font:400 15px/1.55 'IBM Plex Sans',sans-serif;color:#4A5560;letter-spacing:0;margin-bottom:28px;text-transform:none}
    .login-fg{display:flex;flex-direction:column;gap:6px;margin-bottom:16px}
    .login-fg label{font:500 11px/1.2 'IBM Plex Mono',monospace;color:#4A5560;letter-spacing:.08em;text-transform:uppercase}
    .login-fg input{border:1px solid #C9D0D6;border-radius:4px;height:40px;padding:0 12px;font:400 14px/1.2 'IBM Plex Sans',sans-serif;color:#0F1419;background:#FFFFFF;outline:none;transition:border-color 120ms cubic-bezier(.2,0,.38,.9)}
    .login-fg input::placeholder{color:#7A8792}
    .login-fg input:focus{border-width:2px;border-color:#002247;padding:0 11px}
    .login-btn{width:100%;height:44px;padding:0 16px;margin-top:24px;background:#F8BC05;color:#002247;border:none;border-radius:4px;font:600 15px/1.2 'IBM Plex Sans',sans-serif;cursor:pointer;transition:background-color 120ms cubic-bezier(.2,0,.38,.9);letter-spacing:0}
    .login-btn:hover{background:#DCA704}
    .login-btn:disabled{background:#E4E8EC;color:#7A8792;cursor:not-allowed}
    .login-error{background:#FFFFFF;color:#0F1419;border:1px solid #E4E8EC;border-left:3px solid #B3261E;border-radius:4px;padding:12px 16px;font:400 13px/1.45 'IBM Plex Sans',sans-serif;margin-bottom:16px}
    .login-footer{text-align:left;font:500 11px/1.2 'IBM Plex Mono',monospace;color:#4A5560;margin-top:32px;letter-spacing:.06em}
    .login-back{text-align:left;margin-top:12px;font:500 14px/1.2 'IBM Plex Sans',sans-serif;color:#002247;cursor:pointer}
    .login-back:hover{text-decoration:underline}
    @media(max-width:900px){
      .login-page{grid-template-columns:1fr}
      .login-left{padding:40px 24px;gap:32px}
      .login-left-integra-img{height:40px}
      .login-left-sub{max-width:100%}
      .login-right{padding:40px 24px}
    }
  
  `;

  return (
    <>
      <style>{loginCSS}</style>
      <div className="login-page">
        <div className="login-bg-lines" />
        <div className="login-bg-overlay" />
        <div className="login-split">
          <div className="login-left">
            <div className="login-left-integra-wrap">
              <img src="/integra-logo-white-noclaim.svg" alt="INTEGRA" className="login-left-integra-img" />
            </div>
            <div className="login-left-divider" />
            <div className="login-left-company">
              <img src="/PL.png" alt="PL Offshore" className="login-left-company-logo" />
              <div className="login-left-company-name">PL Offshore | Víveres</div>
            </div>
            <div className="login-left-line" />
            <div className="login-left-sub">We Find the Way, or We Make One.</div>
          </div>
          <div className="login-right">
            <div className="login-card">
              <div className="login-card-eyebrow">PL Offshore | Víveres</div>
              <div className="login-card-title">Acceso al portal</div>
              <div className="login-card-sub">Solo personal autorizado</div>
              {error && <div className="login-error">{error}</div>}
              <div className="login-fg">
                <label>Email</label>
                <input type="email" value={email} onChange={e => setEmail(e.target.value)} onKeyDown={handleKey} placeholder="usuario@paranalogistica.com.ar" autoFocus />
              </div>
              <div className="login-fg">
                <label>Contraseña</label>
                <input type="password" value={pass} onChange={e => setPass(e.target.value)} onKeyDown={handleKey} placeholder="••••••••" />
              </div>
              <button className="login-btn" onClick={handleLogin} disabled={loading || !email || !pass}>
                {loading ? "Ingresando..." : "Ingresar →"}
              </button>
              <div className="login-footer">PL Offshore · Acceso restringido</div>
              <div className="login-back" onClick={() => window.location.href = PORTAL_URL}>← Volver a Grupo PL</div>
            </div>
          </div>
        </div>
      </div>
    </>
  );
}

//  ROOT APP 
//  PAGE PIVOT 
function PagePivot() {
  const [pedidos,    setPedidos]    = useState([]);
  const [loading,    setLoading]    = useState(true);
  const [error,      setError]      = useState(null);
  const [filas,      setFilas]      = useState("categoria");
  const [columnas,   setColumnas]   = useState("pedido");
  const [metrica,    setMetrica]    = useState("cantidad");
  const [filtBuque,  setFiltBuque]  = useState("");
  const [filtEstado, setFiltEstado] = useState("");
  const [filtDesde,  setFiltDesde]  = useState("");
  const [filtHasta,  setFiltHasta]  = useState("");
  const [busqueda,   setBusqueda]   = useState("");
  const [expandidos, setExpandidos] = useState({});
  //  Multi-selección de pedidos 
  const [seleccionados, setSeleccionados] = useState(new Set()); // Set de ids
  const [modoSel, setModoSel] = useState(false); // false = todos los pedidos filtrados

  useEffect(() => {
    api.getPedidos({}).then(d => { setPedidos(d); setLoading(false); }).catch(e => { setError(e.message); setLoading(false); });
  }, []);

  const pedidosBase = useMemo(() => pedidos.filter(p => {
    if (filtBuque  && p.base_buque !== filtBuque)  return false;
    if (filtEstado && p.status     !== filtEstado)  return false;
    if (filtDesde  && (p.fecha_pedido||"") < filtDesde) return false;
    if (filtHasta  && (p.fecha_pedido||"") > filtHasta) return false;
    return true;
  }), [pedidos, filtBuque, filtEstado, filtDesde, filtHasta]);

  // Si modoSel activo, solo los marcados; si no, todos los filtrados
  const pedidosFilt = useMemo(() =>
    modoSel && seleccionados.size > 0
      ? pedidosBase.filter(p => seleccionados.has(p.id))
      : pedidosBase,
  [pedidosBase, modoSel, seleccionados]);

  const buques  = useMemo(() => [...new Set(pedidos.map(p => p.base_buque).filter(Boolean))].sort(), [pedidos]);
  const estados = useMemo(() => [...new Set(pedidos.map(p => p.status).filter(Boolean))].sort(), [pedidos]);

  //  Motor pivot 
  const { tabla, filaKeys, colKeys } = useMemo(() => {
    const map = new Map();
    const eventos = [];

    pedidosFilt.forEach(p => {
      const pax = p.pax || 1, dias = p.dias || 1, paxDias = pax * dias;
      const colKey = columnas === "pedido"
        ? `${(p.fecha_pedido||"—").slice(0,10)} · ${p.base_buque||"?"}`
        : p.base_buque || "Sin buque";

      // Usamos lo AUTORIZADO por el comprador cuando ya está aprobado; si
      // todavía no se aprobó, no hay autorizado y se usa lo pedido.
      (p.viveres_pedido_items || []).filter(it => cantEfectiva(it) > 0).forEach(it => {
        const base = metrica === "volumen"
          ? cantEfectiva(it) * (it.volumen_peso||1)
          : cantEfectiva(it);
        const valor = columnas === "dia" ? base / dias : columnas === "pax_dia" ? base / paxDias : base;

        const filaKey = filas === "item"         ? (it.descripcion   || "—")
                      : filas === "subcategoria" ? (it.subcategoria  || "Sin subcategoría")
                      : filas === "buque"        ? (p.base_buque     || "—")
                      :                           (it.categoria      || "Sin categoría");

        if (busqueda && !filaKey.toLowerCase().includes(busqueda.toLowerCase()) &&
            !(it.descripcion||"").toLowerCase().includes(busqueda.toLowerCase())) return;

        if (!map.has(filaKey)) map.set(filaKey, new Map());
        const row = map.get(filaKey);
        row.set(colKey, (row.get(colKey)||0) + valor);
        eventos.push({ filaKey, colKey, it, p, valor });
      });
    });

    const filaKeys = [...map.keys()].sort((a,b) => {
      const tA = [...map.get(a).values()].reduce((s,v)=>s+v,0);
      const tB = [...map.get(b).values()].reduce((s,v)=>s+v,0);
      return tB - tA;
    });
    const colSet = new Set(); eventos.forEach(e => colSet.add(e.colKey));
    const colKeys = [...colSet].sort();
    return { tabla: map, filaKeys, colKeys };
  }, [pedidosFilt, filas, columnas, metrica, busqueda]);

  //  Desglose por ítem dentro de una fila 
  const buildSubFilas = useCallback((fk) => {
    const subMap = new Map();
    pedidosFilt.forEach(p => {
      const pax = p.pax||1, dias = p.dias||1, paxDias = pax*dias;
      const colKey = columnas === "pedido"
        ? `${(p.fecha_pedido||"—").slice(0,10)} · ${p.base_buque||"?"}`
        : p.base_buque || "Sin buque";
      (p.viveres_pedido_items||[]).filter(it => {
        if (cantEfectiva(it) <= 0) return false;
        if (filas === "categoria"    && (it.categoria    ||"Sin categoría")    !== fk) return false;
        if (filas === "subcategoria" && (it.subcategoria ||"Sin subcategoría") !== fk) return false;
        if (filas === "buque"        && (p.base_buque    ||"—")               !== fk) return false;
        return true;
      }).forEach(it => {
        const base = metrica === "volumen" ? cantEfectiva(it)*(it.volumen_peso||1) : cantEfectiva(it);
        const valor = columnas === "dia" ? base/dias : columnas === "pax_dia" ? base/paxDias : base;
        const desc = it.descripcion || "—";
        if (!subMap.has(desc)) subMap.set(desc, new Map());
        subMap.get(desc).set(colKey, (subMap.get(desc).get(colKey)||0) + valor);
      });
    });
    const subKeys = [...subMap.keys()].sort((a,b) => {
      const tA = [...(subMap.get(a)?.values()||[])].reduce((s,v)=>s+v,0);
      const tB = [...(subMap.get(b)?.values()||[])].reduce((s,v)=>s+v,0);
      return tB - tA;
    });
    return { map: subMap, keys: subKeys };
  }, [pedidosFilt, filas, columnas, metrica]);

  //  Helpers 
  const totalFila = (fk) => [...(tabla.get(fk)?.values()||[])].reduce((s,v)=>s+v,0);
  const maxTotal  = useMemo(() => Math.max(1, ...filaKeys.map(f => totalFila(f))), [filaKeys, tabla]);
  const maxCell   = useMemo(() => { let m=1; filaKeys.forEach(fk => colKeys.forEach(ck => { const v=tabla.get(fk)?.get(ck)||0; if(v>m) m=v; })); return m; }, [filaKeys,colKeys,tabla]);

  const heatBg = (v, max) => {
    if (!v) return "transparent";
    const pct = Math.min(v/max, 1);
    return `rgba(35,92,150,${0.07 + pct * 0.28})`;
  };

  const fmtVal = (v) => {
    if (!v || v < 0.0001) return <span style={{ color:"var(--muted2)", fontSize:11 }}>—</span>;
    const sfx = metrica==="volumen" ? " kg" : columnas==="dia" ? "/día" : columnas==="pax_dia" ? "/p·d" : "";
    const num  = v < 0.01 ? v.toFixed(4) : v < 10 ? v.toFixed(2) : v % 1 === 0 ? v.toFixed(0) : v.toFixed(1);
    return <span style={{ fontFamily:"var(--mono)", fontSize:12, fontWeight:600 }}>{num}<span style={{ fontSize:9, color:"var(--muted)", marginLeft:1 }}>{sfx}</span></span>;
  };

  const totalGlobal = filaKeys.reduce((s,f)=>s+totalFila(f),0);

  if (loading) return <div className="state-empty">Cargando historial de pedidos...</div>;
  if (error)   return <div className="state-empty" style={{color:"var(--danger)"}}>Error: {error}</div>;

  const FILA_OPT = [{v:"categoria",l:"Categoría"},{v:"item",l:"Ítem"},{v:"subcategoria",l:"Subcategoría"},{v:"buque",l:"Buque"}];
  const COL_OPT  = [{v:"pedido",l:"Total por pedido"},{v:"dia",l:"Promedio por día"},{v:"pax_dia",l:"Por pax·día"}];
  const MET_OPT  = [{v:"cantidad",l:"Cantidad (unidades)"},{v:"volumen",l:"Volumen analítico (kg)"}];

  return (
    <div>
      {/*  Selector de pedidos  */}
      <div className="card" style={{marginBottom:14}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom: modoSel ? 10 : 0}}>
          <div style={{fontWeight:700,fontSize:13,color:"var(--navy)"}}>
             Pedidos a analizar
            <span style={{fontWeight:400,fontSize:11,color:"var(--muted)",marginLeft:8}}>
              {modoSel && seleccionados.size > 0
                ? `${seleccionados.size} pedido${seleccionados.size>1?"s":""} seleccionado${seleccionados.size>1?"s":""}`
                : `${pedidosBase.length} pedido${pedidosBase.length!==1?"s":""} (todos los filtrados)`}
            </span>
          </div>
          <div style={{display:"flex",gap:6}}>
            {modoSel && seleccionados.size > 0 &&
              <button className="btn btn-ghost btn-sm" onClick={()=>setSeleccionados(new Set())}>Limpiar selección</button>
            }
            <button
              className={`btn btn-sm ${modoSel ? "btn-primary" : "btn-ghost"}`}
              onClick={()=>{ setModoSel(v=>!v); if(modoSel) setSeleccionados(new Set()); }}
            >
              {modoSel ? "✓ Selección activa" : "Seleccionar pedidos"}
            </button>
          </div>
        </div>

        {modoSel && (
          <div>
            <div style={{display:"flex",gap:6,marginBottom:8}}>
              <button className="btn btn-ghost btn-sm" onClick={()=>setSeleccionados(new Set(pedidosBase.map(p=>p.id)))}>Seleccionar todos</button>
              <button className="btn btn-ghost btn-sm" onClick={()=>setSeleccionados(new Set())}>Deseleccionar todos</button>
            </div>
            <div style={{maxHeight:220,overflowY:"auto",display:"flex",flexDirection:"column",gap:4}}>
              {pedidosBase.map(p => {
                const checked = seleccionados.has(p.id);
                const itemsCnt = (p.viveres_pedido_items||[]).filter(it=>(it.cantidad_pedida||0)>0).length;
                return (
                  <label key={p.id} style={{display:"flex",alignItems:"center",gap:10,padding:"7px 10px",borderRadius:6,border:`1px solid ${checked?"var(--accent)":"var(--border)"}`,background:checked?"#EFF6FF":"var(--surface)",cursor:"pointer",transition:"all .12s"}}>
                    <input
                      type="checkbox"
                      checked={checked}
                      onChange={()=>setSeleccionados(prev=>{
                        const next = new Set(prev);
                        checked ? next.delete(p.id) : next.add(p.id);
                        return next;
                      })}
                      style={{width:"auto",accentColor:"var(--accent)"}}
                    />
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
                        <span style={{fontWeight:600,fontSize:12}}>{p.base_buque || "Sin buque"}</span>
                        <span style={{fontSize:11,color:"var(--muted)",fontFamily:"var(--mono)"}}>{(p.fecha_pedido||"").slice(0,10)}</span>
                        <span style={{fontSize:10,padding:"1px 6px",borderRadius:4,background:"#DBEAFE",color:"#1E40AF",fontWeight:600}}>{p.pax} PAX · {p.dias} días</span>
                        {p.solicitado_por && <span style={{fontSize:10,color:"var(--muted)"}}>{p.solicitado_por}</span>}
                      </div>
                    </div>
                    <span style={{fontSize:10,color:"var(--muted)",whiteSpace:"nowrap"}}>{itemsCnt} ítems</span>
                  </label>
                );
              })}
              {pedidosBase.length === 0 && <div style={{fontSize:12,color:"var(--muted)",padding:"8px 0"}}>Sin pedidos para los filtros actuales</div>}
            </div>
          </div>
        )}
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:12, marginBottom:14 }}>
        {[[FILA_OPT,"filas",setFilas,filas,"Filas — agrupar por"],[COL_OPT,"columnas",setColumnas,columnas,"Columnas — normalizar"],[MET_OPT,"metrica",setMetrica,metrica,"Métrica"]].map(([opts,name,setter,val,titulo]) => (
          <div key={name} className="card" style={{padding:"12px 14px"}}>
            <div style={{fontSize:10,color:"var(--muted)",fontWeight:700,textTransform:"uppercase",letterSpacing:".5px",marginBottom:8}}>{titulo}</div>
            {opts.map(o => (
              <label key={o.v} style={{display:"flex",alignItems:"center",gap:7,fontSize:12,cursor:"pointer",padding:"3px 0"}}>
                <input type="radio" name={name} value={o.v} checked={val===o.v} onChange={()=>{setter(o.v); if(name==="filas") setExpandidos({});}} style={{width:"auto",accentColor:"var(--accent)"}} />
                {o.l}
              </label>
            ))}
          </div>
        ))}
      </div>

      {/* Filtros */}
      <div className="filter-row" style={{marginBottom:14}}>
        <select value={filtBuque} onChange={e=>setFiltBuque(e.target.value)} style={{minWidth:160}}>
          <option value="">Todos los buques</option>
          {buques.map(b=><option key={b}>{b}</option>)}
        </select>
        <select value={filtEstado} onChange={e=>setFiltEstado(e.target.value)} style={{minWidth:140}}>
          <option value="">Todos los estados</option>
          {estados.map(s=><option key={s}>{s}</option>)}
        </select>
        <input type="date" value={filtDesde} onChange={e=>setFiltDesde(e.target.value)} style={{width:140}} />
        <span style={{fontSize:11,color:"var(--muted)"}}>→</span>
        <input type="date" value={filtHasta} onChange={e=>setFiltHasta(e.target.value)} style={{width:140}} />
        <input value={busqueda} onChange={e=>setBusqueda(e.target.value)} placeholder="Buscar ítem / categoría..." style={{flex:1,minWidth:160}} />
        {(filtBuque||filtEstado||filtDesde||filtHasta||busqueda) && (
          <button className="btn btn-ghost btn-sm" onClick={()=>{setFiltBuque("");setFiltEstado("");setFiltDesde("");setFiltHasta("");setBusqueda("");}}>✕ Limpiar</button>
        )}
      </div>

      {/* KPIs rápidos */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:14}}>
        {[
          {l:"Pedidos",       v:pedidosFilt.length},
          {l:"Filas",         v:filaKeys.length},
          {l:"Columnas",      v:colKeys.length},
          {l:"Total acum.",   v:(totalGlobal%(1)===0?totalGlobal.toFixed(0):totalGlobal.toFixed(1))+(metrica==="volumen"?" kg":" u")},
        ].map(k=>(
          <div key={k.l} className="card" style={{padding:"10px 14px"}}>
            <div style={{fontSize:10,color:"var(--muted)",fontWeight:700,textTransform:"uppercase",letterSpacing:".5px",marginBottom:3}}>{k.l}</div>
            <div style={{fontSize:20,fontWeight:800,fontFamily:"var(--mono)",color:"var(--navy)"}}>{k.v}</div>
          </div>
        ))}
      </div>

      {/* Tabla pivot */}
      {filaKeys.length === 0
        ? <div className="state-empty">Sin datos para los filtros seleccionados</div>
        : (
        <div className="card">
          <div className="card-title" style={{fontSize:12}}>
            {FILA_OPT.find(o=>o.v===filas)?.l}
            <span style={{fontWeight:400,color:"var(--muted)",margin:"0 6px"}}>×</span>
            {COL_OPT.find(o=>o.v===columnas)?.l}
            <span style={{fontWeight:400,color:"var(--muted)",fontSize:11,marginLeft:8}}>· {MET_OPT.find(o=>o.v===metrica)?.l}</span>
            <span style={{fontWeight:400,color:"var(--muted2)",fontSize:10,marginLeft:8}}>({filaKeys.length} filas · {colKeys.length} col.)</span>
          </div>
          <div className="table-wrap" style={{maxHeight:580,overflowX:"auto",overflowY:"auto"}}>
            <table style={{borderCollapse:"collapse",fontSize:12,minWidth:"100%"}}>
              <thead>
                <tr style={{position:"sticky",top:0,zIndex:4}}>
                  <th style={{minWidth:220,textAlign:"left",padding:"8px 12px",background:"#1A2A46",color:"#fff",fontWeight:700,fontSize:11,letterSpacing:".4px",textTransform:"uppercase",position:"sticky",left:0,zIndex:5}}>
                    {FILA_OPT.find(o=>o.v===filas)?.l}
                  </th>
                  {colKeys.map(ck=>(
                    <th key={ck} title={ck} style={{minWidth:120,textAlign:"right",padding:"8px 10px",background:"#1A2A46",color:"rgba(255,255,255,.7)",fontWeight:600,fontSize:10,letterSpacing:".3px",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",maxWidth:150}}>
                      {ck.length>18?ck.slice(0,18)+"…":ck}
                    </th>
                  ))}
                  <th style={{minWidth:100,textAlign:"right",padding:"8px 10px",background:"#0B1629",color:"var(--gold-light,#D4AA3A)",fontWeight:800,fontSize:11,letterSpacing:".4px",textTransform:"uppercase",position:"sticky",right:0}}>
                    TOTAL
                  </th>
                </tr>
              </thead>
              <tbody>
                {filaKeys.flatMap(fk => {
                  const tot = totalFila(fk);
                  const barPct = Math.min(tot/maxTotal, 1);
                  const isExp = !!expandidos[fk];
                  const canExpand = filas !== "item";
                  const subData = isExp ? buildSubFilas(fk) : null;

                  const mainRow = (
                    <tr
                      key={fk}
                      style={{cursor:canExpand?"pointer":"default", borderBottom:"1px solid var(--border)"}}
                      onClick={()=>canExpand&&setExpandidos(prev=>({...prev,[fk]:!prev[fk]}))}
                    >
                      <td style={{padding:"8px 12px",background:"var(--surface)",position:"sticky",left:0,zIndex:2,borderRight:"1px solid var(--border)"}}>
                        <div style={{display:"flex",alignItems:"center",gap:7}}>
                          {canExpand && <span style={{color:"var(--muted2)",fontSize:10,width:10,flexShrink:0}}>{isExp?"":""}</span>}
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontWeight:600,fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{fk}</div>
                            <div style={{marginTop:3,height:2,borderRadius:1,background:"var(--border)"}}>
                              <div style={{height:"100%",width:`${barPct*100}%`,background:"var(--accent)",borderRadius:1}}/>
                            </div>
                          </div>
                        </div>
                      </td>
                      {colKeys.map(ck => {
                        const v = tabla.get(fk)?.get(ck)||0;
                        return (
                          <td key={ck} style={{padding:"7px 10px",textAlign:"right",background:heatBg(v,maxCell),borderBottom:"1px solid var(--border)"}}>
                            {fmtVal(v)}
                          </td>
                        );
                      })}
                      <td style={{padding:"7px 10px",textAlign:"right",fontWeight:700,background:"#F0F4F8",position:"sticky",right:0,borderLeft:"1px solid var(--border)"}}>
                        {fmtVal(tot)}
                      </td>
                    </tr>
                  );

                  const subRows = (isExp && subData)
                    ? subData.keys.map(sk => {
                        const sTot = [...(subData.map.get(sk)?.values()||[])].reduce((s,v)=>s+v,0);
                        return (
                          <tr key={`${fk}__${sk}`} style={{background:"#FAFBFD",borderBottom:"1px solid var(--border)"}}>
                            <td style={{padding:"5px 12px 5px 36px",position:"sticky",left:0,background:"#FAFBFD",borderRight:"1px solid var(--border)"}}>
                              <span style={{fontSize:11,color:"var(--text)"}}>↳ </span>
                              <span style={{fontSize:11,color:"var(--text)",fontWeight:500}}>{sk}</span>
                            </td>
                            {colKeys.map(ck => {
                              const v = subData.map.get(sk)?.get(ck)||0;
                              return (
                                <td key={ck} style={{padding:"5px 10px",textAlign:"right",fontSize:11,background:v?heatBg(v,maxCell*0.5):"transparent"}}>
                                  {fmtVal(v)}
                                </td>
                              );
                            })}
                            <td style={{padding:"5px 10px",textAlign:"right",fontWeight:600,fontSize:11,background:"#F0F4F8",position:"sticky",right:0,borderLeft:"1px solid var(--border)"}}>
                              {fmtVal(sTot)}
                            </td>
                          </tr>
                        );
                      })
                    : [];

                  return [mainRow, ...subRows];
                })}
              </tbody>
              <tfoot>
                <tr style={{background:"#1A2A46",position:"sticky",bottom:0}}>
                  <td style={{padding:"8px 12px",color:"#fff",fontWeight:700,fontSize:11,textTransform:"uppercase",letterSpacing:".4px",position:"sticky",left:0,background:"#1A2A46"}}>
                    TOTAL GENERAL
                  </td>
                  {colKeys.map(ck=>{
                    const colTot = filaKeys.reduce((s,fk)=>s+(tabla.get(fk)?.get(ck)||0),0);
                    return (
                      <td key={ck} style={{padding:"8px 10px",textAlign:"right",color:"rgba(255,255,255,.9)"}}>
                        {fmtVal(colTot)}
                      </td>
                    );
                  })}
                  <td style={{padding:"8px 10px",textAlign:"right",color:"#D4AA3A",fontWeight:800,position:"sticky",right:0,background:"#0B1629"}}>
                    {fmtVal(totalGlobal)}
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function ViveresApp() {
  const [page, setPage] = useState("inbox");
  const [notif, setNotif] = useState(null);
  const [inboxCount, setInboxCount] = useState(0);
  const notify = useCallback((text, type = "info") => { setNotif({ text, type }); setTimeout(() => setNotif(null), 4000); }, []);
  const loadCounts = useCallback(async () => { try { const d = await api.getPedidos({ status: "enviado" }); setInboxCount(d.length); } catch (e) { console.error(e); } }, []);
  useEffect(() => { loadCounts(); }, [loadCounts]);

  const [navOpen, setNavOpen] = useState(true);

  /* Iconos de línea · trazo 1,6 · sin relleno · toman el color del texto. */
  const Ico = ({ d, size = 18 }) => (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor"
         strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{d}</svg>
  );
  const ICONS = {
    inbox:  <><path d="M3 12h5l1.5 2.5h5L16 12h5" /><path d="M4.6 5.4h14.8l1.6 6.6v6a1.6 1.6 0 0 1-1.6 1.6H4.6A1.6 1.6 0 0 1 3 18V12z" /></>,
    cart:   <><circle cx="9.5" cy="19" r="1.4" /><circle cx="17.5" cy="19" r="1.4" /><path d="M3 4h2.2l2.4 10.2h10.6L21 7.5H6.4" /></>,
    list:   <><path d="M9 6h11M9 12h11M9 18h11" /><path d="M4.5 6h.01M4.5 12h.01M4.5 18h.01" /></>,
    chart:  <><path d="M4 20V10M10 20V4M16 20v-7M22 20H2" /></>,
    box:    <><path d="M12 3l8.5 4.5v9L12 21 3.5 16.5v-9z" /><path d="M3.5 7.5L12 12l8.5-4.5M12 12v9" /></>,
    grid:   <><rect x="3" y="3" width="18" height="18" rx="2" /><path d="M3 9h18M3 15h18M9 3v18M15 3v18" /></>,
    panel:  <><rect x="3" y="4" width="18" height="16" rx="2" /><path d="M9.5 4v16" /></>,
    bell:   <><path d="M18 8a6 6 0 1 0-12 0c0 7-3 8-3 8h18s-3-1-3-8" /><path d="M13.7 21a2 2 0 0 1-3.4 0" /></>,
    help:   <><circle cx="12" cy="12" r="9" /><path d="M9.5 9.5a2.5 2.5 0 1 1 3.6 2.3c-.7.4-1.1 1-1.1 1.7v.3" /><path d="M12 17.5h.01" /></>,
    users:  <><path d="M16 20v-1.5a3.5 3.5 0 0 0-3.5-3.5h-5A3.5 3.5 0 0 0 4 18.5V20" /><circle cx="10" cy="8" r="3.2" /><path d="M20 20v-1.5a3.5 3.5 0 0 0-2.6-3.4M15.5 4.8a3.2 3.2 0 0 1 0 6.2" /></>,
  };

  const SECCIONES = {
    inbox:     { grupo: "Bandeja",     titulo: "Pedidos por revisar",  sub: "Pedidos de víveres que esperan revisión antes de pasar a compra." },
    nuevo:     { grupo: "Pedidos",     titulo: "Nuevo pedido",         sub: "Cargá el pedido por embarcación. La dieta y la dotación definen las cantidades." },
    historial: { grupo: "Pedidos",     titulo: "Historial de pedidos", sub: "Todos los pedidos cargados, con su estado y su costo por cabeza y día." },
    tracker:   { grupo: "Seguimiento", titulo: "Seguimiento de entregas", sub: "Avance de cada pedido desde la compra hasta la recepción a bordo." },
    stock_vuelta: { grupo: "Seguimiento", titulo: "Stock vuelta a puerto", sub: "Registrá el stock que queda a bordo cuando un buque vuelve a puerto, para completar el próximo pedido de ese buque." },
    catalogo:  { grupo: "Datos",       titulo: "Catálogo de víveres",  sub: "Artículos habilitados, con unidad, rubro y precio de referencia." },
    solicitantes: { grupo: "Datos",    titulo: "Solicitantes",         sub: "Nombres habilitados para crear pedidos. Estandarizá quién puede solicitar víveres." },
    pivot:     { grupo: "Datos",       titulo: "Análisis pivot",       sub: "Consumo y costo cruzados por embarcación, rubro y período." },
  };

  const NAV = [
    { titulo: "Bandeja", items: [
      { id: "inbox", icon: "inbox", label: "Pedidos por revisar", count: inboxCount },
    ]},
    { titulo: "Pedidos", items: [
      { id: "nuevo",     icon: "cart", label: "Nuevo pedido", count: 0 },
      { id: "historial", icon: "list", label: "Historial",    count: 0 },
    ]},
    { titulo: "Seguimiento", items: [
      { id: "tracker", icon: "chart", label: "Seguimiento de entregas", count: 0 },
      { id: "stock_vuelta", icon: "box", label: "Stock vuelta a puerto", count: 0 },
    ]},
    { titulo: "Datos", items: [
      { id: "catalogo",     icon: "box",   label: "Catálogo",       count: 0 },
      { id: "solicitantes", icon: "users", label: "Solicitantes",   count: 0 },
      { id: "pivot",        icon: "grid",  label: "Análisis pivot", count: 0 },
    ]},
  ];

  const seccion = SECCIONES[page] || { grupo: "Víveres", titulo: page, sub: "" };
  const inicial = (USUARIO || "C").replace(/@.*$/, "").slice(0, 2).toUpperCase();

  return (
    <>
      <style>{CSS}</style>

      <header className="appbar">
        <img src="/integra-isotipo-white.svg" alt="INTEGRA" className="appbar-iso" />
        <span className="appbar-div" />
        <span className="appbar-instance">PL Offshore</span>
        <input className="appbar-search" type="search" disabled placeholder="Buscar en todo INTEGRA" aria-label="Buscar" />
        <div className="appbar-tools">
          <span style={{ color: "rgba(255,255,255,.86)", display: "block" }}><Ico d={ICONS.bell} /></span>
          <span style={{ color: "rgba(255,255,255,.86)", display: "block" }}><Ico d={ICONS.help} /></span>
          <span className="appbar-div" />
          <span className="appbar-avatar">{inicial}</span>
          <span className="appbar-user">{USUARIO}</span>
          <button className="appbar-link" onClick={() => window.location.href = PORTAL_URL}>Volver al portal</button>
        </div>
      </header>

      <div className={`shell ${navOpen ? "" : "is-collapsed"}`}>
        <nav className="sidebar">
          <div className="sidebar-header">
            <img src="/PL.png" alt="PL Offshore" className="sidebar-logo-img" />
            {navOpen && (
              <div>
                <div className="sidebar-logo-main">Víveres</div>
                <div className="sidebar-logo-sub">PL Offshore</div>
              </div>
            )}
          </div>

          <div className="sidebar-nav">
            {NAV.map(grupo => (
              <div key={grupo.titulo} style={{ marginBottom: 8 }}>
                {navOpen && <div className="nav-section">{grupo.titulo}</div>}
                {grupo.items.map(it => (
                  <button
                    key={it.id}
                    className={`ni ${page === it.id ? "active" : ""}`}
                    onClick={() => setPage(it.id)}
                    title={it.label}
                  >
                    <span className="ni-ico"><Ico d={ICONS[it.icon]} /></span>
                    {navOpen && <span className="ni-label">{it.label}</span>}
                    {it.count > 0 && <span className="ni-badge">{it.count}</span>}
                  </button>
                ))}
              </div>
            ))}
          </div>

          <div className="sidebar-foot">
            <button className="sidebar-foot-btn" onClick={() => setNavOpen(v => !v)}>
              <span style={{ display: "block", color: "var(--muted2)" }}><Ico d={ICONS.panel} size={16} /></span>
              {navOpen && <span style={{ flex: 1, textAlign: "left" }}>Colapsar menú</span>}
            </button>
            {navOpen && (
              <div className="sidebar-foot-meta">
                <div>VÍVERES v2.2</div>
                <div>POWERED BY INTEGRA</div>
              </div>
            )}
          </div>
        </nav>

        <div className="main">
          <div className="pagehead">
            <div className="crumb">
              <button onClick={() => window.location.href = PORTAL_URL}>Portal</button>
              <span>/</span>
              <button onClick={() => setPage("inbox")}>Víveres</button>
              <span>/</span>
              <span className="crumb-current">{seccion.titulo}</span>
            </div>
            <div className="pagehead-row">
              <div>
                <h1>{seccion.titulo}</h1>
                {seccion.sub && <p>{seccion.sub}</p>}
              </div>
              {page !== "nuevo" && (
                <div className="pagehead-actions">
                  <button className="btn btn-primary" onClick={() => setPage("nuevo")}>Nuevo pedido</button>
                </div>
              )}
            </div>
          </div>

          <div className="content">
            {page === "inbox"     && <PageInbox notify={notify} onNeedRefresh={loadCounts} />}
            {page === "nuevo"     && <PageNuevo notify={notify} onSaved={() => { setPage("historial"); loadCounts(); }} onCancel={() => setPage("historial")} />}
            {page === "historial" && <PageHistorial onNuevo={() => setPage("nuevo")} notify={notify} />}
            {page === "tracker"   && <PageTracker notify={notify} />}
            {page === "stock_vuelta" && <PageStockVuelta notify={notify} />}
            {page === "catalogo"  && <PageCatalogo notify={notify} />}
            {page === "solicitantes" && <PageSolicitantes notify={notify} />}
            {page === "pivot"     && <PagePivot />}
          </div>
        </div>
      </div>

      <Notif msg={notif} onClose={() => setNotif(null)} />

      <nav className="mobile-nav">
        {[
          { id: "inbox",     label: "Bandeja",  icon: "inbox", count: inboxCount },
          { id: "nuevo",     label: "Nuevo",    icon: "cart",  count: 0 },
          { id: "historial", label: "Historial",icon: "list",  count: 0 },
          { id: "tracker",   label: "Entregas", icon: "chart", count: 0 },
          { id: "stock_vuelta", label: "Stock", icon: "box", count: 0 },
          { id: "catalogo",  label: "Catálogo", icon: "box",   count: 0 },
        ].map(it => (
          <div
            key={it.id}
            className={`mobile-nav-item ${page === it.id ? "active" : ""}`}
            onClick={() => setPage(it.id)}
          >
            <span className="mobile-nav-icon"><Ico d={ICONS[it.icon]} size={18} /></span>
            <span className="mobile-nav-label">{it.label}</span>
            {it.count > 0 && <span className="mobile-nav-badge">{it.count}</span>}
          </div>
        ))}
      </nav>
    </>
  );
}

export default function App() {
  const [session, setSession] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      setLoading(false);
    });

    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
    });

    return () => subscription.unsubscribe();
  }, []);

  if (loading) return (
    <div style={{ minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", background:"#213363" }}>
      <div style={{ fontFamily:"'DM Mono',monospace", fontSize:10, color:"rgba(255,255,255,0.3)", letterSpacing:3, textTransform:"uppercase" }}>Cargando...</div>
    </div>
  );

  if (!session) return <LoginPage />;

  return <ViveresApp />;
}
