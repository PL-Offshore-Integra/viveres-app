import { supabase } from "./supabase";

// Los buckets de Storage son privados. En la base queda guardada la URL pública del archivo
// (/storage/v1/object/public/<bucket>/<ruta>) y, para abrirlo, se pide una URL firmada de
// corta duración. Los links que no son de Storage (SharePoint, etc.) se devuelven tal cual.
const DURACION_SEG = 600;

export async function firmarUrl(url) {
  const m = /\/storage\/v1\/object\/(?:public|sign)\/([^/]+)\/([^?#]+)/.exec(url || "");
  if (!m) return url;
  let ruta = m[2];
  try { ruta = decodeURIComponent(ruta); } catch { /* la ruta no venía codificada */ }
  const { data, error } = await supabase.storage.from(m[1]).createSignedUrl(ruta, DURACION_SEG);
  if (error) throw error;
  return data.signedUrl;
}

// Reemplazo de <a href={url} target="_blank">. La pestaña se abre dentro del click (si se abre
// después del await, el navegador la bloquea como popup) y se redirige cuando llega la URL firmada.
export function ArchivoLink({ href, onClick, children, ...props }) {
  const abrir = async (e) => {
    onClick?.(e);
    e.preventDefault();
    const win = window.open("", "_blank");
    try {
      const url = await firmarUrl(href);
      if (win) win.location.href = url; else window.location.href = url;
    } catch (err) {
      win?.close();
      alert("No se pudo abrir el archivo: " + (err?.message || err));
    }
  };
  return <a href={href} target="_blank" rel="noreferrer" {...props} onClick={abrir}>{children}</a>;
}
