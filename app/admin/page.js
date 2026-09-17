"use client";

import { useEffect, useState } from "react";

function formatearFecha(iso) {
  if (!iso) return "";
  return new Date(iso).toLocaleString("es-AR");
}

function formatearBytes(bytes) {
  if (!bytes && bytes !== 0) return "";
  if (bytes < 1024) return `${bytes} B`;
  return `${(bytes / 1024).toFixed(1)} KB`;
}

export default function Admin() {
  const [clientes, setClientes] = useState([]);
  const [cargandoClientes, setCargandoClientes] = useState(true);
  const [editando, setEditando] = useState(null);
  const [nombre, setNombre] = useState("");
  const [porcentaje, setPorcentaje] = useState("");
  const [mensajeClientes, setMensajeClientes] = useState(null);

  const [baseInfo, setBaseInfo] = useState(null);
  const [archivoBase, setArchivoBase] = useState(null);
  const [subiendoBase, setSubiendoBase] = useState(false);
  const [mensajeBase, setMensajeBase] = useState(null);

  const [archivoOrdenar, setArchivoOrdenar] = useState(null);
  const [ordenando, setOrdenando] = useState(false);
  const [mensajeOrdenar, setMensajeOrdenar] = useState(null);
  const [resumenOrdenar, setResumenOrdenar] = useState(null);

  function cargarClientes() {
    setCargandoClientes(true);
    return fetch("/api/clientes")
      .then((r) => r.json())
      .then(setClientes)
      .catch(() => setMensajeClientes({ tipo: "error", texto: "No se pudo cargar la lista de clientes." }))
      .finally(() => setCargandoClientes(false));
  }

  function cargarBaseInfo() {
    return fetch("/api/base-excel")
      .then((r) => r.json())
      .then(setBaseInfo)
      .catch(() => {});
  }

  useEffect(() => {
    cargarClientes();
    cargarBaseInfo();
  }, []);

  function limpiarFormulario() {
    setEditando(null);
    setNombre("");
    setPorcentaje("");
  }

  function editarCliente(cliente) {
    setEditando(cliente);
    setNombre(cliente.nombre);
    setPorcentaje(String(cliente.porcentaje));
  }

  async function guardarCliente(e) {
    e.preventDefault();
    setMensajeClientes(null);

    const payload = { nombre: nombre.trim(), porcentaje: Number(porcentaje) };
    if (!payload.nombre) {
      setMensajeClientes({ tipo: "error", texto: "El nombre es obligatorio." });
      return;
    }
    if (Number.isNaN(payload.porcentaje)) {
      setMensajeClientes({ tipo: "error", texto: "El porcentaje debe ser un número." });
      return;
    }

    try {
      const url = editando ? `/api/clientes/${editando.clienteId}` : "/api/clientes";
      const method = editando ? "PUT" : "POST";
      const respuesta = await fetch(url, {
        method,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (!respuesta.ok) {
        const error = await respuesta.json().catch(() => ({}));
        throw new Error(error.error || "No se pudo guardar el cliente.");
      }
      setMensajeClientes({ tipo: "ok", texto: editando ? "Cliente actualizado." : "Cliente creado." });
      limpiarFormulario();
      cargarClientes();
    } catch (err) {
      setMensajeClientes({ tipo: "error", texto: err.message });
    }
  }

  async function borrarCliente(cliente) {
    if (!confirm(`¿Borrar a "${cliente.nombre}"?`)) return;
    try {
      const respuesta = await fetch(`/api/clientes/${cliente.clienteId}`, { method: "DELETE" });
      if (!respuesta.ok) throw new Error("No se pudo borrar el cliente.");
      if (editando?.clienteId === cliente.clienteId) limpiarFormulario();
      cargarClientes();
    } catch (err) {
      setMensajeClientes({ tipo: "error", texto: err.message });
    }
  }

  async function subirBase(e) {
    e.preventDefault();
    if (!archivoBase) return;

    setSubiendoBase(true);

    try {
      setMensajeBase({ tipo: "pendiente", texto: "Preparando subida..." });
      const presignRes = await fetch("/api/base-excel/presigned-url", { method: "POST" });
      const presignData = await presignRes.json().catch(() => ({}));
      if (!presignRes.ok) throw new Error(presignData.error || "No se pudo iniciar la subida.");
      const { url, key } = presignData;

      setMensajeBase({ tipo: "pendiente", texto: "Subiendo archivo a S3..." });
      const putRes = await fetch(url, {
        method: "PUT",
        headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        body: archivoBase,
      });
      if (!putRes.ok) throw new Error("No se pudo subir el archivo a S3.");

      setMensajeBase({ tipo: "pendiente", texto: "Validando columnas..." });
      const confirmRes = await fetch("/api/base-excel", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ key, tamanioEsperado: archivoBase.size }),
      });
      const data = await confirmRes.json().catch(() => ({}));
      if (!confirmRes.ok) throw new Error(data.error || "No se pudo confirmar el excel base.");

      setBaseInfo(data);
      setArchivoBase(null);
      setMensajeBase({ tipo: "ok", texto: "Excel base actualizado." });
    } catch (err) {
      setMensajeBase({ tipo: "error", texto: err.message });
    } finally {
      setSubiendoBase(false);
    }
  }

  async function ordenarExcel(e) {
    e.preventDefault();
    if (!archivoOrdenar) return;

    setOrdenando(true);
    setResumenOrdenar(null);

    try {
      setMensajeOrdenar({ tipo: "pendiente", texto: "Subiendo archivo..." });
      const presignRes = await fetch("/api/ordenar-excel/upload-url", { method: "POST" });
      const presignData = await presignRes.json().catch(() => ({}));
      if (!presignRes.ok) throw new Error(presignData.error || "No se pudo iniciar la subida.");
      const { url, key } = presignData;

      const putRes = await fetch(url, {
        method: "PUT",
        headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        body: archivoOrdenar,
      });
      if (!putRes.ok) throw new Error("No se pudo subir el archivo a S3.");

      setMensajeOrdenar({ tipo: "pendiente", texto: "Clasificando imágenes con IA... puede tardar un rato." });
      const respuesta = await fetch("/api/ordenar-excel", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ key, tamanioEsperado: archivoOrdenar.size }),
      });
      const data = await respuesta.json().catch(() => ({}));
      if (!respuesta.ok) throw new Error(data.error || "No se pudo ordenar el excel.");

      const a = document.createElement("a");
      a.href = data.downloadUrl;
      a.download = "ordenado.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();

      setResumenOrdenar(data.resumen || null);
      setArchivoOrdenar(null);
      setMensajeOrdenar({ tipo: "ok", texto: "Listo, se descargó ordenado.xlsx." });
    } catch (err) {
      setMensajeOrdenar({ tipo: "error", texto: err.message });
    } finally {
      setOrdenando(false);
    }
  }

  return (
    <div className="page">
      <div>
        <nav className="top-nav">
          <a href="/">Descargar</a>
          <a href="/admin">Administración</a>
        </nav>

        <div className="card">
          <h2>Ordenar excel por imagen</h2>
          <p className="subtitle">
            Sub&iacute; cualquier excel con fotos de productos y lo reordena agrupando los que se parecen (usa IA para
            clasificar cada foto).
          </p>

          <form onSubmit={ordenarExcel}>
            <div className="field">
              <label htmlFor="archivoOrdenar">Excel con imágenes</label>
              <input
                type="file"
                id="archivoOrdenar"
                accept=".xlsx"
                onChange={(e) => setArchivoOrdenar(e.target.files[0] || null)}
              />
            </div>
            <button type="submit" disabled={!archivoOrdenar || ordenando}>
              Ordenar y descargar
            </button>
          </form>

          {mensajeOrdenar && <div className={`mensaje ${mensajeOrdenar.tipo}`}>{mensajeOrdenar.texto}</div>}

          {resumenOrdenar && (
            <ul className="lista-clientes" style={{ marginTop: 12 }}>
              {resumenOrdenar.map((r) => (
                <li key={r.categoria}>
                  <span className="cliente-nombre">{r.categoria}</span>
                  <span className="cliente-porcentaje">{r.cantidad} fila(s)</span>
                </li>
              ))}
            </ul>
          )}
        </div>

        <div className="card">
          <h2>Excel base</h2>
          <p className="subtitle">
            {baseInfo?.existe
              ? `Actualizado el ${formatearFecha(baseInfo.ultimaActualizacion)} · ${formatearBytes(baseInfo.tamanioBytes)}`
              : "Todavía no se subió ningún excel base."}
          </p>

          <form onSubmit={subirBase}>
            <div className="field">
              <label htmlFor="archivoBase">
                Reemplazar excel base <span className="hint">(columnas "Item No." y "Precio")</span>
              </label>
              <input
                type="file"
                id="archivoBase"
                accept=".xlsx"
                onChange={(e) => setArchivoBase(e.target.files[0] || null)}
              />
            </div>
            <button type="submit" disabled={!archivoBase || subiendoBase}>
              Subir nueva base
            </button>
          </form>

          {mensajeBase && <div className={`mensaje ${mensajeBase.tipo}`}>{mensajeBase.texto}</div>}
        </div>

        <div className="card">
          <h2>Clientes</h2>
          <p className="subtitle">Cada cliente tiene un porcentaje que se aplica sobre el precio base.</p>

          <form onSubmit={guardarCliente}>
            <div className="grid-2">
              <div className="field">
                <label htmlFor="nombre">Nombre</label>
                <input
                  type="text"
                  id="nombre"
                  value={nombre}
                  onChange={(e) => setNombre(e.target.value)}
                  placeholder="Ej: Distribuidora Sur"
                />
              </div>
              <div className="field">
                <label htmlFor="porcentaje">Porcentaje</label>
                <input
                  type="number"
                  id="porcentaje"
                  step="any"
                  value={porcentaje}
                  onChange={(e) => setPorcentaje(e.target.value)}
                  placeholder="Ej: 25"
                />
              </div>
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button type="submit">{editando ? "Guardar cambios" : "Crear cliente"}</button>
              {editando && (
                <button type="button" className="secundario" onClick={limpiarFormulario}>
                  Cancelar
                </button>
              )}
            </div>
          </form>

          {mensajeClientes && <div className={`mensaje ${mensajeClientes.tipo}`}>{mensajeClientes.texto}</div>}

          <div style={{ marginTop: 20 }}>
            {cargandoClientes && <p className="subtitle">Cargando...</p>}
            {!cargandoClientes && clientes.length === 0 && <p className="subtitle">Todavía no hay clientes.</p>}
            {!cargandoClientes && clientes.length > 0 && (
              <ul className="lista-clientes">
                {clientes.map((c) => (
                  <li key={c.clienteId}>
                    <span className="cliente-nombre">{c.nombre}</span>
                    <span className="cliente-porcentaje">{c.porcentaje}%</span>
                    <div className="acciones-fila">
                      <button type="button" className="secundario" onClick={() => editarCliente(c)}>
                        Editar
                      </button>
                      <button type="button" className="peligro" onClick={() => borrarCliente(c)}>
                        Borrar
                      </button>
                    </div>
                  </li>
                ))}
              </ul>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
