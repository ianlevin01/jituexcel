"use client";

import { useEffect, useMemo, useState } from "react";

export default function Home() {
  const [clientes, setClientes] = useState([]);
  const [busqueda, setBusqueda] = useState("");
  const [clienteSeleccionado, setClienteSeleccionado] = useState(null);
  const [archivo, setArchivo] = useState(null);
  const [procesando, setProcesando] = useState(false);
  const [mensaje, setMensaje] = useState(null);

  useEffect(() => {
    fetch("/api/clientes")
      .then((r) => r.json())
      .then(setClientes)
      .catch(() => setMensaje({ tipo: "error", texto: "No se pudo cargar la lista de clientes." }));
  }, []);

  const clientesFiltrados = useMemo(() => {
    const q = busqueda.trim().toLowerCase();
    if (!q) return clientes;
    return clientes.filter((c) => c.nombre.toLowerCase().includes(q));
  }, [clientes, busqueda]);

  function cambiarBusqueda(valor) {
    setBusqueda(valor);
    if (clienteSeleccionado && valor !== clienteSeleccionado.nombre) {
      setClienteSeleccionado(null);
    }
  }

  function limpiarSeleccion() {
    setClienteSeleccionado(null);
    setBusqueda("");
  }

  async function descargar() {
    if (!clienteSeleccionado || !archivo) return;

    setProcesando(true);
    setMensaje({ tipo: "pendiente", texto: "Procesando..." });

    try {
      const formData = new FormData();
      formData.append("clienteId", clienteSeleccionado.clienteId);
      formData.append("file", archivo);

      const respuesta = await fetch("/api/reprice", { method: "POST", body: formData });

      if (!respuesta.ok) {
        const error = await respuesta.json().catch(() => ({}));
        throw new Error(error.error || "No se pudo generar el excel.");
      }

      const filasOmitidas = Number(respuesta.headers.get("X-Filas-Omitidas") || "0");
      const blob = await respuesta.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "actualizados.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      let texto = `Listo. Se descargó actualizados.xlsx para "${clienteSeleccionado.nombre}" (${clienteSeleccionado.porcentaje}%).`;
      if (filasOmitidas > 0) {
        texto += `\nSe omitieron ${filasOmitidas} fila(s) sin coincidencia en el excel base.`;
      }
      setMensaje({ tipo: "ok", texto });
    } catch (err) {
      setMensaje({ tipo: "error", texto: err.message });
    } finally {
      setProcesando(false);
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
          <div className="card-header">
            <div className="icon">
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M4 4h11l5 5v11a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1Z" stroke="white" strokeWidth="1.6" strokeLinejoin="round" />
                <path d="M14 4v5h5" stroke="white" strokeWidth="1.6" strokeLinejoin="round" />
                <path d="M8 13.5h2M8 16.5h5M14.5 13.5v3" stroke="white" strokeWidth="1.6" strokeLinecap="round" />
              </svg>
            </div>
            <h1>Descargar precios de cliente</h1>
          </div>
          <p className="subtitle">
            Buscá un cliente, subí su excel y descargá los precios actualizados con su porcentaje.
          </p>

          <div className="field">
            <label htmlFor="busqueda">Cliente</label>
            <input
              type="text"
              id="busqueda"
              placeholder="Buscar por nombre..."
              value={busqueda}
              onChange={(e) => cambiarBusqueda(e.target.value)}
              autoComplete="off"
            />
          </div>

          {busqueda.trim() && !clienteSeleccionado && (
            <ul className="lista-clientes" style={{ marginBottom: 20 }}>
              {clientesFiltrados.length === 0 && <li>Sin resultados.</li>}
              {clientesFiltrados.map((c) => (
                <li
                  key={c.clienteId}
                  className="seleccionable"
                  onClick={() => {
                    setClienteSeleccionado(c);
                    setBusqueda(c.nombre);
                  }}
                >
                  <span className="cliente-nombre">{c.nombre}</span>
                  <span className="cliente-porcentaje">{c.porcentaje}%</span>
                </li>
              ))}
            </ul>
          )}

          {clienteSeleccionado && (
            <ul className="lista-clientes" style={{ marginBottom: 20 }}>
              <li className="seleccionado">
                <span className="cliente-nombre">{clienteSeleccionado.nombre}</span>
                <span className="cliente-porcentaje">{clienteSeleccionado.porcentaje}%</span>
                <div className="acciones-fila">
                  <button type="button" className="secundario" onClick={limpiarSeleccion}>
                    Cambiar
                  </button>
                </div>
              </li>
            </ul>
          )}

          <div className="field">
            <label htmlFor="archivo">Excel a modificar</label>
            <input
              type="file"
              id="archivo"
              accept=".xlsx"
              onChange={(e) => setArchivo(e.target.files[0] || null)}
            />
          </div>

          <button onClick={descargar} disabled={!clienteSeleccionado || !archivo || procesando}>
            Descargar
          </button>

          {mensaje && <div className={`mensaje ${mensaje.tipo}`}>{mensaje.texto}</div>}
        </div>
      </div>
    </div>
  );
}
